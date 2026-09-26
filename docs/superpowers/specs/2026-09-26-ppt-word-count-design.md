# PowerPoint Word Count add-in — design

Date: 2026-09-26 · Status: approved by the user ("go ahead and build it") · Approach A (plain Office.js)

## 1. Goal

A PowerPoint add-in that shows **Words**, **Characters (no spaces)** and **Characters (with spaces)** for whatever is currently selected, updating live. It is used by a small team on current Microsoft 365 PowerPoint (Mac, Windows, web) and is shared with them through static hosting.

## 2. User-facing behaviour

A **Word Count** button in its own group on the **Home** ribbon tab opens a task pane. The pane shows the three counts in large type and refreshes itself when the selection changes. What is counted:

| Selection | What is counted | Context line |
|---|---|---|
| Text highlighted inside a shape, placeholder or table cell | Only the highlighted text | "Highlighted text in *Title 1*" |
| One or more whole shapes (text boxes, placeholders, shapes with text, WordArt) | All their text, summed | "*TextBox 3*" / "3 shapes selected" |
| A table (as a whole shape) | Every cell (merged cells counted once) | "*Table 2*" |
| A group (any nesting depth) | Every text shape and table inside it | "*Group 5*" |
| Nothing, a slide thumbnail, or only shapes without text (pictures, lines…) | — | Empty state: "Select a text box, table or group, or highlight some text." |

- When two or more shapes are selected, a compact **per-shape breakdown** (name + words) is shown under the totals.
- **Charts, SmartArt, diagrams and embedded (OLE) objects** can't be read by the Office.js API. If any are selected, a note says "N charts/SmartArt not included".
- A small **Refresh** button re-reads the selection by hand, in case PowerPoint misses an update.
- A **Help** link opens a page explaining the counting rules, known limitations and how to install the add-in.
- The pane follows PowerPoint's light/dark theme.

## 3. Counting rules (match Microsoft Word)

These rules come from measuring Word for Mac 16.113.2's own statistics engine on 528 strings. A second agent re-checked the measurements. Implemented in `src/count.js` as a single pass over **Unicode code points**. Intl.Segmenter, `String.length` and regex splitting are not used; each of them disagrees with Word on common text.

Character classes:

| Class | Characters | Word boundary? | In "with spaces"? | In "no spaces"? |
|---|---|---|---|---|
| Paragraph break | CR, LF (CRLF = one), U+000C, U+000E, U+2029 | yes | no | no |
| Space | space, tab, U+000B (Shift+Enter line break), NBSP U+00A0, U+2028, ideographic space U+3000 | yes | yes | no |
| Dash separator | en dash U+2013, em dash U+2014, bullet U+2022 | yes (never a word itself) | yes | yes |
| Ignorable | ZWSP U+200B, ZWNJ U+200C, ZWJ U+200D (except after an emoji base), BOM U+FEFF, U+001F, variation selectors U+FE00–FE0F and U+E0100–E01EF, tags U+E0000–E007F | no | no | no |
| East Asian | Han, Hiragana, Katakana, Bopomofo, CJK symbols/punctuation, full-width forms, circled numbers, small form variants, ※ ‥ ․ ‧ (see `count.js` for the exact table) | each character is its own word | yes | yes |
| Everything else | letters, digits, all other punctuation and symbols, emoji | joins into words | yes | yes |

Extra rules: two or more consecutive ASCII hyphens end the word just after the run (`word--word` is 2 words). A ZWJ after an emoji base (Extended_Pictographic / Emoji_Modifier) counts as a character. Hangul defaults to space-delimited "Korean words".

Examples (Word-verified): `Hello world` → 2/10/11 · `word - word` → 3/9/11 · `word—word` → 2/9/9 · `well-known` → 1 · `https://example.com/a/b?x=1&y=2` → 1 word · `世界` → 2 words · `Hello世界` → 3 · `👍🏽` → 1/2/2.

**Deliberate deviations** (documented in the help page):
- U+2028 is treated as a line break and U+2029 as a paragraph break. Word treats them as East Asian characters.
- Cyrillic and Greek are always alphabetic. Word sometimes flips them to East Asian depending on document state.
- A lone U+FE0F is 0 words; Word counts it as 1.
- A ZWJ after rare text-default symbols (© ™ ⭐ ✅ …) is counted.
- Automatic bullets and numbering are not in PowerPoint's text, so they are never counted. Word counts them.

**Multiple containers:** each text frame, table cell and highlighted range is counted on its own and the results are summed. Strings are never joined, so words never merge across boxes or cells.

## 4. Architecture

Plain HTML/CSS/ES modules with no bundler or framework. XML (add-in only) manifest. Office.js is loaded in `<head>` from `https://officeapis.public.onecdn.static.microsoft/1/office.js`, which Microsoft now recommends.

```
manifest.xml                  dev manifest (https://localhost:3100, "Word Count (dev)")
src/
  taskpane.html               loads office.js + app, contains only <div id="app">
  taskpane.css                all styles; theme via CSS custom properties
  taskpane.js                 bootstrap: Office.onReady → start({ Office, PowerPoint, root })
  app.js                      controller + rendering (framework-free); dependency-injected
  refresher.js                refresh scheduling (debounce, single-flight, retry) — pure, timer-injected
  selection.js                readSelection(context) → SelectionSnapshot (the only Office API reader)
  count.js                    countText(), sumCounts() — pure, no Office dependency
  help.html                   help: rules, limitations, install guide, manifest download
assets/icon-{16,20,24,32,40,48,64,80,128}.png
dev/
  fake-powerpoint.js          in-memory PowerPoint/Office fake with real load/sync semantics
  preview.html                browser preview: scenario picker + 320px pane + theme toggle
scripts/
  serve.mjs                   dev server (HTTPS with dev certs; --http for the preview)
  build.mjs                   dist/ for hosting (+ production manifest)
  make-icons.py               regenerates icons (Pillow)
test/                         node --test suites
.github/workflows/deploy.yml  GitHub Pages deploy
```

### 4.1 URL layout
Dev and production share one layout. `/` serves `src/*` (so `/taskpane.html`, `/app.js`, `/help.html`), `/assets/*` serves icons and `/dev/*` serves the preview (dev only). The dev server also serves `/manifest.xml`. In production the build copies `src/*` to `dist/` and `assets/` to `dist/assets/`, and writes `dist/word-count-manifest.xml`.

### 4.2 `count.js`
```js
export const ZERO = Object.freeze({ words: 0, characters: 0, charactersWithSpaces: 0 });
export function countText(text, { hangulPerSyllable = false } = {}) → { words, characters, charactersWithSpaces }
export function sumCounts(countsArray) → counts
export function countChunks(strings) → counts   // sum of countText over each string
```

### 4.3 `selection.js`
```js
/** @returns {Promise<SelectionSnapshot>} */
export async function readSelection(context)
SelectionSnapshot = {
  source: "highlight" | "shapes" | "none",
  selectedCount: number,                    // shapes selected, including pictures (for "3 shapes selected")
  items: Array<{ id: string, name: string, kind: "highlight" | "text" | "table" | "group", texts: string[] }>,
  unsupported: Array<{ id: string, name: string, type: string }>,   // charts, SmartArt, diagrams, OLE
}
```
Algorithm. It needs PowerPointApi 1.10 and uses only documented calls:
1. **Sync 1:** `presentation.getSelectedTextRangeOrNullObject().load("text")` and `presentation.getSelectedShapes().load("items/id,items/name,items/type,items/level")`. Never use the throwing `getSelectedTextRange()`, and never navigate to the range's parent (that fails for table cells).
2. **Highlight rule:** if the range is not null, its text is non-empty and at most one shape is selected, return `source: "highlight"` with one item (named after the selected shape if there is one).
3. **Shapes:** walk the selection breadth-first, **one sync per group-nesting level plus one for text**. Track a `seen` set of shape ids so a group and a child selected together are never counted twice. Each descendant's texts go to its top-level selected shape's item. Dispatch on `type`:
   - `Group` → `shape.group.shapes.load("items/id,items/name,items/type,items/level")` (only after type is known).
   - `Table` → `shape.getTable().load("rowCount,columnCount")`, then `table.getCellOrNullObject(r, c).load("text")` for every cell in the next round. Each cell is one chunk. Cells hidden under a merge are null objects (documented) and are skipped. `Table.values` isn't used because what it reports for merged-over cells is undocumented.
   - `Chart`, `SmartArt`, `Diagram`, `Ole` → recorded in `unsupported`.
   - `Image`, `Line`, `Media`, `Model3D`, `Ink`, `ContentApp`, `Graphic` → no text; ignored.
   - anything else (TextBox, Placeholder, GeometricShape, Callout, Freeform, Unsupported…) → `shape.getTextFrameOrNullObject().load("hasText")`, then in a follow-up sync `textFrame.textRange.load("text")` for non-null frames with text. For `Placeholder`, also load `placeholderFormat.containedType`. If the frame is null and `containedType === "Table"` (or `type === "Unsupported"`), try `getTable()` in its own guarded sync. If that also fails, record it as unsupported.
   - Never touch `shape.textFrame`, `group`, `parentGroup` or `placeholderFormat` on a shape whose type doesn't allow it. Each of these throws and kills the whole batch.
4. **Failure isolation:** every unit of work is a job with `queue()` and `read()`. If a batched `context.sync()` rejects, each job in that batch is re-queued and synced **on its own** inside try/catch. A job that still fails is skipped. It is recorded as unsupported if it is a shape; otherwise it is ignored.

### 4.4 `app.js` (controller)
- `start({ Office, PowerPoint, root, info })`. If the host isn't PowerPoint, show "Open this add-in from PowerPoint". If `!isSetSupported("PowerPointApi", "1.10")`, show "Please update PowerPoint (needs Mac 16.105 / Windows 2601 or newer)".
- Subscribes to `Office.EventType.DocumentSelectionChanged` via `Office.context.document.addHandlerAsync`.
- **Refresh engine:** triggers (selection event, focus, visibility, Refresh button, timer) call `schedule()`. It debounces by 120 ms and runs **one `PowerPoint.run` at a time**. If a trigger arrives mid-run, the run is marked dirty and repeats once afterwards. Results carry a generation number, and stale results are dropped.
- **Typing:** DocumentSelectionChanged isn't documented to fire on keystrokes. On desktop (`platform` Mac/PC), a 1.5 s light re-read runs while the pane is visible and idle. It is **disabled on PowerPoint on the web**, where read-only syncs have triggered autosave.
- **Errors:** a failed run retries once after 300 ms. That covers a known web bug where the first batch after `onReady` is "trampled". If the retry also fails, the last good numbers stay and a subtle "Couldn't read the selection, Retry" status appears.
- **Spurious empties:** PowerPoint occasionally reports an empty selection for a moment. When a read comes back empty right after a non-empty one, it is confirmed by one more read 150 ms later before the empty state is shown.
- The scheduling logic (debounce, single-flight, dirty rerun, generation, retry, empty confirmation) lives in `src/refresher.js`. It takes injected timers so it can be unit-tested without a DOM or Office.
- **Rendering:** three stat rows with `toLocaleString()` numbers and tabular digits, the context line, the breakdown (≥2 items), the unsupported note, the empty state, and a footer ("Counted the way Microsoft Word counts", Refresh, Help).
- **Theme:** reads `Office.context.officeTheme` defensively (it can be undefined or `{}`) at start and on focus. Valid hex colors set CSS variables, and `isDarkTheme` (or background luminance) sets `data-theme`. Without a usable theme it falls back to `prefers-color-scheme`.

### 4.5 Manifest
TaskPaneApp, XSD element order as validated. Dev `Id` `25c6e699-85a5-4a2d-bc50-ecc66e7483be`, production `Id` `1b331966-f37d-4165-9499-7f94f173c9c1`. Host `Presentation`. `<Set Name="PowerPointApi" MinVersion="1.10"/>`. `<Permissions>ReadWriteDocument</Permissions>`, which PowerPoint.run and addHandlerAsync require even though the add-in never writes. `IconUrl` 32 px, `HighResolutionIconUrl` 64 px, `SupportUrl` `/help.html`. VersionOverrides 1.0: GetStarted strings, `TabHome` → one group "Word Count" → one `ShowTaskpane` button (`TaskpaneId` `WordCountTaskpane`) with 16/32/80 icons (plus 20/24/40/48/64). No FunctionFile.

### 4.6 Build and hosting
`npm run build` needs `BASE_URL`, for example `https://joezandstra.github.io/ppt-word-count` (a trailing slash is stripped). It:
- copies `src/` and `assets/` to `dist/`;
- writes `dist/word-count-manifest.xml` from `manifest.xml` by replacing `https://localhost:3100` with BASE_URL, the dev Id with the production Id, and "Word Count (dev)" with "Word Count";
- adds `?v=<content-hash>` to every local `<script src>`/`<link href>` in the HTML and to every relative `import` in the JS. GitHub Pages caches files for 10 minutes and this can't be changed, so without the hash a deploy could briefly mix old and new files.

The GitHub Actions workflow (`checkout@v7`, `setup-node@v7` with Node 24, `configure-pages@v6`, `upload-pages-artifact@v5`, `deploy-pages@v5`) runs `npm ci`, `npm test` and `npm run build` with `BASE_URL = steps.pages.outputs.base_url`, then deploys `dist/`. Pages must be enabled with source "GitHub Actions". The repo must be public on GitHub Free.

Sharing options for colleagues, documented in the help page and README:
- (a) Microsoft 365 admin, Integrated apps, upload custom app by manifest URL. This is best when the team shares a work tenant; it can take up to 24 h to appear.
- (b) Mac: copy the manifest into `~/Library/Containers/com.microsoft.Powerpoint/Data/Documents/wef`.
- (c) Windows: Home > Add-ins > More add-ins > My Add-ins > Upload.
- (d) Web: Home > Add-ins > More Settings > Upload My Add-in.

Web-only updates reach users automatically. Manifest changes need a `<Version>` bump and a reinstall.

## 5. Known platform limitations (shown in help)
1. On PowerPoint for Mac and Windows, with the cursor inside a word but nothing highlighted, PowerPoint reports the whole word as "selected" (office-js #6839). The pane then counts that word instead of the whole box. This will be verified on 16.113; mitigate if possible.
2. With the slide-thumbnail pane hidden, PowerPoint may not report highlighted text (#6813). The whole shape is counted instead.
3. Charts, SmartArt and embedded objects can't be read.
4. Shapes on the slide master/layouts aren't reported by PowerPoint's selection API.
5. Partial highlights inside table cells may not be reported, depending on the platform. The whole table is counted instead.

## 6. Testing
- `test/count.test.js`: Word-verified vectors (≥60) plus the full measured fixture (`test/fixtures/word-measurements.json`, the rows measured in fresh-document conditions), with the deliberate deviations asserted separately.
- `test/selection.test.js`: every selection scenario against `dev/fake-powerpoint.js`. The fake enforces load-before-read, throws at sync for textFrame on tables/groups, group on non-groups and getTable on non-tables, and supports injected transient failures.
- `test/refresher.test.js`: refresh engine behaviour (debounce, single-flight, dirty rerun, stale drop, retry, empty confirmation) with fake timers.
- `test/build.test.js`: build with a sample BASE_URL; asserts manifest URLs, Id and name swapped, no `localhost` left, cache-busting applied, every referenced file present.
- `npm run validate`: Microsoft's manifest validation service (online).
- Browser preview (`npm run preview`, `/dev/preview.html`): every scenario in light and dark at 320 px wide.
- Live in PowerPoint for Mac 16.113 once the dev certificate is renewed: text box, highlight, caret, multi-select, table with merged cells, nested group, chart.
