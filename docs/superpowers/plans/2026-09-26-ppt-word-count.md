# PowerPoint Word Count Add-in Implementation Plan

> **For agentic workers:** REQUIRED SUB-SKILL: Use superpowers:subagent-driven-development (recommended) or superpowers:executing-plans to implement this plan task-by-task. Steps use checkbox (`- [ ]`) syntax for tracking.

**Goal:** Build a PowerPoint task-pane add-in that shows live Words / Characters (no spaces) / Characters (with spaces) for the selected text, shapes, tables and groups, counted the way Microsoft Word counts, and make it shareable through GitHub Pages.

**Architecture:** Plain HTML/CSS/ES-module Office.js add-in with no bundler or framework. A pure counting module, a single PowerPoint API reader (`selection.js`), a pure refresh scheduler, and a dependency-injected UI controller (`app.js`). An in-memory fake of the PowerPoint API drives both the unit tests and a browser preview page. A Node build script produces `dist/` and a production manifest, and a GitHub Actions workflow deploys it to Pages.

**Tech Stack:** Office.js (PowerPointApi 1.10, XML add-in-only manifest), vanilla JS (ES2022 modules), CSS custom properties, Node ≥20 built-ins only (`node:test`, `node:fs`, `node:crypto`), dev tools `office-addin-dev-certs` / `office-addin-debugging` / `office-addin-manifest`, GitHub Actions + Pages.

**Spec:** `docs/superpowers/specs/2026-09-26-ppt-word-count-design.md`

## Global Constraints

- Requirement set: `PowerPointApi` **1.10** (manifest `<Set Name="PowerPointApi" MinVersion="1.10"/>`, runtime `isSetSupported("PowerPointApi", "1.10")`).
- Manifest permission: `ReadWriteDocument`. PowerPoint.run and addHandlerAsync need it; the add-in never writes.
- Office.js URL, in `<head>`: `https://officeapis.public.onecdn.static.microsoft/1/office.js`.
- Dev origin: `https://localhost:3100`. Dev Id `25c6e699-85a5-4a2d-bc50-ecc66e7483be`, display name `Word Count (dev)`. Production Id `1b331966-f37d-4165-9499-7f94f173c9c1`, display name `Word Count`, file `word-count-manifest.xml`.
- URL layout (dev and production alike): `/` = files of `src/` (`/taskpane.html`, `/app.js`, `/help.html`…), `/assets/*` = icons, `/dev/*` = preview (dev only).
- No runtime dependencies. No bundler. No framework. Tests use `node:test` and `node:assert/strict` only.
- Never use `innerHTML` with anything derived from the document (shape names, text). Use `textContent`.
- Never call `shape.textFrame`, `presentation.getSelectedTextRange()`, `shape.group`, `shape.placeholderFormat` or `getTable()` unless the shape's `type` has been synced and allows it.
- Count every text container separately and add the results; never concatenate strings before counting.
- User-facing copy: plain language, sentence case, no jargon. Stat labels exactly: `Words`, `Characters (no spaces)`, `Characters (with spaces)`.
- Code style: 2-space indent, double quotes, semicolons, ES modules, JSDoc typedefs for shared shapes. Comments explain *why*, briefly.
- Do not `git commit` inside tasks run in parallel by subagents. The orchestrator commits after review. Sequential executors may commit per task.

## File map

| File | Responsibility | Task |
|---|---|---|
| `src/count.js` | `countText`, `sumCounts`, `countChunks`, `ZERO` (Word rules) | 1 |
| `test/count.test.js` | vectors + Word measurement fixture | 1 |
| `test/fixtures/word-measurements.json` | 432 Word-measured rows + 6 deliberate deviations (already committed) | 1 |
| `dev/fake-powerpoint.js` | `createFakeHost()`: fake `Office` + `PowerPoint` with real load/sync rules | 2 |
| `dev/scenarios.js` | `SCENARIOS`: named selections for tests and the preview | 2 |
| `test/fake-powerpoint.test.js` | proves the fake enforces the rules | 2 |
| `src/selection.js` | `readSelection(context)` → `SelectionSnapshot`, `tableCellTexts` | 3 |
| `test/selection.test.js` | every selection scenario | 3 |
| `src/refresher.js` | `createRefresher()` (debounce, single-flight, retry, empty confirm, poll) | 4 |
| `test/refresher.test.js` | scheduler behaviour with mock timers | 4 |
| `src/app.js` | `start()`, `summarize()`, `describeSelection()`, `unsupportedNote()`, `themeFrom()`, view | 5 |
| `src/taskpane.html`, `src/taskpane.css`, `src/taskpane.js` | pane page, styles, bootstrap | 5 |
| `dev/preview.html` | browser preview with scenario picker | 5 |
| `scripts/serve.mjs` (modify) | URL layout mapping | 5 |
| `test/app.test.js` | pure functions of app.js | 5 |
| `manifest.xml` | dev manifest | 6 |
| `src/help.html` | help, limitations, install guide | 6 |
| `test/manifest.test.js` | manifest sanity checks | 6 |
| `scripts/build.mjs` | `build({ baseUrl, root, outDir })` → dist + production manifest | 7 |
| `test/build.test.js` | build output checks | 7 |
| `.github/workflows/deploy.yml` | Pages deploy | 7 |
| `README.md` | user guide + maintainer notes | 8 |

## Shared interfaces (all tasks rely on these exact names)

```js
/** @typedef {{ words: number, characters: number, charactersWithSpaces: number }} Counts */
// src/count.js
export const ZERO;                                            // frozen Counts of zeros
export function countText(text, { hangulPerSyllable = false } = {}) /* → Counts */;
export function sumCounts(list /* Counts[] */) /* → Counts */;
export function countChunks(texts /* string[] */, options) /* → Counts */;

/**
 * @typedef {{ id: string, name: string, kind: "highlight" | "text" | "table" | "group", texts: string[] }} SelectionItem
 * @typedef {{ id: string, name: string, type: string }} UnsupportedShape
 * @typedef {{ source: "highlight" | "shapes" | "none", selectedCount: number, items: SelectionItem[], unsupported: UnsupportedShape[] }} SelectionSnapshot
 */
// src/selection.js
export async function readSelection(context) /* → SelectionSnapshot */;
export function tableCellTexts(cells /* Array<{ isNullObject: boolean, text: string }> */) /* → string[] */;

// src/refresher.js
export function createRefresher({ read, onResult, onError, isEmpty, timers, debounceMs, retryMs, confirmEmptyMs })
  /* → { schedule(): void, refreshNow(): Promise<void> | undefined, poll(intervalMs, shouldPoll): void, stop(): void } */;

/**
 * @typedef {{ id: string, name: string, kind: SelectionItem["kind"], counts: Counts }} SummaryItem
 * @typedef {{ source: SelectionSnapshot["source"], selectedCount: number, items: SummaryItem[], unsupported: UnsupportedShape[], total: Counts }} Summary
 */
// src/app.js
export function start({ Office, PowerPoint, root, info }) /* → { refresher, view } | null */;
export function summarize(snapshot) /* → Summary */;
export function describeSelection(summary) /* → { title: string, detail: string, note: string, empty: boolean } */;
export function unsupportedNote(count) /* → string */;
export function themeFrom(officeTheme) /* → { background: string, foreground: string, dark: boolean } | null */;

// dev/fake-powerpoint.js
export class FakeOfficeError extends Error { code }
export function createFakeHost(scenario, { platform = "Mac", apiVersion = "1.10", theme = null, latencyMs = 0 } = {})
  /* → { Office, PowerPoint, host } ; host: { scenario, stats: { runs, syncs, forbidden: string[] }, setScenario(s), failNextSyncs(n), fireSelectionChanged(), latencyMs } */;
// dev/scenarios.js
export const SCENARIOS; // { [key]: { label: string, highlight: string | null, selected: FakeShape[] } }

// scripts/build.mjs
export async function build({ baseUrl, root, outDir }) /* → { outDir: string, version: string, manifestPath: string } */;
```

---

### Task 1: Counting engine

**Files:**
- Create: `src/count.js`
- Create: `test/count.test.js`
- Use (exists): `test/fixtures/word-measurements.json`. Format: `{ description, rows: [{ id, text, words, characters, charactersWithSpaces, hangulPerSyllable? }], deviations: [{ ...row, reason }] }`

**Interfaces:**
- Consumes: nothing.
- Produces: `ZERO`, `countText`, `sumCounts`, `countChunks` exactly as in *Shared interfaces*.

- [ ] **Step 1: Write the failing test** `test/count.test.js`

```js
import assert from "node:assert/strict";
import { readFileSync } from "node:fs";
import { test } from "node:test";

import { countChunks, countText, sumCounts, ZERO } from "../src/count.js";

const fixture = JSON.parse(readFileSync(new URL("./fixtures/word-measurements.json", import.meta.url), "utf8"));

const counts = (words, characters, charactersWithSpaces) => ({ words, characters, charactersWithSpaces });

// [input, expected, why]. All verified against Microsoft Word for Mac 16.113.2
// unless marked "deviation".
const VECTORS = [
  ["Hello world", counts(2, 10, 11), "basic"],
  ["Hello  world", counts(2, 10, 12), "each space counts"],
  ["  Hello world  ", counts(2, 10, 15), "leading/trailing spaces count"],
  ["a\tb", counts(2, 2, 3), "tab is a space"],
  ["a\rb", counts(2, 2, 2), "paragraph mark not counted"],
  ["a\nb", counts(2, 2, 2), "LF is a paragraph"],
  ["a\r\nb", counts(2, 2, 2), "CRLF is one paragraph"],
  ["a\u000Bb", counts(2, 2, 3), "Shift+Enter line break counts as a space"],
  ["word word", counts(2, 8, 9), "NBSP separates and counts as a space"],
  ["a b", counts(1, 3, 3), "thin space is not a space for Word"],
  ["a　b", counts(2, 2, 3), "ideographic space is a space"],
  ["word - word", counts(3, 9, 11), "spaced hyphen is a word"],
  ["word & word", counts(3, 9, 11), "standalone symbol is a word"],
  ["word — word", counts(2, 9, 11), "spaced em dash is not a word but is a character"],
  ["word—word", counts(2, 9, 9), "em dash splits words"],
  ["word–word", counts(2, 9, 9), "en dash splits words"],
  ["well-known", counts(1, 10, 10), "hyphen joins"],
  ["e-mail", counts(1, 6, 6), "hyphen joins"],
  ["don’t", counts(1, 5, 5), "apostrophe joins"],
  ["3.14", counts(1, 4, 4), "decimal"],
  ["U.S.A.", counts(1, 6, 6), "abbreviation"],
  ["https://example.com/a/b?x=1&y=2", counts(1, 31, 31), "URL is one word"],
  ["a/b", counts(1, 3, 3), "slash joins"],
  [". . .", counts(3, 3, 5), "each dot is a word"],
  ["The court … concluded", counts(4, 18, 21), "ellipsis character is a word"],
  ["—", counts(0, 1, 1), "lone em dash"],
  ["word--word", counts(2, 10, 10), "double hyphen breaks"],
  ["--word", counts(2, 6, 6), "leading double hyphen"],
  ["• item", counts(1, 5, 6), "bullet character is a separator"],
  ["word−word", counts(1, 9, 9), "minus sign joins"],
  ["word​word", counts(1, 8, 8), "zero-width space ignored"],
  ["word­word", counts(1, 9, 9), "literal soft hyphen counted"],
  ["\"Remember, we—\"", counts(3, 14, 15), "closing quote after em dash is a word"],
  ["世界", counts(2, 2, 2), "each CJK character is a word"],
  ["Hello世界", counts(3, 7, 7), "CJK breaks a Latin run"],
  ["你好，世界。", counts(6, 6, 6), "full-width punctuation are words"],
  ["Hi， this is a test", counts(6, 14, 18), "full-width comma"],
  ["世界,", counts(3, 3, 3), "ASCII comma after CJK is a word"],
  ["日本語123", counts(4, 6, 6), "digit run is one word"],
  ["ｶﾀｶﾅ", counts(4, 4, 4), "half-width katakana per character"],
  ["ＡＢＣ", counts(3, 3, 3), "full-width Latin per character"],
  ["안녕하세요 세계", counts(2, 7, 8), "Korean words by default"],
  ["สวัสดีครับ", counts(1, 10, 10), "Thai is not segmented"],
  ["\u{1F44D}", counts(1, 1, 1), "emoji is one character"],
  ["\u{1F44D}\u{1F3FD}", counts(1, 2, 2), "skin tone adds one"],
  ["\u{1F468}‍\u{1F469}‍\u{1F467}", counts(1, 5, 5), "joiners inside emoji count"],
  ["\u{1F1FA}\u{1F1F8}", counts(1, 2, 2), "flag is two"],
  ["❤️", counts(1, 1, 1), "variation selector ignored"],
  ["Hi\u{1F44D}", counts(1, 3, 3), "emoji joins a word"],
  ["café", counts(1, 5, 5), "combining accent counted"],
  [" ", counts(0, 0, 1), "lone space"],
  ["​", counts(0, 0, 0), "lone zero-width space"],
  ["", counts(0, 0, 0), "empty"],
  ["The quick brown fox — it jumped – over 3.5 lazy dogs.", counts(10, 42, 53), "mixed dashes"],
  ["\t•\tBullet text", counts(2, 11, 14), "typed bullet with tabs"],
  ["Title\rBody text here", counts(4, 17, 19), "paragraphs"],
  ["Line one\u000BLine two", counts(4, 14, 17), "soft line break"],
  ["\"This — as expected — irritated.\"", counts(4, 28, 33), "quoted with em dashes"],
  ["a b", counts(2, 2, 3), "deviation: U+2028 is a line break (Word: 3/3/3)"],
  ["a b", counts(2, 2, 2), "deviation: U+2029 is a paragraph break (Word: 3/3/3)"],
  ["a※b", counts(3, 3, 3), "reference mark is East Asian (fresh-document Word)"],
];

for (const [input, expected, why] of VECTORS) {
  test(`countText ${JSON.stringify(input)} (${why})`, () => {
    assert.deepEqual(countText(input), expected);
  });
}

test("counts Hangul per syllable when asked (Word's untagged behaviour)", () => {
  assert.deepEqual(countText("안녕하세요 세계", { hangulPerSyllable: true }), counts(7, 7, 8));
});

test("word count for a sentence with dash variants matches Word", () => {
  const text = "How about some dash stuff–things like hyphen-dashes and em-dashes. 1-248-434-5508.";
  assert.equal(countText(text).words, 11);
});

test(`matches all ${fixture.rows.length} Word-measured rows`, () => {
  const failures = [];
  for (const row of fixture.rows) {
    const actual = countText(row.text, { hangulPerSyllable: row.hangulPerSyllable === true });
    const expected = counts(row.words, row.characters, row.charactersWithSpaces);
    if (JSON.stringify(actual) !== JSON.stringify(expected)) {
      failures.push(`${row.id} ${JSON.stringify(row.text)}: expected ${JSON.stringify(expected)}, got ${JSON.stringify(actual)}`);
    }
  }
  assert.deepEqual(failures, []);
});

test("each deliberate deviation really differs from Word (so the list stays honest)", () => {
  for (const row of fixture.deviations) {
    const actual = countText(row.text, { hangulPerSyllable: row.hangulPerSyllable === true });
    assert.notDeepEqual(actual, counts(row.words, row.characters, row.charactersWithSpaces), row.id);
  }
});

test("sumCounts adds field by field and returns a new object", () => {
  const total = sumCounts([counts(1, 2, 3), counts(10, 20, 30)]);
  assert.deepEqual(total, counts(11, 22, 33));
  assert.deepEqual(sumCounts([]), ZERO);
  assert.notEqual(sumCounts([]), ZERO);
});

test("countChunks never lets words run together across chunks", () => {
  assert.deepEqual(countChunks(["Hello", "world"]), counts(2, 10, 10));
  assert.deepEqual(countChunks([]), ZERO);
});

test("ZERO is frozen", () => {
  assert.ok(Object.isFrozen(ZERO));
});
```

- [ ] **Step 2: Run test to verify it fails**

Run: `node --test test/count.test.js`
Expected: FAIL with `Cannot find module '.../src/count.js'`.

- [ ] **Step 3: Write the implementation** `src/count.js`

```js
// Word-compatible counting of words and characters.
//
// The rules reproduce Microsoft Word's own statistics, measured on Word for Mac
// 16.113.2 (docs/superpowers/specs/2026-09-26-ppt-word-count-design.md, section 3).
// Counting walks Unicode code points: String.length, Intl.Segmenter and regex
// splitting all disagree with Word on everyday text (hyphens, URLs, CJK, emoji).

/** @typedef {{ words: number, characters: number, charactersWithSpaces: number }} Counts */

/** @type {Readonly<Counts>} */
export const ZERO = Object.freeze({ words: 0, characters: 0, charactersWithSpaces: 0 });

// Paragraph breaks end a word and are never counted.
const PARAGRAPH = new Set([0x0a, 0x0c, 0x0d, 0x0e, 0x2029]);
// Spaces end a word and are counted only in "characters with spaces".
const SPACE = new Set([0x09, 0x0b, 0x20, 0xa0, 0x2028, 0x3000]);
// En dash, em dash and bullet end a word and count as characters, but are never words themselves.
const DASH_SEPARATOR = new Set([0x2013, 0x2014, 0x2022]);

const HYPHEN = 0x2d;
const ZERO_WIDTH_JOINER = 0x200d;

// Each of these characters is a word on its own, like Word's "Asian characters".
const EAST_ASIAN = new RegExp(
  "[" +
    "\\p{Script=Han}\\p{Script=Hiragana}\\p{Script=Katakana}" +
    "\\u3100-\\u312F" + // Bopomofo (Bopomofo Extended is not East Asian to Word)
    "\\u2024\\u2025\\u2027\\u203B" + // one/two dot leaders, hyphenation point, reference mark
    "\\u2160-\\u217F" + // Roman numerals
    "\\u2460-\\u24B5\\u24D0-\\u24FE" + // circled and parenthesized numbers, circled small letters
    "\\u2E80-\\u2FDF" + // CJK and Kangxi radicals
    "\\u3001-\\u303F" + // CJK symbols and punctuation
    "\\u3099-\\u309C\\u30A0\\u30FB\\u30FC" + // kana marks, middle dot, prolonged sound mark
    "\\u3190-\\u319F" + // Kanbun
    "\\u3200-\\u33FF" + // enclosed CJK letters, CJK compatibility
    "\\uFE30-\\uFE4F" + // CJK compatibility forms
    "\\uFE50-\\uFE6B" + // small form variants
    "\\uFF01-\\uFFEF" + // half-width and full-width forms
    "]",
  "u",
);
const HANGUL = /\p{Script=Hangul}/u;
// A zero-width joiner after one of these belongs to an emoji sequence and is counted.
const EMOJI_BASE = /[\p{Extended_Pictographic}\p{Emoji_Modifier}]/u;

// Invisible format characters that are neither counted nor word breaks.
function isIgnorable(cp) {
  return (
    cp === 0x200b || // zero-width space
    cp === 0x200c || // zero-width non-joiner
    cp === ZERO_WIDTH_JOINER ||
    cp === 0xfeff || // byte order mark
    cp === 0x1f || // Word's optional hyphen
    (cp >= 0xfe00 && cp <= 0xfe0f) || // variation selectors
    (cp >= 0xe0100 && cp <= 0xe01ef) || // variation selectors supplement
    (cp >= 0xe0000 && cp <= 0xe007f) // tags
  );
}

/**
 * Counts one piece of text the way Microsoft Word does.
 * @param {string} text
 * @param {{ hangulPerSyllable?: boolean }} [options] Count each Hangul syllable as a
 *   word (Word's behaviour for text not tagged as Korean). Default: space-delimited Korean words.
 * @returns {Counts}
 */
export function countText(text, { hangulPerSyllable = false } = {}) {
  let words = 0;
  let withSpaces = 0;
  let spaces = 0;
  let inWord = false;
  let breakAfterHyphens = false; // after "--", the next non-hyphen starts a new word
  let previous = -1; // previous code point in the current word
  let previousVisible = -1; // previous non-ignorable code point, for the joiner rule

  const endWord = () => {
    inWord = false;
    breakAfterHyphens = false;
    previous = -1;
  };

  for (const ch of text) {
    const cp = ch.codePointAt(0);
    if (cp === ZERO_WIDTH_JOINER && previousVisible !== -1 && EMOJI_BASE.test(String.fromCodePoint(previousVisible))) {
      withSpaces++;
      continue;
    }
    if (isIgnorable(cp)) continue;
    previousVisible = cp;

    if (PARAGRAPH.has(cp)) {
      endWord();
      continue;
    }
    withSpaces++;
    if (SPACE.has(cp)) {
      spaces++;
      endWord();
      continue;
    }
    if (DASH_SEPARATOR.has(cp)) {
      endWord();
      continue;
    }
    if (EAST_ASIAN.test(ch) || (hangulPerSyllable && HANGUL.test(ch))) {
      words++;
      endWord();
      continue;
    }
    if (breakAfterHyphens && cp !== HYPHEN) endWord();
    if (!inWord) {
      words++;
      inWord = true;
    }
    if (cp === HYPHEN && previous === HYPHEN) breakAfterHyphens = true;
    previous = cp;
  }

  return { words, characters: withSpaces - spaces, charactersWithSpaces: withSpaces };
}

/**
 * Adds counts together.
 * @param {Counts[]} list
 * @returns {Counts}
 */
export function sumCounts(list) {
  const total = { ...ZERO };
  for (const c of list) {
    total.words += c.words;
    total.characters += c.characters;
    total.charactersWithSpaces += c.charactersWithSpaces;
  }
  return total;
}

/**
 * Counts several separate pieces of text (text boxes, table cells...) and adds the
 * results. Pieces are never joined, so words can't run together across them.
 * @param {string[]} texts
 * @param {{ hangulPerSyllable?: boolean }} [options]
 * @returns {Counts}
 */
export function countChunks(texts, options) {
  return sumCounts(texts.map((text) => countText(text, options)));
}
```

- [ ] **Step 4: Run test to verify it passes**

Run: `node --test test/count.test.js`
Expected: PASS (all tests). If a VECTORS row fails, compare against `test/fixtures/word-measurements.json` (a Word measurement); the fixture wins over the plan.

- [ ] **Step 5: Commit** (sequential executors only)

```bash
git add src/count.js test/count.test.js
git commit -m "feat: add Word-compatible word and character counting"
```

---

### Task 2: Fake PowerPoint host and scenarios

**Files:**
- Create: `dev/fake-powerpoint.js`
- Create: `dev/scenarios.js`
- Create: `test/fake-powerpoint.test.js`

**Interfaces:**
- Consumes: nothing.
- Produces: `createFakeHost`, `FakeOfficeError` (see *Shared interfaces*) and `SCENARIOS`. Fake shapes are plain objects:
  - text shape: `{ id, name, type, text }`. `text: ""` is an empty text frame; omit `text` for "no text frame".
  - table: `{ id, name, type: "Table", rows: string[][], merged?: [{ row, column, rowCount, columnCount }] }`
  - group: `{ id, name, type: "Group", shapes: FakeShape[] }`
  - table inside a content placeholder: `{ id, name, type: "Placeholder", containedType: "Table", rows }`
  - any shape + `failTextFrame: true`: its `getTextFrameOrNullObject()` load fails the batch.
- Supported fake API surface: `PowerPoint.run(cb)`, `context.sync()`, `context.presentation.getSelectedShapes()`, `.getSelectedTextRangeOrNullObject()`, `.getSelectedTextRange()` (recorded as forbidden), `ShapeCollection.items`, `Shape.id/name/type/level`, `Shape.textFrame` (forbidden, throws on non-text), `Shape.getTextFrameOrNullObject()`, `Shape.getTable()`, `Shape.group`, `Shape.placeholderFormat`, `TextFrame.hasText/textRange`, `TextRange.text`, `Table.rowCount/columnCount/values`, `Table.getCellOrNullObject(r, c)`, `TableCell.text`, `ShapeGroup.id/shapes`, `PlaceholderFormat.containedType/type`, every proxy's `isNullObject`, and `load()` with `"a,b"`, `["a","b"]`, `"items/x"` and `"nav/prop"` paths.

- [ ] **Step 1: Write the failing test** `test/fake-powerpoint.test.js`

```js
import assert from "node:assert/strict";
import { test } from "node:test";

import { createFakeHost, FakeOfficeError } from "../dev/fake-powerpoint.js";
import { SCENARIOS } from "../dev/scenarios.js";

const textBox = { id: "1", name: "TextBox 1", type: "TextBox", text: "Hello world" };
const table = {
  id: "2",
  name: "Table 1",
  type: "Table",
  rows: [
    ["A", "B"],
    ["merged", ""],
  ],
  merged: [{ row: 1, column: 0, rowCount: 1, columnCount: 2 }],
};

test("properties must be loaded and synced before they're read", async () => {
  const { PowerPoint } = createFakeHost({ highlight: null, selected: [textBox] });
  await PowerPoint.run(async (context) => {
    const shapes = context.presentation.getSelectedShapes();
    assert.throws(() => shapes.items, { code: "PropertyNotLoaded" });
    shapes.load("items/id,items/name,items/type");
    assert.throws(() => shapes.items, { code: "PropertyNotLoaded" });
    await context.sync();
    assert.equal(shapes.items.length, 1);
    assert.equal(shapes.items[0].name, "TextBox 1");
    assert.throws(() => shapes.items[0].level, { code: "PropertyNotLoaded" });
  });
});

test("an invalid navigation rejects the whole batch and applies nothing", async () => {
  const { PowerPoint } = createFakeHost({ highlight: null, selected: [textBox, table] });
  await PowerPoint.run(async (context) => {
    const shapes = context.presentation.getSelectedShapes();
    shapes.load("items/id,items/type");
    await context.sync();
    const [box, tbl] = shapes.items;
    const range = box.getTextFrameOrNullObject().textRange;
    range.load("text");
    tbl.textFrame.load("hasText"); // tables have no text frame
    await assert.rejects(context.sync(), (error) => error instanceof FakeOfficeError && error.code === "InvalidArgument");
    assert.throws(() => range.text, { code: "PropertyNotLoaded" });
    // The context keeps working after a failed batch.
    range.load("text");
    await context.sync();
    assert.equal(range.text, "Hello world");
  });
});

test("group, placeholderFormat and getTable fail on the wrong kind of shape", async () => {
  const { PowerPoint } = createFakeHost({ highlight: null, selected: [textBox] });
  await PowerPoint.run(async (context) => {
    const shapes = context.presentation.getSelectedShapes();
    shapes.load("items/type");
    await context.sync();
    const [box] = shapes.items;
    for (const load of [() => box.group.shapes.load("items/id"), () => box.placeholderFormat.load("containedType"), () => box.getTable().load("rowCount")]) {
      load();
      await assert.rejects(context.sync(), { code: "GeneralException" });
    }
  });
});

test("OrNullObject methods return null objects instead of failing", async () => {
  const picture = { id: "3", name: "Picture 1", type: "Image" };
  const { PowerPoint } = createFakeHost({ highlight: null, selected: [picture] });
  await PowerPoint.run(async (context) => {
    const range = context.presentation.getSelectedTextRangeOrNullObject();
    range.load("text");
    const shapes = context.presentation.getSelectedShapes();
    shapes.load("items/type");
    await context.sync();
    assert.equal(range.isNullObject, true);
    const frame = shapes.items[0].getTextFrameOrNullObject();
    frame.load("hasText");
    await context.sync();
    assert.equal(frame.isNullObject, true);
    // Navigating from a null object is undocumented in Office.js, so the fake refuses.
    frame.textRange.load("text");
    await assert.rejects(context.sync(), { code: "InvalidArgument" });
  });
});

test("table cells hidden under a merge are null objects; values repeats merged text", async () => {
  const { PowerPoint } = createFakeHost({ highlight: null, selected: [table] });
  await PowerPoint.run(async (context) => {
    const shapes = context.presentation.getSelectedShapes();
    shapes.load("items/type");
    await context.sync();
    const t = shapes.items[0].getTable();
    t.load("rowCount,columnCount,values");
    await context.sync();
    assert.equal(t.rowCount, 2);
    assert.equal(t.columnCount, 2);
    assert.deepEqual(t.values, [
      ["A", "B"],
      ["merged", "merged"],
    ]);
    const anchor = t.getCellOrNullObject(1, 0);
    const hidden = t.getCellOrNullObject(1, 1);
    anchor.load("text");
    hidden.load("text");
    await context.sync();
    assert.equal(anchor.isNullObject, false);
    assert.equal(anchor.text, "merged");
    assert.equal(hidden.isNullObject, true);
  });
});

test("group children are reachable through group.shapes", async () => {
  const group = { id: "9", name: "Group 1", type: "Group", shapes: [textBox, table] };
  const { PowerPoint } = createFakeHost({ highlight: null, selected: [group] });
  await PowerPoint.run(async (context) => {
    const shapes = context.presentation.getSelectedShapes();
    shapes.load("items/id,items/type");
    await context.sync();
    const children = shapes.items[0].group.shapes;
    children.load("items/id,items/name,items/type");
    await context.sync();
    assert.deepEqual(
      children.items.map((s) => s.name),
      ["TextBox 1", "Table 1"],
    );
  });
});

test("proxies can't be used after their PowerPoint.run has finished", async () => {
  const { PowerPoint } = createFakeHost({ highlight: null, selected: [textBox] });
  let leaked;
  await PowerPoint.run(async (context) => {
    leaked = context.presentation.getSelectedShapes();
  });
  await PowerPoint.run(async () => {
    assert.throws(() => leaked.load("items/id"), { code: "InvalidObjectPath" });
  });
});

test("records forbidden APIs and counts runs and syncs", async () => {
  const { PowerPoint, host } = createFakeHost({ highlight: null, selected: [textBox] });
  await PowerPoint.run(async (context) => {
    context.presentation.getSelectedTextRange();
    const shapes = context.presentation.getSelectedShapes();
    shapes.load("items/id");
    await context.sync();
    void shapes.items[0].textFrame;
  });
  assert.deepEqual(host.stats.forbidden, ["Presentation.getSelectedTextRange", "Shape.textFrame"]);
  assert.equal(host.stats.runs, 1);
  assert.equal(host.stats.syncs, 1);
});

test("failNextSyncs makes the next syncs fail like a transient host error", async () => {
  const { PowerPoint, host } = createFakeHost({ highlight: null, selected: [textBox] });
  host.failNextSyncs(1);
  await PowerPoint.run(async (context) => {
    const shapes = context.presentation.getSelectedShapes();
    shapes.load("items/id");
    await assert.rejects(context.sync(), { code: "GeneralException" });
    shapes.load("items/id");
    await context.sync();
    assert.equal(shapes.items.length, 1);
  });
});

test("Office fake: onReady, requirement sets, selection events", async () => {
  const { Office, host } = createFakeHost({ highlight: null, selected: [] }, { platform: "PC", apiVersion: "1.10" });
  assert.deepEqual(await Office.onReady(), { host: "PowerPoint", platform: "PC" });
  assert.equal(Office.context.requirements.isSetSupported("PowerPointApi", "1.10"), true);
  assert.equal(Office.context.requirements.isSetSupported("PowerPointApi", "1.8"), true);
  assert.equal(Office.context.requirements.isSetSupported("PowerPointApi", "1.11"), false);
  let fired = 0;
  let status;
  Office.context.document.addHandlerAsync(Office.EventType.DocumentSelectionChanged, () => fired++, (result) => (status = result.status));
  assert.equal(status, Office.AsyncResultStatus.Succeeded);
  host.fireSelectionChanged();
  assert.equal(fired, 1);
});

test("every scenario has a label and a selection", () => {
  assert.ok(Object.keys(SCENARIOS).length >= 10);
  for (const [key, scenario] of Object.entries(SCENARIOS)) {
    assert.equal(typeof scenario.label, "string", key);
    assert.ok(Array.isArray(scenario.selected), key);
    assert.ok(scenario.highlight === null || typeof scenario.highlight === "string", key);
  }
});
```

- [ ] **Step 2: Run test to verify it fails**

Run: `node --test test/fake-powerpoint.test.js`
Expected: FAIL with `Cannot find module '.../dev/fake-powerpoint.js'`.

- [ ] **Step 3: Write the implementation** `dev/fake-powerpoint.js`

```js
// In-memory stand-in for the parts of the Office.js PowerPoint API this add-in
// uses. The unit tests and the browser preview (dev/preview.html) run against it.
//
// It copies the behaviour of the real proxy model that can trip code up:
// - properties must be load()ed and synced before they can be read;
// - navigating to something a shape doesn't have (textFrame on a table, group on
//   a non-group, getTable() on a text box, placeholderFormat on a non-placeholder)
//   is only detected at context.sync(), and it rejects the whole batch;
// - *OrNullObject methods return an object whose isNullObject is true instead;
// - proxies stop working once the PowerPoint.run that created them has finished.
//
// Fake shapes are plain objects: see dev/scenarios.js for examples.

export class FakeOfficeError extends Error {
  constructor(code, message) {
    super(message);
    this.name = "RichApi.Error";
    this.code = code;
  }
}

const NULL = Symbol("null object");
const invalidArgument = () => new FakeOfficeError("InvalidArgument", "The argument is invalid or missing or has an incorrect format.");
const generalException = () => new FakeOfficeError("GeneralException", "There was an internal error while processing the request.");

class FakeProxy {
  constructor(context, resolve, { nullable = false } = {}) {
    this._context = context;
    this._resolve = resolve; // () => data or NULL; throws FakeOfficeError when invalid
    this._nullable = nullable;
    this._values = new Map();
  }

  load(spec = "") {
    this._context._assertOpen();
    const paths = (Array.isArray(spec) ? spec : String(spec).split(",")).map((p) => p.trim()).filter(Boolean);
    const own = [];
    const items = [];
    for (const path of paths) {
      const slash = path.indexOf("/");
      if (slash === -1) own.push(path);
      else if (path.slice(0, slash) === "items" && this instanceof FakeCollection) items.push(path.slice(slash + 1));
      else this[path.slice(0, slash)].load(path.slice(slash + 1));
    }
    this._context._queue.push({ proxy: this, own, items });
    return this;
  }

  get isNullObject() {
    return this._get("isNullObject");
  }

  _get(property) {
    if (!this._values.has(property)) {
      throw new FakeOfficeError(
        "PropertyNotLoaded",
        `The property '${property}' is not available. Before reading the property's value, call the load method on the containing object and call "context.sync()" on the associated request context.`,
      );
    }
    return this._values.get(property);
  }

  _apply(data, own) {
    this._values.set("isNullObject", data === NULL);
    for (const property of own) this._values.set(property, data === NULL ? null : this._read(data, property));
  }

  _read(_data, property) {
    throw new FakeOfficeError("InvalidArgument", `Unknown property '${property}'.`);
  }

  _child(key, create) {
    this[key] ??= create();
    return this[key];
  }
}

class FakeCollection extends FakeProxy {
  constructor(context, resolve, Item) {
    super(context, resolve);
    this._Item = Item;
  }

  get items() {
    return this._get("items");
  }

  _apply(list, own, items = []) {
    super._apply(list, own.filter((p) => p !== "items"));
    const props = own.includes("items") ? this._Item.SCALARS : items;
    this._values.set(
      "items",
      list.map((data) => {
        const item = new this._Item(this._context, () => data);
        item._apply(data, props);
        return item;
      }),
    );
  }
}

function hasTextFrame(shape) {
  return typeof shape.text === "string" && shape.type !== "Table" && shape.type !== "Group";
}

class FakeShape extends FakeProxy {
  static SCALARS = ["id", "name", "type", "level"];

  get id() {
    return this._get("id");
  }
  get name() {
    return this._get("name");
  }
  get type() {
    return this._get("type");
  }
  get level() {
    return this._get("level");
  }

  _read(shape, property) {
    if (property === "level") return shape.level ?? 0;
    if (FakeShape.SCALARS.includes(property)) return shape[property];
    return super._read(shape, property);
  }

  /** Forbidden in the add-in: throws for shapes without a text frame. */
  get textFrame() {
    this._context._host._forbidden("Shape.textFrame");
    return this._child("_textFrame", () =>
      new FakeTextFrame(this._context, () => {
        const shape = this._resolve();
        if (!hasTextFrame(shape)) throw invalidArgument();
        return shape;
      }),
    );
  }

  getTextFrameOrNullObject() {
    return new FakeTextFrame(
      this._context,
      () => {
        const shape = this._resolve();
        if (shape.failTextFrame) throw generalException();
        return hasTextFrame(shape) ? shape : NULL;
      },
      { nullable: true },
    );
  }

  getTable() {
    return new FakeTable(this._context, () => {
      const shape = this._resolve();
      if (!Array.isArray(shape.rows)) throw generalException();
      return shape;
    });
  }

  get group() {
    return this._child("_group", () =>
      new FakeShapeGroup(this._context, () => {
        const shape = this._resolve();
        if (shape.type !== "Group") throw generalException();
        return shape;
      }),
    );
  }

  get placeholderFormat() {
    return this._child("_placeholderFormat", () =>
      new FakePlaceholderFormat(this._context, () => {
        const shape = this._resolve();
        if (shape.type !== "Placeholder") throw generalException();
        return shape;
      }),
    );
  }
}

class FakeTextFrame extends FakeProxy {
  get hasText() {
    return this._get("hasText");
  }

  _read(shape, property) {
    if (property === "hasText") return shape.text.length > 0;
    return super._read(shape, property);
  }

  get textRange() {
    return this._child("_textRange", () =>
      new FakeTextRange(this._context, () => {
        const shape = this._resolve();
        // Office.js doesn't document navigating from a null object, so refuse it.
        if (shape === NULL) throw invalidArgument();
        return { text: shape.text };
      }),
    );
  }
}

class FakeTextRange extends FakeProxy {
  get text() {
    return this._get("text");
  }

  _read(range, property) {
    if (property === "text") return range.text;
    return super._read(range, property);
  }
}

function mergeCovering(shape, row, column) {
  return (shape.merged ?? []).find(
    (m) => row >= m.row && row < m.row + m.rowCount && column >= m.column && column < m.column + m.columnCount,
  );
}

class FakeTable extends FakeProxy {
  get rowCount() {
    return this._get("rowCount");
  }
  get columnCount() {
    return this._get("columnCount");
  }
  get values() {
    return this._get("values");
  }

  _read(shape, property) {
    if (property === "rowCount") return shape.rows.length;
    if (property === "columnCount") return Math.max(0, ...shape.rows.map((r) => r.length));
    if (property === "values") {
      // PowerPoint doesn't document what it reports for cells hidden under a
      // merge. The fake repeats the merged text there (the worst case), so any
      // code that counts from `values` double counts and fails its tests.
      return shape.rows.map((cells, r) =>
        cells.map((text, c) => {
          const merge = mergeCovering(shape, r, c);
          return merge ? shape.rows[merge.row][merge.column] : text;
        }),
      );
    }
    return super._read(shape, property);
  }

  getCellOrNullObject(row, column) {
    return new FakeTableCell(
      this._context,
      () => {
        const shape = this._resolve();
        if (row < 0 || row >= shape.rows.length || column < 0 || column >= (shape.rows[row]?.length ?? 0)) return NULL;
        const merge = mergeCovering(shape, row, column);
        if (merge && (merge.row !== row || merge.column !== column)) return NULL;
        return { text: shape.rows[row][column], rowIndex: row, columnIndex: column };
      },
      { nullable: true },
    );
  }
}

class FakeTableCell extends FakeProxy {
  get text() {
    return this._get("text");
  }

  _read(cell, property) {
    if (property in cell) return cell[property];
    return super._read(cell, property);
  }
}

class FakeShapeGroup extends FakeProxy {
  get id() {
    return this._get("id");
  }

  _read(shape, property) {
    if (property === "id") return shape.id;
    return super._read(shape, property);
  }

  get shapes() {
    return this._child("_shapes", () => new FakeCollection(this._context, () => this._resolve().shapes, FakeShape));
  }
}

class FakePlaceholderFormat extends FakeProxy {
  get containedType() {
    return this._get("containedType");
  }
  get type() {
    return this._get("type");
  }

  _read(shape, property) {
    if (property === "containedType") return shape.containedType ?? null;
    if (property === "type") return shape.placeholderType ?? "Body";
    return super._read(shape, property);
  }
}

class FakePresentation {
  constructor(context) {
    this._context = context;
  }

  getSelectedShapes() {
    return new FakeCollection(this._context, () => this._context._host.scenario.selected, FakeShape);
  }

  getSelectedTextRangeOrNullObject() {
    return new FakeTextRange(
      this._context,
      () => {
        const highlight = this._context._host.scenario.highlight;
        return highlight == null ? NULL : { text: highlight };
      },
      { nullable: true },
    );
  }

  /** Forbidden in the add-in: throws when no text is selected. */
  getSelectedTextRange() {
    this._context._host._forbidden("Presentation.getSelectedTextRange");
    return new FakeTextRange(this._context, () => {
      const highlight = this._context._host.scenario.highlight;
      if (highlight == null) throw generalException();
      return { text: highlight };
    });
  }
}

class FakeRequestContext {
  constructor(host) {
    this._host = host;
    this._queue = [];
    this.presentation = new FakePresentation(this);
  }

  _assertOpen() {
    if (this._closed) {
      throw new FakeOfficeError("InvalidObjectPath", "The object path isn't working for what you're trying to do. Objects can't be used after their PowerPoint.run has finished.");
    }
  }

  async sync() {
    const host = this._host;
    host.stats.syncs++;
    const queue = this._queue;
    this._queue = [];
    if (host.latencyMs > 0) await new Promise((resolve) => setTimeout(resolve, host.latencyMs));
    if (host._failures > 0) {
      host._failures--;
      throw new FakeOfficeError("GeneralException", "The property 'items' is not available. Before reading the property's value, call the load method.");
    }
    // Resolve everything first: one failure rejects the whole batch and applies nothing.
    const resolved = queue.map((op) => op.proxy._resolve());
    queue.forEach((op, i) => op.proxy._apply(resolved[i], op.own, op.items));
  }
}

function compareVersions(a, b) {
  const pa = String(a).split(".").map(Number);
  const pb = String(b).split(".").map(Number);
  for (let i = 0; i < Math.max(pa.length, pb.length); i++) {
    const diff = (pa[i] ?? 0) - (pb[i] ?? 0);
    if (diff !== 0) return Math.sign(diff);
  }
  return 0;
}

/**
 * Creates a fake Office + PowerPoint pair.
 * @param {{ highlight: string | null, selected: object[] }} scenario
 * @param {{ platform?: string, apiVersion?: string, theme?: object | null, latencyMs?: number }} [options]
 */
export function createFakeHost(scenario = { highlight: null, selected: [] }, { platform = "Mac", apiVersion = "1.10", theme = null, latencyMs = 0 } = {}) {
  const selectionHandlers = [];
  const host = {
    scenario,
    latencyMs,
    stats: { runs: 0, syncs: 0, forbidden: [] },
    _failures: 0,
    _forbidden(api) {
      host.stats.forbidden.push(api);
    },
    setScenario(next) {
      host.scenario = next;
    },
    failNextSyncs(count = 1) {
      host._failures = count;
    },
    fireSelectionChanged() {
      for (const handler of selectionHandlers) handler({ type: "documentSelectionChanged" });
    },
  };

  const PowerPoint = {
    async run(callback) {
      host.stats.runs++;
      const context = new FakeRequestContext(host);
      try {
        return await callback(context);
      } finally {
        context._closed = true;
      }
    },
  };

  const Office = {
    HostType: { PowerPoint: "PowerPoint" },
    PlatformType: { PC: "PC", OfficeOnline: "OfficeOnline", Mac: "Mac", iOS: "iOS", Android: "Android", Universal: "Universal" },
    EventType: { DocumentSelectionChanged: "documentSelectionChanged" },
    AsyncResultStatus: { Succeeded: "succeeded", Failed: "failed" },
    onReady: async () => ({ host: "PowerPoint", platform }),
    context: {
      platform,
      officeTheme: theme ?? undefined,
      requirements: {
        isSetSupported: (name, version = "1.1") => name === "PowerPointApi" && compareVersions(apiVersion, version) >= 0,
      },
      document: {
        addHandlerAsync(eventType, handler, optionsOrCallback, maybeCallback) {
          const callback = typeof optionsOrCallback === "function" ? optionsOrCallback : maybeCallback;
          if (eventType === "documentSelectionChanged") selectionHandlers.push(handler);
          callback?.({ status: "succeeded", value: undefined });
        },
      },
    },
  };

  return { Office, PowerPoint, host };
}
```

- [ ] **Step 4: Write** `dev/scenarios.js`

```js
// Named selections for the unit tests and the browser preview (dev/preview.html).
// Shape formats are described in dev/fake-powerpoint.js.

const SENTENCE = "The quick brown fox jumps over the lazy dog.";

const titlePlaceholder = { id: "101", name: "Title 1", type: "Placeholder", text: "Quarterly results" };
const bodyPlaceholder = {
  id: "102",
  name: "Content Placeholder 2",
  type: "Placeholder",
  text: "Revenue grew 12% year over year\rNew customers in three regions\rCosts held flat",
};
const textBox = { id: "103", name: "TextBox 3", type: "TextBox", text: SENTENCE };
const picture = { id: "104", name: "Picture 4", type: "Image" };
const table = {
  id: "105",
  name: "Table 5",
  type: "Table",
  rows: [
    ["Region", "Q1", "Q2"],
    ["North", "1,200 units", "1,450 units"],
    ["South — combined total", "", ""],
  ],
  merged: [{ row: 2, column: 0, rowCount: 1, columnCount: 3 }],
};

export const SCENARIOS = {
  nothing: { label: "Nothing selected", highlight: null, selected: [] },
  textBox: { label: "A text box", highlight: null, selected: [textBox] },
  title: { label: "A title placeholder", highlight: null, selected: [titlePlaceholder] },
  highlight: { label: "Highlighted text", highlight: "quick brown fox", selected: [textBox] },
  several: { label: "Three shapes (one is a picture)", highlight: null, selected: [titlePlaceholder, bodyPlaceholder, picture] },
  table: { label: "Table with merged cells", highlight: null, selected: [table] },
  group: {
    label: "Nested group",
    highlight: null,
    selected: [
      {
        id: "106",
        name: "Group 6",
        type: "Group",
        shapes: [
          { id: "107", name: "Rectangle 7", type: "GeometricShape", text: "Step one: plan" },
          {
            id: "108",
            name: "Group 8",
            type: "Group",
            shapes: [
              { id: "109", name: "Oval 9", type: "GeometricShape", text: "Step two: build" },
              { id: "110", name: "Arrow 10", type: "GeometricShape", text: "" },
              { id: "111", name: "Picture 11", type: "Image" },
            ],
          },
        ],
      },
    ],
  },
  chart: { label: "Chart and a text box", highlight: null, selected: [{ id: "112", name: "Chart 12", type: "Chart" }, textBox] },
  picture: { label: "A picture (no text)", highlight: null, selected: [picture] },
  emptyPlaceholder: { label: "An empty placeholder", highlight: null, selected: [{ id: "113", name: "Subtitle 2", type: "Placeholder", text: "" }] },
  tablePlaceholder: {
    label: "Table inside a content placeholder",
    highlight: null,
    selected: [{ id: "114", name: "Content Placeholder 3", type: "Placeholder", containedType: "Table", rows: [["Name", "Role"], ["Ana", "Design lead"]] }],
  },
  cjk: {
    label: "Japanese and English",
    highlight: null,
    selected: [{ id: "115", name: "TextBox 15", type: "TextBox", text: "東京で会議があります。 Meeting at 10:00 in Tokyo." }],
  },
  longNames: {
    label: "Many shapes with long names",
    highlight: null,
    selected: Array.from({ length: 6 }, (_, i) => ({
      id: String(200 + i),
      name: `Speaker notes callout with a very long descriptive name ${i + 1}`,
      type: "TextBox",
      text: "Lorem ipsum dolor sit amet ".repeat(i + 1).trim(),
    })),
  },
};
```

- [ ] **Step 5: Run test to verify it passes**

Run: `node --test test/fake-powerpoint.test.js`
Expected: PASS.

- [ ] **Step 6: Commit** (sequential executors only)

```bash
git add dev/fake-powerpoint.js dev/scenarios.js test/fake-powerpoint.test.js
git commit -m "test: add fake PowerPoint host with real load/sync rules"
```

---
### Task 3: Selection reader

**Files:**
- Create: `src/selection.js`
- Create: `test/selection.test.js`

**Interfaces:**
- Consumes: `createFakeHost` (Task 2) and `SCENARIOS` (Task 2), in tests only.
- Produces: `readSelection(context)` → `SelectionSnapshot`, and `tableCellTexts(cells)` (see *Shared interfaces*). Items only include shapes that yielded at least one text container, and `selectedCount` is the number of selected shapes, including pictures.

- [ ] **Step 1: Write the failing test** `test/selection.test.js`

```js
import assert from "node:assert/strict";
import { test } from "node:test";

import { createFakeHost } from "../dev/fake-powerpoint.js";
import { SCENARIOS } from "../dev/scenarios.js";
import { readSelection, tableCellTexts } from "../src/selection.js";

const SENTENCE = "The quick brown fox jumps over the lazy dog.";

async function read(scenario) {
  const { PowerPoint, host } = createFakeHost(scenario);
  const snapshot = await PowerPoint.run((context) => readSelection(context));
  return { snapshot, host };
}

test("nothing selected", async () => {
  const { snapshot } = await read(SCENARIOS.nothing);
  assert.deepEqual(snapshot, { source: "none", selectedCount: 0, items: [], unsupported: [] });
});

test("a text box: its whole text, in three syncs", async () => {
  const { snapshot, host } = await read(SCENARIOS.textBox);
  assert.deepEqual(snapshot, {
    source: "shapes",
    selectedCount: 1,
    items: [{ id: "103", name: "TextBox 3", kind: "text", texts: [SENTENCE] }],
    unsupported: [],
  });
  assert.ok(host.stats.syncs <= 3, `used ${host.stats.syncs} syncs`);
});

test("highlighted text counts only the highlight and names its shape, in one sync", async () => {
  const { snapshot, host } = await read(SCENARIOS.highlight);
  assert.deepEqual(snapshot, {
    source: "highlight",
    selectedCount: 1,
    items: [{ id: "103", name: "TextBox 3", kind: "highlight", texts: ["quick brown fox"] }],
    unsupported: [],
  });
  assert.equal(host.stats.syncs, 1);
});

test("a highlight with no selected shape (e.g. speaker notes) has no name", async () => {
  const { snapshot } = await read({ highlight: "notes text", selected: [] });
  assert.equal(snapshot.source, "highlight");
  assert.deepEqual(snapshot.items, [{ id: "", name: "", kind: "highlight", texts: ["notes text"] }]);
});

test("an empty highlight (just a cursor) counts the whole shape", async () => {
  const { snapshot } = await read({ highlight: "", selected: SCENARIOS.textBox.selected });
  assert.equal(snapshot.source, "shapes");
  assert.deepEqual(snapshot.items[0].texts, [SENTENCE]);
});

test("a highlight is ignored when several shapes are selected", async () => {
  const { snapshot } = await read({ highlight: "Quarterly", selected: SCENARIOS.several.selected });
  assert.equal(snapshot.source, "shapes");
});

test("several shapes: pictures are left out but still counted as selected", async () => {
  const { snapshot } = await read(SCENARIOS.several);
  assert.equal(snapshot.selectedCount, 3);
  assert.deepEqual(
    snapshot.items.map((i) => [i.name, i.kind, i.texts.length]),
    [
      ["Title 1", "text", 1],
      ["Content Placeholder 2", "text", 1],
    ],
  );
});

test("a table: every cell once, merged cells not repeated", async () => {
  const { snapshot, host } = await read(SCENARIOS.table);
  assert.deepEqual(snapshot.items, [
    {
      id: "105",
      name: "Table 5",
      kind: "table",
      texts: ["Region", "Q1", "Q2", "North", "1,200 units", "1,450 units", "South — combined total"],
    },
  ]);
  assert.ok(host.stats.syncs <= 3, `used ${host.stats.syncs} syncs`);
});

test("nested groups: every text shape inside, attributed to the top-level group", async () => {
  const { snapshot, host } = await read(SCENARIOS.group);
  assert.deepEqual(snapshot.items, [{ id: "106", name: "Group 6", kind: "group", texts: ["Step one: plan", "Step two: build", ""] }]);
  assert.ok(host.stats.syncs <= 5, `used ${host.stats.syncs} syncs`);
});

test("a group containing a table reads the table's cells", async () => {
  const group = {
    id: "g",
    name: "Group 1",
    type: "Group",
    shapes: [
      { id: "t", name: "Table 1", type: "Table", rows: [["a", "b"]] },
      { id: "x", name: "TextBox 1", type: "TextBox", text: "c" },
    ],
  };
  const { snapshot } = await read({ highlight: null, selected: [group] });
  assert.equal(snapshot.items[0].kind, "group");
  assert.deepEqual([...snapshot.items[0].texts].sort(), ["a", "b", "c"]);
});

test("a shape selected together with its group is counted once", async () => {
  const inner = { id: "t", name: "TextBox 1", type: "TextBox", text: "inside" };
  const other = { id: "u", name: "TextBox 2", type: "TextBox", text: "sibling" };
  const group = { id: "g", name: "Group 1", type: "Group", shapes: [inner, other] };
  const { snapshot } = await read({ highlight: null, selected: [group, inner] });
  assert.deepEqual(snapshot.items.flatMap((i) => i.texts).sort(), ["inside", "sibling"]);
});

test("charts are reported as not countable; the rest still counts", async () => {
  const { snapshot } = await read(SCENARIOS.chart);
  assert.deepEqual(snapshot.unsupported, [{ id: "112", name: "Chart 12", type: "Chart" }]);
  assert.deepEqual(
    snapshot.items.map((i) => i.name),
    ["TextBox 3"],
  );
  assert.equal(snapshot.selectedCount, 2);
});

test("a picture alone: nothing to count", async () => {
  const { snapshot } = await read(SCENARIOS.picture);
  assert.deepEqual(snapshot, { source: "shapes", selectedCount: 1, items: [], unsupported: [] });
});

test("an empty placeholder is an item with empty text", async () => {
  const { snapshot } = await read(SCENARIOS.emptyPlaceholder);
  assert.deepEqual(
    snapshot.items.map((i) => i.texts),
    [[""]],
  );
});

test("a table inside a content placeholder is read as a table", async () => {
  const { snapshot } = await read(SCENARIOS.tablePlaceholder);
  assert.deepEqual(snapshot.items, [{ id: "114", name: "Content Placeholder 3", kind: "table", texts: ["Name", "Role", "Ana", "Design lead"] }]);
});

test("a table reported as 'Unsupported' (older builds) is still counted", async () => {
  const { snapshot } = await read({ highlight: null, selected: [{ id: "u", name: "Table 1", type: "Unsupported", rows: [["a", "b"]] }] });
  assert.deepEqual(snapshot.items, [{ id: "u", name: "Table 1", kind: "table", texts: ["a", "b"] }]);
  assert.deepEqual(snapshot.unsupported, []);
});

test("an 'Unsupported' shape that isn't a table is reported, not fatal", async () => {
  const { snapshot } = await read({ highlight: null, selected: [{ id: "z", name: "Zoom 1", type: "Unsupported" }, ...SCENARIOS.textBox.selected] });
  assert.deepEqual(snapshot.unsupported, [{ id: "z", name: "Zoom 1", type: "Unsupported" }]);
  assert.deepEqual(
    snapshot.items.map((i) => i.name),
    ["TextBox 3"],
  );
});

test("one shape failing doesn't hide the others", async () => {
  const broken = { id: "9", name: "Broken 9", type: "TextBox", text: "x", failTextFrame: true };
  const { snapshot } = await read({ highlight: null, selected: [SCENARIOS.textBox.selected[0], broken, SCENARIOS.title.selected[0]] });
  assert.deepEqual(
    snapshot.items.map((i) => i.name),
    ["TextBox 3", "Title 1"],
  );
  assert.deepEqual(snapshot.unsupported, [{ id: "9", name: "Broken 9", type: "TextBox" }]);
});

test("never uses the APIs that throw on the wrong kind of shape", async () => {
  for (const [key, scenario] of Object.entries(SCENARIOS)) {
    const { host } = await read(scenario);
    assert.deepEqual(host.stats.forbidden, [], key);
  }
});

test("a failed first sync rejects, so the caller can retry", async () => {
  const { PowerPoint, host } = createFakeHost(SCENARIOS.textBox);
  host.failNextSyncs(1);
  await assert.rejects(PowerPoint.run((context) => readSelection(context)), { code: "GeneralException" });
});

test("tableCellTexts skips cells hidden under a merge", () => {
  assert.deepEqual(
    tableCellTexts([
      { isNullObject: false, text: "a" },
      { isNullObject: true, text: null },
      { isNullObject: false, text: "" },
    ]),
    ["a", ""],
  );
});
```

- [ ] **Step 2: Run test to verify it fails**

Run: `node --test test/selection.test.js`
Expected: FAIL with `Cannot find module '.../src/selection.js'`.

- [ ] **Step 3: Write the implementation** `src/selection.js`

```js
// Reads the current PowerPoint selection and returns the text to count.
//
// This is the only module that talks to the PowerPoint API. Several Office.js
// calls fail the whole context.sync() when used on the wrong kind of shape
// (textFrame on a table, group on a non-group...), so each step checks the
// shape's type first, and a failed batch is retried one job at a time so one
// bad shape can't hide the rest.
// Design: docs/superpowers/specs/2026-09-26-ppt-word-count-design.md, section 4.3.

/**
 * @typedef {{ id: string, name: string, kind: "highlight" | "text" | "table" | "group", texts: string[] }} SelectionItem
 * @typedef {{ id: string, name: string, type: string }} UnsupportedShape
 * @typedef {{ source: "highlight" | "shapes" | "none", selectedCount: number, items: SelectionItem[], unsupported: UnsupportedShape[] }} SelectionSnapshot
 */

const SHAPE_FIELDS = "items/id,items/name,items/type";
// Shapes that never hold text.
const NO_TEXT = new Set(["Image", "Line", "Media", "Model3D", "Ink", "ContentApp", "Graphic"]);
// Shapes that show text the API can't read.
const UNREADABLE = new Set(["Chart", "SmartArt", "Diagram", "Ole"]);

/**
 * @param {PowerPoint.RequestContext} context
 * @returns {Promise<SelectionSnapshot>}
 */
export async function readSelection(context) {
  const range = context.presentation.getSelectedTextRangeOrNullObject();
  range.load("text");
  const selection = context.presentation.getSelectedShapes();
  selection.load(SHAPE_FIELDS);
  await context.sync();

  const shapes = selection.items;
  // Highlighted text wins unless several shapes are selected (what the range means
  // then is undocumented). Never navigate from the range to its shape: that fails
  // for text inside table cells.
  if (!range.isNullObject && range.text.length > 0 && shapes.length <= 1) {
    const shape = shapes[0];
    return {
      source: "highlight",
      selectedCount: shapes.length,
      items: [{ id: shape?.id ?? "", name: shape?.name ?? "", kind: "highlight", texts: [range.text] }],
      unsupported: [],
    };
  }
  if (shapes.length === 0) return { source: "none", selectedCount: 0, items: [], unsupported: [] };
  return readShapes(context, shapes);
}

/**
 * Text of each table cell, skipping cells hidden under a merged area. PowerPoint
 * returns a null object for those; only a merge's top-left cell holds its text.
 * @param {Array<{ isNullObject: boolean, text: string }>} cells
 * @returns {string[]}
 */
export function tableCellTexts(cells) {
  return cells.filter((cell) => !cell.isNullObject).map((cell) => cell.text);
}

// Walks the selected shapes breadth first. Each round is one context.sync(), so a
// selection costs one sync per group-nesting level plus two for reading text.
async function readShapes(context, shapes) {
  const items = [];
  const unsupported = [];
  const seen = new Set();
  let pending = [];

  const reportUnsupported = (shape) => {
    if (!unsupported.some((u) => u.id === shape.id)) unsupported.push({ id: shape.id, name: shape.name, type: shape.type });
  };
  // Queues the work for one shape. Its text goes to `item`, its top-level selected shape.
  const visit = (shape, item) => {
    if (seen.has(shape.id)) return; // a group and its child can both be selected
    seen.add(shape.id);
    if (NO_TEXT.has(shape.type)) return;
    if (UNREADABLE.has(shape.type)) return reportUnsupported(shape);
    if (shape.type === "Group") pending.push(groupJob(shape, item));
    else if (shape.type === "Table") pending.push(tableJob(shape, item));
    else pending.push(textFrameJob(shape, item));
  };

  for (const shape of shapes) {
    const item = { id: shape.id, name: shape.name, kind: kindOf(shape.type), texts: [] };
    items.push(item);
    visit(shape, item);
  }

  while (pending.length > 0) {
    const round = pending;
    pending = [];
    for (const { job, outcome } of await runJobs(context, round, reportUnsupported)) {
      for (const child of outcome.children ?? []) visit(child, job.item);
      if (outcome.next) pending.push(outcome.next);
      if (outcome.unsupported) reportUnsupported(job.shape);
    }
  }

  return {
    source: "shapes",
    selectedCount: shapes.length,
    items: items.filter((item) => item.texts.length > 0),
    unsupported,
  };
}

// Syncs a round of jobs as one batch. When the batch fails, each job is retried
// on its own; jobs that still fail are reported instead of breaking the count.
async function runJobs(context, jobs, reportUnsupported) {
  for (const job of jobs) job.queue();
  let batchFailed = false;
  try {
    await context.sync();
  } catch {
    batchFailed = true;
  }
  const done = [];
  for (const job of jobs) {
    try {
      if (batchFailed) {
        job.queue();
        await context.sync();
      }
      done.push({ job, outcome: job.read() });
    } catch {
      reportUnsupported(job.shape);
    }
  }
  return done;
}

function kindOf(type) {
  if (type === "Group") return "group";
  if (type === "Table") return "table";
  return "text";
}

// A job queues loads for one shape (queue) and interprets them after the sync
// (read). read() can return group children to visit, a follow-up job for the
// next round, or `unsupported: true`.

function groupJob(shape, item) {
  let children;
  return {
    shape,
    item,
    queue() {
      children = shape.group.shapes;
      children.load(SHAPE_FIELDS);
    },
    read: () => ({ children: children.items }),
  };
}

function tableJob(shape, item) {
  let table;
  return {
    shape,
    item,
    queue() {
      table = shape.getTable();
      table.load("rowCount,columnCount");
    },
    read: () => ({ next: tableCellsJob(shape, item, table) }),
  };
}

// Reads cells one by one rather than Table.values: getCellOrNullObject is
// documented to return a null object for cells hidden under a merge, but what
// `values` holds for them is not documented.
function tableCellsJob(shape, item, table) {
  let cells;
  return {
    shape,
    item,
    queue() {
      cells = [];
      for (let row = 0; row < table.rowCount; row++) {
        for (let column = 0; column < table.columnCount; column++) {
          const cell = table.getCellOrNullObject(row, column);
          cell.load("text");
          cells.push(cell);
        }
      }
    },
    read() {
      if (item.id === shape.id) item.kind = "table";
      item.texts.push(...tableCellTexts(cells));
      return {};
    },
  };
}

function textFrameJob(shape, item) {
  let frame;
  let format = null;
  return {
    shape,
    item,
    queue() {
      frame = shape.getTextFrameOrNullObject();
      frame.load("hasText");
      if (shape.type === "Placeholder") {
        format = shape.placeholderFormat;
        format.load("containedType");
      }
    },
    read() {
      if (!frame.isNullObject) return { next: textRangeJob(shape, item, frame) };
      const contained = format?.containedType;
      // Tables in content placeholders, and tables on builds that report their
      // type as "Unsupported", have no text frame but can still be read as tables.
      if (contained === "Table" || shape.type === "Unsupported") return { next: tableJob(shape, item) };
      if (UNREADABLE.has(contained)) return { unsupported: true };
      return {};
    },
  };
}

// The frame's text is read in a second round: navigating from a frame that might
// be a null object into its textRange is not documented to be safe.
function textRangeJob(shape, item, frame) {
  let range;
  return {
    shape,
    item,
    queue() {
      range = frame.textRange;
      range.load("text");
    },
    read() {
      item.texts.push(range.text);
      return {};
    },
  };
}
```

- [ ] **Step 4: Run test to verify it passes**

Run: `node --test test/selection.test.js`
Expected: PASS.

- [ ] **Step 5: Commit** (sequential executors only)

```bash
git add src/selection.js test/selection.test.js
git commit -m "feat: read text from highlighted text, shapes, tables and groups"
```

---

### Task 4: Refresh scheduler

**Files:**
- Create: `src/refresher.js`
- Create: `test/refresher.test.js`

**Interfaces:**
- Consumes: nothing.
- Produces: `createRefresher({ read, onResult, onError, isEmpty = () => false, timers = globalThis, debounceMs = 120, retryMs = 300, confirmEmptyMs = 150 })` → `{ schedule, refreshNow, poll, stop }`.
  - `schedule()`: debounced trigger.
  - `refreshNow()`: immediate read. Returns the run's promise, or `undefined` if a run is already in flight (it is then marked dirty).
  - `poll(intervalMs, shouldPoll)`: periodic read while `shouldPoll()` is true and nothing is running or scheduled.
  - `stop()`: cancels everything; later calls do nothing.

- [ ] **Step 1: Write the failing test** `test/refresher.test.js`

```js
import assert from "node:assert/strict";
import { afterEach, beforeEach, mock, test } from "node:test";

import { createRefresher } from "../src/refresher.js";

// Lets pending promise callbacks run (setImmediate is not mocked).
const flush = async () => {
  for (let i = 0; i < 10; i++) await new Promise((resolve) => setImmediate(resolve));
};

function deferred() {
  let resolve;
  let reject;
  const promise = new Promise((res, rej) => {
    resolve = res;
    reject = rej;
  });
  return { promise, resolve, reject };
}

// Each read() returns the next entry: a value, an Error (rejects), or a deferred.
function setup(reads, options = {}) {
  const log = { reads: 0, results: [], errors: [] };
  const read = () => {
    const next = reads[Math.min(log.reads, reads.length - 1)];
    log.reads++;
    if (next instanceof Error) return Promise.reject(next);
    if (next && typeof next.promise?.then === "function") return next.promise;
    return Promise.resolve(next);
  };
  const refresher = createRefresher({
    read,
    onResult: (result) => log.results.push(result),
    onError: (error) => log.errors.push(error),
    ...options,
  });
  return { refresher, log };
}

beforeEach(() => mock.timers.enable({ apis: ["setTimeout", "setInterval"] }));
afterEach(() => mock.timers.reset());

test("schedule() debounces a burst of triggers into one read", async () => {
  const { refresher, log } = setup(["A"]);
  refresher.schedule();
  refresher.schedule();
  refresher.schedule();
  mock.timers.tick(119);
  await flush();
  assert.equal(log.reads, 0);
  mock.timers.tick(1);
  await flush();
  assert.equal(log.reads, 1);
  assert.deepEqual(log.results, ["A"]);
});

test("refreshNow() reads straight away", async () => {
  const { refresher, log } = setup(["A"]);
  refresher.refreshNow();
  await flush();
  assert.deepEqual(log.results, ["A"]);
});

test("only one read runs at a time; a trigger during a read causes exactly one re-read and the stale result is dropped", async () => {
  const first = deferred();
  const { refresher, log } = setup([first, "second"]);
  refresher.refreshNow();
  await flush();
  refresher.schedule();
  mock.timers.tick(120);
  refresher.refreshNow();
  await flush();
  assert.equal(log.reads, 1);
  first.resolve("first");
  await flush();
  assert.equal(log.reads, 2);
  assert.deepEqual(log.results, ["second"]);
});

test("a failed read is retried once after 300 ms", async () => {
  const { refresher, log } = setup([new Error("transient"), "ok"]);
  refresher.refreshNow();
  await flush();
  assert.equal(log.reads, 1);
  mock.timers.tick(300);
  await flush();
  assert.equal(log.reads, 2);
  assert.deepEqual(log.results, ["ok"]);
  assert.deepEqual(log.errors, []);
});

test("two failures report an error and keep the last good result", async () => {
  const { refresher, log } = setup(["good", new Error("a"), new Error("b")]);
  refresher.refreshNow();
  await flush();
  refresher.refreshNow();
  await flush();
  mock.timers.tick(300);
  await flush();
  assert.deepEqual(log.results, ["good"]);
  assert.equal(log.errors.length, 1);
});

test("a sudden empty result is confirmed by a second read before it is shown", async () => {
  const { refresher, log } = setup(["full", "empty", "empty"], { isEmpty: (r) => r === "empty" });
  refresher.refreshNow();
  await flush();
  refresher.refreshNow();
  await flush();
  assert.deepEqual(log.results, ["full"]);
  mock.timers.tick(150);
  await flush();
  assert.equal(log.reads, 3);
  assert.deepEqual(log.results, ["full", "empty"]);
});

test("if the confirming read has text again, that is shown instead", async () => {
  const { refresher, log } = setup(["full", "empty", "full again"], { isEmpty: (r) => r === "empty" });
  refresher.refreshNow();
  await flush();
  refresher.refreshNow();
  await flush();
  mock.timers.tick(150);
  await flush();
  assert.deepEqual(log.results, ["full", "full again"]);
});

test("empty after empty needs no confirmation", async () => {
  const { refresher, log } = setup(["empty", "empty"], { isEmpty: (r) => r === "empty" });
  refresher.refreshNow();
  await flush();
  refresher.refreshNow();
  await flush();
  assert.equal(log.reads, 2);
  assert.deepEqual(log.results, ["empty", "empty"]);
});

test("poll() reads on an interval only while shouldPoll() is true", async () => {
  let visible = true;
  const { refresher, log } = setup(["A"]);
  refresher.poll(1000, () => visible);
  mock.timers.tick(1000);
  await flush();
  assert.equal(log.reads, 1);
  visible = false;
  mock.timers.tick(1000);
  await flush();
  assert.equal(log.reads, 1);
});

test("poll() skips a tick while a read is still running", async () => {
  const slow = deferred();
  const { refresher, log } = setup([slow]);
  refresher.poll(1000, () => true);
  mock.timers.tick(1000);
  await flush();
  mock.timers.tick(1000);
  await flush();
  assert.equal(log.reads, 1);
});

test("stop() cancels pending and future work", async () => {
  const { refresher, log } = setup(["A"]);
  refresher.poll(1000, () => true);
  refresher.schedule();
  refresher.stop();
  mock.timers.tick(5000);
  await flush();
  assert.equal(refresher.refreshNow(), undefined);
  await flush();
  assert.equal(log.reads, 0);
});
```

- [ ] **Step 2: Run test to verify it fails**

Run: `node --test test/refresher.test.js`
Expected: FAIL with `Cannot find module '.../src/refresher.js'`.

- [ ] **Step 3: Write the implementation** `src/refresher.js`

```js
// Decides when to re-read the selection. It debounces bursts of triggers, runs
// one read at a time, retries a failed read once, and double-checks a sudden
// empty selection (PowerPoint occasionally reports one for a moment). This is pure
// scheduling logic: the caller supplies read() and the callbacks, and timers can
// be injected for tests.

/**
 * @template T
 * @param {{
 *   read: () => Promise<T>,
 *   onResult: (result: T) => void,
 *   onError: (error: unknown) => void,
 *   isEmpty?: (result: T) => boolean,
 *   timers?: Pick<typeof globalThis, "setTimeout" | "clearTimeout" | "setInterval" | "clearInterval">,
 *   debounceMs?: number,
 *   retryMs?: number,
 *   confirmEmptyMs?: number,
 * }} options
 */
export function createRefresher({ read, onResult, onError, isEmpty = () => false, timers = globalThis, debounceMs = 120, retryMs = 300, confirmEmptyMs = 150 }) {
  let debounceTimer = null;
  let pollTimer = null;
  let running = false;
  let dirty = false; // a trigger arrived while a read was running
  let stopped = false;
  let last;
  let hasLast = false;

  const wait = (ms) => new Promise((resolve) => timers.setTimeout(resolve, ms));

  async function readWithRetry() {
    try {
      return await read();
    } catch {
      await wait(retryMs);
      return read();
    }
  }

  async function run() {
    if (stopped) return;
    if (running) {
      dirty = true;
      return;
    }
    running = true;
    dirty = false;
    try {
      let result = await readWithRetry();
      if (!dirty && hasLast && isEmpty(result) && !isEmpty(last)) {
        await wait(confirmEmptyMs);
        if (!dirty) result = await readWithRetry();
      }
      // If the selection changed during the read, this result is already stale.
      if (!dirty && !stopped) {
        last = result;
        hasLast = true;
        onResult(result);
      }
    } catch (error) {
      if (!dirty && !stopped) onError(error);
    } finally {
      running = false;
    }
    if (dirty && !stopped) run();
  }

  return {
    schedule() {
      if (stopped) return;
      timers.clearTimeout(debounceTimer);
      debounceTimer = timers.setTimeout(() => {
        debounceTimer = null;
        run();
      }, debounceMs);
    },

    refreshNow() {
      if (stopped) return undefined;
      timers.clearTimeout(debounceTimer);
      debounceTimer = null;
      return run();
    },

    poll(intervalMs, shouldPoll = () => true) {
      timers.clearInterval(pollTimer);
      pollTimer = timers.setInterval(() => {
        if (!running && debounceTimer === null && shouldPoll()) run();
      }, intervalMs);
    },

    stop() {
      stopped = true;
      timers.clearTimeout(debounceTimer);
      timers.clearInterval(pollTimer);
    },
  };
}
```

- [ ] **Step 4: Run test to verify it passes**

Run: `node --test test/refresher.test.js`
Expected: PASS.

- [ ] **Step 5: Commit** (sequential executors only)

```bash
git add src/refresher.js test/refresher.test.js
git commit -m "feat: add debounced single-flight refresh scheduler"
```

---
### Task 5: Task pane UI and browser preview

**Files:**
- Create: `src/app.js`, `src/taskpane.html`, `src/taskpane.js`, `src/taskpane.css`
- Create: `dev/preview.html`
- Create: `test/app.test.js`
- Modify: `scripts/serve.mjs` (full replacement below: URL layout, `/` → `src/`)

**Interfaces:**
- Consumes: `countChunks`, `sumCounts` (Task 1); `readSelection` (Task 3); `createRefresher` (Task 4); `createFakeHost`, `SCENARIOS` (Task 2, preview only).
- Produces: `start`, `summarize`, `describeSelection`, `unsupportedNote`, `themeFrom` (see *Shared interfaces*). `start` returns `null` when it only shows a message (wrong host or old PowerPoint), otherwise `{ refresher, view }`.

UI requirements (from the spec):
- Layout at 320 px: a context header (title + muted detail); a full-width **Words** tile with a big accent-coloured number; two half-width tiles for the character counts; an optional note; an optional "By shape" breakdown (≥2 items); an error status with "Try again"; and a footer ("Counted the way Microsoft Word counts", Refresh, Help).
- Empty states keep the tiles in place with "–" so the layout never jumps.
- Numbers use `Intl.NumberFormat` and tabular digits.
- The view skips DOM updates when the summary hasn't changed, because polling re-reads often.
- Theme: PowerPoint colours, when valid, set `--office-bg` / `--office-fg` and `data-theme`. Otherwise `prefers-color-scheme` applies.

- [ ] **Step 1: Write the failing test** `test/app.test.js`

```js
import assert from "node:assert/strict";
import { test } from "node:test";

import { describeSelection, summarize, themeFrom, unsupportedNote } from "../src/app.js";

const summary = (overrides) => ({
  source: "shapes",
  selectedCount: 1,
  items: [],
  unsupported: [],
  total: { words: 0, characters: 0, charactersWithSpaces: 0 },
  ...overrides,
});
const item = (name, kind = "text") => ({ id: name, name, kind, counts: { words: 1, characters: 1, charactersWithSpaces: 1 } });

test("summarize counts each item and the total", () => {
  const result = summarize({
    source: "shapes",
    selectedCount: 2,
    items: [
      { id: "1", name: "TextBox 1", kind: "text", texts: ["Hello world"] },
      { id: "2", name: "Table 1", kind: "table", texts: ["a", "b c"] },
    ],
    unsupported: [],
  });
  assert.deepEqual(result.items[0], { id: "1", name: "TextBox 1", kind: "text", counts: { words: 2, characters: 10, charactersWithSpaces: 11 } });
  assert.deepEqual(result.items[1].counts, { words: 3, characters: 3, charactersWithSpaces: 4 });
  assert.deepEqual(result.total, { words: 5, characters: 13, charactersWithSpaces: 15 });
  assert.equal(result.source, "shapes");
  assert.equal(result.selectedCount, 2);
});

test("describe: nothing selected", () => {
  assert.deepEqual(describeSelection(summary({ source: "none", selectedCount: 0 })), {
    title: "Nothing selected",
    detail: "Select a text box, table or group, or highlight some text.",
    note: "",
    empty: true,
  });
});

test("describe: highlighted text, with and without a shape name", () => {
  assert.deepEqual(describeSelection(summary({ source: "highlight", items: [item("TextBox 3", "highlight")] })), {
    title: "Highlighted text",
    detail: "in TextBox 3",
    note: "",
    empty: false,
  });
  assert.equal(describeSelection(summary({ source: "highlight", selectedCount: 0, items: [item("", "highlight")] })).detail, "");
});

test("describe: one shape uses its name; tables and groups say what's counted", () => {
  assert.deepEqual(describeSelection(summary({ items: [item("TextBox 3")] })), { title: "TextBox 3", detail: "", note: "", empty: false });
  assert.equal(describeSelection(summary({ items: [item("Table 2", "table")] })).detail, "Whole table");
  assert.equal(describeSelection(summary({ items: [item("Group 5", "group")] })).detail, "Whole group");
});

test("describe: several shapes", () => {
  assert.equal(describeSelection(summary({ selectedCount: 3, items: [item("A"), item("B")] })).title, "3 shapes selected");
});

test("describe: selected shapes without text", () => {
  assert.deepEqual(describeSelection(summary({ selectedCount: 1 })), {
    title: "No text to count",
    detail: "The selected shape doesn't contain text.",
    note: "",
    empty: true,
  });
  assert.equal(describeSelection(summary({ selectedCount: 2 })).detail, "The selected shapes don't contain text.");
});

test("describe: carries the note about objects that can't be counted", () => {
  const unsupported = [{ id: "c", name: "Chart 1", type: "Chart" }];
  assert.equal(describeSelection(summary({ selectedCount: 2, items: [item("A")], unsupported })).note, unsupportedNote(1));
  assert.equal(describeSelection(summary({ selectedCount: 1, unsupported })).note, unsupportedNote(1));
});

test("unsupportedNote wording", () => {
  assert.equal(unsupportedNote(0), "");
  assert.equal(unsupportedNote(1), "1 selected object can't be counted: PowerPoint doesn't let add-ins read text in charts, SmartArt and some other objects.");
  assert.equal(unsupportedNote(2), "2 selected objects can't be counted: PowerPoint doesn't let add-ins read text in charts, SmartArt and some other objects.");
});

test("themeFrom uses PowerPoint's colours and dark flag", () => {
  assert.deepEqual(themeFrom({ bodyBackgroundColor: "#1F1F1F", bodyForegroundColor: "#FFFFFF", isDarkTheme: true }), {
    background: "#1f1f1f",
    foreground: "#ffffff",
    dark: true,
  });
});

test("themeFrom works out dark or light from the background when the flag is missing", () => {
  assert.equal(themeFrom({ bodyBackgroundColor: "#202020", bodyForegroundColor: "#F0F0F0" }).dark, true);
  assert.equal(themeFrom({ bodyBackgroundColor: "FFFFFF", bodyForegroundColor: "000000" }).dark, false);
});

test("themeFrom ignores missing, empty or invalid themes", () => {
  assert.equal(themeFrom(undefined), null);
  assert.equal(themeFrom({}), null);
  assert.equal(themeFrom({ bodyBackgroundColor: "red", bodyForegroundColor: "#000000" }), null);
});
```

- [ ] **Step 2: Run test to verify it fails**

Run: `node --test test/app.test.js`
Expected: FAIL with `Cannot find module '.../src/app.js'`.

- [ ] **Step 3: Write the implementation** `src/app.js`

```js
// The task pane: wires PowerPoint's selection events to the refresher and draws
// the counts. Office and PowerPoint are passed in, so the browser preview
// (dev/preview.html) runs this same code against a fake PowerPoint.

import { countChunks, sumCounts } from "./count.js";
import { createRefresher } from "./refresher.js";
import { readSelection } from "./selection.js";

/**
 * @typedef {import("./count.js").Counts} Counts
 * @typedef {import("./selection.js").SelectionSnapshot} SelectionSnapshot
 * @typedef {{ id: string, name: string, kind: string, counts: Counts }} SummaryItem
 * @typedef {{ source: SelectionSnapshot["source"], selectedCount: number, items: SummaryItem[], unsupported: SelectionSnapshot["unsupported"], total: Counts }} Summary
 */

const REQUIRED_API = "1.10";
// PowerPoint fires no event while you type, so desktop panes re-read every so
// often. Not on the web, where frequent reads have been seen to trigger autosave.
const POLL_MS = 1500;
const STATS = ["words", "characters", "charactersWithSpaces"];
const KIND_DETAIL = { text: "", highlight: "", table: "Whole table", group: "Whole group" };
const numberFormat = new Intl.NumberFormat();

/**
 * Starts the pane inside `root`.
 * @param {{ Office: any, PowerPoint: any, root: HTMLElement, info: { host: string | null, platform: string | null } }} deps
 */
export function start({ Office, PowerPoint, root, info }) {
  const doc = root.ownerDocument;
  const view = createView(root);

  if (info?.host !== Office.HostType.PowerPoint) {
    view.showMessage("Open Word Count from PowerPoint", "This page is the Word Count add-in. It works inside PowerPoint: choose Word Count on the Home tab.");
    return null;
  }
  if (!Office.context.requirements.isSetSupported("PowerPointApi", REQUIRED_API)) {
    view.showMessage(
      "Please update PowerPoint",
      "Word Count needs a recent version of PowerPoint: 16.105 or later on Mac, Version 2601 or later on Windows, or PowerPoint on the web.",
    );
    return null;
  }

  const applyOfficeTheme = () => applyTheme(doc.documentElement, Office.context.officeTheme);
  applyOfficeTheme();

  const refresher = createRefresher({
    read: () => PowerPoint.run((context) => readSelection(context)),
    onResult: (snapshot) => view.render(summarize(snapshot)),
    onError: () => view.showError(),
    isEmpty: (snapshot) => snapshot.items.length === 0,
  });

  Office.context.document.addHandlerAsync(
    Office.EventType.DocumentSelectionChanged,
    () => refresher.schedule(),
    (result) => {
      if (result?.status === Office.AsyncResultStatus.Failed) {
        view.showNotice("Counts won't update by themselves here. Choose Refresh after changing the selection.");
      }
    },
  );
  doc.defaultView.addEventListener("focus", () => {
    applyOfficeTheme();
    refresher.schedule();
  });
  doc.addEventListener("visibilitychange", () => {
    if (doc.visibilityState === "visible") refresher.schedule();
  });
  view.onRefresh(() => refresher.refreshNow());

  if (info.platform !== Office.PlatformType.OfficeOnline) {
    refresher.poll(POLL_MS, () => doc.visibilityState === "visible");
  }
  // PowerPoint on the web can drop the first request made while the add-in starts,
  // so give start-up a moment. The refresher also retries a failed read.
  doc.defaultView.setTimeout(() => refresher.refreshNow(), 50);
  return { refresher, view };
}

/**
 * Counts every item of a selection snapshot and the total.
 * @param {SelectionSnapshot} snapshot
 * @returns {Summary}
 */
export function summarize(snapshot) {
  const items = snapshot.items.map(({ id, name, kind, texts }) => ({ id, name, kind, counts: countChunks(texts) }));
  return {
    source: snapshot.source,
    selectedCount: snapshot.selectedCount,
    items,
    unsupported: snapshot.unsupported,
    total: sumCounts(items.map((item) => item.counts)),
  };
}

/**
 * The words shown above the counts.
 * @param {Summary} summary
 * @returns {{ title: string, detail: string, note: string, empty: boolean }}
 */
export function describeSelection({ source, selectedCount, items, unsupported }) {
  const note = unsupportedNote(unsupported.length);
  if (source === "highlight") {
    const name = items[0]?.name;
    return { title: "Highlighted text", detail: name ? `in ${name}` : "", note, empty: false };
  }
  if (items.length === 0) {
    if (selectedCount === 0) {
      return { title: "Nothing selected", detail: "Select a text box, table or group, or highlight some text.", note, empty: true };
    }
    const detail = selectedCount === 1 ? "The selected shape doesn't contain text." : "The selected shapes don't contain text.";
    return { title: "No text to count", detail, note, empty: true };
  }
  if (selectedCount === 1) {
    const [only] = items;
    return { title: only.name || "Selected shape", detail: KIND_DETAIL[only.kind] ?? "", note, empty: false };
  }
  return { title: `${selectedCount} shapes selected`, detail: "", note, empty: false };
}

/** @param {number} count */
export function unsupportedNote(count) {
  if (count === 0) return "";
  const what = count === 1 ? "1 selected object can't" : `${count} selected objects can't`;
  return `${what} be counted: PowerPoint doesn't let add-ins read text in charts, SmartArt and some other objects.`;
}

const HEX_COLOR = /^#?([0-9a-f]{6})$/i;

function hexColor(value) {
  const match = typeof value === "string" ? HEX_COLOR.exec(value.trim()) : null;
  return match ? `#${match[1].toLowerCase()}` : null;
}

function relativeLuminance(hex) {
  const n = Number.parseInt(hex.slice(1), 16);
  const [r, g, b] = [(n >> 16) & 255, (n >> 8) & 255, n & 255].map((v) => {
    const c = v / 255;
    return c <= 0.03928 ? c / 12.92 : ((c + 0.055) / 1.055) ** 2.4;
  });
  return 0.2126 * r + 0.7152 * g + 0.0722 * b;
}

/**
 * PowerPoint's UI colours. Office.context.officeTheme can be missing, or an empty
 * object, on some platforms, so every value is checked.
 * @returns {{ background: string, foreground: string, dark: boolean } | null}
 */
export function themeFrom(officeTheme) {
  const background = hexColor(officeTheme?.bodyBackgroundColor);
  const foreground = hexColor(officeTheme?.bodyForegroundColor);
  if (!background || !foreground) return null;
  const dark = typeof officeTheme.isDarkTheme === "boolean" ? officeTheme.isDarkTheme : relativeLuminance(background) < 0.2;
  return { background, foreground, dark };
}

function applyTheme(element, officeTheme) {
  const theme = themeFrom(officeTheme);
  if (!theme) {
    delete element.dataset.theme;
    element.style.removeProperty("--office-bg");
    element.style.removeProperty("--office-fg");
    return;
  }
  element.dataset.theme = theme.dark ? "dark" : "light";
  element.style.setProperty("--office-bg", theme.background);
  element.style.setProperty("--office-fg", theme.foreground);
}

// Static markup only. Anything that comes from the presentation is set with textContent.
const TEMPLATE = `
  <main class="pane">
    <header class="context">
      <h1 class="context__title">Reading the selection…</h1>
      <p class="context__detail" hidden></p>
    </header>
    <dl class="stats" aria-live="polite">
      <div class="stat stat--words">
        <dt class="stat__label">Words</dt>
        <dd class="stat__value" data-stat="words">–</dd>
      </div>
      <div class="stat">
        <dt class="stat__label">Characters <span class="stat__qualifier">(no spaces)</span></dt>
        <dd class="stat__value" data-stat="characters">–</dd>
      </div>
      <div class="stat">
        <dt class="stat__label">Characters <span class="stat__qualifier">(with spaces)</span></dt>
        <dd class="stat__value" data-stat="charactersWithSpaces">–</dd>
      </div>
    </dl>
    <p class="note" hidden></p>
    <section class="breakdown" hidden>
      <h2 class="breakdown__title">By shape</h2>
      <ol class="breakdown__list"></ol>
    </section>
    <p class="status" role="status" hidden>
      <span>Couldn't read the selection.</span>
      <button type="button" class="link-button status__retry">Try again</button>
    </p>
    <p class="notice" hidden></p>
    <footer class="footer">
      <span>Counted the way Microsoft Word counts</span>
      <span class="footer__actions">
        <button type="button" class="link-button footer__refresh">Refresh</button>
        <a class="link-button" href="help.html" target="_blank" rel="noopener">Help</a>
      </span>
    </footer>
  </main>`;

function createView(root) {
  const doc = root.ownerDocument;
  root.innerHTML = TEMPLATE;
  const $ = (selector) => root.querySelector(selector);
  const pane = $(".pane");
  const title = $(".context__title");
  const detail = $(".context__detail");
  const stats = $(".stats");
  const values = Object.fromEntries(STATS.map((key) => [key, $(`[data-stat="${key}"]`)]));
  const note = $(".note");
  const breakdown = $(".breakdown");
  const list = $(".breakdown__list");
  const status = $(".status");
  const notice = $(".notice");
  let lastRendered = "";

  const setText = (element, text) => {
    element.textContent = text;
    element.hidden = !text;
  };

  const breakdownRow = (item) => {
    const row = doc.createElement("li");
    row.className = "breakdown__row";
    const name = row.appendChild(doc.createElement("span"));
    name.className = "breakdown__name";
    name.textContent = item.name || "Shape";
    name.title = item.name;
    const count = row.appendChild(doc.createElement("span"));
    count.className = "breakdown__count";
    count.textContent = `${numberFormat.format(item.counts.words)} ${item.counts.words === 1 ? "word" : "words"}`;
    return row;
  };

  return {
    /** @param {Summary} summary */
    render(summary) {
      status.hidden = true;
      const key = JSON.stringify(summary);
      if (key === lastRendered) return; // polling re-reads often; keep the DOM and screen readers quiet
      lastRendered = key;
      const description = describeSelection(summary);
      title.textContent = description.title;
      setText(detail, description.detail);
      setText(note, description.note);
      pane.classList.toggle("is-empty", description.empty);
      for (const stat of STATS) {
        values[stat].textContent = description.empty ? "–" : numberFormat.format(summary.total[stat]);
      }
      const rows = summary.items.length >= 2 ? summary.items : [];
      list.replaceChildren(...rows.map(breakdownRow));
      breakdown.hidden = rows.length === 0;
    },

    showError() {
      status.hidden = false;
    },

    showNotice(text) {
      setText(notice, text);
    },

    showMessage(heading, text) {
      pane.classList.add("is-message");
      title.textContent = heading;
      setText(detail, text);
      stats.hidden = true;
      $(".footer__refresh").hidden = true;
    },

    onRefresh(handler) {
      $(".footer__refresh").addEventListener("click", handler);
      $(".status__retry").addEventListener("click", handler);
    },
  };
}
```

- [ ] **Step 4: Run test to verify it passes**

Run: `node --test test/app.test.js`
Expected: PASS.

- [ ] **Step 5: Write** `src/taskpane.html`

```html
<!doctype html>
<html lang="en">
  <head>
    <meta charset="utf-8" />
    <meta name="viewport" content="width=device-width, initial-scale=1" />
    <title>Word Count</title>
    <!-- Office.js must be loaded in <head>, before the add-in's own scripts. -->
    <script src="https://officeapis.public.onecdn.static.microsoft/1/office.js"></script>
    <link rel="stylesheet" href="taskpane.css" />
    <script type="module" src="taskpane.js"></script>
  </head>
  <body>
    <div id="app"></div>
  </body>
</html>
```

- [ ] **Step 6: Write** `src/taskpane.js`

```js
// Entry point: waits for Office.js, then starts the pane.
import { start } from "./app.js";

const root = document.getElementById("app");

if (typeof Office === "undefined") {
  root.textContent = "Word Count couldn't load Office.js. Check your internet connection, then close and reopen the pane.";
} else {
  Office.onReady((info) => start({ Office, PowerPoint: globalThis.PowerPoint, root, info }));
}
```

- [ ] **Step 7: Write** `src/taskpane.css`

```css
/* Word Count task pane and help page.
   Colours are tokens. PowerPoint's own UI colours, set by app.js as --office-bg
   and --office-fg, win when available so the pane blends into PowerPoint. */

:root {
  color-scheme: light;
  --bg: #ffffff;
  --fg: #242424;
  --accent: #1d5bd8;
  --danger: #b42318;
  --pane-bg: var(--office-bg, var(--bg));
  --pane-fg: var(--office-fg, var(--fg));
  --muted: color-mix(in srgb, var(--pane-fg) 64%, var(--pane-bg));
  --surface: color-mix(in srgb, var(--pane-fg) 5%, var(--pane-bg));
  --border: color-mix(in srgb, var(--pane-fg) 14%, var(--pane-bg));
  --font: -apple-system, BlinkMacSystemFont, "Segoe UI Variable Text", "Segoe UI", system-ui, sans-serif;
}

@media (prefers-color-scheme: dark) {
  :root:not([data-theme="light"]) {
    color-scheme: dark;
    --bg: #1f1f1f;
    --fg: #f2f2f2;
    --accent: #8cb4ff;
    --danger: #ff8f85;
  }
}

:root[data-theme="dark"] {
  color-scheme: dark;
  --bg: #1f1f1f;
  --fg: #f2f2f2;
  --accent: #8cb4ff;
  --danger: #ff8f85;
}

*,
*::before,
*::after {
  box-sizing: border-box;
}

[hidden] {
  display: none !important;
}

html,
body {
  margin: 0;
}

body {
  background: var(--pane-bg);
  color: var(--pane-fg);
  font: 14px/1.45 var(--font);
  -webkit-font-smoothing: antialiased;
}

/* ---------- Task pane ---------- */

.pane {
  display: flex;
  flex-direction: column;
  gap: 16px;
  min-height: 100vh;
  padding: 16px;
}

.context__title {
  margin: 0;
  font-size: 15px;
  font-weight: 600;
  line-height: 1.3;
  overflow-wrap: anywhere;
}

.context__detail {
  margin: 2px 0 0;
  color: var(--muted);
  font-size: 13px;
}

.stats {
  display: grid;
  grid-template-columns: 1fr 1fr;
  gap: 8px;
  margin: 0;
}

.stat {
  display: flex;
  flex-direction: column-reverse; /* number above its label; the label is still read first */
  justify-content: flex-end;
  gap: 4px;
  min-width: 0;
  padding: 12px;
  border: 1px solid var(--border);
  border-radius: 8px;
  background: var(--surface);
}

.stat--words {
  grid-column: 1 / -1;
  padding: 14px 16px 16px;
}

.stat__value {
  margin: 0;
  font-size: 24px;
  font-weight: 600;
  line-height: 1.1;
  font-variant-numeric: tabular-nums;
  letter-spacing: -0.01em;
  overflow-wrap: anywhere;
}

.stat--words .stat__value {
  font-size: 44px;
  color: var(--accent);
}

.stat__label {
  color: var(--muted);
  font-size: 12px;
  line-height: 1.3;
}

.stat__qualifier {
  display: block;
}

.stat--words .stat__label {
  font-size: 13px;
}

.is-empty .stat__value {
  color: var(--muted);
}

.note {
  margin: 0;
  padding: 8px 12px;
  border-left: 3px solid var(--accent);
  border-radius: 0 6px 6px 0;
  background: var(--surface);
  color: var(--muted);
  font-size: 12px;
}

.breakdown__title {
  margin: 0 0 4px;
  color: var(--muted);
  font-size: 12px;
  font-weight: 600;
}

.breakdown__list {
  margin: 0;
  padding: 0;
  list-style: none;
  border-top: 1px solid var(--border);
}

.breakdown__row {
  display: flex;
  justify-content: space-between;
  gap: 12px;
  padding: 7px 0;
  border-bottom: 1px solid var(--border);
  font-size: 13px;
}

.breakdown__name {
  min-width: 0;
  overflow: hidden;
  text-overflow: ellipsis;
  white-space: nowrap;
}

.breakdown__count {
  flex: none;
  color: var(--muted);
  font-variant-numeric: tabular-nums;
}

.status {
  margin: 0;
  color: var(--danger);
  font-size: 13px;
}

.notice {
  margin: 0;
  color: var(--muted);
  font-size: 12px;
}

.footer {
  display: flex;
  flex-wrap: wrap;
  justify-content: space-between;
  gap: 4px 12px;
  margin-top: auto;
  padding-top: 12px;
  border-top: 1px solid var(--border);
  color: var(--muted);
  font-size: 12px;
}

.footer__actions {
  display: flex;
  gap: 12px;
}

.link-button {
  padding: 0;
  border: 0;
  background: none;
  color: var(--accent);
  font: inherit;
  text-decoration: none;
  cursor: pointer;
}

.link-button:hover {
  text-decoration: underline;
}

.link-button:focus-visible {
  outline: 2px solid var(--accent);
  outline-offset: 2px;
  border-radius: 2px;
}

@media (max-width: 250px) {
  .stats {
    grid-template-columns: 1fr;
  }
}

/* ---------- Help page ---------- */

.help {
  max-width: 680px;
  margin: 0 auto;
  padding: 24px 16px 48px;
}

.help h1 {
  margin: 0 0 8px;
  font-size: 24px;
  line-height: 1.25;
}

.help h2 {
  margin: 32px 0 8px;
  font-size: 18px;
}

.help h3 {
  margin: 20px 0 6px;
  font-size: 15px;
}

.help p,
.help li {
  max-width: 65ch;
}

.help li + li {
  margin-top: 4px;
}

.help__lead {
  color: var(--muted);
  font-size: 16px;
}

.help code {
  padding: 1px 4px;
  border-radius: 4px;
  background: var(--surface);
  font-size: 0.92em;
  overflow-wrap: anywhere;
}

.help a {
  color: var(--accent);
}

.help__table {
  width: 100%;
  border-collapse: collapse;
  font-size: 13px;
}

.help__table th,
.help__table td {
  padding: 8px;
  border-bottom: 1px solid var(--border);
  text-align: left;
  vertical-align: top;
}

.help__download {
  display: inline-block;
  padding: 8px 14px;
  border-radius: 6px;
  background: var(--accent);
  color: var(--pane-bg) !important;
  font-weight: 600;
  text-decoration: none;
}
```

- [ ] **Step 8: Replace** `scripts/serve.mjs`

```js
// Local server for developing the add-in (https://localhost:3100).
//
// It reads the certificate that `npm run certs` installs in ~/.office-addin-dev-certs
// instead of calling office-addin-dev-certs itself, so starting the server never
// triggers a surprise keychain/password prompt.
//
// `--http` serves plain HTTP with no certificate. PowerPoint won't accept that, but
// it's enough for the browser preview at /dev/preview.html.
//
// URL layout (the same as the production build): src/ is served at the root,
// plus /assets/*, /dev/* and the dev manifest.

import { X509Certificate } from "node:crypto";
import { existsSync, readFileSync } from "node:fs";
import { readFile, stat } from "node:fs/promises";
import { createServer as createHttpServer } from "node:http";
import { createServer } from "node:https";
import { homedir } from "node:os";
import { extname, join, normalize, resolve, sep } from "node:path";
import { fileURLToPath } from "node:url";

const ROOT = resolve(fileURLToPath(new URL("..", import.meta.url)));
const PORT = Number(process.env.PORT) || 3100;
const CERT_DIR = join(homedir(), ".office-addin-dev-certs");
const MANIFEST = join(ROOT, "manifest.xml");
const SERVED_DIRS = ["src", "assets", "dev"].map((dir) => join(ROOT, dir) + sep);

const MIME = {
  ".html": "text/html; charset=utf-8",
  ".js": "text/javascript; charset=utf-8",
  ".mjs": "text/javascript; charset=utf-8",
  ".css": "text/css; charset=utf-8",
  ".json": "application/json; charset=utf-8",
  ".png": "image/png",
  ".svg": "image/svg+xml",
  ".ico": "image/x-icon",
  ".xml": "application/xml; charset=utf-8",
};

function loadCertificate() {
  const keyPath = join(CERT_DIR, "localhost.key");
  const certPath = join(CERT_DIR, "localhost.crt");
  if (!existsSync(keyPath) || !existsSync(certPath)) {
    fail("No development certificate was found.");
  }
  const cert = readFileSync(certPath);
  const expires = new Date(new X509Certificate(cert).validTo);
  if (expires <= new Date()) {
    fail(`The development certificate expired on ${expires.toDateString()}.`);
  }
  return { key: readFileSync(keyPath), cert };
}

function fail(reason) {
  console.error(`\n${reason}\n\nRun this once (it asks for your Mac password), then try again:\n\n  npm run certs\n`);
  process.exit(1);
}

function resolvePath(pathname) {
  if (pathname === "/") return join(ROOT, "src", "taskpane.html");
  if (pathname === "/manifest.xml" || pathname === "/word-count-manifest.xml") return MANIFEST;
  const top = pathname.split("/")[1];
  if (top === "assets" || top === "dev") return join(ROOT, pathname);
  return join(ROOT, "src", pathname);
}

async function handle(req, res) {
  let pathname;
  try {
    pathname = decodeURIComponent(new URL(req.url, "https://localhost").pathname);
  } catch {
    return send(res, 400, "Bad request");
  }
  const target = normalize(resolvePath(pathname));
  if (target !== MANIFEST && !SERVED_DIRS.some((dir) => target.startsWith(dir))) {
    return send(res, 404, "Not found");
  }
  try {
    if (!(await stat(target)).isFile()) return send(res, 404, "Not found");
    const body = await readFile(target);
    res.writeHead(200, {
      "Content-Type": MIME[extname(target)] ?? "application/octet-stream",
      // Office caches add-in files aggressively; never cache during development.
      "Cache-Control": "no-store",
    });
    res.end(body);
  } catch {
    send(res, 404, "Not found");
  }
}

function send(res, status, text) {
  res.writeHead(status, { "Content-Type": "text/plain; charset=utf-8" });
  res.end(text);
}

if (process.argv.includes("--http")) {
  createHttpServer(handle).listen(PORT, () => {
    console.log(`Browser preview: http://localhost:${PORT}/dev/preview.html`);
  });
} else {
  createServer(loadCertificate(), handle).listen(PORT, () => {
    console.log(`Word Count add-in is being served at https://localhost:${PORT}`);
    console.log("Leave this window open while you use the add-in. Press Ctrl+C to stop.");
  });
}
```

- [ ] **Step 9: Write** `dev/preview.html`

```html
<!doctype html>
<html lang="en">
  <head>
    <meta charset="utf-8" />
    <meta name="viewport" content="width=device-width, initial-scale=1" />
    <title>Word Count preview</title>
    <link rel="stylesheet" href="/taskpane.css" />
    <style>
      body.preview-body {
        background: color-mix(in srgb, var(--pane-fg) 10%, var(--pane-bg));
      }
      .preview {
        display: flex;
        flex-wrap: wrap;
        gap: 24px;
        align-items: flex-start;
        padding: 24px;
      }
      .preview-controls {
        display: grid;
        gap: 12px;
        width: 280px;
        font-size: 13px;
      }
      .preview-controls h1 {
        margin: 0;
        font-size: 18px;
      }
      .preview-controls p {
        margin: 0;
        color: var(--muted);
      }
      .preview-controls label {
        display: grid;
        gap: 4px;
        font-weight: 600;
      }
      .preview-controls select,
      .preview-controls button {
        font: inherit;
        padding: 6px 8px;
      }
      .preview-frame {
        width: 320px;
        height: 600px;
        overflow: auto;
        border: 1px solid var(--border);
        border-radius: 6px;
        background: var(--pane-bg);
        box-shadow: 0 6px 24px rgb(0 0 0 / 0.14);
      }
      .preview-frame #app {
        height: 100%;
      }
      .preview-frame .pane {
        min-height: 100%;
      }
    </style>
  </head>
  <body class="preview-body">
    <div class="preview">
      <form class="preview-controls" onsubmit="return false">
        <h1>Word Count preview</h1>
        <p>The real task pane code, running against a fake PowerPoint.</p>
        <label>Selection <select id="scenario"></select></label>
        <label
          >Office theme
          <select id="theme">
            <option value="none">None (follow the system)</option>
            <option value="light">Light</option>
            <option value="dark">Dark</option>
          </select>
        </label>
        <label
          >Platform
          <select id="platform">
            <option value="Mac">Mac</option>
            <option value="PC">Windows</option>
            <option value="OfficeOnline">Web (no polling)</option>
          </select>
        </label>
        <button type="button" id="type">Simulate typing in the first text box</button>
        <button type="button" id="fail">Make the next read fail (twice)</button>
      </form>
      <div class="preview-frame"><div id="app"></div></div>
    </div>
    <script type="module">
      import { createFakeHost } from "/dev/fake-powerpoint.js";
      import { SCENARIOS } from "/dev/scenarios.js";
      import { start } from "/app.js";

      const THEMES = {
        light: { bodyBackgroundColor: "#FFFFFF", bodyForegroundColor: "#242424", controlBackgroundColor: "#FFFFFF", controlForegroundColor: "#242424", isDarkTheme: false },
        dark: { bodyBackgroundColor: "#292929", bodyForegroundColor: "#F2F2F2", controlBackgroundColor: "#1F1F1F", controlForegroundColor: "#F2F2F2", isDarkTheme: true },
      };
      const params = new URLSearchParams(location.search);
      const state = {
        scenario: SCENARIOS[params.get("scenario")] ? params.get("scenario") : "textBox",
        theme: THEMES[params.get("theme")] ? params.get("theme") : "none",
        platform: ["Mac", "PC", "OfficeOnline"].includes(params.get("platform")) ? params.get("platform") : "Mac",
      };
      const saveState = () => history.replaceState(null, "", `?${new URLSearchParams(state)}`);

      // Copies, so "simulate typing" never edits the shared scenarios.
      const fake = createFakeHost(structuredClone(SCENARIOS[state.scenario]), {
        platform: state.platform,
        theme: THEMES[state.theme] ?? null,
        latencyMs: 25,
      });
      start({ Office: fake.Office, PowerPoint: fake.PowerPoint, root: document.getElementById("app"), info: { host: "PowerPoint", platform: state.platform } });

      const scenarioSelect = document.getElementById("scenario");
      for (const [key, { label }] of Object.entries(SCENARIOS)) scenarioSelect.add(new Option(label, key, false, key === state.scenario));
      scenarioSelect.addEventListener("change", () => {
        state.scenario = scenarioSelect.value;
        saveState();
        fake.host.setScenario(structuredClone(SCENARIOS[state.scenario]));
        fake.host.fireSelectionChanged();
      });

      // Theme and platform are read once at start-up, as in PowerPoint, so reload.
      for (const id of ["theme", "platform"]) {
        const select = document.getElementById(id);
        select.value = state[id];
        select.addEventListener("change", () => {
          state[id] = select.value;
          saveState();
          location.reload();
        });
      }

      const firstTextShape = (shapes) => {
        for (const shape of shapes) {
          if (typeof shape.text === "string") return shape;
          const inner = shape.shapes && firstTextShape(shape.shapes);
          if (inner) return inner;
        }
        return null;
      };
      document.getElementById("type").addEventListener("click", () => {
        const shape = firstTextShape(fake.host.scenario.selected);
        if (shape) shape.text += " More words";
        // No selection event: on Mac/Windows the pane's polling picks this up.
      });
      document.getElementById("fail").addEventListener("click", () => {
        fake.host.failNextSyncs(2);
        fake.host.fireSelectionChanged();
      });
    </script>
  </body>
</html>
```

- [ ] **Step 10: Check it in a browser**

Run: `npm run preview`, then open `http://localhost:3100/dev/preview.html`.
Expected: the pane shows "TextBox 3" and 9 words / 36 characters / 44 characters with spaces. Walk every scenario in the Selection menu and both Office themes. Check that "Simulate typing" updates the count within about 2 s on Mac/Windows and never on Web, and that "Make the next read fail" shows "Couldn't read the selection. Try again".

- [ ] **Step 11: Run all tests**

Run: `npm test`
Expected: PASS.

- [ ] **Step 12: Commit** (sequential executors only)

```bash
git add src/app.js src/taskpane.html src/taskpane.js src/taskpane.css dev/preview.html scripts/serve.mjs test/app.test.js
git commit -m "feat: add live task pane UI and browser preview"
```

---
### Task 6: Manifest and help page

**Files:**
- Create: `manifest.xml`
- Create: `src/help.html`
- Create: `test/manifest.test.js`

**Interfaces:**
- Consumes: `assets/icon-*.png` (committed), `src/taskpane.html` (Task 5), `src/taskpane.css` (Task 5; the help page uses its `.help*` styles).
- Produces: the dev manifest that `scripts/build.mjs` (Task 7) rewrites. It must contain `https://localhost:3100`, the dev Id and `DefaultValue="Word Count (dev)"`.

- [ ] **Step 1: Write the failing test** `test/manifest.test.js`

```js
import assert from "node:assert/strict";
import { existsSync, readFileSync } from "node:fs";
import { test } from "node:test";

const root = new URL("../", import.meta.url);
// Comments are dropped so a mention of an element in a comment isn't mistaken for the element.
const xml = readFileSync(new URL("manifest.xml", root), "utf8").replace(/<!--[\s\S]*?-->/g, "");

test("targets PowerPoint with PowerPointApi 1.10 and read/write permission", () => {
  assert.match(xml, /<Host Name="Presentation"\s*\/>/);
  assert.match(xml, /<Set Name="PowerPointApi" MinVersion="1\.10"\s*\/>/);
  assert.match(xml, /<Permissions>ReadWriteDocument<\/Permissions>/);
});

test("uses the dev identity", () => {
  assert.match(xml, /<Id>25c6e699-85a5-4a2d-bc50-ecc66e7483be<\/Id>/);
  assert.match(xml, /<DisplayName DefaultValue="Word Count \(dev\)"\s*\/>/);
});

test("every web address points at https://localhost:3100", () => {
  const urls = [...xml.matchAll(/(?:DefaultValue="|<AppDomain>)(https?:\/\/[^"<]+)/g)].map((m) => m[1]);
  assert.ok(urls.length >= 10, `only ${urls.length} URLs`);
  for (const url of urls) assert.ok(url.startsWith("https://localhost:3100"), url);
});

test("every referenced icon exists", () => {
  const icons = [...xml.matchAll(/https:\/\/localhost:3100\/(assets\/[^"]+\.png)/g)].map((m) => m[1]);
  assert.ok(icons.length >= 5);
  for (const icon of new Set(icons)) assert.ok(existsSync(new URL(icon, root)), icon);
});

test("the ribbon button and group have the required 16, 32 and 80 px icons", () => {
  for (const size of [16, 32, 80]) assert.match(xml, new RegExp(`<bt:Image size="${size}" resid="Icon\\.${size}"\\s*/>`));
});

test("the pages it points at exist in src/", () => {
  for (const page of ["taskpane.html", "help.html"]) {
    assert.ok(xml.includes(`https://localhost:3100/${page}`), page);
    assert.ok(existsSync(new URL(`src/${page}`, root)), page);
  }
});

test("top-level elements are in the order the schema requires", () => {
  const order = [
    "Id",
    "Version",
    "ProviderName",
    "DefaultLocale",
    "DisplayName",
    "Description",
    "IconUrl",
    "HighResolutionIconUrl",
    "SupportUrl",
    "AppDomains",
    "Hosts",
    "Requirements",
    "DefaultSettings",
    "Permissions",
    "VersionOverrides",
  ];
  const positions = order.map((name) => xml.indexOf(`<${name}`));
  positions.forEach((position, i) => assert.notEqual(position, -1, `${order[i]} is missing`));
  assert.deepEqual(
    [...positions].sort((a, b) => a - b),
    positions,
  );
});
```

- [ ] **Step 2: Run test to verify it fails**

Run: `node --test test/manifest.test.js`
Expected: FAIL with `ENOENT: no such file or directory, open '.../manifest.xml'`.

- [ ] **Step 3: Write** `manifest.xml`

```xml
<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<!--
  Development manifest (https://localhost:3100). `npm run build` turns it into the
  production manifest (dist/word-count-manifest.xml) with the hosting address, its
  own Id and the name "Word Count". Bump <Version> whenever this file changes.
-->
<OfficeApp xmlns="http://schemas.microsoft.com/office/appforoffice/1.1"
  xmlns:xsi="http://www.w3.org/2001/XMLSchema-instance"
  xmlns:bt="http://schemas.microsoft.com/office/officeappbasictypes/1.0"
  xmlns:ov="http://schemas.microsoft.com/office/taskpaneappversionoverrides"
  xsi:type="TaskPaneApp">
  <Id>25c6e699-85a5-4a2d-bc50-ecc66e7483be</Id>
  <Version>1.0.0.0</Version>
  <ProviderName>Joe Zandstra</ProviderName>
  <DefaultLocale>en-US</DefaultLocale>
  <DisplayName DefaultValue="Word Count (dev)"/>
  <Description DefaultValue="Live word, character and characters-with-spaces counts for the selected text, shapes, tables and groups, counted the way Microsoft Word counts."/>
  <IconUrl DefaultValue="https://localhost:3100/assets/icon-32.png"/>
  <HighResolutionIconUrl DefaultValue="https://localhost:3100/assets/icon-64.png"/>
  <SupportUrl DefaultValue="https://localhost:3100/help.html"/>
  <AppDomains>
    <AppDomain>https://localhost:3100</AppDomain>
  </AppDomains>
  <Hosts>
    <Host Name="Presentation"/>
  </Hosts>
  <Requirements>
    <Sets DefaultMinVersion="1.1">
      <Set Name="PowerPointApi" MinVersion="1.10"/>
    </Sets>
  </Requirements>
  <DefaultSettings>
    <SourceLocation DefaultValue="https://localhost:3100/taskpane.html"/>
  </DefaultSettings>
  <!-- Read/write is required by PowerPoint.run and selection events; the add-in never changes the presentation. -->
  <Permissions>ReadWriteDocument</Permissions>
  <VersionOverrides xmlns="http://schemas.microsoft.com/office/taskpaneappversionoverrides" xsi:type="VersionOverridesV1_0">
    <Hosts>
      <Host xsi:type="Presentation">
        <DesktopFormFactor>
          <GetStarted>
            <Title resid="GetStarted.Title"/>
            <Description resid="GetStarted.Description"/>
            <LearnMoreUrl resid="Help.Url"/>
          </GetStarted>
          <ExtensionPoint xsi:type="PrimaryCommandSurface">
            <OfficeTab id="TabHome">
              <Group id="WordCount.Group">
                <Label resid="Group.Label"/>
                <Icon>
                  <bt:Image size="16" resid="Icon.16"/>
                  <bt:Image size="32" resid="Icon.32"/>
                  <bt:Image size="80" resid="Icon.80"/>
                </Icon>
                <Control xsi:type="Button" id="WordCount.ShowTaskpane">
                  <Label resid="Button.Label"/>
                  <Supertip>
                    <Title resid="Button.Label"/>
                    <Description resid="Button.Tooltip"/>
                  </Supertip>
                  <Icon>
                    <bt:Image size="16" resid="Icon.16"/>
                    <bt:Image size="20" resid="Icon.20"/>
                    <bt:Image size="24" resid="Icon.24"/>
                    <bt:Image size="32" resid="Icon.32"/>
                    <bt:Image size="40" resid="Icon.40"/>
                    <bt:Image size="48" resid="Icon.48"/>
                    <bt:Image size="64" resid="Icon.64"/>
                    <bt:Image size="80" resid="Icon.80"/>
                  </Icon>
                  <Action xsi:type="ShowTaskpane">
                    <TaskpaneId>WordCountTaskpane</TaskpaneId>
                    <SourceLocation resid="Taskpane.Url"/>
                  </Action>
                </Control>
              </Group>
            </OfficeTab>
          </ExtensionPoint>
        </DesktopFormFactor>
      </Host>
    </Hosts>
    <Resources>
      <bt:Images>
        <bt:Image id="Icon.16" DefaultValue="https://localhost:3100/assets/icon-16.png"/>
        <bt:Image id="Icon.20" DefaultValue="https://localhost:3100/assets/icon-20.png"/>
        <bt:Image id="Icon.24" DefaultValue="https://localhost:3100/assets/icon-24.png"/>
        <bt:Image id="Icon.32" DefaultValue="https://localhost:3100/assets/icon-32.png"/>
        <bt:Image id="Icon.40" DefaultValue="https://localhost:3100/assets/icon-40.png"/>
        <bt:Image id="Icon.48" DefaultValue="https://localhost:3100/assets/icon-48.png"/>
        <bt:Image id="Icon.64" DefaultValue="https://localhost:3100/assets/icon-64.png"/>
        <bt:Image id="Icon.80" DefaultValue="https://localhost:3100/assets/icon-80.png"/>
      </bt:Images>
      <bt:Urls>
        <bt:Url id="Taskpane.Url" DefaultValue="https://localhost:3100/taskpane.html"/>
        <bt:Url id="Help.Url" DefaultValue="https://localhost:3100/help.html"/>
      </bt:Urls>
      <bt:ShortStrings>
        <bt:String id="GetStarted.Title" DefaultValue="Word Count is ready"/>
        <bt:String id="Group.Label" DefaultValue="Word Count"/>
        <bt:String id="Button.Label" DefaultValue="Word Count"/>
      </bt:ShortStrings>
      <bt:LongStrings>
        <bt:String id="GetStarted.Description" DefaultValue="Choose Word Count on the Home tab, then select some text, a shape, a table or a group."/>
        <bt:String id="Button.Tooltip" DefaultValue="Show live word and character counts for whatever is selected."/>
      </bt:LongStrings>
    </Resources>
  </VersionOverrides>
</OfficeApp>
```

- [ ] **Step 4: Write** `src/help.html`

```html
<!doctype html>
<html lang="en">
  <head>
    <meta charset="utf-8" />
    <meta name="viewport" content="width=device-width, initial-scale=1" />
    <title>Word Count for PowerPoint – help</title>
    <link rel="stylesheet" href="taskpane.css" />
  </head>
  <body>
    <main class="help">
      <h1>Word Count for PowerPoint</h1>
      <p class="help__lead">See how many words and characters are in whatever you've selected: highlighted text, text boxes, placeholders, tables and groups.</p>

      <h2>Using it</h2>
      <ol>
        <li>On the <strong>Home</strong> tab, choose <strong>Word Count</strong>. A pane opens on the right.</li>
        <li>Select something. The counts update by themselves.</li>
      </ol>
      <table class="help__table">
        <thead>
          <tr>
            <th scope="col">You select…</th>
            <th scope="col">What's counted</th>
          </tr>
        </thead>
        <tbody>
          <tr>
            <td>Some text inside a box (highlighted)</td>
            <td>Only the highlighted text</td>
          </tr>
          <tr>
            <td>One or more text boxes or placeholders</td>
            <td>All of their text, added together</td>
          </tr>
          <tr>
            <td>A table</td>
            <td>Every cell (merged cells once)</td>
          </tr>
          <tr>
            <td>A group</td>
            <td>Every text box and table inside it, however deeply grouped</td>
          </tr>
        </tbody>
      </table>
      <p>When you select several shapes, the pane also lists each one with its word count.</p>

      <h2>How words and characters are counted</h2>
      <p>The counts follow the same rules as Microsoft Word's Word Count, so you get the same numbers Word would give for the same text.</p>
      <ul>
        <li><strong>Words</strong> are separated by spaces. A hyphenated word (<em>well-known</em>), a web address or a number such as <em>3.14</em> is one word. An en or em dash between words (<em>before—after</em>) splits them. In Chinese and Japanese, each character counts as a word.</li>
        <li><strong>Characters (no spaces)</strong> are letters, numbers, punctuation and symbols. An emoji counts as one character.</li>
        <li><strong>Characters (with spaces)</strong> are the same, plus spaces and tabs. Line and paragraph breaks aren't counted.</li>
      </ul>

      <h2>Known limitations</h2>
      <ul>
        <li><strong>Charts, SmartArt and embedded objects</strong> can't be counted, because PowerPoint doesn't let add-ins read their text. The pane tells you when one has been left out.</li>
        <li><strong>A cursor inside a word:</strong> on Mac and Windows, when the cursor is inside a word and nothing is highlighted, PowerPoint may report that word as selected, so the pane counts just that word. Click the border of the box so the whole box is selected to count all of it.</li>
        <li><strong>Slide thumbnails hidden:</strong> if the slide thumbnails are hidden, PowerPoint may not report highlighted text, and the whole box is counted instead.</li>
        <li><strong>Part of a table cell:</strong> highlighting part of a cell may count the whole table, depending on your version of PowerPoint.</li>
        <li><strong>Slide master and layouts:</strong> PowerPoint doesn't report selections there, so they can't be counted.</li>
        <li><strong>Automatic bullets and numbering</strong> aren't counted. Word counts them in documents, so a Word count of the same bulleted text can be slightly higher.</li>
      </ul>

      <h2 id="install">Installing Word Count</h2>
      <p>Word Count needs a current version of PowerPoint: 16.105 or later on Mac, Version 2601 or later on Windows, or PowerPoint on the web.</p>
      <p><a class="help__download" href="word-count-manifest.xml" download>Download the add-in file</a></p>

      <h3>For a whole organisation (recommended)</h3>
      <p>
        A Microsoft 365 administrator can install it for everyone at once. In the Microsoft 365 admin center, go to <strong>Settings › Integrated apps › Upload custom apps</strong>, choose <strong>Office Add-in</strong>, and upload the file or paste this page's address with
        <code>help.html</code> replaced by <code>word-count-manifest.xml</code>. It appears on the Home tab within a day, on Mac, Windows and the web.
      </p>

      <h3>On a Mac</h3>
      <ol>
        <li>Quit PowerPoint.</li>
        <li>In Finder, choose <strong>Go › Go to Folder…</strong>, paste <code>~/Library/Containers/com.microsoft.Powerpoint/Data/Documents/wef</code> and press Return. If Finder can't find it, go to <code>…/Data/Documents</code> instead and create a folder named <code>wef</code> there.</li>
        <li>Copy the downloaded <code>word-count-manifest.xml</code> into that folder. Leave anything else in the folder alone.</li>
        <li>Open PowerPoint and a presentation, choose <strong>Home › Add-ins</strong>, and pick <strong>Word Count</strong>.</li>
      </ol>

      <h3>On Windows</h3>
      <ol>
        <li>In PowerPoint, choose <strong>Home › Add-ins › More Add-ins</strong>.</li>
        <li>Open the <strong>My Add-ins</strong> tab, choose <strong>Manage My Add-ins › Upload My Add-in</strong>, and pick the downloaded file.</li>
      </ol>

      <h3>In PowerPoint on the web</h3>
      <ol>
        <li>Open a presentation and choose <strong>Home › Add-ins › More Settings</strong> (in some versions, <strong>Insert › Add-ins › My Add-ins</strong>).</li>
        <li>Choose <strong>Upload My Add-in</strong> and pick the downloaded file. Your browser remembers it, so you'll need to do this again in a different browser or after clearing your browsing data.</li>
      </ol>

      <h2>Troubleshooting</h2>
      <ul>
        <li><strong>The numbers look stuck:</strong> choose <strong>Refresh</strong> at the bottom of the pane.</li>
        <li><strong>The pane looks out of date after an update:</strong> on a Mac, open the small menu at the top of the pane and choose <strong>Clear Web Cache</strong>, then reopen the pane.</li>
        <li><strong>It says "Please update PowerPoint":</strong> install the latest Office updates (on a Mac: <strong>Help › Check for Updates</strong>).</li>
      </ul>
    </main>
  </body>
</html>
```

- [ ] **Step 5: Run the tests**

Run: `node --test test/manifest.test.js`
Expected: PASS. It needs `src/taskpane.html` from Task 5.

- [ ] **Step 6: Validate with Microsoft's service** (needs internet; it sends the manifest to Microsoft)

Run: `npm run validate`
Expected: `The manifest is valid.`

- [ ] **Step 7: Commit** (sequential executors only)

```bash
git add manifest.xml src/help.html test/manifest.test.js
git commit -m "feat: add dev manifest and help page"
```

---

### Task 7: Production build and GitHub Pages deploy

**Files:**
- Create: `scripts/build.mjs`
- Create: `test/build.test.js`
- Create: `.github/workflows/deploy.yml`

**Interfaces:**
- Consumes: `manifest.xml` (Task 6); everything in `src/` and `assets/`.
- Produces: `build({ baseUrl, root, outDir })` → `{ outDir, version, manifestPath }`; also exports `normalizeBaseUrl(baseUrl)`, `productionManifest(devXml, baseUrl)`, `DEV_ORIGIN`, `DEV_ID`, `PROD_ID`, `MANIFEST_NAME`.

- [ ] **Step 1: Write the failing test** `test/build.test.js`

```js
import assert from "node:assert/strict";
import { existsSync, readFileSync } from "node:fs";
import { mkdtemp, rm } from "node:fs/promises";
import { tmpdir } from "node:os";
import { join } from "node:path";
import { after, before, test } from "node:test";

import { build, DEV_ID, normalizeBaseUrl, PROD_ID, productionManifest } from "../scripts/build.mjs";

const BASE = "https://example.github.io/ppt-word-count/";
let outDir;
let result;

before(async () => {
  outDir = await mkdtemp(join(tmpdir(), "word-count-build-"));
  result = await build({ baseUrl: BASE, outDir });
});
after(() => rm(outDir, { recursive: true, force: true }));

const read = (path) => readFileSync(join(outDir, path), "utf8");

test("copies the pane, help page, icons and production manifest, but not dev files", () => {
  const expected = [
    "taskpane.html",
    "taskpane.css",
    "taskpane.js",
    "app.js",
    "count.js",
    "selection.js",
    "refresher.js",
    "help.html",
    "assets/icon-16.png",
    "assets/icon-80.png",
    "word-count-manifest.xml",
  ];
  for (const file of expected) assert.ok(existsSync(join(outDir, file)), file);
  assert.ok(!existsSync(join(outDir, "dev")));
  assert.ok(!existsSync(join(outDir, "manifest.xml")));
});

test("the production manifest points at the hosting address with its own identity", () => {
  const manifest = read("word-count-manifest.xml");
  assert.match(manifest, /<SourceLocation DefaultValue="https:\/\/example\.github\.io\/ppt-word-count\/taskpane\.html"\s*\/>/);
  assert.match(manifest, /<AppDomain>https:\/\/example\.github\.io<\/AppDomain>/);
  assert.ok(manifest.includes(`<Id>${PROD_ID}</Id>`));
  assert.ok(!manifest.includes(DEV_ID));
  assert.match(manifest, /<DisplayName DefaultValue="Word Count"\s*\/>/);
  assert.doesNotMatch(manifest, /localhost/i);
});

test("every address in the production manifest has a matching file", () => {
  const manifest = read("word-count-manifest.xml");
  const urls = [...manifest.matchAll(/DefaultValue="(https:\/\/[^"]+)"/g)].map((m) => m[1]);
  assert.ok(urls.length >= 10);
  for (const url of urls) {
    assert.ok(url.startsWith(BASE), url);
    assert.ok(existsSync(join(outDir, url.slice(BASE.length))), url);
  }
});

test("local scripts, styles and module imports are version-stamped; the Office.js CDN is not", () => {
  const v = result.version;
  assert.match(v, /^[0-9a-f]{10}$/);
  const html = read("taskpane.html");
  assert.ok(html.includes(`href="taskpane.css?v=${v}"`));
  assert.ok(html.includes(`src="taskpane.js?v=${v}"`));
  assert.ok(html.includes('src="https://officeapis.public.onecdn.static.microsoft/1/office.js"'));
  assert.ok(read("help.html").includes(`href="taskpane.css?v=${v}"`));
  assert.ok(read("taskpane.js").includes(`from "./app.js?v=${v}"`));
  const app = read("app.js");
  for (const module of ["./count.js", "./refresher.js", "./selection.js"]) assert.ok(app.includes(`"${module}?v=${v}"`), module);
});

test("the same sources always give the same version", async () => {
  const again = await mkdtemp(join(tmpdir(), "word-count-build-"));
  try {
    const second = await build({ baseUrl: "https://example.github.io/ppt-word-count", outDir: again });
    assert.equal(second.version, result.version);
  } finally {
    await rm(again, { recursive: true, force: true });
  }
});

test("rejects addresses Office can't use", () => {
  assert.throws(() => normalizeBaseUrl("http://example.com"), /https/);
  assert.throws(() => normalizeBaseUrl(""), /BASE_URL/);
  assert.throws(() => normalizeBaseUrl(undefined), /BASE_URL/);
  assert.throws(() => normalizeBaseUrl("https://localhost:3100"), /localhost/);
  assert.equal(normalizeBaseUrl("https://example.com/a//"), "https://example.com/a");
});

test("refuses a dev manifest it no longer recognises", () => {
  assert.throws(() => productionManifest("<OfficeApp/>", "https://example.com"), /manifest\.xml/);
});
```

- [ ] **Step 2: Run test to verify it fails**

Run: `node --test test/build.test.js`
Expected: FAIL with `Cannot find module '.../scripts/build.mjs'`.

- [ ] **Step 3: Write the implementation** `scripts/build.mjs`

```js
// Builds dist/: the files to host (for example on GitHub Pages) plus the
// production manifest that colleagues install.
//
//   BASE_URL=https://you.github.io/ppt-word-count npm run build
//
// The production manifest is manifest.xml with the hosting address, its own Id
// (so it can sit next to the dev copy) and the name "Word Count". Every local
// script and stylesheet reference gets ?v=<content hash>, because GitHub Pages
// lets browsers cache files for 10 minutes and a deploy could otherwise mix old
// and new files.

import { createHash } from "node:crypto";
import { mkdir, readdir, readFile, rm, writeFile } from "node:fs/promises";
import { dirname, extname, join, relative, resolve } from "node:path";
import { fileURLToPath, pathToFileURL } from "node:url";

const ROOT = resolve(fileURLToPath(new URL("..", import.meta.url)));
export const DEV_ORIGIN = "https://localhost:3100";
export const DEV_ID = "25c6e699-85a5-4a2d-bc50-ecc66e7483be";
export const PROD_ID = "1b331966-f37d-4165-9499-7f94f173c9c1";
export const MANIFEST_NAME = "word-count-manifest.xml";

/** @returns {string} the address without a trailing slash */
export function normalizeBaseUrl(baseUrl) {
  let url;
  try {
    url = new URL(String(baseUrl ?? ""));
  } catch {
    throw new Error(`BASE_URL must be a full web address, like https://you.github.io/ppt-word-count (got "${baseUrl ?? ""}").`);
  }
  if (url.protocol !== "https:") throw new Error("BASE_URL must start with https:// because Office only loads add-ins over HTTPS.");
  if (url.hostname === "localhost" || url.hostname === "127.0.0.1") throw new Error("BASE_URL can't be localhost; use the hosting address.");
  if (url.search || url.hash) throw new Error("BASE_URL can't contain ? or #.");
  return url.href.replace(/\/+$/, "");
}

/** Turns the dev manifest into the production one. */
export function productionManifest(devManifest, baseUrl) {
  for (const required of [DEV_ORIGIN, DEV_ID, 'DefaultValue="Word Count (dev)"']) {
    if (!devManifest.includes(required)) throw new Error(`manifest.xml no longer contains ${required}; update scripts/build.mjs to match.`);
  }
  const manifest = devManifest
    .replaceAll(`<AppDomain>${DEV_ORIGIN}</AppDomain>`, `<AppDomain>${new URL(baseUrl).origin}</AppDomain>`)
    .replaceAll(DEV_ORIGIN, baseUrl)
    .replaceAll(DEV_ID, PROD_ID)
    .replaceAll('DefaultValue="Word Count (dev)"', 'DefaultValue="Word Count"');
  if (/localhost/i.test(manifest)) throw new Error("The production manifest still mentions localhost.");
  return manifest;
}

function addVersion(source, extension, version) {
  if (extension === ".html") {
    // Local .js/.css references only: absolute URLs (the Office.js CDN) are left alone.
    return source.replace(/\b(src|href)="(?![a-z][a-z0-9+.-]*:|\/\/|#)([^"?#]+\.(?:js|css))"/gi, `$1="$2?v=${version}"`);
  }
  if (extension === ".js") {
    return source.replace(/(\bfrom\s*|\bimport\s*\(\s*|\bimport\s+)(["'])(\.{1,2}\/[^"'?#]+\.js)\2/g, `$1$2$3?v=${version}$2`);
  }
  return source;
}

async function listFiles(dir) {
  const entries = await readdir(dir, { withFileTypes: true });
  const nested = await Promise.all(entries.map((entry) => (entry.isDirectory() ? listFiles(join(dir, entry.name)) : [join(dir, entry.name)])));
  return nested.flat().sort();
}

/**
 * @param {{ baseUrl: string, root?: string, outDir?: string }} options
 * @returns {Promise<{ outDir: string, version: string, manifestPath: string }>}
 */
export async function build({ baseUrl, root = ROOT, outDir = join(root, "dist") }) {
  const base = normalizeBaseUrl(baseUrl);
  const sources = [
    ...(await listFiles(join(root, "src"))).map((file) => ({ file, target: relative(join(root, "src"), file) })),
    ...(await listFiles(join(root, "assets"))).map((file) => ({ file, target: join("assets", relative(join(root, "assets"), file)) })),
  ];
  const contents = await Promise.all(sources.map(({ file }) => readFile(file)));

  const hash = createHash("sha256");
  sources.forEach(({ target }, i) => hash.update(target).update("\0").update(contents[i]).update("\0"));
  const version = hash.digest("hex").slice(0, 10);

  await rm(outDir, { recursive: true, force: true });
  for (const [i, { target }] of sources.entries()) {
    const destination = join(outDir, target);
    await mkdir(dirname(destination), { recursive: true });
    const extension = extname(target);
    const body = extension === ".html" || extension === ".js" ? addVersion(contents[i].toString("utf8"), extension, version) : contents[i];
    await writeFile(destination, body);
  }

  const manifestPath = join(outDir, MANIFEST_NAME);
  await writeFile(manifestPath, productionManifest(await readFile(join(root, "manifest.xml"), "utf8"), base));
  return { outDir, version, manifestPath };
}

if (process.argv[1] && import.meta.url === pathToFileURL(resolve(process.argv[1])).href) {
  try {
    const { outDir, version, manifestPath } = await build({ baseUrl: process.env.BASE_URL });
    console.log(`Built ${relative(ROOT, outDir)}/ (version ${version}).`);
    console.log(`Production manifest: ${relative(ROOT, manifestPath)}`);
  } catch (error) {
    console.error(error.message);
    process.exitCode = 1;
  }
}
```

- [ ] **Step 4: Run test to verify it passes**

Run: `node --test test/build.test.js`
Expected: PASS.

- [ ] **Step 5: Write** `.github/workflows/deploy.yml`

```yaml
# Builds the add-in and publishes it to GitHub Pages on every push to main.
# One-time setup: repository Settings > Pages > Build and deployment > Source: GitHub Actions.
name: Deploy to GitHub Pages

on:
  push:
    branches: [main]
  workflow_dispatch:

permissions:
  contents: read
  pages: write
  id-token: write

concurrency:
  group: pages
  cancel-in-progress: false

jobs:
  build:
    runs-on: ubuntu-latest
    steps:
      - uses: actions/checkout@v7
      - uses: actions/setup-node@v7
        with:
          node-version: 24
      - id: pages
        uses: actions/configure-pages@v6
      # The add-in has no runtime dependencies, so tests and the build need no npm install.
      - run: npm test
      - run: npm run build
        env:
          BASE_URL: ${{ steps.pages.outputs.base_url }}
      # Schema check by Microsoft's online validator. (Its "-p" store mode also fetches
      # the icons from the live site, which fails before the first deploy.)
      - name: Check the production manifest with Microsoft's validator
        run: npx --yes office-addin-manifest@3 validate dist/word-count-manifest.xml
        continue-on-error: true
      - uses: actions/upload-pages-artifact@v5
        with:
          path: dist

  deploy:
    needs: build
    runs-on: ubuntu-latest
    environment:
      name: github-pages
      url: ${{ steps.deployment.outputs.page_url }}
    steps:
      - id: deployment
        uses: actions/deploy-pages@v5
```

- [ ] **Step 6: Try a local build**

Run: `BASE_URL=https://example.github.io/ppt-word-count npm run build && ls dist`
Expected: `Built dist/ (version …)`; the listing shows `app.js assets count.js help.html refresher.js selection.js taskpane.css taskpane.html taskpane.js word-count-manifest.xml`.

- [ ] **Step 7: Commit** (sequential executors only)

```bash
git add scripts/build.mjs test/build.test.js .github/workflows/deploy.yml
git commit -m "feat: add production build and GitHub Pages deploy"
```

---

### Task 8: README and final verification

**Files:**
- Create: `README.md`

**Interfaces:**
- Consumes: every npm script in `package.json`: `certs`, `serve`, `preview`, `sideload`, `start`, `stop`, `test`, `validate`, `build`.
- Produces: user-facing documentation. No code.

- [ ] **Step 1: Write** `README.md`

````markdown
# Word Count for PowerPoint

A PowerPoint add-in that shows **Words**, **Characters (no spaces)** and **Characters (with spaces)** for whatever you select: highlighted text, text boxes, placeholders, tables and groups. It updates as you click around, and it counts the same way Microsoft Word does.

Works in current PowerPoint for Mac (16.105+), Windows (Version 2601+) and PowerPoint on the web. How the counts work, the known limitations and install steps for colleagues are all on the add-in's help page ([`src/help.html`](src/help.html)).

## Try it on this Mac

You only need to do step 1 once a year.

1. **Trust the local certificate.** PowerPoint only loads add-ins over a secure connection, so your Mac needs to trust a small local certificate. Run the command below. Your Mac asks for your password, possibly once in the terminal (nothing appears as you type; that's normal) and once in a pop-up window.

   ```bash
   npm run certs
   ```

2. **Start the add-in.** This starts a small local web server and opens PowerPoint with the add-in loaded.

   ```bash
   npm start
   ```

   In PowerPoint, choose **Home › Word Count** (it's called **Word Count (dev)** in this local version) and select some text.

3. **When you're done**, stop the server:

   ```bash
   npm stop
   ```

### Preview in a browser (no PowerPoint needed)

```bash
npm run preview
```

Then open <http://localhost:3100/dev/preview.html>. This runs the real pane against a pretend PowerPoint, with a menu of example selections.

## Share it with colleagues

Colleagues can't use the local version, so the add-in has to live on the web. The project is set up to publish itself to **GitHub Pages** for free.

1. **Put the project on GitHub** as a *public* repository named `ppt-word-count`. (On GitHub's free plan, Pages only works for public repositories. There's nothing private in this project.)
2. **Turn on Pages:** on GitHub, open the repository's **Settings › Pages** and set **Source** to **GitHub Actions**.
3. **Publish:** every time changes reach the `main` branch, GitHub builds and publishes the add-in; watch it under the **Actions** tab. The first time, if the run failed because Pages wasn't turned on yet, choose **Re-run all jobs**.
4. **Send colleagues the help page:** `https://<your-github-username>.github.io/ppt-word-count/help.html#install`. It has the download link and step-by-step install instructions for Mac, Windows and the web, and for a Microsoft 365 administrator who wants to install it for everyone.

### Updating

- **Changes to the pane** (anything in `src/`) reach everyone automatically after publishing, within about 10 minutes.
- **Changes to `manifest.xml`** (name, icons, ribbon button) also need the `<Version>` number in it raised (e.g. `1.0.0.0` → `1.0.1.0`), and colleagues need to install the new manifest file again.

## Troubleshooting

- **"The development certificate expired" or "No development certificate was found":** run `npm run certs` again.
- **The pane is blank or shows an old version:** open the small menu at the top of the pane and choose **Clear Web Cache**, then reopen it.
- **Word Count doesn't appear under Home › Add-ins:** quit PowerPoint completely and run `npm start` again.

## For developers

- `npm test`: unit tests (`node --test`, no dependencies needed).
- `npm run validate`: checks `manifest.xml` with Microsoft's online validator.
- `BASE_URL=https://example.github.io/ppt-word-count npm run build`: builds `dist/` for hosting, including `dist/word-count-manifest.xml`.
- Design and research: [`docs/superpowers/specs/2026-09-26-ppt-word-count-design.md`](docs/superpowers/specs/2026-09-26-ppt-word-count-design.md). Implementation plan: [`docs/superpowers/plans/2026-09-26-ppt-word-count.md`](docs/superpowers/plans/2026-09-26-ppt-word-count.md).

| Path | What it is |
|---|---|
| `src/count.js` | Word-compatible counting rules |
| `src/selection.js` | Reads the PowerPoint selection (the only code that talks to PowerPoint) |
| `src/refresher.js` | Decides when to re-read the selection |
| `src/app.js` | The pane's behaviour and drawing |
| `src/taskpane.*`, `src/help.html` | The pane page, styles and help page |
| `dev/` | Fake PowerPoint, example selections and the browser preview |
| `scripts/` | Local server, production build, icon generator |
| `test/` | Unit tests, including 432 counts measured in Microsoft Word |
| `manifest.xml` | Add-in definition for local development |
````

- [ ] **Step 2: Run everything**

Run: `npm test && npm run validate && BASE_URL=https://example.github.io/ppt-word-count npm run build`
Expected: all tests pass; `The manifest is valid.`; `Built dist/`.

- [ ] **Step 3: Commit**

```bash
git add README.md
git commit -m "docs: add README with setup, sharing and troubleshooting"
```

---

## Self-review (done while writing)

- **Spec coverage:** §2 behaviour → Tasks 3 and 5. §3 counting → Task 1. §4.1 URL layout → Task 5 (serve) and Task 7 (build). §4.2 → Task 1. §4.3 → Task 3; the spec's `Table.values` + merged-areas idea was replaced by `getCellOrNullObject`, which is documented to return null objects for merged-over cells, and the spec's interface gains `selectedCount`. §4.4 → Tasks 4 and 5. §4.5 manifest → Task 6. §4.6 build and hosting → Task 7 and README. §5 limitations → help page (Task 6). §6 testing → tests in every task, plus preview and validate. The live PowerPoint check is done by the orchestrator after Task 8, because it needs the user's renewed certificate.
- **Placeholders:** none. Every code step has full code.
- **Type consistency:** `SelectionSnapshot` (`source`, `selectedCount`, `items[{id,name,kind,texts}]`, `unsupported[{id,name,type}]`) is used identically in Tasks 3 and 5. `Summary` adds `counts` and `total`. The refresher API is `schedule`/`refreshNow`/`poll`/`stop`. The fake host API is `setScenario`/`failNextSyncs`/`fireSelectionChanged`/`stats`.

---

## Addendum: changes after the multi-agent review (2026-09-26)

Five reviewers (Office API, counting, live updates, build/deploy, docs/UX) produced findings. An adversarial verifier then checked each one: 28 were confirmed (about 14 distinct issues) and 12 were rejected. Changes made on top of the tasks above:

- `src/selection.js`: `runJobs` passes on (rather than hiding) a `read()` failure after a successful sync (office-js #6363) and a round where every job fails on its own, so the refresher retries the whole read. Guesses (`getTable()` on "Unsupported" shapes) run in their own batch (`guess: true`) so a wrong guess can't force the round into one-at-a-time retries.
- `dev/fake-powerpoint.js`: `failSync(n)` and `trampleSync(n)` (a sync that resolves but loads nothing). `failNextSyncs` is now built on them. New scenario `smartArt`.
- `src/refresher.js`: a `timeoutMs` (10 s) on each read, so a request that never answers can't freeze the pane.
- `src/app.js`:
  - A selection of only charts or SmartArt says "Can't count this" once.
  - The note reads "N object(s) isn't/aren't included".
  - The error status stays in the page as a live region (visually hidden when quiet).
  - Focus moves to Refresh when "Try again" disappears.
  - One spoken summary per selection change replaces `aria-live` on the numbers.
- `src/help.html`: line breaks and emoji described correctly; added the web typing limitation, a "doesn't appear" and version-check troubleshooting entry, and the full Mac path.
- `scripts/serve.mjs`: the preview uses port 3101, with a clear message when a port is in use.
- `package.json`: `npm start` runs `npm run certs` first, so renewals last a year.
- `manifest.xml`: ribbon labels "Word Count (dev)".
- `scripts/build.mjs`: the production manifest gets a neutral header comment.
- `src/index.html`: redirects the site root to the help page.
- `README.md`: where to type commands, prerequisites, how to stop the preview, more troubleshooting.

## Addendum: findings from live testing in PowerPoint for Mac 16.113 (2026-09-26)

Tested with a generated deck (title, body with a Shift+Enter break, table with a merged row, nested group, picture, chart, Japanese text with an emoji). Selections were driven by AppleScript, and each reading was checked from the pane's dev log.

| Case | Result |
|---|---|
| Title (whole shape) | 2/16/17 "Title 1" ✔ (after the fix below) |
| Body placeholder with paragraphs and a Shift+Enter break | 15/71/84 ✔, confirming how PowerPoint returns line-break characters |
| Table with a merged row | 11/54/59 "Table 1", merged cell counted once ✔ |
| Nested group | 6/25/29 "Group 1" ✔ |
| Chart | "Can't count this" ✔ (after the fix below) |
| Chart + text box | "2 shapes selected", 9/36/44, 1 not included ✔ |
| Japanese + emoji | 17/35/41 ✔ |
| Highlighted word ("Quarterly") | 1/9/9 "Highlighted text in Title 1" ✔, live update via selection event |
| Partial highlight in a table cell | counts the whole table: the documented limitation |
| Picture | PowerPoint's add-in API reports no selection, so the pane shows "Nothing selected" |

Fixes:
- `selection.js`: PowerPoint for Mac reports a whole selected box as a text range holding all its text. The reader now reads the shape and treats the range as the whole shape when it covers all of the shape's text, ignoring whitespace. PowerPoint also returns a non-null range with `text: null` for a selected chart, so the reader guards against null text.
- Added `src/devlog.js` and request logging in `scripts/serve.mjs`: pane start-up steps, reads and errors go to the dev server log on localhost.

Observed: a pane in a window that macOS considers hidden (covered by other windows) has `visibilityState` "hidden", renders nothing to window captures and pauses polling. It refreshes when it becomes visible again, as designed.

## Addendum: highlighted text inside table cells (2026-09-26)

The user asked for highlighted text in table cells to be counted. Probed live on PowerPoint for Mac 16.113, with text highlighted in a table cell:
- `getSelectedTextRangeOrNullObject()` loading `text` gives a non-null range with `text: null`. Loading `start`/`length` throws InvalidArgument.
- `Office.context.document.getSelectedDataAsync(Office.CoercionType.Text)` returns exactly the highlighted text ("1,200"). For a selected cell it returns that cell's text plus "\r\n". For a whole selected table it returns every cell joined by "\r\n" (merged cells once).
- The preview API (@types/office-js-preview) adds nothing for table-cell text.

Implementation:
- `readSelection(context, { selectedText })`: when exactly one table is selected, the plain-text selection is compared with the table's text, ignoring whitespace. Anything less than the whole table is counted as a highlight.
- The first sync is retried without the text range if it fails.
- `app.js` exports `selectedTextReader(Office)`: a promise wrapper with a 1.5 s timeout, because office-js #4200 shows the call can fail to answer.
- The fake gained `getSelectedDataAsync`, `CoercionType`, `highlight: { error }`, `hangSelectedData` and a `tableHighlight` scenario.

Verified live: "1,200" → 1/5/5; "combined total" in the merged cell → 2/13/14; one cell → 2/10/11; whole table → 11/54/59; neighbouring text box unaffected.

## Addendum: a cursor inside a word (2026-09-26)

Probed live on PowerPoint for Mac 16.113. The user clicked once inside "platform", then double-clicked it:
- Cursor only: `getSelectedTextRangeOrNullObject()` returned "platform" (start 20, length 8), which is office-js #6839. `getSelectedDataAsync(Text)` returned "".
- Double-click: the range returned "platform", and `getSelectedDataAsync(Text)` returned "platform".
- Nothing selected: `getSelectedDataAsync(Text)` failed with code 1001, "The current selection is not supported."

Implementation: in the highlight branch (non-empty range, one shape), `readSelection` asks `selectedText()`. An empty string means only a cursor, so the whole shape is counted. `null` (unavailable or failed) keeps the range. The help page no longer lists the limitation. PowerPoint for Windows has the same range bug per #6839, and the fix relies on the same common API there, but it hasn't been tested on Windows. The fake's plain-text selection now defaults to the scenario's highlight. New scenario: `cursor`.

## Addendum: second review (2026-09-26)

A review of the changes above (3 lenses, each with an adversarial verifier) confirmed 8 findings, 5 distinct; 10 were rejected. Fixed:
- A selected empty cell ("\r\n") counts as 0 rather than the whole table. Only "" or null means no selection, and an entirely empty table stays "whole table".
- The whole-box check for text shapes ignores only paragraph breaks, so an unselected edge space keeps a highlight. The loose comparison remains for tables.
- The shapes-only retry of the first sync runs only for InvalidArgument. Transient errors reject again, so the refresher retries without losing a highlight, and the preview's "fail" button reaches the error state again.
- Spec sections 4.3 and 5 are marked as superseded by these addenda. The README explains how to watch the dev log alongside `npm start`.

Re-verified live: the whole multi-paragraph box reads "TextBox 1" (95/611/706); a group reads "Group 4".
