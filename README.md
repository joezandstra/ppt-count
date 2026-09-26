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
