# Word Count for PowerPoint

A PowerPoint add-in that shows **Words**, **Characters (no spaces)** and **Characters (with spaces)** for whatever you select: highlighted text, text boxes, placeholders, tables and groups. It updates as you click around, and it counts the same way Microsoft Word does.

Works in current PowerPoint for Mac (16.105+), Windows (Version 2601+) and PowerPoint on the web. How the counts work, the known limitations and install steps for colleagues are all on the add-in's help page ([`src/help.html`](src/help.html)).

## Try it on this Mac

**Where to type the commands:** open the **Terminal** app and go to the project folder first:

```bash
cd ~/Developer/PPT-word-char-count
```

Type every command below in that same window. (On a fresh copy of the project, or on another Mac, first install [Node.js](https://nodejs.org) version 20 or later and run `npm install` once in the project folder.)

You only need to do step 1 once a year.

1. **Trust the local certificate.** PowerPoint only loads add-ins over a secure connection, so your Mac needs to trust a small local certificate. Run the command below. Your Mac asks for your password, possibly once in the terminal (nothing appears as you type; that's normal) and once in a pop-up window.

   ```bash
   npm run certs
   ```

2. **Start the add-in.** This starts a small local web server in the background and opens PowerPoint with the add-in loaded. (If the certificate from step 1 has run out, it asks for your password again and renews it for another year.)

   ```bash
   npm start
   ```

   In PowerPoint, choose **Home › Word Count (dev)** and select some text. ("(dev)" marks this local copy; the version you share with colleagues is just called Word Count.)

3. **When you're done**, stop the server. It otherwise keeps running until you restart your Mac.

   ```bash
   npm stop
   ```

### Preview in a browser (no PowerPoint needed)

```bash
npm run preview
```

Then open <http://localhost:3101/dev/preview.html>. This runs the real pane against a pretend PowerPoint, with a menu of example selections. Press **Ctrl+C** in the Terminal window to stop it.

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

- **`npm start` asks for your Mac password:** that's the yearly certificate renewal. Type your password (nothing appears as you type).
- **The pane is blank or shows an error after restarting your Mac:** the local server stopped. Run `npm start` again, or `npm stop` to remove the local add-in from PowerPoint.
- **The pane shows an old version:** open the small menu at the top of the pane and choose **Clear Web Cache**, then reopen it.
- **Word Count (dev) doesn't appear on the Home tab:** quit PowerPoint completely and run `npm start` again.
- **"Port 3100 is already in use":** a copy of the server is already running. Run `npm stop`, then `npm start`.

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
