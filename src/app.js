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
