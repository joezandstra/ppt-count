// Entry point: waits for Office.js, then starts the pane.
import { start } from "./app.js";

const root = document.getElementById("app");

if (typeof Office === "undefined") {
  root.textContent = "Word Count couldn't load Office.js. Check your internet connection, then close and reopen the pane.";
} else {
  Office.onReady((info) => start({ Office, PowerPoint: globalThis.PowerPoint, root, info }));
}
