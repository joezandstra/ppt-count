// Entry point: waits for Office.js, then starts the pane.
import { start } from "./app.js";

const root = document.getElementById("app");
const devLog = globalThis.devLog ?? (() => {}); // see devlog.js

if (typeof Office === "undefined") {
  devLog("Office.js is missing");
  root.textContent = "Word Count couldn't load Office.js. Check your internet connection, then close and reopen the pane.";
} else {
  devLog("waiting for Office.onReady");
  Office.onReady((info) => {
    devLog(`Office.onReady: host=${info.host} platform=${info.platform}`);
    start({ Office, PowerPoint: globalThis.PowerPoint, root, info, log: devLog });
    devLog("pane started");
  });
}
