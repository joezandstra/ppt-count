// Local development only: sends start-up steps and errors from the pane to the dev
// server's log (scripts/serve.mjs), because PowerPoint's pane has no visible console.
// A classic script, loaded before the modules, so it also catches module load errors.
// Does nothing once the add-in is hosted anywhere other than localhost.
(function () {
  var enabled = location.hostname === "localhost" && typeof navigator.sendBeacon === "function";
  window.devLog = function (message) {
    if (enabled) navigator.sendBeacon("/__log", String(message));
  };
  window.addEventListener("error", function (event) {
    window.devLog("error: " + event.message + " (" + event.filename + ":" + event.lineno + ")");
  });
  window.addEventListener("unhandledrejection", function (event) {
    var reason = event.reason;
    window.devLog("unhandled rejection: " + (reason && reason.message ? reason.message : reason));
  });
  window.devLog("pane loading");
})();
