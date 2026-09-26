// Local server for developing the add-in (https://localhost:3100).
//
// It reads the certificate that `npm run certs` installs in ~/.office-addin-dev-certs
// instead of calling office-addin-dev-certs itself, so starting the server never
// triggers a surprise keychain/password prompt.
//
// `--http` serves plain HTTP with no certificate on port 3101. PowerPoint won't
// accept that, but it's enough for the browser preview at /dev/preview.html. It uses
// a different port so a forgotten preview can't get in the way of `npm start`.
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
const HTTP_ONLY = process.argv.includes("--http");
const PORT = Number(process.env.PORT) || (HTTP_ONLY ? 3101 : 3100);
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
  if (req.method === "POST" && req.url === "/__log") return logFromPane(req, res);
  logRequest(req, res);
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

// One line per request, so you can see what PowerPoint asked for (handy when a pane stays blank).
function logRequest(req, res) {
  res.on("finish", () => console.log(`${new Date().toISOString().slice(11, 19)} ${res.statusCode} ${req.method} ${req.url}`));
}

// Messages from src/devlog.js in the pane.
function logFromPane(req, res) {
  let body = "";
  req.setEncoding("utf8");
  req.on("data", (chunk) => {
    if (body.length < 2000) body += chunk;
  });
  req.on("end", () => {
    console.log(`${new Date().toISOString().slice(11, 19)} [pane] ${body.slice(0, 2000)}`);
    res.writeHead(204);
    res.end();
  });
}

function send(res, status, text) {
  res.writeHead(status, { "Content-Type": "text/plain; charset=utf-8" });
  res.end(text);
}

const server = HTTP_ONLY ? createHttpServer(handle) : createServer(loadCertificate(), handle);
server.on("error", (error) => {
  if (error.code !== "EADDRINUSE") throw error;
  console.error(`\nPort ${PORT} is already in use, probably by another copy of this server.`);
  console.error(HTTP_ONLY ? "Close the other preview window, or press Ctrl+C there.\n" : "Run `npm stop`, then try again.\n");
  process.exit(1);
});
server.listen(PORT, () => {
  if (HTTP_ONLY) {
    console.log(`Browser preview: http://localhost:${PORT}/dev/preview.html`);
    console.log("Press Ctrl+C to stop it.");
  } else {
    console.log(`Word Count add-in is being served at https://localhost:${PORT}`);
    console.log("Leave this window open while you use the add-in. Press Ctrl+C to stop.");
  }
});
