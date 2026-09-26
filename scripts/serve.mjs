// Local HTTPS server for developing the add-in (https://localhost:3100).
//
// It reads the certificate that `npm run certs` installs in ~/.office-addin-dev-certs
// instead of calling office-addin-dev-certs itself, so starting the server never
// triggers a surprise keychain/password prompt.
//
// `--http` serves plain HTTP with no certificate. PowerPoint won't accept that, but
// it's enough for the browser preview at /dev/preview.html.

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

// Only these top-level paths are served; everything else is a 404.
const PUBLIC = ["src", "assets", "dev", "manifest.xml", "README.md"];

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
  ".md": "text/markdown; charset=utf-8",
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

async function handle(req, res) {
  const url = new URL(req.url, `https://localhost:${PORT}`);
  let path = decodeURIComponent(url.pathname);
  if (path === "/") path = "/src/taskpane.html";

  const target = normalize(join(ROOT, path));
  const top = target.slice(ROOT.length + 1).split(sep)[0];
  if (!target.startsWith(ROOT + sep) || !PUBLIC.includes(top)) {
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
