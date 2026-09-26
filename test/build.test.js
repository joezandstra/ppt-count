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
    "index.html",
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
  assert.doesNotMatch(manifest, /\(dev\)|Development manifest/);
});

test("the site's front page leads to the help and install page", () => {
  assert.match(read("index.html"), /url=help\.html/);
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
