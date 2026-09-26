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
