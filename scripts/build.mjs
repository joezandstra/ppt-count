// Builds dist/: the files to host (for example on GitHub Pages) plus the
// production manifest that colleagues install.
//
//   BASE_URL=https://you.github.io/ppt-word-count npm run build
//
// The production manifest is manifest.xml with the hosting address, its own Id
// (so it can sit next to the dev copy) and the name "Word Count". Every local
// script and stylesheet reference gets ?v=<content hash>, because GitHub Pages
// lets browsers cache files for 10 minutes and a deploy could otherwise mix old
// and new files.

import { createHash } from "node:crypto";
import { mkdir, readdir, readFile, rm, writeFile } from "node:fs/promises";
import { dirname, extname, join, relative, resolve } from "node:path";
import { fileURLToPath, pathToFileURL } from "node:url";

const ROOT = resolve(fileURLToPath(new URL("..", import.meta.url)));
export const DEV_ORIGIN = "https://localhost:3100";
export const DEV_ID = "25c6e699-85a5-4a2d-bc50-ecc66e7483be";
export const PROD_ID = "1b331966-f37d-4165-9499-7f94f173c9c1";
export const MANIFEST_NAME = "word-count-manifest.xml";

/** @returns {string} the address without a trailing slash */
export function normalizeBaseUrl(baseUrl) {
  let url;
  try {
    url = new URL(String(baseUrl ?? ""));
  } catch {
    throw new Error(`BASE_URL must be a full web address, like https://you.github.io/ppt-word-count (got "${baseUrl ?? ""}").`);
  }
  if (url.protocol !== "https:") throw new Error("BASE_URL must start with https:// because Office only loads add-ins over HTTPS.");
  if (url.hostname === "localhost" || url.hostname === "127.0.0.1") throw new Error("BASE_URL can't be localhost; use the hosting address.");
  if (url.search || url.hash) throw new Error("BASE_URL can't contain ? or #.");
  return url.href.replace(/\/+$/, "");
}

/** Turns the dev manifest into the production one. */
export function productionManifest(devManifest, baseUrl) {
  for (const required of [DEV_ORIGIN, DEV_ID, 'DefaultValue="Word Count (dev)"']) {
    if (!devManifest.includes(required)) throw new Error(`manifest.xml no longer contains ${required}; update scripts/build.mjs to match.`);
  }
  const manifest = devManifest
    .replace(/<!--\s*Development manifest[\s\S]*?-->/, "<!-- Word Count for PowerPoint. Generated from manifest.xml by scripts/build.mjs. -->")
    .replaceAll(`<AppDomain>${DEV_ORIGIN}</AppDomain>`, `<AppDomain>${new URL(baseUrl).origin}</AppDomain>`)
    .replaceAll(DEV_ORIGIN, baseUrl)
    .replaceAll(DEV_ID, PROD_ID)
    .replaceAll('DefaultValue="Word Count (dev)"', 'DefaultValue="Word Count"');
  if (/localhost/i.test(manifest)) throw new Error("The production manifest still mentions localhost.");
  return manifest;
}

function addVersion(source, extension, version) {
  if (extension === ".html") {
    // Local .js/.css references only: absolute URLs (the Office.js CDN) are left alone.
    return source.replace(/\b(src|href)="(?![a-z][a-z0-9+.-]*:|\/\/|#)([^"?#]+\.(?:js|css))"/gi, `$1="$2?v=${version}"`);
  }
  if (extension === ".js") {
    return source.replace(/(\bfrom\s*|\bimport\s*\(\s*|\bimport\s+)(["'])(\.{1,2}\/[^"'?#]+\.js)\2/g, `$1$2$3?v=${version}$2`);
  }
  return source;
}

async function listFiles(dir) {
  const entries = await readdir(dir, { withFileTypes: true });
  const nested = await Promise.all(entries.map((entry) => (entry.isDirectory() ? listFiles(join(dir, entry.name)) : [join(dir, entry.name)])));
  return nested.flat().sort();
}

/**
 * @param {{ baseUrl: string, root?: string, outDir?: string }} options
 * @returns {Promise<{ outDir: string, version: string, manifestPath: string }>}
 */
export async function build({ baseUrl, root = ROOT, outDir = join(root, "dist") }) {
  const base = normalizeBaseUrl(baseUrl);
  const sources = [
    ...(await listFiles(join(root, "src"))).map((file) => ({ file, target: relative(join(root, "src"), file) })),
    ...(await listFiles(join(root, "assets"))).map((file) => ({ file, target: join("assets", relative(join(root, "assets"), file)) })),
  ];
  const contents = await Promise.all(sources.map(({ file }) => readFile(file)));

  const hash = createHash("sha256");
  sources.forEach(({ target }, i) => hash.update(target).update("\0").update(contents[i]).update("\0"));
  const version = hash.digest("hex").slice(0, 10);

  await rm(outDir, { recursive: true, force: true });
  for (const [i, { target }] of sources.entries()) {
    const destination = join(outDir, target);
    await mkdir(dirname(destination), { recursive: true });
    const extension = extname(target);
    const body = extension === ".html" || extension === ".js" ? addVersion(contents[i].toString("utf8"), extension, version) : contents[i];
    await writeFile(destination, body);
  }

  const manifestPath = join(outDir, MANIFEST_NAME);
  await writeFile(manifestPath, productionManifest(await readFile(join(root, "manifest.xml"), "utf8"), base));
  return { outDir, version, manifestPath };
}

if (process.argv[1] && import.meta.url === pathToFileURL(resolve(process.argv[1])).href) {
  try {
    const { outDir, version, manifestPath } = await build({ baseUrl: process.env.BASE_URL });
    console.log(`Built ${relative(ROOT, outDir)}/ (version ${version}).`);
    console.log(`Production manifest: ${relative(ROOT, manifestPath)}`);
  } catch (error) {
    console.error(error.message);
    process.exitCode = 1;
  }
}
