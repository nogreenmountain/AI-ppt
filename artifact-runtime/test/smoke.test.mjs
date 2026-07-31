// End-to-end smoke test. Synthesises the fixtures (a 1x1 transparent PNG
// inlined so the test does not depend on the filesystem), invokes the
// builder via both the high-level API and the CLI entry point, and
// writes both the PPTX and a PNG preview into test/_out/smoke.pptx +
// smoke.png.

import { mkdir, writeFile, stat, readFile } from "node:fs/promises";
import { dirname, resolve } from "node:path";
import { fileURLToPath } from "node:url";

import {
  buildPptxFromSchema,
  writePptx,
  writePngPreview,
  runBuild,
  loadSchemaFromPath,
  resolveSlideSize,
  sortByZ,
  buildElement,
} from "../src/index.mjs";
import { loadImageBytes } from "../src/image-loader.mjs";

// 1x1 transparent PNG (smallest legal PNG). Embedding it here keeps the
// test self-contained — no fixture files required on disk for the image
// bytes themselves.
const PNG_1x1 = new Uint8Array([
  0x89, 0x50, 0x4e, 0x47, 0x0d, 0x0a, 0x1a, 0x0a,
  0x00, 0x00, 0x00, 0x0d, 0x49, 0x48, 0x44, 0x52,
  0x00, 0x00, 0x00, 0x01, 0x00, 0x00, 0x00, 0x01,
  0x08, 0x06, 0x00, 0x00, 0x00, 0x1f, 0x15, 0xc4,
  0x89, 0x00, 0x00, 0x00, 0x0d, 0x49, 0x44, 0x41,
  0x54, 0x08, 0x99, 0x63, 0x00, 0x01, 0x00, 0x00,
  0x05, 0x00, 0x01, 0x0d, 0x0a, 0x2d, 0xb4, 0x00,
  0x00, 0x00, 0x00, 0x49, 0x45, 0x4e, 0x44, 0xae,
  0x42, 0x60, 0x82,
]);

const here = dirname(fileURLToPath(import.meta.url));
const fixturesDir = resolve(here, "fixtures");
const outDir = resolve(here, "_out");
const schemaPath = resolve(fixturesDir, "smoke.schema.json");
const bgPath = resolve(fixturesDir, "source.png");
const inlinePath = resolve(fixturesDir, "inline.png");
const pptxPath = resolve(outDir, "smoke.pptx");
const previewPath = resolve(outDir, "smoke.png");

async function ensureFixturePngs() {
  await mkdir(fixturesDir, { recursive: true });
  await writeFile(bgPath, PNG_1x1);
  await writeFile(inlinePath, PNG_1x1);
}

let pass = 0;
let fail = 0;
function check(cond, msg) {
  if (cond) { pass++; process.stdout.write(`  ok  ${msg}\n`); }
  else { fail++; process.stdout.write(`  FAIL ${msg}\n`); }
}

async function main() {
  await mkdir(outDir, { recursive: true });
  await ensureFixturePngs();

  const schema = await loadSchemaFromPath(schemaPath);
  const { widthPx, heightPx } = resolveSlideSize(schema.slide);
  check(widthPx === 1280 && heightPx === 720, "slide size from schema");

  const kinds = new Set(schema.elements.map((e) => e.kind));
  check(kinds.has("text") && kinds.has("shape") && kinds.has("image"), "fixture covers all kinds");

  const geos = new Set(schema.elements.filter((e) => e.kind === "shape").map((e) => e.geometry));
  for (const g of ["rect", "roundRect", "ellipse", "line"]) {
    check(geos.has(g), `fixture covers shape geometry ${g}`);
  }

  const nodes = schema.elements.map((el) => buildElement(el, null)).filter(Boolean);
  const sorted = sortByZ(nodes);
  check(sorted[0].z <= sorted[sorted.length - 1].z, "sortByZ is non-decreasing");

  // Build via the high-level entry. The loader supplies bytes for the
  // background + image element.
  const loadedPaths = new Set();
  const loader = async (_id, p) => {
    loadedPaths.add(p);
    return loadImageBytes(p, fixturesDir);
  };

  let result;
  try {
    result = await buildPptxFromSchema(schema, { baseDir: fixturesDir, imageLoader: loader });
  } catch (err) {
    process.stderr.write(`build failed: ${err.stack || err.message}\n`);
    fail++;
    return;
  }
  check(Boolean(result?.bytes), "buildPptxFromSchema returned bytes");
  check(typeof result.mime === "string" && result.mime.length > 0, "blob has a mime type");
  check(result.bytes instanceof Uint8Array, "blob is a Uint8Array");
  check(result.bytes.length > 0, `PPTX blob is non-empty (${result.bytes.length} bytes)`);

  const written = await writePptx(result.bytes, pptxPath);
  check(written > 0, `wrote ${written} PPTX bytes`);
  const onDisk = await stat(pptxPath);
  check(onDisk.size > 0, "PPTX exists on disk and is non-empty");

  // PNG preview (best-effort — skia-canvas is bundled with @oai/artifact-tool
  // but the binary may not load on every machine).
  try {
    const previewBytes = await writePngPreview(schema, previewPath, { baseDir: fixturesDir });
    check(previewBytes > 0, `wrote ${previewBytes} preview bytes`);
    const previewOnDisk = await stat(previewPath);
    check(previewOnDisk.size > 0, "PNG preview exists on disk");
  } catch (err) {
    process.stdout.write(`  note: PNG preview skipped (${err.message})\n`);
  }

  // Exercise runBuild (the CLI entry point).
  try {
    const cliResult = await runBuild({
      input: schemaPath,
      out: pptxPath,
      preview: previewPath,
      skipPreview: false,
      quiet: true,
      cwd: here,
    });
    check(Boolean(cliResult?.pptxPath), "runBuild returned pptxPath");
  } catch (err) {
    process.stderr.write(`runBuild failed: ${err.stack || err.message}\n`);
    fail++;
  }

  check(loadedPaths.size >= 2, `image loader invoked for both images (${loadedPaths.size} paths)`);

  // Confirm the file we wrote is recognised as a PPTX (ZIP magic bytes).
  const head = (await readFile(pptxPath)).subarray(0, 4);
  check(
    head[0] === 0x50 && head[1] === 0x4b && head[2] === 0x03 && head[3] === 0x04,
    "PPTX starts with ZIP magic bytes (PK\\x03\\x04)",
  );

  check(result.mime.includes("presentationml.presentation"), "PPTX MIME type is reported");
}

main().then(() => {
  if (fail > 0) {
    process.stderr.write(`\n${fail} failure(s)\n`);
    process.exit(1);
  }
  process.stdout.write(`\nsmoke: all ${pass} assertions passed\n`);
}, (err) => {
  process.stderr.write(`smoke fatal: ${err.stack || err.message}\n`);
  process.exit(99);
});
