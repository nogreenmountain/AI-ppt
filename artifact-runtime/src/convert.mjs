#!/usr/bin/env node

import { dirname, resolve } from "node:path";

import {
  buildPptxFromSchema,
  loadSchemaFromPath,
  writePptx,
} from "./index.mjs";

function parseArgs(argv) {
  const args = { spec: null, out: null, preview: null, selfTest: false, help: false };
  for (let i = 2; i < argv.length; i += 1) {
    const arg = argv[i];
    if (arg === "--spec") args.spec = argv[++i];
    else if (arg === "--out") args.out = argv[++i];
    else if (arg === "--preview") args.preview = argv[++i];
    else if (arg === "--self-test") args.selfTest = true;
    else if (arg === "--help" || arg === "-h") args.help = true;
  }
  return args;
}

function printHelp() {
  process.stdout.write([
    "Usage:",
    "  node src/convert.mjs --spec <detected.json> --out <out.pptx> [--preview <out.png>]",
    "  node src/convert.mjs --self-test",
    "",
    "Notes:",
    "  The open-source builder writes PPTX files with pptxgenjs.",
    "  PNG previews are produced by the optional report step, not by this CLI.",
    "",
  ].join("\n"));
}

async function runBuild(args) {
  if (!args.spec) throw new Error("--spec <detected.json> is required");
  if (!args.out) throw new Error("--out <out.pptx> is required");

  const specPath = resolve(args.spec);
  const outPath = resolve(args.out);
  const schema = await loadSchemaFromPath(specPath);
  const result = await buildPptxFromSchema(schema, { baseDir: dirname(specPath) });
  const written = await writePptx(result.bytes, outPath);
  process.stdout.write(`[convert] wrote pptx: ${outPath} (${written} bytes)\n`);
  if (args.preview) {
    process.stdout.write("[convert] preview skipped: run the report step to render the PPTX via PowerPoint.\n");
  }
}

async function selfTest() {
  const { mkdtemp, writeFile, stat } = await import("node:fs/promises");
  const { tmpdir } = await import("node:os");
  const { join } = await import("node:path");

  const png = "iVBORw0KGgoAAAANSUhEUgAAAAEAAAABCAYAAAAfFcSJAAAADUlEQVR4nGNgYGD4DwABBAEAfbLI3wAAAABJRU5ErkJggg==";
  const dir = await mkdtemp(join(tmpdir(), "slide2pptx-js-"));
  await writeFile(join(dir, "bg.png"), Buffer.from(png, "base64"));
  const spec = {
    version: "1.0",
    source: { image_path: "bg.png", width_px: 1280, height_px: 720 },
    slide: { width: 1280, height: 720, unit: "px" },
    background: { strategy: "original", image_path: "bg.png" },
    elements: [
      {
        id: "el_title",
        kind: "text",
        bbox: { left: 80, top: 80, width: 900, height: 90 },
        z: 2,
        editable_score: 1,
        render_strategy: "native",
        text: "slide2pptx self-test",
        font_family: "Calibri",
        font_size: 36,
        font_weight: "bold",
        text_color: "#111827",
      },
      {
        id: "el_rect",
        kind: "shape",
        bbox: { left: 80, top: 220, width: 360, height: 100 },
        z: 1,
        editable_score: 1,
        render_strategy: "native",
        geometry: "rect",
        fill: "#2563EB",
        line_color: "#1E40AF",
        line_width: 2,
      },
    ],
  };
  const specPath = join(dir, "detected.json");
  const pptxPath = join(dir, "selftest.pptx");
  await writeFile(specPath, JSON.stringify(spec, null, 2), "utf8");
  await runBuild({ spec: specPath, out: pptxPath });
  const size = (await stat(pptxPath)).size;
  if (size <= 0) throw new Error("self-test produced an empty PPTX");
  process.stdout.write(`[self-test] OK ${pptxPath} (${size} bytes)\n`);
}

async function main() {
  const args = parseArgs(process.argv);
  if (args.help) return printHelp();
  if (args.selfTest) return selfTest();
  return runBuild(args);
}

main().catch((err) => {
  process.stderr.write(`[convert] ${err?.stack || err}\n`);
  process.exit(1);
});
