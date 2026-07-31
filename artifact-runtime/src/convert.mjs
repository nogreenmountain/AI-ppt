#!/usr/bin/env node
// convert.mjs
//
// Documented facade for @oai/artifact-tool. Builds a PPTX (and optionally
// a PNG preview) from a detected JSON spec. This file is intentionally
// self-contained and uses only the public, documented API:
//
//   import { Presentation, PresentationFile } from "@oai/artifact-tool";
//
//   Presentation.create({ slideSize })
//   presentation.slides.add()                         -> slide
//   slide.images.add({ ... })
//   slide.shapes.add({ ... })                         -> shape
//   shape.text = "..."
//   shape.text.style = { ... }
//   PresentationFile.exportPptx(presentation)         -> bytes
//   presentation.export({ slide, format: "png", scale: 1 }) -> bytes
//
// CLI:
//   node src/convert.mjs --spec <path.json> [--out <pptx>] [--preview <png>]
//   node src/convert.mjs --self-test
//
// The --self-test mode embeds a valid tiny PNG as base64 (no Python, no
// network, no scratch files) and walks every API path so a missing import,
// a renamed export, or a missing method surfaces immediately.

import fs from "node:fs/promises";
import path from "node:path";
import { fileURLToPath } from "node:url";
import { Presentation, PresentationFile } from "@oai/artifact-tool";

// ---------------------------------------------------------------------------
// CLI parsing
// ---------------------------------------------------------------------------

function parseArgs(argv) {
  const args = { spec: null, out: null, preview: null, selfTest: false };
  for (let i = 2; i < argv.length; i += 1) {
    const a = argv[i];
    if (a === "--spec") {
      args.spec = argv[++i];
    } else if (a === "--out") {
      args.out = argv[++i];
    } else if (a === "--preview") {
      args.preview = argv[++i];
    } else if (a === "--self-test") {
      args.selfTest = true;
    } else if (a === "--help" || a === "-h") {
      args.help = true;
    }
  }
  return args;
}

function printHelp() {
  process.stdout.write(
    [
      "Usage:",
      "  node src/convert.mjs --spec <spec.json> [--out <out.pptx>] [--preview <out.png>]",
      "  node src/convert.mjs --self-test",
      "",
    ].join("\n")
  );
}

// ---------------------------------------------------------------------------
// Bytes helpers
// ---------------------------------------------------------------------------

// Normalise whatever the artifact-tool runtime hands back into a Uint8Array
// we can write to disk. The facade is documented to return byte-backed
// results, but we still defend against ArrayBuffer / Blob / { data } shapes.
async function toBytes(result) {
  if (result == null) {
    throw new Error("toBytes: result is null/undefined");
  }
  if (result instanceof Uint8Array) return result;
  if (result instanceof ArrayBuffer) return new Uint8Array(result);
  if (ArrayBuffer.isView(result)) {
    return new Uint8Array(result.buffer, result.byteOffset, result.byteLength);
  }
  if (typeof result === "object" && result.data) {
    if (result.data instanceof Uint8Array) return result.data;
    if (result.data instanceof ArrayBuffer) return new Uint8Array(result.data);
  }
  if (typeof result.arrayBuffer === "function") {
    return new Uint8Array(await result.arrayBuffer());
  }
  throw new Error(
    "toBytes: unsupported result type " + (typeof result)
  );
}

// ---------------------------------------------------------------------------
// Asset / image loading
// ---------------------------------------------------------------------------

function resolveAsset(assetPath, baseDir) {
  if (!assetPath) return null;
  if (path.isAbsolute(assetPath)) return assetPath;
  return path.resolve(baseDir, assetPath);
}

async function readImageBytes(assetPath, baseDir) {
  const abs = resolveAsset(assetPath, baseDir);
  const buf = await fs.readFile(abs);
  return new Uint8Array(buf);
}

function imageContentType(assetPath) {
  const ext = path.extname(assetPath || "").toLowerCase();
  if (ext === ".jpg" || ext === ".jpeg") return "image/jpeg";
  if (ext === ".webp") return "image/webp";
  return "image/png";
}

function toArrayBuffer(bytes) {
  return bytes.buffer.slice(bytes.byteOffset, bytes.byteOffset + bytes.byteLength);
}

// ---------------------------------------------------------------------------
// Style / element translation
// ---------------------------------------------------------------------------

function buildTextStyle(el, styleOverride = {}) {
  const style = {};
  // OCR boxes describe the visible ink bounds, while PowerPoint text boxes
  // normally add their own padding and use different Chinese font metrics.
  // Microsoft YaHei is available on the target Windows machine and gives a
  // much closer match than Arial's Chinese fallback.
  style.fontFamily = /[\u3400-\u9fff]/u.test(String(el.text || ""))
    ? "Microsoft YaHei"
    : (el.font_family || "Arial");
  if (el.font_size != null) style.fontSize = el.font_size * 0.91;
  if (el.font_weight) style.bold = el.font_weight === "bold";
  if (el.font_style) style.italic = el.font_style === "italic";
  if (el.text_color) style.color = el.text_color;
  if (el.align) style.alignment = el.align;
  return { ...style, ...styleOverride };
}

// High-confidence visual review corrections for OCR slips visible in the
// supplied reference image. Keeping them here avoids mutating detected.json,
// which remains the raw detector output for reproducibility.
const TEXT_CORRECTIONS = new Map([
  ["东方仿旗", "东方仿真"],
  ["统一技术底座：SimforAl虚拟化工现场平台", "统一技术底座：Sim for AI虚拟化工现场平台"],
  ["化工、炼化、能等生产企业", "化工、炼化、能源等生产企业"],
  ["面向产业真实同题的化工+AI大案", "面向产业真实问题的化工+AI大赛"],
  ["成果验证与选拨", "成果验证与选拔"],
  ["真实问题转化为察题", "真实问题转化为赛题"],
  ["高校生与科研困队参", "高校师生与科研团队参赛"],
  ["优要方案验证与选", "优秀方案验证与筛选"],
  ["·数字生联合实验室", "·数字孪生联合实验室"],
  ["·定制化数据与机型服务", "·定制化数据与模型服务"],
  ["共建资源|协同攻关|成转化", "共建资源丨协同攻关丨成果转化"],
  ["数共擎", "数据共享"],
  ["润向产业雄题", "面向产业难题"],
  ["望共瞳", "模型共建"],
  ["开联合研究", "开展联合研究"],
  ["高校科研困队、", "高校科研团队、"],
  ["研发皖所、研究机构", "研发院所、研究机构"],
  ["工业数据乏", "工业数据匮乏"],
  ["盘产业提供真实问题", "产业提供真实问题"],
]);

function correctedText(el, localCorrections = {}) {
  const raw = el.text != null ? String(el.text) : "";
  return localCorrections[raw] || TEXT_CORRECTIONS.get(raw) || raw;
}

function textCenterInsideRegions(el, regions = []) {
  if (el.kind !== "text") return false;
  const b = bboxOf(el);
  const cx = b.left + b.width / 2;
  const cy = b.top + b.height / 2;
  return regions.some((region) => {
    const r = region.bbox || region;
    return (
      cx >= (r.left || 0) &&
      cy >= (r.top || 0) &&
      cx <= (r.left || 0) + (r.width || 0) &&
      cy <= (r.top || 0) + (r.height || 0)
    );
  });
}

function textBoxPosition(el) {
  const b = bboxOf(el);
  // Give glyphs a small amount of horizontal breathing room. The expansion
  // is centred so alignment to the reference image remains stable.
  const extraW = Math.max(4, Math.min(18, b.width * 0.05));
  const extraH = Math.max(2, Math.min(6, b.height * 0.16));
  return {
    left: Math.max(0, b.left - extraW / 2),
    top: Math.max(0, b.top - extraH / 2),
    width: b.width + extraW,
    height: b.height + extraH,
  };
}

function bboxOf(el) {
  const b = el.bbox || {};
  return {
    left: b.left || 0,
    top: b.top || 0,
    width: b.width || 0,
    height: b.height || 0,
  };
}

// Add one element from the spec to the slide. Background image is added
// separately so it always sits behind every other element.
async function addElement(
  el,
  slide,
  baseDir,
  localCorrections = {},
  textStyleOverrides = {}
) {
  const position = el.kind === "text" ? textBoxPosition(el) : bboxOf(el);
  const { left, top, width, height } = position;

  if (el.kind === "text") {
    // A lone low-confidence "B" was OCR noise from the circular factory
    // icon. The icon itself is already preserved in the cleaned background.
    if (el.text === "B" && (el.confidence?.ocr || 0) < 0.8) return;

    const shape = slide.shapes.add({
      geometry: "textbox",
      position,
      fill: "none",
      line: { style: "solid", fill: "none", width: 0 },
    });
    shape.text = correctedText(el, localCorrections);
    shape.text.style = buildTextStyle(
      el,
      textStyleOverrides[el.id] || textStyleOverrides[String(el.text || "")] || {}
    );
    shape.text.verticalAlignment = el.valign || "middle";
    shape.text.autoFit = "shrinkText";
    shape.text.wrap = "none";
    shape.text.insets = { top: 0, right: 0, bottom: 0, left: 0 };
    return;
  }

  if (el.kind === "shape") {
    if (el.geometry === "line") {
      // The documented facade does not guarantee a line primitive, so
      // represent lines with a thin filled rect that matches the spec's
      // bbox height (falling back to max(line_width, 1) if too small).
      const thin = Math.max(height || 0, el.line_width || 1, 1);
      slide.shapes.add({
        geometry: "rect",
        position: { left, top, width, height: thin },
        fill: el.line_color || "#000000",
        line: { style: "solid", fill: "none", width: 0 },
      });
      return;
    }
    const geom = el.geometry || "rect";
    if (geom !== "rect" && geom !== "roundRect" && geom !== "ellipse") {
      // Unknown geometry -> fall back to rect rather than throwing.
      slide.shapes.add({
        geometry: "rect",
        position,
        fill: el.fill || "none",
        line: el.line_color
          ? { style: "solid", fill: el.line_color, width: el.line_width || 1 }
          : { style: "solid", fill: "none", width: 0 },
      });
      return;
    }
    slide.shapes.add({
      geometry: geom,
      position,
      fill: el.fill || "none",
      line: el.line_color
        ? { style: "solid", fill: el.line_color, width: el.line_width || 1 }
        : { style: "solid", fill: "none", width: 0 },
    });
    return;
  }

  if (el.kind === "image") {
    const bytes = await readImageBytes(el.image_path, baseDir);
    slide.images.add({
      blob: toArrayBuffer(bytes),
      contentType: imageContentType(el.image_path),
      alt: el.id || "Detected image element",
      fit: el.fit || "cover",
      position,
    });
    return;
  }

  // Unknown kind -> skip silently.
}

// ---------------------------------------------------------------------------
// Build flow
// ---------------------------------------------------------------------------

async function buildFromSpec(specPath) {
  const absSpec = path.resolve(specPath);
  const raw = await fs.readFile(absSpec, "utf-8");
  let schema = JSON.parse(raw);
  const baseDir = path.dirname(absSpec);

  // A compact review spec may extend raw detector output with screenshot
  // overlays, OCR corrections, and text-suppression regions. This preserves
  // detected.json as an auditable raw result while keeping review adjustments
  // reusable and easy to inspect.
  if (schema.base_spec) {
    const baseSpecPath = resolveAsset(schema.base_spec, baseDir);
    const baseRaw = await fs.readFile(baseSpecPath, "utf-8");
    const baseSchema = JSON.parse(baseRaw);
    schema = {
      ...baseSchema,
      ...schema,
      slide: schema.slide || baseSchema.slide,
      background: schema.background || baseSchema.background,
      elements: [
        ...(baseSchema.elements || []),
        ...(schema.elements || []),
      ],
      text_corrections: {
        ...(baseSchema.text_corrections || {}),
        ...(schema.text_corrections || {}),
      },
      text_style_overrides: {
        ...(baseSchema.text_style_overrides || {}),
        ...(schema.text_style_overrides || {}),
      },
      suppress_text_regions: [
        ...(baseSchema.suppress_text_regions || []),
        ...(schema.suppress_text_regions || []),
      ],
    };
  }

  const slideW = (schema.slide && schema.slide.width) || 1280;
  const slideH = (schema.slide && schema.slide.height) || 720;

  const presentation = Presentation.create({
    slideSize: { width: slideW, height: slideH },
  });

  const slide = presentation.slides.add();

  // Background image first -> it lives at the bottom of the z stack.
  if (
    schema.background &&
    schema.background.image_path &&
    schema.background.strategy !== "none"
  ) {
    const bytes = await readImageBytes(
      schema.background.image_path,
      baseDir
    );
    slide.images.add({
      blob: toArrayBuffer(bytes),
      contentType: imageContentType(schema.background.image_path),
      alt: "Reconstructed slide background",
      fit: "cover",
      position: { left: 0, top: 0, width: slideW, height: slideH },
    });
  }

  // Native elements in ascending z order.
  const elements = (schema.elements || []).slice();
  elements.sort((a, b) => (a.z || 0) - (b.z || 0));
  for (const el of elements) {
    if (textCenterInsideRegions(el, schema.suppress_text_regions)) continue;
    await addElement(
      el,
      slide,
      baseDir,
      schema.text_corrections,
      schema.text_style_overrides
    );
  }

  return { presentation, slide, slideW, slideH };
}

// ---------------------------------------------------------------------------
// Self-test
//
// Embeds a valid 1x1 red PNG (RFC-2083 base64) and exercises every
// documented facade entry point. No Python, no network, no scratch files.
// ---------------------------------------------------------------------------

// 1x1 transparent-ish PNG (8-bit RGBA), generated offline and pasted here.
const EMBEDDED_PNG_B64 =
  "iVBORw0KGgoAAAANSUhEUgAAAAEAAAABCAYAAAAfFcSJAAAADUlEQVR4nGNgYGD4DwABBAEAfbLI3wAAAABJRU5ErkJggg==";

function embeddedPngBytes() {
  return new Uint8Array(Buffer.from(EMBEDDED_PNG_B64, "base64"));
}

async function selfTest() {
  process.stdout.write("[self-test] starting\n");

  // 1. Validate the runtime actually exposes the documented API.
  if (typeof Presentation !== "object" && typeof Presentation !== "function") {
    throw new Error("Presentation export is missing from @oai/artifact-tool");
  }
  if (typeof Presentation.create !== "function") {
    throw new Error("Presentation.create is not a function");
  }
  if (typeof PresentationFile !== "object" && typeof PresentationFile !== "function") {
    throw new Error(
      "PresentationFile export is missing from @oai/artifact-tool"
    );
  }
  if (typeof PresentationFile.exportPptx !== "function") {
    throw new Error("PresentationFile.exportPptx is not a function");
  }

  // 2. Build a presentation with a full-slide background + native textbox.
  const presentation = Presentation.create({
    slideSize: { width: 1280, height: 720 },
  });

  const slide = presentation.slides.add();

  if (!slide || typeof slide.images.add !== "function") {
    throw new Error("slide.images.add is missing");
  }
  if (!slide || typeof slide.shapes.add !== "function") {
    throw new Error("slide.shapes.add is missing");
  }

  // Full-slide background image (byte-backed).
  slide.images.add({
    blob: toArrayBuffer(embeddedPngBytes()),
    contentType: "image/png",
    alt: "Self-test background",
    fit: "cover",
    position: { left: 0, top: 0, width: 1280, height: 720 },
  });

  // Native text element using the documented textbox geometry.
  const titleShape = slide.shapes.add({
    geometry: "textbox",
    position: { left: 60, top: 60, width: 1160, height: 80 },
    fill: "none",
    line: { style: "solid", fill: "none", width: 0 },
  });
  if (!titleShape || !("text" in titleShape)) {
    throw new Error("shape does not expose a writable text property");
  }
  titleShape.text = "Self-test OK";
  titleShape.text.style = {
    fontFamily: "Calibri",
    fontSize: 44,
    bold: true,
    color: "#1F2937",
    alignment: "left",
  };

  // Native shapes: rect, roundRect, ellipse (geometry coverage).
  slide.shapes.add({
    geometry: "rect",
    position: { left: 800, top: 170, width: 400, height: 110 },
    fill: "#2563EB",
    line: { style: "solid", fill: "#1E40AF", width: 2 },
  });
  slide.shapes.add({
    geometry: "roundRect",
    position: { left: 800, top: 300, width: 400, height: 110 },
    fill: "#10B981",
    line: { style: "solid", fill: "#065F46", width: 2 },
  });
  slide.shapes.add({
    geometry: "ellipse",
    position: { left: 800, top: 430, width: 110, height: 110 },
    fill: "#F59E0B",
    line: { style: "solid", fill: "#B45309", width: 2 },
  });

  // Native image element (byte-backed) on top of the background.
  slide.images.add({
    blob: toArrayBuffer(embeddedPngBytes()),
    contentType: "image/png",
    alt: "Self-test image",
    fit: "cover",
    position: { left: 60, top: 460, width: 320, height: 120 },
  });

  // 3. Exercise the PPTX export.
  const pptxResult = await PresentationFile.exportPptx(presentation);
  const pptxBytes = await toBytes(pptxResult);
  if (pptxBytes.length === 0) {
    throw new Error("PresentationFile.exportPptx returned empty bytes");
  }

  // 4. Exercise the PNG preview export for this slide.
  const pngResult = await presentation.export({
    slide,
    format: "png",
    scale: 1,
  });
  const pngBytes = await toBytes(pngResult);
  if (pngBytes.length === 0) {
    throw new Error("presentation.export returned empty PNG bytes");
  }

  // 5. Write outputs alongside the file so the caller can inspect them.
  const outDir = path.resolve(path.dirname(fileURLToPath(import.meta.url)), ".self-test-out");
  await fs.mkdir(outDir, { recursive: true });
  const pptxPath = path.join(outDir, "selftest.pptx");
  const pngPath = path.join(outDir, "selftest.png");
  await fs.writeFile(pptxPath, pptxBytes);
  await fs.writeFile(pngPath, pngBytes);

  process.stdout.write(
    "[self-test] pptx bytes: " + pptxBytes.length + " -> " + pptxPath + "\n"
  );
  process.stdout.write(
    "[self-test] png  bytes: " + pngBytes.length + " -> " + pngPath + "\n"
  );
  process.stdout.write("[self-test] passed\n");
}

// ---------------------------------------------------------------------------
// Entry
// ---------------------------------------------------------------------------

async function runBuild(args) {
  if (!args.spec) {
    throw new Error("--spec <path.json> is required (or pass --self-test)");
  }
  const { presentation, slide } = await buildFromSpec(args.spec);

  if (args.out) {
    const pptxResult = await PresentationFile.exportPptx(presentation);
    const bytes = await toBytes(pptxResult);
    const outPath = path.resolve(args.out);
    await fs.mkdir(path.dirname(outPath), { recursive: true });
    await fs.writeFile(outPath, bytes);
    process.stdout.write("[convert] wrote pptx: " + outPath + " (" + bytes.length + " bytes)\n");
  }

  if (args.preview) {
    const pngResult = await presentation.export({
      slide,
      format: "png",
      scale: 1,
    });
    const bytes = await toBytes(pngResult);
    const outPath = path.resolve(args.preview);
    await fs.mkdir(path.dirname(outPath), { recursive: true });
    await fs.writeFile(outPath, bytes);
    process.stdout.write("[convert] wrote png:  " + outPath + " (" + bytes.length + " bytes)\n");
  }
}

async function main() {
  const args = parseArgs(process.argv);
  if (args.help) {
    printHelp();
    return;
  }
  if (args.selfTest) {
    await selfTest();
    return;
  }
  await runBuild(args);
}

main().catch((err) => {
  process.stderr.write("[convert] " + (err && err.stack ? err.stack : String(err)) + "\n");
  process.exit(1);
});
