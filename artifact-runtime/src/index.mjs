import { mkdir, writeFile } from "node:fs/promises";
import { dirname } from "node:path";

import PptxGenJS from "pptxgenjs";

import { bboxToInches, pxToIn, resolveSlideSize } from "./coordinates.mjs";
import { buildElement, sortByZ } from "./elements.mjs";
import { imageDataUri, loadImageBytes } from "./image-loader.mjs";
import { loadSchemaFromPath } from "./schema.mjs";

const MIME_PPTX = "application/vnd.openxmlformats-officedocument.presentationml.presentation";

function toBytes(value) {
  if (value instanceof Uint8Array) return value;
  if (value instanceof ArrayBuffer) return new Uint8Array(value);
  if (ArrayBuffer.isView(value)) {
    return new Uint8Array(value.buffer, value.byteOffset, value.byteLength);
  }
  if (Buffer.isBuffer(value)) return new Uint8Array(value);
  throw new Error(`Unsupported PPTX output type: ${typeof value}`);
}

async function ensureParent(filePath) {
  await mkdir(dirname(filePath), { recursive: true });
}

async function applyBackground(schema, slide, imageLoader, size) {
  if (schema.background?.strategy === "solid") {
    slide.background = { color: (schema.background.fill || "#FFFFFF").replace("#", "") };
    return;
  }
  if (!schema.background?.image_path) return;
  const loaded = await imageLoader("background", schema.background.image_path);
  slide.addImage({
    data: imageDataUri(loaded.bytes, loaded.contentType),
    x: 0,
    y: 0,
    w: size.widthIn,
    h: size.heightIn,
  });
}

function addText(slide, node) {
  const box = bboxToInches(node.bbox);
  const p = node.props;
  slide.addText(p.children.join("\n"), {
    ...box,
    margin: 0,
    breakLine: false,
    wrap: false,
    fit: "shrink",
    fontFace: p.fontFamily,
    fontSize: p.fontSize,
    bold: p.bold,
    italic: p.italic,
    color: p.color.replace("#", ""),
    align: p.align,
    valign: p.valign === "middle" ? "mid" : p.valign,
    rotate: p.rotation || undefined,
    transparency: 0,
    marginPt: 0,
  });
}

function addShape(slide, node) {
  const box = bboxToInches(node.bbox);
  const p = node.props;
  const line = p.lineColor
    ? { color: p.lineColor.replace("#", ""), width: p.lineWidth || 1 }
    : { color: "FFFFFF", transparency: 100, width: 0 };

  if (p.geometry === "line") {
    slide.addShape("line", {
      x: box.x,
      y: box.y,
      w: box.w,
      h: box.h || 0,
      line,
      rotate: p.rotation || undefined,
    });
    return;
  }

  slide.addShape(p.geometry, {
    ...box,
    fill: p.paint
      ? { color: p.paint.replace("#", "") }
      : { color: "FFFFFF", transparency: 100 },
    line,
    rotate: p.rotation || undefined,
  });
}

async function addImage(slide, node, baseDir, imageLoader) {
  const loaded = node.props.imageBytes
    ? { bytes: node.props.imageBytes, contentType: "image/png" }
    : await imageLoader(node.id, node.props.imagePath);
  slide.addImage({
    data: imageDataUri(loaded.bytes, loaded.contentType),
    ...bboxToInches(node.bbox),
    rotate: node.props.rotation || undefined,
    altText: node.id || "Detected image",
  });
}

export async function buildPptxFromSchema(schema, options = {}) {
  const baseDir = options.baseDir || process.cwd();
  const imageLoader = options.imageLoader || ((_id, p) => loadImageBytes(p, baseDir));
  const size = resolveSlideSize(schema.slide);

  const pptx = new PptxGenJS();
  pptx.author = "slide2pptx";
  pptx.subject = "Generated from a detected slide image";
  pptx.title = "Reconstructed slide";
  pptx.company = "";
  pptx.lang = "zh-CN";
  pptx.defineLayout({ name: "SLIDE2PPTX_CUSTOM", width: size.widthIn, height: size.heightIn });
  pptx.layout = "SLIDE2PPTX_CUSTOM";

  const slide = pptx.addSlide();
  await applyBackground(schema, slide, imageLoader, size);

  const nodes = sortByZ((schema.elements || []).map((el) => buildElement(el)).filter(Boolean));
  for (const node of nodes) {
    if (node.type === "text") addText(slide, node);
    else if (node.type === "shape") addShape(slide, node);
    else if (node.type === "image") await addImage(slide, node, baseDir, imageLoader);
  }

  const output = await pptx.write({ outputType: "nodebuffer", compression: true });
  return {
    bytes: toBytes(output),
    mime: MIME_PPTX,
    slideSize: size,
  };
}

export async function writePptx(bytes, outPath) {
  await ensureParent(outPath);
  await writeFile(outPath, bytes);
  return bytes.length;
}

export async function writePngPreview(_schema, _outPath, _options = {}) {
  throw new Error("PNG preview is not available in the open-source builder; use the report step with PowerPoint rendering.");
}

export async function runBuild(args = {}) {
  const input = args.input || args.spec;
  if (!input) throw new Error("runBuild requires input/spec");
  const schema = await loadSchemaFromPath(input);
  const baseDir = args.baseDir || dirname(input);
  const result = await buildPptxFromSchema(schema, { baseDir, imageLoader: args.imageLoader });
  const out = args.out;
  if (!out) return result;
  const written = await writePptx(result.bytes, out);
  if (!args.quiet) {
    process.stdout.write(`[build] wrote pptx: ${out} (${written} bytes)\n`);
  }
  if (args.preview && !args.skipPreview && !args.quiet) {
    process.stdout.write("[build] PNG preview skipped; run the report step for a rendered preview.\n");
  }
  return { ...result, pptxPath: out, bytesWritten: written };
}

export {
  bboxToInches,
  buildElement,
  loadImageBytes,
  loadSchemaFromPath,
  pxToIn,
  resolveSlideSize,
  sortByZ,
};
