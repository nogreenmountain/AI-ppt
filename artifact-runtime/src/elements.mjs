import { normaliseBBox, pxToPt } from "./coordinates.mjs";

const COLOR_RE = /^#[0-9a-fA-F]{6}([0-9a-fA-F]{2})?$/;
const CHINESE_RE = /[\u3400-\u9fff]/u;
const OCR_FONT_CALIBRATION = 0.91;
const SINGLE_LATIN_RE = /^[A-Za-z]$/u;

export function resolveColor(value, fallback = "#000000") {
  if (typeof value !== "string" || !COLOR_RE.test(value)) return fallback;
  return value.toUpperCase();
}

export function resolveFontFamily(value, text = "") {
  if (CHINESE_RE.test(String(text))) return "Microsoft YaHei";
  return typeof value === "string" && value.trim() ? value.trim() : "Calibri";
}

export function colorLuminance(hexColor) {
  const value = resolveColor(hexColor, "#000000").slice(1, 7);
  const r = Number.parseInt(value.slice(0, 2), 16);
  const g = Number.parseInt(value.slice(2, 4), 16);
  const b = Number.parseInt(value.slice(4, 6), 16);
  return 0.2126 * r + 0.7152 * g + 0.0722 * b;
}

export function isSpuriousOcrGlyph(el) {
  const text = String(el.text ?? "").trim();
  if (!SINGLE_LATIN_RE.test(text)) return false;
  const bbox = normaliseBBox(el.bbox);
  const confidence = Number(el.confidence?.ocr ?? el.editable_score ?? 1);
  const isSmallBox = bbox.width <= 30 && bbox.height <= 25;
  const isWhiteInk = colorLuminance(el.text_color) >= 245;
  return isSmallBox && isWhiteInk && confidence < 0.82;
}

export function inferFontSizePt(heightPx) {
  const h = Number(heightPx);
  if (!Number.isFinite(h) || h <= 0) return 18;
  return Math.max(8, Math.min(96, Math.round(h * 2 / 3)));
}

export function buildTextElement(el) {
  const rawBBox = normaliseBBox(el.bbox);
  const extraW = Math.max(4, Math.min(18, rawBBox.width * 0.05));
  const extraH = Math.max(2, Math.min(6, rawBBox.height * 0.16));
  const bbox = {
    left: Math.max(0, rawBBox.left - extraW / 2),
    top: Math.max(0, rawBBox.top - extraH / 2),
    width: rawBBox.width + extraW,
    height: rawBBox.height + extraH,
  };
  const explicitFontSize = Number(el.font_size);
  const fontSize = Number.isFinite(explicitFontSize) && explicitFontSize > 0
    ? pxToPt(explicitFontSize) * OCR_FONT_CALIBRATION
    : inferFontSizePt(rawBBox.height);
  const fontFamily = resolveFontFamily(el.font_family, el.text);
  const color = resolveColor(el.text_color);
  const bold = el.font_weight === "bold";
  const italic = el.font_style === "italic";
  const align = ["left", "center", "right"].includes(el.align) ? el.align : "left";
  const valign = ["top", "middle", "bottom"].includes(el.valign) ? el.valign : "top";
  const children = String(el.text ?? "").split(/\r?\n/);
  const textStyleSource = [
    `${fontSize}pt`,
    fontFamily,
    bold ? "bold" : "",
    italic ? "italic" : "",
    color,
    align,
    valign,
  ].filter(Boolean).join(" ");
  return {
    id: el.id,
    z: Number(el.z) || 0,
    type: "text",
    bbox,
    props: {
      children,
      textStyleSource,
      fontFamily,
      fontSize,
      bold,
      italic,
      color,
      align,
      valign,
      rotation: Number(el.rotation) || 0,
    },
  };
}

export function buildShapeElement(el) {
  const geometry = ["rect", "roundRect", "ellipse", "line"].includes(el.geometry)
    ? el.geometry
    : "rect";
  const paint = el.fill ? resolveColor(el.fill, undefined)?.toLowerCase() : undefined;
  const lineColor = el.line_color ? resolveColor(el.line_color, "#000000") : undefined;
  const lineWidth = Number.isFinite(Number(el.line_width)) ? Number(el.line_width) : 0;
  return {
    id: el.id,
    z: Number(el.z) || 0,
    type: "shape",
    bbox: normaliseBBox(el.bbox),
    props: {
      geometry,
      paint,
      strokeSource: lineColor ? `${lineWidth || 1}px ${lineColor}` : "",
      lineColor,
      lineWidth,
      rotation: Number(el.rotation) || 0,
    },
  };
}

export function buildImageElement(el, imageBytes = null) {
  return {
    id: el.id,
    z: Number(el.z) || 0,
    type: "image",
    bbox: normaliseBBox(el.bbox),
    props: {
      imagePath: el.image_path,
      imageBytes,
      rotation: Number(el.rotation) || 0,
    },
  };
}

export function buildElement(el, imageBytes = null) {
  if (el.kind === "text") return isSpuriousOcrGlyph(el) ? null : buildTextElement(el);
  if (el.kind === "shape") return buildShapeElement(el);
  if (el.kind === "image") return buildImageElement(el, imageBytes);
  return null;
}

export function sortByZ(nodes) {
  return [...nodes].sort((a, b) => (a.z || 0) - (b.z || 0));
}
