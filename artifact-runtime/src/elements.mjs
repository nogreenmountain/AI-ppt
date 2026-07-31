import { normaliseBBox } from "./coordinates.mjs";

const COLOR_RE = /^#[0-9a-fA-F]{6}([0-9a-fA-F]{2})?$/;

export function resolveColor(value, fallback = "#000000") {
  if (typeof value !== "string" || !COLOR_RE.test(value)) return fallback;
  return value.toUpperCase();
}

export function resolveFontFamily(value) {
  return typeof value === "string" && value.trim() ? value.trim() : "Calibri";
}

export function inferFontSizePt(heightPx) {
  const h = Number(heightPx);
  if (!Number.isFinite(h) || h <= 0) return 18;
  return Math.max(8, Math.min(96, Math.round(h * 2 / 3)));
}

export function buildTextElement(el) {
  const bbox = normaliseBBox(el.bbox);
  const fontSize = Number(el.font_size) || inferFontSizePt(bbox.height);
  const fontFamily = resolveFontFamily(el.font_family);
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
  if (el.kind === "text") return buildTextElement(el);
  if (el.kind === "shape") return buildShapeElement(el);
  if (el.kind === "image") return buildImageElement(el, imageBytes);
  return null;
}

export function sortByZ(nodes) {
  return [...nodes].sort((a, b) => (a.z || 0) - (b.z || 0));
}
