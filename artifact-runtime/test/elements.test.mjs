// Tests for element conversion + text defaults + z-order sort. Pure JS
// (no artifact-tool runtime required).

import {
  buildTextElement,
  buildShapeElement,
  buildImageElement,
  buildElement,
  sortByZ,
  inferFontSizePt,
  resolveFontFamily,
  resolveColor,
} from "../src/elements.mjs";

let pass = 0;
let fail = 0;
function assert(cond, msg) {
  if (cond) { pass++; process.stdout.write(`  ok  ${msg}\n`); }
  else { fail++; process.stdout.write(`  FAIL ${msg}\n`); }
}
function assertEq(a, b, msg) {
  assert(a === b, `${msg} (got ${JSON.stringify(a)} expected ${JSON.stringify(b)})`);
}

process.stdout.write("elements.mjs\n");

// Text element with explicit font size + bold + color
const textEl = buildTextElement({
  id: "el_1",
  kind: "text",
  bbox: { left: 10, top: 20, width: 200, height: 40 },
  z: 0,
  text: "Hello",
  font_size: 24,
  font_weight: "bold",
  text_color: "#112233",
});
assertEq(textEl.type, "text", "text node has type=text");
assertEq(textEl.bbox.left, 10, "text bbox left");
assert(textEl.props.textStyleSource.includes("24pt"), "text style embeds pt size");
assert(textEl.props.textStyleSource.includes("bold"), "text style embeds bold");
assert(textEl.props.textStyleSource.includes("#112233"), "text style embeds color");
assertEq(textEl.props.children.length, 1, "single-line text produces one child");

// Text element with inferred font size
const textEl2 = buildTextElement({
  id: "el_2",
  kind: "text",
  bbox: { left: 0, top: 0, width: 100, height: 30 },
  z: 0,
  text: "Hi",
});
assert(textEl2.props.textStyleSource.includes("20pt"), "20pt inferred from 30px height");

// Text element with multi-line
const textEl3 = buildTextElement({
  id: "el_3",
  kind: "text",
  bbox: { left: 0, top: 0, width: 100, height: 60 },
  z: 0,
  text: "Line A\nLine B",
});
assertEq(textEl3.props.children.length, 2, "multi-line text splits");

// Shape: rect, no fill, with stroke
const rectEl = buildShapeElement({
  id: "el_4",
  kind: "shape",
  bbox: { left: 0, top: 0, width: 100, height: 50 },
  z: 0,
  geometry: "rect",
  fill: "#ff0000",
  line_color: "#000000",
  line_width: 2,
});
assertEq(rectEl.type, "shape", "shape node has type=shape");
assertEq(rectEl.props.geometry, "rect", "shape preserves geometry");
assertEq(rectEl.props.paint, "#ff0000", "shape fill normalised lowercase");
assert(rectEl.props.strokeSource.includes("2px"), "shape stroke has px width");

// Shape: ellipse
const ellEl = buildShapeElement({
  id: "el_5",
  kind: "shape",
  bbox: { left: 0, top: 0, width: 100, height: 100 },
  z: 0,
  geometry: "ellipse",
});
assertEq(ellEl.props.geometry, "ellipse", "ellipse geometry passes through");

// Shape: roundRect
const rrEl = buildShapeElement({
  id: "el_6",
  kind: "shape",
  bbox: { left: 0, top: 0, width: 100, height: 100 },
  z: 0,
  geometry: "roundRect",
});
assertEq(rrEl.props.geometry, "roundRect", "roundRect geometry passes through");

// Shape: line
const lineEl = buildShapeElement({
  id: "el_7",
  kind: "shape",
  bbox: { left: 0, top: 0, width: 200, height: 4 },
  z: 0,
  geometry: "line",
  line_color: "#00ff00",
  line_width: 3,
});
assertEq(lineEl.props.geometry, "line", "line geometry passes through");

// Image element with pre-loaded bytes
const bytes = new Uint8Array([1, 2, 3, 4]);
const imgEl = buildImageElement({
  id: "el_8",
  kind: "image",
  bbox: { left: 0, top: 0, width: 200, height: 100 },
  z: 0,
  image_path: "x.png",
}, bytes);
assertEq(imgEl.type, "image", "image node has type=image");
assert(imgEl.props.imageBytes === bytes, "image bytes are attached");

// Unknown kind -> null
const unknown = buildElement({ id: "el_x", kind: "weird", bbox: {}, z: 0 }, null);
assertEq(unknown, null, "unknown kind produces null");

// Z-order sort (lower z paints first)
const nodes = [
  buildTextElement({ id: "a", kind: "text", bbox: {}, z: 5, text: "top" }),
  buildTextElement({ id: "b", kind: "text", bbox: {}, z: 1, text: "bottom" }),
  buildTextElement({ id: "c", kind: "text", bbox: {}, z: 3, text: "mid" }),
];
const sorted = sortByZ(nodes);
assertEq(sorted[0].id, "b", "lowest z sorts first");
assertEq(sorted[1].id, "c", "mid z sorts middle");
assertEq(sorted[2].id, "a", "highest z sorts last");

// Default font size inference
assertEq(inferFontSizePt(18), 12, "inferFontSizePt 18px -> 12pt");
assertEq(inferFontSizePt(36), 24, "inferFontSizePt 36px -> 24pt");
assertEq(inferFontSizePt(3), 8, "inferFontSizePt 3px -> clamped to 8pt");
assertEq(inferFontSizePt(400), 96, "inferFontSizePt 400px -> clamped to 96pt");
assertEq(inferFontSizePt(0), 18, "inferFontSizePt 0 -> default 18");

// Font family default
assertEq(resolveFontFamily(), "Calibri", "resolveFontFamily default -> Calibri");
assertEq(resolveFontFamily("Arial"), "Arial", "resolveFontFamily passthrough");

// Color normalisation
assertEq(resolveColor("#aabbcc"), "#AABBCC", "color uppercased");
assertEq(resolveColor(), "#000000", "missing color -> #000000");

if (fail > 0) {
  process.stderr.write(`\n${fail} failure(s)\n`);
  process.exit(1);
}
process.stdout.write(`\nall ${pass} assertions passed\n`);