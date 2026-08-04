import { buildElement } from "../src/elements.mjs";

let pass = 0;
let fail = 0;

function assert(cond, msg) {
  if (cond) {
    pass++;
    process.stdout.write(`  ok  ${msg}\n`);
  } else {
    fail++;
    process.stdout.write(`  FAIL ${msg}\n`);
  }
}

function approxEq(actual, expected, msg, tolerance = 1e-6) {
  assert(Math.abs(actual - expected) < tolerance, `${msg} (got ${actual} expected ${expected})`);
}

process.stdout.write("elements.mjs\n");

const text = buildElement({
  id: "title",
  kind: "text",
  text: "中文标题",
  bbox: { left: 40, top: 10, width: 600, height: 44 },
  font_family: "Arial",
  font_size: 44,
  render_strategy: "native",
});

assert(text.props.fontFamily === "Microsoft YaHei", "Chinese text uses Microsoft YaHei");
approxEq(text.props.fontSize, 44 * 72 / 96 * 0.91, "OCR pixel size is converted to calibrated points");
assert(text.bbox.width > 600 && text.bbox.height > 44, "text box gets breathing room");

const latin = buildElement({
  id: "latin",
  kind: "text",
  text: "Sim for AI",
  bbox: { left: 0, top: 0, width: 200, height: 30 },
  font_family: "Arial",
  font_size: 30,
});
assert(latin.props.fontFamily === "Arial", "Latin text keeps the requested font");

const spuriousGlyph = buildElement({
  id: "badge-noise",
  kind: "text",
  text: "B",
  bbox: { left: 65.31, top: 176, width: 24.97, height: 20.48 },
  font_size: 20.5,
  text_color: "#FEFEFE",
  editable_score: 0.756,
  confidence: { ocr: 0.756 },
});
assert(spuriousGlyph === null, "white low-confidence badge glyph is skipped");

const intentionalGlyph = buildElement({
  id: "label-a",
  kind: "text",
  text: "A",
  bbox: { left: 20, top: 20, width: 20, height: 20 },
  font_size: 18,
  text_color: "#003B8F",
  editable_score: 0.95,
  confidence: { ocr: 0.95 },
});
assert(intentionalGlyph !== null, "high-confidence dark single-letter label is kept");

if (fail > 0) {
  process.stderr.write(`\n${fail} failure(s)\n`);
  process.exit(1);
}
process.stdout.write(`\nall ${pass} assertions passed\n`);
