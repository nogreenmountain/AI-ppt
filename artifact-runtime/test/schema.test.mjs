// Tests for the schema loader. Pure JS, no runtime dependency.

import {
  loadSchemaFromObject,
  SchemaError,
} from "../src/schema.mjs";

let pass = 0;
let fail = 0;
function assert(cond, msg) {
  if (cond) { pass++; process.stdout.write(`  ok  ${msg}\n`); }
  else { fail++; process.stdout.write(`  FAIL ${msg}\n`); }
}
function assertThrows(fn, ctor, msg) {
  try {
    fn();
    fail++; process.stdout.write(`  FAIL ${msg} (no throw)\n`);
  } catch (err) {
    if (err instanceof ctor) { pass++; process.stdout.write(`  ok  ${msg}\n`); }
    else { fail++; process.stdout.write(`  FAIL ${msg} (wrong type ${err?.name})\n`); }
  }
}

process.stdout.write("schema.mjs\n");

const minimal = {
  version: "1.0",
  source: { image_path: "x.png", width_px: 1280, height_px: 720 },
  slide: { width: 1280, height: 720, unit: "px" },
  background: { strategy: "solid", image_path: null, fill: "#FFFFFF" },
  elements: [],
};

assert(loadSchemaFromObject(minimal) === minimal, "accepts a minimal valid schema");

assertThrows(
  () => loadSchemaFromObject({ ...minimal, version: "2.0" }),
  SchemaError,
  "rejects wrong version",
);
assertThrows(
  () => loadSchemaFromObject({ ...minimal, slide: undefined }),
  SchemaError,
  "rejects missing slide",
);
assertThrows(
  () => loadSchemaFromObject({ ...minimal, elements: "not an array" }),
  SchemaError,
  "rejects non-array elements",
);
assertThrows(
  () => loadSchemaFromObject({
    ...minimal,
    elements: [{ id: "el_a", kind: "weird", bbox: { left: 0, top: 0, width: 10, height: 10 }, z: 0 }],
  }),
  SchemaError,
  "rejects unknown element kind",
);
assertThrows(
  () => loadSchemaFromObject({
    ...minimal,
    elements: [{ id: "el_a", kind: "text", bbox: { left: 0, top: 0, width: 10, height: 10 } }], // no z
  }),
  SchemaError,
  "rejects element without z",
);
assertThrows(
  () => loadSchemaFromObject({
    ...minimal,
    elements: [{ id: "el_a", kind: "text", z: 0 }], // no bbox
  }),
  SchemaError,
  "rejects element without bbox",
);

if (fail > 0) {
  process.stderr.write(`\n${fail} failure(s)\n`);
  process.exit(1);
}
process.stdout.write(`\nall ${pass} assertions passed\n`);