// Focused unit tests for the coordinate / unit helpers. These are pure
// functions, so they can run without the pptxgenjs runtime being present.

import {
  pxToEmu,
  emuToPx,
  pxToPt,
  ptToPx,
  normaliseBBox,
  resolveSlideSize,
  DEFAULT_SLIDE_WIDTH_PX,
  DEFAULT_SLIDE_HEIGHT_PX,
  EMU_PER_INCH,
} from "../src/coordinates.mjs";

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
function assertEq(a, b, msg) {
  assert(a === b, `${msg} (got ${a} expected ${b})`);
}
function approxEq(a, b, msg, tol = 1e-6) {
  assert(Math.abs(a - b) < tol, `${msg} (got ${a} expected ${b})`);
}

process.stdout.write("coordinates.mjs\n");
assertEq(pxToEmu(96), EMU_PER_INCH, "pxToEmu(96) == EMU_PER_INCH");
assertEq(emuToPx(EMU_PER_INCH), 96, "emuToPx(EMU_PER_INCH) == 96");
assertEq(pxToEmu(0), 0, "pxToEmu(0) == 0");
assertEq(pxToEmu(1280), 12192000, "pxToEmu(1280) == 12192000 (1280x720 slide width)");
assertEq(pxToEmu(720), 6858000, "pxToEmu(720) == 6858000 (1280x720 slide height)");
approxEq(pxToPt(96), 72, "pxToPt(96) == 72");
approxEq(ptToPx(72), 96, "ptToPx(72) == 96");
approxEq(ptToPx(18), 24, "ptToPx(18) == 24");

assertEq(
  JSON.stringify(normaliseBBox({ left: 1, top: 2, width: 3, height: 4 })),
  JSON.stringify({ left: 1, top: 2, width: 3, height: 4 }),
  "normaliseBBox passes through left/top/width/height",
);
assertEq(
  JSON.stringify(normaliseBBox({ x: 1, y: 2, w: 3, h: 4 })),
  JSON.stringify({ left: 1, top: 2, width: 3, height: 4 }),
  "normaliseBBox converts x/y/w/h",
);

const slidePx = resolveSlideSize({ width: 1280, height: 720, unit: "px" });
assertEq(slidePx.widthPx, 1280, "resolveSlideSize defaults to px");
assertEq(slidePx.heightPx, 720, "resolveSlideSize defaults to px");

const slidePt = resolveSlideSize({ width: 960, height: 540, unit: "pt" });
assertEq(slidePt.widthPx, 1280, "pt -> px: 960pt -> 1280px");
assertEq(slidePt.heightPx, 720, "pt -> px: 540pt -> 720px");

assertEq(DEFAULT_SLIDE_WIDTH_PX, 1280, "DEFAULT_SLIDE_WIDTH_PX == 1280");
assertEq(DEFAULT_SLIDE_HEIGHT_PX, 720, "DEFAULT_SLIDE_HEIGHT_PX == 720");

if (fail > 0) {
  process.stderr.write(`\n${fail} failure(s)\n`);
  process.exit(1);
}
process.stdout.write(`\nall ${pass} assertions passed\n`);
