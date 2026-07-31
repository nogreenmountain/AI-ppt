export const PX_PER_INCH = 96;
export const PT_PER_INCH = 72;
export const EMU_PER_INCH = 914400;
export const DEFAULT_SLIDE_WIDTH_PX = 1280;
export const DEFAULT_SLIDE_HEIGHT_PX = 720;

export function pxToIn(px) {
  return Number(px || 0) / PX_PER_INCH;
}

export function inToPx(inches) {
  return Number(inches || 0) * PX_PER_INCH;
}

export function pxToEmu(px) {
  return Math.round(pxToIn(px) * EMU_PER_INCH);
}

export function emuToPx(emu) {
  return Math.round((Number(emu || 0) / EMU_PER_INCH) * PX_PER_INCH);
}

export function pxToPt(px) {
  return (Number(px || 0) / PX_PER_INCH) * PT_PER_INCH;
}

export function ptToPx(pt) {
  return (Number(pt || 0) / PT_PER_INCH) * PX_PER_INCH;
}

export function normaliseBBox(bbox = {}) {
  const num = (value) => {
    const parsed = Number(value);
    return Number.isFinite(parsed) ? parsed : 0;
  };
  return {
    left: num(bbox.left ?? bbox.x),
    top: num(bbox.top ?? bbox.y),
    width: num(bbox.width ?? bbox.w),
    height: num(bbox.height ?? bbox.h),
  };
}

export function resolveSlideSize(slide = {}) {
  const unit = slide.unit || "px";
  const width = Number(slide.width) || DEFAULT_SLIDE_WIDTH_PX;
  const height = Number(slide.height) || DEFAULT_SLIDE_HEIGHT_PX;
  const widthPx = unit === "pt" ? ptToPx(width) : width;
  const heightPx = unit === "pt" ? ptToPx(height) : height;
  return {
    widthPx: Math.round(widthPx),
    heightPx: Math.round(heightPx),
    widthIn: pxToIn(widthPx),
    heightIn: pxToIn(heightPx),
  };
}

export function bboxToInches(bbox = {}) {
  const b = normaliseBBox(bbox);
  return {
    x: pxToIn(b.left),
    y: pxToIn(b.top),
    w: pxToIn(b.width),
    h: pxToIn(b.height),
  };
}
