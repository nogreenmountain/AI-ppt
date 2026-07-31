"""Core detection logic for slide2pptx.

This module turns a slide screenshot into a structured
``detected.json`` description plus a background PNG, following
``spec/detected.schema.json``. The implementation is intentionally
small and degrades gracefully when optional dependencies are missing:

* Pillow + NumPy are guaranteed baselines (always required).
* ``rapidocr_onnxruntime`` (optional) extracts OCR boxes/text/confidence
  and produces native text elements.
* ``cv2`` / OpenCV (optional) inpaints the OCR regions into
  ``cleaned-background.png``; otherwise the original image is copied
  as ``original-background.png`` and a warning is recorded.

The slide canvas is always mapped to 1280x720 -- bounding boxes are
scaled linearly from the source image dimensions.

No network calls are performed and no shape detection is included in
this module (per the task scope).
"""

from __future__ import annotations

import io
import json
import uuid
from dataclasses import dataclass, field
from pathlib import Path
from typing import Any, Iterable

import numpy as np
from PIL import Image

try:  # pragma: no cover - exercised only when optional dep is installed
    from rapidocr_onnxruntime import RapidOCR  # type: ignore
except Exception:  # noqa: BLE001 - any import failure is treated as "missing"
    RapidOCR = None  # type: ignore[assignment]

try:  # pragma: no cover - exercised only when optional dep is installed
    import cv2  # type: ignore
except Exception:  # noqa: BLE001
    cv2 = None  # type: ignore[assignment]


SCHEMA_VERSION = "1.0"
SLIDE_WIDTH = 1280.0
SLIDE_HEIGHT = 720.0


# ---------------------------------------------------------------------------
# Helpers
# ---------------------------------------------------------------------------


def _scale_box(
    box: Iterable[float],
    src_w: int,
    src_h: int,
    dst_w: float = SLIDE_WIDTH,
    dst_h: float = SLIDE_HEIGHT,
) -> dict[str, float]:
    """Convert a raw 4-corner OCR box (in source pixels) to a schema bbox.

    ``box`` is interpreted as ``[[x0, y0], [x1, y1], [x2, y2], [x3, y3]]``
    where the corners follow the order returned by RapidOCR. The smallest
    enclosing axis-aligned rectangle is computed and rescaled to the
    target 1280x720 canvas.
    """

    pts = np.asarray(box, dtype=np.float64).reshape(-1, 2)
    xs = pts[:, 0]
    ys = pts[:, 1]
    left = float(xs.min())
    top = float(ys.min())
    right = float(xs.max())
    bottom = float(ys.max())
    sx = dst_w / float(src_w)
    sy = dst_h / float(src_h)
    return {
        "left": round(left * sx, 2),
        "top": round(top * sy, 2),
        "width": round((right - left) * sx, 2),
        "height": round((bottom - top) * sy, 2),
    }


def _scale_xywh(
    x: float,
    y: float,
    w: float,
    h: float,
    src_w: int,
    src_h: int,
    dst_w: float = SLIDE_WIDTH,
    dst_h: float = SLIDE_HEIGHT,
) -> dict[str, float]:
    sx = dst_w / float(src_w)
    sy = dst_h / float(src_h)
    return {
        "left": round(float(x) * sx, 2),
        "top": round(float(y) * sy, 2),
        "width": round(float(w) * sx, 2),
        "height": round(float(h) * sy, 2),
    }


def _sample_text_color(rgb: np.ndarray, mask: np.ndarray) -> str:
    """Estimate text colour as pixels furthest from the local background.

    ``rgb`` is an HxWx3 uint8 array, ``mask`` is a boolean array of the
    same HxW shape selecting the pixels inside the OCR box. Text is usually
    the minority colour, so the per-channel median is a useful background
    estimate for both dark-on-light and light-on-dark labels. Returns
    ``#RRGGBB``.
    """

    if not mask.any():
        # Fallback: opaque black. We have nothing to sample.
        return "#000000"

    pixels = rgb[mask].astype(np.float32)
    background = np.median(pixels, axis=0)
    distance = np.linalg.norm(pixels - background, axis=1)
    threshold = float(np.percentile(distance, 85))
    ink = pixels[distance >= threshold]
    if not len(ink):
        ink = pixels
    colour = np.median(ink, axis=0)
    colour = np.clip(colour, 0, 255)
    r, g, b = (int(round(c)) for c in colour)
    return f"#{r:02X}{g:02X}{b:02X}"


def _estimate_font_size(box_height_px: float) -> float:
    """Roughly estimate a font size (points) from the box height.

    Artifact-tool uses CSS-like pixel font sizes. OCR boxes include glyph
    ascenders/descenders and anti-alias padding, so roughly 80% of the box
    height is a safer starting point than converting the height to points.
    """

    if box_height_px <= 0:
        return 12.0
    return max(round(box_height_px, 1), 8.0)


def _make_id(prefix: str = "el") -> str:
    return f"{prefix}_{uuid.uuid4().hex[:12]}"


def _inflate_xyxy(
    box: tuple[float, float, float, float],
    pad: float,
    width: int,
    height: int,
) -> tuple[int, int, int, int]:
    x0, y0, x1, y1 = box
    return (
        max(int(np.floor(x0 - pad)), 0),
        max(int(np.floor(y0 - pad)), 0),
        min(int(np.ceil(x1 + pad)), width),
        min(int(np.ceil(y1 + pad)), height),
    )


def _make_box_mask(
    boxes: list[tuple[float, float, float, float]],
    width: int,
    height: int,
    pad: float,
) -> np.ndarray:
    mask = np.zeros((height, width), dtype=bool)
    for box in boxes:
        x0, y0, x1, y1 = _inflate_xyxy(box, pad, width, height)
        if x1 > x0 and y1 > y0:
            mask[y0:y1, x0:x1] = True
    return mask


def _component_text_overlap(
    bbox: tuple[int, int, int, int],
    text_boxes: list[tuple[float, float, float, float]],
) -> float:
    x, y, w, h = bbox
    area = max(w * h, 1)
    overlap = 0.0
    x1 = x + w
    y1 = y + h
    for tx0, ty0, tx1, ty1 in text_boxes:
        ix0 = max(float(x), tx0)
        iy0 = max(float(y), ty0)
        ix1 = min(float(x1), tx1)
        iy1 = min(float(y1), ty1)
        if ix1 > ix0 and iy1 > iy0:
            overlap += (ix1 - ix0) * (iy1 - iy0)
    return float(overlap / area)


def _component_overlap_ratio(
    a: tuple[int, int, int, int],
    b: tuple[float, float, float, float],
) -> float:
    """Return how much of ``a`` is covered by xyxy box ``b``."""

    x, y, w, h = a
    if w <= 0 or h <= 0:
        return 0.0
    ax0 = float(x)
    ay0 = float(y)
    ax1 = float(x + w)
    ay1 = float(y + h)
    bx0, by0, bx1, by1 = b
    ix0 = max(ax0, bx0)
    iy0 = max(ay0, by0)
    ix1 = min(ax1, bx1)
    iy1 = min(ay1, by1)
    if ix1 <= ix0 or iy1 <= iy0:
        return 0.0
    return float(((ix1 - ix0) * (iy1 - iy0)) / max(w * h, 1))


def _bbox_iou(
    a: tuple[int, int, int, int],
    b: tuple[int, int, int, int],
) -> float:
    ax, ay, aw, ah = a
    bx, by, bw, bh = b
    if aw <= 0 or ah <= 0 or bw <= 0 or bh <= 0:
        return 0.0
    ax1 = ax + aw
    ay1 = ay + ah
    bx1 = bx + bw
    by1 = by + bh
    ix0 = max(ax, bx)
    iy0 = max(ay, by)
    ix1 = min(ax1, bx1)
    iy1 = min(ay1, by1)
    if ix1 <= ix0 or iy1 <= iy0:
        return 0.0
    inter = float((ix1 - ix0) * (iy1 - iy0))
    union = float(aw * ah + bw * bh) - inter
    return inter / union if union > 0 else 0.0


def _bbox_is_near_duplicate(
    a: tuple[int, int, int, int],
    b: tuple[int, int, int, int],
) -> bool:
    area_a = max(1, a[2] * a[3])
    area_b = max(1, b[2] * b[3])
    return _bbox_iou(a, b) > 0.72 and min(area_a, area_b) / max(area_a, area_b) > 0.45


def _detect_text_like_regions(
    src_img: Image.Image,
    known_text_boxes: list[tuple[float, float, float, float]] | None = None,
) -> list[tuple[float, float, float, float]]:
    """Find text-like ink regions for visual-component suppression.

    This is intentionally *not* OCR. It does not try to read text or emit
    native text elements. Its only job is to protect visual extraction from
    swallowing nearby glyphs when OCR is unavailable or incomplete.
    """

    if cv2 is None:
        return []

    known_text_boxes = known_text_boxes or []
    rgb = np.asarray(src_img.convert("RGB"), dtype=np.uint8)
    src_h, src_w = rgb.shape[:2]
    if src_w <= 0 or src_h <= 0:
        return []

    gray = cv2.cvtColor(rgb, cv2.COLOR_RGB2GRAY)
    hsv = cv2.cvtColor(rgb, cv2.COLOR_RGB2HSV)
    saturation = hsv[:, :, 1]

    # Text is usually local high-frequency ink. A large blur approximates the
    # local background so both dark and coloured text can be picked up without
    # hard-coding slide colours.
    blur_size = max(31, int(round(min(src_w, src_h) / 18)) | 1)
    background = cv2.GaussianBlur(gray, (blur_size, blur_size), 0)
    contrast = cv2.absdiff(gray, background)
    dark_or_coloured = (gray < 220) | (saturation > 35)
    raw = ((contrast > 18) & dark_or_coloured).astype(np.uint8) * 255

    # Remove obvious thin separators/lines before glyph filtering. Text glyphs
    # survive because they are broken into smaller blobs at this stage.
    line_kernel_w = max(35, int(round(src_w * 0.035)))
    horizontal_kernel = cv2.getStructuringElement(cv2.MORPH_RECT, (line_kernel_w, 1))
    horizontal = cv2.morphologyEx(raw, cv2.MORPH_OPEN, horizontal_kernel, iterations=1)
    raw[horizontal > 0] = 0

    num_labels, labels, stats, _ = cv2.connectedComponentsWithStats(raw, 8)
    glyph_mask = np.zeros((src_h, src_w), dtype=np.uint8)
    total_area = float(src_w * src_h)
    min_area = max(6.0, total_area * 0.000003)
    max_area = max(1500.0, total_area * 0.0025)
    min_h = max(5, int(round(src_h * 0.004)))
    max_h = max(28, int(round(src_h * 0.075)))

    for label in range(1, num_labels):
        x = int(stats[label, cv2.CC_STAT_LEFT])
        y = int(stats[label, cv2.CC_STAT_TOP])
        w = int(stats[label, cv2.CC_STAT_WIDTH])
        h = int(stats[label, cv2.CC_STAT_HEIGHT])
        area = float(stats[label, cv2.CC_STAT_AREA])
        if area < min_area or area > max_area:
            continue
        if h < min_h or h > max_h:
            continue
        if w < 2 or w > src_w * 0.16:
            continue
        aspect = w / max(h, 1)
        coverage = area / float(max(w * h, 1))
        if aspect < 0.04 or aspect > 12.0:
            continue
        if coverage < 0.05 or coverage > 0.88:
            continue
        # Keep known OCR boxes in the protection mask, but avoid promoting
        # large icon pieces that sit well outside any text-like dimensions.
        glyph_mask[labels == label] = 255

    if not np.any(glyph_mask):
        return []

    # Join characters into words/short lines, not whole panels. Horizontal
    # dilation is purposely modest so adjacent icons do not get absorbed.
    join_w = max(9, int(round(src_w * 0.012)))
    join_h = max(3, int(round(src_h * 0.004)))
    join_kernel = cv2.getStructuringElement(cv2.MORPH_RECT, (join_w, join_h))
    line_mask = cv2.dilate(glyph_mask, join_kernel, iterations=1)

    num_lines, line_labels, line_stats, _ = cv2.connectedComponentsWithStats(line_mask, 8)
    boxes: list[tuple[float, float, float, float]] = []
    for label in range(1, num_lines):
        x = int(line_stats[label, cv2.CC_STAT_LEFT])
        y = int(line_stats[label, cv2.CC_STAT_TOP])
        w = int(line_stats[label, cv2.CC_STAT_WIDTH])
        h = int(line_stats[label, cv2.CC_STAT_HEIGHT])
        area = float(line_stats[label, cv2.CC_STAT_AREA])
        if area < max(18.0, total_area * 0.000006):
            continue
        if h < min_h or h > src_h * 0.10:
            continue
        if w < max(5, src_w * 0.004):
            continue
        aspect = w / max(h, 1)
        if aspect < 0.18 or aspect > 38.0:
            continue
        # Exclude regions already nearly covered by OCR boxes; callers append
        # OCR boxes explicitly, so duplicating them only inflates masks.
        if any(_component_overlap_ratio((x, y, w, h), box) > 0.72 for box in known_text_boxes):
            continue
        pad_x = max(2, int(round(h * 0.18)))
        pad_y = max(1, int(round(h * 0.08)))
        x0, y0, x1, y1 = _inflate_xyxy((x, y, x + w, y + h), max(pad_x, pad_y), src_w, src_h)
        boxes.append((float(x0), float(y0), float(x1), float(y1)))

    # Large slides can produce many tiny title glyph groups; merge boxes that
    # touch after padding so suppression remains stable and cheap.
    boxes.sort(key=lambda b: (b[1], b[0]))
    merged: list[tuple[float, float, float, float]] = []
    for box in boxes:
        bx0, by0, bx1, by1 = box
        merged_into_existing = False
        for idx, current in enumerate(merged):
            cx0, cy0, cx1, cy1 = current
            same_line = min(by1, cy1) - max(by0, cy0) > 0.35 * min(by1 - by0, cy1 - cy0)
            close = bx0 <= cx1 + max(10, (cy1 - cy0) * 1.6) and bx1 >= cx0 - max(10, (cy1 - cy0) * 1.6)
            if same_line and close:
                merged[idx] = (
                    min(cx0, bx0),
                    min(cy0, by0),
                    max(cx1, bx1),
                    max(cy1, by1),
                )
                merged_into_existing = True
                break
        if not merged_into_existing:
            merged.append(box)

    return merged


def _median_mask_color(rgb: np.ndarray, mask: np.ndarray) -> str:
    if not mask.any():
        return "#000000"
    pixels = rgb[mask].reshape(-1, 3)
    colour = np.median(pixels, axis=0)
    r, g, b = (int(round(c)) for c in np.clip(colour, 0, 255))
    return f"#{r:02X}{g:02X}{b:02X}"


def _mask_color_std(rgb: np.ndarray, mask: np.ndarray) -> float:
    if not mask.any():
        return 255.0
    pixels = rgb[mask].reshape(-1, 3).astype(np.float32)
    return float(np.mean(np.std(pixels, axis=0)))


def _bbox_ink_coverage(rgb: np.ndarray, bbox: tuple[int, int, int, int]) -> float:
    x, y, w, h = bbox
    if w <= 0 or h <= 0:
        return 0.0
    crop = rgb[y : y + h, x : x + w].astype(np.float32)
    if crop.size == 0:
        return 0.0
    border_parts = [
        crop[: max(1, h // 12), :, :],
        crop[-max(1, h // 12) :, :, :],
        crop[:, : max(1, w // 12), :],
        crop[:, -max(1, w // 12) :, :],
    ]
    border = np.concatenate([part.reshape(-1, 3) for part in border_parts], axis=0)
    background = np.median(border, axis=0)
    distance = np.linalg.norm(crop - background, axis=2)
    return float(np.mean(distance > 28.0))


def _classify_simple_shape(
    component_mask: np.ndarray,
    rgb: np.ndarray,
    bbox: tuple[int, int, int, int],
) -> tuple[str | None, str]:
    """Return a native PPT geometry when a component is very simple."""

    if cv2 is None:
        return None, "#000000"

    x, y, w, h = bbox
    if w <= 0 or h <= 0:
        return None, "#000000"

    fill = _median_mask_color(rgb, component_mask)
    if _mask_color_std(rgb, component_mask) > 22.0:
        return None, fill

    crop_mask = component_mask[y : y + h, x : x + w]
    mask_u8 = crop_mask.astype(np.uint8) * 255
    contours, _ = cv2.findContours(mask_u8, cv2.RETR_EXTERNAL, cv2.CHAIN_APPROX_SIMPLE)
    if not contours:
        return None, fill

    contour = max(contours, key=cv2.contourArea)
    area = float(cv2.contourArea(contour))
    bbox_area = float(max(w * h, 1))
    extent = area / bbox_area
    perimeter = float(cv2.arcLength(contour, True))
    circularity = (4.0 * np.pi * area / (perimeter * perimeter)) if perimeter > 0 else 0.0

    if min(w, h) <= max(4, int(max(w, h) * 0.08)):
        return "line", fill
    if extent > 0.82:
        return "rect", fill
    if circularity > 0.72 and extent > 0.58:
        return "ellipse", fill
    return None, fill


def _save_component_png(
    rgb: np.ndarray,
    component_mask: np.ndarray,
    bbox: tuple[int, int, int, int],
    assets_dir: Path,
    element_id: str,
) -> Path:
    x, y, w, h = bbox
    crop_rgb = rgb[y : y + h, x : x + w]
    crop_mask = component_mask[y : y + h, x : x + w]
    rgba = np.dstack([crop_rgb, crop_mask.astype(np.uint8) * 255])
    assets_dir.mkdir(parents=True, exist_ok=True)
    out_path = assets_dir / f"{element_id}.png"
    Image.fromarray(rgba, mode="RGBA").save(out_path, format="PNG")
    return out_path


def _save_component_alpha_png(
    rgb: np.ndarray,
    alpha_mask: np.ndarray,
    bbox: tuple[int, int, int, int],
    assets_dir: Path,
    element_id: str,
) -> Path:
    x, y, w, h = bbox
    crop_rgb = rgb[y : y + h, x : x + w]
    crop_alpha = alpha_mask[y : y + h, x : x + w]
    rgba = np.dstack([crop_rgb, crop_alpha.astype(np.uint8)])
    assets_dir.mkdir(parents=True, exist_ok=True)
    out_path = assets_dir / f"{element_id}.png"
    Image.fromarray(rgba, mode="RGBA").save(out_path, format="PNG")
    return out_path


def _bbox_from_mask(mask: np.ndarray) -> tuple[int, int, int, int] | None:
    ys, xs = np.where(mask)
    if not len(xs):
        return None
    x0 = int(xs.min())
    y0 = int(ys.min())
    x1 = int(xs.max()) + 1
    y1 = int(ys.max()) + 1
    return x0, y0, x1 - x0, y1 - y0


def _extract_visual_components(
    src_img: Image.Image,
    out_dir: Path,
    text_boxes: list[tuple[float, float, float, float]],
    max_components: int = 28,
) -> tuple[list[dict[str, Any]], list[np.ndarray], list[str]]:
    """Extract non-text visual blocks as native shapes or transparent PNGs."""

    warnings: list[str] = []
    if cv2 is None:
        return [], [], warnings

    rgb = np.asarray(src_img.convert("RGB"), dtype=np.uint8)
    src_h, src_w = rgb.shape[:2]
    if src_w <= 0 or src_h <= 0:
        return [], [], warnings

    gray = cv2.cvtColor(rgb, cv2.COLOR_RGB2GRAY)
    hsv = cv2.cvtColor(rgb, cv2.COLOR_RGB2HSV)
    saturation = hsv[:, :, 1]

    edge = cv2.Canny(gray, 45, 140)
    blur_size = max(21, int(round(min(src_w, src_h) / 28)) | 1)
    blurred = cv2.GaussianBlur(gray, (blur_size, blur_size), 0)
    contrast = cv2.absdiff(gray, blurred)
    contrast_mask = (contrast > 18).astype(np.uint8) * 255
    colour_mask = ((saturation > 48) & (contrast > 10)).astype(np.uint8) * 255
    raw_mask = cv2.bitwise_or(edge, cv2.bitwise_or(contrast_mask, colour_mask))

    text_pad = max(5, int(round(min(src_w, src_h) * 0.008)))
    text_mask = _make_box_mask(text_boxes, src_w, src_h, text_pad)
    raw_mask[text_mask] = 0

    kernel = cv2.getStructuringElement(
        cv2.MORPH_ELLIPSE,
        (
            max(3, int(round(src_w * 0.006)) | 1),
            max(3, int(round(src_h * 0.009)) | 1),
        ),
    )
    mask = cv2.morphologyEx(raw_mask, cv2.MORPH_CLOSE, kernel, iterations=2)
    mask = cv2.dilate(mask, kernel, iterations=1)
    mask[text_mask] = 0

    num_labels, labels, stats, _ = cv2.connectedComponentsWithStats(mask, 8)
    total_area = float(src_w * src_h)
    min_area = max(90.0, total_area * 0.00035)
    max_area = total_area * 0.12
    min_w = max(8, int(round(src_w * 0.01)))
    min_h = max(8, int(round(src_h * 0.012)))

    candidates: list[tuple[int, int, int, int, int, float]] = []
    for label in range(1, num_labels):
        x = int(stats[label, cv2.CC_STAT_LEFT])
        y = int(stats[label, cv2.CC_STAT_TOP])
        w = int(stats[label, cv2.CC_STAT_WIDTH])
        h = int(stats[label, cv2.CC_STAT_HEIGHT])
        area = float(stats[label, cv2.CC_STAT_AREA])
        if area < min_area or area > max_area:
            continue
        if w > src_w * 0.65 and h > src_h * 0.45:
            continue
        if w < min_w and h < min_h:
            continue
        aspect = w / max(h, 1)
        coverage = area / float(max(w * h, 1))
        if aspect < 0.04 or aspect > 25:
            continue
        ink_coverage = _bbox_ink_coverage(rgb, (x, y, w, h))
        if aspect > 2.0 and ink_coverage < 0.38 and h < src_h * 0.16:
            continue
        if _component_text_overlap((x, y, w, h), text_boxes) > 0.35:
            continue
        candidates.append((label, x, y, w, h, area))

    candidates.sort(key=lambda item: item[5], reverse=True)
    assets_dir = out_dir / "assets"
    elements: list[dict[str, Any]] = []
    masks: list[np.ndarray] = []

    for idx, (label, x, y, w, h, area) in enumerate(candidates[:max_components]):
        component_mask = labels == label
        component_u8 = component_mask.astype(np.uint8) * 255
        contours, _ = cv2.findContours(component_u8, cv2.RETR_EXTERNAL, cv2.CHAIN_APPROX_SIMPLE)
        filled = np.zeros_like(component_u8)
        if contours:
            cv2.drawContours(filled, contours, -1, 255, thickness=cv2.FILLED)
            component_mask = filled.astype(bool)

        pad = max(2, int(round(min(w, h) * 0.04)))
        x0, y0, x1, y1 = _inflate_xyxy((x, y, x + w, y + h), pad, src_w, src_h)
        bbox = (x0, y0, x1 - x0, y1 - y0)
        padded_mask = np.zeros((src_h, src_w), dtype=bool)
        padded_mask[y0:y1, x0:x1] = component_mask[y0:y1, x0:x1]

        geometry, fill = _classify_simple_shape(padded_mask, rgb, bbox)
        element_id = _make_id("el")
        scaled = _scale_xywh(bbox[0], bbox[1], bbox[2], bbox[3], src_w, src_h)
        common: dict[str, Any] = {
            "id": element_id,
            "bbox": scaled,
            "z": -100 + idx,
            "editable_score": 0.62 if geometry else 0.42,
            "confidence": {
                "visual_component": round(min(area / max(total_area * 0.02, 1), 1.0), 3)
            },
            "metadata": {
                "detector": "opencv_residual_components",
                "source_bbox_px": {
                    "left": bbox[0],
                    "top": bbox[1],
                    "width": bbox[2],
                    "height": bbox[3],
                },
                "source_area_px": int(area),
            },
        }

        if geometry:
            element = {
                **common,
                "kind": "shape",
                "render_strategy": "native",
                "geometry": geometry,
                "fill": fill,
                "line_color": fill,
                "line_width": 0,
            }
        else:
            try:
                png_path = _save_component_png(rgb, padded_mask, bbox, assets_dir, element_id)
            except Exception as exc:  # noqa: BLE001
                warnings.append(f"visual component export failed: {exc}")
                continue
            element = {
                **common,
                "kind": "image",
                "render_strategy": "image",
                "image_path": str(png_path.resolve()),
                "fit": "contain",
            }
        elements.append(element)
        masks.append(padded_mask)

    if candidates and not elements:
        warnings.append("visual components were detected but none could be exported.")

    return elements, masks, warnings


def _extract_iterative_residual_components(
    residual_img: Image.Image,
    out_dir: Path,
    *,
    pass_index: int = 2,
    max_components: int = 96,
) -> tuple[list[dict[str, Any]], list[np.ndarray], list[str]]:
    """Extract visual leftovers after text/first-pass components are removed.

    The first detector intentionally stays conservative so text and nearby
    icons do not get glued together. This second pass runs on the inpainted
    residual and looks for lower-contrast slide structures: container frames,
    gradient bars, arrows, logo pieces, dashed curves and remaining icons.
    """

    warnings: list[str] = []
    if cv2 is None:
        return [], [], warnings

    rgb = np.asarray(residual_img.convert("RGB"), dtype=np.uint8)
    src_h, src_w = rgb.shape[:2]
    if src_w <= 0 or src_h <= 0:
        return [], [], warnings

    hsv = cv2.cvtColor(rgb, cv2.COLOR_RGB2HSV)
    gray = cv2.cvtColor(rgb, cv2.COLOR_RGB2GRAY)
    blurred = cv2.GaussianBlur(gray, (61, 61), 0)
    contrast = cv2.absdiff(gray, blurred)

    hue = hsv[:, :, 0]
    saturation = hsv[:, :, 1]
    value = hsv[:, :, 2]
    raw_masks: list[tuple[str, np.ndarray, int]] = [
        (
            "blue_visual",
            (hue >= 85) & (hue <= 125) & (saturation > 35) & (value > 45),
            3,
        ),
        (
            "green_visual",
            (hue >= 35) & (hue <= 90) & (saturation > 25) & (value > 55),
            3,
        ),
        (
            "gold_red_visual",
            ((hue <= 30) | (hue >= 165)) & (saturation > 35) & (value > 55),
            3,
        ),
    ]

    colour_union = np.zeros((src_h, src_w), dtype=bool)
    for _name, raw, _priority in raw_masks:
        colour_union |= raw

    candidates: list[dict[str, Any]] = []
    for name, raw, priority in raw_masks:
        mask = raw.astype(np.uint8) * 255
        mask = cv2.morphologyEx(
            mask,
            cv2.MORPH_CLOSE,
            cv2.getStructuringElement(cv2.MORPH_ELLIPSE, (3, 3)),
            iterations=1,
        )
        mask = cv2.morphologyEx(
            mask,
            cv2.MORPH_OPEN,
            cv2.getStructuringElement(cv2.MORPH_RECT, (2, 2)),
            iterations=1,
        )
        num_labels, labels, stats, _ = cv2.connectedComponentsWithStats(mask, 8)
        for label in range(1, num_labels):
            x = int(stats[label, cv2.CC_STAT_LEFT])
            y = int(stats[label, cv2.CC_STAT_TOP])
            w = int(stats[label, cv2.CC_STAT_WIDTH])
            h = int(stats[label, cv2.CC_STAT_HEIGHT])
            area = int(stats[label, cv2.CC_STAT_AREA])
            if area < 45 or w < 3 or h < 3:
                continue
            if w > src_w * 0.96 and h > src_h * 0.25:
                continue
            aspect = w / max(h, 1)
            if aspect > 160 or aspect < 0.015:
                continue
            component_mask = labels == label
            bbox = _bbox_from_mask(component_mask)
            if bbox is None:
                continue
            candidates.append(
                {
                    "class": name,
                    "bbox": bbox,
                    "area": area,
                    "mask": component_mask,
                    "priority": priority,
                }
            )

    # Faint structures: low-contrast frames, separators and shadows. Subtract
    # colour so this layer does not swallow icons already represented above.
    expanded_colour = cv2.dilate(
        colour_union.astype(np.uint8),
        cv2.getStructuringElement(cv2.MORPH_ELLIPSE, (7, 7)),
        iterations=1,
    ).astype(bool)
    structural_raw = (contrast > 8) & (gray < 249) & ~expanded_colour
    structural_mask = structural_raw.astype(np.uint8) * 255
    structural_mask = cv2.morphologyEx(
        structural_mask,
        cv2.MORPH_CLOSE,
        cv2.getStructuringElement(cv2.MORPH_RECT, (5, 5)),
        iterations=1,
    )
    num_labels, labels, stats, _ = cv2.connectedComponentsWithStats(structural_mask, 8)
    total_area = float(src_w * src_h)
    for label in range(1, num_labels):
        x = int(stats[label, cv2.CC_STAT_LEFT])
        y = int(stats[label, cv2.CC_STAT_TOP])
        w = int(stats[label, cv2.CC_STAT_WIDTH])
        h = int(stats[label, cv2.CC_STAT_HEIGHT])
        area = int(stats[label, cv2.CC_STAT_AREA])
        if area < 120 or w < 8 or h < 3:
            continue
        if w > src_w * 0.94 and h > src_h * 0.24:
            continue
        if w * h > total_area * 0.18:
            continue
        aspect = w / max(h, 1)
        if aspect > 300 or aspect < 0.02:
            continue
        component_mask = labels == label
        bbox = _bbox_from_mask(component_mask)
        if bbox is None:
            continue
        candidates.append(
            {
                "class": "faint_structure",
                "bbox": bbox,
                "area": area,
                "mask": component_mask,
                "priority": 1,
            }
        )

    candidates.sort(key=lambda c: (int(c["priority"]), int(c["area"])), reverse=True)
    kept: list[dict[str, Any]] = []
    for candidate in candidates:
        bbox = candidate["bbox"]
        if any(_bbox_is_near_duplicate(bbox, other["bbox"]) for other in kept):
            continue
        kept.append(candidate)
        if len(kept) >= max_components:
            break

    # Structural elements should sit closest to the background. Remaining
    # coloured objects are layered above them, but below first-pass objects and
    # native text.
    kept.sort(key=lambda c: (int(c["priority"]), int(c["area"])))
    assets_dir = out_dir / "assets"
    elements: list[dict[str, Any]] = []
    masks: list[np.ndarray] = []
    for idx, candidate in enumerate(kept):
        x, y, w, h = candidate["bbox"]
        local_alpha = candidate["mask"].astype(np.uint8) * 255
        local_alpha = cv2.dilate(
            local_alpha,
            cv2.getStructuringElement(cv2.MORPH_ELLIPSE, (3, 3)),
            iterations=1,
        )
        local_alpha = cv2.GaussianBlur(local_alpha, (3, 3), 0)
        alpha_mask = np.zeros((src_h, src_w), dtype=np.uint8)
        alpha_mask[y : y + h, x : x + w] = local_alpha[y : y + h, x : x + w]

        element_id = _make_id(f"el_pass{pass_index}")
        try:
            png_path = _save_component_alpha_png(
                rgb,
                alpha_mask,
                (x, y, w, h),
                assets_dir,
                element_id,
            )
        except Exception as exc:  # noqa: BLE001
            warnings.append(f"iterative component export failed: {exc}")
            continue

        area = int(candidate["area"])
        elements.append(
            {
                "id": element_id,
                "kind": "image",
                "bbox": _scale_xywh(x, y, w, h, src_w, src_h),
                "z": -300 + idx,
                "editable_score": 0.38,
                "render_strategy": "image",
                "image_path": str(png_path.resolve()),
                "fit": "contain",
                "confidence": {
                    "second_pass_visual": round(min(area / 12000.0, 1.0), 3)
                },
                "metadata": {
                    "detector": "iterative_residual_components",
                    "pass": pass_index,
                    "class": str(candidate["class"]),
                    "source_bbox_px": {
                        "left": int(x),
                        "top": int(y),
                        "width": int(w),
                        "height": int(h),
                    },
                    "source_area_px": area,
                },
            }
        )
        masks.append(alpha_mask > 8)

    return elements, masks, warnings


def _inpaint_rgb_with_cv2(rgb: np.ndarray, masks: list[np.ndarray]) -> np.ndarray:
    """Inpaint an RGB array over the union of masks."""

    assert cv2 is not None
    bgr = cv2.cvtColor(rgb, cv2.COLOR_RGB2BGR)
    union = np.zeros(bgr.shape[:2], dtype=np.uint8)
    h, w = union.shape
    for mask in masks:
        if mask.shape != union.shape:
            resized = np.zeros(union.shape, dtype=bool)
            hh = min(mask.shape[0], h)
            ww = min(mask.shape[1], w)
            resized[:hh, :ww] = mask[:hh, :ww]
            mask = resized
        union[mask] = 255
    if np.any(union):
        union = cv2.dilate(
            union,
            cv2.getStructuringElement(cv2.MORPH_ELLIPSE, (5, 5)),
            iterations=1,
        )
    inpainted = cv2.inpaint(bgr, union, 3.0, cv2.INPAINT_TELEA)
    return cv2.cvtColor(inpainted, cv2.COLOR_BGR2RGB)


def _inpaint_with_cv2(image_path: Path, masks: list[np.ndarray]) -> np.ndarray:
    """Inpaint ``image_path`` using ``cv2.inpaint`` over the union of masks.

    Each mask in ``masks`` is a boolean HxW array. The function returns
    the inpainted RGB uint8 array.
    """

    assert cv2 is not None  # caller guarantees this
    bgr = cv2.imread(str(image_path), cv2.IMREAD_COLOR)
    if bgr is None:  # pragma: no cover - defensive
        raise RuntimeError(f"cv2 could not read {image_path}")
    union = np.zeros(bgr.shape[:2], dtype=np.uint8)
    h, w = union.shape
    for m in masks:
        if m.shape != union.shape:
            # Crop / pad the mask if needed (shouldn't happen in practice).
            resized = np.zeros(union.shape, dtype=bool)
            hh = min(m.shape[0], h)
            ww = min(m.shape[1], w)
            resized[:hh, :ww] = m[:hh, :ww]
            m = resized
        union[m] = 255
    # 3 px radius matches typical OCR character widths.
    inpainted = cv2.inpaint(bgr, union, 3.0, cv2.INPAINT_TELEA)
    rgb = cv2.cvtColor(inpainted, cv2.COLOR_BGR2RGB)
    return rgb


# ---------------------------------------------------------------------------
# Synthetic image (for --self-test)
# ---------------------------------------------------------------------------


def _make_synthetic_slide(size: tuple[int, int] = (1280, 720)) -> Image.Image:
    """Return a small synthetic slide with two text-like dark patches."""

    w, h = size
    # Light grey background with a soft vertical gradient.
    arr = np.zeros((h, w, 3), dtype=np.uint8)
    for y in range(h):
        v = int(230 + (y / max(h - 1, 1)) * 20)
        arr[y, :, :] = (v, v, v)

    img = Image.fromarray(arr, mode="RGB")
    # Paint two dark rectangles that look like text lines.
    from PIL import ImageDraw

    draw = ImageDraw.Draw(img)
    # Title bar
    draw.rectangle((80, 90, 720, 170), fill=(20, 20, 30))
    # Sub line
    draw.rectangle((80, 220, 560, 260), fill=(40, 40, 50))
    return img


# ---------------------------------------------------------------------------
# Result container
# ---------------------------------------------------------------------------


@dataclass
class DetectResult:
    payload: dict[str, Any]
    background_path: Path
    detections_json_path: Path
    warnings: list[str] = field(default_factory=list)


# ---------------------------------------------------------------------------
# Public API
# ---------------------------------------------------------------------------


def detect(
    image_path: Path | str,
    out_dir: Path | str,
    *,
    visual_passes: int = 2,
    second_pass_max_components: int = 96,
) -> DetectResult:
    """Run detection on ``image_path`` and write outputs to ``out_dir``.

    Always writes ``detected.json`` and a background PNG (``original-`` or
    ``cleaned-`` depending on optional dependencies). Returns a
    :class:`DetectResult` summarising what was produced.
    """

    src_path = Path(image_path)
    if not src_path.is_file():
        raise FileNotFoundError(f"Input image not found: {src_path}")

    out_dir = Path(out_dir)
    out_dir.mkdir(parents=True, exist_ok=True)

    src_img = Image.open(src_path).convert("RGB")
    src_w, src_h = src_img.size

    warnings: list[str] = []

    # Background handling -----------------------------------------------------
    original_bg_path = out_dir / "original-background.png"
    src_img.save(original_bg_path, format="PNG")

    background_strategy = "original"
    background_image_path: Path | None = original_bg_path

    # Try OCR first so we know which boxes to mask before inpainting.
    ocr_text_elements: list[dict[str, Any]] = []
    ocr_masks: list[np.ndarray] = []
    ocr_boxes_src: list[tuple[float, float, float, float]] = []

    if RapidOCR is not None:
        try:
            engine = RapidOCR()
            ocr_result = engine(str(src_path))
            # rapidocr-onnxruntime returns either:
            #   (boxes, texts, scores)            -- older versions
            #   (boxes, texts, scores, det_elapse) -- newer versions
            #   or None when nothing is detected.
            boxes: list[Any] = []
            texts: list[str] = []
            scores: list[float] = []
            if ocr_result is not None:
                # rapidocr_onnxruntime 1.4 returns ``(lines, elapsed)`` where
                # each line is ``[box, text, score]``. Older releases returned
                # ``(boxes, texts, scores, ...)``. Accept both layouts.
                if (
                    len(ocr_result) == 2
                    and isinstance(ocr_result[0], (list, tuple))
                    and ocr_result[0]
                    and isinstance(ocr_result[0][0], (list, tuple))
                    and len(ocr_result[0][0]) >= 3
                ):
                    lines = list(ocr_result[0])
                    boxes = [line[0] for line in lines]
                    texts = [line[1] for line in lines]
                    scores = [line[2] for line in lines]
                elif len(ocr_result) >= 3:
                    boxes = list(ocr_result[0] or [])
                    texts = list(ocr_result[1] or [])
                    scores = list(ocr_result[2] or [])
                else:  # pragma: no cover - unusual shape
                    boxes = list(ocr_result[0] or [])
            for idx, box in enumerate(boxes):
                if box is None:
                    continue
                try:
                    bbox = _scale_box(box, src_w, src_h)
                except Exception:
                    continue
                text = texts[idx] if idx < len(texts) else ""
                if text is None:
                    text = ""
                text_str = str(text).strip()
                try:
                    conf = float(scores[idx]) if idx < len(scores) else 0.0
                except (TypeError, ValueError):
                    conf = 0.0
                height_px = float(bbox["height"])
                font_size = _estimate_font_size(height_px)
                element: dict[str, Any] = {
                    "id": _make_id("el"),
                    "kind": "text",
                    "bbox": bbox,
                    "z": idx,
                    "editable_score": round(min(max(conf, 0.0), 1.0), 3),
                    "render_strategy": "native",
                    "text": text_str,
                    "font_size": font_size,
                    "align": "left",
                    "valign": "middle",
                    "confidence": {"ocr": round(min(max(conf, 0.0), 1.0), 3)},
                }
                ocr_text_elements.append(element)
                # Track source-pixel box for later masking / colour sampling.
                pts = np.asarray(box, dtype=np.float64).reshape(-1, 2)
                xs = pts[:, 0]
                ys = pts[:, 1]
                ocr_boxes_src.append(
                    (float(xs.min()), float(ys.min()), float(xs.max()), float(ys.max()))
                )
        except Exception as exc:  # noqa: BLE001 - never let OCR kill detection
            warnings.append(f"OCR failed: {exc}")
    else:
        warnings.append(
            "rapidocr_onnxruntime not installed; native text elements skipped."
        )

    visual_text_boxes_src = list(ocr_boxes_src)
    heuristic_text_boxes_src: list[tuple[float, float, float, float]] = []
    if cv2 is not None:
        try:
            heuristic_text_boxes_src = _detect_text_like_regions(
                src_img,
                known_text_boxes=ocr_boxes_src,
            )
            visual_text_boxes_src.extend(heuristic_text_boxes_src)
            if heuristic_text_boxes_src and not ocr_text_elements:
                warnings.append(
                    "heuristic text-like regions detected for component separation; "
                    "install rapidocr_onnxruntime to emit editable text."
                )
        except Exception as exc:  # noqa: BLE001
            warnings.append(f"heuristic text-region detection failed: {exc}")

    # Build per-box masks for optional inpainting.
    if ocr_boxes_src:
        rgb_array = np.asarray(src_img, dtype=np.uint8)
        for (x0, y0, x1, y1) in ocr_boxes_src:
            padding = max(4, int(round((y1 - y0) * 0.12)))
            ix0, iy0, ix1, iy1 = _inflate_xyxy(
                (x0, y0, x1, y1), padding, src_w, src_h
            )
            mask = np.zeros((src_h, src_w), dtype=bool)
            mask[iy0:iy1, ix0:ix1] = True
            ocr_masks.append(mask)
        # Estimate text colour from the first available mask.
        for element, mask in zip(ocr_text_elements, ocr_masks):
            colour = _sample_text_color(rgb_array, mask)
            element["text_color"] = colour
            # Best-effort weight guess: dark text on light bg is bold-ish.
            luminance = _hex_luminance(colour)
            element["font_weight"] = "bold" if luminance < 90 else "normal"
            element["font_style"] = "normal"
            element["font_family"] = "Arial"

    visual_elements: list[dict[str, Any]] = []
    visual_masks: list[np.ndarray] = []
    second_pass_visual_elements: list[dict[str, Any]] = []
    second_pass_visual_masks: list[np.ndarray] = []
    if cv2 is not None:
        try:
            visual_elements, visual_masks, visual_warnings = _extract_visual_components(
                src_img,
                out_dir,
                visual_text_boxes_src,
            )
            warnings.extend(visual_warnings)
        except Exception as exc:  # noqa: BLE001 - preserve OCR output on failure
            warnings.append(f"visual component extraction failed: {exc}")
    else:
        warnings.append("cv2 not installed; visual component extraction skipped.")

    # Iterative second pass: remove first-pass text/visual masks, then detect
    # lower-contrast remaining structures on the residual image. First-pass
    # elements are kept and added back to the final payload.
    first_pass_masks = [*ocr_masks, *visual_masks]
    if visual_passes >= 2 and first_pass_masks and cv2 is not None:
        try:
            src_rgb = np.asarray(src_img, dtype=np.uint8)
            first_residual_rgb = _inpaint_rgb_with_cv2(src_rgb, first_pass_masks)
            first_residual_path = out_dir / "residual-pass1-background.png"
            Image.fromarray(first_residual_rgb, mode="RGB").save(
                first_residual_path,
                format="PNG",
            )
            (
                second_pass_visual_elements,
                second_pass_visual_masks,
                second_pass_warnings,
            ) = _extract_iterative_residual_components(
                Image.fromarray(first_residual_rgb, mode="RGB"),
                out_dir,
                pass_index=2,
                max_components=second_pass_max_components,
            )
            warnings.extend(second_pass_warnings)
        except Exception as exc:  # noqa: BLE001
            warnings.append(f"iterative visual component extraction failed: {exc}")

    # Inpaint when cv2 is present and we have something to mask out.
    all_inpaint_masks = [*first_pass_masks, *second_pass_visual_masks]
    if all_inpaint_masks and cv2 is not None:
        try:
            cleaned_rgb = _inpaint_rgb_with_cv2(
                np.asarray(src_img, dtype=np.uint8),
                all_inpaint_masks,
            )
            cleaned_path = out_dir / "cleaned-background.png"
            Image.fromarray(cleaned_rgb, mode="RGB").save(cleaned_path, format="PNG")
            background_strategy = "cleaned"
            background_image_path = cleaned_path
        except Exception as exc:  # noqa: BLE001 - fall back to original
            warnings.append(f"cv2 inpainting failed: {exc}; keeping original.")
    elif all_inpaint_masks and cv2 is None:
        warnings.append(
            "cv2 not installed; OCR regions kept in original-background.png."
        )

    # Solid fill (best-effort average corner colour, useful for downstream).
    solid_fill = _sample_corner_color(np.asarray(src_img))

    payload: dict[str, Any] = {
        "version": SCHEMA_VERSION,
        "source": {
            "image_path": str(src_path.resolve()),
            "width_px": int(src_w),
            "height_px": int(src_h),
        },
        "slide": {
            "width": SLIDE_WIDTH,
            "height": SLIDE_HEIGHT,
            "unit": "px",
        },
        "background": {
            "strategy": background_strategy,
            "image_path": (
                str(background_image_path.resolve())
                if background_image_path is not None
                else None
            ),
            "fill": solid_fill,
        },
        "elements": [
            *second_pass_visual_elements,
            *visual_elements,
            *ocr_text_elements,
        ],
        "metrics": {
            "element_count": (
                len(second_pass_visual_elements)
                + len(visual_elements)
                + len(ocr_text_elements)
            ),
            "text_element_count": len(ocr_text_elements),
            "visual_component_count": len(visual_elements)
            + len(second_pass_visual_elements),
            "first_pass_visual_component_count": len(visual_elements),
            "second_pass_visual_component_count": len(second_pass_visual_elements),
            "heuristic_text_region_count": len(heuristic_text_boxes_src),
            "visual_passes": int(max(1, visual_passes)),
            "source_width_px": int(src_w),
            "source_height_px": int(src_h),
        },
        "warnings": list(warnings),
    }

    json_path = out_dir / "detected.json"
    json_path.write_text(json.dumps(payload, indent=2, ensure_ascii=False), encoding="utf-8")

    return DetectResult(
        payload=payload,
        background_path=background_image_path if background_image_path is not None else original_bg_path,
        detections_json_path=json_path,
        warnings=warnings,
    )


def _sample_corner_color(rgb: np.ndarray) -> str:
    """Return ``#RRGGBB`` for the average colour of the top-left 8x8 patch."""

    h, w = rgb.shape[:2]
    patch = rgb[: max(min(8, h), 1), : max(min(8, w), 1)]
    mean = patch.reshape(-1, 3).mean(axis=0)
    r, g, b = (int(round(c)) for c in np.clip(mean, 0, 255))
    return f"#{r:02X}{g:02X}{b:02X}"


def _hex_luminance(hex_color: str) -> float:
    """Return ITU-R BT.601 luminance for ``#RRGGBB``."""

    s = hex_color.lstrip("#")
    if len(s) != 6:
        return 128.0
    r = int(s[0:2], 16)
    g = int(s[2:4], 16)
    b = int(s[4:6], 16)
    return 0.2126 * r + 0.7152 * g + 0.0722 * b


def run_self_test() -> dict[str, Any]:
    """Create a temporary synthetic slide and run ``detect`` on it.

    Returns the parsed JSON payload as a dict. Raises ``AssertionError``
    if the schema or output files are wrong -- callers can assert on the
    return value to confirm the basics look correct.
    """

    import tempfile

    with tempfile.TemporaryDirectory() as tmp:
        tmp_path = Path(tmp)
        img_path = tmp_path / "synthetic.png"
        out_dir = tmp_path / "out"
        _make_synthetic_slide().save(img_path, format="PNG")
        result = detect(img_path, out_dir)

        assert result.detections_json_path.is_file(), "detected.json missing"
        assert result.background_path.is_file(), "background PNG missing"
        assert result.payload["version"] == SCHEMA_VERSION
        assert "elements" in result.payload
        assert isinstance(result.payload["elements"], list)
        for el in result.payload["elements"]:
            assert el["kind"] in {"text", "shape", "image"}
            assert set(el["bbox"].keys()) == {"left", "top", "width", "height"}
            assert el["bbox"]["width"] >= 0
            assert el["bbox"]["height"] >= 0
        # Slide canvas is always 1280x720.
        assert result.payload["slide"]["width"] == 1280
        assert result.payload["slide"]["height"] == 720
        assert result.payload["slide"]["unit"] == "px"

        return result.payload


__all__ = ["detect", "run_self_test", "DetectResult"]
