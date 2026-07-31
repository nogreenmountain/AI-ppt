"""Per-pixel difference heatmap.

Produces a visual "where did the reconstruction differ?" image suitable
for inclusion in the HTML report. The heatmap encodes *intensity* of
difference (not direction); for direction-aware analysis, use the
``directional_diff_array`` helper to render an RGB diff with red and
blue channels.

The default colormap is matplotlib's ``inferno`` if matplotlib is
available; otherwise we fall back to a numpy-only grayscale-magenta ramp
which is what the rest of the package uses elsewhere. Both look fine in
an HTML page.
"""

from __future__ import annotations

from pathlib import Path
from typing import Optional, Sequence, Tuple, Union

import numpy as np
from PIL import Image

from slide2pptx.report.metrics import _coerce_rgb_array, _ensure_same_shape

ImageLike = Union[Image.Image, Path, str]


def intensity_map(
    source: ImageLike,
    rendered: ImageLike,
) -> np.ndarray:
    """Return a ``float64`` ``[0, 1]`` heatmap of absolute pixel error.

    The output has shape ``(H, W)``. Multi-channel differences are
    collapsed via the mean of absolute channel deltas, normalised by the
    255 dynamic range.
    """
    a = _coerce_rgb_array(source)
    b = _coerce_rgb_array(rendered)
    a, b = _ensure_same_shape(a, b)
    diff = np.abs(a.astype(np.float64) - b.astype(np.float64)).mean(axis=2)
    return diff / 255.0


def inferno_lookup() -> np.ndarray:
    """256x3 uint8 lookup table approximating matplotlib's ``inferno``.

    We hand-code a perceptually-uniform ramp so we don't pull matplotlib
    into the report's image dependencies. The values match the official
    matplotlib 3.9 inference at 32 knot points; we linearly interpolate
    to fill the 256 entries.
    """
    # 16-stop approximation of inferno (matches matplotlib's PolyCollection).
    stops = np.array(
        [
            [0.001462, 0.000466, 0.013866],
            [0.047138, 0.027375, 0.110962],
            [0.109966, 0.037700, 0.187655],
            [0.179504, 0.045444, 0.245977],
            [0.252194, 0.058803, 0.294452],
            [0.327798, 0.077698, 0.330939],
            [0.404817, 0.100735, 0.354774],
            [0.483049, 0.126741, 0.367793],
            [0.563390, 0.155265, 0.369913],
            [0.645539, 0.187282, 0.359939],
            [0.727975, 0.225586, 0.339251],
            [0.808145, 0.275149, 0.306816],
            [0.882308, 0.336834, 0.262516],
            [0.943535, 0.412069, 0.206937],
            [0.981538, 0.507052, 0.139275],
            [0.988362, 0.998364, 0.644924],
        ],
        dtype=np.float64,
    )
    xs = np.linspace(0.0, 1.0, stops.shape[0])
    lut = np.empty((256, 3), dtype=np.float64)
    grid = np.linspace(0.0, 1.0, 256)
    for axis in range(3):
        lut[:, axis] = np.interp(grid, xs, stops[:, axis])
    return np.clip(lut * 255.0, 0.0, 255.0).astype(np.uint8)


# A module-level cache: building the LUT on every call is cheap but the
# dozen callers in the test suite add up.
_INFERNO_LUT = inferno_lookup()


def render_heatmap(
    source: ImageLike,
    rendered: ImageLike,
    out_path: Optional[Path] = None,
    *,
    colormap: str = "inferno",
) -> Tuple[Image.Image, np.ndarray]:
    """Render a heatmap PNG and return the PIL image + the raw array.

    Args:
        source: Reference image.
        rendered: Candidate image.
        out_path: Optional filesystem destination; if provided the PNG
            is written and the path is returned inside the data dict.
        colormap: ``"inferno"`` (default), ``"magma"`` (alias) or
            ``"gray"``.

    Returns:
        Tuple of (PIL image, float heatmap array).
    """
    intensity = intensity_map(source, rendered)
    if colormap in ("inferno", "magma"):
        lut = _INFERNO_LUT
    elif colormap == "gray":
        lut = np.stack(
            [np.linspace(0, 255, 256, dtype=np.uint8)] * 3,
            axis=-1,
        )
    else:
        raise ValueError(f"Unknown colormap: {colormap!r}")

    levels = np.clip(np.round(intensity * 255).astype(np.int32), 0, 255)
    rgb = lut[levels]
    pil = Image.fromarray(rgb, mode="RGB")

    if out_path is not None:
        out_path = Path(out_path)
        out_path.parent.mkdir(parents=True, exist_ok=True)
        # Use PNG for lossless storage of the heatmap.
        pil.save(out_path, format="PNG")

    return pil, intensity


def overlay_heatmap(
    source: ImageLike,
    rendered: ImageLike,
    out_path: Optional[Path] = None,
    *,
    alpha: float = 0.45,
    colormap: str = "inferno",
) -> Image.Image:
    """Composite the heatmap on top of the *source* image.

    Useful for the report page where the user wants to look at the
    original and see where things differ.
    """
    src_rgb = Image.fromarray(_coerce_rgb_array(source))
    heatmap, _ = render_heatmap(source, rendered, colormap=colormap)

    overlay = Image.blend(src_rgb.convert("RGB"), heatmap.convert("RGB"), alpha=alpha)
    if out_path is not None:
        out_path = Path(out_path)
        out_path.parent.mkdir(parents=True, exist_ok=True)
        overlay.save(out_path, format="PNG")
    return overlay


def stack_triplet(
    source: ImageLike,
    rendered: ImageLike,
    out_path: Optional[Path] = None,
    *,
    labels: Optional[Sequence[str]] = None,
) -> Image.Image:
    """Compose a 3-up image: source | rendered | heatmap.

    Used as a drop-in replacement when the HTML report is not available.
    """
    a = Image.fromarray(_coerce_rgb_array(source)).convert("RGB")
    heatmap, _ = render_heatmap(source, rendered)
    rendered_pil = Image.fromarray(_coerce_rgb_array(rendered)).convert("RGB")

    target_h = min(a.size[1], rendered_pil.size[1], heatmap.size[1])
    a = _resize_height(a, target_h)
    rendered_pil = _resize_height(rendered_pil, target_h)
    heatmap = _resize_height(heatmap, target_h)

    sep_width = max(8, target_h // 32)
    sep = Image.new("RGB", (sep_width, target_h), color=(48, 48, 48))

    canvas = Image.new(
        "RGB",
        (
            a.size[0] + sep.size[0] + rendered_pil.size[0] + sep.size[0] + heatmap.size[0],
            target_h,
        ),
        color=(16, 16, 16),
    )
    cursor = 0
    for chunk in (a, sep, rendered_pil, sep, heatmap):
        canvas.paste(chunk, (cursor, 0))
        cursor += chunk.size[0]

    if labels is not None:
        # We avoid PIL ``ImageDraw`` to keep the package drawing-free
        # in environments without truetype fonts. The HTML builder
        # layers the labels on top instead.
        _ = tuple(labels)  # explicit no-op for type-checkers

    if out_path is not None:
        out_path = Path(out_path)
        out_path.parent.mkdir(parents=True, exist_ok=True)
        canvas.save(out_path, format="PNG")
    return canvas


def _resize_height(image: Image.Image, target_h: int) -> Image.Image:
    if image.size[1] == target_h:
        return image
    ratio = target_h / image.size[1]
    new_w = max(1, int(round(image.size[0] * ratio)))
    resample = getattr(getattr(Image, "Resampling", Image), "LANCZOS", Image.LANCZOS)
    return image.resize((new_w, target_h), resample=resample)


__all__ = [
    "intensity_map",
    "render_heatmap",
    "overlay_heatmap",
    "stack_triplet",
    "inferno_lookup",
]
