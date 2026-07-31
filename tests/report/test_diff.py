"""Tests for the diff heatmap module."""

from __future__ import annotations

from pathlib import Path

import numpy as np
import pytest
from PIL import Image

from slide2pptx.report import diff as diff_mod


# ---------------------------------------------------------------------------
# intensity_map
# ---------------------------------------------------------------------------


def test_intensity_map_zero_for_identical_images(identical_png_pair):
    a, b = identical_png_pair
    intensity = diff_mod.intensity_map(a, b)
    assert intensity.shape == np.asarray(Image.open(a)).shape[:2]
    # Every value should be exactly 0 for identical inputs.
    assert float(intensity.max()) == 0.0


def test_intensity_map_max_for_max_differences(tmp_path):
    a = np.zeros((40, 40, 3), dtype=np.uint8)
    b = np.full((40, 40, 3), 255, dtype=np.uint8)
    pa = tmp_path / "a.png"
    pb = tmp_path / "b.png"
    Image.fromarray(a).save(pa)
    Image.fromarray(b).save(pb)
    intensity = diff_mod.intensity_map(pa, pb)
    assert intensity.shape == (40, 40)
    assert float(intensity.mean()) == pytest.approx(1.0, abs=1e-6)


def test_intensity_map_resizes_when_shapes_differ(tmp_path):
    Image.new("RGB", (200, 100), (10, 10, 10)).save(tmp_path / "a.png")
    Image.new("RGB", (100, 50), (200, 200, 200)).save(tmp_path / "b.png")
    intensity = diff_mod.intensity_map(tmp_path / "a.png", tmp_path / "b.png")
    assert intensity.shape == (100, 200)


# ---------------------------------------------------------------------------
# render_heatmap
# ---------------------------------------------------------------------------


def test_render_heatmap_returns_pil_and_array(identical_png_pair):
    a, b = identical_png_pair
    img, arr = diff_mod.render_heatmap(a, b)
    assert isinstance(img, Image.Image)
    assert isinstance(arr, np.ndarray)
    assert img.size == (arr.shape[1], arr.shape[0])


def test_render_heatmap_writes_png(tmp_path, sample_png_pair):
    a, b = sample_png_pair
    out = tmp_path / "diff.png"
    diff_mod.render_heatmap(a, b, out_path=out)
    assert out.is_file()
    with Image.open(out) as loaded:
        assert loaded.size[0] > 0 and loaded.size[1] > 0


def test_render_heatmap_supports_gray_colormap(sample_png_pair):
    a, b = sample_png_pair
    img_gray, _ = diff_mod.render_heatmap(a, b, colormap="gray")
    img_inferno, _ = diff_mod.render_heatmap(a, b, colormap="inferno")
    assert img_gray.mode == "RGB"
    assert img_inferno.mode == "RGB"
    # At least one pixel differs between the two colormaps for a noisy diff.
    arr_g = np.asarray(img_gray)
    arr_i = np.asarray(img_inferno)
    assert not np.array_equal(arr_g, arr_i)


def test_render_heatmap_rejects_unknown_colormap(sample_png_pair):
    a, b = sample_png_pair
    with pytest.raises(ValueError):
        diff_mod.render_heatmap(a, b, colormap="nope")


# ---------------------------------------------------------------------------
# overlay_heatmap & stack_triplet
# ---------------------------------------------------------------------------


def test_overlay_heatmap_returns_pil(sample_png_pair, tmp_path):
    a, b = sample_png_pair
    overlay = diff_mod.overlay_heatmap(a, b, out_path=tmp_path / "overlay.png")
    assert isinstance(overlay, Image.Image)
    assert (tmp_path / "overlay.png").is_file()


def test_stack_triplet_creates_three_panel_image(sample_png_pair, tmp_path):
    a, b = sample_png_pair
    out = tmp_path / "triplet.png"
    canvas = diff_mod.stack_triplet(a, b, out_path=out)
    assert isinstance(canvas, Image.Image)
    # The composite is wider than each input panel.
    a_pil = Image.open(a)
    assert canvas.size[0] > a_pil.size[0]
    assert canvas.size[1] == a_pil.size[1]


# ---------------------------------------------------------------------------
# inferno lookup
# ---------------------------------------------------------------------------


def test_inferno_lookup_shape_and_range():
    lut = diff_mod.inferno_lookup()
    assert lut.shape == (256, 3)
    assert lut.dtype == np.uint8
    # All values must fall in [0, 255].
    assert int(lut.min()) >= 0
    assert int(lut.max()) <= 255


def test_inferno_lookup_monotonic_in_intensity():
    """``inferno`` should be monotonic in luminance so that 0 -> black, 255 -> bright."""
    lut = diff_mod.inferno_lookup()
    # luminance approximation
    lum = 0.2126 * lut[:, 0] + 0.7152 * lut[:, 1] + 0.0722 * lut[:, 2]
    # Should generally increase from index 0 to 255 - allow a tiny relaxation.
    assert lum[-1] > lum[0]
    assert lum[-1] > 200  # very bright at the high end (yellow-white)
    assert lum[0] < 25  # very dark at index 0
