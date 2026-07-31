"""Tests for the report metrics module.

These tests use synthetic NumPy arrays so they are deterministic and
do not require Pillow to find installed system fonts.
"""

from __future__ import annotations

import math
from pathlib import Path

import numpy as np
import pytest
from PIL import Image

from slide2pptx.report import metrics
from slide2pptx.report.models import DiffMetrics


# ---------------------------------------------------------------------------
# MAE / RMSE
# ---------------------------------------------------------------------------


def test_mae_zero_on_identical_arrays(identical_png_pair):
    a, b = identical_png_pair
    assert metrics.mean_absolute_error(a, b) == 0.0


def test_mae_matches_manual_computation(sample_png_pair):
    a, b = sample_png_pair
    actual = metrics.mean_absolute_error(a, b)
    arr_a = np.asarray(Image.open(a)).astype(np.int16)
    arr_b = np.asarray(Image.open(b)).astype(np.int16)
    expected = float(np.abs(arr_a - arr_b).mean())
    assert math.isclose(actual, expected, rel_tol=1e-6)


def test_mae_resizes_when_shapes_differ(tmp_path):
    """If the input shapes differ we expect MAE to align the rendered to the source."""
    Image.new("RGB", (300, 200), (200, 0, 0)).save(tmp_path / "big.png")
    Image.new("RGB", (150, 100), (200, 0, 0)).save(tmp_path / "small.png")
    out = metrics.mean_absolute_error(tmp_path / "big.png", tmp_path / "small.png")
    # Same colour, resized -> MAE 0
    assert out == 0.0


def test_rmse_penalises_large_errors(sample_png_pair):
    a, b = sample_png_pair
    rmse_val = metrics.root_mean_squared_error(a, b)
    mae_val = metrics.mean_absolute_error(a, b)
    # RMSE >= MAE is a textbook invariant.
    assert rmse_val >= mae_val - 1e-6
    assert rmse_val > 0.0


# ---------------------------------------------------------------------------
# SSIM
# ---------------------------------------------------------------------------


def test_ssim_one_for_identical_images(identical_png_pair):
    a, b = identical_png_pair
    ssim_val = metrics.global_ssim(a, b)
    # SSIM should be exactly 1 on identical images (modulo float precision).
    assert ssim_val == pytest.approx(1.0, abs=1e-4)


def test_ssim_decreases_for_large_noise_diff(tmp_path):
    """Different images with heavy noise should give a clearly lower SSIM."""
    rng = np.random.default_rng(seed=12)
    a = np.full((96, 160, 3), 200, dtype=np.uint8) + rng.integers(0, 8, (96, 160, 3), np.uint8)
    b = a + rng.integers(0, 200, (96, 160, 3), np.uint8)
    pa = tmp_path / "a.png"
    pb = tmp_path / "b.png"
    Image.fromarray(a).save(pa)
    Image.fromarray(b).save(pb)
    ssim_val = metrics.global_ssim(pa, pb)
    assert 0.0 <= ssim_val <= 1.0
    # noisy identical->~1.0; random->~0.0 typically
    assert ssim_val < 0.5


def test_ssim_invariance_to_window_size_is_bounded(tmp_path):
    rng = np.random.default_rng(seed=99)
    arr = rng.integers(0, 255, (128, 160, 3), np.uint8)
    pa = tmp_path / "a.png"
    pb = tmp_path / "b.png"
    Image.fromarray(arr).save(pa)
    Image.fromarray(arr).save(pb)
    ssim_w11 = metrics.global_ssim(pa, pb, window_size=11)
    ssim_w7 = metrics.global_ssim(pa, pb, window_size=7)
    # Should both be ~1.0 for identical images; let's just assert the
    # symmetric window picks produce near-1 values.
    assert ssim_w11 == pytest.approx(1.0, abs=1e-3)
    assert ssim_w7 == pytest.approx(1.0, abs=1e-3)


# ---------------------------------------------------------------------------
# pixel_diff_ratio and compute_metrics
# ---------------------------------------------------------------------------


def test_pixel_diff_ratio_lower_for_smaller_diff(identical_png_pair, sample_png_pair):
    ident_a, ident_b = identical_png_pair
    sample_a, sample_b = sample_png_pair
    ident_ratio = metrics.pixel_diff_ratio(ident_a, ident_b)
    sample_ratio = metrics.pixel_diff_ratio(sample_a, sample_b)
    assert ident_ratio == 0.0
    assert sample_ratio > ident_ratio


def test_pixel_diff_ratio_threshold_controls_share(tmp_path):
    # Build image pair with difference = 50 in the green channel only.
    a = np.zeros((50, 50, 3), dtype=np.uint8)
    b = a.copy()
    b[..., 1] = 50
    pa = tmp_path / "a.png"
    pb = tmp_path / "b.png"
    Image.fromarray(a).save(pa)
    Image.fromarray(b).save(pb)
    ratio_low = metrics.pixel_diff_ratio(pa, pb, threshold=20)
    ratio_high = metrics.pixel_diff_ratio(pa, pb, threshold=60)
    # Below threshold -> many changed pixels; above -> none.
    assert ratio_low > 0.5
    assert ratio_high == 0.0


def test_compute_metrics_returns_dataclass(sample_png_pair):
    a, b = sample_png_pair
    result = metrics.compute_metrics(a, b)
    assert isinstance(result, DiffMetrics)
    assert 0.0 <= result.ssim <= 1.0
    assert 0.0 <= result.pixel_diff_ratio <= 1.0
    assert result.mae >= 0.0
    assert result.rmse >= result.mae - 1e-6
    assert result.width > 0
    assert result.height > 0
    assert result.threshold == 30


def test_compute_metrics_resizes_rendered(sample_png_pair, tmp_path):
    a, _ = sample_png_pair
    # Make a tiny copy of the source
    img = Image.open(a).resize((32, 32))
    tiny = tmp_path / "tiny.png"
    img.save(tiny)
    result = metrics.compute_metrics(a, tiny)
    # No exception means resize worked; metrics must be finite.
    assert math.isfinite(result.mae)
    assert math.isfinite(result.rmse)


# ---------------------------------------------------------------------------
# Misc
# ---------------------------------------------------------------------------


def test_coerce_rgb_array_accepts_path(tmp_path):
    Path(tmp_path / "x.png")  # ensure file exists
    image = Image.new("RGB", (10, 10), (10, 200, 30))
    image.save(tmp_path / "x.png")
    arr = metrics._coerce_rgb_array(tmp_path / "x.png")
    assert arr.shape == (10, 10, 3)


def test_metrics_works_with_pil_image_objects(sample_png_pair):
    a, b = sample_png_pair
    img_a = Image.open(a)
    img_b = Image.open(b)
    result = metrics.compute_metrics(img_a, img_b)
    assert isinstance(result.mae, float)


def test_metrics_wrong_input_type_raises():
    with pytest.raises(TypeError):
        metrics._coerce_rgb_array(123)  # type: ignore[arg-type]
