"""Tests for the HTML report builder."""

from __future__ import annotations

from pathlib import Path

import pytest
from PIL import Image

from slide2pptx.report import html_builder
from slide2pptx.report.models import (
    DiffMetrics,
    EditabilityStatus,
    EditabilitySummary,
    ReportInputs,
)


def _inputs(
    source_path: Path,
    rendered_path: Path,
    heatmap_path: Path,
    *,
    editability=None,
    diff_metrics=None,
    job_id="job_test",
) -> ReportInputs:
    return ReportInputs(
        job_id=job_id,
        source_image=source_path,
        rendered_image=rendered_path,
        heatmap=heatmap_path,
        diff_metrics=diff_metrics or DiffMetrics(
            mae=1.23, rmse=4.56, ssim=0.91, pixel_diff_ratio=0.02,
            width=200, height=120, threshold=30,
        ),
        editability=editability or EditabilitySummary(
            total_elements=3,
            editable_count=2,
            bitmap_fallback_count=0,
            image_render_count=1,
            native_render_count=2,
            low_confidence_count=0,
            total_text_chars=10,
            avg_editable_score=0.85,
            status=EditabilityStatus.FULLY_EDITABLE,
        ),
        pptx_path=Path("dummy.pptx"),
        detected_json_path=Path("detected.json"),
    )


# ---------------------------------------------------------------------------
# Basic rendering
# ---------------------------------------------------------------------------


def test_render_report_returns_html_string(identical_png_pair, tmp_path):
    a, b = identical_png_pair
    heatmap = tmp_path / "heat.png"
    Image.new("RGB", (120, 80), (50, 60, 70)).save(heatmap)
    inputs = _inputs(a, b, heatmap)
    html = html_builder.render_report(inputs)
    assert isinstance(html, str)
    assert "<html" in html.lower()
    assert "Slide2PPTX" in html


def test_render_report_inlines_images_as_data_uris(identical_png_pair, tmp_path):
    a, b = identical_png_pair
    heatmap = tmp_path / "heat.png"
    Image.new("RGB", (120, 80), (50, 60, 70)).save(heatmap)
    inputs = _inputs(a, b, heatmap)
    html = html_builder.render_report(inputs)
    # All three image panels must appear as base64 data URIs.
    assert html.count("data:image/png;base64,") >= 3


def test_render_report_writes_file(identical_png_pair, tmp_path):
    a, b = identical_png_pair
    heatmap = tmp_path / "heat.png"
    Image.new("RGB", (120, 80), (50, 60, 70)).save(heatmap)
    inputs = _inputs(a, b, heatmap)
    out = tmp_path / "report.html"
    html_builder.render_report(inputs, out_path=out)
    assert out.is_file()
    content = out.read_text(encoding="utf-8")
    assert "<html" in content.lower()
    assert "data:image/png;base64," in content


def test_render_report_missing_heatmap_does_not_crash(tmp_path):
    """Missing images become empty data URIs (page renders with alt text)."""
    a = tmp_path / "a.png"
    b = tmp_path / "b.png"
    Image.new("RGB", (50, 50), (10, 10, 10)).save(a)
    Image.new("RGB", (50, 50), (10, 10, 10)).save(b)
    inputs = _inputs(a, b, Path("does-not-exist.png"))
    html = html_builder.render_report(inputs)
    assert "data:" in html  # at least something inline-able is present


# ---------------------------------------------------------------------------
# Element table
# ---------------------------------------------------------------------------


def test_render_report_includes_elements_table(identical_png_pair, tmp_path, minimal_detected_json):
    a, b = identical_png_pair
    heatmap = tmp_path / "heat.png"
    Image.new("RGB", (40, 30), (50, 50, 50)).save(heatmap)
    inputs = _inputs(a, b, heatmap)
    elements = [
        {"id": "el_0001", "kind": "text", "render_strategy": "native",
         "editable_score": 0.85, "bbox": {"left": 0, "top": 0, "width": 100, "height": 20}},
        {"id": "el_0002", "kind": "shape", "render_strategy": "native",
         "editable_score": 0.5, "bbox": {"left": 0, "top": 30, "width": 100, "height": 20}},
    ]
    html = html_builder.render_report(inputs, elements)
    assert "el_0001" in html
    assert "el_0002" in html
    # Pill classes are rendered for native/image/background.
    assert 'pill native' in html


def test_render_report_truncates_long_text(identical_png_pair, tmp_path):
    a, b = identical_png_pair
    heatmap = tmp_path / "heat.png"
    Image.new("RGB", (10, 10), (0, 0, 0)).save(heatmap)
    inputs = _inputs(a, b, heatmap)
    long_text = "x" * 800
    elements = [{"id": "el_long", "kind": "text",
                 "render_strategy": "native",
                 "editable_score": 0.7,
                 "bbox": {"left": 0, "top": 0, "width": 100, "height": 20},
                 "text": long_text}]
    html = html_builder.render_report(inputs, elements)
    assert "…" in html  # the long text gets truncated in the table


def test_render_report_handles_no_elements(identical_png_pair, tmp_path):
    a, b = identical_png_pair
    heatmap = tmp_path / "heat.png"
    Image.new("RGB", (10, 10), (0, 0, 0)).save(heatmap)
    inputs = _inputs(a, b, heatmap)
    html = html_builder.render_report(inputs, detected_elements=[])
    assert "Elements (0)" in html


# ---------------------------------------------------------------------------
# Status banner
# ---------------------------------------------------------------------------


def test_status_banner_classes():
    assert html_builder._status_banner("fully_editable")[0] == "full"
    assert html_builder._status_banner(EditabilityStatus.BITMAP_FALLBACK)[0] == "fallback"


def test_ssim_class_thresholds():
    assert html_builder._ssim_class(0.9) == "good"
    assert html_builder._ssim_class(0.7) == "warn"
    assert html_builder._ssim_class(0.4) == "bad"


def test_diff_class_thresholds():
    assert html_builder._diff_class(0.02) == "good"
    assert html_builder._diff_class(0.10) == "warn"
    assert html_builder._diff_class(0.20) == "bad"


# ---------------------------------------------------------------------------
# Data URI helper
# ---------------------------------------------------------------------------


def test_data_uri_missing_file_returns_empty():
    assert html_builder._data_uri(Path("/nonexistent/file.png")) == ""


def test_image_dimensions_fallback_on_missing(tmp_path):
    fake = tmp_path / "missing.png"
    assert html_builder._image_dimensions(fake) == "?"
