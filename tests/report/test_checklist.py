"""Tests for the editability / checklist module."""

from __future__ import annotations

import json

import pytest

from slide2pptx.report import checklist
from slide2pptx.report.models import (
    EditabilityStatus,
    EditabilitySummary,
    ElementStats,
)


# ---------------------------------------------------------------------------
# Basic ingestion
# ---------------------------------------------------------------------------


def test_compute_editability_minimal(detected_json_path):
    summary = checklist.compute_editability(detected_json_path)
    assert isinstance(summary, EditabilitySummary)
    assert summary.total_elements == 4
    # Two native text/shape with score >= 0.4 should be editable.
    assert summary.editable_count == 2
    # A background strategy counts as image_render; an image strategy counts as
    # image_render; native counts as native.
    assert summary.native_render_count == 2
    assert summary.image_render_count == 2
    assert summary.bitmap_fallback_count == 0


def test_compute_editability_accepts_dict(minimal_detected_json):
    summary = checklist.compute_editability(minimal_detected_json)
    assert summary.total_elements == 4
    assert summary.avg_editable_score > 0


def test_compute_editability_dict_min_fields_raises(tmp_path):
    bad_path = tmp_path / "bad.json"
    bad_path.write_text('{"version": "1.0"}', encoding="utf-8")
    with pytest.raises(checklist.DetectedJsonError):
        checklist.compute_editability(bad_path)


def test_compute_editability_malformed_json(tmp_path):
    p = tmp_path / "broken.json"
    p.write_text("not even json {", encoding="utf-8")
    with pytest.raises(checklist.DetectedJsonError):
        checklist.compute_editability(p)


def test_compute_editability_non_object_top_level(tmp_path):
    p = tmp_path / "list.json"
    p.write_text("[]", encoding="utf-8")
    with pytest.raises(checklist.DetectedJsonError):
        checklist.compute_editability(p)


def test_compute_editability_no_elements(empty_detected_json):
    summary = checklist.compute_editability(empty_detected_json)
    assert summary.total_elements == 0
    assert summary.editable_ratio == 0
    # Empty -> unknown status
    assert summary.status == EditabilityStatus.UNKNOWN
    # The warning is added because no elements were seen.
    assert any("no elements" in w for w in summary.warnings)


# ---------------------------------------------------------------------------
# Detail checks
# ---------------------------------------------------------------------------


def test_per_kind_buckets(detected_json_path):
    summary = checklist.compute_editability(detected_json_path)
    assert "text" in summary.per_kind
    assert "shape" in summary.per_kind
    assert "image" in summary.per_kind
    # The single text element has editable_score 0.95.
    text_bucket = summary.per_kind["text"]
    assert isinstance(text_bucket, ElementStats)
    assert text_bucket.count == 1
    assert text_bucket.avg_editable_score == pytest.approx(0.95, abs=1e-6)


def test_per_strategy_buckets(detected_json_path):
    summary = checklist.compute_editability(detected_json_path)
    assert set(summary.per_strategy.keys()) == {"native", "image", "background"}


def test_warnings_passthrough(minimal_detected_json):
    summary = checklist.compute_editability(minimal_detected_json)
    assert "test warning" in summary.warnings


def test_low_confidence_classification(minimal_detected_json):
    # Force the text element into low-confidence by lowering both editable_score
    # and the confidence payload.
    minimal_detected_json["elements"][0]["editable_score"] = 0.10
    minimal_detected_json["elements"][0]["confidence"] = {"text": 0.40}
    summary = checklist.compute_editability(minimal_detected_json)
    assert summary.low_confidence_count == 1


def test_element_confidence_helper(minimal_detected_json):
    element = minimal_detected_json["elements"][0]
    confidence = checklist.element_confidence(element)
    assert confidence == pytest.approx(0.92, abs=1e-6)


def test_element_confidence_no_mapping_returns_none():
    assert checklist.element_confidence({"id": "el_x"}) is None
    assert checklist.element_confidence({"confidence": None}) is None


# ---------------------------------------------------------------------------
# list_elements
# ---------------------------------------------------------------------------


def test_list_elements_returns_annotated_rows(detected_json_path):
    rows = checklist.list_elements(detected_json_path)
    assert len(rows) == 4
    for row in rows:
        assert "__kind__" in row
        assert "__strategy__" in row


def test_list_elements_with_invalid_shape(tmp_path):
    p = tmp_path / "detected.json"
    p.write_text('{"version":"1.0","source":{"image_path":"x","width_px":10,"height_px":10},"slide":{"width":10,"height":10},"background":{"strategy":"solid","image_path":null},"elements":[{"id":"el_0001","kind":"text","bbox":{"left":0,"top":0,"width":10,"height":10},"z":0,"editable_score":0.5,"render_strategy":"native"}]}', encoding="utf-8")
    rows = checklist.list_elements(p)
    assert rows[0]["__kind__"] == "text"
    assert rows[0]["__strategy__"] == "native"


# ---------------------------------------------------------------------------
# Editability status decision
# ---------------------------------------------------------------------------


def test_status_fully_editable():
    status = checklist._resolve_status(
        total_elements=3, editable_count=3,
        image_render_count=0, bitmap_count=0,
        avg_score=0.9,
    )
    assert status == EditabilityStatus.FULLY_EDITABLE


def test_status_bitmap_fallback():
    status = checklist._resolve_status(
        total_elements=2, editable_count=0,
        image_render_count=2, bitmap_count=0,
        avg_score=0.1,
    )
    assert status == EditabilityStatus.BITMAP_FALLBACK


def test_status_low_confidence_for_mixed():
    status = checklist._resolve_status(
        total_elements=4, editable_count=2,
        image_render_count=1, bitmap_count=1,
        avg_score=0.45,
    )
    assert status in {EditabilityStatus.LOW_CONFIDENCE, EditabilityStatus.BITMAP_FALLBACK}


def test_editable_ratio_property(empty_detected_json):
    summary = checklist.compute_editability(empty_detected_json)
    assert summary.editable_ratio == 0
