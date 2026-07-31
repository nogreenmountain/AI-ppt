"""Dataclasses describing the inputs and outputs of the report pipeline.

Keep these small, immutable where possible. They are the public contract
for the report module: every helper either takes one of these as input or
produces one as output. ``report_cli`` and the HTML builder rely on them.
"""

from __future__ import annotations

from dataclasses import dataclass, field
from enum import Enum
from pathlib import Path
from typing import Any, Dict, List, Optional, Sequence


# ---------------------------------------------------------------------------
# Element statistics
# ---------------------------------------------------------------------------


class EditabilityStatus(str, Enum):
    """Coarse editability verdict derived from the detection JSON.

    The enum lives here so both ``checklist`` and ``html_builder`` can use
    the same set of labels without re-defining them.
    """

    FULLY_EDITABLE = "fully_editable"
    BITMAP_FALLBACK = "bitmap_fallback"
    LOW_CONFIDENCE = "low_confidence"
    UNKNOWN = "unknown"


@dataclass(frozen=True)
class ElementStats:
    """Aggregated counts for a single ``render_strategy`` bucket.

    Attributes:
        strategy: Value of ``render_strategy`` from
            ``spec/detected.schema.json`` (``"native"``,
            ``"image"``, ``"background"``).
        kind: Element kind (``"text"``, ``"shape"``, ``"image"``).
        count: How many elements of this (strategy, kind) were detected.
        avg_editable_score: Mean ``editable_score`` over the bucket.
        min_editable_score: Lowest ``editable_score`` seen.
        examples: Up to three sample element ``id`` values for tooltip use.
    """

    strategy: str
    kind: str
    count: int
    avg_editable_score: float
    min_editable_score: float
    examples: List[str] = field(default_factory=list)


@dataclass(frozen=True)
class EditabilitySummary:
    """Per-element editability roll-up produced by :mod:`checklist`."""

    total_elements: int
    editable_count: int
    bitmap_fallback_count: int
    image_render_count: int
    native_render_count: int
    low_confidence_count: int
    total_text_chars: int
    avg_editable_score: float
    per_kind: Dict[str, ElementStats] = field(default_factory=dict)
    per_strategy: Dict[str, ElementStats] = field(default_factory=dict)
    warnings: List[str] = field(default_factory=list)
    status: EditabilityStatus = EditabilityStatus.UNKNOWN

    @property
    def editable_ratio(self) -> float:
        """Return the share of elements that are individually editable.

        A native text or shape rendered with a high editable score counts
        as editable; bitmap fallbacks and low-confidence renderings do not.
        """
        if self.total_elements == 0:
            return 0.0
        return self.editable_count / self.total_elements


# ---------------------------------------------------------------------------
# Visual metrics
# ---------------------------------------------------------------------------


@dataclass(frozen=True)
class DiffMetrics:
    """Pixel-level metrics between source and rendered slide images.

    All values are floats in ``[0, 1]`` (except ``rmse`` which is on the
    same 0-255 scale as the source PNGs' channels, but reported as mean
    across the three channels). See :func:`slide2pptx.report.metrics.compute_metrics`.
    """

    mae: float
    rmse: float
    ssim: float
    pixel_diff_ratio: float
    width: int
    height: int
    threshold: int = 30


# ---------------------------------------------------------------------------
# HTML report input
# ---------------------------------------------------------------------------


@dataclass(frozen=True)
class ReportInputs:
    """All artefacts that the HTML report will reference.

    ``source_image``, ``rendered_image`` and ``heatmap`` are filesystem
    paths; the builder reads them, base64-encodes them, and inlines them.
    """

    job_id: str
    source_image: Path
    rendered_image: Path
    heatmap: Path
    diff_metrics: DiffMetrics
    editability: EditabilitySummary
    pptx_path: Optional[Path] = None
    detected_json_path: Optional[Path] = None
    warnings: List[str] = field(default_factory=list)
    timings_ms: Dict[str, float] = field(default_factory=dict)
    extra: Dict[str, Any] = field(default_factory=dict)


# ---------------------------------------------------------------------------
# Renderer return
# ---------------------------------------------------------------------------


@dataclass(frozen=True)
class RenderResult:
    """Outcome of a single PowerPoint COM export."""

    pptx_path: Path
    rendered_image: Path
    duration_ms: float
    slide_index: int = 1
    warnings: Sequence[str] = field(default_factory=tuple)


__all__ = [
    "DiffMetrics",
    "EditabilityStatus",
    "EditabilitySummary",
    "ElementStats",
    "ElementStats",
    "RenderResult",
    "ReportInputs",
]
