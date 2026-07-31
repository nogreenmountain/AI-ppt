"""Editability statistics derived from ``detected.json``.

This module operates on the JSON contract documented in
``spec/detected.schema.json``. The schema is intentionally permissive
(``additionalProperties: true``); we therefore validate the JSON
programmatically by walking the known fields rather than via
``jsonschema`` so the report keeps working even when downstream test
fixtures add extra keys.

The output is :class:`slide2pptx.report.models.EditabilitySummary`,
which the HTML builder consumes.
"""

from __future__ import annotations

import json
from collections import defaultdict
from pathlib import Path
from typing import Any, Dict, Iterable, List, Mapping, Optional, Sequence, Tuple, Union

from slide2pptx.report.models import (
    EditabilityStatus,
    EditabilitySummary,
    ElementStats,
)

JSONLike = Union[str, Path, Mapping[str, Any]]


# ---------------------------------------------------------------------------
# Validation helpers
# ---------------------------------------------------------------------------


_REQUIRED_TOP_LEVEL = ("version", "source", "slide", "background", "elements")


class DetectedJsonError(ValueError):
    """Raised when a detected JSON file violates the schema contract."""


def _load_json(source: JSONLike) -> Mapping[str, Any]:
    if isinstance(source, (str, Path)):
        text = Path(source).read_text(encoding="utf-8")
        try:
            data = json.loads(text)
        except json.JSONDecodeError as exc:
            raise DetectedJsonError(f"Invalid JSON in {source}: {exc}") from exc
    elif isinstance(source, Mapping):
        data = source
    else:
        raise TypeError(f"Cannot load JSON from {type(source).__name__}")
    if not isinstance(data, Mapping):
        raise DetectedJsonError("Top-level value must be a JSON object")
    return data


def _validate_top_level(payload: Mapping[str, Any]) -> None:
    missing = [k for k in _REQUIRED_TOP_LEVEL if k not in payload]
    if missing:
        raise DetectedJsonError(
            f"Detected JSON missing required keys: {missing!r}"
        )


# ---------------------------------------------------------------------------
# Element iteration
# ---------------------------------------------------------------------------


def _iter_elements(payload: Mapping[str, Any]) -> Iterable[Mapping[str, Any]]:
    elements = payload.get("elements")
    if elements is None:
        return ()
    if not isinstance(elements, list):
        raise DetectedJsonError("'elements' must be an array")
    for idx, element in enumerate(elements):
        if not isinstance(element, Mapping):
            raise DetectedJsonError(f"element[{idx}] is not an object")
        yield element


def _kind_of(element: Mapping[str, Any]) -> str:
    kind = element.get("kind")
    if kind in ("text", "shape", "image"):
        return kind
    return "unknown"


def _strategy_of(element: Mapping[str, Any]) -> str:
    strategy = element.get("render_strategy")
    if strategy in ("native", "image", "background"):
        return strategy
    return "unknown"


def _is_text(element: Mapping[str, Any]) -> bool:
    return _kind_of(element) == "text"


# ---------------------------------------------------------------------------
# Statistics computation
# ---------------------------------------------------------------------------


_DEFAULT_EDITABLE_THRESHOLD = 0.4


def _build_bucket(
    rows: Sequence[Tuple[Dict[str, Any], str]],
) -> ElementStats:
    """Combine a bucket of (element, id) tuples into ElementStats."""
    rows_list = list(rows)
    if not rows_list:
        return ElementStats(strategy="", kind="", count=0,
                            avg_editable_score=0.0, min_editable_score=0.0)
    scores = [float(r[0].get("editable_score", 0.0)) for r in rows_list]
    examples = [r[1] for r in rows_list[:3]]
    return ElementStats(
        strategy=rows_list[0][0].get("__strategy__", ""),
        kind=rows_list[0][0].get("__kind__", ""),
        count=len(rows_list),
        avg_editable_score=sum(scores) / len(scores),
        min_editable_score=min(scores),
        examples=examples,
    )


def compute_editability(
    detected: JSONLike,
    *,
    editable_threshold: float = _DEFAULT_EDITABLE_THRESHOLD,
    low_confidence_threshold: float = 0.75,
) -> EditabilitySummary:
    """Compute the editability summary for a detected-JSON.

    The function is tolerant: unknown ``kind`` or ``render_strategy``
    values count as fallback elements, mirroring the prototype
    blueprint's "grep everything" approach.
    """
    payload = _load_json(detected)
    _validate_top_level(payload)
    elements = list(_iter_elements(payload))

    warnings: List[str] = list(payload.get("warnings") or [])
    if not isinstance(warnings, list):
        warnings = []

    per_kind_buckets: Dict[str, List[Tuple[Dict[str, Any], str]]] = defaultdict(list)
    per_strategy_buckets: Dict[str, List[Tuple[Dict[str, Any], str]]] = defaultdict(list)

    editable_count = 0
    bitmap_count = 0
    image_render_count = 0
    native_render_count = 0
    low_confidence_count = 0
    total_text_chars = 0
    score_sum = 0.0
    score_count = 0

    for element in elements:
        kind = _kind_of(element)
        strategy = _strategy_of(element)
        if kind == "unknown":
            warnings.append(
                f"element '{element.get('id', '?')}' has unknown kind"
            )
        if strategy == "unknown":
            warnings.append(
                f"element '{element.get('id', '?')}' has unknown render_strategy"
            )
        # Annotate the element with derived fields so the bucket helper
        # has the necessary context.
        element = dict(element)
        element["__kind__"] = kind
        element["__strategy__"] = strategy

        per_kind_buckets[kind].append((element, str(element.get("id", ""))))
        per_strategy_buckets[strategy].append((element, str(element.get("id", ""))))

        score = float(element.get("editable_score", 0.0) or 0.0)
        score_sum += score
        score_count += 1

        # Strategy counts
        if strategy == "background":
            image_render_count += 1  # raster backdrop
        elif strategy == "image":
            image_render_count += 1
        elif strategy == "native":
            native_render_count += 1
            if kind in ("text", "shape"):
                if score >= editable_threshold:
                    editable_count += 1
        elif strategy == "unknown":
            bitmap_count += 1

        # Confidence categorisation (used for low_confidence warning).
        confidence = element.get("confidence")
        if isinstance(confidence, Mapping):
            text_conf = confidence.get("text") or confidence.get("ocr")
            if isinstance(text_conf, (int, float)) and text_conf < low_confidence_threshold:
                low_confidence_count += 1

        if _is_text(element):
            txt = element.get("text")
            if isinstance(txt, str):
                total_text_chars += len(txt)

    per_kind = {k: _build_bucket(rows) for k, rows in per_kind_buckets.items()}
    per_strategy = {k: _build_bucket(rows) for k, rows in per_strategy_buckets.items()}

    avg_score = score_sum / score_count if score_count else 0.0

    status = _resolve_status(
        total_elements=len(elements),
        editable_count=editable_count,
        image_render_count=image_render_count,
        bitmap_count=bitmap_count,
        avg_score=avg_score,
    )

    if not elements:
        warnings.append("detected JSON contains no elements")

    return EditabilitySummary(
        total_elements=len(elements),
        editable_count=editable_count,
        bitmap_fallback_count=bitmap_count,
        image_render_count=image_render_count,
        native_render_count=native_render_count,
        low_confidence_count=low_confidence_count,
        total_text_chars=total_text_chars,
        avg_editable_score=avg_score,
        per_kind=per_kind,
        per_strategy=per_strategy,
        warnings=warnings,
        status=status,
    )


def _resolve_status(
    *,
    total_elements: int,
    editable_count: int,
    image_render_count: int,
    bitmap_count: int,
    avg_score: float,
) -> EditabilityStatus:
    """Pick the high-level status from the stats."""
    if total_elements == 0:
        return EditabilityStatus.UNKNOWN
    if editable_count == total_elements:
        return EditabilityStatus.FULLY_EDITABLE
    if bitmap_count + image_render_count == total_elements:
        return EditabilityStatus.BITMAP_FALLBACK
    if avg_score < 0.6:
        return EditabilityStatus.LOW_CONFIDENCE
    return EditabilityStatus.LOW_CONFIDENCE  # any nonzero bitmap makes this


# ---------------------------------------------------------------------------
# Element-level ops (useful for the HTML table)
# ---------------------------------------------------------------------------


def list_elements(detected: JSONLike) -> List[Dict[str, Any]]:
    """Return a flat list of detection elements with annotated derived fields."""
    payload = _load_json(detected)
    _validate_top_level(payload)
    out: List[Dict[str, Any]] = []
    for element in _iter_elements(payload):
        annotated = dict(element)
        annotated["__kind__"] = _kind_of(element)
        annotated["__strategy__"] = _strategy_of(element)
        out.append(annotated)
    return out


def element_confidence(element: Mapping[str, Any]) -> Optional[float]:
    """Extract the most-relevant confidence from an element.

    Returns the highest confidence value found in ``element.confidence``
    (a free-form object in the schema). Returns ``None`` if there are no
    numeric confidences.
    """
    confidence = element.get("confidence")
    if not isinstance(confidence, Mapping):
        return None
    values = [v for v in confidence.values() if isinstance(v, (int, float))]
    return max(values) if values else None


__all__ = [
    "DetectedJsonError",
    "compute_editability",
    "list_elements",
    "element_confidence",
]
