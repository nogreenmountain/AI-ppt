"""Shared pytest fixtures for the report test suite.

The tests do NOT require PowerPoint to be installed: the renderer is
mocked and the diff metrics operate on synthetic PNGs that we generate
in-process. We *do* honour a ``RUN_RENDERER_INTEGRATION`` environment
variable (off by default) to allow devs who have PowerPoint + Office to
run an end-to-end smoke test manually.
"""

from __future__ import annotations

import json
import os
import uuid
from pathlib import Path
from typing import Dict, List, Tuple

import numpy as np
import pytest


def pytest_addoption(parser):
    parser.addoption(
        "--run-integration",
        action="store_true",
        default=False,
        help="Run tests that open Microsoft PowerPoint through COM.",
    )
from PIL import Image, ImageDraw, ImageFont


# ---------------------------------------------------------------------------
# Image factories
# ---------------------------------------------------------------------------


def _make_image(size: Tuple[int, int], *, fill: Tuple[int, int, int] = (240, 240, 240)) -> Image.Image:
    """Create a solid-colour RGB PIL image."""
    return Image.new("RGB", size, fill)


def _make_text_image(
    size: Tuple[int, int],
    text: str,
    *,
    fill: Tuple[int, int, int] = (255, 255, 255),
    text_fill: Tuple[int, int, int] = (12, 12, 12),
    size_pt: int = 32,
) -> Image.Image:
    """Create a simple PIL image that has rendered text on a solid background."""
    img = _make_image(size, fill=fill)
    draw = ImageDraw.Draw(img)
    try:
        font = ImageFont.truetype("arial.ttf", size_pt)
    except (OSError, IOError):
        font = ImageFont.load_default()
    draw.text((size_pt, size_pt), text, fill=text_fill, font=font)
    return img


@pytest.fixture(scope="session")
def synthetic_image_factory():
    """Expose the image factories as a session fixture."""
    return {"make_image": _make_image, "make_text_image": _make_text_image}


@pytest.fixture
def sample_png_pair(tmp_path) -> Tuple[Path, Path]:
    """Create a pair of small test PNGs.

    The two PNGs are derived from the same source with a deterministic
    noise pattern; this lets us test the metric invariants without
    needing actual PowerPoint in the loop.
    """
    rng = np.random.default_rng(seed=42)
    base = np.full((120, 200, 3), 245, dtype=np.uint8)
    base[40:80, 30:170] = (210, 24, 24)  # red banner
    image_a = base + rng.integers(0, 5, size=base.shape, dtype=np.uint8)
    image_b = base + rng.integers(0, 60, size=base.shape, dtype=np.uint8)
    a_path = tmp_path / "source.png"
    b_path = tmp_path / "rendered.png"
    Image.fromarray(image_a, mode="RGB").save(a_path)
    Image.fromarray(image_b, mode="RGB").save(b_path)
    return a_path, b_path


@pytest.fixture
def identical_png_pair(tmp_path) -> Tuple[Path, Path]:
    """Two PNGs with identical content (used for SSIM ~ 1 invariants)."""
    img = _make_image((80, 120), fill=(204, 230, 255))
    a = tmp_path / "a.png"
    b = tmp_path / "b.png"
    img.save(a)
    img.save(b)
    return a, b


# ---------------------------------------------------------------------------
# Detected-JSON fixtures
# ---------------------------------------------------------------------------


@pytest.fixture
def minimal_detected_json() -> Dict[str, object]:
    """A small but schema-valid ``detected.json``."""
    return {
        "version": "1.0",
        "source": {
            "image_path": "source.png",
            "width_px": 1920,
            "height_px": 1080,
        },
        "slide": {"width": 1920, "height": 1080, "unit": "px"},
        "background": {
            "strategy": "solid",
            "image_path": None,
            "fill": "#ffffff",
        },
        "elements": [
            {
                "id": "el_0001",
                "kind": "text",
                "bbox": {"left": 32, "top": 32, "width": 600, "height": 80},
                "z": 1,
                "editable_score": 0.95,
                "render_strategy": "native",
                "text": "Slide title",
                "font_family": "Arial",
                "font_size": 36,
                "text_color": "#101010",
                "confidence": {"text": 0.92},
            },
            {
                "id": "el_0002",
                "kind": "shape",
                "bbox": {"left": 1200, "top": 700, "width": 600, "height": 200},
                "z": 0,
                "editable_score": 0.88,
                "render_strategy": "native",
                "geometry": "roundRect",
                "fill": "#2266dd",
            },
            {
                "id": "el_0003",
                "kind": "image",
                "bbox": {"left": 800, "top": 200, "width": 400, "height": 300},
                "z": 2,
                "editable_score": 0.20,
                "render_strategy": "image",
                "image_path": "assets/el_0003.png",
            },
            {
                "id": "el_0004",
                "kind": "shape",
                "bbox": {"left": 50, "top": 950, "width": 800, "height": 80},
                "z": -1,
                "editable_score": 0.10,
                "render_strategy": "background",
                "geometry": "rect",
            },
        ],
        "warnings": ["test warning"],
    }


@pytest.fixture
def detected_json_path(tmp_path, minimal_detected_json) -> Path:
    """A schema-valid detected JSON on disk."""
    p = tmp_path / "detected.json"
    p.write_text(json.dumps(minimal_detected_json), encoding="utf-8")
    return p


@pytest.fixture
def empty_detected_json() -> Dict[str, object]:
    """A valid detected.json with zero elements."""
    return {
        "version": "1.0",
        "source": {"image_path": "x.png", "width_px": 100, "height_px": 100},
        "slide": {"width": 100, "height": 100, "unit": "px"},
        "background": {"strategy": "solid", "image_path": None, "fill": "#000000"},
        "elements": [],
        "warnings": [],
    }


# ---------------------------------------------------------------------------
# CLI testing helpers
# ---------------------------------------------------------------------------


@pytest.fixture
def integration_env():
    """Expose whether integration tests are enabled.

    Skips when ``RUN_RENDERER_INTEGRATION`` is unset/false.
    """
    return {"enabled": bool(os.environ.get("RUN_RENDERER_INTEGRATION"))}


@pytest.fixture
def job_id() -> str:
    return uuid.uuid4().hex[:8]
