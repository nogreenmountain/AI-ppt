"""Evaluation and reporting sub-package.

This module owns:

* :mod:`slide2pptx.report.renderer` - PowerPoint COM renderer (Windows only).
* :mod:`slide2pptx.report.metrics` - pixel-level metrics (MAE, RMSE, SSIM).
* :mod:`slide2pptx.report.diff` - per-pixel diff heatmap.
* :mod:`slide2pptx.report.checklist` - editability statistics from
  ``detected.json``.
* :mod:`slide2pptx.report.html_builder` - self-contained HTML report.
* :mod:`slide2pptx.report.models` - dataclasses describing the pipeline
  artefacts.
"""

from importlib import import_module
from types import ModuleType

__all__ = [
    "checklist",
    "diff",
    "html_builder",
    "metrics",
    "models",
    "renderer",
]


def __getattr__(name: str) -> ModuleType:
    if name in __all__:
        module = import_module(f"{__name__}.{name}")
        globals()[name] = module
        return module
    raise AttributeError(f"module {__name__!r} has no attribute {name!r}")
