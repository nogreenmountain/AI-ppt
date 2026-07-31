"""Detection sub-package for slide2pptx.

This sub-package turns a slide screenshot into a structured
``detected.json`` description (see ``spec/detected.schema.json``) plus a
background PNG. Pillow + NumPy are required; ``rapidocr_onnxruntime``
and OpenCV are optional -- their absence is logged in
``payload["warnings"]`` and the output still gets produced.
"""

from .core import DetectResult, detect, run_self_test

__all__ = ["detect", "DetectResult", "run_self_test"]