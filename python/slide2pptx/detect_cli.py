"""Command-line entry point for slide2pptx detection.

Usage::

    python detect_cli.py INPUT_IMAGE --out OUTPUT_DIR

The CLI always writes ``detected.json`` and a background PNG
(``original-background.png`` or ``cleaned-background.png`` depending on
which optional dependencies are installed). Optional dependencies:

* ``rapidocr_onnxruntime`` -- OCR for native text elements.
* ``opencv-python`` (``cv2``) -- inpainting for the cleaned background.

``--self-test`` creates a temporary synthetic image, runs the detector
on it and verifies the JSON output without writing to the chosen
``--out`` directory.
"""

from __future__ import annotations

import argparse
import json
import sys
from pathlib import Path

# Make ``slide2pptx`` importable when the CLI is invoked from any cwd.
_THIS_DIR = Path(__file__).resolve().parent
if str(_THIS_DIR.parent) not in sys.path:
    sys.path.insert(0, str(_THIS_DIR.parent))

from slide2pptx.detect import detect, run_self_test  # noqa: E402


def _parse_args(argv: list[str] | None = None) -> argparse.Namespace:
    parser = argparse.ArgumentParser(
        prog="slide2pptx-detect",
        description=(
            "Detect slide elements from an image and write detected.json "
            "plus a background PNG to the output directory."
        ),
    )
    parser.add_argument(
        "input_image",
        nargs="?",
        type=Path,
        help="Path to the input slide image (PNG/JPG/etc.).",
    )
    parser.add_argument(
        "--out",
        dest="out_dir",
        type=Path,
        default=Path("out"),
        help="Directory to write detected.json and the background PNG to.",
    )
    parser.add_argument(
        "--self-test",
        action="store_true",
        help=(
            "Run the built-in self-test on a synthetic image and exit. "
            "Does not require INPUT_IMAGE or --out."
        ),
    )
    parser.add_argument(
        "--visual-passes",
        type=int,
        default=2,
        choices=[1, 2],
        help="Number of visual extraction passes to run (default: 2).",
    )
    parser.add_argument(
        "--second-pass-max-components",
        type=int,
        default=96,
        help="Maximum residual visual components to export in pass 2.",
    )
    return parser.parse_args(argv)


def main(argv: list[str] | None = None) -> int:
    args = _parse_args(argv)

    if args.self_test:
        payload = run_self_test()
        print(
            json.dumps(
                {
                    "ok": True,
                    "element_count": len(payload["elements"]),
                    "version": payload["version"],
                    "warnings": payload.get("warnings", []),
                },
                indent=2,
            )
        )
        return 0

    if args.input_image is None:
        print(
            "ERROR: INPUT_IMAGE is required unless --self-test is supplied.",
            file=sys.stderr,
        )
        return 2

    result = detect(
        args.input_image,
        args.out_dir,
        visual_passes=args.visual_passes,
        second_pass_max_components=args.second_pass_max_components,
    )

    print(
        json.dumps(
            {
                "ok": True,
                "detected_json": str(result.detections_json_path.resolve()),
                "background": str(result.background_path.resolve()),
                "element_count": len(result.payload["elements"]),
                "warnings": result.warnings,
            },
            indent=2,
        )
    )
    return 0


if __name__ == "__main__":
    raise SystemExit(main())
