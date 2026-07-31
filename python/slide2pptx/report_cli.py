"""Command-line entry point for the slide2pptx report module.

Usage
-----
PowerShell (Windows)::

    python -m slide2pptx.report_cli render --pptx slide.pptx --out out_dir
    python -m slide2pptx.report_cli diff --source source.png --rendered rendered.png --out out_dir
    python -m slide2pptx.report_cli full --source source.png --pptx slide.pptx --detected detected.json --out out_dir

Bash / POSIX::

    python -m slide2pptx.report_cli --help

Each subcommand is intentionally narrow so callers can compose them
with shell pipes. ``full`` is the convenience command that mirrors the
end-to-end workflow described in ``research/prototype-blueprint.md``.
"""

from __future__ import annotations

import argparse
import json
import logging
import sys
import time
from pathlib import Path
from typing import Optional

from slide2pptx.report import (
    checklist,
    diff,
    html_builder,
    metrics,
    models,
    renderer,
)

LOG = logging.getLogger("slide2pptx.report_cli")

# Exit codes loosely aligned with the prototype blueprint §5.3.
EXIT_OK = 0
EXIT_INPUT = 10
EXIT_RENDER = 40
EXIT_OTHER = 99


# ---------------------------------------------------------------------------
# Argument parser
# ---------------------------------------------------------------------------


def _parser() -> argparse.ArgumentParser:
    """Construct the top-level argument parser (used by ``main``)."""
    parser = argparse.ArgumentParser(
        prog="python -m slide2pptx.report_cli",
        description="Slide2PPTX report utilities (renderer, diff, editability, HTML).",
    )
    parser.add_argument("--log-level", default="INFO",
                        help="Python logging level (default: INFO)")
    sub = parser.add_subparsers(dest="command", required=True)

    render_cmd = sub.add_parser("render", help="Export the first slide of a PPTX as PNG.")
    render_cmd.add_argument("--pptx", required=True, type=Path,
                            help="Input .pptx file")
    render_cmd.add_argument("--out", required=True, type=Path,
                            help="Output PNG path")
    render_cmd.add_argument("--slide-index", type=int, default=1,
                            help="1-based slide number (default: 1)")
    render_cmd.add_argument("--timeout", type=float, default=90.0,
                            help="Subprocess timeout in seconds (default: 90)")

    diff_cmd = sub.add_parser("diff", help="Compute pixel-level metrics between two PNGs.")
    diff_cmd.add_argument("--source", required=True, type=Path, help="Source PNG")
    diff_cmd.add_argument("--rendered", required=True, type=Path, help="Rendered PNG")
    diff_cmd.add_argument("--out", required=True, type=Path,
                          help="Output directory; writes metrics.json and diff.png")
    diff_cmd.add_argument("--threshold", type=int, default=30,
                          help="Per-channel threshold (0-255) for diff ratio")

    report_cmd = sub.add_parser("full", help="Run the entire report chain on a job.")
    report_cmd.add_argument("--source", required=True, type=Path)
    report_cmd.add_argument("--rendered", required=True, type=Path)
    report_cmd.add_argument("--out", required=True, type=Path)
    report_cmd.add_argument("--pptx", type=Path, default=None,
                            help="Optional: PPTX path recorded in the report footer")
    report_cmd.add_argument("--detected", type=Path, default=None,
                            help="Optional: detected.json path for editability stats")
    report_cmd.add_argument("--threshold", type=int, default=30)

    return parser


# ---------------------------------------------------------------------------
# Subcommand implementations
# ---------------------------------------------------------------------------


def _cmd_render(args: argparse.Namespace) -> int:
    """``render`` subcommand implementation."""
    pptx = Path(args.pptx)
    out = Path(args.out)
    if not pptx.is_file():
        LOG.error("PPTX not found: %s", pptx)
        return EXIT_INPUT

    options = models.RenderOptions if hasattr(models, "RenderOptions") else renderer.RenderOptions
    opt = options(timeout_seconds=args.timeout, slide_index=args.slide_index)
    ren = renderer.SlideRenderer(options=opt)

    try:
        result = ren.render(pptx, out, slide_index=args.slide_index)
    except renderer.RendererError as exc:
        LOG.error("Render failed: %s", exc)
        return EXIT_RENDER
    except FileNotFoundError as exc:
        LOG.error("%s", exc)
        return EXIT_INPUT

    LOG.info("Rendered %s -> %s in %.0f ms",
             result.pptx_path, result.rendered_image, result.duration_ms)
    return EXIT_OK


def _cmd_diff(args: argparse.Namespace) -> int:
    """``diff`` subcommand implementation."""
    src = Path(args.source)
    rendered = Path(args.rendered)
    out_dir = Path(args.out)
    if not src.is_file() or not rendered.is_file():
        LOG.error("Source and rendered PNGs must exist: %s, %s", src, rendered)
        return EXIT_INPUT

    out_dir.mkdir(parents=True, exist_ok=True)

    metric_values = metrics.compute_metrics(src, rendered, threshold=args.threshold)
    heatmap_path = out_dir / "diff.png"
    overlay_path = out_dir / "diff_overlay.png"
    diff.render_heatmap(src, rendered, heatmap_path)
    diff.overlay_heatmap(src, rendered, overlay_path)

    payload = {
        "source": str(src),
        "rendered": str(rendered),
        "threshold": metric_values.threshold,
        "width": metric_values.width,
        "height": metric_values.height,
        "mae": metric_values.mae,
        "rmse": metric_values.rmse,
        "ssim": metric_values.ssim,
        "pixel_diff_ratio": metric_values.pixel_diff_ratio,
        "heatmap": str(heatmap_path),
        "overlay": str(overlay_path),
    }
    metrics_path = out_dir / "metrics.json"
    metrics_path.write_text(json.dumps(payload, indent=2), encoding="utf-8")
    LOG.info("Wrote %s and %s", metrics_path, heatmap_path)
    return EXIT_OK


def _cmd_full(args: argparse.Namespace) -> int:
    """``full`` subcommand implementation."""
    src = Path(args.source)
    rendered = Path(args.rendered)
    out_dir = Path(args.out)
    if not src.is_file() or not rendered.is_file():
        LOG.error("Source and rendered PNGs must exist: %s, %s", src, rendered)
        return EXIT_INPUT

    out_dir.mkdir(parents=True, exist_ok=True)
    timings: dict = {}

    started = time.monotonic()
    metric_values = metrics.compute_metrics(src, rendered, threshold=args.threshold)
    timings["metrics.compute_metrics_ms"] = (time.monotonic() - started) * 1000

    started = time.monotonic()
    heatmap_path = out_dir / "diff.png"
    diff.render_heatmap(src, rendered, heatmap_path)
    timings["diff.render_heatmap_ms"] = (time.monotonic() - started) * 1000

    elements: list = []
    editability_summary: Optional[models.EditabilitySummary] = None
    if args.detected:
        det_path = Path(args.detected)
        if det_path.is_file():
            started = time.monotonic()
            editability_summary = checklist.compute_editability(det_path)
            elements = checklist.list_elements(det_path)
            timings["checklist.compute_ms"] = (time.monotonic() - started) * 1000
        else:
            LOG.warning("Detected JSON not found: %s", det_path)

    if editability_summary is None:
        editability_summary = models.EditabilitySummary(
            total_elements=len(elements),
            editable_count=0,
            bitmap_fallback_count=0,
            image_render_count=0,
            native_render_count=0,
            low_confidence_count=0,
            total_text_chars=0,
            avg_editable_score=0.0,
            status=models.EditabilityStatus.UNKNOWN,
        )

    started = time.monotonic()
    report_inputs = models.ReportInputs(
        job_id=out_dir.name or "report",
        source_image=src,
        rendered_image=rendered,
        heatmap=heatmap_path,
        diff_metrics=metric_values,
        editability=editability_summary,
        pptx_path=args.pptx,
        detected_json_path=args.detected,
        warnings=[],
        timings_ms=timings,
    )
    html_path = out_dir / "report.html"
    html_builder.render_report(report_inputs, elements, out_path=html_path)
    timings["html_builder.render_ms"] = (time.monotonic() - started) * 1000

    LOG.info("Report written to %s", html_path)
    return EXIT_OK


# ---------------------------------------------------------------------------
# Module entry point
# ---------------------------------------------------------------------------


def main(argv: Optional[list] = None) -> int:
    """Python ``__main__`` entry point."""
    parser = _parser()
    args = parser.parse_args(argv)
    logging.basicConfig(
        level=getattr(logging, args.log_level.upper(), logging.INFO),
        format="%(asctime)s %(levelname)s %(name)s | %(message)s",
    )
    try:
        if args.command == "render":
            return _cmd_render(args)
        if args.command == "diff":
            return _cmd_diff(args)
        if args.command == "full":
            return _cmd_full(args)
        parser.error("Unknown command")
        return EXIT_OTHER
    except Exception as exc:  # pragma: no cover - defensive
        LOG.exception("Unhandled error: %s", exc)
        return EXIT_OTHER


if __name__ == "__main__":
    sys.exit(main())
