"""Compact end-to-end orchestration CLI for slide2pptx.

Usage::

    python -m slide2pptx.pipeline_cli INPUT_FILE --out OUTPUT_DIR [--skip-report]

The orchestrator runs the existing pieces of the project in order:

1. :func:`slide2pptx.detect.detect` runs on an input image and writes
   ``detected.json`` + a background PNG into ``OUTPUT_DIR/detect``.
2. ``node artifact-runtime/src/convert.mjs`` re-renders the detected JSON
   into ``reconstructed.pptx`` plus ``artifact-preview.png`` inside
   ``OUTPUT_DIR/build``.
3. Unless ``--skip-report`` is given, :class:`SlideRenderer` exports the
   reconstructed PPTX to ``powerPoint-render.png`` and the existing report
   modules compute MAE/RMSE/SSIM, a heatmap and a self-contained HTML
   report under ``OUTPUT_DIR/report``.

The CLI deliberately does not re-implement detection, builder or report
logic - it only glues those modules together. All subprocess calls use
list-form arguments; no shell is spawned. Node discovery follows a fixed
    order: ``SLIDE2PPTX_NODE`` environment override, then
    ``shutil.which('node')``.

Exit codes:
    0   success
    10  input file missing or invalid
    20  Node runtime missing
    30  convert.mjs exited non-zero
    40  PPTX render failed
    50  Report assembly failed (any other orchestration failure)
    99  unexpected internal error
"""

from __future__ import annotations

import argparse
import json
import logging
import os
import shutil
import subprocess
import sys
import time
from pathlib import Path
from typing import Any

# Make ``slide2pptx`` importable when the CLI is invoked from any cwd
# (mirrors the pattern in detect_cli.py / report_cli.py).
_THIS_FILE = Path(__file__).resolve()
_PYTHON_DIR = _THIS_FILE.parent.parent  # python/
if str(_PYTHON_DIR) not in sys.path:
    sys.path.insert(0, str(_PYTHON_DIR))

from slide2pptx.detect import detect  # noqa: E402
from slide2pptx import ppt_input  # noqa: E402


LOG = logging.getLogger("slide2pptx.pipeline_cli")

# Exit codes loosely aligned with the rest of the project.
EXIT_OK = 0
EXIT_INPUT = 10
EXIT_NODE_MISSING = 20
EXIT_BUILDER = 30
EXIT_RENDER = 40
EXIT_REPORT = 50
EXIT_OTHER = 99
PRESENTATION_SUFFIXES = {".ppt", ".pptx"}


# ---------------------------------------------------------------------------
# Node discovery & repo root
# ---------------------------------------------------------------------------


def resolve_repo_root(start: Path) -> Path:
    """Walk upward from ``start`` until a recognised repo marker is found.

    The repository ships with a ``python/`` directory directly under the
    root (this CLI lives at ``python/slide2pptx/pipeline_cli.py``). When
    that marker is present we trust it. Otherwise we fall back to the
    parent of the directory holding ``slide2pptx`` and finally to ``start``.
    """
    override = os.environ.get("SLIDE2PPTX_REPO_ROOT", "").strip()
    if override:
        candidate = Path(override)
        if (candidate / "artifact-runtime").is_dir() and (candidate / "python").is_dir():
            return candidate.resolve()
        if (candidate / "artifact-runtime").is_dir():
            return candidate.resolve()

    start = start.resolve()
    candidates: list[Path] = [start]
    for parent in start.parents:
        candidates.append(parent)
        if parent.name == "python" and (parent.parent / "artifact-runtime").is_dir():
            return parent.parent
        if (parent / "artifact-runtime").is_dir() and (parent / "python").is_dir():
            return parent
    # As a last resort, return the highest ancestor we walked up to so the
    # caller still produces a useful error message rather than a crash.
    return candidates[-1]


def find_node(runtime_root: Path | None = None) -> str:
    """Discover the Node.js executable to pass to subprocess.

    Discovery order:
        1. ``SLIDE2PPTX_NODE`` environment override (must exist or
           ``FileNotFoundError`` is raised).
        2. ``shutil.which('node')`` (whatever is on ``PATH``).
    """
    override = os.environ.get("SLIDE2PPTX_NODE", "").strip()
    if override:
        candidate = Path(override)
        if candidate.is_file():
            return str(candidate)
        raise FileNotFoundError(
            f"SLIDE2PPTX_NODE={override!r} does not point to an existing file"
        )
    on_path = shutil.which("node")
    if on_path:
        return on_path
    raise FileNotFoundError(
        "Node.js 20+ executable not found. Run convert_image_to_ppt.bat to "
        "auto-install project dependencies, set SLIDE2PPTX_NODE, or add node to PATH."
    )


# ---------------------------------------------------------------------------
# Step 1: detection
# ---------------------------------------------------------------------------


def run_detect(
    image_path: Path,
    detect_dir: Path,
    *,
    visual_passes: int = 2,
    second_pass_max_components: int = 96,
) -> dict[str, Any]:
    """Run :func:`slide2pptx.detect.detect` and return a normalised summary.

    The summary is plain JSON-serialisable and is bubbled up into the
    final result.
    """
    detect_dir.mkdir(parents=True, exist_ok=True)
    started = time.monotonic()
    result = detect(
        image_path,
        detect_dir,
        visual_passes=visual_passes,
        second_pass_max_components=second_pass_max_components,
    )
    duration_ms = (time.monotonic() - started) * 1000
    return {
        "detected_json": str(result.detections_json_path.resolve()),
        "background": str(result.background_path.resolve()),
        "element_count": len(result.payload.get("elements", [])),
        "warnings": list(result.warnings),
        "duration_ms": round(duration_ms, 2),
    }


# ---------------------------------------------------------------------------
# Step 2: artifact builder (Node)
# ---------------------------------------------------------------------------


def run_convert(
    node_exe: str,
    repo_root: Path,
    detected_json: Path,
    build_dir: Path,
    *,
    attempts: int = 3,
    retry_delay_s: float = 0.75,
) -> dict[str, Any]:
    """Invoke ``artifact-runtime/src/convert.mjs`` to produce PPTX + preview.

    Both files are written under ``build_dir`` so the rest of the
    pipeline never has to guess a name. Any non-zero exit is surfaced as
    a :class:`BuilderFailure` so the orchestrator can map it to a useful
    exit code.
    """
    build_dir.mkdir(parents=True, exist_ok=True)
    pptx_out = build_dir / "reconstructed.pptx"
    preview_out = build_dir / "artifact-preview.png"

    convert_script = repo_root / "artifact-runtime" / "src" / "convert.mjs"
    if not convert_script.is_file():
        raise FileNotFoundError(
            f"convert.mjs not found at {convert_script!s}; "
            "set SLIDE2PPTX_REPO_ROOT or run from the repo."
        )

    cmd = [
        node_exe,
        str(convert_script),
        "--spec", str(detected_json.resolve()),
        "--out", str(pptx_out.resolve()),
        "--preview", str(preview_out.resolve()),
    ]
    attempts = max(1, attempts)
    attempt_summaries: list[dict[str, Any]] = []
    last_summary: dict[str, Any] | None = None

    for attempt in range(1, attempts + 1):
        for stale_path in (pptx_out, preview_out):
            try:
                stale_path.unlink()
            except FileNotFoundError:
                pass

        LOG.info("Running convert.mjs (attempt %s/%s): %s", attempt, attempts, cmd)
        started = time.monotonic()
        proc = subprocess.run(
            cmd,
            capture_output=True,
            text=True,
            encoding="utf-8",
            errors="replace",
            check=False,
            # No shell; all arguments are explicit strings.
        )
        duration_ms = (time.monotonic() - started) * 1000

        summary = {
            "attempt": attempt,
            "exit_code": proc.returncode,
            "stdout_tail": "\n".join((proc.stdout or "").splitlines()[-10:]),
            "stderr_tail": "\n".join((proc.stderr or "").splitlines()[-10:]),
            "duration_ms": round(duration_ms, 2),
            "pptx": str(pptx_out.resolve()) if pptx_out.is_file() else None,
            "preview": str(preview_out.resolve()) if preview_out.is_file() else None,
        }
        attempt_summaries.append(summary)
        last_summary = summary

        if proc.returncode == 0 and pptx_out.is_file():
            return {
                **summary,
                "attempts": attempt_summaries,
                "retry_count": attempt - 1,
            }

        if proc.returncode == 0 and not pptx_out.is_file():
            summary["stderr_tail"] = (
                summary["stderr_tail"]
                or f"convert.mjs did not write {pptx_out}"
            )

        if attempt < attempts:
            LOG.warning(
                "PPTX build failed on attempt %s/%s (exit=%s); retrying.",
                attempt,
                attempts,
                proc.returncode,
            )
            time.sleep(retry_delay_s)

    failure = {
        **(last_summary or {}),
        "attempts": attempt_summaries,
        "retry_count": max(0, attempts - 1),
    }
    raise BuilderFailure(failure)


class BuilderFailure(RuntimeError):
    """Wraps a non-zero ``convert.mjs`` outcome for the CLI to handle."""

    def __init__(self, payload: dict[str, Any]) -> None:
        super().__init__(f"convert.mjs failed: exit={payload.get('exit_code')}")
        self.payload = payload


# ---------------------------------------------------------------------------
# Step 3: render + report (PowerPoint COM + metrics + HTML)
# ---------------------------------------------------------------------------


def run_render(pptx_path: Path, render_png: Path) -> dict[str, Any]:
    """Render ``pptx_path`` to ``render_png`` via the existing renderer.

    This reuses :class:`slide2pptx.report.renderer.SlideRenderer` so the
    CLI never duplicates the COM/cleanup logic.
    """
    from slide2pptx.report import renderer

    render_png.parent.mkdir(parents=True, exist_ok=True)
    started = time.monotonic()
    opt = renderer.RenderOptions()
    ren = renderer.SlideRenderer(options=opt)
    try:
        result = ren.render(pptx_path, render_png)
    except renderer.RendererError as exc:
        raise RenderFailure(str(exc)) from exc
    except FileNotFoundError as exc:
        raise RenderFailure(str(exc)) from exc
    duration_ms = (time.monotonic() - started) * 1000
    return {
        "rendered_image": str(result.rendered_image.resolve()),
        "duration_ms": round(duration_ms, 2),
        "slide_index": result.slide_index,
        "warnings": list(result.warnings),
    }


class RenderFailure(RuntimeError):
    """Wraps a :class:`SlideRenderer` failure for the CLI to handle."""


def run_report(
    source_image: Path,
    rendered_image: Path,
    detected_json: Path,
    pptx_path: Path,
    report_dir: Path,
    *,
    threshold: int = 30,
) -> dict[str, Any]:
    """Build the metrics, heatmap and HTML report.

    The implementation mirrors :func:`slide2pptx.report_cli._cmd_full`
    but is inlined here so we can capture timings and produce a single
    JSON-friendly summary. No duplicated metric logic; everything is
    delegated to existing modules.
    """
    from slide2pptx.report import checklist, diff, html_builder, metrics, models

    report_dir.mkdir(parents=True, exist_ok=True)
    timings: dict[str, float] = {}

    started = time.monotonic()
    metric_values = metrics.compute_metrics(source_image, rendered_image, threshold=threshold)
    timings["metrics.compute_metrics_ms"] = round((time.monotonic() - started) * 1000, 2)

    started = time.monotonic()
    heatmap_path = report_dir / "diff.png"
    diff.render_heatmap(source_image, rendered_image, heatmap_path)
    timings["diff.render_heatmap_ms"] = round((time.monotonic() - started) * 1000, 2)

    editability_summary: models.EditabilitySummary | None = None
    elements: list = []
    if detected_json.is_file():
        try:
            started = time.monotonic()
            editability_summary = checklist.compute_editability(detected_json)
            elements = checklist.list_elements(detected_json)
            timings["checklist.compute_ms"] = round((time.monotonic() - started) * 1000, 2)
        except Exception as exc:  # noqa: BLE001 - defensive, continue without checklist
            LOG.warning("checklist failed: %s", exc)
            editability_summary = None
            elements = []

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
        job_id=report_dir.name or "report",
        source_image=source_image,
        rendered_image=rendered_image,
        heatmap=heatmap_path,
        diff_metrics=metric_values,
        editability=editability_summary,
        pptx_path=pptx_path,
        detected_json_path=detected_json,
        warnings=[],
        timings_ms=timings,
    )
    html_path = report_dir / "report.html"
    html_builder.render_report(report_inputs, elements, out_path=html_path)
    timings["html_builder.render_ms"] = round((time.monotonic() - started) * 1000, 2)

    return {
        "metrics": {
            "mae": metric_values.mae,
            "rmse": metric_values.rmse,
            "ssim": metric_values.ssim,
            "pixel_diff_ratio": metric_values.pixel_diff_ratio,
            "width": metric_values.width,
            "height": metric_values.height,
            "threshold": metric_values.threshold,
        },
        "heatmap": str(heatmap_path.resolve()),
        "html_report": str(html_path.resolve()),
        "editability": {
            "total_elements": editability_summary.total_elements,
            "editable_count": editability_summary.editable_count,
            "native_render_count": editability_summary.native_render_count,
            "image_render_count": editability_summary.image_render_count,
            "bitmap_fallback_count": editability_summary.bitmap_fallback_count,
            "low_confidence_count": editability_summary.low_confidence_count,
            "avg_editable_score": round(editability_summary.avg_editable_score, 4),
            "status": editability_summary.status.value,
        },
        "timings_ms": timings,
    }


def run_image_pipeline(
    image_path: Path,
    out_dir: Path,
    *,
    repo_root: Path,
    node_exe: str,
    skip_report: bool,
    threshold: int = 30,
    visual_passes: int = 2,
    second_pass_max_components: int = 96,
) -> dict[str, Any]:
    """Run detect -> build -> optional report for a single source image."""
    out_dir.mkdir(parents=True, exist_ok=True)
    detect_dir = out_dir / "detect"
    build_dir = out_dir / "build"
    report_dir = out_dir / "report"

    result: dict[str, Any] = {
        "ok": True,
        "input_image": str(image_path.resolve()),
        "out_dir": str(out_dir.resolve()),
        "skip_report": bool(skip_report),
    }

    LOG.info("Detecting elements in %s", image_path)
    result["detect"] = run_detect(
        image_path,
        detect_dir,
        visual_passes=visual_passes,
        second_pass_max_components=second_pass_max_components,
    )
    detected_json = Path(result["detect"]["detected_json"])

    LOG.info("Building PPTX via %s", node_exe)
    result["build"] = run_convert(node_exe, repo_root, detected_json, build_dir)

    pptx_path = Path(result["build"]["pptx"])
    if not pptx_path.is_file():
        raise BuilderFailure({
            "exit_code": 0,
            "stderr_tail": f"builder reported success but {pptx_path} missing",
            "pptx": None,
        })

    if skip_report:
        LOG.info("--skip-report set; skipping PowerPoint render + report.")
        result["report"] = {"skipped": True}
    else:
        render_png = report_dir / "powerPoint-render.png"
        LOG.info("Rendering PPTX -> %s", render_png)
        result["render"] = run_render(pptx_path, render_png)

        LOG.info("Computing metrics + heatmap + HTML report")
        result["report"] = run_report(
            source_image=image_path,
            rendered_image=render_png,
            detected_json=detected_json,
            pptx_path=pptx_path,
            report_dir=report_dir,
            threshold=threshold,
        )
    return result


def run_presentation_pipeline(
    presentation_path: Path,
    out_dir: Path,
    *,
    repo_root: Path,
    node_exe: str,
    skip_report: bool,
    threshold: int = 30,
    visual_passes: int = 2,
    second_pass_max_components: int = 96,
) -> dict[str, Any]:
    """Export a PPT/PPTX to images, then rebuild each slide independently."""
    out_dir.mkdir(parents=True, exist_ok=True)
    pages_dir = out_dir / "source-slides"
    export = ppt_input.export_presentation_slides(presentation_path, pages_dir)

    result: dict[str, Any] = {
        "ok": True,
        "input_presentation": str(presentation_path.resolve()),
        "out_dir": str(out_dir.resolve()),
        "slide_count": len(export.slides),
        "source_slides": str(pages_dir.resolve()),
        "skip_report": bool(skip_report),
        "slides": [],
        "failed_slides": [],
    }

    for slide in export.slides:
        slide_dir = out_dir / f"slide-{slide.index:03d}"
        LOG.info("Processing exported slide %s -> %s", slide.index, slide_dir)
        try:
            slide_result = run_image_pipeline(
                slide.image_path,
                slide_dir,
                repo_root=repo_root,
                node_exe=node_exe,
                skip_report=skip_report,
                threshold=threshold,
                visual_passes=visual_passes,
                second_pass_max_components=second_pass_max_components,
            )
            result["slides"].append({
                "ok": True,
                "index": slide.index,
                "source_image": str(slide.image_path),
                "out_dir": str(slide_dir.resolve()),
                "pptx": slide_result.get("build", {}).get("pptx"),
                "detected_json": slide_result.get("detect", {}).get("detected_json"),
                "warnings": slide_result.get("detect", {}).get("warnings", []),
                "retry_count": slide_result.get("build", {}).get("retry_count", 0),
            })
        except BuilderFailure as exc:
            failure = {
                "ok": False,
                "index": slide.index,
                "stage": "build",
                "source_image": str(slide.image_path),
                "out_dir": str(slide_dir.resolve()),
                "error": str(exc),
                "details": exc.payload,
            }
            LOG.error("Slide %s failed while building PPTX: %s", slide.index, exc)
            result["slides"].append(failure)
            result["failed_slides"].append(failure)
        except RenderFailure as exc:
            failure = {
                "ok": False,
                "index": slide.index,
                "stage": "render",
                "source_image": str(slide.image_path),
                "out_dir": str(slide_dir.resolve()),
                "error": str(exc),
            }
            LOG.error("Slide %s failed while rendering report: %s", slide.index, exc)
            result["slides"].append(failure)
            result["failed_slides"].append(failure)
    if result["failed_slides"]:
        result["ok"] = False
    return result


# ---------------------------------------------------------------------------
# Argument parsing & main
# ---------------------------------------------------------------------------


def _parse_args(argv: list[str] | None = None) -> argparse.Namespace:
    parser = argparse.ArgumentParser(
        prog="python -m slide2pptx.pipeline_cli",
        description=(
            "End-to-end orchestrator: detect -> artifact-build -> render -> report. "
            "Reuses slide2pptx.detect, slide2pptx.report and artifact-runtime/src/convert.mjs."
        ),
    )
    parser.add_argument(
        "input_file",
        type=Path,
        help="Path to an input slide image (PNG/JPG/etc.) or PPT/PPTX deck.",
    )
    parser.add_argument(
        "--out",
        dest="out_dir",
        type=Path,
        required=True,
        help="Output directory; detection/build/report sub-dirs are created inside.",
    )
    parser.add_argument(
        "--skip-report",
        action="store_true",
        help="Skip the PowerPoint render + report assembly (only detect + build).",
    )
    parser.add_argument(
        "--threshold",
        type=int,
        default=30,
        help="Pixel-diff threshold (0-255) forwarded to the report (default: 30).",
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
    parser.add_argument(
        "--log-level",
        default="INFO",
        help="Python logging level (default: INFO).",
    )
    return parser.parse_args(argv)


def _build_result(payload: dict[str, Any]) -> dict[str, Any]:
    """Flatten the orchestration payload into a JSON-safe summary.

    Path objects are coerced to strings; everything else is left alone
    because every step only returns dicts/strings/lists of primitives.
    """
    def _coerce(value: Any) -> Any:
        if isinstance(value, Path):
            return str(value.resolve())
        if isinstance(value, dict):
            return {k: _coerce(v) for k, v in value.items()}
        if isinstance(value, (list, tuple)):
            return [_coerce(v) for v in value]
        return value

    return _coerce(payload)


def main(argv: list[str] | None = None) -> int:
    parser = _parse_args(argv)
    logging.basicConfig(
        level=getattr(logging, parser.log_level.upper(), logging.INFO),
        format="%(asctime)s %(levelname)s %(name)s | %(message)s",
    )

    input_path: Path = parser.input_file
    out_dir: Path = parser.out_dir
    if not input_path.is_file():
        print(
            json.dumps(
                {"ok": False, "stage": "input", "error": f"input file not found: {input_path}"},
                indent=2,
            ),
            file=sys.stderr,
        )
        return EXIT_INPUT

    repo_root = resolve_repo_root(_THIS_FILE.parent)

    try:
        try:
            node_exe = find_node()
        except FileNotFoundError as exc:
            print(
                json.dumps(
                    {"ok": False, "stage": "node", "error": str(exc)},
                    indent=2,
                ),
                file=sys.stderr,
            )
            return EXIT_NODE_MISSING

        if input_path.suffix.lower() in PRESENTATION_SUFFIXES:
            result = run_presentation_pipeline(
                input_path,
                out_dir,
                repo_root=repo_root,
                node_exe=node_exe,
                skip_report=parser.skip_report,
                threshold=parser.threshold,
                visual_passes=parser.visual_passes,
                second_pass_max_components=parser.second_pass_max_components,
            )
        else:
            result = run_image_pipeline(
                input_path,
                out_dir,
                repo_root=repo_root,
                node_exe=node_exe,
                skip_report=parser.skip_report,
                threshold=parser.threshold,
                visual_passes=parser.visual_passes,
                second_pass_max_components=parser.second_pass_max_components,
            )

        print(json.dumps(_build_result(result), indent=2))
        if result.get("ok") is False:
            failed_slides = result.get("failed_slides") or []
            if failed_slides and failed_slides[0].get("stage") == "render":
                return EXIT_RENDER
            return EXIT_BUILDER
        return EXIT_OK
    except BuilderFailure as exc:
        print(
            json.dumps(
                {
                    "ok": False,
                    "stage": "build",
                    "error": str(exc),
                    "details": exc.payload,
                },
                indent=2,
            ),
            file=sys.stderr,
        )
        return EXIT_BUILDER
    except RenderFailure as exc:
        print(
            json.dumps({"ok": False, "stage": "render", "error": str(exc)}, indent=2),
            file=sys.stderr,
        )
        return EXIT_RENDER
    except ppt_input.PresentationInputError as exc:
        print(
            json.dumps({"ok": False, "stage": "presentation", "error": str(exc)}, indent=2),
            file=sys.stderr,
        )
        return EXIT_INPUT
    except Exception as exc:  # pragma: no cover - defensive
        LOG.exception("Unexpected error: %s", exc)
        print(
            json.dumps(
                {"ok": False, "stage": "orchestrator", "error": str(exc)},
                indent=2,
            ),
            file=sys.stderr,
        )
        return EXIT_OTHER


if __name__ == "__main__":
    sys.exit(main())
