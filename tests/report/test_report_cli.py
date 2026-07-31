"""Tests for the CLI entry point."""

from __future__ import annotations

import json
import sys
from pathlib import Path
from unittest import mock

import pytest
from PIL import Image

from slide2pptx.report import report_cli


@pytest.fixture
def call_cli(monkeypatch):
    """Helper: run main() with the supplied argv list and return exit code."""
    captured = {}

    def _run(argv):
        with mock.patch.object(sys, "argv", ["report_cli"] + argv):
            try:
                exit_code = report_cli.main(argv)
            except SystemExit as exc:
                exit_code = exc.code if isinstance(exc.code, int) else 1
        return exit_code

    return _run


# ---------------------------------------------------------------------------
# Top-level
# ---------------------------------------------------------------------------


def test_cli_help_exits_cleanly(capsys):
    with mock.patch.object(sys, "argv", ["report_cli", "--help"]):
        with pytest.raises(SystemExit):
            report_cli.main()
    captured = capsys.readouterr()
    assert "render" in captured.out
    assert "diff" in captured.out
    assert "full" in captured.out


# ---------------------------------------------------------------------------
# diff subcommand
# ---------------------------------------------------------------------------


def test_cli_diff_writes_metrics_and_heatmap(tmp_path, sample_png_pair, call_cli, monkeypatch):
    src, rendered = sample_png_pair
    out_dir = tmp_path / "out"
    argv = ["--log-level", "WARNING", "diff",
            "--source", str(src),
            "--rendered", str(rendered),
            "--out", str(out_dir),
            "--threshold", "20"]
    code = call_cli(argv)
    assert code == report_cli.EXIT_OK
    metrics_file = out_dir / "metrics.json"
    heatmap_file = out_dir / "diff.png"
    assert metrics_file.is_file()
    assert heatmap_file.is_file()
    payload = json.loads(metrics_file.read_text(encoding="utf-8"))
    assert "mae" in payload and "ssim" in payload
    assert payload["threshold"] == 20


def test_cli_diff_missing_input_returns_input_code(tmp_path, call_cli):
    code = call_cli(["diff", "--source", str(tmp_path / "no.png"),
                     "--rendered", str(tmp_path / "no.png"),
                     "--out", str(tmp_path / "out")])
    assert code == report_cli.EXIT_INPUT


# ---------------------------------------------------------------------------
# render subcommand
# ---------------------------------------------------------------------------


def test_cli_render_invokes_renderer(tmp_path, call_cli, monkeypatch):
    pptx = tmp_path / "x.pptx"
    pptx.write_bytes(b"PK")
    out_png = tmp_path / "x.png"

    monkeypatch.setattr(
        report_cli.renderer, "_resolve_powershell",
        lambda: "C:/pwsh.exe",
    )

    captured_kwargs = {}

    def fake_render(self, pptx_path, out_png_path, slide_index=None):
        captured_kwargs["pptx"] = pptx_path
        captured_kwargs["out"] = out_png_path
        captured_kwargs["slide"] = slide_index
        return report_cli.models.RenderResult(
            pptx_path=Path(pptx_path),
            rendered_image=Path(out_png_path),
            duration_ms=42.0,
            slide_index=slide_index or 1,
        )

    monkeypatch.setattr(report_cli.renderer.SlideRenderer, "render", fake_render)

    code = call_cli([
        "--log-level", "WARNING",
        "render", "--pptx", str(pptx), "--out", str(out_png),
        "--slide-index", "2",
    ])
    assert code == report_cli.EXIT_OK
    assert captured_kwargs["slide"] == 2


def test_cli_render_input_error(tmp_path, call_cli):
    code = call_cli([
        "render",
        "--pptx", str(tmp_path / "missing.pptx"),
        "--out", str(tmp_path / "out.png"),
    ])
    assert code == report_cli.EXIT_INPUT


def test_cli_render_renderer_error(tmp_path, call_cli, monkeypatch):
    pptx = tmp_path / "x.pptx"
    pptx.write_bytes(b"PK")
    out_png = tmp_path / "x.png"

    def fake_render(self, pptx_path, out_png_path, slide_index=None):
        raise report_cli.renderer.RendererError("boom")

    monkeypatch.setattr(report_cli.renderer.SlideRenderer, "render", fake_render)
    code = call_cli(["render", "--pptx", str(pptx), "--out", str(out_png)])
    assert code == report_cli.EXIT_RENDER


# ---------------------------------------------------------------------------
# full subcommand
# ---------------------------------------------------------------------------


def test_cli_full_pipeline(tmp_path, sample_png_pair, detected_json_path, call_cli):
    src, rendered = sample_png_pair
    out_dir = tmp_path / "out"
    out_dir.mkdir()
    code = call_cli([
        "--log-level", "WARNING",
        "full",
        "--source", str(src),
        "--rendered", str(rendered),
        "--out", str(out_dir),
        "--detected", str(detected_json_path),
        "--pptx", str(tmp_path / "fake.pptx"),
    ])
    assert code == report_cli.EXIT_OK
    report_html = out_dir / "report.html"
    diff_png = out_dir / "diff.png"
    assert report_html.is_file()
    assert diff_png.is_file()


def test_cli_full_works_without_detected(tmp_path, sample_png_pair, call_cli):
    src, rendered = sample_png_pair
    out_dir = tmp_path / "out"
    out_dir.mkdir()
    code = call_cli([
        "--log-level", "WARNING",
        "full",
        "--source", str(src),
        "--rendered", str(rendered),
        "--out", str(out_dir),
    ])
    assert code == report_cli.EXIT_OK
    assert (out_dir / "report.html").is_file()


def test_cli_full_missing_inputs(tmp_path, call_cli):
    code = call_cli([
        "full",
        "--source", str(tmp_path / "a.png"),
        "--rendered", str(tmp_path / "b.png"),
        "--out", str(tmp_path / "out"),
    ])
    assert code == report_cli.EXIT_INPUT
