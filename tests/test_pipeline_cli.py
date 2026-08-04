from __future__ import annotations

import subprocess
from pathlib import Path

import pytest

from slide2pptx import pipeline_cli, ppt_input


def _fake_repo(tmp_path: Path) -> Path:
    repo_root = tmp_path / "repo"
    script = repo_root / "artifact-runtime" / "src" / "convert.mjs"
    script.parent.mkdir(parents=True)
    script.write_text("// test placeholder\n", encoding="utf-8")
    return repo_root


def test_run_convert_retries_transient_failure(tmp_path, monkeypatch):
    repo_root = _fake_repo(tmp_path)
    detected_json = tmp_path / "detected.json"
    detected_json.write_text("{}", encoding="utf-8")
    build_dir = tmp_path / "build"
    calls = 0

    def fake_run(cmd, **kwargs):
        nonlocal calls
        calls += 1
        out_path = Path(cmd[cmd.index("--out") + 1])
        preview_path = Path(cmd[cmd.index("--preview") + 1])
        if calls == 1:
            return subprocess.CompletedProcess(cmd, 3221225786, "", "")
        out_path.write_bytes(b"PK")
        preview_path.write_bytes(b"PNG")
        return subprocess.CompletedProcess(cmd, 0, "", "")

    monkeypatch.setattr(pipeline_cli.subprocess, "run", fake_run)
    monkeypatch.setattr(pipeline_cli.time, "sleep", lambda _seconds: None)

    result = pipeline_cli.run_convert(
        "node.exe",
        repo_root,
        detected_json,
        build_dir,
        retry_delay_s=0,
    )

    assert calls == 2
    assert result["exit_code"] == 0
    assert result["retry_count"] == 1
    assert len(result["attempts"]) == 2
    assert Path(result["pptx"]).is_file()


def test_run_convert_raises_after_all_retries(tmp_path, monkeypatch):
    repo_root = _fake_repo(tmp_path)
    detected_json = tmp_path / "detected.json"
    detected_json.write_text("{}", encoding="utf-8")

    def fake_run(cmd, **kwargs):
        return subprocess.CompletedProcess(cmd, 3221225786, "", "")

    monkeypatch.setattr(pipeline_cli.subprocess, "run", fake_run)
    monkeypatch.setattr(pipeline_cli.time, "sleep", lambda _seconds: None)

    with pytest.raises(pipeline_cli.BuilderFailure) as exc_info:
        pipeline_cli.run_convert(
            "node.exe",
            repo_root,
            detected_json,
            tmp_path / "build",
            attempts=3,
            retry_delay_s=0,
        )

    assert exc_info.value.payload["exit_code"] == 3221225786
    assert exc_info.value.payload["retry_count"] == 2
    assert len(exc_info.value.payload["attempts"]) == 3


def test_presentation_pipeline_continues_after_slide_failure(tmp_path, monkeypatch):
    deck = tmp_path / "deck.pptx"
    deck.write_bytes(b"PK")
    slide_1 = tmp_path / "slide-001.png"
    slide_2 = tmp_path / "slide-002.png"
    slide_1.write_bytes(b"PNG1")
    slide_2.write_bytes(b"PNG2")
    export = ppt_input.PresentationExportResult(
        input_path=deck,
        output_dir=tmp_path / "pages",
        slides=(
            ppt_input.ExportedSlide(index=1, image_path=slide_1),
            ppt_input.ExportedSlide(index=2, image_path=slide_2),
        ),
        duration_ms=1.0,
    )

    monkeypatch.setattr(
        pipeline_cli.ppt_input,
        "export_presentation_slides",
        lambda _path, _out_dir: export,
    )

    def fake_image_pipeline(image_path, out_dir, **kwargs):
        if image_path == slide_1:
            raise pipeline_cli.BuilderFailure({"exit_code": 3221225786})
        return {
            "detect": {"detected_json": str(out_dir / "detect" / "detected.json"), "warnings": []},
            "build": {"pptx": str(out_dir / "build" / "reconstructed.pptx"), "retry_count": 0},
        }

    monkeypatch.setattr(pipeline_cli, "run_image_pipeline", fake_image_pipeline)

    result = pipeline_cli.run_presentation_pipeline(
        deck,
        tmp_path / "out",
        repo_root=tmp_path,
        node_exe="node.exe",
        skip_report=True,
    )

    assert result["ok"] is False
    assert [slide["index"] for slide in result["slides"]] == [1, 2]
    assert result["slides"][0]["ok"] is False
    assert result["slides"][1]["ok"] is True
    assert result["failed_slides"][0]["index"] == 1
