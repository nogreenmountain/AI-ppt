from __future__ import annotations

from unittest import mock

import pytest

from slide2pptx import ppt_input


def test_export_presentation_slides_parses_payload(tmp_path, monkeypatch):
    deck = tmp_path / "deck.pptx"
    deck.write_bytes(b"PK")
    out_dir = tmp_path / "pages"
    slide_1 = out_dir / "slide-001.png"
    slide_2 = out_dir / "slide-002.png"
    out_dir.mkdir()
    slide_1.write_bytes(b"png1")
    slide_2.write_bytes(b"png2")
    slide_1_json = str(slide_1).replace("\\", "\\\\")
    slide_2_json = str(slide_2).replace("\\", "\\\\")

    fake_proc = mock.Mock()
    fake_proc.returncode = 0
    fake_proc.stderr = ""
    fake_proc.stdout = (
        '{"success":true,"slideCount":2,"width":960,"height":540,'
        f'"slides":[{{"index":1,"image":"{slide_1_json}"}},'
        f'{{"index":2,"image":"{slide_2_json}"}}]}}'
    )

    monkeypatch.setattr(ppt_input.sys, "platform", "win32")
    monkeypatch.setattr(ppt_input.subprocess, "run", lambda *a, **kw: fake_proc)

    result = ppt_input.export_presentation_slides(
        deck,
        out_dir,
        powershell_executable="C:/Windows/System32/WindowsPowerShell/v1.0/powershell.exe",
    )

    assert len(result.slides) == 2
    assert result.slides[0].index == 1
    assert result.slides[0].image_path == slide_1.resolve()
    assert result.slides[1].width_pt == 960


def test_export_presentation_slides_rejects_non_windows(tmp_path, monkeypatch):
    deck = tmp_path / "deck.pptx"
    deck.write_bytes(b"PK")
    monkeypatch.setattr(ppt_input.sys, "platform", "linux")

    with pytest.raises(ppt_input.PresentationAppUnavailableError):
        ppt_input.export_presentation_slides(deck, tmp_path / "pages")


def test_export_presentation_slides_missing_file(tmp_path, monkeypatch):
    monkeypatch.setattr(ppt_input.sys, "platform", "win32")

    with pytest.raises(FileNotFoundError):
        ppt_input.export_presentation_slides(
            tmp_path / "missing.pptx",
            tmp_path / "pages",
            powershell_executable="powershell.exe",
        )


def test_export_presentation_slides_failure_payload(tmp_path, monkeypatch):
    deck = tmp_path / "deck.pptx"
    deck.write_bytes(b"PK")
    fake_proc = mock.Mock()
    fake_proc.returncode = 0
    fake_proc.stderr = ""
    fake_proc.stdout = '{"success":false,"error":"PowerPoint refused"}'

    monkeypatch.setattr(ppt_input.sys, "platform", "win32")
    monkeypatch.setattr(ppt_input.subprocess, "run", lambda *a, **kw: fake_proc)

    with pytest.raises(ppt_input.PresentationInputError, match="refused"):
        ppt_input.export_presentation_slides(
            deck,
            tmp_path / "pages",
            powershell_executable="powershell.exe",
        )
