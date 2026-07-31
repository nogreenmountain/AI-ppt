"""Tests for the PowerPoint COM renderer.

We deliberately do NOT call PowerPoint in the unit tests; instead we
patch the renderer internals so the call boundary is exercised without
Office. Integration testing on real hardware is opt-in via the
``RUN_RENDERER_INTEGRATION=1`` environment variable.
"""

from __future__ import annotations

import sys
from pathlib import Path
from unittest import mock

import pytest

from slide2pptx.report import renderer
from slide2pptx.report.models import RenderResult
from slide2pptx.report.renderer import (
    PowerPointUnavailableError,
    RenderOptions,
    RendererError,
    RenderTimeoutError,
    SlideRenderer,
    invoke_com_export,
)


# ---------------------------------------------------------------------------
# PowerShell resolution
# ---------------------------------------------------------------------------


def test_resolve_powershell_uses_env_override(tmp_path):
    target = tmp_path / "pwsh.exe"
    target.write_text("", encoding="utf-8")
    with mock.patch.dict("os.environ", {"SLIDE2PPTX_POWERSHELL": str(target)}):
        resolved = renderer._resolve_powershell()
    assert resolved == str(target)


def test_resolve_powershell_env_override_rejects_missing(tmp_path):
    with mock.patch.dict("os.environ", {"SLIDE2PPTX_POWERSHELL": str(tmp_path / "nope.exe")}):
        with pytest.raises(PowerPointUnavailableError):
            renderer._resolve_powershell()


def test_resolve_powershell_falls_back_to_which():
    with mock.patch.dict("os.environ", {}, clear=False), \
         mock.patch.object(renderer.shutil, "which", return_value=None):
        with pytest.raises(PowerPointUnavailableError):
            renderer._resolve_powershell()


# ---------------------------------------------------------------------------
# invoke_com_export happy + error paths
# ---------------------------------------------------------------------------


@pytest.fixture
def sample_pptx(tmp_path) -> Path:
    p = tmp_path / "mini.pptx"
    p.write_bytes(b"PK\x03\x04")  # not a real pptx but enough for our tests
    return p


def test_invoke_com_export_returns_parsed_payload(tmp_path, sample_pptx, monkeypatch):
    out_png = tmp_path / "out.png"
    out_png.write_bytes(b"fake-png")
    fake_proc = mock.Mock()
    fake_proc.stdout = '{"success": true, "slideIndex": 1, "width": 1920, "height": 1080}'
    fake_proc.stderr = ""
    fake_proc.returncode = 0

    monkeypatch.setattr(renderer.subprocess, "run",
                        lambda *a, **kw: fake_proc)
    payload = invoke_com_export(
        sample_pptx, out_png,
        powershell_executable="C:/pwsh.exe",
    )
    assert payload["success"] is True
    assert payload["slideIndex"] == 1


def test_invoke_com_export_missing_pptx_raises(tmp_path):
    with pytest.raises(FileNotFoundError):
        invoke_com_export(tmp_path / "missing.pptx", tmp_path / "out.png")


def test_invoke_com_export_parses_failure_payload(sample_pptx, tmp_path, monkeypatch):
    out_png = tmp_path / "out.png"
    fake_proc = mock.Mock()
    fake_proc.stdout = '{"success": false, "error": "PowerPoint refused"}'
    fake_proc.stderr = "stderr noise"
    fake_proc.returncode = 0
    monkeypatch.setattr(renderer.subprocess, "run",
                        lambda *a, **kw: fake_proc)
    payload = invoke_com_export(sample_pptx, out_png,
                                powershell_executable="C:/pwsh.exe")
    assert payload["success"] is False
    assert "refused" in payload["error"]


def test_invoke_com_export_handles_empty_stdout(sample_pptx, tmp_path, monkeypatch):
    fake_proc = mock.Mock()
    fake_proc.stdout = ""
    fake_proc.stderr = "boom"
    fake_proc.returncode = 2
    monkeypatch.setattr(renderer.subprocess, "run",
                        lambda *a, **kw: fake_proc)
    payload = invoke_com_export(sample_pptx, tmp_path / "out.png",
                                powershell_executable="C:/pwsh.exe")
    assert payload["success"] is False
    assert "boom" in payload["error"] or "exit=" in payload["error"]


def test_invoke_com_export_handles_invalid_json(sample_pptx, tmp_path, monkeypatch):
    fake_proc = mock.Mock()
    fake_proc.stdout = "not-json"
    fake_proc.stderr = ""
    fake_proc.returncode = 0
    monkeypatch.setattr(renderer.subprocess, "run",
                        lambda *a, **kw: fake_proc)
    with pytest.raises(RendererError):
        invoke_com_export(sample_pptx, tmp_path / "out.png",
                          powershell_executable="C:/pwsh.exe")


def test_invoke_com_export_translates_timeout(sample_pptx, tmp_path, monkeypatch):
    def _raise(*a, **kw):
        raise renderer.subprocess.TimeoutExpired(cmd="powershell", timeout=5)
    monkeypatch.setattr(renderer.subprocess, "run", _raise)
    with pytest.raises(RenderTimeoutError):
        invoke_com_export(sample_pptx, tmp_path / "out.png",
                          powershell_executable="C:/pwsh.exe")


# ---------------------------------------------------------------------------
# SlideRenderer high-level behaviour
# ---------------------------------------------------------------------------


def test_renderer_refuses_non_windows():
    with mock.patch.object(renderer.sys, "platform", "linux"):
        ren = SlideRenderer()
        with pytest.raises(PowerPointUnavailableError):
            ren.render(Path("x.pptx"), Path("out.png"))


def test_renderer_render_success(tmp_path, monkeypatch):
    sample_pptx = tmp_path / "in.pptx"
    sample_pptx.write_bytes(b"PK")
    out_png = tmp_path / "out.png"
    out_png.write_bytes(b"png-data")

    fake_payload = {"success": True, "slideIndex": 1, "duration": 0.5}

    with mock.patch.object(renderer.sys, "platform", "win32"), \
         mock.patch.object(renderer, "invoke_com_export",
                           return_value=fake_payload) as invoke, \
         mock.patch.object(SlideRenderer, "_list_powerpoint_pids",
                           return_value=set()), \
         mock.patch.object(SlideRenderer, "_kill_powerpoint",
                           return_value=False), \
         mock.patch.object(SlideRenderer, "_finalize",
                           lambda self, **kw: None):
        ren = SlideRenderer(options=RenderOptions(timeout_seconds=15))
        result = ren.render(sample_pptx, out_png)

    invoke.assert_called_once()
    assert isinstance(result, RenderResult)
    assert result.rendered_image == out_png
    assert result.duration_ms >= 0


def test_renderer_render_failure_raises(tmp_path, monkeypatch):
    sample_pptx = tmp_path / "in.pptx"
    sample_pptx.write_bytes(b"PK")
    out_png = tmp_path / "out.png"

    fake_payload = {"success": False, "error": "Office missing"}

    with mock.patch.object(renderer.sys, "platform", "win32"), \
         mock.patch.object(renderer, "invoke_com_export",
                           return_value=fake_payload), \
         mock.patch.object(SlideRenderer, "_list_powerpoint_pids",
                           return_value=set()), \
         mock.patch.object(SlideRenderer, "_kill_powerpoint",
                           return_value=False), \
         mock.patch.object(SlideRenderer, "_finalize",
                           lambda self, **kw: None):
        ren = SlideRenderer()
        with pytest.raises(RendererError):
            ren.render(sample_pptx, out_png)


def test_renderer_finalize_kills_leaked_pids(tmp_path):
    ren = SlideRenderer()
    ren._kill_powerpoint = mock.Mock(return_value=True)
    ren._list_powerpoint_pids = mock.Mock(return_value={1234, 5678})
    ren._finalize(pre_pids={1234}, options=RenderOptions(timeout_seconds=5))
    # Only the leaked pid (5678) should be killed.
    ren._kill_powerpoint.assert_called_once_with(5678)


def test_renderer_finalize_skips_when_cleanup_disabled():
    ren = SlideRenderer(options=RenderOptions(timeout_seconds=5, cleanup_zombie=False))
    ren._list_powerpoint_pids = mock.Mock(return_value={9999})
    ren._kill_powerpoint = mock.Mock()
    ren._finalize(pre_pids=set(), options=RenderOptions(timeout_seconds=5, cleanup_zombie=False))
    ren._kill_powerpoint.assert_not_called()


# ---------------------------------------------------------------------------
# PID helpers
# ---------------------------------------------------------------------------


def test_list_powerpoint_pids_fallback_empty_when_tasklist_missing(monkeypatch):
    def _raise(*a, **kw):
        raise FileNotFoundError("no tasklist")
    monkeypatch.setattr(renderer.subprocess, "run", _raise)
    assert renderer.SlideRenderer()._list_powerpoint_pids_fallback() == set()


def test_list_powerpoint_pids_fallback_parses_csv(monkeypatch):
    fake_proc = mock.Mock()
    fake_proc.stdout = (
        '"POWERPNT.EXE","1234","Console","1","100,000 K"\r\n'
        '"POWERPNT.EXE","5678","Console","1","200,000 K"\r\n'
        '"notepad.exe","9999","Console","1","40,000 K"'
    )
    fake_proc.returncode = 0
    monkeypatch.setattr(renderer.subprocess, "run",
                        lambda *a, **kw: fake_proc)
    pids = renderer.SlideRenderer()._list_powerpoint_pids_fallback()
    assert pids == {1234, 5678}


def test_list_powerpoint_pids_uses_psutil(monkeypatch):
    """When psutil is available, the higher-level helper should prefer it."""
    class _Pid:
        pid = 111
        def info_get(self, key, default=None):
            return {"name": "POWERPNT.EXE"}.get(key, default)

    # Build a process-like object
    class _Proc:
        info = {"name": "POWERPNT.EXE"}
        @property
        def pid(self):
            return 222

    class _Psutil:
        def process_iter(self, attrs=None):
            return [_Proc()]

    fake_psutil = _Psutil()
    monkeypatch.setitem(sys.modules, "psutil", fake_psutil)
    ren = renderer.SlideRenderer()
    ren._list_powerpoint_pids_fallback = mock.Mock(return_value=set())
    pids = ren._list_powerpoint_pids()
    assert 222 in pids
    ren._list_powerpoint_pids_fallback.assert_not_called()


# ---------------------------------------------------------------------------
# Integration gate
# ---------------------------------------------------------------------------


@pytest.mark.skipif(
    "not config.getoption('--run-integration')",
    reason="integration gate off",
)
def test_render_integration_when_allowed(tmp_path, sample_pptx, integration_env):
    """Gate-tested integration smoke. Off by default; enabled with --run-integration."""
    if not integration_env["enabled"]:
        pytest.skip("integration disabled")
    out_png = tmp_path / "real.png"
    ren = SlideRenderer(options=RenderOptions(timeout_seconds=120))
    result = ren.render(sample_pptx, out_png)
    assert out_png.is_file()
    assert result.duration_ms >= 0
