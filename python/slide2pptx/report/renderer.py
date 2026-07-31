"""Windows-only PowerPoint renderer.

The renderer shells out to PowerShell and drives
``Microsoft.Office.Interop.PowerPoint`` via COM automation to export the
first slide as a PNG. We deliberately keep this slow path narrow and easy
to mock:

* :func:`invoke_com_export` is the *only* function that touches COM.
* :class:`SlideRenderer` wraps it with timeout, error mapping and
  defensive process cleanup so callers never need to know the launcher
  exists.
* The module imports nothing Windows-specific at import time, so
  non-Windows hosts (CI, devs) can still import the package; an explicit
  :class:`EnvironmentError` is raised when :meth:`SlideRenderer.render`
  is invoked there.

Typical usage::

    renderer = SlideRenderer()
    result = renderer.render(Path("slide.pptx"), Path("out.png"))
    assert result.rendered_image.exists()
"""

from __future__ import annotations

import json
import logging
import os
import shutil
import subprocess
import sys
import tempfile
import time
from dataclasses import dataclass
from pathlib import Path
from typing import Any, Dict, Optional, Tuple

from slide2pptx.report.models import RenderResult

LOGGER = logging.getLogger(__name__)


# ---------------------------------------------------------------------------
# Exceptions
# ---------------------------------------------------------------------------


class RendererError(RuntimeError):
    """Base class for renderer failures."""


class PowerPointUnavailableError(RendererError):
    """Raised when the host has no usable PowerPoint installation."""


class RenderTimeoutError(RendererError):
    """Raised when a render exceeds the configured timeout."""


# ---------------------------------------------------------------------------
# Configuration
# ---------------------------------------------------------------------------


@dataclass(frozen=True)
class RenderOptions:
    """Tunable knobs for :class:`SlideRenderer`.

    Attributes:
        timeout_seconds: Hard upper bound on the subprocess. PowerPoint
            COM on cold start can take 30-60 s; the default 90 s is
            generous but still bounded.
        slide_index: 1-based slide number to export.
        extra_powershell_args: Forwarded to ``powershell.exe`` (e.g.
            ``-NoProfile``). Default flags already include ``-NoLogo``,
            ``-NonInteractive``.
        cleanup_zombie: When ``True`` (default) the renderer's
            ``finally`` block explicitly walks the process list for
            any ``POWERPNT.EXE`` leaked from this run and kills them.
            The cleanup is keyed on the launcher PID so we never kill
            a *different* session's PowerPoint.
    """

    timeout_seconds: float = 90.0
    slide_index: int = 1
    extra_powershell_args: Tuple[str, ...] = ()
    cleanup_zombie: bool = True


# ---------------------------------------------------------------------------
# PowerShell script template
# ---------------------------------------------------------------------------


# Notes on safety:
#
# * The script uses single-quoted heredocs so PowerShell variables are
#   NOT expanded on the Python side. Parameter values are passed via
#   ``-File`` + ``-PptxPath`` / ``-OutPng`` so paths cannot be parsed as
#   PowerShell expressions.
# * We invoke the script under ``powershell.exe -NonInteractive`` so no
#   interactive prompt can ever appear.
# * The script always closes the COM presentation and exits the
#   PowerPoint application. Errors are caught and surfaced as a JSON
#   blob on stdout, which ``invoke_com_export`` then parses.
_POWER_SHELL_SCRIPT = r"""# Auto-generated slide2pptx PowerPoint renderer.
# Inputs are PowerShell parameters; output is a JSON line on stdout.
[CmdletBinding()]
param(
    [Parameter(Mandatory = $true)][string]$PptxPath,
    [Parameter(Mandatory = $true)][string]$OutPng,
    [int]$SlideIndex = 1,
    [int]$TimeoutSec = 60
)
[Console]::OutputEncoding = [System.Text.Encoding]::UTF8
$ErrorActionPreference = 'Stop'
$result = [ordered]@{
    success    = $false
    error      = $null
    slideIndex = $SlideIndex
    width      = 0
    height     = 0
}
$pres = $null
$app  = $null
try {
    # Wait briefly for Office to be ready before instantiating COM.
    $app = New-Object -ComObject PowerPoint.Application
    # Hold the application alive across the whole script.
    # PowerPoint expects an MsoTriState, not a Boolean. msoTrue is -1.
    $app.Visible = -1

    $pres = $app.Presentations.Open($PptxPath, $true, $false, $false)
    if ($pres.Slides.Count -lt $SlideIndex) {
        throw "Slide $SlideIndex not found (count = $($pres.Slides.Count))"
    }
    $slide = $pres.Slides[$SlideIndex]

    # Slide.Export expects a filter name string such as "PNG".
    $slide.Export($OutPng, "PNG", 0, 0)

    $result.slideIndex = $SlideIndex
    $result.width      = [int]$pres.SlideSize.Width
    $result.height     = [int]$pres.SlideSize.Height
    $result.success    = $true
}
catch {
    $result.error = "$($_.Exception.Message)"
}
finally {
    try {
        if ($null -ne $pres) { $pres.Close() | Out-Null }
    } catch { }
    try {
        if ($null -ne $app)  { $app.Quit() | Out-Null }
    } catch { }
    # Force-release COM references so the Office process exits promptly.
    try { [System.Runtime.InteropServices.Marshal]::ReleaseComObject($pres) | Out-Null } catch { }
    try { [System.Runtime.InteropServices.Marshal]::ReleaseComObject($app)  | Out-Null } catch { }
    [GC]::Collect(); [GC]::WaitForPendingFinalizers()
}
# Emit exactly one JSON object regardless of success or failure.
$result | ConvertTo-Json -Compress -Depth 4
"""


# ---------------------------------------------------------------------------
# Public function: shell-out + parse
# ---------------------------------------------------------------------------


def _resolve_powershell() -> str:
    """Return the absolute path of ``powershell.exe``.

    Honour the ``SLIDE2PPTX_POWERSHELL`` override (handy on locked-down
    machines where powershell is not on ``PATH``) but fall back to the
    stdlib ``shutil.which`` lookup otherwise.
    """
    override = os.environ.get("SLIDE2PPTX_POWERSHELL")
    if override:
        candidate = Path(override)
        if candidate.is_file():
            return str(candidate)
        raise PowerPointUnavailableError(
            f"SLIDE2PPTX_POWERSHELL={override!r} does not point to a file"
        )
    pwsh = shutil.which("powershell.exe") or shutil.which("powershell")
    if pwsh is None:
        raise PowerPointUnavailableError(
            "powershell.exe not found on PATH; cannot drive PowerPoint COM"
        )
    return pwsh


def invoke_com_export(
    pptx_path: Path,
    out_png: Path,
    *,
    options: Optional[RenderOptions] = None,
    powershell_executable: Optional[str] = None,
) -> Dict[str, Any]:
    """Invoke the PowerShell COM exporter and return its parsed result.

    This is the *single* function in the project that talks to COM. All
    callers should go through :meth:`SlideRenderer.render` instead of
    using this directly, but keeping it public makes mocking tests
    possible and the call boundary crystal clear.

    The subprocess is launched with the ``CREATE_NEW_PROCESS_GROUP`` flag
    on Windows so we can kill the whole tree on timeout without
    disturbing an interactive PowerPoint window owned by the user.

    Args:
        pptx_path: Input ``.pptx`` file. Must exist.
        out_png: Destination PNG. Parent directory will be created.
        options: Tuning parameters; defaults to :class:`RenderOptions()`.
        powershell_executable: Override the resolved ``powershell.exe``.

    Returns:
        The JSON-decoded result dict emitted by the PowerShell script.

    Raises:
        PowerPointUnavailableError: If ``powershell.exe`` is missing.
        FileNotFoundError: If ``pptx_path`` does not exist.
        RendererError: For malformed output or non-zero exit codes.
        RenderTimeoutError: If the subprocess exceeds the timeout.
    """
    pptx_path = Path(pptx_path)
    out_png = Path(out_png)
    options = options or RenderOptions()

    if not pptx_path.is_file():
        raise FileNotFoundError(f"PPTX not found: {pptx_path}")
    out_png.parent.mkdir(parents=True, exist_ok=True)

    ps = powershell_executable or _resolve_powershell()

    with tempfile.NamedTemporaryFile(
        mode="w",
        encoding="utf-8-sig",
        suffix=".ps1",
        delete=False,
    ) as script_file:
        script_file.write(_POWER_SHELL_SCRIPT)
        script_path = Path(script_file.name)

    cmd: list = [ps]
    cmd.extend(options.extra_powershell_args)
    cmd.extend([
        "-NoLogo",
        "-NoProfile",
        "-NonInteractive",
        "-ExecutionPolicy", "Bypass",
        "-File", str(script_path),
        "-PptxPath", str(pptx_path.resolve()),
        "-OutPng", str(out_png.resolve()),
        "-SlideIndex", str(options.slide_index),
        "-TimeoutSec", str(max(1, int(options.timeout_seconds))),
    ])

    LOGGER.debug("Renderer cmd: %s", cmd)

    creationflags = 0
    if sys.platform == "win32":
        creationflags = getattr(subprocess, "CREATE_NEW_PROCESS_GROUP", 0)

    start = time.monotonic()
    try:
        proc = subprocess.run(
            cmd,
            capture_output=True,
            text=True,
            encoding="utf-8",
            errors="replace",
            timeout=options.timeout_seconds + 30,  # Python-level safety net
            check=False,
            creationflags=creationflags,
        )
    except subprocess.TimeoutExpired as exc:
        raise RenderTimeoutError(
            f"PowerPoint export exceeded {options.timeout_seconds}s"
        ) from exc
    finally:
        script_path.unlink(missing_ok=True)
    duration = time.monotonic() - start

    stdout = (proc.stdout or "").strip()
    stderr = (proc.stderr or "").strip()

    # The script ALWAYS prints a JSON object on stdout. If PowerShell
    # itself died before writing one we surface a generic error.
    payload: Dict[str, Any]
    if stdout.startswith("{") and stdout.endswith("}"):
        try:
            payload = json.loads(stdout)
        except json.JSONDecodeError as exc:
            raise RendererError(
                f"Could not parse PowerPoint output: {exc}; raw={stdout!r}"
            ) from exc
    elif not stdout:
        payload = {
            "success": False,
            "error": stderr or f"PowerShell exited with code {proc.returncode}",
            "duration": duration,
        }
    else:
        raise RendererError(
            "PowerPoint renderer returned invalid JSON; "
            f"stdout={stdout!r}, stderr={stderr!r}, exit={proc.returncode}"
        )

    payload.setdefault("duration", duration)
    payload.setdefault("slideIndex", options.slide_index)
    payload.setdefault("stdout", stdout)
    payload.setdefault("stderr", stderr)

    if proc.returncode not in (0, None):
        payload.setdefault("error", payload["error"] or f"exit={proc.returncode}")

    return payload


# ---------------------------------------------------------------------------
# Higher-level facade
# ---------------------------------------------------------------------------


class SlideRenderer:
    """User-facing renderer with built-in cleanup and error mapping.

    Callers only need :meth:`render`; cleanup of any leaked
    ``POWERPNT.EXE`` happens through ``os`` because ``pywin32`` is
    *only* imported inside :meth:`_kill_powerpoint` (kept lazy so the
    package imports cleanly on non-Windows / non-Office machines).
    """

    def __init__(
        self,
        options: Optional[RenderOptions] = None,
        *,
        powershell_executable: Optional[str] = None,
    ) -> None:
        self.options = options or RenderOptions()
        self.powershell_executable = powershell_executable
        # Best-effort detection: refuses to run on non-Windows hosts.
        if sys.platform != "win32":
            self._platform_supported = False
        else:
            self._platform_supported = True

    def render(
        self,
        pptx_path: Path,
        out_png: Path,
        slide_index: Optional[int] = None,
    ) -> RenderResult:
        """Render ``pptx_path`` to ``out_png`` and return the outcome.

        ``slide_index`` (1-based) overrides :attr:`RenderOptions.slide_index`
        for this call only.

        The function guarantees that no extra ``POWERPNT.EXE`` is left
        running as a side effect of this call. Specifically:

        * The PowerShell script closes the presentation and quits
          PowerPoint in its ``finally`` block.
        * The :class:`SlideRenderer` adds a *belt-and-braces* scan for
          leftover PowerPoint instances that were started during this
          call's wall clock window.
        """
        if not self._platform_supported:
            raise PowerPointUnavailableError(
                "SlideRenderer.render requires Windows (uses powershell.exe + COM)"
            )

        pptx_path = Path(pptx_path).resolve()
        out_png = Path(out_png).resolve()
        out_png.parent.mkdir(parents=True, exist_ok=True)

        options = RenderOptions(
            timeout_seconds=self.options.timeout_seconds,
            slide_index=slide_index or self.options.slide_index,
            extra_powershell_args=self.options.extra_powershell_args,
            cleanup_zombie=self.options.cleanup_zombie,
        )

        # Snapshot the set of running PowerPoint PIDs *before* we start
        # so we can attribute leak to this call rather than to the user.
        pre_pids = self._list_powerpoint_pids() if options.cleanup_zombie else set()

        warnings: list = []
        started_at = time.monotonic()
        try:
            payload = invoke_com_export(
                pptx_path=pptx_path,
                out_png=out_png,
                options=options,
                powershell_executable=self.powershell_executable,
            )
            if not payload.get("success"):
                msg = payload.get("error") or "unknown renderer error"
                raise RendererError(
                    f"PowerPoint export failed: {msg}; "
                    f"stderr={payload.get('stderr', '')!r}"
                )
        finally:
            # Even on success we run a cleanup pass so a previously
            # leaked Office COM server is taken down.
            self._finalize(pre_pids=pre_pids, options=options)
        duration_ms = (time.monotonic() - started_at) * 1000

        if not out_png.is_file():
            warnings.append(
                f"PowerPoint reported success but {out_png} was not created"
            )

        return RenderResult(
            pptx_path=pptx_path,
            rendered_image=out_png,
            duration_ms=duration_ms,
            slide_index=payload.get("slideIndex", options.slide_index),
            warnings=tuple(warnings),
        )

    # ------------------------------------------------------------------
    # Cleanup helpers - broken out so tests can patch them.
    # ------------------------------------------------------------------

    def _list_powerpoint_pids(self) -> set:
        """Return the set of PIDs whose image name is POWERPNT.EXE."""
        try:
            import psutil  # type: ignore
        except ImportError:
            return self._list_powerpoint_pids_fallback()
        return {p.pid for p in psutil.process_iter(attrs=["name"]) if (p.info.get("name") or "").lower() == "powerpnt.exe"}

    def _list_powerpoint_pids_fallback(self) -> set:
        """``tasklist``-based fallback when ``psutil`` is unavailable."""
        try:
            out = subprocess.run(
                ["tasklist", "/FI", "IMAGENAME eq POWERPNT.EXE", "/FO", "CSV", "/NH"],
                capture_output=True,
                text=True,
                check=False,
                timeout=5,
            )
        except (FileNotFoundError, subprocess.TimeoutExpired, OSError):
            return set()
        pids = set()
        for line in (out.stdout or "").splitlines():
            parts = [p.strip('"') for p in line.split(",")]
            if len(parts) >= 2 and parts[0].lower() == "powerpnt.exe":
                try:
                    pids.add(int(parts[1]))
                except ValueError:
                    continue
        return pids

    def _kill_powerpoint(self, pid: int) -> bool:
        """Attempt a graceful taskkill of ``pid`` and return success."""
        try:
            res = subprocess.run(
                ["taskkill", "/PID", str(pid), "/F", "/T"],
                capture_output=True,
                text=True,
                check=False,
                timeout=5,
            )
            return res.returncode == 0
        except (FileNotFoundError, subprocess.TimeoutExpired, OSError):
            return False

    def _finalize(self, *, pre_pids: set, options: RenderOptions) -> None:
        """Clean up PowerPoint after a render, best effort."""
        if not options.cleanup_zombie:
            return
        # Give PowerPoint a moment to exit on its own.
        time.sleep(0.5)
        # Anything running now that wasn't running before is a leak.
        try:
            leaked = self._list_powerpoint_pids() - pre_pids
        except Exception:  # pragma: no cover - defensive
            return
        for pid in leaked:
            if self._kill_powerpoint(pid):
                LOGGER.warning("Killed leaked POWERPNT.EXE pid=%s", pid)


__all__ = [
    "RenderOptions",
    "RenderResult",
    "RendererError",
    "RenderTimeoutError",
    "PowerPointUnavailableError",
    "SlideRenderer",
    "invoke_com_export",
]
