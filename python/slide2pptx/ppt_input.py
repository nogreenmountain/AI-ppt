"""PowerPoint input support for slide2pptx.

This module exports every slide in a local PPT/PPTX deck to PNG files so the
existing image pipeline can rebuild them one page at a time. It uses Windows
PowerPoint COM through PowerShell and imports no Windows-only modules at import
time, so the package remains importable on CI and non-Windows machines.
"""

from __future__ import annotations

import json
import shutil
import subprocess
import sys
import tempfile
import time
from dataclasses import dataclass
from pathlib import Path
from typing import Any


class PresentationInputError(RuntimeError):
    """Base class for PPT/PPTX input failures."""


class PresentationAppUnavailableError(PresentationInputError):
    """Raised when PowerPoint automation is unavailable."""


class PresentationExportTimeoutError(PresentationInputError):
    """Raised when PPT/PPTX slide export exceeds the timeout."""


@dataclass(frozen=True)
class ExportedSlide:
    """A rendered source slide image."""

    index: int
    image_path: Path
    width_pt: int = 0
    height_pt: int = 0


@dataclass(frozen=True)
class PresentationExportResult:
    """Summary of a PPT/PPTX export."""

    input_path: Path
    output_dir: Path
    slides: tuple[ExportedSlide, ...]
    duration_ms: float


_POWER_SHELL_EXPORT_ALL = r"""# Auto-generated slide2pptx deck exporter.
[CmdletBinding()]
param(
    [Parameter(Mandatory = $true)][string]$PptPath,
    [Parameter(Mandatory = $true)][string]$OutDir
)
[Console]::OutputEncoding = [System.Text.Encoding]::UTF8
$ErrorActionPreference = 'Stop'
$result = [ordered]@{
    success = $false
    error = $null
    slideCount = 0
    width = 0
    height = 0
    slides = @()
}
$pres = $null
$app = $null
try {
    New-Item -ItemType Directory -Force -Path $OutDir | Out-Null
    $app = New-Object -ComObject PowerPoint.Application
    $app.Visible = -1
    $pres = $app.Presentations.Open($PptPath, $true, $false, $false)
    $result.slideCount = [int]$pres.Slides.Count
    $result.width = [int]$pres.SlideSize.Width
    $result.height = [int]$pres.SlideSize.Height
    for ($i = 1; $i -le $pres.Slides.Count; $i++) {
        $fileName = "slide-{0:D3}.png" -f $i
        $outPath = Join-Path $OutDir $fileName
        $pres.Slides[$i].Export($outPath, "PNG", 0, 0)
        $result.slides += [ordered]@{
            index = $i
            image = $outPath
        }
    }
    $result.success = $true
}
catch {
    $result.error = "$($_.Exception.Message)"
}
finally {
    try { if ($null -ne $pres) { $pres.Close() | Out-Null } } catch { }
    try { if ($null -ne $app) { $app.Quit() | Out-Null } } catch { }
    try { [System.Runtime.InteropServices.Marshal]::ReleaseComObject($pres) | Out-Null } catch { }
    try { [System.Runtime.InteropServices.Marshal]::ReleaseComObject($app) | Out-Null } catch { }
    [GC]::Collect(); [GC]::WaitForPendingFinalizers()
}
$result | ConvertTo-Json -Compress -Depth 5
"""


def _resolve_powershell() -> str:
    pwsh = shutil.which("powershell.exe") or shutil.which("powershell")
    if pwsh is None:
        raise PresentationAppUnavailableError(
            "powershell.exe not found; PPT/PPTX input requires Windows PowerPoint."
        )
    return pwsh


def _parse_payload(stdout: str, stderr: str, returncode: int) -> dict[str, Any]:
    text = (stdout or "").strip()
    if text.startswith("{") and text.endswith("}"):
        try:
            return json.loads(text)
        except json.JSONDecodeError as exc:
            raise PresentationInputError(f"Could not parse PowerPoint export JSON: {exc}") from exc
    if not text:
        return {
            "success": False,
            "error": stderr or f"PowerShell exited with code {returncode}",
        }
    raise PresentationInputError(f"PowerPoint exporter returned invalid output: {text!r}")


def export_presentation_slides(
    ppt_path: Path,
    out_dir: Path,
    *,
    timeout_seconds: float = 600.0,
    powershell_executable: str | None = None,
) -> PresentationExportResult:
    """Export every slide in ``ppt_path`` to PNG images under ``out_dir``.

    Args:
        ppt_path: Input .ppt or .pptx file.
        out_dir: Destination folder for slide-001.png, slide-002.png, ...
        timeout_seconds: Hard timeout for the whole deck export.
        powershell_executable: Test hook or custom PowerShell path.
    """
    if sys.platform != "win32":
        raise PresentationAppUnavailableError("PPT/PPTX input requires Windows PowerPoint.")

    ppt_path = Path(ppt_path).resolve()
    out_dir = Path(out_dir).resolve()
    if not ppt_path.is_file():
        raise FileNotFoundError(f"Presentation not found: {ppt_path}")
    out_dir.mkdir(parents=True, exist_ok=True)

    ps = powershell_executable or _resolve_powershell()
    with tempfile.NamedTemporaryFile(
        mode="w",
        encoding="utf-8-sig",
        suffix=".ps1",
        delete=False,
    ) as script_file:
        script_file.write(_POWER_SHELL_EXPORT_ALL)
        script_path = Path(script_file.name)

    cmd = [
        ps,
        "-NoLogo",
        "-NoProfile",
        "-NonInteractive",
        "-ExecutionPolicy",
        "Bypass",
        "-File",
        str(script_path),
        "-PptPath",
        str(ppt_path),
        "-OutDir",
        str(out_dir),
    ]

    started = time.monotonic()
    try:
        proc = subprocess.run(
            cmd,
            capture_output=True,
            text=True,
            encoding="utf-8",
            errors="replace",
            timeout=timeout_seconds,
            check=False,
        )
    except subprocess.TimeoutExpired as exc:
        raise PresentationExportTimeoutError(
            f"PowerPoint deck export exceeded {timeout_seconds}s"
        ) from exc
    finally:
        script_path.unlink(missing_ok=True)

    payload = _parse_payload(proc.stdout, proc.stderr, proc.returncode)
    if not payload.get("success"):
        raise PresentationInputError(payload.get("error") or "PowerPoint deck export failed")

    slides = tuple(
        ExportedSlide(
            index=int(item["index"]),
            image_path=Path(item["image"]).resolve(),
            width_pt=int(payload.get("width") or 0),
            height_pt=int(payload.get("height") or 0),
        )
        for item in payload.get("slides", [])
    )
    if not slides:
        raise PresentationInputError("PowerPoint exported zero slides.")
    missing = [str(slide.image_path) for slide in slides if not slide.image_path.is_file()]
    if missing:
        raise PresentationInputError(f"PowerPoint reported slides that were not created: {missing}")

    return PresentationExportResult(
        input_path=ppt_path,
        output_dir=out_dir,
        slides=slides,
        duration_ms=round((time.monotonic() - started) * 1000, 2),
    )


__all__ = [
    "ExportedSlide",
    "PresentationAppUnavailableError",
    "PresentationExportResult",
    "PresentationExportTimeoutError",
    "PresentationInputError",
    "export_presentation_slides",
]
