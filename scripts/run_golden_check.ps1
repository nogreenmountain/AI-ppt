param(
  [string]$OutDir
)

$ErrorActionPreference = "Stop"
$RepoRoot = (Resolve-Path (Join-Path $PSScriptRoot "..")).Path

if ([string]::IsNullOrWhiteSpace($OutDir)) {
  $stamp = Get-Date -Format "yyyyMMdd-HHmmss"
  $OutDir = Join-Path $RepoRoot "outputs\golden-check-$stamp"
}

$env:SLIDE2PPTX_NO_PAUSE = "1"
$inputImage = Join-Path $RepoRoot "samples\industry-teaching-research.png"

& (Join-Path $PSScriptRoot "setup_and_run.ps1") $inputImage $OutDir
if ($LASTEXITCODE -ne 0) { exit $LASTEXITCODE }

$python = Join-Path $RepoRoot ".venv\Scripts\python.exe"
& $python (Join-Path $PSScriptRoot "compare_golden.py") $OutDir
exit $LASTEXITCODE
