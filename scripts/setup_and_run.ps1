param(
  [Parameter(Position = 0)]
  [string]$InputImage,

  [Parameter(Position = 1)]
  [string]$OutDir,

  [switch]$Report,
  [switch]$SelfTest
)

$ErrorActionPreference = "Stop"
$RepoRoot = (Resolve-Path (Join-Path $PSScriptRoot "..")).Path
$VenvPython = Join-Path $RepoRoot ".venv\Scripts\python.exe"
$ArtifactRuntime = Join-Path $RepoRoot "artifact-runtime"

function Write-Step([string]$Message) {
  Write-Host ""
  Write-Host "==> $Message"
}

function Find-CommandPath([string[]]$Names) {
  foreach ($name in $Names) {
    $cmd = Get-Command $name -ErrorAction SilentlyContinue
    if ($cmd) { return $cmd.Source }
  }
  return $null
}

function Refresh-PathFromRegistry {
  $machine = [Environment]::GetEnvironmentVariable("Path", "Machine")
  $user = [Environment]::GetEnvironmentVariable("Path", "User")
  $env:Path = "$machine;$user"
}

function Install-WithWinget([string]$PackageId) {
  $winget = Find-CommandPath @("winget")
  if (-not $winget) { return $false }
  Write-Step "Installing $PackageId with winget"
  & $winget install --id $PackageId --exact --silent --accept-package-agreements --accept-source-agreements
  Refresh-PathFromRegistry
  return $true
}

function Ensure-Python {
  if (Test-Path $VenvPython) { return $VenvPython }

  $python = Find-CommandPath @("py", "python", "python3")
  if (-not $python) {
    if (-not (Install-WithWinget "Python.Python.3.12")) {
      throw "Python 3.10+ was not found and winget is unavailable. Install Python, then run this script again."
    }
    $python = Find-CommandPath @("py", "python", "python3")
  }
  if (-not $python) { throw "Python was installed but is not visible on PATH yet. Open a new terminal and run again." }

  Write-Step "Creating Python virtual environment"
  if ((Split-Path -Leaf $python) -ieq "py.exe") {
    & $python -3 -m venv (Join-Path $RepoRoot ".venv")
  } else {
    & $python -m venv (Join-Path $RepoRoot ".venv")
  }
  if ($LASTEXITCODE -ne 0) { throw "Failed to create Python virtual environment." }
  return $VenvPython
}

function Ensure-Node {
  $node = Find-CommandPath @("node")
  if (-not $node) {
    if (-not (Install-WithWinget "OpenJS.NodeJS.LTS")) {
      throw "Node.js 20+ was not found and winget is unavailable. Install Node.js LTS, then run this script again."
    }
    $node = Find-CommandPath @("node")
  }
  if (-not $node) { throw "Node.js was installed but is not visible on PATH yet. Open a new terminal and run again." }

  $major = [int](& $node -p "process.versions.node.split('.')[0]")
  if ($major -lt 20) {
    throw "Node.js 20+ is required. Found Node $(& $node -v). Install Node.js LTS and run again."
  }
  return $node
}

function Ensure-PythonDeps([string]$PythonExe) {
  Write-Step "Installing Python dependencies"
  & $PythonExe -m pip install --upgrade pip
  if ($LASTEXITCODE -ne 0) { throw "pip upgrade failed." }
  & $PythonExe -m pip install -r (Join-Path $RepoRoot "requirements.txt")
  if ($LASTEXITCODE -ne 0) { throw "Python dependency installation failed." }
}

function Ensure-NodeDeps {
  Write-Step "Installing Node dependencies"
  Push-Location $ArtifactRuntime
  try {
    $npm = Find-CommandPath @("npm")
    $pnpm = Find-CommandPath @("pnpm")
    if ($npm) {
      & $npm install
    } elseif ($pnpm) {
      & $pnpm install
    } else {
      throw "Neither npm nor pnpm was found after Node.js setup."
    }
    if ($LASTEXITCODE -ne 0) { throw "Node dependency installation failed." }
  } finally {
    Pop-Location
  }
}

function Run-SelfTest([string]$PythonExe, [string]$NodeExe) {
  Write-Step "Running Python tests"
  Push-Location $RepoRoot
  try {
    $env:PYTHONPATH = Join-Path $RepoRoot "python"
    & $PythonExe -m pytest
    if ($LASTEXITCODE -ne 0) { throw "Python tests failed." }

    Write-Step "Running JS tests"
    Push-Location $ArtifactRuntime
    try {
      & $NodeExe "test/run.mjs"
      if ($LASTEXITCODE -ne 0) { throw "JS tests failed." }
      & $NodeExe "src/convert.mjs" --self-test
      if ($LASTEXITCODE -ne 0) { throw "JS self-test failed." }
    } finally {
      Pop-Location
    }
  } finally {
    Pop-Location
  }
}

$PythonExe = Ensure-Python
$NodeExe = Ensure-Node
Ensure-PythonDeps $PythonExe
Ensure-NodeDeps

if ($SelfTest) {
  Run-SelfTest $PythonExe $NodeExe
  Write-Host ""
  Write-Host "[OK] setup_and_test completed."
  exit 0
}

if ([string]::IsNullOrWhiteSpace($InputImage)) {
  $InputImage = Join-Path $RepoRoot "samples\source.png"
  Write-Host "[INFO] No input image supplied; using sample: $InputImage"
}
if (-not (Test-Path $InputImage)) {
  throw "Input image not found: $InputImage"
}
if ([string]::IsNullOrWhiteSpace($OutDir)) {
  $stamp = Get-Date -Format "yyyyMMdd-HHmmss"
  $OutDir = Join-Path $RepoRoot "outputs\one-click-$stamp"
}

$env:PYTHONPATH = Join-Path $RepoRoot "python"
$env:SLIDE2PPTX_NODE = $NodeExe

Write-Step "Converting image to PPTX"
$pipelineArgs = @(
  "-m", "slide2pptx.pipeline_cli",
  $InputImage,
  "--out", $OutDir,
  "--visual-passes", "2",
  "--second-pass-max-components", "96"
)
if (-not $Report) { $pipelineArgs += "--skip-report" }
& $PythonExe @pipelineArgs
if ($LASTEXITCODE -ne 0) { throw "Conversion failed with exit code $LASTEXITCODE." }

Write-Host ""
Write-Host "[OK] Done."
Write-Host "PPTX: $OutDir\build\reconstructed.pptx"
Write-Host "Detection JSON: $OutDir\detect\detected.json"
if ($Report) {
  Write-Host "Report: $OutDir\report\report.html"
} else {
  Write-Host "Tip: run convert_image_to_ppt.bat <image> <out_dir> -Report to also render the HTML report on Windows with PowerPoint installed."
}
