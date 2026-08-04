param(
  [Parameter(Position = 0)]
  [string]$InputFile,

  [Parameter(Position = 1)]
  [string]$OutDir,

  [switch]$Report,
  [switch]$SelfTest
)

$ErrorActionPreference = "Stop"
$RepoRoot = (Resolve-Path (Join-Path $PSScriptRoot "..")).Path
$VenvPython = Join-Path $RepoRoot ".venv\Scripts\python.exe"
$ArtifactRuntime = Join-Path $RepoRoot "artifact-runtime"
$RequiredPythonMajor = 3
$RequiredPythonMinor = 12

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

function Add-ToPathIfNeeded([string]$Dir) {
  if ([string]::IsNullOrWhiteSpace($Dir)) { return }
  if (-not ($env:Path -split ';' | Where-Object { $_ -ieq $Dir })) {
    $env:Path = "$Dir;$env:Path"
  }
}

function Refresh-PathFromRegistry {
  $machine = [Environment]::GetEnvironmentVariable("Path", "Machine")
  $user = [Environment]::GetEnvironmentVariable("Path", "User")
  $env:Path = "$machine;$user"
}

function Install-WithWinget([string]$PackageId) {
  $winget = Find-CommandPath @("winget")
  if (-not $winget) { return $false }
  Write-Step "Install $PackageId with winget"
  & $winget install --id $PackageId --exact --silent --accept-package-agreements --accept-source-agreements
  Refresh-PathFromRegistry
  return $true
}

function Get-PythonVersionTag([string]$PythonExe, [string[]]$PythonArgs = @()) {
  try {
    $version = & $PythonExe @PythonArgs -c "import sys; print(f'{sys.version_info[0]}.{sys.version_info[1]}.{sys.version_info[2]}')"
    if ($LASTEXITCODE -ne 0) { return $null }
    return ($version | Select-Object -First 1).Trim()
  } catch {
    return $null
  }
}

function Test-RequiredPython([string]$PythonExe, [string[]]$PythonArgs = @()) {
  try {
    & $PythonExe @PythonArgs -c "import sys; raise SystemExit(0 if sys.version_info[:2] == ($RequiredPythonMajor, $RequiredPythonMinor) else 1)" | Out-Null
    return $LASTEXITCODE -eq 0
  } catch {
    return $false
  }
}

function Move-AsideExistingVenv {
  $venvDir = Join-Path $RepoRoot ".venv"
  if (-not (Test-Path $venvDir)) { return }

  $version = Get-PythonVersionTag $VenvPython
  if ([string]::IsNullOrWhiteSpace($version)) { $version = "unknown" }
  $stamp = Get-Date -Format "yyyyMMdd-HHmmss"
  $archive = Join-Path $RepoRoot (".venv-py{0}-{1}" -f ($version -replace '\.', ''), $stamp)
  Write-Step "Archive incompatible .venv ($version) -> $(Split-Path -Leaf $archive)"
  Move-Item -LiteralPath $venvDir -Destination $archive
}

function New-VenvWithPython([string]$PythonExe, [string[]]$PythonArgs = @()) {
  Write-Step "Create Python $RequiredPythonMajor.$RequiredPythonMinor virtual environment"
  & $PythonExe @PythonArgs -m venv (Join-Path $RepoRoot ".venv")
  if ($LASTEXITCODE -ne 0) { throw "Failed to create Python virtual environment." }
  return $VenvPython
}

function Resolve-Python312Candidate {
  $override = $env:SLIDE2PPTX_PYTHON
  if (-not [string]::IsNullOrWhiteSpace($override)) {
    $overridePath = Resolve-Path $override -ErrorAction SilentlyContinue
    if (-not $overridePath) {
      throw "SLIDE2PPTX_PYTHON points to a missing file: $override"
    }
    if (-not (Test-RequiredPython $overridePath.Path)) {
      $found = Get-PythonVersionTag $overridePath.Path
      throw "SLIDE2PPTX_PYTHON must point to Python 3.12. Found: $found"
    }
    return @{ Exe = $overridePath.Path; Args = @() }
  }

  $candidateFiles = @(
    (Join-Path $env:LOCALAPPDATA "Programs\Python\Python312\python.exe"),
    (Join-Path $env:USERPROFILE ".cache\codex-runtimes\codex-primary-runtime\dependencies\python\python.exe")
  )
  foreach ($candidate in $candidateFiles) {
    if ((Test-Path $candidate) -and (Test-RequiredPython $candidate)) {
      return @{ Exe = $candidate; Args = @() }
    }
  }

  $pyLauncher = Find-CommandPath @("py")
  if ($pyLauncher -and (Test-RequiredPython $pyLauncher @("-3.12"))) {
    return @{ Exe = $pyLauncher; Args = @("-3.12") }
  }

  foreach ($name in @("python3.12", "python", "python3")) {
    $cmd = Find-CommandPath @($name)
    if ($cmd -and (Test-RequiredPython $cmd)) {
      return @{ Exe = $cmd; Args = @() }
    }
  }

  return $null
}

function Ensure-Python {
  if (Test-Path $VenvPython) {
    if (Test-RequiredPython $VenvPython) { return $VenvPython }
    Move-AsideExistingVenv
  }

  $python = Resolve-Python312Candidate
  if (-not $python) {
    if (-not (Install-WithWinget "Python.Python.3.12")) {
      throw "Python 3.12 was not found and winget is unavailable. Install Python 3.12, set SLIDE2PPTX_PYTHON, then run this script again."
    }
    $python = Resolve-Python312Candidate
  }
  if (-not $python) { throw "Python 3.12 was installed but is not visible yet. Open a new terminal and run again." }

  return New-VenvWithPython $python.Exe $python.Args
}

function Ensure-Node {
  $override = $env:SLIDE2PPTX_NODE
  if (-not [string]::IsNullOrWhiteSpace($override)) {
    $overridePath = Resolve-Path $override -ErrorAction SilentlyContinue
    if (-not $overridePath) {
      throw "SLIDE2PPTX_NODE points to a missing file: $override"
    }
    $node = $overridePath.Path
    $nodeDir = Split-Path -Parent $node
    Add-ToPathIfNeeded $nodeDir
  }
  $node = Find-CommandPath @("node")
  if (-not $node) {
    $bundledNode = Join-Path $env:USERPROFILE ".cache\codex-runtimes\codex-primary-runtime\dependencies\node\bin\node.exe"
    if (Test-Path $bundledNode) {
      $node = $bundledNode
      Add-ToPathIfNeeded (Split-Path -Parent $node)
      Add-ToPathIfNeeded (Join-Path $env:USERPROFILE ".cache\codex-runtimes\codex-primary-runtime\dependencies\bin\fallback")
    }
  }
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
  Write-Step "Install Python dependencies"
  & $PythonExe -m pip install --upgrade pip
  if ($LASTEXITCODE -ne 0) { throw "pip upgrade failed." }
  & $PythonExe -m pip install -r (Join-Path $RepoRoot "requirements.txt")
  if ($LASTEXITCODE -ne 0) { throw "Python dependency installation failed." }
}

function Ensure-NodeDeps {
  Write-Step "Install Node dependencies"
  Push-Location $ArtifactRuntime
  try {
    $pnpm = Find-CommandPath @("pnpm", "pnpm.cmd", "pnpm.ps1")
    $npm = Find-CommandPath @("npm", "npm.cmd")
    if ($pnpm) {
      & $pnpm install --frozen-lockfile
    } elseif ($npm) {
      & $npm exec --yes --package pnpm@11.13.1 -- pnpm install --frozen-lockfile
    } else {
      throw "Neither pnpm nor npm was found after Node.js setup."
    }
    if ($LASTEXITCODE -ne 0) { throw "Node dependency installation failed." }
  } finally {
    Pop-Location
  }
}

function Run-SelfTest([string]$PythonExe, [string]$NodeExe) {
  Write-Step "Run Python tests"
  Push-Location $RepoRoot
  try {
    $env:PYTHONPATH = Join-Path $RepoRoot "python"
    & $PythonExe -m pytest
    if ($LASTEXITCODE -ne 0) { throw "Python tests failed." }

    Write-Step "Run JavaScript tests"
    Push-Location $ArtifactRuntime
    try {
      & $NodeExe "test/run.mjs"
      if ($LASTEXITCODE -ne 0) { throw "JavaScript tests failed." }
      & $NodeExe "src/convert.mjs" --self-test
      if ($LASTEXITCODE -ne 0) { throw "JavaScript self-test failed." }
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

if ([string]::IsNullOrWhiteSpace($InputFile)) {
  $goldenSample = Join-Path $RepoRoot "samples\industry-teaching-research.png"
  $InputFile = if (Test-Path $goldenSample) { $goldenSample } else { Join-Path $RepoRoot "samples\source.png" }
  Write-Host "[INFO] No input file supplied; using sample: $InputFile"
}
if (-not (Test-Path $InputFile)) {
  throw "Input file not found: $InputFile"
}
if ([string]::IsNullOrWhiteSpace($OutDir)) {
  $stamp = Get-Date -Format "yyyyMMdd-HHmmss"
  $OutDir = Join-Path $RepoRoot "outputs\one-click-$stamp"
}

$env:PYTHONPATH = Join-Path $RepoRoot "python"
$env:SLIDE2PPTX_NODE = $NodeExe

Write-Step "Split input to editable PPTX"
$pipelineArgs = @(
  "-m", "slide2pptx.pipeline_cli",
  $InputFile,
  "--out", $OutDir,
  "--visual-passes", "2",
  "--second-pass-max-components", "96"
)
if (-not $Report) { $pipelineArgs += "--skip-report" }
& $PythonExe @pipelineArgs
if ($LASTEXITCODE -ne 0) { throw "Conversion failed with exit code $LASTEXITCODE." }

Write-Host ""
Write-Host "[OK] Done."
Write-Host "Output: $OutDir"
Write-Host "Image input PPTX: $OutDir\build\reconstructed.pptx"
Write-Host "PPT/PPTX page outputs: $OutDir\slide-001, slide-002, ..."
if ($Report) {
  Write-Host "Report: $OutDir\report\report.html"
} else {
  Write-Host "Tip: add -Report to generate an HTML comparison report when PowerPoint is installed."
}
