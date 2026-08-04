param(
  [switch]$Install
)

$ErrorActionPreference = "Stop"
$RepoRoot = (Resolve-Path (Join-Path $PSScriptRoot "..")).Path
$VenvPython = Join-Path $RepoRoot ".venv\Scripts\python.exe"
$ArtifactRuntime = Join-Path $RepoRoot "artifact-runtime"
$PackageRoot = Join-Path $RepoRoot "build\desktop-runtime"
$PackagedArtifactRuntime = Join-Path $PackageRoot "artifact-runtime"
$AppName = -join @(
  [char]0x0041, [char]0x0049, [char]0x0020,
  [char]0x0050, [char]0x0050, [char]0x0054, [char]0x0020,
  [char]0x62C6, [char]0x9875, [char]0x5668
)
$DistApp = Join-Path $RepoRoot ("dist\" + $AppName)

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

Write-Step "Prepare dependencies"
$env:SLIDE2PPTX_NO_PAUSE = "1"
& (Join-Path $PSScriptRoot "setup_and_run.ps1") -SelfTest
if ($LASTEXITCODE -ne 0) { throw "Setup/self-test failed." }

$NodeExe = Find-CommandPath @("node")
if (-not $NodeExe) {
  throw "Node.js is required to build the offline desktop app."
}

Write-Step "Install PyInstaller"
& $VenvPython -m pip install pyinstaller
if ($LASTEXITCODE -ne 0) { throw "PyInstaller installation failed." }

Write-Step "Prepare packaged Node runtime"
if (Test-Path $PackageRoot) {
  $resolvedPackageRoot = (Resolve-Path $PackageRoot).Path
  if (-not $resolvedPackageRoot.StartsWith((Join-Path $RepoRoot "build"))) {
    throw "Refusing to clean unexpected package path: $resolvedPackageRoot"
  }
  Remove-Item -LiteralPath $resolvedPackageRoot -Recurse -Force
}
New-Item -ItemType Directory -Force -Path $PackagedArtifactRuntime | Out-Null
Copy-Item -LiteralPath (Join-Path $ArtifactRuntime "package.json") -Destination $PackagedArtifactRuntime
Copy-Item -LiteralPath (Join-Path $ArtifactRuntime "pnpm-lock.yaml") -Destination $PackagedArtifactRuntime
Copy-Item -LiteralPath (Join-Path $ArtifactRuntime "src") -Destination $PackagedArtifactRuntime -Recurse
$pnpm = Find-CommandPath @("pnpm", "pnpm.cmd")
if (-not $pnpm) {
  throw "pnpm is required to prepare a self-contained Node dependency tree."
}
Push-Location $PackagedArtifactRuntime
try {
  $env:CI = "true"
  & $pnpm install --prod --node-linker=hoisted
  if ($LASTEXITCODE -ne 0) { throw "Packaged Node dependency installation failed." }
} finally {
  Pop-Location
}

Write-Step "Build desktop app"
$addArtifact = "$PackagedArtifactRuntime;artifact-runtime"
$addPython = "$(Join-Path $RepoRoot 'python');python"
$addNode = "$NodeExe;runtime\node"
$guiEntry = Join-Path $RepoRoot "python\slide2pptx\gui_app.py"

Push-Location $RepoRoot
try {
  & $VenvPython -m PyInstaller `
    --noconfirm `
    --clean `
    --onedir `
    --windowed `
    --name $AppName `
    --paths (Join-Path $RepoRoot "python") `
    --collect-all rapidocr_onnxruntime `
    --collect-all onnxruntime `
    --add-data $addArtifact `
    --add-data $addPython `
    --add-binary $addNode `
    $guiEntry
  if ($LASTEXITCODE -ne 0) { throw "PyInstaller build failed." }
} finally {
  Pop-Location
}

Write-Host ""
Write-Host ("[OK] Built: " + (Join-Path $DistApp ($AppName + ".exe")))

if ($Install) {
  $InstallDir = Join-Path $env:LOCALAPPDATA ("Programs\" + $AppName)
  Write-Step ("Install to " + $InstallDir)
  New-Item -ItemType Directory -Force -Path $InstallDir | Out-Null
  Copy-Item -LiteralPath (Join-Path $DistApp "*") -Destination $InstallDir -Recurse -Force

  $Exe = Join-Path $InstallDir ($AppName + ".exe")
  $DesktopShortcut = Join-Path ([Environment]::GetFolderPath("Desktop")) ($AppName + ".lnk")
  $StartMenuDir = Join-Path $env:APPDATA "Microsoft\Windows\Start Menu\Programs"
  $StartShortcut = Join-Path $StartMenuDir ($AppName + ".lnk")
  $ShortcutDescription = -join @(
    [char]0x672C, [char]0x5730, [char]0x56FE, [char]0x7247,
    [char]0x3001, [char]0x0050, [char]0x0050, [char]0x0054,
    [char]0x3001, [char]0x0050, [char]0x0050, [char]0x0054,
    [char]0x0058, [char]0x0020, [char]0x62C6, [char]0x9875,
    [char]0x5DE5, [char]0x5177
  )

  $shell = New-Object -ComObject WScript.Shell
  foreach ($shortcutPath in @($DesktopShortcut, $StartShortcut)) {
    $shortcut = $shell.CreateShortcut($shortcutPath)
    $shortcut.TargetPath = $Exe
    $shortcut.WorkingDirectory = $InstallDir
    $shortcut.Description = $ShortcutDescription
    $shortcut.Save()
  }

  Write-Host ""
  Write-Host ("[OK] Installed: " + $Exe)
  Write-Host "Shortcuts: Desktop and Start Menu"
}
