@echo off
setlocal EnableExtensions
chcp 65001 >nul

set "PROJECT_DIR=%~dp0"
set "BUNDLED_PY=C:\Users\tangvx\.cache\codex-runtimes\codex-primary-runtime\dependencies\python\python.exe"
set "BUNDLED_NODE=C:\Users\tangvx\.cache\codex-runtimes\codex-primary-runtime\dependencies\node\bin\node.exe"

if "%~1"=="" (
  echo.
  echo Usage:
  echo   convert_image_to_ppt.bat "input_image_path" ["output_dir"]
  echo.
  echo Example:
  echo   convert_image_to_ppt.bat "C:\Users\tangvx\Desktop\slide.png"
  echo   convert_image_to_ppt.bat "C:\Users\tangvx\Desktop\slide.png" "C:\Users\tangvx\Desktop\out"
  echo.
  if not "%SLIDE2PPTX_NO_PAUSE%"=="1" pause
  exit /b 2
)

set "INPUT_IMAGE=%~1"
if not exist "%INPUT_IMAGE%" (
  echo [ERROR] Input image not found: "%INPUT_IMAGE%"
  if not "%SLIDE2PPTX_NO_PAUSE%"=="1" pause
  exit /b 10
)

if "%~2"=="" (
  for /f %%I in ('powershell -NoProfile -Command "Get-Date -Format yyyyMMdd-HHmmss"') do set "STAMP=%%I"
  set "OUT_DIR=%PROJECT_DIR%outputs\bat-run-%STAMP%"
) else (
  set "OUT_DIR=%~2"
)

if exist "%BUNDLED_PY%" (
  set "PY_EXE=%BUNDLED_PY%"
) else (
  set "PY_EXE=python"
  echo [WARN] Bundled Python 3.12 was not found. Falling back to python on PATH.
  echo [WARN] If OCR text extraction fails, use Python 3.12 with rapidocr_onnxruntime installed.
)

if exist "%BUNDLED_NODE%" (
  set "SLIDE2PPTX_NODE=%BUNDLED_NODE%"
)

pushd "%PROJECT_DIR%" >nul
set "PYTHONPATH=%PROJECT_DIR%python"

echo [INFO] Input image: "%INPUT_IMAGE%"
echo [INFO] Output dir: "%OUT_DIR%"
echo [INFO] Running two-pass image-to-PPT conversion...

"%PY_EXE%" -m slide2pptx.pipeline_cli "%INPUT_IMAGE%" --out "%OUT_DIR%" --visual-passes 2 --second-pass-max-components 96 --skip-report
set "EXIT_CODE=%ERRORLEVEL%"

if "%EXIT_CODE%"=="0" (
  echo.
  echo [OK] Done:
  echo   PPTX: "%OUT_DIR%\build\reconstructed.pptx"
  echo   Preview: "%OUT_DIR%\build\artifact-preview.png"
  echo   Detection JSON: "%OUT_DIR%\detect\detected.json"
) else (
  echo.
  echo [ERROR] Conversion failed. Exit code: %EXIT_CODE%
)

popd >nul
if not "%SLIDE2PPTX_NO_PAUSE%"=="1" pause
exit /b %EXIT_CODE%
