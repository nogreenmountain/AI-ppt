@echo off
setlocal EnableExtensions
chcp 65001 >nul

powershell -NoProfile -ExecutionPolicy Bypass -File "%~dp0scripts\build_windows_app.ps1" -Install
set "EXIT_CODE=%ERRORLEVEL%"

if not "%SLIDE2PPTX_NO_PAUSE%"=="1" pause
exit /b %EXIT_CODE%
