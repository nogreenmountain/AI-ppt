@echo off
setlocal EnableExtensions
chcp 65001 >nul

powershell -NoProfile -ExecutionPolicy Bypass -File "%~dp0scripts\setup_and_run.ps1" -SelfTest
set "EXIT_CODE=%ERRORLEVEL%"

if not "%SLIDE2PPTX_NO_PAUSE%"=="1" pause
exit /b %EXIT_CODE%
