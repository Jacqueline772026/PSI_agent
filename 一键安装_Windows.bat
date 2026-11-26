@echo off
REM PSI agent one-click setup (Windows). Double-click to run.
cd /d "%~dp0"
powershell -NoProfile -ExecutionPolicy Bypass -File "%~dp0scripts\setup_windows.ps1"
echo.
pause
