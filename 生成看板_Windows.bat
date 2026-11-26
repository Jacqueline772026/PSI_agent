@echo off
REM Build the PSI dashboard (Windows). Double-click to run.
cd /d "%~dp0"
powershell -NoProfile -ExecutionPolicy Bypass -File "%~dp0scripts\dashboard_windows.ps1"
echo.
pause
