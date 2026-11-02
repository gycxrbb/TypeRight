@echo off
rem TypeRight installer wrapper (ASCII-only; all logic+Chinese in ps1)
cd /d "%~dp0"
powershell -NoProfile -ExecutionPolicy Bypass -File "dist-scripts\install-core.ps1"
pause
