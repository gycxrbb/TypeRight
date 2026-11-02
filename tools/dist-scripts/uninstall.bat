@echo off
rem TypeRight uninstall wrapper
cd /d "%~dp0"
powershell -NoProfile -ExecutionPolicy Bypass -File "dist-scripts\uninstall-core.ps1"
pause
