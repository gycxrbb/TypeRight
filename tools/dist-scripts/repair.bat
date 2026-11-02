@echo off
rem TypeRight repair wrapper: use when the TypeRight tab disappears
cd /d "%~dp0"
powershell -NoProfile -ExecutionPolicy Bypass -File "dist-scripts\repair-core.ps1"
pause
