@echo off
rem NeonCenterline - installer. All logic is in install.ps1 (finds the newest CorelDRAW by version).
powershell -NoProfile -ExecutionPolicy Bypass -File "%~dp0install.ps1"
echo.
pause
