@echo off
rem NeonCenterline - installer. Copies NeonCenterline.gms into the CorelDRAW user GMS folder(s)
rem and the toolbar add-on into CorelDRAW Addons folder(s) (needs administrator rights).
setlocal
set "HERE=%~dp0"
net session >nul 2>&1
if errorlevel 1 (
  echo Requesting administrator rights for the Addons folder...
  powershell -NoProfile -Command "Start-Process -FilePath '%~f0' -Verb RunAs"
  exit /b
)
set FOUND=0
for /d %%D in ("%APPDATA%\Corel\*") do (
  if exist "%%D\Draw\GMS" (
    copy /Y "%HERE%NeonCenterline.gms" "%%D\Draw\GMS\" >nul && echo GMS   -^> %%D\Draw\GMS && set FOUND=1
  )
)
for /d %%D in ("%ProgramFiles%\Corel\*") do (
  if exist "%%D\Programs64\Addons" (
    xcopy /E /I /Y "%HERE%addon\NeonCenterline" "%%D\Programs64\Addons\NeonCenterline" >nul && echo Addon -^> %%D\Programs64\Addons\NeonCenterline
  )
)
if "%FOUND%"=="0" echo CorelDRAW GMS folder not found. Copy NeonCenterline.gms manually (see README).
echo.
echo Done. Close CorelDRAW (if running) and start it again.
pause
