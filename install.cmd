@echo off
rem NeonCenterline - installer.
rem 1) NeonCenterline.gms -> CorelDRAW user GMS folder(s) of the CURRENT user (no admin rights needed).
rem 2) Toolbar add-on -> CorelDRAW Addons folder(s) in Program Files (asks for administrator rights).
setlocal
set "NEON_INST=%~f0"
set "HERE=%~dp0"
if /I "%~1"=="addon" goto addon

set FOUND=0
for /d %%D in ("%APPDATA%\Corel\*") do (
  if exist "%%D\Draw\GMS" (
    copy /Y "%HERE%NeonCenterline.gms" "%%D\Draw\GMS\" >nul && echo GMS   -^> %%D\Draw\GMS && set FOUND=1
  )
)
if "%FOUND%"=="0" echo CorelDRAW GMS folder not found. Copy NeonCenterline.gms manually (see README).

echo Installing the toolbar (administrator rights)...
powershell -NoProfile -Command "Start-Process -FilePath $env:ComSpec -ArgumentList '/c', ('\"' + $env:NEON_INST + '\" addon') -Verb RunAs -Wait" 2>nul
goto done

:addon
for /d %%D in ("%ProgramFiles%\Corel\*") do (
  if exist "%%D\Programs64\Addons" (
    xcopy /E /I /Y "%HERE%addon\NeonCenterline" "%%D\Programs64\Addons\NeonCenterline" >nul && echo Addon -^> %%D\Programs64\Addons\NeonCenterline
  )
)
timeout /t 3 >nul
exit /b

:done
echo.
echo Done. Close CorelDRAW (if running) and start it again.
pause
