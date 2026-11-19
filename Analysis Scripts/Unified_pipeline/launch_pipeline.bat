@echo off
rem Start the EMT Pipeline (Master Analysis Launcher) in the project's uv env.
rem Double-click this file. First-time setup: setup.bat
setlocal
cd /d "%~dp0"

set "UV=uv"
where uv >nul 2>nul
if errorlevel 1 (
  if exist "%USERPROFILE%\.local\bin\uv.exe" (
    set "UV=%USERPROFILE%\.local\bin\uv.exe"
  ) else (
    echo Could not find uv. Install it ^(see README.md^), then run setup.bat.
    pause
    exit /b 1
  )
)

"%UV%" run python Python_analyze\Master_Analysis_Launcher.py %*
if errorlevel 1 (
  echo.
  echo The launcher exited with an error ^(see above^).
  pause
)
