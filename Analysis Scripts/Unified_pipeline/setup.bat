@echo off
rem One-time setup on Windows (re-run after a git pull that changes uv.lock or
rem R_reports\renv.lock):
rem   1. Python env + Quarto from uv.lock   (uv sync)
rem   2. R packages from R_reports\renv.lock (renv::restore)
setlocal EnableDelayedExpansion
cd /d "%~dp0"

set "UV=uv"
where uv >nul 2>nul
if errorlevel 1 (
  if exist "%USERPROFILE%\.local\bin\uv.exe" (
    set "UV=%USERPROFILE%\.local\bin\uv.exe"
  ) else (
    echo uv is not installed. In PowerShell, run:
    echo   powershell -ExecutionPolicy ByPass -c "irm https://astral.sh/uv/install.ps1 | iex"
    echo then open a new window and run setup.bat again.
    pause
    exit /b 1
  )
)

echo == Python env and Quarto (uv sync) ==
"%UV%" sync
if errorlevel 1 goto :failed

rem Find Rscript: PATH, then machine-wide and per-user installs (newest last wins)
set "RSCRIPT="
for /f "delims=" %%r in ('where Rscript 2^>nul') do if not defined RSCRIPT set "RSCRIPT=%%r"
if not defined RSCRIPT (
  for /d %%d in ("%ProgramFiles%\R\R-*" "%LOCALAPPDATA%\Programs\R\R-*") do (
    if exist "%%~d\bin\Rscript.exe" set "RSCRIPT=%%~d\bin\Rscript.exe"
  )
)
if not defined RSCRIPT (
  echo.
  echo R is not installed, so the reports can't run ^(the analyses still work^).
  echo Install R 4.5 from https://cran.r-project.org/bin/windows/base/ and run setup.bat again.
  pause
  exit /b 1
)
echo Using !RSCRIPT!

echo.
echo == R packages (renv::restore, first time takes a few minutes) ==
pushd R_reports
"!RSCRIPT!" -e "renv::restore(prompt = FALSE)"
if errorlevel 1 (popd & goto :failed)
"!RSCRIPT!" -e "renv::status()"
popd

echo.
echo Setup done. Start the pipeline by double-clicking launch_pipeline.bat.
pause
exit /b 0

:failed
echo.
echo Setup failed; see the messages above.
pause
exit /b 1
