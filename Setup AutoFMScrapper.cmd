@echo off
setlocal
cd /d "%~dp0"
if not exist "%~dp0.venv\Scripts\python.exe" (
  where py >nul 2>&1
  if not errorlevel 1 (
    py -3 -m venv "%~dp0.venv"
  ) else (
    python -m venv "%~dp0.venv"
  )
  if errorlevel 1 (
    echo Install Python for Windows, then run this file again.
    pause
    exit /b 1
  )
)
"%~dp0.venv\Scripts\python.exe" -m pip install -r "%~dp0requirements.txt"
if errorlevel 1 (
  echo Setup could not finish. Check the internet connection and run this file again.
  pause
  exit /b 1
)
echo AutoFMScrapper is ready. Open AutoFMScrapper.cmd starts the app.
