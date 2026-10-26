@echo off
setlocal
cd /d "%~dp0"
if exist "%~dp0.venv\Scripts\pythonw.exe" (
  start "" "%~dp0.venv\Scripts\pythonw.exe" "%~dp0AutoFMScrapper.py" --ui
  exit /b 0
)
call "%~dp0Setup AutoFMScrapper.cmd"
if errorlevel 1 exit /b 1
start "" "%~dp0.venv\Scripts\pythonw.exe" "%~dp0AutoFMScrapper.py" --ui
