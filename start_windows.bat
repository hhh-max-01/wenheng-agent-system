@echo off
cd /d "%~dp0"
if not exist ".venv\Scripts\python.exe" (
  echo The project environment is missing. Please run install_dependencies.bat first.
  pause
  exit /b 1
)
".venv\Scripts\python.exe" server.py
pause
