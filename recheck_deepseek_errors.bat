@echo off
cd /d "%~dp0"
if not exist ".venv\Scripts\python.exe" (
  echo The project environment is missing. Please run install_dependencies.bat first.
  pause
  exit /b 1
)
echo Rechecking only the cases that were wrong in the previous evaluation.
echo.
".venv\Scripts\python.exe" tests\recheck_deepseek_errors.py
echo.
pause
