@echo off
cd /d "%~dp0"
if not exist ".venv\Scripts\python.exe" (
  echo The project environment is missing. Please run install_dependencies.bat first.
  pause
  exit /b 1
)
echo This test will call DeepSeek once for each readable public sample.
echo Please keep this window open until the summary appears.
echo.
".venv\Scripts\python.exe" tests\evaluate_with_deepseek.py
echo.
pause
