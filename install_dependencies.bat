@echo off
cd /d "%~dp0"
echo Creating a private Python environment for this project...
where py >nul 2>nul
if %errorlevel%==0 (
  py -m venv .venv
) else (
  python -m venv .venv
)
if not exist ".venv\Scripts\python.exe" (
  echo Failed to create .venv. Please check whether Python is installed correctly.
  pause
  exit /b 1
)
echo Installing document-processing packages...
".venv\Scripts\python.exe" -m pip install -r requirements.txt
if errorlevel 1 (
  echo Installation failed. Please copy the error message and send it to Codex.
  pause
  exit /b 1
)
echo.
echo Installation completed successfully.
pause
