@echo off
setlocal
cd /d "%~dp0"

if not exist "venv\Scripts\python.exe" (
  echo Creating Python virtual environment...
  py -3 -m venv venv
  if errorlevel 1 (
    echo Could not create the environment. Install Python 3.11 or newer and try again.
    pause
    exit /b 1
  )
)

echo Installing or checking backend dependencies...
venv\Scripts\python.exe -m pip install --upgrade pip
venv\Scripts\python.exe -m pip install -r requirements.txt
if errorlevel 1 (
  echo Dependency installation failed. Check internet access and Python version.
  pause
  exit /b 1
)

echo Starting Presentation Proficiency API on port 8011...
venv\Scripts\python.exe -m uvicorn app:app --host 127.0.0.1 --port 8011
pause
