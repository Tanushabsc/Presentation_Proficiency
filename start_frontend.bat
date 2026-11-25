@echo off
setlocal
cd /d "%~dp0"
if not exist node_modules (
  echo Installing frontend dependencies. This requires internet access...
  call npm install
  if errorlevel 1 (
    echo Frontend dependency installation failed.
    pause
    exit /b 1
  )
)
echo Starting frontend...
call npm run dev
pause
