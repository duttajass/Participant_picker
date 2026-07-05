@echo off
setlocal

title Teams Picker Launcher
cd /d "%~dp0"

echo ==========================================
echo      Teams Picker - Windows Launcher
echo ==========================================
echo.

where node >nul 2>nul
if errorlevel 1 (
  echo [ERROR] Node.js is not installed or not in PATH.
  echo Install Node.js 18+ from: https://nodejs.org
  pause
  exit /b 1
)

where npm >nul 2>nul
if errorlevel 1 (
  echo [ERROR] npm is not available in PATH.
  pause
  exit /b 1
)

if not exist "node_modules" (
  echo [SETUP] Installing npm dependencies...
  call npm install
  if errorlevel 1 (
    echo [ERROR] npm install failed.
    pause
    exit /b 1
  )
)

if not exist ".playwright-installed" (
  echo [SETUP] Installing Playwright Chromium browser...
  call npm run install:browsers
  if errorlevel 1 (
    echo [WARN] Browser install step returned an error.
    echo        You can retry manually: npm run install:browsers
  ) else (
    type nul > .playwright-installed
  )
)

echo [INFO] Starting Teams Picker server...
echo [INFO] Opening app in browser after startup...
start "" cmd /c "timeout /t 5 /nobreak >nul && start http://localhost:3000"

call npm start

echo.
echo Server stopped.
pause
exit /b 0
