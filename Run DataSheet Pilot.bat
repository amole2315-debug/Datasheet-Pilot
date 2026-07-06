@echo off
cd /d "%~dp0"
where node >nul 2>nul
if errorlevel 1 (
  echo Node.js was not found. Please install Node.js or run a local web server in this folder.
  pause
  exit /b 1
)
start "" "http://127.0.0.1:8787/"
node "%~dp0pilot_server.js"
pause
