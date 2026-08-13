@echo off
cd /d "%~dp0"
where node >nul 2>nul
if errorlevel 1 (
  echo Node.js was not found. Please install Node.js or run a local web server in this folder.
  pause
  exit /b 1
)
start "" "http://127.0.0.1:8787/"
start "DataSheet Pilot Server" /min cmd /c "node ""%~dp0pilot_server.js"" >> ""%~dp0pilot_server.log"" 2>&1"
exit /b
