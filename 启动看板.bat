@echo off
chcp 65001 >nul 2>&1
cd /d "%~dp0"
cd tools
where python >nul 2>nul && set "PY=python"
if not defined PY ( where py >nul 2>nul && set "PY=py" )
if not defined PY ( echo Python not found. Please install Python 3 and add it to PATH. & pause & exit /b 1 )
"%PY%" generate_dashboard.py
echo Generating dashboard ...
echo Starting local server on port 8080 ...
start "" "%PY%" server.py
timeout /t 2 >nul
start "" "http://localhost:8080"
