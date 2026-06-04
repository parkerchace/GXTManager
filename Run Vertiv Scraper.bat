@echo off
REM GXTManager launcher for Windows
REM Double-click this file to run.

cd /d "%~dp0"

python --version >nul 2>&1
if errorlevel 1 (
    echo.
    echo Python was not found on this computer.
    echo Download it from https://www.python.org/downloads/
    echo.
    echo IMPORTANT: On the installer's first screen, check the box that says
    echo "Add Python to PATH" before clicking Install Now.
    echo.
    pause
    exit /b 1
)

if not exist ".venv" (
    echo Setting up for first run...
    python -m venv .venv
    if errorlevel 1 (
        echo.
        echo Failed to create a virtual environment.
        echo Make sure Python installed correctly and try again.
        echo.
        pause
        exit /b 1
    )
)

.venv\Scripts\pip install -q -q -r requirements.txt 2>nul

echo Starting GXTManager...
echo.

.venv\Scripts\python vertiv_battery_scraper.py

if errorlevel 1 (
    echo.
    echo GXTManager closed with an error.
    pause
)
