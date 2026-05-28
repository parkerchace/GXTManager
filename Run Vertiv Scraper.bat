@echo off
REM GXTManager launcher for Windows
REM Double-click this file to run.

cd /d "%~dp0"

REM Check that Python is installed and on PATH
python --version >nul 2>&1
if errorlevel 1 (
    echo.
    echo Python was not found on this computer.
    echo Please download and install it from https://www.python.org/downloads/
    echo.
    echo IMPORTANT: On the installer's first screen, check the box that says
    echo "Add Python to PATH" before clicking Install Now.
    echo.
    pause
    exit /b 1
)

REM Create a virtual environment on first run so dependencies stay isolated
if not exist ".venv" (
    echo First run -- setting up a virtual environment...
    python -m venv .venv
    echo Done.
)

REM Install or update required packages
echo Checking requirements...
.venv\Scripts\pip install -q --upgrade pip
.venv\Scripts\pip install -q -r requirements.txt
echo All good.
echo.

REM Launch the app
.venv\Scripts\python vertiv_battery_scraper.py
