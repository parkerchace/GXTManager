#!/bin/zsh
# GXTManager launcher for macOS
# Double-click this file in Finder to run. If macOS blocks it, right-click and choose Open.

cd "$(dirname "$0")"

# Make sure Python 3 is available
if ! command -v python3 &>/dev/null; then
    echo ""
    echo "Python 3 was not found on this computer."
    echo "Please download and install it from https://www.python.org/downloads/"
    echo "Then try running this again."
    echo ""
    read -r -p "Press Enter to close..."
    exit 1
fi

# Create a virtual environment on first run so dependencies stay isolated
if [ ! -d ".venv" ]; then
    echo "First run -- setting up a virtual environment..."
    python3 -m venv .venv
    echo "Done."
fi

# Install or update required packages
echo "Checking requirements..."
.venv/bin/pip install -q --upgrade pip
.venv/bin/pip install -q -r requirements.txt
echo "All good."
echo ""

# Launch the app
.venv/bin/python3 vertiv_battery_scraper.py
