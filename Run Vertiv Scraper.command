#!/bin/zsh
# GXTManager launcher for macOS
# Double-click this file in Finder to run. If macOS blocks it, right-click and choose Open.

cd "$(dirname "$0")"

if ! command -v python3 &>/dev/null; then
    echo ""
    echo "Python 3 is not installed."
    echo "Download it from https://www.python.org/downloads/ and try again."
    echo ""
    read -r -p "Press Enter to close..."
    exit 1
fi

if [ ! -d ".venv" ]; then
    echo "Setting up for first run..."
    python3 -m venv .venv 2>/dev/null
fi

# Install requirements silently -- -q -q suppresses warnings, 2>/dev/null drops the rest
.venv/bin/pip install -q -q -r requirements.txt 2>/dev/null

echo "Starting GXTManager..."

# Filter out the noisy macOS window-state restore message before launching
.venv/bin/python3 vertiv_battery_scraper.py 2> >(grep -v "ApplePersistenceIgnoreState" >&2)
