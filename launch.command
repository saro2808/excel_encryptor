#!/bin/bash
# Mac launcher – double-click this file in Finder to start the app.
# (If macOS asks "Are you sure?", click Open.)

cd "$(dirname "$0")"

echo "============================================="
echo " Excel Anonymiser & Restorer"
echo "============================================="

if [ ! -d ".venv" ]; then
    echo "First-time setup (this takes about a minute, please wait)..."
    echo ""
    echo "[1/2] Creating virtual environment..."
    python3 -m venv .venv
    if [ $? -ne 0 ]; then
        echo ""
        echo "ERROR: Python 3 not found."
        echo "Please install it from https://www.python.org"
        read -p "Press Enter to close..."
        exit 1
    fi
    echo "[2/2] Downloading and installing packages (you will see them listed below)..."
    echo ""
    .venv/bin/pip install -r requirements.txt
    if [ $? -ne 0 ]; then
        echo ""
        echo "ERROR: Failed to install dependencies."
        read -p "Press Enter to close..."
        exit 1
    fi
    echo ""
    echo "Setup complete!"
fi

echo "Starting app... a browser tab will open automatically."
echo "Close this window to stop the app."
echo ""
.venv/bin/streamlit run app.py
