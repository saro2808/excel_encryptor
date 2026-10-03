@echo off
title Excel Anonymiser
cd /d "%~dp0"

echo =============================================
echo  Excel Anonymiser ^& Restorer
echo =============================================

if not exist ".venv" (
    echo First-time setup ^(this takes about a minute, please wait^)...
    echo.
    echo [1/2] Creating virtual environment...
    python -m venv .venv
    if errorlevel 1 (
        echo.
        echo ERROR: Python not found. Please install Python from https://www.python.org
        pause
        exit /b 1
    )
    echo [2/2] Downloading and installing packages ^(you will see them listed below^)...
    echo.
    .venv\Scripts\pip install -r requirements.txt
    if errorlevel 1 (
        echo.
        echo ERROR: Failed to install dependencies.
        pause
        exit /b 1
    )
    echo.
    echo Setup complete!
)

echo Starting app... a browser tab will open automatically.
echo Close this window to stop the app.
echo.
.venv\Scripts\streamlit run app.py
pause
