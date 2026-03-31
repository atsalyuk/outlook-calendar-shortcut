@echo off
where python >nul 2>&1
if %errorlevel% neq 0 (
    echo Python is not installed or not in PATH.
    echo Download from https://www.python.org/downloads/
    pause
    exit /b 1
)
echo Creating virtual environment...
python -m venv .venv
echo Installing dependencies...
.venv\Scripts\pip install -r requirements.txt
echo.
echo Done! Run the app with: run.bat
pause
