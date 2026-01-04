@echo off
echo ================================================================
echo    Pattern-Based Academic Document Formatter
echo    Backend Server Startup Script (Windows)
echo ================================================================
echo.

cd /d "%~dp0"
cd backend

echo [1/4] Checking Python installation...
python --version >nul 2>&1
if errorlevel 1 (
    echo ERROR: Python is not installed or not in PATH
    echo Please install Python 3.8+ from https://www.python.org/downloads/
    pause
    exit /b 1
)

echo [2/4] Setting up virtual environment...
if not exist "venv" (
    echo Creating virtual environment...
    python -m venv venv
    if errorlevel 1 (
        echo ERROR: Failed to create virtual environment
        pause
        exit /b 1
    )
)

echo [3/4] Activating virtual environment and installing dependencies...
call venv\Scripts\activate.bat
pip install -r requirements.txt --quiet
if errorlevel 1 (
    echo ERROR: Failed to install dependencies
    pause
    exit /b 1
)

echo [4/4] Creating required directories...
if not exist "uploads" mkdir uploads
if not exist "outputs" mkdir outputs

echo.
echo ================================================================
echo    Server is starting...
echo    API will be available at: http://localhost:5000
echo    Health check: http://localhost:5000/health
echo    Press Ctrl+C to stop the server
echo ================================================================
echo.

python pattern_formatter_backend.py

pause
