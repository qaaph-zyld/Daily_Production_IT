@echo off
REM ============================================================================
REM Adient Production Dashboard - Direct Dependency Installation
REM This script installs dependencies directly (no virtual environment)
REM Better for network drives where venv creation hangs
REM ============================================================================

echo.
echo ============================================================================
echo   Adient Production Dashboard - Installing Dependencies
echo ============================================================================
echo.

REM Get the directory where this script is located
set SCRIPT_DIR=%~dp0
cd /d "%SCRIPT_DIR%"

echo [1/3] Checking Python installation...
python --version >nul 2>&1
if errorlevel 1 (
    echo ERROR: Python is not installed or not in PATH
    echo Please install Python 3.8 or higher from https://www.python.org/
    pause
    exit /b 1
)
python --version
echo.

echo [2/3] Upgrading pip...
python -m pip install --upgrade pip --proxy 104.129.196.38:10563 --user
if errorlevel 1 (
    echo WARNING: Failed with proxy, trying without proxy...
    python -m pip install --upgrade pip --user
)
echo.

echo [3/3] Installing dependencies...
echo Using proxy: 104.129.196.38:10563
pip install -r requirements.txt --proxy 104.129.196.38:10563 --user
if errorlevel 1 (
    echo WARNING: Installation with proxy failed, trying without proxy...
    pip install -r requirements.txt --user
    if errorlevel 1 (
        echo ERROR: Failed to install dependencies
        echo.
        echo Please check your internet connection and try again.
        pause
        exit /b 1
    )
)
echo.

echo ============================================================================
echo   Installation Complete!
echo ============================================================================
echo.
echo Installed packages:
pip list | findstr "Flask pyodbc dotenv"
echo.
echo ============================================================================
echo   Next Steps:
echo ============================================================================
echo.
echo 1. Run "start_dashboard_simple.bat" to launch the dashboard
echo 2. Open http://localhost:5000 in your web browser
echo.
echo ============================================================================
pause
