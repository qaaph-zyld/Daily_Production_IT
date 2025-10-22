@echo off
REM ============================================================================
REM Adient Production Dashboard - Virtual Environment Setup Script
REM This script creates a virtual environment and installs all dependencies
REM Supports proxy configuration for corporate firewalls
REM ============================================================================

echo.
echo ============================================================================
echo   Adient Production Dashboard - Virtual Environment Setup
echo ============================================================================
echo.

REM Get the directory where this script is located
set SCRIPT_DIR=%~dp0
cd /d "%SCRIPT_DIR%"

echo [1/5] Checking Python installation...
python --version >nul 2>&1
if errorlevel 1 (
    echo ERROR: Python is not installed or not in PATH
    echo Please install Python 3.8 or higher from https://www.python.org/
    pause
    exit /b 1
)
python --version
echo.

echo [2/5] Creating virtual environment...
if exist "venv" (
    echo Virtual environment already exists. Removing old one...
    rmdir /s /q venv
)
echo Creating virtual environment (this may take 1-2 minutes on network drives)...
python -m venv venv --without-pip
if errorlevel 1 (
    echo ERROR: Failed to create virtual environment
    pause
    exit /b 1
)
echo Virtual environment created successfully!
echo.

echo [3/5] Activating virtual environment...
call venv\Scripts\activate.bat
if errorlevel 1 (
    echo ERROR: Failed to activate virtual environment
    pause
    exit /b 1
)
echo Virtual environment activated!
echo.

echo [4/5] Installing pip (bootstrapping)...
echo Downloading get-pip.py...
curl https://bootstrap.pypa.io/get-pip.py -o get-pip.py --proxy 104.129.196.38:10563
if errorlevel 1 (
    echo WARNING: Failed with proxy, trying without proxy...
    curl https://bootstrap.pypa.io/get-pip.py -o get-pip.py
)
python get-pip.py --proxy 104.129.196.38:10563
if errorlevel 1 (
    echo WARNING: Failed with proxy, trying without proxy...
    python get-pip.py
)
del get-pip.py
echo Pip installed successfully!
echo.

echo [5/5] Installing dependencies from requirements.txt...
echo Using proxy: 104.129.196.38:10563
pip install -r requirements.txt --proxy 104.129.196.38:10563
if errorlevel 1 (
    echo WARNING: Installation with proxy failed, trying without proxy...
    pip install -r requirements.txt
    if errorlevel 1 (
        echo ERROR: Failed to install dependencies
        echo.
        echo Please check your internet connection and try again.
        echo If behind a corporate firewall, ensure proxy settings are correct.
        pause
        exit /b 1
    )
)
echo.

echo ============================================================================
echo   Setup Complete!
echo ============================================================================
echo.
echo Virtual environment created at: %SCRIPT_DIR%venv
echo.
echo Installed packages:
pip list
echo.
echo ============================================================================
echo   Next Steps:
echo ============================================================================
echo.
echo 1. Review the .env file and ensure database credentials are correct
echo 2. Run "start_dashboard.bat" to launch the dashboard
echo 3. Open http://localhost:5000 in your web browser
echo.
echo ============================================================================
pause
