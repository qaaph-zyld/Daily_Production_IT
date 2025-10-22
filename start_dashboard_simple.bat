@echo off
REM ============================================================================
REM Adient Production Dashboard - Simple Launcher (No Virtual Environment)
REM This script starts the dashboard using system Python
REM ============================================================================

echo.
echo ============================================================================
echo   Adient Production Dashboard - Starting Server
echo ============================================================================
echo.

REM Get the directory where this script is located
set SCRIPT_DIR=%~dp0
cd /d "%SCRIPT_DIR%"

echo [1/2] Checking Python installation...
python --version >nul 2>&1
if errorlevel 1 (
    echo ERROR: Python is not installed or not in PATH
    pause
    exit /b 1
)
python --version
echo.

echo [2/2] Starting dashboard server...
echo.
echo ============================================================================
python dashboard_server.py

REM If server stops, keep window open
echo.
echo ============================================================================
echo   Server stopped
echo ============================================================================
pause
