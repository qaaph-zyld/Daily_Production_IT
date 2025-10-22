@echo off
REM ============================================================================
REM Adient Production Dashboard - Launcher Script (Local VENV aware)
REM Prefers per-user local venv at %LOCALAPPDATA%\AdientDashboard\venv
REM Falls back to running bootstrap_local_venv_and_start.ps1
REM ============================================================================

echo.
echo ============================================================================
echo   Adient Production Dashboard - Starting Server
echo ============================================================================
echo.

set SCRIPT_DIR=%~dp0
cd /d "%SCRIPT_DIR%"

set LOCAL_VENV_PY=%LOCALAPPDATA%\AdientDashboard\venv\Scripts\python.exe

if exist "%LOCALAPPDATA%\AdientDashboard\venv\Scripts\python.exe" (
  echo Using local virtual environment: %LOCALAPPDATA%\AdientDashboard\venv
  echo.
  "%LOCALAPPDATA%\AdientDashboard\venv\Scripts\python.exe" "%SCRIPT_DIR%\dashboard_server.py"
  goto :eof
)

echo Local virtual environment not found. Bootstrapping now...
echo.
powershell -NoProfile -ExecutionPolicy Bypass -File .\bootstrap_local_venv_and_start.ps1

REM If the server stops, keep window open
echo.
echo ============================================================================
echo   Server stopped
echo ============================================================================
pause
