@echo off
setlocal
set "SCRIPT_DIR=%~dp0"
schtasks /Create /TN "AdientDashboardServer" /SC ONLOGON /TR "%SCRIPT_DIR%start_dashboard.bat" /F
endlocal
