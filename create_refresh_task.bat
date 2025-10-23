@echo off
setlocal
set SCRIPT_DIR=%~dp0
set "PS=%SystemRoot%\System32\WindowsPowerShell\v1.0\powershell.exe"
set "TASK_CMD=\"%PS%\" -NoProfile -ExecutionPolicy Bypass -File \"%SCRIPT_DIR%refresh_ping.ps1\""
schtasks /Create /TN "AdientDashboardRefresh" /SC MINUTE /MO 15 /TR "%TASK_CMD%" /F
endlocal
