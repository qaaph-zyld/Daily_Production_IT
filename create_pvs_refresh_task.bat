@echo off
setlocal
set SCRIPT_DIR=%~dp0
set "PS=%SystemRoot%\System32\WindowsPowerShell\v1.0\powershell.exe"
set "TASK_CMD=\"%PS%\" -NoProfile -ExecutionPolicy Bypass -File \"%SCRIPT_DIR%refresh_pvs_daily.ps1\""

schtasks /Create /TN "AdientPVSRefreshDaily8" /SC WEEKLY /D MON,TUE,WED,THU,FRI /ST 08:00 /TR "%TASK_CMD" /F

echo Scheduled task 'AdientPVSRefreshDaily8' created for weekdays at 08:00.
endlocal
