@echo off
setlocal
set "SCRIPT_DIR=%~dp0"
set "EXE=%SCRIPT_DIR%dist\AdientDashboardService.exe"
set "SVC=AdientDashboard"
set "LOGDIR=%SCRIPT_DIR%logs"
set "SVC_ACCOUNT=%~1"

if not exist "%EXE%" (
  echo EXE not found: %EXE%
  echo Run build_exe.ps1 first to create the executable.
  exit /b 1
)

rem Locate NSSM (place nssm.exe next to this script or install NSSM)
set "NSSM=%SCRIPT_DIR%nssm.exe"
if not exist "%NSSM%" set "NSSM=%ProgramFiles%\nssm\win64\nssm.exe"
if not exist "%NSSM%" set "NSSM=%ProgramFiles%\nssm\win32\nssm.exe"
if not exist "%NSSM%" set "NSSM=%ProgramFiles(x86)%\nssm\win64\nssm.exe"
if not exist "%NSSM%" set "NSSM=%ProgramFiles(x86)%\nssm\win32\nssm.exe"

if not exist "%NSSM%" (
  echo NSSM not found. Please install NSSM or place nssm.exe next to this script.
  echo Download: https://nssm.cc/download
  exit /b 1
)

"%NSSM%" stop "%SVC%" >nul 2>nul
"%NSSM%" remove "%SVC%" confirm >nul 2>nul

"%NSSM%" install "%SVC%" "%EXE%"
"%NSSM%" set "%SVC%" AppDirectory "%SCRIPT_DIR%"
"%NSSM%" set "%SVC%" Start SERVICE_AUTO_START
"%NSSM%" set "%SVC%" AppStopMethodSkip 6

rem Enable logging
if not exist "%LOGDIR%" mkdir "%LOGDIR%" >nul 2>nul
"%NSSM%" set "%SVC%" AppStdout "%LOGDIR%\service_stdout.log"
"%NSSM%" set "%SVC%" AppStderr "%LOGDIR%\service_stderr.log"
"%NSSM%" set "%SVC%" AppRotateFiles 1
"%NSSM%" set "%SVC%" AppRotateOnline 1
"%NSSM%" set "%SVC%" AppEnvironmentExtra "PYTHONUNBUFFERED=1"

rem Optional: run service under a specified account
if not "%SVC_ACCOUNT%"=="" (
  echo Configuring service to run as %SVC_ACCOUNT%
  for /f "usebackq delims=" %%P in (`powershell -NoProfile -Command "$p=Read-Host -AsSecureString 'Password for %SVC_ACCOUNT%';[Runtime.InteropServices.Marshal]::PtrToStringAuto([Runtime.InteropServices.Marshal]::SecureStringToBSTR($p))"`) do set "SVC_PASSWORD=%%P"
  if "%SVC_PASSWORD%"=="" (
    echo No password entered. Skipping account assignment.
  ) else (
    "%NSSM%" set "%SVC%" ObjectName "%SVC_ACCOUNT%" "%SVC_PASSWORD%"
  )
)

"%NSSM%" start "%SVC%"
echo Service %SVC% installed and started.
endlocal
