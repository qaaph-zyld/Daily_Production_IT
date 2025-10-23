@echo off
setlocal
set "SCRIPT_DIR=%~dp0"
set "EXE=%SCRIPT_DIR%dist\AdientDashboardService.exe"
set "SVC=AdientDashboard"

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

"%NSSM%" start "%SVC%"
echo Service %SVC% installed and started.
endlocal
