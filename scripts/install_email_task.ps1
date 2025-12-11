# Install-EmailTask.ps1
# Creates a Windows Scheduled Task to send PVS email reports at 8:45 AM daily
#
# Run as Administrator:
#   powershell -ExecutionPolicy Bypass -File install_email_task.ps1

param(
    [string]$ServiceAccount = "",
    [switch]$Test,
    [switch]$Uninstall
)

$ErrorActionPreference = "Stop"

# Configuration
$TaskName = "PVS_Daily_Email"
$TaskDescription = "Sends PVS production report via email at 8:45 AM daily"
$ProjectDir = Split-Path -Parent (Split-Path -Parent $MyInvocation.MyCommand.Path)
$ScriptPath = Join-Path $ProjectDir "scripts\send_pvs_email_auto.py"
$LogDir = Join-Path $ProjectDir "logs"

# Ensure logs directory exists
if (-not (Test-Path $LogDir)) {
    New-Item -ItemType Directory -Path $LogDir -Force | Out-Null
}

# Check for admin rights
$isAdmin = ([Security.Principal.WindowsPrincipal][Security.Principal.WindowsIdentity]::GetCurrent()).IsInRole([Security.Principal.WindowsBuiltInRole]::Administrator)
if (-not $isAdmin) {
    Write-Host "ERROR: This script must be run as Administrator" -ForegroundColor Red
    Write-Host "Right-click PowerShell and select 'Run as Administrator'" -ForegroundColor Yellow
    exit 1
}

# Uninstall if requested
if ($Uninstall) {
    Write-Host "Removing scheduled task: $TaskName" -ForegroundColor Yellow
    Unregister-ScheduledTask -TaskName $TaskName -Confirm:$false -ErrorAction SilentlyContinue
    Write-Host "Task removed successfully" -ForegroundColor Green
    exit 0
}

Write-Host "======================================" -ForegroundColor Cyan
Write-Host " PVS Daily Email - Scheduled Task Setup" -ForegroundColor Cyan
Write-Host "======================================" -ForegroundColor Cyan
Write-Host ""

# Find Python executable
$PythonPath = ""
$VenvPython = Join-Path $ProjectDir ".venv\Scripts\python.exe"
if (Test-Path $VenvPython) {
    $PythonPath = $VenvPython
    Write-Host "Using virtual environment Python: $PythonPath" -ForegroundColor Green
} else {
    $PythonPath = (Get-Command python -ErrorAction SilentlyContinue).Source
    if (-not $PythonPath) {
        Write-Host "ERROR: Python not found. Please install Python or create a virtual environment." -ForegroundColor Red
        exit 1
    }
    Write-Host "Using system Python: $PythonPath" -ForegroundColor Yellow
}

# Verify script exists
if (-not (Test-Path $ScriptPath)) {
    Write-Host "ERROR: Script not found at $ScriptPath" -ForegroundColor Red
    exit 1
}

# Build command arguments
$Arguments = """$ScriptPath"""
if ($Test) {
    $Arguments = """$ScriptPath"" --test"
    Write-Host "TEST MODE: Will send to test recipient only" -ForegroundColor Yellow
}

# Create the action
$Action = New-ScheduledTaskAction -Execute $PythonPath -Argument $Arguments -WorkingDirectory $ProjectDir

# Create trigger for 8:45 AM daily (Monday-Friday)
$Trigger = New-ScheduledTaskTrigger -Daily -At "08:45"

# Create settings
$Settings = New-ScheduledTaskSettingsSet -AllowStartIfOnBatteries -DontStopIfGoingOnBatteries -StartWhenAvailable -RunOnlyIfNetworkAvailable -WakeToRun

# Determine principal (run as user or service account)
if ($ServiceAccount) {
    Write-Host ""
    Write-Host "Service Account: $ServiceAccount" -ForegroundColor Cyan
    $SecurePassword = Read-Host "Enter password for $ServiceAccount" -AsSecureString
    $BSTR = [System.Runtime.InteropServices.Marshal]::SecureStringToBSTR($SecurePassword)
    $Password = [System.Runtime.InteropServices.Marshal]::PtrToStringAuto($BSTR)
    
    $Principal = New-ScheduledTaskPrincipal -UserId $ServiceAccount -LogonType Password -RunLevel Highest
    
    # Register task with password
    Register-ScheduledTask -TaskName $TaskName -Action $Action -Trigger $Trigger -Settings $Settings -Principal $Principal -Description $TaskDescription -User $ServiceAccount -Password $Password -Force
} else {
    Write-Host ""
    Write-Host "No service account specified. Task will run as SYSTEM." -ForegroundColor Yellow
    Write-Host "To use a service account, run with: -ServiceAccount 'DOMAIN\username'" -ForegroundColor Yellow
    
    $Principal = New-ScheduledTaskPrincipal -UserId "SYSTEM" -LogonType ServiceAccount -RunLevel Highest
    
    # Register task as SYSTEM
    Register-ScheduledTask -TaskName $TaskName -Action $Action -Trigger $Trigger -Settings $Settings -Principal $Principal -Description $TaskDescription -Force
}

Write-Host ""
Write-Host "======================================" -ForegroundColor Green
Write-Host " Task Created Successfully!" -ForegroundColor Green
Write-Host "======================================" -ForegroundColor Green
Write-Host ""
Write-Host "Task Name:    $TaskName"
Write-Host "Schedule:     Daily at 8:45 AM"
Write-Host "Script:       $ScriptPath"
Write-Host "Python:       $PythonPath"
Write-Host ""
Write-Host "To test the task manually:" -ForegroundColor Cyan
Write-Host "  schtasks /run /tn ""$TaskName"""
Write-Host ""
Write-Host "To view task status:" -ForegroundColor Cyan
Write-Host "  schtasks /query /tn ""$TaskName"" /v"
Write-Host ""
Write-Host "To uninstall:" -ForegroundColor Cyan
Write-Host "  .\install_email_task.ps1 -Uninstall"
Write-Host ""
