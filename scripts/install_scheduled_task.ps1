<#
.SYNOPSIS
    Creates a Windows Scheduled Task to refresh the Excel file at 08:45 CET, Monday-Friday.

.DESCRIPTION
    This script creates a scheduled task that:
    - Runs at 08:45 AM local time (CET)
    - Only on weekdays (Monday-Friday)
    - Executes refresh_excel_scheduled.ps1
    - Runs under the specified service account (or current user if not specified)

.PARAMETER ServiceAccount
    Optional. The service account to run the task under (e.g., "DOMAIN\ServiceAccount").
    If not specified, runs under the current user.

.EXAMPLE
    .\install_scheduled_task.ps1
    .\install_scheduled_task.ps1 -ServiceAccount "ADIENT\svc_pvs_dashboard"

.NOTES
    Requires Administrator privileges to create scheduled tasks.
#>

param(
    [string]$ServiceAccount = ""
)

$ErrorActionPreference = "Stop"

# ============================================================================
# CONFIGURATION
# ============================================================================
$TaskName = "AdientPVS_ExcelRefresh"
$TaskDescription = "Refreshes PVS Planned_qtys.xlsx at 08:45 AM CET, Monday-Friday"
$ScriptDir = Split-Path -Parent $MyInvocation.MyCommand.Path
$RefreshScript = Join-Path $ScriptDir "refresh_excel_scheduled.ps1"
$ProjectRoot = Split-Path -Parent $ScriptDir

# ============================================================================
# VALIDATION
# ============================================================================
if (-not (Test-Path $RefreshScript)) {
    Write-Error "Refresh script not found: $RefreshScript"
    exit 1
}

Write-Host "=============================================================" -ForegroundColor Cyan
Write-Host "  Adient PVS - Scheduled Task Installer" -ForegroundColor Cyan
Write-Host "=============================================================" -ForegroundColor Cyan
Write-Host ""
Write-Host "Task Name: $TaskName"
Write-Host "Script: $RefreshScript"
Write-Host "Schedule: 08:45 AM, Monday-Friday"
Write-Host ""

# ============================================================================
# REMOVE EXISTING TASK
# ============================================================================
$ExistingTask = Get-ScheduledTask -TaskName $TaskName -ErrorAction SilentlyContinue
if ($ExistingTask) {
    Write-Host "Removing existing task..." -ForegroundColor Yellow
    Unregister-ScheduledTask -TaskName $TaskName -Confirm:$false
}

# ============================================================================
# CREATE TASK ACTION
# ============================================================================
$Action = New-ScheduledTaskAction `
    -Execute "powershell.exe" `
    -Argument "-NoProfile -ExecutionPolicy Bypass -File `"$RefreshScript`"" `
    -WorkingDirectory $ProjectRoot

# ============================================================================
# CREATE TRIGGER (08:45 AM, Monday-Friday)
# ============================================================================
$Trigger = New-ScheduledTaskTrigger `
    -Weekly `
    -DaysOfWeek Monday,Tuesday,Wednesday,Thursday,Friday `
    -At "08:45"

# ============================================================================
# CREATE SETTINGS
# ============================================================================
$Settings = New-ScheduledTaskSettingsSet `
    -AllowStartIfOnBatteries `
    -DontStopIfGoingOnBatteries `
    -StartWhenAvailable `
    -ExecutionTimeLimit (New-TimeSpan -Minutes 30) `
    -RestartCount 3 `
    -RestartInterval (New-TimeSpan -Minutes 1)

# ============================================================================
# CREATE PRINCIPAL (USER CONTEXT)
# ============================================================================
if ($ServiceAccount) {
    Write-Host "Configuring to run as: $ServiceAccount" -ForegroundColor Yellow
    Write-Host "You will be prompted for the password..." -ForegroundColor Yellow
    
    $Principal = New-ScheduledTaskPrincipal `
        -UserId $ServiceAccount `
        -LogonType Password `
        -RunLevel Highest
    
    # Get password securely
    $SecurePassword = Read-Host -AsSecureString "Enter password for $ServiceAccount"
    $BSTR = [System.Runtime.InteropServices.Marshal]::SecureStringToBSTR($SecurePassword)
    $Password = [System.Runtime.InteropServices.Marshal]::PtrToStringAuto($BSTR)
    [System.Runtime.InteropServices.Marshal]::ZeroFreeBSTR($BSTR)
    
    Register-ScheduledTask `
        -TaskName $TaskName `
        -Description $TaskDescription `
        -Action $Action `
        -Trigger $Trigger `
        -Settings $Settings `
        -User $ServiceAccount `
        -Password $Password `
        -Force
} else {
    Write-Host "Configuring to run as current user (interactive)" -ForegroundColor Yellow
    
    $Principal = New-ScheduledTaskPrincipal `
        -UserId $env:USERNAME `
        -LogonType Interactive `
        -RunLevel Highest
    
    Register-ScheduledTask `
        -TaskName $TaskName `
        -Description $TaskDescription `
        -Action $Action `
        -Trigger $Trigger `
        -Settings $Settings `
        -Principal $Principal `
        -Force
}

# ============================================================================
# VERIFY
# ============================================================================
$Task = Get-ScheduledTask -TaskName $TaskName -ErrorAction SilentlyContinue
if ($Task) {
    Write-Host ""
    Write-Host "SUCCESS: Scheduled task created!" -ForegroundColor Green
    Write-Host ""
    Write-Host "Task Details:" -ForegroundColor Cyan
    Write-Host "  Name: $TaskName"
    Write-Host "  Status: $($Task.State)"
    Write-Host "  Next Run: Check Task Scheduler for next run time"
    Write-Host ""
    Write-Host "To test the task manually, run:" -ForegroundColor Yellow
    Write-Host "  Start-ScheduledTask -TaskName '$TaskName'"
    Write-Host ""
} else {
    Write-Error "Failed to create scheduled task"
    exit 1
}
