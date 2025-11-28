# install_netlify_task.ps1
# Creates a Windows Scheduled Task to deploy PVS dashboard to Netlify daily at 08:45 CET

param(
    [string]$NetlifySiteId = "",
    [string]$NetlifyAuthToken = ""
)

$ErrorActionPreference = "Stop"

# Must run as admin
$isAdmin = ([Security.Principal.WindowsPrincipal][Security.Principal.WindowsIdentity]::GetCurrent()).IsInRole([Security.Principal.WindowsBuiltInRole]::Administrator)
if (-not $isAdmin) {
    Write-Host "ERROR: This script must be run as Administrator" -ForegroundColor Red
    exit 1
}

$TaskName = "PVS_Netlify_Deploy"
$ScriptDir = Split-Path -Parent $MyInvocation.MyCommand.Path
$ProjectRoot = Split-Path -Parent $ScriptDir
$DeployScript = Join-Path $ScriptDir "deploy_to_netlify.ps1"

Write-Host "Installing Netlify deployment scheduled task..." -ForegroundColor Cyan
Write-Host "Project root: $ProjectRoot"
Write-Host "Deploy script: $DeployScript"

# Remove existing task if present
$existing = Get-ScheduledTask -TaskName $TaskName -ErrorAction SilentlyContinue
if ($existing) {
    Write-Host "Removing existing task..."
    Unregister-ScheduledTask -TaskName $TaskName -Confirm:$false
}

# Build argument string
$argString = "-NoProfile -ExecutionPolicy Bypass -File `"$DeployScript`""
if ($NetlifySiteId) {
    $argString += " -SiteId `"$NetlifySiteId`""
}
if ($NetlifyAuthToken) {
    $argString += " -AuthToken `"$NetlifyAuthToken`""
}

# Create action
$action = New-ScheduledTaskAction `
    -Execute "powershell.exe" `
    -Argument $argString `
    -WorkingDirectory $ProjectRoot

# Create trigger: Monday-Friday at 08:45 (CET = UTC+1, but Windows uses local time)
$trigger = New-ScheduledTaskTrigger -Weekly -DaysOfWeek Monday,Tuesday,Wednesday,Thursday,Friday -At "08:45"

# Settings
$settings = New-ScheduledTaskSettingsSet `
    -AllowStartIfOnBatteries `
    -DontStopIfGoingOnBatteries `
    -StartWhenAvailable `
    -RunOnlyIfNetworkAvailable `
    -ExecutionTimeLimit (New-TimeSpan -Minutes 30)

# Principal (run as current user)
$principal = New-ScheduledTaskPrincipal -UserId $env:USERNAME -LogonType S4U -RunLevel Highest

# Register task
Register-ScheduledTask `
    -TaskName $TaskName `
    -Action $action `
    -Trigger $trigger `
    -Settings $settings `
    -Principal $principal `
    -Description "Daily deployment of PVS dashboard static snapshot to Netlify at 08:45 CET"

Write-Host ""
Write-Host "SUCCESS: Scheduled task '$TaskName' created!" -ForegroundColor Green
Write-Host ""
Write-Host "The task will run Monday-Friday at 08:45 local time."
Write-Host ""
Write-Host "To test manually, run:" -ForegroundColor Yellow
Write-Host "  Start-ScheduledTask -TaskName '$TaskName'"
Write-Host ""
Write-Host "To view task status:" -ForegroundColor Yellow
Write-Host "  Get-ScheduledTask -TaskName '$TaskName' | Get-ScheduledTaskInfo"
