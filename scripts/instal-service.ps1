# install-service.ps1
param(
    [string]$ServiceAccount = "LocalSystem"
)

$ErrorActionPreference = "Stop"

# Configuration
$ServiceName = "AdientPVSService"
$DisplayName = "Adient PVS Dashboard"
$Description = "Displays production metrics dashboard on the TV"
$ExePath = Join-Path $PSScriptRoot "dist\AdientPVSService.exe"
$WorkingDir = Join-Path $PSScriptRoot "dist"
$LogDir = Join-Path $PSScriptRoot "logs"

# Check if running as admin
$isAdmin = ([Security.Principal.WindowsPrincipal][Security.Principal.WindowsIdentity]::GetCurrent()).IsInRole([Security.Principal.WindowsBuiltInRole]::Administrator)
if (-not $isAdmin) {
    Write-Host "This script must be run as Administrator" -ForegroundColor Red
    Write-Host "Please right-click the script and select 'Run as administrator'"
    exit 1
}

# Check if executable exists
if (-not (Test-Path $ExePath)) {
    Write-Host "Error: Executable not found at $ExePath" -ForegroundColor Red
    Write-Host "Please run build_pvs_exe.ps1 first to create the executable."
    exit 1
}

# Create logs directory if it doesn't exist
if (-not (Test-Path $LogDir)) {
    New-Item -ItemType Directory -Path $LogDir -Force | Out-Null
}

# Check if service exists and stop/remove it
$service = Get-Service -Name $ServiceName -ErrorAction SilentlyContinue
if ($service) {
    Write-Host "Stopping and removing existing service..." -ForegroundColor Yellow
    Stop-Service -Name $ServiceName -Force -ErrorAction SilentlyContinue
    sc.exe delete $ServiceName | Out-Null
    Start-Sleep -Seconds 2
}

# Create the service
Write-Host "`nInstalling $ServiceName service..." -ForegroundColor Cyan
$serviceParams = @{
    Name = $ServiceName
    DisplayName = $DisplayName
    Description = $Description
    BinaryPathName = "`"$ExePath`""
    StartupType = "Automatic"
}

# Create the service
$service = New-Service @serviceParams -ErrorAction Stop

# Configure service to auto-restart on failure
$service | Set-Service -StartupType Automatic -ErrorAction Stop
sc.exe failure $ServiceName reset= 86400 actions= restart/60000/restart/60000/restart/60000 | Out-Null

# Set service account
if ($ServiceAccount -eq "LocalSystem") {
    sc.exe config $ServiceName obj= "LocalSystem" | Out-Null
} else {
    sc.exe config $ServiceName obj= "$env:COMPUTERNAME\$ServiceAccount" password= "" | Out-Null
}

# Set working directory
$key = "HKLM:\SYSTEM\CurrentControlSet\Services\$ServiceName"
New-ItemProperty -Path $key -Name "WorkingDirectory" -Value $WorkingDir -PropertyType String -Force | Out-Null

# Start the service
Write-Host "Starting $ServiceName..." -ForegroundColor Cyan
Start-Service -Name $ServiceName -ErrorAction Stop

# Verify service is running
$service = Get-Service -Name $ServiceName
if ($service.Status -eq 'Running') {
    Write-Host "`n✅ Service installed and running successfully!" -ForegroundColor Green
    Write-Host "   Name: $($service.DisplayName)"
    Write-Host "   Status: $($service.Status)"
    Write-Host "   URL: http://localhost:5051"
} else {
    Write-Host "`n⚠️  Service installed but failed to start. Status: $($service.Status)" -ForegroundColor Yellow
    Write-Host "   Check the Windows Event Viewer for errors."
}

Write-Host "`nTo access the dashboard, open a browser and navigate to:"
Write-Host "   http://localhost:5051" -ForegroundColor Cyan
Write-Host "`nService account: $ServiceAccount" -ForegroundColor Gray