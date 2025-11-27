$ErrorActionPreference = "Stop"
Write-Host "`n=== Build Adient Dashboard EXE ===`n"

$scriptDir = Split-Path -Parent $MyInvocation.MyCommand.Path
Set-Location $scriptDir

# Choose python (prefer local venv)
$py = Join-Path $env:LOCALAPPDATA "AdientDashboard/venv/Scripts/python.exe"
if (-not (Test-Path $py)) { $py = "python" }

# Ensure dependencies
& $py -m pip install --upgrade pip > $null
if (Test-Path (Join-Path $scriptDir "requirements.txt")) {
  & $py -m pip install -r (Join-Path $scriptDir "requirements.txt") > $null
}
& $py -m pip install pyinstaller > $null

# Clean old build
if (Test-Path "$scriptDir/dist") { Remove-Item -Recurse -Force "$scriptDir/dist" }
if (Test-Path "$scriptDir/build") { Remove-Item -Recurse -Force "$scriptDir/build" }

# Build onefile EXE with bundled templates/static
$addDataTpl = "templates;templates"
$addDataSta = "static;static"

& $py -m PyInstaller `
  --noconfirm `
  --clean `
  --onefile `
  --name "AdientDashboardService" `
  --add-data $addDataTpl `
  --add-data $addDataSta `
  --hidden-import pyodbc `
  --hidden-import dotenv `
  "$scriptDir/service_launcher.py"

Write-Host "`nBuild finished. EXE: $scriptDir\dist\AdientDashboardService.exe`n"
if (Test-Path (Join-Path $scriptDir ".env")) {
  Copy-Item -Force (Join-Path $scriptDir ".env") (Join-Path $scriptDir "dist\.env")
}
