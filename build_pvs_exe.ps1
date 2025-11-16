$ErrorActionPreference = "Stop"
Write-Host "`n=== Build Adient PVS EXE ===`n"

$scriptDir = Split-Path -Parent $MyInvocation.MyCommand.Path
Set-Location $scriptDir

$py = Join-Path $scriptDir ".venv/Scenes/python.exe"
if (-not (Test-Path $py)) { $py = Join-Path $scriptDir ".venv/Scripts/python.exe" }
if (-not (Test-Path $py)) { $py = Join-Path $env:LOCALAPPDATA "AdientDashboard/venv/Scripts/python.exe" }
if (-not (Test-Path $py)) { $py = "python" }

& $py -m pip install --upgrade pip > $null
if (Test-Path (Join-Path $scriptDir "requirements_pvs.txt")) {
  & $py -m pip install -r (Join-Path $scriptDir "requirements_pvs.txt") > $null
} elseif (Test-Path (Join-Path $scriptDir "requirements.txt")) {
  & $py -m pip install -r (Join-Path $scriptDir "requirements.txt") > $null
}
& $py -m pip install pyinstaller > $null

if (Test-Path "$scriptDir/dist") { Remove-Item -Recurse -Force "$scriptDir/dist" }
if (Test-Path "$scriptDir/build") { Remove-Item -Recurse -Force "$scriptDir/build" }

$addDataTpl = "templates;templates"

& $py -m PyInstaller `
  --noconfirm `
  --clean `
  --onefile `
  --name "AdientPVSService" `
  --add-data $addDataTpl `
  --hidden-import pyodbc `
  --hidden-import dotenv `
  "$scriptDir/pvs_service_launcher.py"

# Ensure runtime config files are available next to EXE
if (-not (Test-Path (Join-Path $scriptDir 'dist'))) { New-Item -ItemType Directory -Force -Path (Join-Path $scriptDir 'dist') | Out-Null }
if (Test-Path (Join-Path $scriptDir ".env")) {
  Copy-Item -Force (Join-Path $scriptDir ".env") (Join-Path $scriptDir "dist\.env")
}
# Copy PVS planned CSV
$dstPvs = Join-Path $scriptDir 'dist/PVS'
New-Item -ItemType Directory -Force -Path $dstPvs | Out-Null
$srcCsv = Join-Path $scriptDir 'PVS/Planned_qtys.csv'
if (Test-Path $srcCsv) { Copy-Item -Force $srcCsv (Join-Path $dstPvs 'Planned_qtys.csv') }
# Copy PVS planned XLSX (primary source)
$srcXlsx = Join-Path $scriptDir 'PVS/Planned_qtys.xlsx'
if (Test-Path $srcXlsx) { Copy-Item -Force $srcXlsx (Join-Path $dstPvs 'Planned_qtys.xlsx') }
 # Copy PVS mapping CSV
 $srcMap = Join-Path $scriptDir 'PVS/ProdLine_Project_Map.csv'
 if (Test-Path $srcMap) { Copy-Item -Force $srcMap (Join-Path $dstPvs 'ProdLine_Project_Map.csv') }

Write-Host "`nBuild finished. EXE: $scriptDir\dist\AdientPVSService.exe`n"
