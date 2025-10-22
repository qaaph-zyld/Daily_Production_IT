$ErrorActionPreference = "Stop"
$proxy = "http://104.129.196.38:10563"

Write-Host "`n============================================================="
Write-Host "  Adient Dashboard - Local VENV Bootstrap & Start"
Write-Host "=============================================================\n"

# Move to script directory
$scriptDir = Split-Path -Parent $MyInvocation.MyCommand.Path
Set-Location $scriptDir

# 0) Kill any stuck python/pip
Write-Host "[0/7] Stopping any running python/pip..."
Get-Process python -ErrorAction SilentlyContinue | Stop-Process -Force -ErrorAction SilentlyContinue
Get-Process pip -ErrorAction SilentlyContinue | Stop-Process -Force -ErrorAction SilentlyContinue

# 1) Configure proxy for current session and pip.ini
Write-Host "[1/7] Configuring proxy and pip..."
$env:HTTP_PROXY = $proxy
$env:HTTPS_PROXY = $proxy
[Environment]::SetEnvironmentVariable("PIP_DISABLE_PIP_VERSION_CHECK","1","User")
[Environment]::SetEnvironmentVariable("PIP_DEFAULT_TIMEOUT","180","User")
[Environment]::SetEnvironmentVariable("HTTP_PROXY", $proxy, "User")
[Environment]::SetEnvironmentVariable("HTTPS_PROXY", $proxy, "User")

$pipDir = Join-Path $env:APPDATA "pip"
New-Item -ItemType Directory -Force -Path $pipDir | Out-Null
$pipIni = @"
[global]
proxy = $proxy
trusted-host = pypi.org files.pythonhosted.org pypi.python.org
timeout = 180
disable-pip-version-check = yes
"@
$pipIni | Set-Content -Path (Join-Path $pipDir "pip.ini") -Encoding Ascii

# 2) Pick a LOCAL Python interpreter (avoid network paths)
Write-Host "[2/7] Selecting local Python interpreter..."
$candidates = @(
  "$env:LOCALAPPDATA\Programs\Python\Python313\python.exe",
  "$env:LOCALAPPDATA\Programs\Python\Python312\python.exe",
  "C:\\Program Files\\Python313\\python.exe",
  "C:\\Program Files\\Python312\\python.exe",
  "C:\\Python313\\python.exe",
  "C:\\Python312\\python.exe"
)
$python = $null
foreach ($p in $candidates) {
  if (Test-Path $p) { $python = $p; break }
}
if (-not $python) {
  # Fall back to default 'python' (may be on network)
  $python = "python"
}
Write-Host ("Using Python: " + (& $python -c "import sys; print(sys.executable)"))

# 3) Create LOCAL venv (not on the network drive)
Write-Host "[3/7] Creating local virtual environment..."
$venvRoot = Join-Path $env:LOCALAPPDATA "AdientDashboard"
$venvPath = Join-Path $venvRoot "venv"
New-Item -ItemType Directory -Force -Path $venvRoot | Out-Null
if (Test-Path $venvPath) {
  Write-Host "Removing existing local venv..."
  Remove-Item -Recurse -Force -Path $venvPath
}

# Using --without-pip to avoid copying from network; install pip via ensurepip afterwards
& $python --version
& $python -m venv "$venvPath" --without-pip

# 4) Install pip into venv (offline, uses bundled wheels)
Write-Host "[4/7] Bootstrapping pip inside venv (offline ensurepip)..."
& "$venvPath\Scripts\python.exe" -m ensurepip --upgrade

# 5) Upgrade pip and install requirements with proxy
Write-Host "[5/7] Upgrading pip (via proxy)..."
& "$venvPath\Scripts\python.exe" -m pip install --upgrade pip --proxy $proxy --trusted-host pypi.org --trusted-host files.pythonhosted.org --trusted-host pypi.pythonhosted.org
if ($LASTEXITCODE -ne 0) { Write-Error "Failed to upgrade pip. Check proxy connectivity."; exit 1 }

Write-Host "[6/7] Installing requirements.txt (via proxy)..."
& "$venvPath\Scripts\python.exe" -m pip install -r "$scriptDir\requirements.txt" --proxy $proxy --trusted-host pypi.org --trusted-host files.pythonhosted.org --trusted-host pypi.pythonhosted.org --no-cache-dir
if ($LASTEXITCODE -ne 0) { Write-Error "Failed to install requirements. Check proxy or pip.ini."; exit 1 }

# 6b) Download static frontend assets (Chart.js + Datalabels)
Write-Host "[6b] Ensuring static frontend assets are present (via proxy)..."
$staticDir = Join-Path $scriptDir "static"
New-Item -ItemType Directory -Force -Path $staticDir | Out-Null

$assets = @(
  @{Url='https://cdn.jsdelivr.net/npm/chart.js@4.4.0/dist/chart.umd.min.js'; File='chart.umd.min.js'},
  @{Url='https://cdn.jsdelivr.net/npm/chartjs-plugin-datalabels@2.2.0/dist/chartjs-plugin-datalabels.min.js'; File='chartjs-plugin-datalabels.min.js'}
)
foreach ($a in $assets) {
  $outFile = Join-Path $staticDir $a.File
  try {
    Write-Host ("  - Downloading " + $a.File)
    Invoke-WebRequest -Uri $a.Url -OutFile $outFile -Proxy $proxy -UseBasicParsing -TimeoutSec 120
  } catch {
    Write-Warning ("  Failed to download " + $a.File + ": " + $_.Exception.Message)
  }
}

# 7) Start server
Write-Host "[7/7] Starting dashboard server using local venv...\n"
$env:FLASK_PORT = (Select-String -Path "$scriptDir\.env" -Pattern '^FLASK_PORT=(.+)$').Matches.Groups[1].Value
if (-not $env:FLASK_PORT) { $env:FLASK_PORT = "5000" }

& "$venvPath\Scripts\python.exe" "$scriptDir\dashboard_server.py"
