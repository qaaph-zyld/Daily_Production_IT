$ErrorActionPreference = "Stop"
$proxy = $null

Write-Host "`n============================================================="
Write-Host "  Adient Dashboard - Local VENV Bootstrap & Start"
Write-Host "=============================================================\n"

# Move to script directory
$scriptDir = Split-Path -Parent $MyInvocation.MyCommand.Path
Set-Location $scriptDir

# Load optional proxy from .env (PROXY=...), defaulting to none
$envFile = Join-Path $scriptDir ".env"
if (Test-Path $envFile) {
  $m = Select-String -Path $envFile -Pattern '^\s*PROXY\s*=\s*(.+)\s*$' -ErrorAction SilentlyContinue
  if ($m) { $proxy = $m.Matches.Groups[1].Value.Trim() }
}
if ([string]::IsNullOrWhiteSpace($proxy)) { $proxy = $null }

# 0) Kill any stuck python/pip
Write-Host "[0/7] Stopping any running python/pip..."
Get-Process python -ErrorAction SilentlyContinue | Stop-Process -Force -ErrorAction SilentlyContinue
Get-Process pip -ErrorAction SilentlyContinue | Stop-Process -Force -ErrorAction SilentlyContinue

# 1) Configure pip (proxy optional)
Write-Host "[1/7] Configuring pip (proxy optional)..."
[Environment]::SetEnvironmentVariable("PIP_DISABLE_PIP_VERSION_CHECK","1","User")
[Environment]::SetEnvironmentVariable("PIP_DEFAULT_TIMEOUT","180","User")
if ($proxy) {
  $env:HTTP_PROXY = $proxy
  $env:HTTPS_PROXY = $proxy
  [Environment]::SetEnvironmentVariable("HTTP_PROXY", $proxy, "User")
  [Environment]::SetEnvironmentVariable("HTTPS_PROXY", $proxy, "User")
}

$pipDir = Join-Path $env:APPDATA "pip"
New-Item -ItemType Directory -Force -Path $pipDir | Out-Null
$pipIniLines = @(
  "[global]",
  "trusted-host = pypi.org files.pythonhosted.org pypi.python.org",
  "timeout = 180",
  "disable-pip-version-check = yes"
)
if ($proxy) { $pipIniLines += "proxy = $proxy" }
$pipIniLines -join "`r`n" | Set-Content -Path (Join-Path $pipDir "pip.ini") -Encoding Ascii

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

# 5) Upgrade pip and install requirements (proxy optional)
Write-Host "[5/7] Upgrading pip..."
$pipArgs = @("-m","pip","install","--upgrade","pip","--trusted-host","pypi.org","--trusted-host","files.pythonhosted.org","--trusted-host","pypi.pythonhosted.org")
if ($proxy) { $pipArgs += @("--proxy", $proxy) }
& "$venvPath\Scripts\python.exe" @pipArgs
if ($LASTEXITCODE -ne 0) { Write-Error "Failed to upgrade pip. Check connectivity."; exit 1 }

Write-Host "[6/7] Installing requirements.txt..."
$reqArgs = @("-m","pip","install","-r", "$scriptDir\requirements.txt","--trusted-host","pypi.org","--trusted-host","files.pythonhosted.org","--trusted-host","pypi.pythonhosted.org","--no-cache-dir")
if ($proxy) { $reqArgs += @("--proxy", $proxy) }
& "$venvPath\Scripts\python.exe" @reqArgs
if ($LASTEXITCODE -ne 0) { Write-Error "Failed to install requirements. Check connectivity or pip.ini."; exit 1 }

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
    $iwr = @{ Uri = $a.Url; OutFile = $outFile; UseBasicParsing = $true; TimeoutSec = 120 }
    if ($proxy) { $iwr["Proxy"] = $proxy }
    Invoke-WebRequest @iwr
  } catch {
    Write-Warning ("  Failed to download " + $a.File + ": " + $_.Exception.Message)
  }
}

# 6c) Generate public link config for shared-folder HTML redirect
Write-Host "[6c] Generating public link config..."
$publicDir = Join-Path $scriptDir "public"
New-Item -ItemType Directory -Force -Path $publicDir | Out-Null
$env:FLASK_PORT = (Select-String -Path "$scriptDir\.env" -Pattern '^FLASK_PORT=(.+)$').Matches.Groups[1].Value
if (-not $env:FLASK_PORT) { $env:FLASK_PORT = "5000" }
$hostname = $env:COMPUTERNAME
$baseUrl = "http://$hostname:$env:FLASK_PORT"
@{ baseUrl = $baseUrl } | ConvertTo-Json -Compress | Set-Content -Path (Join-Path $publicDir "config_public.json") -Encoding Ascii
$html = @"
<!DOCTYPE html>
<html lang='en'>
<head>
<meta charset='utf-8'>
<title>Adient Production Dashboard</title>
<meta http-equiv='refresh' content='0; url=$baseUrl' />
<script>location.replace('$baseUrl');</script>
</head>
<body>
<p>Redirecting to <a href='$baseUrl'>$baseUrl</a> ...</p>
</body>
</html>
"@
$html | Set-Content -Path (Join-Path $publicDir "index.html") -Encoding UTF8

# 7) Start server
Write-Host "[7/7] Starting dashboard server using local venv...\n"
$env:FLASK_PORT = (Select-String -Path "$scriptDir\.env" -Pattern '^FLASK_PORT=(.+)$').Matches.Groups[1].Value
if (-not $env:FLASK_PORT) { $env:FLASK_PORT = "5000" }

& "$venvPath\Scripts\python.exe" "$scriptDir\dashboard_server.py"
