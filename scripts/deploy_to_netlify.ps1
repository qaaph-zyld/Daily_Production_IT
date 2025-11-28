# deploy_to_netlify.ps1
# Automated daily deployment of PVS static snapshot to Netlify
# Runs at 08:45 CET Mon-Fri via Windows Task Scheduler

param(
    [string]$SiteId = "",          # Netlify site ID (set after first deploy)
    [string]$AuthToken = ""        # Netlify personal access token
)

$ErrorActionPreference = "Stop"
$ScriptDir = Split-Path -Parent $MyInvocation.MyCommand.Path
$ProjectRoot = Split-Path -Parent $ScriptDir
$StaticDir = Join-Path $ProjectRoot "netlify_static"
$LogFile = Join-Path $ProjectRoot "logs\netlify_deploy.log"

# Ensure logs directory exists
$LogDir = Split-Path -Parent $LogFile
if (-not (Test-Path $LogDir)) {
    New-Item -ItemType Directory -Path $LogDir -Force | Out-Null
}

function Log {
    param([string]$Message)
    $ts = Get-Date -Format "yyyy-MM-dd HH:mm:ss"
    $line = "[$ts] $Message"
    Write-Host $line
    Add-Content -Path $LogFile -Value $line
}

Log "========== Starting PVS Netlify Deployment =========="

# Step 1: Generate fresh static snapshot
Log "Generating static snapshot..."
try {
    Push-Location $ProjectRoot
    $pythonPath = "python"
    $result = & $pythonPath "scripts/generate_static_pvs.py" 2>&1
    $result | ForEach-Object { Log $_ }
    if ($LASTEXITCODE -ne 0) {
        throw "Static generation failed with exit code $LASTEXITCODE"
    }
    Pop-Location
    Log "Static snapshot generated successfully"
} catch {
    Log "ERROR: Failed to generate static snapshot: $_"
    exit 1
}

# Step 2: Deploy to Netlify using CLI
Log "Deploying to Netlify..."

# Check if Netlify CLI is installed
$netlifyCmd = Get-Command netlify -ErrorAction SilentlyContinue
if (-not $netlifyCmd) {
    Log "ERROR: Netlify CLI not found. Install with: npm install -g netlify-cli"
    exit 1
}

try {
    Push-Location $ProjectRoot
    
    # Build deploy command (match manual successful command)
    $deployArgs = @("deploy", "--prod", "--dir=netlify_static", "--no-build")
    
    if ($SiteId) {
        $deployArgs += "--site=$SiteId"
    }
    
    if ($AuthToken) {
        $env:NETLIFY_AUTH_TOKEN = $AuthToken
    }
    
    $result = & netlify @deployArgs 2>&1
    $result | ForEach-Object { Log $_ }
    
    if ($LASTEXITCODE -ne 0) {
        throw "Netlify deploy failed with exit code $LASTEXITCODE"
    }
    
    Pop-Location
    Log "Deployment completed successfully!"
} catch {
    Log "ERROR: Netlify deployment failed: $_"
    exit 1
}

Log "========== Deployment Complete =========="
