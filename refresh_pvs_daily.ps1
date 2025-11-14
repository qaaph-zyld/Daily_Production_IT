# PVS Daily Refresh - Scheduled for 08:45 CET on weekdays
$ErrorActionPreference = "Stop"
$scriptDir = Split-Path -Parent $MyInvocation.MyCommand.Path
$envFile = Join-Path $scriptDir ".env"
$port = "5051"
if (Test-Path $envFile) {
  $m = Select-String -Path $envFile -Pattern '^\s*PVS_PORT\s*=\s*(\d+)\s*$' -ErrorAction SilentlyContinue
  if ($m) { $port = $m.Matches.Groups[1].Value }
}
$log = Join-Path $scriptDir "pvs_refresh_log.txt"
$stamp = Get-Date -Format "yyyy-MM-dd HH:mm:ss"
try {
  $health = Invoke-WebRequest -Uri ("http://localhost:{0}/api/health" -f $port) -UseBasicParsing -TimeoutSec 25
  $data = Invoke-WebRequest -Uri ("http://localhost:{0}/api/pvs?_={1}" -f $port, [DateTimeOffset]::Now.ToUnixTimeSeconds()) -UseBasicParsing -TimeoutSec 60
  Add-Content -Path $log -Value ("{0} OK health:{1} data:{2}" -f $stamp, $health.StatusCode, $data.StatusCode)
} catch {
  Add-Content -Path $log -Value ("{0} ERR {1}" -f $stamp, $_.Exception.Message)
}
