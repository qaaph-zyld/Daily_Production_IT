$ErrorActionPreference = "Stop"
$scriptDir = Split-Path -Parent $MyInvocation.MyCommand.Path
$envFile = Join-Path $scriptDir ".env"
$port = "5000"
if (Test-Path $envFile) {
  $m = Select-String -Path $envFile -Pattern '^\s*FLASK_PORT\s*=\s*(\d+)\s*$' -ErrorAction SilentlyContinue
  if ($m) { $port = $m.Matches.Groups[1].Value }
}
$publicDir = Join-Path $scriptDir "public"
New-Item -ItemType Directory -Force -Path $publicDir | Out-Null
$hostname = $env:COMPUTERNAME
$baseUrl = "http://$($hostname):$port"
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
