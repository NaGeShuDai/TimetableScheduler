$ErrorActionPreference = 'Stop'
Set-Location -LiteralPath $PSScriptRoot
$nodePath = Join-Path $PSScriptRoot 'runtime\node.exe'
if (-not (Test-Path -LiteralPath $nodePath)) {
    Write-Host 'Missing runtime\node.exe. Extract the complete portable package first.'
    Read-Host 'Press Enter to exit'
    exit 1
}
$url = 'http://127.0.0.1:8765'
try { $response = Invoke-WebRequest -UseBasicParsing -Uri "$url/index.html" -TimeoutSec 2; if ($response.Content -notmatch 'id="welcome"') { throw 'Port is occupied' }; Start-Process $url; exit } catch { }
Start-Process -FilePath $nodePath -ArgumentList ('"' + (Join-Path $PSScriptRoot 'server.cjs') + '"') -WorkingDirectory $PSScriptRoot -WindowStyle Hidden
for ($i = 0; $i -lt 30; $i++) {
    Start-Sleep -Milliseconds 200
    try {
        $response = Invoke-WebRequest -UseBasicParsing -Uri "$url/index.html" -TimeoutSec 1
        if ($response.Content -match 'id="welcome"') { Start-Process $url; exit }
    } catch { }
}
Write-Host 'Startup failed. Check port 8765 or run runtime\node.exe server.cjs for diagnostics.'
Read-Host 'Press Enter to exit'
