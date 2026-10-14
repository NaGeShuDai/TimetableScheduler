$listener = Get-NetTCPConnection -LocalAddress 127.0.0.1 -LocalPort 8765 -State Listen -ErrorAction SilentlyContinue
foreach ($entry in $listener) {
    $process = Get-CimInstance Win32_Process -Filter "ProcessId=$($entry.OwningProcess)"
    if ($process.Name -eq 'node.exe' -and $process.CommandLine.Contains((Join-Path $PSScriptRoot 'server.cjs'))) {
        Stop-Process -Id $entry.OwningProcess
        Write-Host 'Local browser server stopped.'
    }
}
