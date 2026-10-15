$ErrorActionPreference = 'Stop'
if (Get-Process -Name TimetableScheduler -ErrorAction SilentlyContinue) {
    throw 'Close TimetableScheduler before clearing its saved data.'
}
$dataFolders = @((Join-Path $PSScriptRoot 'data'), (Join-Path $env:LOCALAPPDATA 'TimetableScheduler\data'))
foreach ($folder in $dataFolders) {
    foreach ($name in @('workspace.json', 'workspace.tmp', 'error.log')) {
        $target = Join-Path $folder $name
        if (Test-Path -LiteralPath $target) { Remove-Item -LiteralPath $target -Force }
    }
}
Write-Host 'Desktop saved data cleared. Original Excel files and resources are unchanged.'
