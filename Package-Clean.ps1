$ErrorActionPreference = 'Stop'
Add-Type -AssemblyName System.IO.Compression.FileSystem
function Write-Package($destination, $files, $root, $prefix) {
    $temporary = $destination + '.tmp'
    if (Test-Path -LiteralPath $temporary) { Remove-Item -LiteralPath $temporary }
    $archive = [System.IO.Compression.ZipFile]::Open($temporary, [System.IO.Compression.ZipArchiveMode]::Create)
    try {
        foreach ($file in $files) {
            $relative = $file.FullName.Substring($root.Length + 1).Replace('\', '/')
            [System.IO.Compression.ZipFileExtensions]::CreateEntryFromFile($archive, $file.FullName, $prefix + $relative, [System.IO.Compression.CompressionLevel]::Optimal) | Out-Null
        }
    } finally { $archive.Dispose() }
    Move-Item -LiteralPath $temporary -Destination $destination -Force
}
$desktopFiles = @('TimetableScheduler.exe', 'README.md', 'Clear-Desktop-Data.cmd', 'Clear-Desktop-Data.ps1') | ForEach-Object { Get-Item -LiteralPath (Join-Path $PSScriptRoot $_) }
$desktopFiles += Get-ChildItem -LiteralPath (Join-Path $PSScriptRoot 'resources') -Recurse -File
Write-Package (Join-Path $PSScriptRoot 'TimetableScheduler-Clean.zip') $desktopFiles $PSScriptRoot 'TimetableScheduler/'
$browserRoot = Join-Path $PSScriptRoot 'browser-version'
if (-not (Test-Path -LiteralPath (Join-Path $browserRoot 'runtime\node.exe'))) { throw 'Missing browser runtime/node.exe' }
$browserFiles = Get-ChildItem -LiteralPath $browserRoot -File | Where-Object { $_.Extension -in @('.cmd', '.ps1', '.md', '.mjs', '.cjs', '.json') }
foreach ($folder in @('public', 'src', 'runtime')) { $browserFiles += Get-ChildItem -LiteralPath (Join-Path $browserRoot $folder) -Recurse -File }
Write-Package (Join-Path $PSScriptRoot 'HNU-BrowserScheduler-Portable.zip') $browserFiles $browserRoot 'browser-version/'
Write-Host 'Clean desktop and portable browser packages generated. No workspace, examples, backups, screenshots or browser profiles are included.'
