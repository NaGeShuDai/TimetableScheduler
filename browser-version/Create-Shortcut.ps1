$ErrorActionPreference = 'Stop'
$shortcutName = 'HNU Browser Scheduler.lnk'
$shell = New-Object -ComObject WScript.Shell
foreach ($folder in @($PSScriptRoot, [Environment]::GetFolderPath('Desktop'))) {
    $shortcut = $shell.CreateShortcut((Join-Path $folder $shortcutName))
    $shortcut.TargetPath = Join-Path $PSScriptRoot 'Start-Browser.cmd'
    $shortcut.WorkingDirectory = $PSScriptRoot
    $shortcut.IconLocation = (Join-Path $PSScriptRoot 'public\assets\app.ico') + ',0'
    $shortcut.Description = 'Hainan University local browser timetable scheduler'
    $shortcut.Save()
}
Write-Host 'Project and desktop shortcuts created.'
