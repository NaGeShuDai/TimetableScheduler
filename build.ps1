$ErrorActionPreference = 'Stop'
Set-Location -LiteralPath $PSScriptRoot
if (-not (Test-Path -LiteralPath '.venv/Scripts/python.exe')) {
    py -3 -m venv .venv
}
& .venv/Scripts/python.exe -m pip install -r requirements.txt
if ($LASTEXITCODE -ne 0) { throw 'Dependency installation failed' }
& .venv/Scripts/python.exe -m PyInstaller --noconfirm --clean --onefile --windowed --name TimetableScheduler --icon resources/app.ico --distpath . --workpath build --collect-all ortools --add-data "resources;resources" app.py
if ($LASTEXITCODE -ne 0) { throw 'Build failed' }
