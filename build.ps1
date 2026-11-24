param([string]$Python = 'python')
$ErrorActionPreference = 'Stop'
Set-Location -LiteralPath $PSScriptRoot
& $Python -m pip install -r requirements-build.txt Pillow
if ($LASTEXITCODE -ne 0) { throw 'Dependency installation failed.' }
& $Python tools/build_assets.py
if ($LASTEXITCODE -ne 0) { throw 'Asset creation failed.' }
& $Python -m unittest discover -s tests -v
if ($LASTEXITCODE -ne 0) { throw 'Tests failed.' }
& $Python -m PyInstaller --noconfirm --clean --onefile --windowed --name 'DBN Email Studio' --icon assets/email-studio-blue.ico --add-data 'web;web' --additional-hooks-dir hooks --collect-data extract_msg --collect-data RTFDE --collect-data lark --exclude-module tkinter --exclude-module numpy --exclude-module pandas --exclude-module matplotlib --exclude-module scipy --exclude-module torch --exclude-module IPython app.py
if ($LASTEXITCODE -ne 0) { throw 'Windows build failed.' }
Write-Host 'Built: dist\DBN Email Studio.exe'
