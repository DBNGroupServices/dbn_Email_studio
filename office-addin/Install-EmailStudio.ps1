<#
.SYNOPSIS
  Installs or upgrades DBN Email Studio for Outlook Classic on this computer (current Windows user only).

.DESCRIPTION
  1. Copies "DBN Email Studio.exe" to %LOCALAPPDATA%\Programs\DBN Email Studio
  2. Turns on "start at Windows sign-in" so the Outlook pane is always available
  3. Starts the engine and prepares the local HTTPS endpoint (https://localhost:8788)
  4. Trusts the engine's local certificate authority for this Windows user
     (Windows shows its own confirmation prompt - choose Yes)
  5. Checks that Outlook will be able to load the add-in pages
  6. Copies manifest.xml next to the app and opens Outlook's custom add-in page

  No administrator rights are needed. Nothing is changed machine-wide.

.PARAMETER ExePath
  Path to "DBN Email Studio.exe". Defaults to ..\dist or the folder above this script.
.PARAMETER NoAutoStart
  Do not start the engine at Windows sign-in (not recommended - the pane will fail to load after a restart).
.PARAMETER NoTrust
  Do not install certificate trust (the pane will not load until trust is installed).
.PARAMETER ClearOfficeCache
  Clear Outlook's cached copy of web add-ins. Close Outlook first.
.PARAMETER NoSideloadPage
  Do not open the Outlook "custom add-ins" page at the end.
#>
[CmdletBinding()]
param(
  [string]$ExePath,
  [switch]$NoAutoStart,
  [switch]$NoTrust,
  [switch]$ClearOfficeCache,
  [switch]$NoSideloadPage
)
$ErrorActionPreference = 'Stop'
$AppName    = 'DBN Email Studio'
$InstallDir = Join-Path $env:LOCALAPPDATA "Programs\$AppName"
$Target     = Join-Path $InstallDir "$AppName.exe"
$DataDir    = Join-Path $env:LOCALAPPDATA $AppName
$EngineUrl  = 'http://127.0.0.1:8767'
$OfficeUrl  = 'https://localhost:8788'
$RunKey     = 'HKCU:\Software\Microsoft\Windows\CurrentVersion\Run'

function Step($text)  { Write-Host ''; Write-Host "==> $text" -ForegroundColor Cyan }
function Ok($text)    { Write-Host "    [OK]   $text" -ForegroundColor Green }
function Warn($text)  { Write-Host "    [WARN] $text" -ForegroundColor Yellow }
function Fail($text)  { Write-Host "    [FAIL] $text" -ForegroundColor Red }

function Get-Health([string]$url) {
  try { return Invoke-RestMethod -Uri ($url + '/health') -TimeoutSec 2 -UseBasicParsing } catch { return $null }
}

# ---------- 1. Locate and install the executable ----------
Step 'Installing the app'
if (-not $ExePath) {
  foreach ($candidate in @((Join-Path $PSScriptRoot "..\dist\$AppName.exe"), (Join-Path $PSScriptRoot "..\$AppName.exe"), (Join-Path $PSScriptRoot "$AppName.exe"))) {
    if (Test-Path -LiteralPath $candidate) { $ExePath = (Resolve-Path -LiteralPath $candidate).Path; break }
  }
}
if (-not $ExePath -or -not (Test-Path -LiteralPath $ExePath)) {
  throw "Could not find '$AppName.exe'. Build it with build.ps1, or run: .\Install-EmailStudio.ps1 -ExePath 'C:\path\to\$AppName.exe'"
}
$sameFile = (Test-Path -LiteralPath $Target) -and ((Resolve-Path -LiteralPath $ExePath).Path -eq (Resolve-Path -LiteralPath $Target).Path)
$running = Get-Process -Name $AppName -ErrorAction SilentlyContinue
if ($running -and -not $sameFile) {
  Write-Host "    Closing the running copy of $AppName so it can be upgraded (your settings and sign-in are kept)."
  $running | Stop-Process -Force
  Start-Sleep -Seconds 2
}
New-Item -ItemType Directory -Force -Path $InstallDir | Out-Null
if (-not $sameFile) { Copy-Item -LiteralPath $ExePath -Destination $Target -Force }
Copy-Item -LiteralPath (Join-Path $PSScriptRoot 'manifest.xml') -Destination (Join-Path $InstallDir 'manifest.xml') -Force
foreach ($script in 'Diagnose-EmailStudio.ps1','Install-EmailStudio.ps1','Fix Email Studio.cmd') {
  $src = Join-Path $PSScriptRoot $script
  if ((Test-Path -LiteralPath $src) -and ((Resolve-Path -LiteralPath $src).Path -ne (Join-Path $InstallDir $script))) { Copy-Item -LiteralPath $src -Destination $InstallDir -Force }
}
Ok "Installed to $InstallDir"

# Start-menu shortcut
try {
  $shell = New-Object -ComObject WScript.Shell
  $lnk = $shell.CreateShortcut((Join-Path ([Environment]::GetFolderPath('Programs')) "$AppName.lnk"))
  $lnk.TargetPath = $Target; $lnk.WorkingDirectory = $InstallDir; $lnk.IconLocation = "$Target,0"; $lnk.Save()
  Ok 'Start menu shortcut created'
} catch { Warn 'Could not create a Start menu shortcut (optional).' }

# ---------- 2. Start at sign-in ----------
Step 'Start at Windows sign-in'
if ($NoAutoStart) {
  Remove-ItemProperty -Path $RunKey -Name $AppName -ErrorAction SilentlyContinue
  Warn 'Skipped. The Outlook pane will show "can''t load the add-in" after a restart until you open the app.'
} else {
  Set-ItemProperty -Path $RunKey -Name $AppName -Value ('"' + $Target + '" --no-browser')
  Ok 'Email Studio will start quietly in the background when you sign in.'
}

# ---------- 3. Start engine and prepare HTTPS ----------
Step 'Starting the engine'
$health = Get-Health $EngineUrl
if (-not $health) {
  Start-Process -FilePath $Target -ArgumentList '--no-browser' -WorkingDirectory $InstallDir
  for ($i = 0; $i -lt 90 -and -not $health; $i++) { Start-Sleep -Milliseconds 500; $health = Get-Health $EngineUrl }
}
if (-not $health -or $health.app -ne 'dbn-email-studio-v1') { throw 'The engine did not start. Run Diagnose-EmailStudio.ps1 for details.' }
Ok ("Engine running (version " + $health.version + ")")

$page  = (Invoke-WebRequest -UseBasicParsing -Uri $EngineUrl -TimeoutSec 5).Content
$token = [regex]::Match($page, 'name="app-token" content="([A-Za-z0-9_-]{43})"').Groups[1].Value
if (-not $token) { throw 'Could not connect to the local engine.' }
$status = Invoke-RestMethod -Uri ($EngineUrl + '/api/office/prepare') -Method Post -ContentType 'application/json' -Headers @{ 'X-App-Token' = $token } -Body '{}'
$cerPath = $status.certificate_path
if (-not (Test-Path -LiteralPath $cerPath)) { throw "Certificate file not found: $cerPath" }
$cert = New-Object System.Security.Cryptography.X509Certificates.X509Certificate2 -ArgumentList $cerPath
if ($cert.Thumbprint -ne $status.thumbprint) { throw 'Local certificate fingerprint mismatch. Stop and run Diagnose-EmailStudio.ps1.' }
Ok "Local HTTPS ready at $OfficeUrl"
Write-Host ("    Certificate authority : " + $cert.Subject)
Write-Host ("    Thumbprint (SHA-1)    : " + $cert.Thumbprint)
Write-Host ("    SHA-256               : " + $status.sha256)
Write-Host ("    Valid until           : " + $cert.NotAfter.ToString('d MMM yyyy'))
Write-Host  '    This authority can only vouch for localhost / 127.0.0.1 (name-constrained).'

# ---------- 4. Trust ----------
Step 'Certificate trust (this Windows user only)'
$trusted = Get-ChildItem Cert:\CurrentUser\Root | Where-Object { $_.Thumbprint -eq $cert.Thumbprint }
if ($trusted) {
  Ok 'Already trusted.'
} elseif ($NoTrust) {
  Warn 'Skipped. Outlook will not load the pane until this certificate is trusted.'
} else {
  Write-Host '    Windows will ask you to confirm installing the certificate. Check the thumbprint above matches, then choose Yes.'
  Import-Certificate -FilePath $cerPath -CertStoreLocation Cert:\CurrentUser\Root | Out-Null
  if (Get-ChildItem Cert:\CurrentUser\Root | Where-Object { $_.Thumbprint -eq $cert.Thumbprint }) { Ok 'Trusted for this Windows user.' } else { Fail 'Trust was not installed (the prompt may have been declined).' }
}

# ---------- 5. Verify what Outlook will load ----------
Step 'Checking the add-in pages load securely'
try {
  $r = Invoke-WebRequest -UseBasicParsing -Uri ($OfficeUrl + '/office/taskpane.html') -TimeoutSec 5
  if ($r.StatusCode -eq 200) { Ok 'Task pane loads over trusted HTTPS.' }
} catch {
  Fail ("Could not load $OfficeUrl/office/taskpane.html : " + $_.Exception.Message)
  Warn 'Run Diagnose-EmailStudio.ps1 for a full check.'
}

# ---------- 6. Office cache and manifest ----------
if ($ClearOfficeCache) {
  Step 'Clearing Outlook web add-in cache'
  if (Get-Process -Name OUTLOOK -ErrorAction SilentlyContinue) {
    Warn 'Outlook is open - close Outlook and re-run with -ClearOfficeCache.'
  } else {
    $wef = Join-Path $env:LOCALAPPDATA 'Microsoft\Office\16.0\Wef'
    if (Test-Path -LiteralPath $wef) { Remove-Item -LiteralPath $wef -Recurse -Force -ErrorAction SilentlyContinue; Ok 'Cache cleared.' } else { Ok 'No cache to clear.' }
  }
}

Step 'Outlook add-in'
$manifest = Join-Path $InstallDir 'manifest.xml'
Write-Host "    Manifest: $manifest"
Write-Host '    If Email Studio is not on your ribbon yet (or you are upgrading from 1.x and want pinning):'
Write-Host '      Outlook > Get Add-ins > My add-ins > Custom add-ins > Add a custom add-in > Add from file > choose the manifest above.'
Write-Host '    Then restart Outlook. Look for "Email Studio" on the Home tab when reading or composing an email.'
if (-not $NoSideloadPage) { try { Start-Process 'https://aka.ms/olksideload' } catch { } }

Write-Host ''
Write-Host 'Done. Email Studio will now be available in Outlook every time you sign in.' -ForegroundColor Green
Write-Host 'If the pane ever shows "can''t load the add-in", run Diagnose-EmailStudio.ps1 -Repair from:'
Write-Host "  $InstallDir"
