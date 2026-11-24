<#
.SYNOPSIS
  Finds out why Outlook says "Sorry, we can't load the add-in" for DBN Email Studio, and optionally fixes it.

.DESCRIPTION
  Checks, in order: the app is installed, the engine is running, the add-in HTTPS port answers,
  the certificate is trusted and in date, start-at-sign-in is on, and Outlook's add-in cache.
  With -Repair it starts the engine, turns on start-at-sign-in and (if Outlook is closed) clears
  Outlook's add-in cache. It never disables certificate checks or changes organisation policy.

.EXAMPLE
  .\Diagnose-EmailStudio.ps1
.EXAMPLE
  .\Diagnose-EmailStudio.ps1 -Repair
#>
[CmdletBinding()]
param([switch]$Repair)
$ErrorActionPreference = 'Continue'
$AppName   = 'DBN Email Studio'
$Exe       = Join-Path $env:LOCALAPPDATA "Programs\$AppName\$AppName.exe"
$EngineUrl = 'http://127.0.0.1:8767'
$OfficeUrl = 'https://localhost:8788'
$RunKey    = 'HKCU:\Software\Microsoft\Windows\CurrentVersion\Run'
$CaFile    = Join-Path $env:LOCALAPPDATA "$AppName\$AppName Local CA.cer"
$script:problems = 0

function Pass($t){ Write-Host "[PASS] $t" -ForegroundColor Green }
function Bad($t,$fix){ $script:problems++; Write-Host "[FAIL] $t" -ForegroundColor Red; if($fix){ Write-Host "       Fix: $fix" -ForegroundColor Yellow } }
function Note($t){ Write-Host "[INFO] $t" -ForegroundColor Gray }
function Health($u){ try { Invoke-RestMethod -Uri ($u + '/health') -TimeoutSec 3 -UseBasicParsing } catch { $null } }

Write-Host "DBN Email Studio - add-in diagnostics  ($(Get-Date -Format 'd MMM yyyy HH:mm'))" -ForegroundColor Cyan
Write-Host ''

# 1. Installed
if (Test-Path -LiteralPath $Exe) { Pass "App installed: $Exe" }
else {
  $proc = Get-Process -Name $AppName -ErrorAction SilentlyContinue | Select-Object -First 1
  if ($proc) { $Exe = $proc.Path; Note "Running from a non-standard location: $Exe" }
  else { Bad 'App not found in the standard location.' 'Run Install-EmailStudio.ps1 from the release folder.' }
}

# 2. Engine
$h = Health $EngineUrl
if (-not $h -and $Repair -and (Test-Path -LiteralPath $Exe)) {
  Note 'Starting the engine...'
  Start-Process -FilePath $Exe -ArgumentList '--no-browser'
  for ($i=0; $i -lt 60 -and -not $h; $i++) { Start-Sleep -Milliseconds 500; $h = Health $EngineUrl }
}
if ($h -and $h.app -eq 'dbn-email-studio-v1') { Pass "Engine running (version $($h.version))" }
elseif ($h) { Bad 'Port 8767 is used by a different program.' 'Close that program, then open DBN Email Studio.' }
else { Bad 'Engine is NOT running - this is the usual cause of "can''t load the add-in".' 'Open DBN Email Studio from the Start menu, or run this script with -Repair.' }

# 3. HTTPS port
$port = Get-NetTCPConnection -LocalPort 8788 -State Listen -ErrorAction SilentlyContinue | Select-Object -First 1
if ($port) {
  $owner = (Get-Process -Id $port.OwningProcess -ErrorAction SilentlyContinue).ProcessName
  if ($owner -and $owner -ne $AppName) { Bad "Port 8788 is held by '$owner', not Email Studio." "Close '$owner' and restart Email Studio." }
  else { Pass 'Add-in port 8788 is listening.' }
} else {
  if ($h) { Bad 'Add-in HTTPS port 8788 is not listening.' 'Run Install-EmailStudio.ps1 once to prepare the Outlook endpoint.' }
}

# 4. Certificate
if (Test-Path -LiteralPath $CaFile) {
  $ca = New-Object System.Security.Cryptography.X509Certificates.X509Certificate2 -ArgumentList $CaFile
  $trusted = Get-ChildItem Cert:\CurrentUser\Root | Where-Object { $_.Thumbprint -eq $ca.Thumbprint }
  if ($trusted) { Pass "Local certificate authority trusted (valid to $($ca.NotAfter.ToString('d MMM yyyy')))." }
  else { Bad 'Local certificate authority is NOT trusted - Outlook will refuse the pane.' 'Run Install-EmailStudio.ps1 and choose Yes on the Windows certificate prompt.' }
  if ($ca.NotAfter -lt (Get-Date).AddDays(30)) { Bad 'Local certificate authority expires within 30 days.' 'Run Install-EmailStudio.ps1 to renew.' }
} else {
  $legacy = Get-ChildItem Cert:\CurrentUser\Root | Where-Object { $_.Subject -like '*DBN Email Studio localhost*' }
  if ($legacy) {
    foreach ($c in $legacy) {
      if ($c.NotAfter -lt (Get-Date)) { Bad "Version 1.x certificate EXPIRED on $($c.NotAfter.ToString('d MMM yyyy'))." 'Install version 2 with Install-EmailStudio.ps1 (certificates then renew automatically).' }
      else { Note "Version 1.x certificate found (expires $($c.NotAfter.ToString('d MMM yyyy'))). Upgrade with Install-EmailStudio.ps1." }
    }
  } else { Bad 'No Email Studio certificate found.' 'Run Install-EmailStudio.ps1.' }
}

# 5. End-to-end page load (uses Windows certificate trust, like Outlook)
try {
  $r = Invoke-WebRequest -UseBasicParsing -Uri ($OfficeUrl + '/office/taskpane.html') -TimeoutSec 5
  if ($r.StatusCode -eq 200) { Pass "Outlook pane page loads: $OfficeUrl/office/taskpane.html" }
} catch {
  $msg = $_.Exception.Message
  if ($msg -match 'trust|certificate|SSL|TLS') { Bad "Pane page blocked by certificate: $msg" 'Run Install-EmailStudio.ps1 to trust the local certificate authority.' }
  else { Bad "Pane page did not load: $msg" 'Start the engine (-Repair) and re-run this check.' }
}

# 6. Start at sign-in
$run = (Get-ItemProperty -Path $RunKey -Name $AppName -ErrorAction SilentlyContinue).$AppName
if ($run) { Pass 'Starts automatically at Windows sign-in.' }
elseif ($Repair -and (Test-Path -LiteralPath $Exe)) { Set-ItemProperty -Path $RunKey -Name $AppName -Value ('"' + $Exe + '" --no-browser'); Pass 'Start at sign-in turned ON.' }
else { Bad 'Not set to start at sign-in - the pane will fail after every restart.' 'Run this script with -Repair, or tick "Start when I sign in" in Email Studio settings.' }

# 7. Outlook cache
$wef = Join-Path $env:LOCALAPPDATA 'Microsoft\Office\16.0\Wef'
if ($Repair) {
  if (Get-Process -Name OUTLOOK -ErrorAction SilentlyContinue) { Note 'Outlook is open, so its add-in cache was not cleared. Close Outlook and re-run -Repair if the pane still fails.' }
  elseif (Test-Path -LiteralPath $wef) { Remove-Item -LiteralPath $wef -Recurse -Force -ErrorAction SilentlyContinue; Pass 'Outlook add-in cache cleared (Outlook will refresh it on next start).' }
} else { Note 'If everything above passes but Outlook still errors: close Outlook and run with -Repair to clear its add-in cache.' }

Write-Host ''
if ($script:problems -eq 0) { Write-Host 'All checks passed. In Outlook, close the Email Studio pane and open it again (or click Retry).' -ForegroundColor Green }
else { Write-Host "$script:problems problem(s) found. Apply the fixes above, or run: .\Diagnose-EmailStudio.ps1 -Repair" -ForegroundColor Yellow }
