# Superseded in version 2.0 by Install-EmailStudio.ps1 (kept so old instructions still work).
param([switch]$TrustLocalCertificate)
& (Join-Path $PSScriptRoot 'Install-EmailStudio.ps1') -NoSideloadPage
