param([switch]$TrustLocalCertificate)
$ErrorActionPreference='Stop'
$taskUrl='http://127.0.0.1:8767'
try {$taskHealth=Invoke-RestMethod ($taskUrl+'/health') -TimeoutSec 2} catch {
    $taskEngine=Join-Path $PSScriptRoot '..\DBN Email Studio.exe'
    if(!(Test-Path -LiteralPath $taskEngine)){$taskEngine=Join-Path $PSScriptRoot '..\dist\DBN Email Studio.exe'}
    if(!(Test-Path -LiteralPath $taskEngine)){throw 'Open DBN Email Studio, then run this setup again.'}
    Start-Process -FilePath $taskEngine -ArgumentList '--no-browser' -WindowStyle Hidden
    for($taskAttempt=0;$taskAttempt -lt 90;$taskAttempt++){
        try{$taskHealth=Invoke-RestMethod ($taskUrl+'/health') -TimeoutSec 1;break}catch{Start-Sleep -Milliseconds 500}
    }
}
if($taskHealth.app -ne 'dbn-email-studio-v1'){throw 'The local service is not DBN Email Studio.'}
$taskHtml=(Invoke-WebRequest -UseBasicParsing $taskUrl).Content
$taskToken=[regex]::Match($taskHtml,'name="app-token" content="([A-Za-z0-9_-]{43})"').Groups[1].Value
if(!$taskToken){throw 'Could not connect to your local studio.'}
$taskStatus=Invoke-RestMethod ($taskUrl+'/api/office/prepare') -Method Post -ContentType 'application/json' -Headers @{'X-App-Token'=$taskToken} -Body '{}'
$taskCertificate=Get-PfxCertificate -FilePath $taskStatus.certificate_path
$taskHash=(Get-FileHash -LiteralPath $taskStatus.certificate_path -Algorithm SHA256).Hash
if($taskHash -ne $taskStatus.sha256){throw 'Local certificate fingerprint mismatch.'}
Write-Host 'Local Outlook add-in prepared. It has not been installed in Outlook.'
Write-Host ('HTTPS address: '+$taskStatus.url)
Write-Host ('Certificate SHA256: '+$taskHash)
Write-Host ('Expires: '+$taskCertificate.NotAfter)
if($TrustLocalCertificate){
    Import-Certificate -FilePath $taskStatus.certificate_path -CertStoreLocation Cert:\CurrentUser\Root | Out-Null
    Write-Host 'The localhost certificate is now trusted for this Windows user.'
}else{
    Write-Host 'No certificate trust was installed. Review the README and certificate before enabling trust.'
}
