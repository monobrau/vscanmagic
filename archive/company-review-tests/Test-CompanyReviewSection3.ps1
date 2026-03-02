#Requires -Version 5.1
param([int]$CompanyId = 15373)

$scriptDir = if ($PSScriptRoot) { $PSScriptRoot } else { (Get-Location).Path }
. (Join-Path $scriptDir "VScanMagic-Modules\VScanMagic-Core.ps1")
. (Join-Path $scriptDir "ConnectSecure-API.ps1")

$creds = Load-ConnectSecureCredentials
Connect-ConnectSecureAPI -BaseUrl $creds.BaseUrl -TenantName $creds.TenantName -ClientId $creds.ClientId -ClientSecret $creds.ClientSecret | Out-Null

$data = Get-ConnectSecureCompanyReviewData -CompanyId $CompanyId -CompanyName 'Test'
Write-Host "External assets count: $($data.ExternalAssets.Count)"
Write-Host "Last External Scan: $($data.LastExternalScan)"
Write-Host "Last Internal Scan: $($data.LastInternalScan)"
foreach ($ea in $data.ExternalAssets) {
    Write-Host "  $($ea.Name): $($ea.Address)"
}
