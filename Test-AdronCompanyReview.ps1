. .\VScanMagic-Modules\VScanMagic-Core.ps1
. .\ConnectSecure-API.ps1
$creds = Load-ConnectSecureCredentials
Connect-ConnectSecureAPI -BaseUrl $creds.BaseUrl -TenantName $creds.TenantName -ClientId $creds.ClientId -ClientSecret $creds.ClientSecret | Out-Null

$data = Get-ConnectSecureCompanyReviewData -CompanyId 12345 -CompanyName 'Example Company'
Write-Host "External assets count: $($data.ExternalAssets.Count)"
$data.ExternalAssets | ForEach-Object { Write-Host "  $($_.Name): $($_.Address)" }
