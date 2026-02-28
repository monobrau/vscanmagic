#Requires -Version 5.1
. $PSScriptRoot\ConnectSecure-API.ps1
$credsPath = Join-Path $env:LOCALAPPDATA 'VScanMagic\ConnectSecure-Credentials.json'
if (-not (Test-Path $credsPath)) {
    Write-Host 'No credentials at' $credsPath -ForegroundColor Red
    exit 1
}
$creds = Get-Content $credsPath -Raw | ConvertFrom-Json
$ok = Connect-ConnectSecureAPI -BaseUrl $creds.BaseUrl -TenantName $creds.TenantName -ClientId $creds.ClientId -ClientSecret $creds.ClientSecret
if (-not $ok) {
    Write-Host 'Auth failed' -ForegroundColor Red
    exit 1
}
Write-Host 'Fetching 3 vulnerabilities (vulnerabilities_details)...' -ForegroundColor Cyan
$v = Get-ConnectSecureVulnerabilities -CompanyId 0 -Limit 3 -FetchAll:$false
Write-Host "Got $($v.Count) rows" -ForegroundColor Green
if ($v.Count -gt 0) {
    $first = $v[0]
    $ah = if ($first.AffectedHosts) { $first.AffectedHosts.Substring(0,[Math]::Min(60,$first.AffectedHosts.Length)) } else { '(empty)' }
    Write-Host "First: problem_name=$($first.problem_name) severity=$($first.severity) HostCount=$($first.HostCount) AffectedHosts=$ah"
}
