#Requires -Version 5.1
<#
.SYNOPSIS
Compares vulnerability endpoints to find which returns less redundant (fewer) rows.
asset_wise = one row per host+vuln (redundant). application = possibly one row per unique vuln.
#>
param(
    [int]$CompanyId = 0,
    [int]$SampleSize = 1000,
    [switch]$UseSavedCredentials
)

$credsPath = Join-Path $env:LOCALAPPDATA "VScanMagic\ConnectSecure-Credentials.json"
if (-not (Test-Path $credsPath)) {
    Write-Host "No saved credentials. Run VScanMagic GUI and save credentials first." -ForegroundColor Red
    exit 1
}
$creds = Get-Content $credsPath -Raw | ConvertFrom-Json

$baseUrl = $creds.BaseUrl.ToString().TrimEnd('/')
$authStr = "$($creds.TenantName)+$($creds.ClientId):$($creds.ClientSecret)"
$b64 = [Convert]::ToBase64String([Text.Encoding]::UTF8.GetBytes($authStr))

$authResp = Invoke-RestMethod -Uri "$baseUrl/w/authorize" -Method Post -Headers @{
    "Client-Auth-Token" = $b64
    "Content-Type" = "application/json"
} -Body "{}"

$token = $authResp.access_token
$userId = $authResp.user_id
if (-not $token -and $authResp.data) { $token = $authResp.data.access_token; $userId = $authResp.data.user_id }
if (-not $token) { Write-Host "Auth failed"; exit 1 }

$headers = @{
    "Authorization" = "Bearer $token"
    "X-USER-ID" = $userId
    "Accept" = "application/json"
}

$query = "limit=$SampleSize&skip=0&sort=severity.keyword:desc"
if ($CompanyId -gt 0) { $query += "&company_id=$CompanyId" }

$endpoints = @(
    @{ Name = "asset_wise_vulnerabilities"; Path = "/r/report_queries/asset_wise_vulnerabilities"; Desc = "One row per host+vuln (redundant)" }
    @{ Name = "application_vulnerabilities"; Path = "/r/report_queries/application_vulnerabilities"; Desc = "Possibly one row per unique vuln" }
    @{ Name = "application_vulnerabilities_by_product"; Path = "/r/report_queries/application_vulnerabilities_by_product"; Desc = "One row per product (most compact)" }
)

Write-Host "`n=== Comparing endpoints (limit=$SampleSize) ===" -ForegroundColor Cyan
Write-Host ""

foreach ($ep in $endpoints) {
    Write-Host "--- $($ep.Name) ---" -ForegroundColor Yellow
    Write-Host "  $($ep.Desc)"
    try {
        $uri = "$baseUrl$($ep.Path)?$query"
        $resp = Invoke-RestMethod -Uri $uri -Method Get -Headers $headers -TimeoutSec 60
        $data = $null
        if ($resp.status -and $resp.data) { $data = $resp.data }
        elseif ($resp.data) { $data = $resp.data }
        
        if (-not $data) {
            Write-Host "  Rows: 0 (no data)" -ForegroundColor Gray
        } else {
            $count = if ($data -is [Array]) { $data.Count } else { @($data).Count }
            Write-Host "  Rows: $count" -ForegroundColor Green
            $first = if ($data -is [Array]) { $data[0] } else { $data }
            $keys = $first.PSObject.Properties.Name -join ", "
            $hasHost = $keys -match "host|ip|asset"
            Write-Host "  Has host/asset fields: $hasHost" -ForegroundColor $(if ($hasHost) { "Green" } else { "Gray" })
            Write-Host "  Keys: $keys" -ForegroundColor DarkGray
        }
    } catch {
        Write-Host "  Error: $($_.Exception.Message)" -ForegroundColor Red
    }
    Write-Host ""
}

Write-Host "=== Summary ===" -ForegroundColor Cyan
Write-Host "Use application_vulnerabilities or by_product for fewer rows (no per-host detail)."
Write-Host "Use asset_wise only when you need Host Name + IP per row."
Write-Host ""
