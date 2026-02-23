# Minimal script to fetch and save API schema - no dot-source of ConnectSecure-API
param([int]$CompanyId = 0)
$credsPath = Join-Path $env:LOCALAPPDATA "VScanMagic\ConnectSecure-Credentials.json"
if (-not (Test-Path $credsPath)) { Write-Host "No credentials. Run GUI and save ConnectSecure credentials first." -ForegroundColor Yellow; exit 1 }
$creds = Get-Content $credsPath -Raw | ConvertFrom-Json
$authRaw = "$($creds.TenantName)+$($creds.ClientId):$($creds.ClientSecret)"
$authB64 = [Convert]::ToBase64String([Text.Encoding]::UTF8.GetBytes($authRaw))
$headers = @{
    "Client-Auth-Token" = $authB64
    "Content-Type" = "application/json"
    "accept" = "application/json"
}
$authUrl = "$($creds.BaseUrl.TrimEnd('/'))/w/authorize"
try {
    $authResp = Invoke-RestMethod -Uri $authUrl -Method Post -Headers $headers -Body "{}" -TimeoutSec 30
    if (-not $authResp.status -or -not $authResp.data.access_token) { Write-Host "Auth failed"; exit 1 }
} catch { Write-Host "Auth error: $_"; exit 1 }
$token = $authResp.data.access_token
$userId = $authResp.data.user_id
$apiHeaders = @{
    "Authorization" = "Bearer $token"
    "X-USER-ID" = $userId
    "accept" = "application/json"
}
$apiUrl = "$($creds.BaseUrl.TrimEnd('/'))/r/report_queries/application_vulnerabilities?limit=3&skip=0"
if ($CompanyId -gt 0) { $apiUrl += "&company_id=$CompanyId" }
$resp = Invoke-RestMethod -Uri $apiUrl -Headers $apiHeaders -TimeoutSec 90
if (-not $resp.status -or -not $resp.data -or $resp.data.Count -eq 0) { Write-Host "No data"; exit 0 }
$first = $resp.data[0]
$outPath = Join-Path $PSScriptRoot "schema-sample.json"
$first | ConvertTo-Json -Depth 10 | Set-Content -Path $outPath -Encoding UTF8
Write-Host "Saved to $outPath" -ForegroundColor Green
Write-Host "`nTop-level keys:" ($first.PSObject.Properties.Name -join ", ")
