#Requires -Version 5.1
<#
.SYNOPSIS
Lists report jobs from ConnectSecure using saved credentials.

.EXAMPLE
.\Get-ReportJobs.ps1
.\Get-ReportJobs.ps1 -CompanyId 15853
.\Get-ReportJobs.ps1 -Limit 50 -Skip 0
#>
param(
    [int]$CompanyId = 0,
    [int]$Limit = 25,
    [int]$Skip = 0
)

$credsPath = Join-Path $env:LOCALAPPDATA "VScanMagic\ConnectSecure-Credentials.json"
if (-not (Test-Path $credsPath)) {
    Write-Host "ERROR: No saved credentials. Run VScanMagic GUI and save credentials first." -ForegroundColor Red
    exit 1
}

$creds = Get-Content $credsPath -Raw | ConvertFrom-Json
$baseUrl = $creds.BaseUrl.ToString().TrimEnd('/')
$tenantName = $creds.TenantName.ToString().Trim()
$clientId = $creds.ClientId.ToString().Trim()
$clientSecret = $creds.ClientSecret.ToString().Trim()

# Auth
$authStr = "${tenantName}+${clientId}:$clientSecret"
$authBytes = [System.Text.Encoding]::UTF8.GetBytes($authStr)
$b64 = [Convert]::ToBase64String($authBytes)
$authResp = Invoke-RestMethod -Uri "$baseUrl/w/authorize" -Method Post -Headers @{
    "Client-Auth-Token" = $b64
    "Content-Type" = "application/json"
} -Body "{}" -ContentType "application/json"

$token = $authResp.access_token
$userId = $authResp.user_id
if (-not $token -and $authResp.data) {
    $token = $authResp.data.access_token
    $userId = $authResp.data.user_id
}
if (-not $token) { Write-Host "Auth failed" -ForegroundColor Red; exit 1 }

# Get report jobs (condition param causes "Failed to retrieve data" - filter client-side)
# When filtering by company, fetch more to increase chance of finding matches
$fetchLimit = if ($CompanyId -gt 0) { [Math]::Max($Limit, 250) } else { $Limit }
$query = "limit=$fetchLimit&skip=$Skip"
$uri = "$baseUrl/r/company/report_jobs_view?$query"
$headers = @{
    "Authorization" = "Bearer $token"
    "X-USER-ID"     = $userId
    "Accept"        = "application/json"
}

try {
    $resp = Invoke-RestMethod -Uri $uri -Method Get -Headers $headers
    $allJobs = @($resp.data)
    if (-not $allJobs) { $allJobs = @() }
    if ($CompanyId -gt 0) {
        $filtered = @()
        foreach ($j in $allJobs) {
            if (([int]$j.company_id) -eq $CompanyId) { $filtered += $j }
        }
        $jobs = $filtered
        Write-Host "Report jobs for company $CompanyId ($($jobs.Count)):" -ForegroundColor Cyan
    } else {
        $jobs = $allJobs
        Write-Host "`nReport jobs ($($jobs.Count)):" -ForegroundColor Cyan
    }
    foreach ($j in $jobs) {
        $jid = $j.job_id
        $st = $j.status
        $typ = $j.type
        $desc = $j.description
        $cid = $j.company_id
        $cname = $j.company_name
        Write-Host "  job_id=$jid status=$st type=$typ company_id=$cid description=$desc"
    }
    $resp
} catch {
    $sc = $null
    if ($_.Exception.Response) { try { $sc = [int]$_.Exception.Response.StatusCode } catch {} }
    Write-Host "Error: HTTP $sc - $($_.Exception.Message)" -ForegroundColor Red
    if ($_.ErrorDetails.Message) { Write-Host $_.ErrorDetails.Message }
}
