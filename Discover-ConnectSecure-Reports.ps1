#Requires -Version 5.1
<#
.SYNOPSIS
Discovers ConnectSecure report builder API - standard reports and report job endpoints.
Use this to find the correct report type IDs/names for your ConnectSecure instance.
Run after loading credentials (or pass credentials).

.EXAMPLE
.\Discover-ConnectSecure-Reports.ps1
# Uses saved credentials from ConnectSecure-API
#>
param(
    [string]$BaseUrl,
    [string]$TenantName,
    [string]$ClientId,
    [string]$ClientSecret
)

$scriptPath = Split-Path -Parent $MyInvocation.MyCommand.Path
. (Join-Path $scriptPath "ConnectSecure-API.ps1")

if (-not $BaseUrl) {
    $credsPath = Join-Path $env:LOCALAPPDATA "VScanMagic\ConnectSecure-Credentials.json"
    if (Test-Path $credsPath) {
        try {
            $creds = Get-Content $credsPath -Raw | ConvertFrom-Json
            $BaseUrl = $creds.BaseUrl
            $TenantName = $creds.TenantName
            $ClientId = $creds.ClientId
            $ClientSecret = $creds.ClientSecret
        } catch { }
    }
}

if (-not $BaseUrl -or -not $TenantName -or -not $ClientId -or -not $ClientSecret) {
    Write-Host "Usage: Provide credentials or ensure Load-ConnectSecureCredentials returns saved creds." -ForegroundColor Yellow
    Write-Host "Example: .\Discover-ConnectSecure-Reports.ps1 -BaseUrl 'https://pod104.myconnectsecure.com' -TenantName 'river-run' -ClientId '...' -ClientSecret '...'" -ForegroundColor Gray
    exit 1
}

Write-Host "`n=== ConnectSecure Report API Discovery ===" -ForegroundColor Cyan
Write-Host "Base URL: $BaseUrl`n" -ForegroundColor Gray

$connected = Connect-ConnectSecureAPI -BaseUrl $BaseUrl -TenantName $TenantName -ClientId $ClientId -ClientSecret $ClientSecret
if (-not $connected) {
    Write-Host "Authentication failed." -ForegroundColor Red
    exit 1
}

Write-Host "`n1. Standard Reports (GET /report_builder/standard_reports, isGlobal=false):" -ForegroundColor Yellow
$reports = Get-ConnectSecureStandardReports -IsGlobal $false
if ($reports -and $reports.Count -gt 0) {
    foreach ($r in $reports) {
        $id = if ($r.id) { $r.id } else { "?" }
        $type = if ($r.reportType) { $r.reportType } else { "?" }
        Write-Host "   id=$id reportType=$type" -ForegroundColor White
    }
    Write-Host "`n   Full JSON:" -ForegroundColor Gray
    $reports | ConvertTo-Json -Depth 5
} else {
    Write-Host "   Trying isGlobal=true..." -ForegroundColor Gray
    $reports = Get-ConnectSecureStandardReports -IsGlobal $true
    if ($reports -and $reports.Count -gt 0) {
        foreach ($r in $reports) { Write-Host "   id=$($r.id) reportType=$($r.reportType)" }
    } else {
        Write-Host "   No standard reports returned. Check Swagger at $BaseUrl/apidocs/" -ForegroundColor Gray
    }
}

Write-Host "`n2. Report Jobs (GET /r/company/report_jobs_view):" -ForegroundColor Yellow
try {
    $jobsResp = Invoke-ConnectSecureRequest -Endpoint "/r/company/report_jobs_view" -Method GET -QueryParameters @{ limit = 5 }
    if ($jobsResp.data -and $jobsResp.data.Count -gt 0) {
        $jobsResp.data | ForEach-Object { Write-Host "   job_id=$($_.job_id) status=$($_.status) type=$($_.type)" }
    } else {
        Write-Host "   No recent report jobs (or empty list)" -ForegroundColor Gray
    }
} catch {
    Write-Host "   Error: $($_.Exception.Message)" -ForegroundColor Red
}

Write-Host "`n3. Swagger/API Docs:" -ForegroundColor Yellow
Write-Host "   Open $BaseUrl/apidocs/ in a browser. Look for:" -ForegroundColor White
Write-Host "   - report_builder / standard_reports" -ForegroundColor Gray
Write-Host "   - report_builder / create_report_job" -ForegroundColor Gray
Write-Host "   - report_jobs_view or report_jobs" -ForegroundColor Gray
Write-Host "   - get_report_link" -ForegroundColor Gray
Write-Host "   Use Get-Token-For-Swagger.ps1 to get a Bearer token for Authorize.`n" -ForegroundColor Gray
