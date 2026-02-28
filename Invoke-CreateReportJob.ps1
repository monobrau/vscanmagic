#Requires -Version 5.1
<#
.SYNOPSIS
Creates a report job via ConnectSecure API using saved credentials.

.EXAMPLE
.\Invoke-CreateReportJob.ps1 -CompanyId 15853 -ReportId "1d091564830b44c485a0ddc35ace9ac6" -ReportName "Suppressed Vulnerabilities"
#>
param(
    [int]$CompanyId = 15853,
    [string]$ReportId = "1d091564830b44c485a0ddc35ace9ac6",
    [string]$ReportName = "Suppressed Vulnerabilities",
    [string]$ReportType = "xlsx"
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

$headers = @{
    "Authorization" = "Bearer $token"
    "X-USER-ID"     = $userId
    "Accept"        = "application/json"
    "Content-Type"  = "application/json"
}

# Fetch company name by company_id
$companyName = "Company $CompanyId"
if ($CompanyId -gt 0) {
    try {
        $companiesResp = Invoke-RestMethod -Uri "$baseUrl/r/company/companies?limit=5000" -Method Get -Headers $headers
        $companies = @()
        if ($companiesResp.data) { $companies = @($companiesResp.data) }
        $match = $companies | Where-Object { $cid = $_.id -or $_.company_id -or $_.companyId; [int]$cid -eq $CompanyId } | Select-Object -First 1
        if ($match) {
            $companyName = ($match.name -or $match.company_name -or $match.companyName -or '').ToString().Trim()
            if ($companyName) { Write-Host "Company: $companyName (id=$CompanyId)" -ForegroundColor Gray }
        }
    } catch {
        Write-Host "Could not fetch company name: $($_.Exception.Message)" -ForegroundColor Yellow
    }
}

# Create report job (payload matches portal: reportType="Standard", isFilter=true)
$reportNameVal = $ReportName -replace '\s',''  # Portal uses "AllVulnerabilitiesReport" (no spaces)
if (-not $reportNameVal) { $reportNameVal = "Report" }
$body = @{
    reportId     = $ReportId
    reportName   = $reportNameVal
    reportType   = "Standard"
    isFilter     = $true
    fileType     = if ($ReportType -eq "docx" -or $ReportType -eq "pdf") { $ReportType } else { "xlsx" }
    reportFilter = @{}
    company_id   = $CompanyId
    company_name = $companyName
} | ConvertTo-Json

try {
    $resp = Invoke-RestMethod -Uri "$baseUrl/report_builder/create_report_job" -Method Post -Headers $headers -Body $body -ContentType "application/json"
    $jobId = $resp.message.job_id
    if (-not $jobId -and $resp.data) { $jobId = $resp.data.job_id }
    if ($jobId) {
        Write-Host "Job created: job_id=$jobId" -ForegroundColor Green
        $resp
    } else {
        Write-Host "Response:" -ForegroundColor Yellow
        $resp | ConvertTo-Json -Depth 5
    }
} catch {
    Write-Host "Error: $($_.Exception.Message)" -ForegroundColor Red
    if ($_.ErrorDetails.Message) { Write-Host $_.ErrorDetails.Message }
}
