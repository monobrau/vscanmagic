#Requires -Version 5.1
<#
.SYNOPSIS
Manually pulls a ConnectSecure report by job ID (e.g. after Report Builder 404).
Uses saved credentials from %LOCALAPPDATA%\VScanMagic\ConnectSecure-Credentials.json

.EXAMPLE
.\Get-ReportByJobId.ps1 -JobId "8096c25f-d0e5-4b77-9ea4-5f95cb55fbee" -CompanyId 15853
#>
param(
    [Parameter(Mandatory=$true)]
    [string]$JobId,
    [int]$CompanyId = 15853,
    [string]$OutputDir = ".",
    [switch]$UseSavedCredentials
)

$ErrorActionPreference = "Stop"

$credsPath = Join-Path $env:LOCALAPPDATA "VScanMagic\ConnectSecure-Credentials.json"
if (-not (Test-Path $credsPath)) {
    Write-Host "ERROR: No saved credentials at $credsPath" -ForegroundColor Red
    exit 1
}
$creds = Get-Content $credsPath -Raw | ConvertFrom-Json
$baseUrl = $creds.BaseUrl.ToString().TrimEnd('/')
$tenantName = $creds.TenantName.ToString().Trim()
$clientId = $creds.ClientId.ToString().Trim()
$clientSecret = $creds.ClientSecret.ToString().Trim()
Write-Host "[*] Loaded credentials. BaseUrl: $baseUrl" -ForegroundColor Cyan

# Auth
Write-Host "[*] Authenticating..." -ForegroundColor Cyan
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
Write-Host "[+] Authenticated. UserId: $userId" -ForegroundColor Green

$headers = @{
    "Authorization" = "Bearer $token"
    "X-USER-ID" = $userId
    "Accept" = "application/json"
    "Content-Type" = "application/json"
}

# 1. Try report_jobs_view/{id}
Write-Host "`n[*] GET /r/company/report_jobs_view/$JobId" -ForegroundColor Cyan
try {
    $statusUrl = "$baseUrl/r/company/report_jobs_view/$JobId"
    $jobResp = Invoke-RestMethod -Uri $statusUrl -Method Get -Headers $headers
    Write-Host "[+] Job status:" -ForegroundColor Green
    $jobResp | ConvertTo-Json -Depth 3 | Write-Host
} catch {
    $sc = $null
    if ($_.Exception.Response) { try { $sc = [int]$_.Exception.Response.StatusCode } catch {} }
    Write-Host "[-] report_jobs_view failed: HTTP $sc - $($_.Exception.Message)" -ForegroundColor Yellow
}

# 2. Try get_report_link (portal sends job_id as JSON array and NO company_id)
Write-Host "`n[*] GET report_builder/get_report_link (job_id=$JobId, isGlobal=false)" -ForegroundColor Cyan
$isGlobal = ($CompanyId -eq 0)
$jobIdArray = "[""$JobId""]"  # Portal format: ["uuid"]
$paramSets = @(
    @{ job_id = $jobIdArray; isGlobal = $isGlobal.ToString().ToLower() },
    @{ job_id = $JobId; isGlobal = $isGlobal.ToString().ToLower() },
    @{ job_id = $jobIdArray; isGlobal = $isGlobal.ToString().ToLower(); company_id = $CompanyId },
    @{ job_id = $JobId; isGlobal = $isGlobal.ToString().ToLower(); company_id = $CompanyId }
)
$downloadUrl = $null
$endpoints = @("/report_builder/get_report_link", "/r/report_builder/get_report_link")

foreach ($ep in $endpoints) {
    foreach ($qp in $paramSets) {
        $q = ($qp.GetEnumerator() | ForEach-Object { "$($_.Key)=$($_.Value)" }) -join "&"
        $uri = "$baseUrl$ep`?$q"
        try {
            $linkResp = Invoke-RestMethod -Uri $uri -Method Get -Headers $headers
            $url = $linkResp.message
            if ([string]::IsNullOrWhiteSpace($url) -and $linkResp.data) {
                $url = $linkResp.data.download_url
                if (-not $url) { $url = $linkResp.data.url }
                if (-not $url) { $url = $linkResp.data.link }
            }
            if (-not [string]::IsNullOrWhiteSpace($url) -and $url -match "https?://") {
                $downloadUrl = $url
                if (-not $downloadUrl.StartsWith("http")) {
                    $downloadUrl = $baseUrl.TrimEnd("/") + "/" + $downloadUrl.TrimStart("/")
                }
                Write-Host "[+] Report link obtained from $ep" -ForegroundColor Green
                Write-Host "    URL: $downloadUrl" -ForegroundColor Gray
                break
            }
        } catch {
            $sc = $null
            if ($_.Exception.Response) { try { $sc = [int]$_.Exception.Response.StatusCode } catch {} }
            Write-Host "[-] $ep : HTTP $sc - $($_.Exception.Message)" -ForegroundColor Yellow
        }
        if ($downloadUrl) { break }
    }
    if ($downloadUrl) { break }
}

# 3. Download if we have a URL
# Pre-signed S3/R2 URLs authenticate via query params - do NOT send Authorization header (causes 400)
if ($downloadUrl) {
    Write-Host "`n[*] Downloading report..." -ForegroundColor Cyan
    $outPath = [System.IO.Path]::GetFullPath($OutputDir)
    if (-not (Test-Path $outPath)) { New-Item -ItemType Directory -Path $outPath -Force | Out-Null }
    $ext = if ($downloadUrl -match '\.(xlsx|xls|docx|doc|pdf|zip)(?:\?|$)') { $Matches[1] } else { "xlsx" }
    $outFile = Join-Path $outPath "Report-$JobId-$((Get-Date).ToString('yyyyMMdd-HHmmss')).$ext"
    try {
        Invoke-WebRequest -Uri $downloadUrl -Method Get -OutFile $outFile -UseBasicParsing
        Write-Host "[+] Saved to: $outFile" -ForegroundColor Green
    } catch {
        Write-Host "[-] Download failed: $($_.Exception.Message)" -ForegroundColor Red
    }
} else {
    Write-Host "`n[-] No download URL obtained. Report may not be ready or endpoint may be unavailable." -ForegroundColor Red
}
