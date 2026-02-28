#Requires -Version 5.1
<#
.SYNOPSIS
Tests ConnectSecure Report Builder API - creates a report job, polls for link, downloads.
Use this to investigate whether Report Builder works and where it fails.

.EXAMPLE
.\Test-ReportBuilder.ps1 -UseSavedCredentials
.\Test-ReportBuilder.ps1 -CompanyId 123
#>
param(
    [int]$CompanyId = 0,
    [string]$OutputDir = ".",
    [switch]$UseSavedCredentials,
    [switch]$SkipDownload,
    [int]$MaxWaitSeconds = 120,
    [int]$PollIntervalSeconds = 5
)

$ErrorActionPreference = "Stop"
Add-Type -AssemblyName System.Web

function Write-Step { param($n, $msg) Write-Host "`n[$n] $msg" -ForegroundColor Cyan }
function Write-Ok   { param($msg) Write-Host "  OK: $msg" -ForegroundColor Green }
function Write-Warn { param($msg) Write-Host "  WARN: $msg" -ForegroundColor Yellow }
function Write-Err  { param($msg) Write-Host "  FAIL: $msg" -ForegroundColor Red }
function Write-Detail { param($msg) Write-Host "  $msg" -ForegroundColor Gray }

# --- Credentials ---
$credsPath = Join-Path $env:LOCALAPPDATA "VScanMagic\ConnectSecure-Credentials.json"
if (-not (Test-Path $credsPath)) {
    Write-Err "No saved credentials. Run VScanMagic GUI, save credentials, then run this script."
    exit 1
}
$creds = Get-Content $credsPath -Raw | ConvertFrom-Json
$baseUrl = $creds.BaseUrl.ToString().TrimEnd('/')
$tenantName = $creds.TenantName.ToString().Trim()
$clientId = $creds.ClientId.ToString().Trim()
$clientSecret = $creds.ClientSecret.ToString().Trim()
Write-Ok "Loaded credentials. BaseUrl: $baseUrl"

# --- Auth ---
Write-Step 1 "Authenticating..."
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
if (-not $token) { Write-Err "Auth failed"; exit 1 }
Write-Ok "Token obtained. UserId: $userId"

$headers = @{
    "Authorization" = "Bearer $token"
    "X-USER-ID" = $userId
    "Accept" = "application/json"
    "Content-Type" = "application/json"
}

# --- Step 2: Standard Reports ---
Write-Step 2 "GET /report_builder/standard_reports?isGlobal=false"
$standardReportsUrl = "$baseUrl/report_builder/standard_reports?isGlobal=false"
try {
    $reportsResp = Invoke-RestMethod -Uri $standardReportsUrl -Method Get -Headers $headers
    $reports = $null
    if ($reportsResp.message) {
        $msg = $reportsResp.message
        if ($msg -is [Array]) {
            foreach ($m in $msg) {
                if ($m.Reports) {
                    foreach ($r in $m.Reports) {
                        if ($r.reports) { $reports = $r.reports; break }
                    }
                }
            }
        }
    }
    if (-not $reports -and $reportsResp.data) { $reports = $reportsResp.data }

    if ($reports -and $reports.Count -gt 0) {
        Write-Ok "standard_reports returned $($reports.Count) report(s)"
        foreach ($r in $reports) {
            Write-Detail "  id=$($r.id) reportType=$($r.reportType) title=$($r.reportName)"
        }
        # Prefer xlsx report; fall back to first
        $xlsxReport = $reports | Where-Object { $_.reportType -eq "xlsx" } | Select-Object -First 1
        $reportId = if ($xlsxReport -and $xlsxReport.id) { $xlsxReport.id } else { $reports[0].id }
        if (-not $reportId) { $reportId = "f836d6a4e4d54ac6a9d2967254796373" }
        Write-Detail "Using report id=$reportId"
    } else {
        Write-Warn "standard_reports returned empty. Using known All Vulnerabilities ID."
        $reportId = "f836d6a4e4d54ac6a9d2967254796373"
    }
} catch {
    Write-Warn "standard_reports failed: $($_.Exception.Message)"
    Write-Detail "Using known All Vulnerabilities report ID."
    $reportId = "f836d6a4e4d54ac6a9d2967254796373"
}

# Report IDs to try (in case first fails)
$reportIdsToTry = @($reportId)
if ($reportId -ne "f836d6a4e4d54ac6a9d2967254796373") {
    $reportIdsToTry += "f836d6a4e4d54ac6a9d2967254796373"  # All Vulnerabilities (known ID)
}

# --- Step 3: Create Report Job ---
Write-Step 3 "POST /report_builder/create_report_job"
$createUrl = "$baseUrl/report_builder/create_report_job"
$jobId = $null
foreach ($rid in $reportIdsToTry) {
    $createBody = @{
        company_id = $CompanyId
        report_format = "xlsx"
        report_id = $rid
    } | ConvertTo-Json
    Write-Detail "Trying report_id=$rid"
    try {
        $createResp = Invoke-RestMethod -Uri $createUrl -Method Post -Headers $headers -Body $createBody
        if ($createResp.status -eq $false -or -not $createResp.status) {
            Write-Warn "API returned status=false: $($createResp.message)"
            continue
        }
        $jid = $null
        if ($createResp.data) {
            $d = $createResp.data
            if ($d.job_id) { $jid = $d.job_id }
            elseif ($d.id) { $jid = $d.id }
            elseif ($d.jobId) { $jid = $d.jobId }
            elseif ($d -is [string]) { $jid = $d }
        }
        if (-not $jid) { $jid = $createResp.job_id }
        if (-not $jid) { $jid = $createResp.id }
        if (-not $jid -and $createResp.message) {
            if ($createResp.message -match "^\d+$") { $jid = $createResp.message }
            elseif ($createResp.message.job_id) { $jid = $createResp.message.job_id }
        }
        if ($jid) {
            $jobId = $jid.ToString()
            Write-Ok "Job created: job_id=$jobId (report_id=$rid)"
            break
        }
        Write-Warn "No job_id in response"
    } catch {
        Write-Warn "create_report_job failed for $rid : $($_.Exception.Message)"
    }
}
if (-not $jobId) {
    Write-Err "create_report_job failed for all report IDs"
    exit 1
}

# --- Step 4: Poll get_report_link ---
Write-Step 4 "Polling get_report_link (max ${MaxWaitSeconds}s, interval ${PollIntervalSeconds}s)"
$isGlobal = ($CompanyId -eq 0)
$getLinkEndpoints = @("/report_builder/get_report_link", "/r/report_builder/get_report_link")
$paramSets = @(
    @{ job_id = $jobId; isGlobal = $isGlobal.ToString().ToLower() },
    @{ jobId = $jobId; isGlobal = $isGlobal.ToString().ToLower() }
)
if (-not $isGlobal -and $CompanyId -gt 0) {
    foreach ($p in $paramSets) { $p["company_id"] = $CompanyId }
}

$downloadUrl = $null
$start = Get-Date
$attempt = 0
while ($true) {
    $elapsed = ((Get-Date) - $start).TotalSeconds
    if ($elapsed -ge $MaxWaitSeconds) {
        Write-Err "Timed out after $MaxWaitSeconds seconds"
        exit 1
    }

    $attempt++
    foreach ($getLinkEndpoint in $getLinkEndpoints) {
        foreach ($qp in $paramSets) {
            $q = ($qp.GetEnumerator() | ForEach-Object { "$($_.Key)=$($_.Value)" }) -join "&"
            $uri = "$baseUrl$getLinkEndpoint`?$q"
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
                Write-Ok "Report ready after $([int]$elapsed)s ($attempt attempt(s))"
                break
            }
            } catch {
                $statusCode = $null
                if ($_.Exception.Response) { try { $statusCode = [int]$_.Exception.Response.StatusCode } catch {} }
                if ($statusCode -eq 404) {
                    Write-Detail "Attempt $attempt : 404 Not Found (report may still be generating)"
                } elseif ($statusCode) {
                    Write-Warn "Attempt $attempt : HTTP $statusCode - $($_.Exception.Message)"
                }
            }
        }
        if ($downloadUrl) { break }
    }
    if ($downloadUrl) { break }

    Write-Detail "Waiting ${PollIntervalSeconds}s..."
    Start-Sleep -Seconds $PollIntervalSeconds
}

# --- Step 5: Download ---
if (-not $SkipDownload -and $downloadUrl) {
    Write-Step 5 "Downloading report"
    $outPath = [System.IO.Path]::GetFullPath($OutputDir)
    if (-not (Test-Path $outPath)) { New-Item -ItemType Directory -Path $outPath -Force | Out-Null }
    $outFile = Join-Path $outPath "ReportBuilder-Test-$((Get-Date).ToString('yyyyMMdd-HHmmss')).xlsx"
    try {
        Invoke-WebRequest -Uri $downloadUrl -Method Get -Headers @{ "Authorization" = "Bearer $token" } -OutFile $outFile -UseBasicParsing
        Write-Ok "Saved to: $outFile"
    } catch {
        Write-Err "Download failed: $($_.Exception.Message)"
    }
} elseif ($SkipDownload) {
    Write-Detail "Skipping download (-SkipDownload)"
}

Write-Host "`n=== Report Builder Test Complete ===" -ForegroundColor Cyan
if ($downloadUrl) {
    Write-Ok "Report Builder is working."
} else {
    Write-Warn "Report Builder did not succeed."
    Write-Detail "If create_report_job returned 'Please Contact Support', Report Builder may be disabled"
    Write-Detail "for your tenant. Contact ConnectSecure support to request Report Builder enablement."
}
