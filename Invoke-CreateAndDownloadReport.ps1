#Requires -Version 5.1
<#
.SYNOPSIS
Creates a report job via ConnectSecure API, polls for the download link, and saves the file.
Uses Swagger schema for create_report_job and multiple endpoint/body variants.

.PARAMETER CompanyId
Company ID (0 for all companies). Required for company-scoped reports.

.PARAMETER ReportType
all-vulnerabilities | suppressed-vulnerabilities | external-vulnerabilities | executive-summary | pending-epss

.PARAMETER MaxWaitSeconds
Max seconds to poll get_report_link before giving up.

.PARAMETER GenerateTopX
When set, after successfully downloading an XLSX vulnerability report, automatically generates
Top X Vulnerabilities Report (DOCX) using VScanMagic logic.
Requires VScanMagic-GUI.ps1 and ConnectSecure-API.ps1 in the same directory.

.PARAMETER TopCount
Number of top vulnerabilities to include (default 10). Used when -GenerateTopX is set.

.EXAMPLE
.\Invoke-CreateAndDownloadReport.ps1 -CompanyId 15853 -ReportType all-vulnerabilities
.\Invoke-CreateAndDownloadReport.ps1 -CompanyId 15853 -ReportType executive-summary -OutputDir .\Reports
.\Invoke-CreateAndDownloadReport.ps1 -CompanyId 15853 -ReportType all-vulnerabilities -GenerateTopX -TopCount 10
#>
param(
    [Parameter(Mandatory=$true)][int]$CompanyId,
    [string]$ReportType = "all-vulnerabilities",
    [string]$OutputDir = ".",
    [int]$MaxWaitSeconds = 300,
    [int]$PollIntervalSeconds = 3,
    [int]$InitialWaitSeconds = 15,
    [switch]$CreateOnly,
    [switch]$GenerateTopX,
    [int]$TopCount = 10
)

$ErrorActionPreference = "Stop"

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

# Report type mapping
$reportMap = @{
    'all-vulnerabilities'      = @{ id='f836d6a4e4d54ac6a9d2967254796373'; fmt='xlsx'; name='All Vulnerabilities Report' }
    'suppressed-vulnerabilities' = @{ id='1d091564830b44c485a0ddc35ace9ac6'; fmt='xlsx'; name='Suppressed Vulnerabilities' }
    'external-vulnerabilities' = @{ id='01beb6b930744e11b690bb9dc25118fb'; fmt='xlsx'; name='External Scan' }
    'executive-summary'        = @{ id='1cd4f45884264d15bee4173dc58b6a57'; fmt='docx'; name='Executive Summary Report' }
    'pending-epss'             = @{ id='85d4913c0dbc4fc782b858f0d27dd180'; fmt='xlsx'; name='Pending Remediation EPSS Score Reports' }
}
$r = $reportMap[$ReportType]
if (-not $r) {
    Write-Host "Unknown ReportType. Use: all-vulnerabilities, suppressed-vulnerabilities, external-vulnerabilities, executive-summary, pending-epss" -ForegroundColor Red
    exit 1
}

# Fetch real company name
$companyName = "Company $CompanyId"
if ($CompanyId -gt 0) {
    try {
        $companiesResp = Invoke-RestMethod -Uri "$baseUrl/r/company/companies?limit=5000" -Method Get -Headers $headers
        $companies = @()
        if ($companiesResp.data) { $companies = @($companiesResp.data) }
        $match = $companies | Where-Object { $cid = $_.id -or $_.company_id -or $_.companyId; [int]$cid -eq $CompanyId } | Select-Object -First 1
        if ($match) {
            $companyName = ($match.name -or $match.company_name -or $match.companyName -or '').ToString().Trim()
            if (-not $companyName) { $companyName = "Company $CompanyId" }
        }
    } catch { }
}

# Create report job - try multiple variants (Swagger schema + alternatives)
$reportId = $r.id
$reportName = $r.name
$reportFmt = $r.fmt

$reportNameCompact = ($reportName -replace '\s','')  # Portal uses "AllVulnerabilitiesReport" (no spaces)
if (-not $reportNameCompact) { $reportNameCompact = "Report" }
$createBodies = @(
    # Portal payload (reportType="Standard", isFilter=true, job_id as array in get_report_link)
    @{ reportId=$reportId; reportName=$reportNameCompact; reportType="Standard"; isFilter=$true; fileType=$reportFmt; reportFilter=@{}; company_id=$CompanyId; company_name=$companyName },
    # Swagger-style fallback
    @{ company_id=$CompanyId; company_name=$companyName; reportId=$reportId; reportName=$reportName; reportType=$reportFmt; fileType=$reportFmt; isFilter=$false; reportFilter=@{} },
    @{ company_id=$CompanyId; company_name=$companyName; report_id=$reportId; report_format=$reportFmt; report_name=$reportName; report_type=$reportFmt; file_type=$reportFmt; is_filter=$false; report_filter=@{} }
)

$createEndpoints = @("$baseUrl/report_builder/create_report_job", "$baseUrl/r/report_builder/create_report_job")
$jobId = $null

foreach ($ep in $createEndpoints) {
    foreach ($body in $createBodies) {
        try {
            $bodyJson = $body | ConvertTo-Json
            $resp = Invoke-RestMethod -Uri $ep -Method Post -Headers $headers -Body $bodyJson -ContentType "application/json"
            if ($resp.status -eq $false) {
                if ($resp.message -match 'Please Contact Support') {
                    Write-Host "  API returned 'Please Contact Support' - Report Builder may be disabled for your tenant." -ForegroundColor Yellow
                }
                continue
            }
            $jid = $null
            if ($resp.data) {
                $d = $resp.data
                if ($d.job_id) { $jid = $d.job_id }
                elseif ($d.jobId) { $jid = $d.jobId }
                elseif ($d.id) { $jid = $d.id }
                elseif ($d -is [string]) { $jid = $d }
            }
            if (-not $jid -and $resp.message) {
                $m = $resp.message
                if ($m -is [string] -and $m -match '^[a-fA-F0-9-]{36}$') { $jid = $m }
                elseif ($m.job_id) { $jid = $m.job_id }
                elseif ($m.jobId) { $jid = $m.jobId }
            }
            if (-not $jid) { $jid = $resp.job_id }
            if (-not $jid) { $jid = $resp.id }
            if ($jid) {
                $jobId = $jid.ToString()
                Write-Host "Job created: job_id=$jobId (via $ep)" -ForegroundColor Green
                break
            }
        } catch {
            Write-Host "  Create failed ($ep): $($_.Exception.Message)" -ForegroundColor Gray
            continue
        }
        if ($jobId) { break }
    }
    if ($jobId) { break }
}

if (-not $jobId) {
    Write-Host "create_report_job failed for all variants. Try creating a report in the ConnectSecure portal, then use:" -ForegroundColor Red
    Write-Host "  .\Get-ReportByJobId.ps1 -JobId <job_id_from_portal> -CompanyId $CompanyId" -ForegroundColor Yellow
    exit 1
}

if ($CreateOnly) {
    Write-Host "Job ID for manual download: $jobId" -ForegroundColor Cyan
    Write-Host "  .\Get-ReportByJobId.ps1 -JobId `"$jobId`" -CompanyId $CompanyId" -ForegroundColor Gray
    Write-Output $jobId
    exit 0
}

# Poll get_report_link
Write-Host "Polling get_report_link (max ${MaxWaitSeconds}s, interval ${PollIntervalSeconds}s)..." -ForegroundColor Cyan
Start-Sleep -Seconds $InitialWaitSeconds

$isGlobal = ($CompanyId -eq 0)
$jobIdArray = "[""$jobId""]"  # Portal sends job_id as JSON array, NO company_id
$paramSets = @(
    @{ job_id=$jobIdArray; isGlobal=$isGlobal.ToString().ToLower() },
    @{ job_id=$jobId; isGlobal=$isGlobal.ToString().ToLower() },
    @{ job_id=$jobIdArray; isGlobal=$isGlobal.ToString().ToLower(); company_id=$CompanyId },
    @{ job_id=$jobId; isGlobal=$isGlobal.ToString().ToLower(); company_id=$CompanyId }
)
$getLinkEndpoints = @("/report_builder/get_report_link", "/r/report_builder/get_report_link")
$downloadUrl = $null
$start = Get-Date
$attempt = 0

while ($true) {
    $elapsed = ((Get-Date) - $start).TotalSeconds
    if ($elapsed -ge $MaxWaitSeconds) {
        Write-Host "Timed out after $MaxWaitSeconds seconds. Try later:" -ForegroundColor Yellow
        Write-Host "  .\Get-ReportByJobId.ps1 -JobId `"$jobId`" -CompanyId $CompanyId" -ForegroundColor Gray
        break
    }
    $attempt++
    foreach ($ep in $getLinkEndpoints) {
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
                    Write-Host "Report ready after $([int]$elapsed)s" -ForegroundColor Green
                    break
                }
            } catch {
                $sc = $null
                if ($_.Exception.Response) { try { $sc = [int]$_.Exception.Response.StatusCode } catch {} }
                if ($sc -eq 404) {
                    if ($attempt % 5 -eq 1) { Write-Host "  Attempt $attempt : 404 (report may still be generating)" -ForegroundColor Gray }
                } elseif ($sc) {
                    Write-Host "  Attempt $attempt : HTTP $sc" -ForegroundColor Gray
                }
            }
        }
        if ($downloadUrl) { break }
    }
    if ($downloadUrl) { break }
    Start-Sleep -Seconds $PollIntervalSeconds
}

# Download
if ($downloadUrl) {
    $safeDir = [System.IO.Path]::GetFullPath($OutputDir)
    if (-not (Test-Path $safeDir)) { New-Item -ItemType Directory -Path $safeDir -Force | Out-Null }
    $ext = if ($downloadUrl -match '\.(xlsx|xls|docx|doc|pdf|zip)(?:\?|$)') { $Matches[1] } else { $reportFmt }
    $outFile = Join-Path $safeDir "Report-$jobId-$((Get-Date).ToString('yyyyMMdd-HHmmss')).$ext"
    try {
        $downloadHeaders = @{}
        if ($downloadUrl -notmatch 'r2\.cloudflarestorage|X-Amz-Signature') {
            $downloadHeaders['Authorization'] = "Bearer $token"
        }
        Invoke-WebRequest -Uri $downloadUrl -Method Get -Headers $downloadHeaders -OutFile $outFile -UseBasicParsing
        Write-Host "Saved: $outFile" -ForegroundColor Green

        # Optional: generate Top X report from XLSX
        $vulnTypes = @('all-vulnerabilities','external-vulnerabilities','suppressed-vulnerabilities')
        if ($GenerateTopX -and $ReportType -in $vulnTypes -and $ext -eq 'xlsx' -and (Test-Path -LiteralPath $outFile)) {
            $scriptDir = Split-Path -Parent $MyInvocation.MyCommand.Path
            $guiPath = Join-Path $scriptDir "VScanMagic-GUI.ps1"
            if (Test-Path $guiPath) {
                try {
                    Write-Host "Generating Top $TopCount Vulnerabilities Report..." -ForegroundColor Cyan
                    $script:IsApiMode = $true
                    . $guiPath
                    $vulnData = Get-VulnerabilityData -ExcelPath $outFile
                    if ($null -ne $vulnData -and $vulnData.Count -gt 0) {
                        $top10 = Get-Top10Vulnerabilities -VulnData $vulnData -Count $TopCount
                        $scanDate = (Get-Date).ToString("MM/dd/yyyy")
                        $ts = (Get-Date).ToString("yyyy-MM-dd_HH-mm-ss")
                        $safeName = $companyName -replace '[^\w\s\-]', '' -replace '\s+', ' '
                        $topXPath = Join-Path $safeDir "$safeName - Top $TopCount Vulnerabilities Report - $ts.docx"
                        New-WordReport -OutputPath $topXPath -ClientName $companyName -ScanDate $scanDate -Top10Data $top10 -TimeEstimates $null -IsRMITPlus $false -GeneralRecommendations @() -ReportTitle "Top $TopCount Vulnerabilities Report"
                        Write-Host "Generated: $topXPath" -ForegroundColor Green
                        Write-Output $topXPath
                    } else {
                        Write-Host "No vulnerability data in file - skipping Top X generation" -ForegroundColor Yellow
                    }
                } catch {
                    Write-Host "Top X generation failed: $($_.Exception.Message)" -ForegroundColor Yellow
                }
            } else {
                Write-Host "VScanMagic-GUI.ps1 not found - skipping Top X generation" -ForegroundColor Yellow
            }
        }

        Write-Output $outFile
    } catch {
        Write-Host "Download failed: $($_.Exception.Message)" -ForegroundColor Red
        exit 1
    }
} else {
    Write-Host "" -ForegroundColor Gray
    Write-Host "API create succeeds but get_report_link returns 404. In testing, API-created jobs do not" -ForegroundColor Yellow
    Write-Host "appear in report_jobs_view and never become available for download (ConnectSecure backend behavior)." -ForegroundColor Yellow
    Write-Host "" -ForegroundColor Gray
    Write-Host "WORKAROUND - use portal-created reports:" -ForegroundColor Cyan
    Write-Host "  1. Create the report in the ConnectSecure portal" -ForegroundColor Gray
    Write-Host "  2. Get the job_id from the download link or: .\Get-ReportJobs.ps1 -CompanyId $CompanyId" -ForegroundColor Gray
    Write-Host "  3. .\Get-ReportByJobId.ps1 -JobId <job_id> -CompanyId $CompanyId" -ForegroundColor Gray
    Write-Host "" -ForegroundColor Gray
    Write-Host "To request API report creation support: Contact ConnectSecure support." -ForegroundColor Gray
    exit 1
}
