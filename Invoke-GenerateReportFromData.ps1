#Requires -Version 5.1
<#
.SYNOPSIS
Generates ConnectSecure reports from data APIs - no portal, no create_report_job.
Fetches vulnerability data via /r/report_queries/* and exports to CSV/XLSX locally.
Uses Excel COM for XLSX when available; falls back to CSV otherwise.

.PARAMETER CompanyId
Company ID (required for company-scoped reports).

.PARAMETER ReportType
all-vulnerabilities | suppressed-vulnerabilities | external-vulnerabilities | executive-summary | pending-epss

.EXAMPLE
.\Invoke-GenerateReportFromData.ps1 -CompanyId 15853 -ReportType suppressed-vulnerabilities
.\Invoke-GenerateReportFromData.ps1 -CompanyId 15853 -ReportType all-vulnerabilities -OutputDir .\Reports
#>
param(
    [Parameter(Mandatory=$true)][int]$CompanyId,
    [switch]$DebugFetch,
    [string]$ReportType = "all-vulnerabilities",
    [string]$OutputDir = ".",
    [int]$TopCount = 10,
    [switch]$ForceCsv
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
}

# Endpoints per report type
$endpointMap = @{
    'all-vulnerabilities'      = '/r/report_queries/vulnerabilities_details'
    'suppressed-vulnerabilities' = '/r/report_queries/vulnerabilities_details_suppressed'
    'external-vulnerabilities' = '/r/report_queries/external_asset_vulnerabilities'
    'pending-epss'             = '/r/report_queries/vulnerabilities_details'
    'executive-summary'        = '/r/report_queries/vulnerabilities_details'
}

$ep = $endpointMap[$ReportType]
if (-not $ep) {
    Write-Host "Unknown ReportType. Use: all-vulnerabilities, suppressed-vulnerabilities, external-vulnerabilities, executive-summary, pending-epss" -ForegroundColor Red
    exit 1
}

# Fetch company name
$clientName = "Company $CompanyId"
try {
    $companiesResp = Invoke-RestMethod -Uri "$baseUrl/r/company/companies?limit=5000" -Method Get -Headers $headers
    $companies = @()
    if ($companiesResp.data) { $companies = @($companiesResp.data) }
    foreach ($c in $companies) {
        $cid = $c.id -or $c.company_id -or $c.companyId
        if ([int]$cid -eq $CompanyId) {
            $clientName = ($c.name -or $c.company_name -or $c.companyName -or '').ToString().Trim()
            if (-not $clientName) { $clientName = "Company $CompanyId" }
            break
        }
    }
} catch { }

# Fetch data - paginate to get all
$allData = @()
$skip = 0
$limit = 5000
$maxFetch = 50000
while ($true) {
    $uri = "$baseUrl$ep`?limit=$limit&skip=$skip&sort=severity.keyword:desc"
    if ($CompanyId -gt 0) { $uri += "&company_id=$CompanyId" }
    $resp = Invoke-RestMethod -Uri $uri -Method Get -Headers $headers
    $chunk = @()
    if ($resp.data) { $chunk = @($resp.data) }
    elseif ($resp.hits -and $resp.hits.hits) {
        $chunk = $resp.hits.hits | ForEach-Object { if ($_._source) { $_._source } else { $_ } }
    }
    if ($chunk.Count -eq 0) {
        if ($DebugFetch -and $allData.Count -eq 0) {
            Write-Host "  First request returned 0 rows. Response keys: $($resp.PSObject.Properties.Name -join ', ')" -ForegroundColor Yellow
        }
        break
    }
    # API ignores company_id param - filter client-side by company_ids
    if ($CompanyId -gt 0) {
        $cidStr = [string]$CompanyId
        $chunk = $chunk | Where-Object {
            $rowCids = $_.company_ids -or $_.companyIds -or $_.company_id
            if ($null -eq $rowCids) { $false }
            elseif ($rowCids -is [array]) { $rowCids -contains $CompanyId }
            else { [string]$rowCids -match "(^|;)\s*$cidStr\s*(;|$)" }
        }
    }
    $allData += $chunk
    if ($chunk.Count -lt $limit -or $allData.Count -ge $maxFetch) { break }
    $skip += $limit
    Write-Host "  Fetched $($allData.Count) rows..." -ForegroundColor Gray
}

if ($ReportType -eq 'executive-summary' -and $allData.Count -gt $TopCount) {
    $allData = $allData | Select-Object -First $TopCount
}

Write-Host "Generating $ReportType for $clientName ($($allData.Count) rows)..." -ForegroundColor Cyan

$safeDir = [System.IO.Path]::GetFullPath($OutputDir)
if (-not (Test-Path $safeDir)) { New-Item -ItemType Directory -Path $safeDir -Force | Out-Null }
$ts = (Get-Date).ToString('yyyy-MM-dd_HHmmss')
if ($ForceCsv -or $ReportType -eq 'executive-summary') {
    $outFile = Join-Path $safeDir "$clientName-$ReportType-$ts.csv"
    $outFile = $outFile -replace '[^\w\s\-\.]', '_'
    if ($allData.Count -gt 0) {
        $props = @()
        foreach ($p in $allData[0].PSObject.Properties) { $props += $p.Name }
        $allData | Select-Object -Property $props | Export-Csv -Path $outFile -NoTypeInformation
    } else {
        Set-Content -Path $outFile -Value "No data"
    }
    Write-Host "Saved: $outFile" -ForegroundColor Green
    Write-Output $outFile
    exit 0
}

# XLSX via Excel COM
try {
    $excel = New-Object -ComObject Excel.Application
    $excel.Visible = $false
    $excel.DisplayAlerts = $false
    $workbook = $excel.Workbooks.Add()
    $sheet = $workbook.Sheets.Item(1)
    $sheet.Name = $ReportType -replace '[^\w]',' '

    if ($allData.Count -gt 0) {
        $props = @($allData[0].PSObject.Properties | Select-Object -ExpandProperty Name)
        for ($c = 1; $c -le $props.Count; $c++) {
            $sheet.Cells.Item(1, $c).Value2 = $props[$c - 1]
        }
        $row = 2
        foreach ($obj in $allData) {
            for ($c = 1; $c -le $props.Count; $c++) {
                $val = $obj.($props[$c - 1])
                if ($null -ne $val) {
                    if ($val -is [array]) { $val = ($val | ForEach-Object { $_ }) -join '; ' }
                    $sheet.Cells.Item($row, $c).Value2 = $val.ToString()
                }
            }
            $row++
        }
    }
    $outFile = Join-Path $safeDir "$clientName-$ReportType-$ts.xlsx"
    $outFile = $outFile -replace '[^\w\s\-\.]', '_'
    $workbook.SaveAs($outFile)
    $workbook.Close($false)
    $excel.Quit()
    [System.Runtime.Interopservices.Marshal]::ReleaseComObject($sheet) | Out-Null
    [System.Runtime.Interopservices.Marshal]::ReleaseComObject($workbook) | Out-Null
    [System.Runtime.Interopservices.Marshal]::ReleaseComObject($excel) | Out-Null
    Write-Host "Saved: $outFile" -ForegroundColor Green
    Write-Output $outFile
} catch {
    Write-Host "Excel unavailable ($($_.Exception.Message)), saving as CSV..." -ForegroundColor Yellow
    $outFile = Join-Path $safeDir "$clientName-$ReportType-$ts.csv"
    $outFile = $outFile -replace '[^\w\s\-\.]', '_'
    if ($allData.Count -gt 0) {
        $props = @($allData[0].PSObject.Properties | Select-Object -ExpandProperty Name)
        $allData | Select-Object -Property $props | Export-Csv -Path $outFile -NoTypeInformation
    } else {
        Set-Content -Path $outFile -Value "No data"
    }
    Write-Host "Saved: $outFile" -ForegroundColor Green
    Write-Output $outFile
}
