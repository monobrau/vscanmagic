#Requires -Version 5.1
<#
.SYNOPSIS
Test script: fetches a small sample of raw vulnerability data from ConnectSecure and outputs CSV.
Standalone - does not dot-source ConnectSecure-API.ps1 (avoids scope/parsing issues).
#>

param(
    [int]$Limit = 3,
    [int]$CompanyId = 0,
    [string]$OutputDir = ".",
    [switch]$UseSavedCredentials
)

Add-Type -AssemblyName System.Web

$reportConfig = @{
    BaseUrl = "https://pod104.myconnectsecure.com"
    TenantName = "your-tenant-name"
    ClientId = "your-client-id"
    ClientSecret = "your-client-secret"
}

$credsPath = Join-Path $env:LOCALAPPDATA "VScanMagic\ConnectSecure-Credentials.json"

$needSavedCreds = $UseSavedCredentials -or
    [string]::IsNullOrWhiteSpace($reportConfig.BaseUrl) -or
    ($reportConfig.TenantName -like "*your-*") -or
    ($reportConfig.ClientId -like "*your-*")

if ($needSavedCreds -and (Test-Path $credsPath)) {
    try {
        $creds = Get-Content $credsPath -Raw | ConvertFrom-Json
        $reportConfig.BaseUrl = $creds.BaseUrl.ToString().Trim()
        $reportConfig.TenantName = $creds.TenantName.ToString().Trim()
        $reportConfig.ClientId = $creds.ClientId.ToString().Trim()
        $reportConfig.ClientSecret = $creds.ClientSecret.ToString().Trim()
        Write-Host "Loaded credentials from $credsPath" -ForegroundColor Green
    } catch {
        Write-Host "Failed to load credentials: $($_.Exception.Message)" -ForegroundColor Red
        exit 1
    }
}

foreach ($f in @('BaseUrl','TenantName','ClientId','ClientSecret')) {
    $v = $reportConfig[$f]
    if ([string]::IsNullOrWhiteSpace($v) -or $v -like "*your-*") {
        Write-Host "ERROR: Configure $f or use -UseSavedCredentials" -ForegroundColor Red
        exit 1
    }
}

# Auth
$baseUrl = $reportConfig.BaseUrl.TrimEnd('/')
$authStr = "$($reportConfig.TenantName)+$($reportConfig.ClientId):$($reportConfig.ClientSecret)"
$authBytes = [System.Text.Encoding]::UTF8.GetBytes($authStr)
$b64 = [System.Convert]::ToBase64String($authBytes)

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
if (-not $token) {
    Write-Host "Authentication failed." -ForegroundColor Red
    exit 1
}

Write-Host "Connected." -ForegroundColor Green
Write-Host ""

# Endpoints - test vulnerabilities_details (compact: one row per vuln)
$endpoints = @(
    "/r/report_queries/vulnerabilities_details"
)

$outPath = [System.IO.Path]::GetFullPath($OutputDir)
if (-not (Test-Path $outPath)) { New-Item -ItemType Directory -Path $outPath -Force | Out-Null }
Write-Host "Output: $outPath" -ForegroundColor Gray
Write-Host ""

$headers = @{
    "Authorization" = "Bearer $token"
    "X-USER-ID" = $userId
    "Accept" = "application/json"
}

$query = "limit=$Limit&skip=0"
if ($CompanyId -gt 0) { $query += "&company_id=$CompanyId" }

function ConvertTo-FlatRow {
    param($obj)
    $row = [ordered]@{}
    foreach ($prop in $obj.PSObject.Properties) {
        $v = $prop.Value
        if ($null -eq $v) { $row[$prop.Name] = "" }
        elseif ($v -is [Array] -or $v -is [System.Collections.IEnumerable]) {
            $arr = @($v)
            $str = ($arr | ForEach-Object { $_.ToString() }) -join ";"
            $row[$prop.Name] = $str
        }
        elseif ($v -is [PSCustomObject] -or $v -is [hashtable]) {
            try { $row[$prop.Name] = ($v | ConvertTo-Json -Compress -Depth 2).Replace('"','''') }
            catch { $row[$prop.Name] = $v.ToString() }
        }
        else { $row[$prop.Name] = $v.ToString() }
    }
    [PSCustomObject]$row
}

$successCount = 0
foreach ($ep in $endpoints) {
    $name = $ep -replace "^/r/|^/w/|^/d/", "" -replace "/", "_"
    Write-Host "$name..." -ForegroundColor Cyan
    try {
        $uri = "$baseUrl$ep`?$query"
        $response = Invoke-RestMethod -Uri $uri -Method Get -Headers $headers -TimeoutSec 90

        $data = $null
        if ($response.PSObject.Properties['status'] -and $response.PSObject.Properties['data']) {
            $data = $response.data
        } elseif ($response.PSObject.Properties['data']) {
            $data = $response.data
        } elseif ($response.PSObject.Properties['hits']) {
            $data = @()
            foreach ($h in $response.hits.hits) {
                if ($h.PSObject.Properties['_source']) { $data += $h._source }
                else { $data += $h }
            }
        }

        if (-not $data) {
            Write-Host "  No records" -ForegroundColor Yellow
            continue
        }

        $arr = @($data)
        $flat = @()
        foreach ($r in $arr) {
            if ($r.PSObject.Properties['_source']) { $flat += $r._source }
            else { $flat += $r }
        }

        $csvPath = Join-Path $outPath "$name.csv"
        $flatRows = foreach ($item in $flat) { ConvertTo-FlatRow -obj $item }
        $flatRows | ConvertTo-Csv -NoTypeInformation | Set-Content -Path $csvPath -Encoding UTF8
        Write-Host "  $($flat.Count) rows -> $csvPath" -ForegroundColor Green
        $successCount++
    } catch {
        Write-Host "  Error: $($_.Exception.Message)" -ForegroundColor Red
    }
}

Write-Host ""
Write-Host "Done. $successCount CSV(s) written to $outPath" -ForegroundColor Cyan
