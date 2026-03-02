#Requires -Version 5.1
<#
.SYNOPSIS
Test script to verify discovery_settings server-side filtering.

.DESCRIPTION
Compares /r/report_queries/discovery_settings with condition vs baseline.
If condition returns fewer records than baseline, server-side filter works.
If same count, condition is ignored (pulling all, filtering locally).

.PARAMETER CompanyId
Company ID to test (e.g. 15373 for Acieta). Required.

.EXAMPLE
.\Test-DiscoverySettingsAPI.ps1 -CompanyId 15373
#>
param(
    [Parameter(Mandatory)]
    [int]$CompanyId
)

$ErrorActionPreference = 'Stop'

$scriptDir = if ($PSScriptRoot) { $PSScriptRoot } else { (Get-Location).Path }
$script:SettingsDirectory = Join-Path $env:LOCALAPPDATA "VScanMagic"
$script:ConnectSecureCredentialsPath = Join-Path $script:SettingsDirectory "ConnectSecure-Credentials.json"

$corePath = Join-Path $scriptDir "VScanMagic-Modules\VScanMagic-Core.ps1"
$apiPath = Join-Path $scriptDir "ConnectSecure-API.ps1"

if (-not (Test-Path $corePath)) { Write-Error "VScanMagic-Core.ps1 not found"; exit 1 }
if (-not (Test-Path $apiPath)) { Write-Error "ConnectSecure-API.ps1 not found"; exit 1 }

. $corePath
. $apiPath

Write-Host "`n=== Discovery Settings API Filter Test ===" -ForegroundColor Cyan
Write-Host "CompanyId: $CompanyId`n" -ForegroundColor Cyan

$creds = Load-ConnectSecureCredentials
if (-not $creds -or [string]::IsNullOrWhiteSpace($creds.BaseUrl) -or [string]::IsNullOrWhiteSpace($creds.ClientId)) {
    Write-Host "ERROR: No ConnectSecure credentials. Configure API Settings in VScanMagic first." -ForegroundColor Red
    exit 1
}

Write-Host "Connecting to $($creds.BaseUrl)..." -ForegroundColor Yellow
$connected = Connect-ConnectSecureAPI -BaseUrl $creds.BaseUrl -TenantName $creds.TenantName -ClientId $creds.ClientId -ClientSecret $creds.ClientSecret
if (-not $connected) {
    Write-Host "ERROR: Failed to authenticate." -ForegroundColor Red
    exit 1
}
Write-Host "Connected.`n" -ForegroundColor Green

$reportEndpoint = '/r/report_queries/discovery_settings'
$companyEndpoint = '/r/company/discovery_settings'

function Invoke-Test {
    param([string]$Label, [string]$Endpoint, [hashtable]$QueryParams)
    $qStr = ($QueryParams.GetEnumerator() | ForEach-Object { "$($_.Key)=$($_.Value)" }) -join "&"
    Write-Host "--- $Label ---" -ForegroundColor Yellow
    Write-Host "  Endpoint: $Endpoint"
    Write-Host "  Query: $qStr"
    try {
        $response = Invoke-ConnectSecureRequest -Endpoint $Endpoint -QueryParameters $QueryParams
        $data = $null
        if ($response -is [array]) { $data = $response }
        elseif ($response.data) { $data = $response.data }
        else { $data = $response }

        if ($null -eq $data) {
            Write-Host "  Result: null" -ForegroundColor Gray
            return $null
        } elseif ($data -is [string]) {
            Write-Host "  Result: STRING (likely error) len=$($data.Length)" -ForegroundColor Red
            Write-Host "  Preview: $($data.Substring(0, [Math]::Min(200, $data.Length)))..." -ForegroundColor Red
            return $null
        } elseif ($data -is [array]) {
            $count = $data.Count
            $color = if ($count -gt 0) { "Green" } else { "Gray" }
            Write-Host "  Result: $count records" -ForegroundColor $color
            if ($count -gt 0) {
                $first = $data[0]
                $keys = @($first.PSObject.Properties.Name)
                Write-Host "  Keys: $($keys -join ', ')" -ForegroundColor Gray
                $hasCid = $keys -contains 'company_id' -or $keys -contains 'companyId'
                Write-Host "  Has company_id/companyId: $hasCid" -ForegroundColor $(if ($hasCid) { "Green" } else { "Yellow" })
                $matchingCompany = ($data | Where-Object { ($_.company_id -eq $CompanyId) -or ($_.companyId -eq $CompanyId) }).Count
                Write-Host "  Records matching company $CompanyId : $matchingCompany" -ForegroundColor $(if ($matchingCompany -eq $count) { "Green" } elseif ($matchingCompany -gt 0) { "Yellow" } else { "Gray" })
                if ($first.company_id -or $first.companyId) { Write-Host "  Sample [0].company_id: $($first.company_id)$($first.companyId)" -ForegroundColor Gray }
            }
            return $count
        } else {
            Write-Host "  Result: $($data.GetType().Name)" -ForegroundColor Gray
            return $null
        }
    } catch {
        Write-Host "  FAIL: $($_.Exception.Message)" -ForegroundColor Red
        return $null
    }
    Write-Host ""
}

# 1. Baseline: report_queries with no filter (pulls all)
$baselineCount = Invoke-Test -Label "1. report_queries baseline (no filter)" -Endpoint $reportEndpoint -QueryParams @{ limit = 5000; skip = 0 }
Write-Host ""

# 2. Portal format: condition=company_id=X
$conditionCount = Invoke-Test -Label "2. report_queries + condition=company_id=$CompanyId (portal format)" -Endpoint $reportEndpoint -QueryParams @{ condition = "company_id=$CompanyId"; limit = 500; skip = 0; order_by = 'updated desc' }
Write-Host ""

# 3. Portal format with externalscan type
$externalscanCount = Invoke-Test -Label "3. report_queries + condition with externalscan type" -Endpoint $reportEndpoint -QueryParams @{ condition = "company_id=$CompanyId and discovery_settings_type='externalscan'"; limit = 500; skip = 0; order_by = 'updated desc' }
Write-Host ""

# 4. company endpoint (no filter) - for comparison
$companyCount = Invoke-Test -Label "4. /r/company/discovery_settings (no filter)" -Endpoint $companyEndpoint -QueryParams @{ limit = 5000; skip = 0 }
Write-Host ""

# Summary
Write-Host "=== Summary ===" -ForegroundColor Cyan
if ($null -ne $baselineCount -and $null -ne $conditionCount) {
    if ($conditionCount -lt $baselineCount) {
        Write-Host "SERVER-SIDE FILTER WORKS: condition returned $conditionCount vs baseline $baselineCount" -ForegroundColor Green
    } elseif ($conditionCount -eq $baselineCount -and $baselineCount -gt 0) {
        Write-Host "CONDITION IGNORED: same count ($conditionCount) - pulling all, filtering locally" -ForegroundColor Red
    } elseif ($conditionCount -eq 0 -and $baselineCount -gt 0) {
        Write-Host "Condition returned 0 - condition may be invalid or company has no discovery_settings" -ForegroundColor Yellow
    }
}
Write-Host "`nDone." -ForegroundColor Cyan
