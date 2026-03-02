#Requires -Version 5.1
<#
.SYNOPSIS
Test script to find which API parameters filter external_asset_externalscan by company.

.DESCRIPTION
Tries company_id, condition, and other query params against /r/report_queries/external_asset_externalscan.
Reports count and sample for each attempt so you can see what works.

.PARAMETER CompanyId
Company ID to test (e.g. 15373 for Acieta). Required.

.EXAMPLE
.\Test-ExternalAssetsAPI.ps1 -CompanyId 15373
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

Write-Host "`n=== External Assets API Filter Test ===" -ForegroundColor Cyan
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

$endpoint = '/r/report_queries/external_asset_externalscan'

function Invoke-Test {
    param([string]$Label, [hashtable]$QueryParams)
    $qStr = ($QueryParams.GetEnumerator() | ForEach-Object { "$($_.Key)=$($_.Value)" }) -join "&"
    Write-Host "--- $Label ---" -ForegroundColor Yellow
    Write-Host "  Query: $qStr"
    try {
        $response = Invoke-ConnectSecureRequest -Endpoint $endpoint -QueryParameters $QueryParams
        $data = $null
        if ($response -is [array]) { $data = $response }
        elseif ($response.data) { $data = $response.data }
        else { $data = $response }

        if ($null -eq $data) {
            Write-Host "  Result: null" -ForegroundColor Gray
        } elseif ($data -is [string]) {
            Write-Host "  Result: STRING (likely error) len=$($data.Length)" -ForegroundColor Red
            Write-Host "  Preview: $($data.Substring(0, [Math]::Min(120, $data.Length)))..." -ForegroundColor Red
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
                Write-Host "  Sample [0]: name=$($first.name) config_name=$($first.config_name) ip=$($first.ip) id=$($first.id)" -ForegroundColor Gray
                if ($hasCid -and $first.company_id) { Write-Host "  Sample [0].company_id: $($first.company_id)" -ForegroundColor Gray }
            }
        } else {
            Write-Host "  Result: $($data.GetType().Name)" -ForegroundColor Gray
        }
    } catch {
        Write-Host "  FAIL: $($_.Exception.Message)" -ForegroundColor Red
    }
    Write-Host ""
}

# Baseline: no filter
Invoke-Test -Label "1. Baseline (no filter)" -QueryParams @{ limit = 100; skip = 0 }

# Server-side attempts
Invoke-Test -Label "2. company_id (query param)" -QueryParams @{ company_id = $CompanyId; limit = 100; skip = 0 }
Invoke-Test -Label "3. condition=company_ids:X" -QueryParams @{ condition = "company_ids:$CompanyId"; limit = 100; skip = 0 }
Invoke-Test -Label "4. condition=company_id:X" -QueryParams @{ condition = "company_id:$CompanyId"; limit = 100; skip = 0 }
Invoke-Test -Label "4b. condition=company_id=X (portal format)" -QueryParams @{ condition = "company_id=$CompanyId"; limit = 100; skip = 0 }
Invoke-Test -Label "5. condition=company_id eq X" -QueryParams @{ condition = "company_id eq $CompanyId"; limit = 100; skip = 0 }
Invoke-Test -Label "6. condition=company_ids eq X" -QueryParams @{ condition = "company_ids eq $CompanyId"; limit = 100; skip = 0 }

# Compare: lightweight_assets with company_id (known to work)
Write-Host "--- Reference: lightweight_assets with company_id ---" -ForegroundColor Cyan
try {
    $lw = Invoke-ConnectSecureRequest -Endpoint '/r/report_queries/lightweight_assets' -QueryParameters @{ company_id = $CompanyId; limit = 10; skip = 0 }
    $lwData = if ($lw.data) { $lw.data } else { $lw }
    $lwCount = if ($lwData -is [array]) { $lwData.Count } else { 0 }
    Write-Host "  lightweight_assets + company_id=$CompanyId : $lwCount records" -ForegroundColor Green
} catch { Write-Host "  FAIL: $($_.Exception.Message)" -ForegroundColor Red }
Write-Host ""

Write-Host "=== Done ===" -ForegroundColor Cyan
Write-Host "If any attempt returns records > 0, that parameter combination works for server-side filtering." -ForegroundColor Gray
