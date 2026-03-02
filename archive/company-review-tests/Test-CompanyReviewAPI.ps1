#Requires -Version 5.1
<#
.SYNOPSIS
Test script for Company Review API endpoints. Verifies ConnectSecure API calls return data.

.DESCRIPTION
Loads credentials from VScanMagic settings, connects to ConnectSecure, and calls each
Company Review endpoint. Outputs raw response info to diagnose API integration issues.

.PARAMETER CompanyId
Company ID to test with. If not provided, uses first company from cache or saved CompanyId.

.EXAMPLE
.\Test-CompanyReviewAPI.ps1
.\Test-CompanyReviewAPI.ps1 -CompanyId 123
#>
param([int]$CompanyId = 0)

$ErrorActionPreference = 'Stop'

# Setup paths (match VScanMagic-GUI.ps1)
$scriptDir = if ($PSScriptRoot) { $PSScriptRoot } else { (Get-Location).Path }
$script:SettingsDirectory = Join-Path $env:LOCALAPPDATA "VScanMagic"
$script:ConnectSecureCredentialsPath = Join-Path $script:SettingsDirectory "ConnectSecure-Credentials.json"
$script:ConnectSecureCompaniesCachePath = Join-Path $script:SettingsDirectory "ConnectSecure-Companies-Cache.json"

# Load Core (credentials, cache)
$corePath = Join-Path $scriptDir "VScanMagic-Modules\VScanMagic-Core.ps1"
if (-not (Test-Path $corePath)) {
    Write-Host "ERROR: VScanMagic-Core.ps1 not found at $corePath" -ForegroundColor Red
    exit 1
}
. $corePath

# Load ConnectSecure API
$apiPath = Join-Path $scriptDir "ConnectSecure-API.ps1"
if (-not (Test-Path $apiPath)) {
    Write-Host "ERROR: ConnectSecure-API.ps1 not found at $apiPath" -ForegroundColor Red
    exit 1
}
. $apiPath

Write-Host "`n=== Company Review API Test ===" -ForegroundColor Cyan
Write-Host ""

# Load credentials and connect
$creds = Load-ConnectSecureCredentials
if (-not $creds -or [string]::IsNullOrWhiteSpace($creds.BaseUrl) -or [string]::IsNullOrWhiteSpace($creds.ClientId)) {
    Write-Host "ERROR: No ConnectSecure credentials. Configure API Settings in VScanMagic first." -ForegroundColor Red
    exit 1
}

Write-Host "Connecting to $($creds.BaseUrl) (tenant: $($creds.TenantName))..." -ForegroundColor Yellow
$connected = Connect-ConnectSecureAPI -BaseUrl $creds.BaseUrl -TenantName $creds.TenantName -ClientId $creds.ClientId -ClientSecret $creds.ClientSecret
if (-not $connected) {
    Write-Host "ERROR: Failed to authenticate." -ForegroundColor Red
    exit 1
}
Write-Host "Connected OK.`n" -ForegroundColor Green

# Resolve company ID
if ($CompanyId -le 0) {
    $CompanyId = $creds.CompanyId
    if ($CompanyId -le 0) {
        $cached = Load-ConnectSecureCompaniesCache -BaseUrl $creds.BaseUrl -TenantName $creds.TenantName
        if ($cached -and $cached.Count -gt 0) {
            $first = $cached | Where-Object { $_.Id -ne 0 } | Select-Object -First 1
            if ($first) { $CompanyId = $first.Id }
        }
    }
}

if ($CompanyId -le 0) {
    Write-Host "ERROR: No company ID. Provide -CompanyId or select a company in VScanMagic and refresh." -ForegroundColor Red
    exit 1
}

Write-Host "Testing with CompanyId: $CompanyId`n" -ForegroundColor Cyan

function Test-Endpoint {
    param([string]$Name, [string]$Endpoint, [hashtable]$QueryParams = @{})
    Write-Host "--- $Name ---" -ForegroundColor Yellow
    Write-Host "  GET $Endpoint"
    if ($QueryParams.Count -gt 0) {
        $qStr = ($QueryParams.GetEnumerator() | ForEach-Object { "$($_.Key)=$($_.Value)" }) -join "&"
        Write-Host "  Query: $qStr"
    }
    try {
        $response = Invoke-ConnectSecureRequest -Endpoint $Endpoint -QueryParameters $QueryParams
        $data = $null
        if ($response -is [array]) { $data = $response }
        elseif ($response.data) { $data = $response.data }
        else { $data = $response }

        $dataType = if ($null -eq $data) { "null" } else { $data.GetType().Name }
        if ($data -is [array]) {
            $count = $data.Count
            Write-Host "  OK: $count records (array)" -ForegroundColor Green
            if ($count -gt 0 -and $count -le 2) {
                $data[0].PSObject.Properties | ForEach-Object { Write-Host "    [0].$($_.Name) = $($_.Value)" }
            }
        } elseif ($data -is [string]) {
            Write-Host "  WARN: data is string (len=$($data.Length)), not array. Raw: $($data.Substring(0, [Math]::Min(80, $data.Length)))..." -ForegroundColor Yellow
        } elseif ($data -and $data.PSObject.Properties) {
            $propCount = @($data.PSObject.Properties).Count
            Write-Host "  OK: object ($dataType) with $propCount properties" -ForegroundColor Green
            $data.PSObject.Properties | Select-Object -First 5 | ForEach-Object { Write-Host "    $($_.Name) = $($_.Value)" }
        } else {
            Write-Host "  OK: $dataType (empty or scalar)" -ForegroundColor Gray
        }
    } catch {
        Write-Host "  FAIL: $($_.Exception.Message)" -ForegroundColor Red
    }
    Write-Host ""
}

# Test each endpoint - NOTE: condition param causes API to return error string; use no condition, filter client-side
$endpoints = @{
    LightweightAssets = @{ Endpoint = '/r/report_queries/lightweight_assets'; Params = @{ company_id = $CompanyId; limit = 500; skip = 0 } }
    CompanyStats = @{ Endpoint = '/r/company/company_stats'; Params = @{ limit = 500; skip = 0 } }
    Agents = @{ Endpoint = '/r/company/agents'; Params = @{ limit = 500; skip = 0 } }
    Credentials = @{ Endpoint = '/r/company/credentials'; Params = @{ limit = 500; skip = 0 } }
    AgentCredMapping = @{ Endpoint = '/r/company/agent_credentials_mapping'; Params = @{ limit = 500; skip = 0 } }
    DiscoverySettings = @{ Endpoint = '/r/company/discovery_settings'; Params = @{ limit = 500; skip = 0 } }
    AgentDiscoveryMapping = @{ Endpoint = '/r/company/agent_discoverysettings_mapping'; Params = @{ limit = 500; skip = 0 } }
    AssetFirewallPolicy = @{ Endpoint = '/r/asset/asset_firewall_policy'; Params = @{ limit = 500; skip = 0 } }
}

foreach ($key in $endpoints.Keys) {
    $e = $endpoints[$key]
    Test-Endpoint -Name $key -Endpoint $e.Endpoint -QueryParams $e.Params
}

# Summary: filter company $CompanyId from each result
Write-Host "--- Client-side filter check (company_id=$CompanyId) ---" -ForegroundColor Yellow
try {
    $agents = Invoke-ConnectSecureRequest -Endpoint '/r/company/agents' -QueryParameters @{ limit = 2000; skip = 0 }
    $data = if ($agents.data) { $agents.data } else { $agents }
    $companyAgents = $data | Where-Object { ($_.company_id -eq $CompanyId) -or ($_.companyId -eq $CompanyId) }
    Write-Host "  Agents: $($data.Count) total, $($companyAgents.Count) for company $CompanyId" -ForegroundColor $(if ($companyAgents.Count -gt 0) { "Green" } else { "Gray" })
} catch { Write-Host "  Agents filter: $($_.Exception.Message)" -ForegroundColor Red }

Write-Host "`n=== Test complete ===" -ForegroundColor Cyan
