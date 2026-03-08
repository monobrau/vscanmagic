#Requires -Version 5.1
<#
.SYNOPSIS
Tests pulling a randomly selected host's "last seen" date from ConnectSecure assets API.

.DESCRIPTION
Fetches assets, picks one at random, and displays its last_discovered_time (last seen).
Run from project root.

.PARAMETER CompanyId
Company ID to filter assets. 0 = all companies (if supported).

.PARAMETER Limit
Max assets to fetch (default 200). Higher = more to choose from.
#>
param([int]$CompanyId = 0, [int]$Limit = 200)

$ErrorActionPreference = 'Stop'

$scriptDir = if ($PSScriptRoot) { $PSScriptRoot } else { (Get-Location).Path }
$script:ScriptDirectory = $scriptDir

$corePath = Join-Path $scriptDir "VScanMagic-Modules\VScanMagic-Core.ps1"
if (-not (Test-Path $corePath)) {
    Write-Host "ERROR: VScanMagic-Core.ps1 not found. Run from project root." -ForegroundColor Red
    exit 1
}
. $corePath

$apiPath = Join-Path $scriptDir "ConnectSecure-API.ps1"
if (-not (Test-Path $apiPath)) {
    Write-Host "ERROR: ConnectSecure-API.ps1 not found." -ForegroundColor Red
    exit 1
}
. $apiPath

# Use same settings path as app (including custom SettingsDirectory)
Load-UserSettings | Out-Null

Write-Host "`n=== Test: Host Last Seen Date ===" -ForegroundColor Cyan

$creds = Load-ConnectSecureCredentials
if (-not $creds -or [string]::IsNullOrWhiteSpace($creds.BaseUrl) -or [string]::IsNullOrWhiteSpace($creds.ClientId)) {
    Write-Host "ERROR: No ConnectSecure credentials. Configure API Settings in VScanMagic first." -ForegroundColor Red
    exit 1
}

$connected = Connect-ConnectSecureAPI -BaseUrl $creds.BaseUrl -TenantName $creds.TenantName -ClientId $creds.ClientId -ClientSecret $creds.ClientSecret
if (-not $connected) {
    Write-Host "ERROR: Failed to authenticate." -ForegroundColor Red
    exit 1
}
Write-Host "Connected to $($creds.BaseUrl)`n" -ForegroundColor Green

Write-Host "Fetching assets (Limit: $Limit, CompanyId: $CompanyId)..." -ForegroundColor Yellow
$assets = Get-ConnectSecureAssets -CompanyId $CompanyId -Limit $Limit -FetchAll:$false
if (-not $assets -or $assets.Count -eq 0) {
    Write-Host "No assets returned." -ForegroundColor Red
    exit 1
}
Write-Host "Retrieved $($assets.Count) assets.`n" -ForegroundColor Green

# Pick one at random
$idx = Get-Random -Minimum 0 -Maximum $assets.Count
$host = $assets[$idx]

$name = if ($host.host_name) { $host.host_name } elseif ($host.hostname) { $host.hostname } elseif ($host.name) { $host.name } elseif ($host.asset_name) { $host.asset_name } elseif ($host.fqdn_name) { $host.fqdn_name } else { "Asset $($host.id)" }
$lastSeen = $host.last_discovered_time
$discovered = $host.discovered

Write-Host "Randomly selected host:" -ForegroundColor Cyan
Write-Host "  Name:        $name"
Write-Host "  Last Seen:   $(if ($lastSeen) { $lastSeen } else { '(not set)' })"
Write-Host "  First Seen:  $(if ($discovered) { $discovered } else { '(not set)' })"
Write-Host "  IP:          $(if ($host.ip) { $host.ip } else { '-' })"
Write-Host "  OS:          $(if ($host.os_name) { $host.os_name } else { '-' })"
Write-Host ""

if ($lastSeen) {
    Write-Host "SUCCESS: last_discovered_time is available for this host." -ForegroundColor Green
} else {
    Write-Host "NOTE: last_discovered_time is empty for this host. Try another run or different CompanyId." -ForegroundColor Yellow
}
