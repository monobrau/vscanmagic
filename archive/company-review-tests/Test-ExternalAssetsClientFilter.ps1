#Requires -Version 5.1
<#
.SYNOPSIS
Test client-side filtering of external assets by company.
Validates: assets -> company filter -> external_asset type -> ID match with external_asset_externalscan.

.PARAMETER CompanyId
Company ID to filter for (e.g. 15373).

.EXAMPLE
.\Test-ExternalAssetsClientFilter.ps1 -CompanyId 15373
#>
param(
    [Parameter(Mandatory)]
    [int]$CompanyId
)

$ErrorActionPreference = 'Stop'

$scriptDir = if ($PSScriptRoot) { $PSScriptRoot } else { (Get-Location).Path }
$script:SettingsDirectory = Join-Path $env:LOCALAPPDATA "VScanMagic"

$corePath = Join-Path $scriptDir "VScanMagic-Modules\VScanMagic-Core.ps1"
$apiPath = Join-Path $scriptDir "ConnectSecure-API.ps1"

if (-not (Test-Path $corePath)) { Write-Error "VScanMagic-Core.ps1 not found"; exit 1 }
if (-not (Test-Path $apiPath)) { Write-Error "ConnectSecure-API.ps1 not found"; exit 1 }

. $corePath
. $apiPath

Write-Host "`n=== External Assets Client-Side Filter Test ===" -ForegroundColor Cyan
Write-Host "CompanyId: $CompanyId`n" -ForegroundColor Cyan

$creds = Load-ConnectSecureCredentials
if (-not $creds -or [string]::IsNullOrWhiteSpace($creds.BaseUrl)) {
    Write-Host "ERROR: No ConnectSecure credentials." -ForegroundColor Red
    exit 1
}

$connected = Connect-ConnectSecureAPI -BaseUrl $creds.BaseUrl -TenantName $creds.TenantName -ClientId $creds.ClientId -ClientSecret $creds.ClientSecret
if (-not $connected) { Write-Host "ERROR: Auth failed." -ForegroundColor Red; exit 1 }
Write-Host "Connected.`n" -ForegroundColor Green

$cidStr = [string]$CompanyId

# Step 1: Fetch assets
Write-Host "--- Step 1: Fetch assets ---" -ForegroundColor Yellow
$assets = @()
try {
    $assets = Get-ConnectSecureAssets -Limit 5000 -FetchAll:$true
    Write-Host "  Total assets: $($assets.Count)" -ForegroundColor Green
} catch {
    Write-Host "  FAIL: $($_.Exception.Message)" -ForegroundColor Red
    exit 1
}

if ($assets.Count -eq 0) {
    Write-Host "  No assets. Cannot proceed." -ForegroundColor Red
    exit 1
}

# Step 2: Company distribution
Write-Host "`n--- Step 2: Company distribution in assets ---" -ForegroundColor Yellow
$byCid = @{}
$noCid = 0
foreach ($a in $assets) {
    $c = $a.company_id; if (-not $c) { $c = $a.companyId }
    if (-not $c -and $a.company_ids) {
        $arr = $a.company_ids
        if ($arr -is [array] -and $arr.Count -gt 0) { $c = $arr[0] }
        elseif ($arr -is [string]) { $c = ($arr -split '[;,]')[0] }
    }
    if ($c) {
        $k = [string]$c
        if (-not $byCid.ContainsKey($k)) { $byCid[$k] = 0 }
        $byCid[$k]++
    } else { $noCid++ }
}
Write-Host "  Unique company_ids: $($byCid.Count)" -ForegroundColor Gray
$byCid.GetEnumerator() | Sort-Object { $_.Value -as [int] } -Descending | Select-Object -First 10 | ForEach-Object {
    $mark = if ($_.Key -eq $cidStr) { " <-- TARGET" } else { "" }
    Write-Host "    company_id=$($_.Key): $($_.Value) assets$mark" -ForegroundColor Gray
}
if ($noCid -gt 0) { Write-Host "    (no company_id): $noCid" -ForegroundColor Gray }

# Step 3: Filter assets by company
Write-Host "`n--- Step 3: Filter assets by company $CompanyId ---" -ForegroundColor Yellow
$companyAssets = $assets | Where-Object { Test-RowMatchesCompanyId -Row $_ -CompanyIdStr $cidStr }
Write-Host "  Company assets (Test-RowMatchesCompanyId): $($companyAssets.Count)" -ForegroundColor $(if ($companyAssets.Count -gt 0) { "Green" } else { "Red" })

# Step 4: Asset type distribution (company assets only)
Write-Host "`n--- Step 4: Asset type distribution (company assets) ---" -ForegroundColor Yellow
$typeCount = @{}
foreach ($a in $companyAssets) {
    $t = $a.asset_type; if (-not $t) { $t = $a.system_type }
    $k = if ($t) { [string]$t } else { "(null)" }
    if (-not $typeCount.ContainsKey($k)) { $typeCount[$k] = 0 }
    $typeCount[$k]++
}
$typeCount.GetEnumerator() | Sort-Object { $_.Value } -Descending | ForEach-Object {
    $mark = if ($_.Key -match 'external') { " <-- external" } else { "" }
    Write-Host "    $($_.Key): $($_.Value)$mark" -ForegroundColor Gray
}

# Step 5: Build company external asset IDs
Write-Host "`n--- Step 5: Build company external asset ID set ---" -ForegroundColor Yellow
$companyExtIds = @{}
foreach ($a in $companyAssets) {
    $t = $a.asset_type; if (-not $t) { $t = $a.system_type }
    if (-not ($t -and [string]$t -match 'external')) { continue }
    $aid = $a.id; if (-not $aid) { $aid = $a.asset_id }
    if ($aid) { $companyExtIds[[string]$aid] = $true }
}
Write-Host "  Company external asset IDs: $($companyExtIds.Count)" -ForegroundColor $(if ($companyExtIds.Count -gt 0) { "Green" } else { "Yellow" })
if ($companyExtIds.Count -gt 0) {
    $sampleIds = @($companyExtIds.Keys | Select-Object -First 5)
    Write-Host "  Sample IDs: $($sampleIds -join ', ')" -ForegroundColor Gray
}

# Step 6: Fetch external_asset_externalscan
Write-Host "`n--- Step 6: Fetch external_asset_externalscan ---" -ForegroundColor Yellow
$extData = @()
try {
    $r = Invoke-ConnectSecureRequest -Endpoint '/r/report_queries/external_asset_externalscan' -QueryParameters @{ limit = 2000; skip = 0 }
    $extData = if ($r.data) { $r.data } else { $r }
    if (-not ($extData -is [array])) { $extData = @() }
    Write-Host "  Total external scan records: $($extData.Count)" -ForegroundColor Green
} catch {
    Write-Host "  FAIL: $($_.Exception.Message)" -ForegroundColor Red
    exit 1
}

if ($extData.Count -gt 0) {
    $sampleExtIds = @($extData[0..4] | ForEach-Object { $_.id })
    Write-Host "  Sample externalscan IDs: $($sampleExtIds -join ', ')" -ForegroundColor Gray
}

# Step 7: ID overlap check
Write-Host "`n--- Step 7: ID overlap (do externalscan.id match asset.id?) ---" -ForegroundColor Yellow
$overlap = 0
$extIdSet = @{}
foreach ($e in $extData) {
    $eid = $e.id; if (-not $eid) { $eid = $e.asset_id }
    if ($eid) { $extIdSet[[string]$eid] = $true }
}
$overlap = ($extIdSet.Keys | Where-Object { $companyExtIds.ContainsKey($_) }).Count
Write-Host "  Externalscan unique IDs: $($extIdSet.Count)" -ForegroundColor Gray
Write-Host "  Overlap with company external asset IDs: $overlap" -ForegroundColor $(if ($overlap -gt 0) { "Green" } else { "Red" })

# Step 8: Filter and result
Write-Host "`n--- Step 8: Filter externalscan by company external asset IDs ---" -ForegroundColor Yellow
$filtered = @($extData | Where-Object {
    $eid = $_.id; if (-not $eid) { $eid = $_.asset_id }
    $companyExtIds.ContainsKey([string]$eid)
})
Write-Host "  Filtered count: $($filtered.Count)" -ForegroundColor $(if ($filtered.Count -gt 0) { "Green" } else { "Yellow" })

if ($filtered.Count -gt 0 -and $filtered.Count -le 5) {
    Write-Host "  Sample filtered:" -ForegroundColor Gray
    foreach ($f in $filtered) {
        Write-Host "    $($f.name): $($f.ip)" -ForegroundColor Gray
    }
}

# Alternative: match by config_name? (externalscan has config_name, assets have name)
Write-Host "`n--- Alternative: Match by config_name/name? ---" -ForegroundColor Yellow
$companyExtNames = @{}
foreach ($a in $companyAssets) {
    $t = $a.asset_type; if (-not $t) { $t = $a.system_type }
    if (-not ($t -and [string]$t -match 'external')) { continue }
    $n = $a.name; if (-not $n) { $n = $a.visible_name }
    if ($n) { $companyExtNames[[string]$n] = $true }
}
$filteredByName = @($extData | Where-Object {
    $cn = $_.config_name; if (-not $cn) { $cn = $_.name }
    $companyExtNames.ContainsKey([string]$cn)
})
Write-Host "  Company external asset names (sample): $(@($companyExtNames.Keys | Select-Object -First 3) -join ', ')" -ForegroundColor Gray
Write-Host "  Filtered by config_name match: $($filteredByName.Count)" -ForegroundColor $(if ($filteredByName.Count -gt 0) { "Green" } else { "Gray" })

Write-Host "`n=== Done ===" -ForegroundColor Cyan
