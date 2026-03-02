#Requires -Version 5.1
<#
.SYNOPSIS
Test external assets (discovery_settings externalscan) for multiple companies.
Helps identify companies where external assets work vs don't.
#>
param(
    [int[]]$CompanyIds = @(15373),
    [string]$CompanyName = '',
    [switch]$ListCompanies
)

$scriptDir = if ($PSScriptRoot) { $PSScriptRoot } else { (Get-Location).Path }
. (Join-Path $scriptDir "VScanMagic-Modules\VScanMagic-Core.ps1")
. (Join-Path $scriptDir "ConnectSecure-API.ps1")

$creds = Load-ConnectSecureCredentials
Connect-ConnectSecureAPI -BaseUrl $creds.BaseUrl -TenantName $creds.TenantName -ClientId $creds.ClientId -ClientSecret $creds.ClientSecret | Out-Null

if ($ListCompanies) {
    $companies = Get-ConnectSecureCompanies
    Write-Host "`nCompanies (search for Example Company or similar):" -ForegroundColor Cyan
    $example = $companies | Where-Object { $_.name -match 'Example' }
    if ($example) {
        Write-Host "  Example Company: $($example.id)" -ForegroundColor Green
    }
    $companies | ForEach-Object { Write-Host "  $($_.id): $($_.name)" } | Out-Null
    $companies | ForEach-Object { Write-Host "  $($_.id): $($_.name)" }
    return
}

# Search by company name (e.g. -CompanyName "Example Company")
if ($CompanyName) {
    $companies = Get-ConnectSecureCompanies
    $match = $companies | Where-Object { $_.name -match [regex]::Escape($CompanyName) } | Select-Object -First 1
    if ($match) {
        Write-Host "Found: $($match.name) (ID=$($match.id))" -ForegroundColor Green
        $CompanyIds = @($match.id)
    } else {
        Write-Host "No company matching '$CompanyName'. Use -ListCompanies to see all." -ForegroundColor Yellow
        return
    }
}

Write-Host "`n=== External Assets by Company ===" -ForegroundColor Cyan
foreach ($cid in $CompanyIds) {
    $ext = Get-ConnectSecureExternalScanDiscoverySettings -CompanyId $cid
    $all = Get-ConnectSecureDiscoverySettings -CompanyId $cid
    $extFromAll = @($all | Where-Object {
        $t = $_.discovery_settings_type; if (-not $t) { $t = $_.type }
        $t -and [string]$t -match 'external|externalscan'
    })
    Write-Host "`nCompany $cid : externalscan=$($ext.Count) fallback=$($extFromAll.Count) all_ds=$($all.Count)" -ForegroundColor $(if ($ext.Count -gt 0 -or $extFromAll.Count -gt 0) { "Green" } else { "Yellow" })
    if ($ext.Count -eq 0 -and $extFromAll.Count -gt 0) {
        $types = $extFromAll | ForEach-Object { $t = $_.discovery_settings_type; if (-not $t) { $t = $_.type }; $t } | Where-Object { $_ } | Select-Object -Unique
        Write-Host "  Fallback types: $($types -join ', ')" -ForegroundColor Gray
    }
}
Write-Host ""
