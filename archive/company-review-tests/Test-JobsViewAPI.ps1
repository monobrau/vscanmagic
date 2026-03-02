#Requires -Version 5.1
<#
.SYNOPSIS
Test /r/company/jobs_view for last external scan date - portal may use this.
#>
param([int]$CompanyId = 15373)

$scriptDir = if ($PSScriptRoot) { $PSScriptRoot } else { (Get-Location).Path }
$script:SettingsDirectory = Join-Path $env:LOCALAPPDATA "VScanMagic"
$script:ConnectSecureCredentialsPath = Join-Path $script:SettingsDirectory "ConnectSecure-Credentials.json"
. (Join-Path $scriptDir "VScanMagic-Modules\VScanMagic-Core.ps1")
. (Join-Path $scriptDir "ConnectSecure-API.ps1")

$creds = Load-ConnectSecureCredentials
Connect-ConnectSecureAPI -BaseUrl $creds.BaseUrl -TenantName $creds.TenantName -ClientId $creds.ClientId -ClientSecret $creds.ClientSecret | Out-Null

$endpoint = '/r/company/jobs_view'
Write-Host "`n=== jobs_view API Test (CompanyId: $CompanyId) ===" -ForegroundColor Cyan

function Invoke-Test {
    param([string]$Label, [hashtable]$QueryParams)
    $qStr = ($QueryParams.GetEnumerator() | ForEach-Object { "$($_.Key)=$($_.Value)" }) -join "&"
    Write-Host "`n--- $Label ---" -ForegroundColor Yellow
    Write-Host "  Query: $qStr"
    try {
        $response = Invoke-ConnectSecureRequest -Endpoint $endpoint -QueryParameters $QueryParams
        $data = $response.data; if (-not $data -and $response -is [array]) { $data = $response }
        if (-not $data -or -not ($data -is [array])) { Write-Host "  Result: no data" -ForegroundColor Gray; return }
        $count = $data.Count
        Write-Host "  Records: $count" -ForegroundColor $(if ($count -gt 0) { "Green" } else { "Gray" })
        if ($count -gt 0) {
            $types = $data | ForEach-Object { $_.type } | Where-Object { $_ } | Select-Object -Unique
            Write-Host "  Types: $($types -join ', ')" -ForegroundColor Gray
            $withCompany = ($data | Where-Object { $_.company_id -eq $CompanyId }).Count
            Write-Host "  Matching company $CompanyId : $withCompany" -ForegroundColor Gray
            $latest = $data | Sort-Object -Property { [DateTime]::Parse($_.updated) } -Descending | Select-Object -First 1
            if ($latest) { Write-Host "  Latest updated: $($latest.updated) type=$($latest.type) company_id=$($latest.company_id)" -ForegroundColor Gray }
        }
    } catch { Write-Host "  FAIL: $($_.Exception.Message)" -ForegroundColor Red }
}

Invoke-Test -Label "1. Baseline (no filter)" -QueryParams @{ limit = 50; skip = 0 }
Invoke-Test -Label "2. condition=company_id=$CompanyId" -QueryParams @{ condition = "company_id=$CompanyId"; limit = 50; skip = 0 }
Invoke-Test -Label "3. condition with type (externalscan)" -QueryParams @{ condition = "company_id=$CompanyId and type='externalscan'"; limit = 50; skip = 0; order_by = "updated desc" }
Invoke-Test -Label "4. condition type 'External Scan'" -QueryParams @{ condition = "company_id=$CompanyId and type='External Scan'"; limit = 50; skip = 0; order_by = "updated desc" }

# List all unique types from company-filtered jobs
Write-Host "`n--- All job types for company $CompanyId ---" -ForegroundColor Yellow
try {
    $r = Invoke-ConnectSecureRequest -Endpoint $endpoint -QueryParameters @{ condition = "company_id=$CompanyId"; limit = 200; skip = 0 }
    $d = $r.data; if (-not $d) { $d = $r }
    $allTypes = $d | ForEach-Object { $_.type } | Where-Object { $_ } | Sort-Object -Unique
    Write-Host "  Types: $($allTypes -join ' | ')" -ForegroundColor Gray
    $extTypes = $allTypes | Where-Object { $_ -match 'external|External|scan|Scan' }
    if ($extTypes) { Write-Host "  External-related: $($extTypes -join ', ')" -ForegroundColor Green }
} catch { Write-Host "  FAIL: $_" -ForegroundColor Red }

Write-Host "`nDone. jobs_view + condition=company_id=X works (server-side). Use max(updated) for last scan." -ForegroundColor Cyan
