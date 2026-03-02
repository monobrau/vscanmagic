#Requires -Version 5.1
# Quick test: which Company Review endpoints support condition=company_id=X?
param([int]$CompanyId = 15373)

$scriptDir = if ($PSScriptRoot) { $PSScriptRoot } else { (Get-Location).Path }
. (Join-Path $scriptDir "VScanMagic-Modules\VScanMagic-Core.ps1")
. (Join-Path $scriptDir "ConnectSecure-API.ps1")

$creds = Load-ConnectSecureCredentials
Connect-ConnectSecureAPI -BaseUrl $creds.BaseUrl -TenantName $creds.TenantName -ClientId $creds.ClientId -ClientSecret $creds.ClientSecret | Out-Null

$endpoints = @(
    @{ Name = 'agents'; Path = '/r/company/agents' }
    @{ Name = 'credentials'; Path = '/r/company/credentials' }
    @{ Name = 'agent_credentials_mapping'; Path = '/r/company/agent_credentials_mapping' }
    @{ Name = 'agent_discoverysettings_mapping'; Path = '/r/company/agent_discoverysettings_mapping' }
    @{ Name = 'asset_firewall_policy'; Path = '/r/asset/asset_firewall_policy' }
    @{ Name = 'company_stats'; Path = '/r/company/company_stats' }
)

Write-Host "`n=== Server-side condition test (company_id=$CompanyId) ===" -ForegroundColor Cyan
foreach ($ep in $endpoints) {
    try {
        $base = Invoke-ConnectSecureRequest -Endpoint $ep.Path -QueryParameters @{ limit = 100; skip = 0 }
        $baseData = $base.data; if (-not $baseData) { $baseData = $base }
        $baseCount = if ($baseData -is [array]) { $baseData.Count } else { 0 }
        $cond = Invoke-ConnectSecureRequest -Endpoint $ep.Path -QueryParameters @{ condition = "company_id=$CompanyId"; limit = 500; skip = 0 }
        $condData = $cond.data; if (-not $condData) { $condData = $cond }
        $condCount = if ($condData -is [array]) { $condData.Count } else { (if ($condData) { 1 } else { 0 }) }
        $works = $condCount -lt $baseCount -or ($condCount -gt 0 -and $condCount -le $baseCount)
        $color = if ($works -and $condCount -gt 0) { "Green" } elseif ($condCount -eq 0 -and $baseCount -gt 0) { "Red" } else { "Yellow" }
        Write-Host "  $($ep.Name): base=$baseCount cond=$condCount -> $(if ($works) { 'OK' } else { 'NO' })" -ForegroundColor $color
    } catch { Write-Host "  $($ep.Name): FAIL $($_.Exception.Message)" -ForegroundColor Red }
}
Write-Host ""
