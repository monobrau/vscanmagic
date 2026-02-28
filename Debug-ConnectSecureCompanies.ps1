#Requires -Version 5.1
<#
.SYNOPSIS
    Debug script - fetches companies from ConnectSecure API and outputs raw response structure.
    Run this to see the actual property names the API returns.
#>

param(
    [Parameter(Mandatory=$true)]
    [string]$BaseUrl,
    [Parameter(Mandatory=$true)]
    [string]$TenantName,
    [Parameter(Mandatory=$true)]
    [string]$ClientId,
    [Parameter(Mandatory=$true)]
    [string]$ClientSecret
)

$scriptPath = Split-Path -Parent $MyInvocation.MyCommand.Path
$apiPath = Join-Path $scriptPath "ConnectSecure-API.ps1"

if (-not (Test-Path $apiPath)) {
    Write-Error "ConnectSecure-API.ps1 not found"
    exit 1
}

. $apiPath | Out-Null

Write-Host "Authenticating..." -ForegroundColor Cyan
$connected = Connect-ConnectSecureAPI -BaseUrl $BaseUrl -TenantName $TenantName -ClientId $ClientId -ClientSecret $ClientSecret
if (-not $connected) {
    Write-Error "Authentication failed"
    exit 1
}

Write-Host "Fetching companies..." -ForegroundColor Cyan
$response = Invoke-ConnectSecureRequest -Endpoint "/r/company/companies" -QueryParameters @{ limit = 10; skip = 0 }

Write-Host "`n=== RAW RESPONSE (full) ===" -ForegroundColor Yellow
$response | ConvertTo-Json -Depth 10

Write-Host "`n=== FIRST COMPANY OBJECT ===" -ForegroundColor Yellow
$companies = @()
if ($response -is [array]) { $companies = $response }
elseif ($response.data) { $companies = $response.data }
elseif ($response.hits -and $response.hits.hits) {
    $companies = $response.hits.hits | ForEach-Object { if ($_._source) { $_ } else { $_ } }
}

if ($companies.Count -gt 0) {
    Write-Host "Companies count: $($companies.Count)"
    $first = $companies[0]
    if ($first -is [bool]) {
        Write-Host "`nNOTE: API returned boolean array - no company objects. Trying company_stats..." -ForegroundColor Magenta
        try {
            $statsResp = Invoke-ConnectSecureRequest -Endpoint "/r/company/company_stats" -QueryParameters @{ limit = 100; skip = 0 }
            Write-Host "`n=== COMPANY_STATS RESPONSE ===" -ForegroundColor Yellow
            $statsResp | ConvertTo-Json -Depth 5
            $stats = if ($statsResp.data) { $statsResp.data } elseif ($statsResp -is [array]) { $statsResp } else { @() }
            if ($stats.Count -gt 0) {
                Write-Host "`nFirst stats object properties:" -ForegroundColor Cyan
                $stats[0] | Get-Member -MemberType NoteProperty | ForEach-Object { Write-Host "  - $($_.Name): $($stats[0].($_.Name))" }
            }
        } catch {
            Write-Host "company_stats failed: $($_.Exception.Message)" -ForegroundColor Red
        }
    } else {
        Write-Host "`nFirst company - Properties:"
        $first | Get-Member -MemberType NoteProperty | ForEach-Object { Write-Host "  - $($_.Name): $($companies[0].($_.Name))" }
        Write-Host "`nFirst company - Full JSON:"
        $first | ConvertTo-Json -Depth 5
    }
} else {
    Write-Host "No companies in response"
}
