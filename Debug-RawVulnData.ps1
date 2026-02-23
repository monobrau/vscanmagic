#Requires -Version 5.1
<#
.SYNOPSIS
Fetches raw vulnerability data and dumps structure to help diagnose type issues.
Run: .\Debug-RawVulnData.ps1
#>

$credsPath = Join-Path $env:LOCALAPPDATA "VScanMagic\ConnectSecure-Credentials.json"
if (-not (Test-Path $credsPath)) {
    Write-Host "No credentials at $credsPath - run GUI or configure first." -ForegroundColor Red
    exit 1
}
$creds = Get-Content $credsPath -Raw | ConvertFrom-Json

. .\ConnectSecure-API.ps1
Connect-ConnectSecureAPI -BaseUrl $creds.BaseUrl -TenantName $creds.TenantName -ClientId $creds.ClientId -ClientSecret $creds.ClientSecret | Out-Null

$limit = 2
Write-Host "Fetching $limit raw vulnerabilities (application)..." -ForegroundColor Cyan

# Test 1: Call Invoke-ConnectSecureVulnerabilityQuery directly with literals
Write-Host "`nTest 1: Direct call with literals..." -ForegroundColor Yellow
try {
    $direct = Invoke-ConnectSecureVulnerabilityQuery -VulnType 'application' -CompanyId 0 -Limit 2 -Skip 0 -Filter 'severity.keyword:(Critical OR High)' -Sort 'severity.keyword:desc' -FetchAll:$false
    Write-Host "Direct call SUCCEEDED, got $($direct.Count) records" -ForegroundColor Green
} catch {
    Write-Host "Direct call FAILED: $($_.Exception.Message)" -ForegroundColor Red
}

# Test 2: Call Get-ConnectSecureVulnerabilities with variable
Write-Host "`nTest 2: Via Get-ConnectSecureVulnerabilities -Limit `$limit -FetchAll:`$false -Raw..." -ForegroundColor Yellow
$raw = $null
try {
    $raw = Get-ConnectSecureVulnerabilities -Limit $limit -FetchAll:$false -Raw
    Write-Host "Get-ConnectSecureVulnerabilities (variable) SUCCEEDED, got $($raw.Count) records" -ForegroundColor Green
} catch {
    Write-Host "Get-ConnectSecureVulnerabilities (variable) FAILED: $($_.Exception.Message)" -ForegroundColor Red
}

# Test 2b: If Test 2 failed, try with literal -Limit 2
if (-not $raw -and $direct) {
    Write-Host "`nTest 2b: Retry with literal -Limit 2..." -ForegroundColor Yellow
    try {
        $raw = Get-ConnectSecureVulnerabilities -Limit 2 -FetchAll:$false -Raw
        Write-Host "Get-ConnectSecureVulnerabilities (literal) SUCCEEDED, got $($raw.Count) records" -ForegroundColor Green
    } catch {
        Write-Host "Get-ConnectSecureVulnerabilities (literal) FAILED: $($_.Exception.Message)" -ForegroundColor Red
    }
}
$dataToUse = if ($raw -and $raw.Count -gt 0) { $raw } elseif ($direct -and $direct.Count -gt 0) { $direct } else { $null }
if (-not $dataToUse -or $dataToUse.Count -eq 0) {
    Write-Host "No data from either test. Check credentials and API." -ForegroundColor Red
    exit 1
}

Write-Host "`n=== RECORD COUNT: $($dataToUse.Count) ===" -ForegroundColor Yellow
Write-Host ""

for ($i = 0; $i -lt [Math]::Min(2, $dataToUse.Count); $i++) {
    $item = $dataToUse[$i]
    Write-Host "--- RECORD $($i+1) ---" -ForegroundColor Green
    Write-Host "  Item type: $($item.GetType().FullName)"
    Write-Host "  Has _source: $($null -ne $item.PSObject.Properties['_source'])"
    
    $obj = if ($item.PSObject.Properties['_source']) { $item._source } else { $item }
    Write-Host "  Obj type: $($obj.GetType().FullName)"
    Write-Host "  Properties:"
    
    foreach ($prop in $obj.PSObject.Properties) {
        $v = $prop.Value
        $vType = if ($null -eq $v) { 'null' } else { $v.GetType().FullName }
        $isArray = $v -is [Array] -or ($v -is [System.Collections.IEnumerable] -and $v -isnot [string])
        $preview = if ($null -eq $v) { 'null' } elseif ($isArray) { "[Array,len=$(@($v).Count)]" } else { $v.ToString().Substring(0, [Math]::Min(50, $v.ToString().Length)) }
        Write-Host "    $($prop.Name): $vType | IsArray=$isArray | Preview: $preview"
    }
    Write-Host ""
}

Write-Host "=== RAW JSON (first record, depth 4) ===" -ForegroundColor Yellow
$first = $dataToUse[0]
$obj = if ($first.PSObject.Properties['_source']) { $first._source } else { $first }
$obj | ConvertTo-Json -Depth 4

Write-Host "`n=== TEST EXPORT (will show exact failure point) ===" -ForegroundColor Yellow
$testPath = Join-Path $PWD "debug_export_test.xlsx"
try {
    Export-ConnectSecureDataToExcel -Data $dataToUse -OutputPath $testPath -SheetName 'Test' -OnProgress $null
    Write-Host "Export SUCCEEDED -> $testPath" -ForegroundColor Green
} catch {
    Write-Host "Export FAILED: $($_.Exception.Message)" -ForegroundColor Red
    Write-Host "Stack: $($_.ScriptStackTrace)" -ForegroundColor Gray
}
