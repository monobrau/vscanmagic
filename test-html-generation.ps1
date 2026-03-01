# Verification script for HTML report generation
# Run from repo root: powershell -ExecutionPolicy Bypass -File test-html-generation.ps1

$ErrorActionPreference = "Stop"
$script:ScriptDirectory = $PSScriptRoot
$modulesDir = Join-Path $script:ScriptDirectory "VScanMagic-Modules"

. (Join-Path $modulesDir "VScanMagic-Core.ps1")
. (Join-Path $modulesDir "VScanMagic-Data.ps1")
. (Join-Path $modulesDir "VScanMagic-Reports.ps1")
. (Join-Path $modulesDir "VScanMagic-Dialogs.ps1")

# Minimal Top10Data for testing
$top10 = @(
    [PSCustomObject]@{
        Product = "Windows 11 (all versions)"
        RiskScore = 8.5
        EPSSScore = 0.1234
        AvgCVSS = 7.2
        VulnCount = 5
        AffectedSystems = @(
            [PSCustomObject]@{ HostName = "WORKSTATION01"; IP = "192.168.1.10"; Username = "user1" }
        )
    }
)

$timeEstimates = @(
    [PSCustomObject]@{
        Product = "Windows 11 (all versions)"
        TimeEstimate = 2.0
        AfterHours = $false
        ThirdParty = $false
        TicketGenerated = $false
    }
)

$outputPath = Join-Path $env:TEMP "VScanMagic_Verify_$(Get-Date -Format 'yyyyMMdd_HHmmss').html"

Write-Host "Testing New-CombinedReportHtml..."
Write-Host "  OutputPath: $outputPath"
Write-Host "  CompanyName: Test Company Inc."
Write-Host ""

New-CombinedReportHtml `
    -OutputPath $outputPath `
    -TopTenData $top10 `
    -TimeEstimates $timeEstimates `
    -IsRMITPlus $false `
    -IncludeTicketInstructions $true `
    -IncludeEmailTemplate $true `
    -IncludeTimeEstimate $true `
    -FilterTopN "10" `
    -CompanyName "Test Company Inc."

Write-Host "Verification:"
if (-not (Test-Path $outputPath)) {
    Write-Host "  FAIL: File was not created" -ForegroundColor Red
    exit 1
}
Write-Host "  OK: File created" -ForegroundColor Green

$content = Get-Content -Path $outputPath -Raw -Encoding UTF8
$checks = @(
    @{ Name = "DOCTYPE"; Pattern = "<!DOCTYPE html>"; Pass = $content -match "<!DOCTYPE html>" }
    @{ Name = "Company name in header"; Pattern = "Test Company Inc."; Pass = $content -match "Test Company Inc\." }
    @{ Name = "report-header class"; Pattern = "report-header"; Pass = $content -match 'class="report-header"' }
    @{ Name = "Tab bar"; Pattern = "tab-bar"; Pass = $content -match 'class="tab-bar"' }
    @{ Name = "Ticket Instructions tab"; Pattern = "Ticket Instructions"; Pass = $content -match "Ticket Instructions" }
    @{ Name = "Email Template tab"; Pattern = "Email Template"; Pass = $content -match "Email Template" }
    @{ Name = "Time Estimate tab"; Pattern = "Time Estimate"; Pass = $content -match "Time Estimate" }
    @{ Name = "Ticket Notes tab"; Pattern = "Ticket Notes"; Pass = $content -match "Ticket Notes" }
    @{ Name = "Closing body"; Pattern = "</body>"; Pass = $content -match "</body>" }
)

$allPass = $true
foreach ($c in $checks) {
    $status = if ($c.Pass) { "OK" } else { "FAIL" }
    $color = if ($c.Pass) { "Green" } else { "Red"; $allPass = $false }
    Write-Host "  $status : $($c.Name)" -ForegroundColor $color
}

Write-Host ""
if ($allPass) {
    Write-Host "All checks passed. HTML created successfully." -ForegroundColor Green
    Write-Host "  Open: $outputPath"
} else {
    Write-Host "Some checks failed." -ForegroundColor Red
    exit 1
}
