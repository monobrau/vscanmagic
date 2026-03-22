#Requires -Version 5.1
<#
.SYNOPSIS
    Lightweight tests for pure PowerShell helpers (no Excel/COM, no network for NVD).
    Run from repo root: .\tests\Run-Tests.ps1
#>
$ErrorActionPreference = 'Stop'
$root = Split-Path -Parent $PSScriptRoot

function Assert-Equal {
    param([string]$Name, $Expected, $Actual)
    if ($Expected -ne $Actual) {
        throw "FAIL $Name : expected '$Expected' got '$Actual'"
    }
}

$nvdPath = Join-Path $root 'VScanMagic-Modules\VScanMagic-NVD.ps1'
. $nvdPath

$bad = Get-NVDRemediationForCve -CveId 'not-a-valid-cve'
Assert-Equal -Name 'InvalidCveEmpty' -Expected '' -Actual $bad

$noCves = Get-NVDRemediationForCveList -CveListText 'no cve ids in this string' -ApiKey ''
Assert-Equal -Name 'EmptyCveListText' -Expected '' -Actual $noCves

Write-Host 'All tests passed.' -ForegroundColor Green
exit 0
