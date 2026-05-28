#Requires -Version 5.1
<#
.SYNOPSIS
    Runs PSScriptAnalyzer on project PowerShell sources (excludes vendor ps2exe).
    Install: Install-Module PSScriptAnalyzer -Scope CurrentUser -Force
#>
param(
    [ValidateSet('Error', 'Warning', 'Information')]
    [string[]]
    $Severity = @('Error', 'Warning')
)

$ErrorActionPreference = 'Stop'
$root = Split-Path -Parent $PSScriptRoot

if (-not (Get-Module -ListAvailable -Name PSScriptAnalyzer)) {
    Write-Warning 'PSScriptAnalyzer is not installed. Install with: Install-Module PSScriptAnalyzer -Scope CurrentUser -Force'
    exit 0
}

Import-Module PSScriptAnalyzer -Force

$paths = @(
    (Join-Path $root 'VScanMagic-GUI.ps1')
    (Join-Path $root 'VScanMagic-API.ps1')
    (Join-Path $root 'ConnectSecure-API.ps1')
    (Join-Path $root 'Export-VulnerabilitiesToCsv.ps1')
    (Join-Path $root 'VScanMagic-Modules')
)

$issues = @()
foreach ($p in $paths) {
    if (-not (Test-Path $p)) { continue }
    $issues += Invoke-ScriptAnalyzer -Path $p -Recurse -Severity $Severity -ErrorAction SilentlyContinue
}

if ($issues.Count -eq 0) {
    Write-Host 'ScriptAnalyzer: no Error/Warning issues.' -ForegroundColor Green
    exit 0
}

$issues | Format-Table -AutoSize
exit 1
