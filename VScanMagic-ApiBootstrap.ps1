# VScanMagic-ApiBootstrap.ps1
# Dot-sourced only by VScanMagic-API.ps1. Loads the minimal module chain for REST report generation:
# Core -> Data -> Reports (no Dialogs, Form, Memberberry, VScanMagic-GUI.ps1).

#Requires -Version 5.1

$apiRoot = $PSScriptRoot
if ([string]::IsNullOrEmpty($apiRoot)) {
    $apiRoot = Split-Path -Parent $MyInvocation.MyCommand.Path
}

$script:ScriptDirectory = $apiRoot
if (-not $script:IsApiMode) {
    $script:IsApiMode = $true
}

# Reports use Application.DoEvents during long COM work; same as VScanMagic-GUI.ps1
Add-Type -AssemblyName System.Windows.Forms -ErrorAction SilentlyContinue
Add-Type -AssemblyName System.Drawing -ErrorAction SilentlyContinue

$modulesDir = Join-Path $apiRoot 'VScanMagic-Modules'
if (-not (Test-Path $modulesDir)) {
    throw "VScanMagic-Modules not found at: $modulesDir"
}

. (Join-Path $modulesDir 'VScanMagic-Core.ps1')
. (Join-Path $modulesDir 'VScanMagic-Data.ps1')
. (Join-Path $modulesDir 'VScanMagic-Reports.ps1')
