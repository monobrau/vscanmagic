#Requires -Modules Microsoft.PowerShell.Utility
<#
.SYNOPSIS
VScanMagic v4 - GUI-based Vulnerability Report Generator
Processes vulnerability scan Excel files and generates professional DOCX reports.

.DESCRIPTION
This script provides a GUI interface for:
- Processing vulnerability scan Excel files
- Calculating composite risk scores
- Generating professional Word documents with executive summaries
- Creating color-coded risk tables
- Providing actionable remediation guidance

.NOTES
Version: 4.0.4
Requires: Microsoft Excel and Microsoft Word installed.
Author: River Run MSP
Modular: Core, Data, Reports, Dialogs, Form loaded from VScanMagic-Modules/
#>

# --- Add Required Assemblies ---
Add-Type -AssemblyName System.Windows.Forms
Add-Type -AssemblyName System.Drawing

# --- Script Directory (for module paths) ---
if ([string]::IsNullOrEmpty($PSScriptRoot)) {
    try {
        $exePath = [System.Diagnostics.Process]::GetCurrentProcess().MainModule.FileName
        $script:ScriptDirectory = [System.IO.Path]::GetDirectoryName($exePath)
    } catch {
        $script:ScriptDirectory = (Get-Location).Path
    }
} else {
    $script:ScriptDirectory = $PSScriptRoot
}

$modulesDir = Join-Path $script:ScriptDirectory "VScanMagic-Modules"
if (-not (Test-Path $modulesDir)) {
    Write-Error "VScanMagic-Modules folder not found at: $modulesDir"
    exit 1
}

# --- Load Modules (order matters: Core -> Data -> Reports -> Dialogs -> Form) ---
$corePath = Join-Path $modulesDir "VScanMagic-Core.ps1"
$dataPath = Join-Path $modulesDir "VScanMagic-Data.ps1"
$reportsPath = Join-Path $modulesDir "VScanMagic-Reports.ps1"
$dialogsPath = Join-Path $modulesDir "VScanMagic-Dialogs.ps1"
$formPath = Join-Path $modulesDir "VScanMagic-Form.ps1"

. $corePath
. $dataPath
. $reportsPath
. $dialogsPath
. $formPath

# --- Main Entry ---
if (-not $script:IsApiMode) {
    Show-VScanMagicGUI
}
