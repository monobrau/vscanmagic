<#
.SYNOPSIS
    Extracts unique software/product names from a ConnectSecure vulnerability Excel file.
.PARAMETER ExcelPath
    Path to the XLSX file (e.g. Global Critical Internal Vulnerabilities report).
.PARAMETER OutputPath
    Optional. Save unique names to CSV. If omitted, outputs to console.
.PARAMETER ExpandCommaLists
    When set, splits comma-separated values into individual software names (ConnectSecure often lists multiple per cell).
.EXAMPLE
    .\Get-UniqueSoftwareNames.ps1 -ExcelPath "K:\...\Global - Critical Internal Vulnerabilities.xlsx"
    .\Get-UniqueSoftwareNames.ps1 -ExcelPath "K:\...\file.xlsx" -OutputPath ".\unique-software.csv" -ExpandCommaLists
#>
param(
    [Parameter(Mandatory=$true)]
    [string]$ExcelPath,
    [string]$OutputPath = "",
    [switch]$ExpandCommaLists = $false
)

$ErrorActionPreference = "Stop"

if (-not (Test-Path -LiteralPath $ExcelPath)) {
    Write-Error "File not found: $ExcelPath"
    exit 1
}

# Copy to temp if OneDrive path (avoids sync locks)
$pathToOpen = $ExcelPath
$tempPath = $null
if ($ExcelPath -match 'OneDrive|iCloud|Dropbox|Google Drive|Box\.com') {
    $tempPath = Join-Path $env:TEMP ("UniqueSoftware_" + [Guid]::NewGuid().ToString("N") + "_" + [System.IO.Path]::GetFileName($ExcelPath))
    Copy-Item -LiteralPath $ExcelPath -Destination $tempPath -Force
    $pathToOpen = $tempPath
}

$excel = $null
$workbook = $null

try {
    $excel = New-Object -ComObject Excel.Application
    $excel.Visible = $false
    $workbook = $excel.Workbooks.Open($pathToOpen, 0, $true)

    $productNames = [System.Collections.Generic.HashSet[string]]::new([StringComparer]::OrdinalIgnoreCase)

    foreach ($sheet in $workbook.Worksheets) {
        $usedRange = $sheet.UsedRange
        $rowCount = $usedRange.Rows.Count
        $colCount = $usedRange.Columns.Count
        if ($rowCount -le 1 -or $colCount -lt 2) { continue }

        $rangeValues = $usedRange.Value2
        if ($null -eq $rangeValues) { continue }

        # Find product/software column - try common header names
        $productColIndex = $null
        for ($c = 1; $c -le $colCount; $c++) {
            $h = [string]$rangeValues[1, $c]
            if ($h -match '^(Software Name|Product Name|Application Name|Product|Software|App Name)$' -or
                $h -match '^(Software|Product)\s*$') {
                $productColIndex = $c
                break
            }
        }
        if (-not $productColIndex) { continue }

        for ($row = 2; $row -le $rowCount; $row++) {
            $val = [string]$rangeValues[$row, $productColIndex]
            $val = $val -replace "^\[|'|\]$", "" -replace "^'|'$", "" -replace "^\s+|\s+$", ""
            if (-not [string]::IsNullOrWhiteSpace($val)) {
                if ($ExpandCommaLists) {
                    $parts = $val -split ',' | ForEach-Object { $_.Trim() } | Where-Object { $_ }
                    foreach ($p in $parts) { $null = $productNames.Add($p) }
                } else {
                    $null = $productNames.Add($val)
                }
            }
        }
    }

    $unique = $productNames | Sort-Object

    if ($OutputPath) {
        $unique | ForEach-Object { [PSCustomObject]@{ SoftwareName = $_ } } | Export-Csv -Path $OutputPath -NoTypeInformation -Encoding UTF8
        Write-Host "Exported $($unique.Count) unique software names to: $OutputPath" -ForegroundColor Green
    } else {
        $unique | ForEach-Object { Write-Output $_ }
        Write-Host "`n$($unique.Count) unique software names" -ForegroundColor Cyan
    }
} finally {
    if ($workbook) { $workbook.Close($false); [System.Runtime.InteropServices.Marshal]::ReleaseComObject($workbook) | Out-Null }
    if ($excel) { $excel.Quit(); [System.Runtime.InteropServices.Marshal]::ReleaseComObject($excel) | Out-Null }
    if ($tempPath -and (Test-Path $tempPath)) { Remove-Item $tempPath -Force -ErrorAction SilentlyContinue }
    [System.GC]::Collect()
}
