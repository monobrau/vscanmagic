#Requires -Modules Microsoft.PowerShell.Utility
<#
.SYNOPSIS
VScanMagic v3 - GUI-based Vulnerability Report Generator
Processes vulnerability scan Excel files and generates professional DOCX reports.

.DESCRIPTION
This script provides a GUI interface for:
- Processing vulnerability scan Excel files
- Calculating composite risk scores
- Generating professional Word documents with executive summaries
- Creating color-coded risk tables
- Providing actionable remediation guidance

.NOTES
Version: 3.0.0
Requires: Microsoft Excel and Microsoft Word installed.
Author: River Run MSP
#>

# --- Add Required Assemblies ---
Add-Type -AssemblyName System.Windows.Forms
Add-Type -AssemblyName System.Drawing

# --- Configuration ---
$script:Config = @{
    AppName = "VScanMagic v3"
    Version = "3.0.0"
    Author = "River Run MSP"

    # Risk Score Calculation
    CVSSEquivalent = @{
        Critical = 9.0
        High = 7.0
        Medium = 5.0
        Low = 3.0
    }

    # Heatmap Color Thresholds for Risk Scores (Low to High gradient)
    RiskColors = @{
        Critical = @{ Threshold = 7500; Color = 'DC143C'; Name = 'Critical'; TextColor = 'FFFFFF' }  # Crimson Red
        VeryHigh = @{ Threshold = 3000; Color = 'FF4500'; Name = 'Very High'; TextColor = 'FFFFFF' }  # Orange Red
        High = @{ Threshold = 1000; Color = 'FF8C00'; Name = 'High'; TextColor = 'FFFFFF' }  # Dark Orange
        MediumHigh = @{ Threshold = 500; Color = 'FFA500'; Name = 'Medium-High'; TextColor = '000000' }  # Orange
        Medium = @{ Threshold = 100; Color = 'FFFF00'; Name = 'Medium'; TextColor = '000000' }  # Yellow
        Low = @{ Threshold = 10; Color = 'ADFF2F'; Name = 'Low'; TextColor = '000000' }  # Green Yellow
        VeryLow = @{ Threshold = 0; Color = '90EE90'; Name = 'Very Low'; TextColor = '000000' }  # Light Green
    }

    # Products to Filter Out
    FilteredProducts = @(
        'Google Chrome',
        'Mozilla Firefox',
        'OS-OUT-OF-SUPPORT'
    )

    # Windows Consolidation Rules
    WindowsConsolidation = @{
        'Windows Server 2012 (all versions)' = @('Windows Server 2012', 'Windows Server 2012 R2')
        'Windows 11 (all versions)' = @('Windows 11', 'Windows 1122H2', 'Windows 1123H2', 'Windows 1124H2')
        'Windows 10 (all versions)' = @('Windows 10', 'Windows 1022H2')
    }

    # Source Sheet Configuration
    SourceSheetPatterns = @("Remediate within *", "Remediate at *")
    ExcludeSheetPatterns = @("Company", "Linux Remediations")
    ConsolidatedSheetName = "Source Data"
    PivotSheetName = "Proposed Remediations (all)"
    SheetToExcludeFormatting = "Company"

    # Excel Formatting Configuration
    ConditionalFormatThreshold = 0.075
    ExcelPathLimit = 200
}

# --- Helper Functions ---

function Write-Log {
    param(
        [string]$Message,
        [ValidateSet('Info', 'Warning', 'Error', 'Success')]
        [string]$Level = 'Info'
    )

    $timestamp = Get-Date -Format "HH:mm:ss"
    $logMessage = "[$timestamp] $Message"

    if ($script:LogTextBox) {
        $script:LogTextBox.AppendText("$logMessage`r`n")
        $script:LogTextBox.ScrollToCaret()
    }

    switch ($Level) {
        'Warning' { Write-Warning $Message }
        'Error' { Write-Error $Message }
        default { Write-Host $Message }
    }
}

function Clear-ComObject {
    param([object]$ComObject)

    if ($null -ne $ComObject -and [System.Runtime.InteropServices.Marshal]::IsComObject($ComObject)) {
        try {
            [System.Runtime.InteropServices.Marshal]::ReleaseComObject($ComObject) | Out-Null
        } catch {
            Write-Log "Error releasing COM object: $($_.Exception.Message)" -Level Warning
        }
    }
}

function Test-FileLocked {
    param([string]$FilePath)

    if (-not (Test-Path $FilePath)) {
        return $false
    }

    try {
        $fileStream = [System.IO.File]::Open($FilePath, [System.IO.FileMode]::Open, [System.IO.FileAccess]::ReadWrite, [System.IO.FileShare]::None)
        $fileStream.Close()
        $fileStream.Dispose()
        return $false
    } catch {
        return $true
    }
}

function Get-RiskScoreColor {
    param([double]$RiskScore)

    # Check heatmap levels from highest to lowest
    if ($RiskScore -ge $script:Config.RiskColors.Critical.Threshold) {
        return $script:Config.RiskColors.Critical
    } elseif ($RiskScore -ge $script:Config.RiskColors.VeryHigh.Threshold) {
        return $script:Config.RiskColors.VeryHigh
    } elseif ($RiskScore -ge $script:Config.RiskColors.High.Threshold) {
        return $script:Config.RiskColors.High
    } elseif ($RiskScore -ge $script:Config.RiskColors.MediumHigh.Threshold) {
        return $script:Config.RiskColors.MediumHigh
    } elseif ($RiskScore -ge $script:Config.RiskColors.Medium.Threshold) {
        return $script:Config.RiskColors.Medium
    } elseif ($RiskScore -ge $script:Config.RiskColors.Low.Threshold) {
        return $script:Config.RiskColors.Low
    } else {
        return $script:Config.RiskColors.VeryLow
    }
}

function ConvertTo-HexColor {
    param([string]$HexColor)

    # Convert hex string to Word color integer (BGR format)
    $r = [Convert]::ToInt32($HexColor.Substring(0, 2), 16)
    $g = [Convert]::ToInt32($HexColor.Substring(2, 2), 16)
    $b = [Convert]::ToInt32($HexColor.Substring(4, 2), 16)

    return $b * 65536 + $g * 256 + $r
}

function Find-ColumnIndex {
    param(
        [hashtable]$Headers,
        [string[]]$PossibleNames
    )

    # Try exact match first (case-insensitive)
    foreach ($name in $PossibleNames) {
        foreach ($header in $Headers.Keys) {
            if ($header -eq $name) {
                return $Headers[$header]
            }
        }
    }

    # Try case-insensitive match
    foreach ($name in $PossibleNames) {
        foreach ($header in $Headers.Keys) {
            if ($header.ToLower() -eq $name.ToLower()) {
                return $Headers[$header]
            }
        }
    }

    # Try partial match
    foreach ($name in $PossibleNames) {
        foreach ($header in $Headers.Keys) {
            if ($header -like "*$name*" -or $name -like "*$header*") {
                return $Headers[$header]
            }
        }
    }

    return $null
}

function Get-SafeNumericValue {
    param(
        [string]$Value,
        [int]$DefaultValue = 0
    )

    if ([string]::IsNullOrWhiteSpace($Value)) {
        return $DefaultValue
    }

    # Remove commas and whitespace
    $cleanValue = $Value -replace '[,\s]', ''

    # Try to parse as integer
    $result = 0
    if ([int]::TryParse($cleanValue, [ref]$result)) {
        return $result
    }

    # Try to parse as double and round
    $doubleResult = 0.0
    if ([double]::TryParse($cleanValue, [ref]$doubleResult)) {
        return [int][Math]::Round($doubleResult)
    }

    return $DefaultValue
}

function Get-SafeDoubleValue {
    param(
        [string]$Value,
        [double]$DefaultValue = 0.0
    )

    if ([string]::IsNullOrWhiteSpace($Value)) {
        return $DefaultValue
    }

    # Remove commas and whitespace
    $cleanValue = $Value -replace '[,\s]', ''

    # Try to parse as double
    $result = 0.0
    if ([double]::TryParse($cleanValue, [ref]$result)) {
        return $result
    }

    return $DefaultValue
}

function Test-SheetMatch {
    param(
        [string]$SheetName,
        [string[]]$Patterns
    )

    foreach ($pattern in $Patterns) {
        if ($SheetName -like $pattern) {
            return $true
        }
    }
    return $false
}

function Read-SheetData {
    param(
        [object]$Worksheet,
        [hashtable]$ColumnIndices
    )

    $data = @()
    $usedRange = $Worksheet.UsedRange
    $rowCount = $usedRange.Rows.Count

    if ($rowCount -le 1) {
        return $data
    }

    # Read data rows with flexible parsing
    for ($row = 2; $row -le $rowCount; $row++) {
        # Show progress for large datasets
        if ($row % 100 -eq 0) {
            Write-Log "  Processing row $row of $rowCount..."
        }

        $rowData = @{
            'Host Name' = ''
            'IP' = ''
            'Product' = ''
            'Critical' = 0
            'High' = 0
            'Medium' = 0
            'Low' = 0
            'Vulnerability Count' = 0
            'EPSS Score' = 0.0
        }

        # Read HostName
        if ($columnIndices.ContainsKey('HostName')) {
            $rowData['Host Name'] = $Worksheet.Cells.Item($row, $columnIndices['HostName']).Text
        }

        # Read IP
        if ($columnIndices.ContainsKey('IP')) {
            $rowData['IP'] = $Worksheet.Cells.Item($row, $columnIndices['IP']).Text
        }

        # Read Product (required)
        if ($columnIndices.ContainsKey('Product')) {
            $rowData['Product'] = $Worksheet.Cells.Item($row, $columnIndices['Product']).Text
        }

        # Skip rows with no product name
        if ([string]::IsNullOrWhiteSpace($rowData['Product'])) {
            continue
        }

        # Read severity counts
        if ($columnIndices.ContainsKey('Critical')) {
            $rowData['Critical'] = Get-SafeNumericValue -Value $Worksheet.Cells.Item($row, $columnIndices['Critical']).Text
        }

        if ($columnIndices.ContainsKey('High')) {
            $rowData['High'] = Get-SafeNumericValue -Value $Worksheet.Cells.Item($row, $columnIndices['High']).Text
        }

        if ($columnIndices.ContainsKey('Medium')) {
            $rowData['Medium'] = Get-SafeNumericValue -Value $Worksheet.Cells.Item($row, $columnIndices['Medium']).Text
        }

        if ($columnIndices.ContainsKey('Low')) {
            $rowData['Low'] = Get-SafeNumericValue -Value $Worksheet.Cells.Item($row, $columnIndices['Low']).Text
        }

        # Read Vulnerability Count
        if ($columnIndices.ContainsKey('VulnCount')) {
            $rowData['Vulnerability Count'] = Get-SafeNumericValue -Value $Worksheet.Cells.Item($row, $columnIndices['VulnCount']).Text
        } else {
            # Calculate from severity counts if not provided
            $rowData['Vulnerability Count'] = $rowData['Critical'] + $rowData['High'] + $rowData['Medium'] + $rowData['Low']
        }

        # Read EPSS Score
        if ($columnIndices.ContainsKey('EPSS')) {
            $rowData['EPSS Score'] = Get-SafeDoubleValue -Value $Worksheet.Cells.Item($row, $columnIndices['EPSS']).Text
        }

        # Only add rows that have at least one vulnerability
        if ($rowData['Vulnerability Count'] -gt 0) {
            $data += [PSCustomObject]$rowData
        }
    }

    return $data
}

function Get-VulnerabilityData {
    param(
        [string]$ExcelPath
    )

    Write-Log "Reading vulnerability data from Excel..."
    Write-Log "Auto-detecting and consolidating remediation sheets..."

    $excel = $null
    $workbook = $null
    $allData = @()

    try {
        $excel = New-Object -ComObject Excel.Application
        $excel.Visible = $false
        $excel.DisplayAlerts = $false

        $workbook = $excel.Workbooks.Open($ExcelPath)

        # Find all sheets that match remediation patterns
        $sourceSheets = @()
        foreach ($sheet in $workbook.Worksheets) {
            $sheetName = $sheet.Name

            # Skip excluded sheets
            $shouldExclude = $false
            foreach ($excludePattern in $script:Config.ExcludeSheetPatterns) {
                if ($sheetName -like $excludePattern -or $sheetName -eq $excludePattern) {
                    $shouldExclude = $true
                    break
                }
            }

            if ($shouldExclude) {
                Write-Log "Excluding sheet: $sheetName"
                Clear-ComObject $sheet
                continue
            }

            # Check if sheet matches any remediation pattern
            $isMatch = Test-SheetMatch -SheetName $sheetName -Patterns $script:Config.SourceSheetPatterns

            if ($isMatch) {
                Write-Log "Found remediation sheet: $sheetName"
                $sourceSheets += $sheet
            } else {
                Clear-ComObject $sheet
            }
        }

        if ($sourceSheets.Count -eq 0) {
            throw "No remediation sheets found. Looking for patterns: $($script:Config.SourceSheetPatterns -join ', '). Excluding: $($script:Config.ExcludeSheetPatterns -join ', ')"
        }

        Write-Log "Processing $($sourceSheets.Count) remediation sheet(s)..."

        # Get headers from first sheet and create column mappings
        $firstSheet = $sourceSheets[0]
        $usedRange = $firstSheet.UsedRange
        $colCount = $usedRange.Columns.Count

        # Get headers
        $headers = @{}
        for ($col = 1; $col -le $colCount; $col++) {
            $headerName = $firstSheet.Cells.Item(1, $col).Text
            if ($headerName) {
                $headers[$headerName] = $col
            }
        }

        Write-Log "Found headers: $($headers.Keys -join ', ')"

        # Define flexible column mappings
        $columnMappings = @{
            'HostName' = @('Host Name', 'Hostname', 'Computer', 'Computer Name', 'Device', 'Device Name', 'System', 'System Name', 'Machine')
            'IP' = @('IP', 'IP Address', 'IPAddress', 'Address')
            'Product' = @('Product', 'Software', 'Application', 'App', 'Program', 'Title', 'Product Name', 'Software Name')
            'Critical' = @('Critical', 'Crit', 'Critical Count', 'Critical Vulnerabilities')
            'High' = @('High', 'High Count', 'High Vulnerabilities')
            'Medium' = @('Medium', 'Med', 'Medium Count', 'Medium Vulnerabilities')
            'Low' = @('Low', 'Low Count', 'Low Vulnerabilities')
            'VulnCount' = @('Vulnerability Count', 'Vuln Count', 'Total Vulnerabilities', 'Total Vulns', 'Count', 'Total Count', 'Number of Vulnerabilities')
            'EPSS' = @('EPSS Score', 'EPSS', 'Exploit Prediction Score', 'Max EPSS Score', 'Max EPSS')
        }

        # Find column indices
        $columnIndices = @{}
        foreach ($key in $columnMappings.Keys) {
            $colIndex = Find-ColumnIndex -Headers $headers -PossibleNames $columnMappings[$key]
            if ($colIndex) {
                $columnIndices[$key] = $colIndex
                Write-Log "Mapped '$key' to column: $($headers.Keys | Where-Object { $headers[$_] -eq $colIndex })"
            } else {
                Write-Log "Could not find column for '$key' (tried: $($columnMappings[$key] -join ', '))" -Level Warning
            }
        }

        # Verify minimum required columns
        $requiredFields = @('Product')
        $missingRequired = @()
        foreach ($field in $requiredFields) {
            if (-not $columnIndices.ContainsKey($field)) {
                $missingRequired += $field
            }
        }

        if ($missingRequired.Count -gt 0) {
            throw "Missing required columns: $($missingRequired -join ', '). Please ensure your Excel file has at least a Product/Software column."
        }

        Write-Log "Successfully mapped $($columnIndices.Count) columns."

        # Read data from all matching sheets
        foreach ($sheet in $sourceSheets) {
            Write-Log "Reading data from: $($sheet.Name)"
            $sheetData = Read-SheetData -Worksheet $sheet -ColumnIndices $columnIndices
            Write-Log "  Found $($sheetData.Count) vulnerability records"
            $allData += $sheetData
        }

        Write-Log "Total vulnerability records consolidated: $($allData.Count)" -Level Success

        # Clean up sheet references
        foreach ($sheet in $sourceSheets) {
            Clear-ComObject $sheet
        }

        return $allData

    } catch {
        Write-Log "Error reading Excel data: $($_.Exception.Message)" -Level Error
        throw
    } finally {
        if ($workbook) {
            $workbook.Close($false)
            Clear-ComObject $workbook
        }
        if ($excel) {
            $excel.Quit()
            Clear-ComObject $excel
        }
        [System.GC]::Collect()
        [System.GC]::WaitForPendingFinalizers()
    }
}

function Get-ConsolidatedProduct {
    param([string]$ProductName)

    if ([string]::IsNullOrWhiteSpace($ProductName)) {
        return $ProductName
    }

    # Normalize the product name for comparison
    $normalizedProduct = $ProductName.Trim()

    # Check against consolidation rules (case-insensitive)
    foreach ($consolidated in $script:Config.WindowsConsolidation.Keys) {
        $patterns = $script:Config.WindowsConsolidation[$consolidated]
        foreach ($pattern in $patterns) {
            # Try exact match (case-insensitive)
            if ($normalizedProduct -eq $pattern) {
                return $consolidated
            }
            # Try wildcard match
            if ($normalizedProduct -like "*$pattern*") {
                return $consolidated
            }
        }
    }

    # Additional common product normalization
    # Remove version numbers at the end (e.g., "Adobe Reader 11.0.23" -> "Adobe Reader")
    if ($normalizedProduct -match '^(.+?)\s+[\d\.]+$') {
        $baseProduct = $matches[1]
        # Check if this base product should be consolidated
        foreach ($consolidated in $script:Config.WindowsConsolidation.Keys) {
            $patterns = $script:Config.WindowsConsolidation[$consolidated]
            foreach ($pattern in $patterns) {
                if ($baseProduct -like "*$pattern*") {
                    return $consolidated
                }
            }
        }
    }

    return $ProductName
}

function Get-AverageCVSS {
    param(
        [int]$Critical,
        [int]$High,
        [int]$Medium,
        [int]$Low
    )

    $total = $Critical + $High + $Medium + $Low

    if ($total -eq 0) {
        return 0
    }

    $weighted = ($Critical * $script:Config.CVSSEquivalent.Critical) +
                ($High * $script:Config.CVSSEquivalent.High) +
                ($Medium * $script:Config.CVSSEquivalent.Medium) +
                ($Low * $script:Config.CVSSEquivalent.Low)

    return [Math]::Round($weighted / $total, 2)
}

function Get-CompositeRiskScore {
    param(
        [int]$VulnCount,
        [double]$EPSSScore,
        [double]$AvgCVSS
    )

    return [Math]::Round($VulnCount * $EPSSScore * $AvgCVSS, 2)
}

function Get-Top10Vulnerabilities {
    param([array]$VulnData)

    Write-Log "Calculating risk scores and identifying top 10 vulnerabilities..."

    # Group by product
    $grouped = $VulnData | Group-Object -Property Product

    $aggregated = @()

    foreach ($group in $grouped) {
        $product = $group.Name

        # Check if product should be filtered
        $shouldFilter = $false
        foreach ($filter in $script:Config.FilteredProducts) {
            if ($product -like "*$filter*") {
                $shouldFilter = $true
                break
            }
        }

        if ($shouldFilter) {
            Write-Log "Filtering out: $product"
            continue
        }

        # Consolidate Windows versions
        $consolidatedProduct = Get-ConsolidatedProduct -ProductName $product

        # Check if we already have this consolidated product
        $existing = $aggregated | Where-Object { $_.Product -eq $consolidatedProduct }

        if ($existing) {
            # Merge with existing
            $existing.Critical += ($group.Group | Measure-Object -Property Critical -Sum).Sum
            $existing.High += ($group.Group | Measure-Object -Property High -Sum).Sum
            $existing.Medium += ($group.Group | Measure-Object -Property Medium -Sum).Sum
            $existing.Low += ($group.Group | Measure-Object -Property Low -Sum).Sum
            $existing.VulnCount += ($group.Group | Measure-Object -Property 'Vulnerability Count' -Sum).Sum

            # Take max EPSS score
            $maxEPSS = ($group.Group.'EPSS Score' | Measure-Object -Maximum).Maximum
            if ($maxEPSS -gt $existing.EPSSScore) {
                $existing.EPSSScore = $maxEPSS
            }

            # Add affected systems
            $existing.AffectedSystems += $group.Group.'Host Name'
        } else {
            # Create new entry
            $critical = ($group.Group | Measure-Object -Property Critical -Sum).Sum
            $high = ($group.Group | Measure-Object -Property High -Sum).Sum
            $medium = ($group.Group | Measure-Object -Property Medium -Sum).Sum
            $low = ($group.Group | Measure-Object -Property Low -Sum).Sum
            $vulnCount = ($group.Group | Measure-Object -Property 'Vulnerability Count' -Sum).Sum
            $epssScore = ($group.Group.'EPSS Score' | Measure-Object -Maximum).Maximum

            $avgCVSS = Get-AverageCVSS -Critical $critical -High $high -Medium $medium -Low $low
            $riskScore = Get-CompositeRiskScore -VulnCount $vulnCount -EPSSScore $epssScore -AvgCVSS $avgCVSS

            $aggregated += [PSCustomObject]@{
                Product = $consolidatedProduct
                Critical = $critical
                High = $high
                Medium = $medium
                Low = $low
                VulnCount = $vulnCount
                EPSSScore = $epssScore
                AvgCVSS = $avgCVSS
                RiskScore = $riskScore
                AffectedSystems = @($group.Group.'Host Name')
            }
        }
    }

    # Recalculate scores for consolidated entries
    foreach ($item in $aggregated) {
        $item.AvgCVSS = Get-AverageCVSS -Critical $item.Critical -High $item.High -Medium $item.Medium -Low $item.Low
        $item.RiskScore = Get-CompositeRiskScore -VulnCount $item.VulnCount -EPSSScore $item.EPSSScore -AvgCVSS $item.AvgCVSS
    }

    # Sort by risk score and take top 10
    $top10 = $aggregated | Sort-Object -Property RiskScore -Descending | Select-Object -First 10

    Write-Log "Identified top 10 vulnerabilities from $($aggregated.Count) unique products"

    return $top10
}

function New-WordReport {
    param(
        [string]$OutputPath,
        [string]$ClientName,
        [string]$ScanDate,
        [array]$Top10Data
    )

    Write-Log "Generating Word document report..."

    $word = $null
    $doc = $null

    try {
        $word = New-Object -ComObject Word.Application
        $word.Visible = $false
        $doc = $word.Documents.Add()

        # Set document properties
        $doc.BuiltInDocumentProperties.Item("Title").Value = "Vulnerability Assessment Report - $ClientName"
        $doc.BuiltInDocumentProperties.Item("Subject").Value = "Security Vulnerability Assessment"
        $doc.BuiltInDocumentProperties.Item("Author").Value = $script:Config.Author
        $doc.BuiltInDocumentProperties.Item("Keywords").Value = "Vulnerability, Security, Assessment, EPSS, CVSS"

        # Set page margins (in points: 1 inch = 72 points)
        # Default margins are usually 1 inch (72 points)
        # Setting to 0.5 inch (36 points) for left and right
        $doc.PageSetup.LeftMargin = 36    # 0.5 inch
        $doc.PageSetup.RightMargin = 36   # 0.5 inch
        $doc.PageSetup.TopMargin = 72     # 1 inch
        $doc.PageSetup.BottomMargin = 72  # 1 inch

        # --- Title Page ---
        Write-Log "Creating title page..."
        $selection = $word.Selection

        # Add some top spacing for title page
        $selection.TypeParagraph()
        $selection.TypeParagraph()
        $selection.TypeParagraph()

        $selection.Font.Name = "Calibri"
        $selection.Font.Size = 32
        $selection.Font.Bold = $true
        $selection.Font.Color = 5855577  # Dark blue color
        $selection.ParagraphFormat.Alignment = 1  # Center
        $selection.TypeText("Vulnerability Assessment Report")
        $selection.TypeParagraph()
        $selection.TypeParagraph()

        # Add horizontal line
        $selection.ParagraphFormat.Borders.Item(3).LineStyle = 1  # Bottom border
        $selection.ParagraphFormat.Borders.Item(3).LineWidth = 24  # Thicker line
        $selection.ParagraphFormat.Borders.Item(3).Color = 5855577  # Dark blue
        $selection.TypeParagraph()
        $selection.ParagraphFormat.Borders.Item(3).LineStyle = 0  # Reset border
        $selection.TypeParagraph()
        $selection.TypeParagraph()

        $selection.Font.Size = 20
        $selection.Font.Bold = $true
        $selection.Font.Color = 0  # Black
        $selection.TypeText($ClientName)
        $selection.TypeParagraph()
        $selection.TypeParagraph()
        $selection.TypeParagraph()

        $selection.Font.Size = 14
        $selection.Font.Bold = $false
        $selection.TypeText("Scan Date: $ScanDate")
        $selection.TypeParagraph()
        $selection.TypeParagraph()

        $selection.Font.Size = 12
        $selection.TypeText("Prepared by: $($script:Config.Author)")
        $selection.TypeParagraph()

        $selection.InsertBreak(7)  # Page break

        # --- Executive Summary ---
        Write-Log "Creating executive summary..."
        $selection.ParagraphFormat.Alignment = 0  # Left align
        $selection.Style = "Heading 1"
        $selection.TypeText("Executive Summary")
        $selection.TypeParagraph()

        $selection.Style = "Normal"
        $selection.Font.Size = 11
        $selection.TypeParagraph()

        $selection.Font.Size = 11
        $selection.Font.Bold = $false
        $selection.TypeText("This vulnerability assessment report summarizes the security posture of $ClientName based on the vulnerability scan conducted on $ScanDate. ")
        $selection.TypeText("The organization utilizes ConnectWise Automate for patch management. WSUS is not currently in use.")
        $selection.TypeParagraph()
        $selection.TypeParagraph()

        $selection.TypeText("This report identifies the top 10 security risks based on a composite risk score that considers vulnerability count, ")
        $selection.TypeText("EPSS (Exploit Prediction Scoring System) scores, and CVSS severity ratings. ")
        $selection.TypeText("Each finding includes specific remediation guidance appropriate for the environment.")
        $selection.TypeParagraph()
        $selection.TypeParagraph()

        # --- Scoring Methodology ---
        Write-Log "Creating scoring methodology section..."
        $selection.Style = "Heading 1"
        $selection.TypeText("Scoring Methodology")
        $selection.TypeParagraph()

        $selection.Style = "Normal"
        $selection.Font.Size = 11
        $selection.TypeParagraph()

        $selection.Font.Size = 11
        $selection.Font.Bold = $false
        $selection.TypeText("The Composite Risk Score is calculated using the following formula:")
        $selection.TypeParagraph()
        $selection.TypeParagraph()

        # Add shaded box for formula
        $selection.ParagraphFormat.Shading.BackgroundPatternColor = 15329769  # Light blue/gray
        $selection.ParagraphFormat.LeftIndent = 18
        $selection.ParagraphFormat.RightIndent = 18
        $selection.ParagraphFormat.SpaceBefore = 6
        $selection.ParagraphFormat.SpaceAfter = 6

        $selection.Font.Name = "Courier New"
        $selection.Font.Bold = $true
        $selection.Font.Size = 10
        $selection.TypeText("Risk Score = Vulnerability Count x EPSS Score x Average CVSS Equivalent")
        $selection.TypeParagraph()

        $selection.Font.Bold = $false
        $selection.TypeText("Where Average CVSS is calculated as:")
        $selection.TypeParagraph()

        $selection.TypeText("(Critical x 9.0 + High x 7.0 + Medium x 5.0 + Low x 3.0) / Total Vulnerabilities")
        $selection.TypeParagraph()

        # Reset formatting
        $selection.ParagraphFormat.Shading.BackgroundPatternColor = 16777215  # White
        $selection.ParagraphFormat.LeftIndent = 0
        $selection.ParagraphFormat.RightIndent = 0
        $selection.ParagraphFormat.SpaceBefore = 0
        $selection.ParagraphFormat.SpaceAfter = 0
        $selection.Font.Name = "Calibri"
        $selection.Font.Size = 11
        $selection.TypeParagraph()

        # --- Risk Score Color Legend ---
        Write-Log "Creating color legend..."
        $selection.Style = "Heading 2"
        $selection.TypeText("Risk Score Color Legend")
        $selection.TypeParagraph()

        $selection.Style = "Normal"
        $selection.TypeParagraph()

        # Create legend table with heatmap gradient (7 levels)
        $legendTable = $doc.Tables.Add($selection.Range, 7, 2)
        $legendTable.Borders.Enable = $true
        $legendTable.Range.Font.Size = 10

        # Row 1: Critical
        $legendTable.Cell(1, 1).Range.Text = $script:Config.RiskColors.Critical.Name
        $legendTable.Cell(1, 1).Shading.BackgroundPatternColor = ConvertTo-HexColor -HexColor $script:Config.RiskColors.Critical.Color
        $legendTable.Cell(1, 1).Range.Font.Color = ConvertTo-HexColor -HexColor $script:Config.RiskColors.Critical.TextColor
        $legendTable.Cell(1, 1).Range.Font.Bold = $true
        $legendTable.Cell(1, 2).Range.Text = "Risk Score >= $($script:Config.RiskColors.Critical.Threshold)"

        # Row 2: Very High
        $legendTable.Cell(2, 1).Range.Text = $script:Config.RiskColors.VeryHigh.Name
        $legendTable.Cell(2, 1).Shading.BackgroundPatternColor = ConvertTo-HexColor -HexColor $script:Config.RiskColors.VeryHigh.Color
        $legendTable.Cell(2, 1).Range.Font.Color = ConvertTo-HexColor -HexColor $script:Config.RiskColors.VeryHigh.TextColor
        $legendTable.Cell(2, 1).Range.Font.Bold = $true
        $legendTable.Cell(2, 2).Range.Text = "Risk Score >= $($script:Config.RiskColors.VeryHigh.Threshold)"

        # Row 3: High
        $legendTable.Cell(3, 1).Range.Text = $script:Config.RiskColors.High.Name
        $legendTable.Cell(3, 1).Shading.BackgroundPatternColor = ConvertTo-HexColor -HexColor $script:Config.RiskColors.High.Color
        $legendTable.Cell(3, 1).Range.Font.Color = ConvertTo-HexColor -HexColor $script:Config.RiskColors.High.TextColor
        $legendTable.Cell(3, 1).Range.Font.Bold = $true
        $legendTable.Cell(3, 2).Range.Text = "Risk Score >= $($script:Config.RiskColors.High.Threshold)"

        # Row 4: Medium-High
        $legendTable.Cell(4, 1).Range.Text = $script:Config.RiskColors.MediumHigh.Name
        $legendTable.Cell(4, 1).Shading.BackgroundPatternColor = ConvertTo-HexColor -HexColor $script:Config.RiskColors.MediumHigh.Color
        $legendTable.Cell(4, 1).Range.Font.Color = ConvertTo-HexColor -HexColor $script:Config.RiskColors.MediumHigh.TextColor
        $legendTable.Cell(4, 1).Range.Font.Bold = $true
        $legendTable.Cell(4, 2).Range.Text = "Risk Score >= $($script:Config.RiskColors.MediumHigh.Threshold)"

        # Row 5: Medium
        $legendTable.Cell(5, 1).Range.Text = $script:Config.RiskColors.Medium.Name
        $legendTable.Cell(5, 1).Shading.BackgroundPatternColor = ConvertTo-HexColor -HexColor $script:Config.RiskColors.Medium.Color
        $legendTable.Cell(5, 1).Range.Font.Color = ConvertTo-HexColor -HexColor $script:Config.RiskColors.Medium.TextColor
        $legendTable.Cell(5, 1).Range.Font.Bold = $true
        $legendTable.Cell(5, 2).Range.Text = "Risk Score >= $($script:Config.RiskColors.Medium.Threshold)"

        # Row 6: Low
        $legendTable.Cell(6, 1).Range.Text = $script:Config.RiskColors.Low.Name
        $legendTable.Cell(6, 1).Shading.BackgroundPatternColor = ConvertTo-HexColor -HexColor $script:Config.RiskColors.Low.Color
        $legendTable.Cell(6, 1).Range.Font.Color = ConvertTo-HexColor -HexColor $script:Config.RiskColors.Low.TextColor
        $legendTable.Cell(6, 1).Range.Font.Bold = $true
        $legendTable.Cell(6, 2).Range.Text = "Risk Score >= $($script:Config.RiskColors.Low.Threshold)"

        # Row 7: Very Low
        $legendTable.Cell(7, 1).Range.Text = $script:Config.RiskColors.VeryLow.Name
        $legendTable.Cell(7, 1).Shading.BackgroundPatternColor = ConvertTo-HexColor -HexColor $script:Config.RiskColors.VeryLow.Color
        $legendTable.Cell(7, 1).Range.Font.Color = ConvertTo-HexColor -HexColor $script:Config.RiskColors.VeryLow.TextColor
        $legendTable.Cell(7, 1).Range.Font.Bold = $true
        $legendTable.Cell(7, 2).Range.Text = "Risk Score >= $($script:Config.RiskColors.VeryLow.Threshold)"

        # AutoFit the legend table
        $legendTable.AutoFitBehavior(1)  # 1 = wdAutoFitContent (fit to content)

        $selection.EndKey(6)  # Move to end of document
        $selection.TypeParagraph()
        $selection.TypeParagraph()

        $selection.InsertBreak(7)  # Page break before Top 10 table

        # --- Top 10 Vulnerabilities Table ---
        Write-Log "Creating top 10 vulnerabilities table..."
        $selection.Style = "Heading 1"
        $selection.TypeText("Top 10 Vulnerabilities by Risk Score")
        $selection.TypeParagraph()

        $selection.Style = "Normal"
        $selection.TypeParagraph()

        # Create table
        $table = $doc.Tables.Add($selection.Range, ($Top10Data.Count + 1), 7)
        $table.Borders.Enable = $true
        $table.Style = "Grid Table 4 - Accent 1"

        # Set table font size to 9 points for better fit
        $table.Range.Font.Size = 9

        # Headers
        $headers = @("Rank", "Product/System", "Risk Score", "EPSS", "Avg CVSS", "Total Vulns", "Affected Systems")
        for ($i = 0; $i -lt $headers.Count; $i++) {
            $table.Cell(1, $i + 1).Range.Text = $headers[$i]
            $table.Cell(1, $i + 1).Range.Font.Bold = $true
        }

        # Data rows
        $rank = 1
        foreach ($item in $Top10Data) {
            $rowIndex = $rank + 1

            $table.Cell($rowIndex, 1).Range.Text = $rank.ToString()
            $table.Cell($rowIndex, 2).Range.Text = $item.Product
            $table.Cell($rowIndex, 3).Range.Text = $item.RiskScore.ToString("N2")
            $table.Cell($rowIndex, 4).Range.Text = $item.EPSSScore.ToString("N4")
            $table.Cell($rowIndex, 5).Range.Text = $item.AvgCVSS.ToString("N2")
            $table.Cell($rowIndex, 6).Range.Text = $item.VulnCount.ToString()
            $table.Cell($rowIndex, 7).Range.Text = $item.AffectedSystems.Count.ToString()

            # Apply color coding based on risk score
            $colorInfo = Get-RiskScoreColor -RiskScore $item.RiskScore
            $bgColor = ConvertTo-HexColor -HexColor $colorInfo.Color
            $textColor = ConvertTo-HexColor -HexColor $colorInfo.TextColor

            for ($col = 1; $col -le 7; $col++) {
                $table.Cell($rowIndex, $col).Shading.BackgroundPatternColor = $bgColor
                $table.Cell($rowIndex, $col).Range.Font.Color = $textColor
            }

            $rank++
        }

        # AutoFit the table for better appearance
        $table.AutoFitBehavior(2)  # 2 = wdAutoFitWindow (fit to window)

        $selection.EndKey(6)
        $selection.TypeParagraph()
        $selection.TypeParagraph()

        $selection.InsertBreak(7)  # Page break before Detailed Findings

        # --- Detailed Findings ---
        Write-Log "Creating detailed findings section..."
        $selection.Style = "Heading 1"
        $selection.Font.Color = 0  # Black
        $selection.TypeText("Detailed Findings and Remediation Guidance")
        $selection.TypeParagraph()

        $selection.Style = "Normal"
        $selection.TypeParagraph()

        $rank = 1
        foreach ($item in $Top10Data) {
            # Vulnerability title
            $selection.Style = "Heading 2"
            $selection.TypeText("$rank. $($item.Product)")
            $selection.TypeParagraph()

            $selection.Style = "Normal"
            $selection.Font.Size = 11

            # Risk metrics in a more compact format
            $selection.Font.Bold = $true
            $selection.TypeText("Risk Metrics:")
            $selection.TypeParagraph()
            $selection.Font.Bold = $false

            $selection.ParagraphFormat.LeftIndent = 36  # Indent for better readability
            $selection.TypeText("Risk Score: $($item.RiskScore.ToString('N2'))")
            $selection.TypeParagraph()
            $selection.TypeText("EPSS Score: $($item.EPSSScore.ToString('N4'))")
            $selection.TypeParagraph()
            $selection.TypeText("Average CVSS: $($item.AvgCVSS.ToString('N2'))")
            $selection.TypeParagraph()
            $selection.TypeText("Total Vulnerabilities: $($item.VulnCount)")
            $selection.TypeParagraph()
            $selection.TypeText("Affected Systems: $($item.AffectedSystems.Count)")
            $selection.TypeParagraph()
            $selection.ParagraphFormat.LeftIndent = 0  # Reset indent
            $selection.TypeParagraph()

            # Affected systems list
            $selection.Font.Bold = $true
            $selection.TypeText("Affected Systems:")
            $selection.TypeParagraph()
            $selection.Font.Bold = $false

            # Display systems as comma-separated list with indent
            $selection.ParagraphFormat.LeftIndent = 36
            $systemsList = ($item.AffectedSystems | Select-Object -Unique) -join ", "
            $selection.TypeText($systemsList)
            $selection.TypeParagraph()
            $selection.ParagraphFormat.LeftIndent = 0
            $selection.TypeParagraph()

            # Remediation guidance
            $selection.Font.Bold = $true
            $selection.TypeText("Remediation Guidance:")
            $selection.TypeParagraph()
            $selection.Font.Bold = $false

            $selection.ParagraphFormat.LeftIndent = 36

            # Determine remediation type
            if ($item.Product -like "*Windows Server 2012*" -or $item.Product -like "*end-of-life*" -or $item.Product -like "*out of support*") {
                $selection.TypeText("This end-of-support operating system represents an infrastructure project beyond the scope of quarterly vulnerability remediation. ")
                $selection.TypeText("Consider planning a migration to a supported operating system version.")
            } elseif ($item.Product -like "*Windows*") {
                $selection.TypeText("Windows patch inconsistencies should be investigated via ConnectWise Automate. ")
                $selection.TypeText("Systems with lower vulnerability counts may indicate that patching is working correctly and awaiting the latest patch cycles. ")
                $selection.TypeText("For systems with high vulnerability counts, verify Windows Update status and investigate any potential issues preventing patch installation.")
            } elseif ($item.Product -like "*printer*" -or $item.Product -like "*Ripple20*") {
                $selection.TypeText("Network printers and IoT devices require manual firmware updates via manufacturer-provided tools and interfaces. ")
                $selection.TypeText("Consult the manufacturer's documentation for firmware update procedures.")
            } elseif ($item.Product -like "*Microsoft Teams*") {
                $selection.TypeText("Microsoft Teams can be updated via RMM script deployed through ConnectWise Automate.")
            } else {
                $selection.TypeText("This application should be updated to the latest version. ")
                $selection.TypeText("If available via ConnectWise Automate/RMM, deploy updates using the patch management system. ")
                $selection.TypeText("Otherwise, manual updates may be required on affected systems.")
            }

            $selection.ParagraphFormat.LeftIndent = 0  # Reset indent
            $selection.TypeParagraph()
            $selection.TypeParagraph()

            # Add a subtle separator between items (except after the last one)
            if ($rank -lt $Top10Data.Count) {
                $selection.ParagraphFormat.Borders.Item(3).LineStyle = 1  # Bottom border
                $selection.ParagraphFormat.Borders.Item(3).LineWidth = 1
                $selection.ParagraphFormat.Borders.Item(3).Color = 12632256  # Light gray
                $selection.TypeParagraph()
                $selection.ParagraphFormat.Borders.Item(3).LineStyle = 0  # Reset border
            }

            $rank++
        }

        # Save document
        Write-Log "Saving document to: $OutputPath"

        # Delete existing file if present
        if (Test-Path $OutputPath) {
            try {
                Remove-Item -Path $OutputPath -Force -ErrorAction Stop
                Write-Log "Removed existing file: $OutputPath"
            } catch {
                throw "Cannot overwrite existing file '$OutputPath'. Please close it if it's open and try again."
            }
        }

        $doc.SaveAs([ref]$OutputPath, [ref]16)  # 16 = wdFormatDocumentDefault (.docx)

        Write-Log "Word document generated successfully" -Level Success

    } catch {
        Write-Log "Error generating Word document: $($_.Exception.Message)" -Level Error
        throw
    } finally {
        if ($doc) {
            $doc.Close($false)
            Clear-ComObject $doc
        }
        if ($word) {
            $word.Quit()
            Clear-ComObject $word
        }
        [System.GC]::Collect()
        [System.GC]::WaitForPendingFinalizers()
    }
}

# Function to generate Excel report with pivot table
function New-ExcelReport {
    param(
        [string]$InputPath,
        [string]$OutputPath
    )

    Write-Log "Starting Excel report generation..." -Level Info

    $excel = $null
    $workbook = $null
    $sourceDataSheet = $null
    $pivotSheet = $null

    try {
        # Create Excel COM Object
        Write-Log "Creating Excel COM object..." -Level Info
        $excel = New-Object -ComObject Excel.Application
        if ($null -eq $excel) {
            throw "Failed to create Excel COM object. Make sure Microsoft Excel is installed."
        }

        $excel.Visible = $false
        $excel.DisplayAlerts = $false
        $excel.ScreenUpdating = $false

        # Open Input Workbook
        Write-Log "Opening input workbook..." -Level Info
        $workbook = $excel.Workbooks.Open($InputPath)
        if ($null -eq $workbook) {
            throw "Failed to open workbook. File may be corrupted or in use."
        }

        # --- 1. Autofit Columns and Rows ---
        Write-Log "Auto-fitting columns and rows..." -Level Info
        foreach ($worksheet in $workbook.Worksheets) {
            if ($worksheet.Name -ne $script:Config.SheetToExcludeFormatting) {
                try {
                    $worksheet.UsedRange.Columns.AutoFit() | Out-Null
                    $worksheet.UsedRange.Rows.AutoFit() | Out-Null
                } catch {
                    Write-Log "Could not AutoFit sheet '$($worksheet.Name)'" -Level Warning
                }
            }
            Clear-ComObject $worksheet
        }

        # --- 2. Consolidate Data ---
        Write-Log "Consolidating data from remediation sheets..." -Level Info

        # Delete existing consolidated sheet if present
        $existingSheet = $null
        foreach ($sheet in $workbook.Worksheets) {
            if ($sheet.Name -eq $script:Config.ConsolidatedSheetName) {
                $existingSheet = $sheet
                break
            }
        }

        if ($null -ne $existingSheet) {
            Write-Log "Deleting existing '$($script:Config.ConsolidatedSheetName)' sheet..." -Level Info
            $existingSheet.Delete()
            Clear-ComObject $existingSheet
        }

        # Create new consolidated sheet
        $lastSheet = $workbook.Worksheets[$workbook.Worksheets.Count]
        $sourceDataSheet = $workbook.Worksheets.Add([System.Reflection.Missing]::Value, $lastSheet)
        $sourceDataSheet.Name = $script:Config.ConsolidatedSheetName
        Write-Log "Created '$($script:Config.ConsolidatedSheetName)' sheet" -Level Info
        Clear-ComObject $lastSheet

        # Find and collect source sheets
        $sourceSheets = @()
        $firstValidSheet = $null

        foreach ($pattern in $script:Config.SourceSheetPatterns) {
            foreach ($sheet in $workbook.Worksheets) {
                $shouldInclude = $false

                # Check if matches pattern
                if ($sheet.Name -like $pattern) {
                    $shouldInclude = $true
                }

                # Check exclusions
                foreach ($excludePattern in $script:Config.ExcludeSheetPatterns) {
                    if ($sheet.Name -like $excludePattern -or $sheet.Name -eq $excludePattern) {
                        $shouldInclude = $false
                        break
                    }
                }

                if ($shouldInclude) {
                    try {
                        if ($null -ne $sheet.UsedRange -and $sheet.UsedRange.Rows.Count -ge 1) {
                            $sourceSheets += $sheet
                            if ($null -eq $firstValidSheet) {
                                $firstValidSheet = $sheet
                                Write-Log "Found first valid sheet: $($firstValidSheet.Name)" -Level Info
                            }
                        }
                    } catch {
                        Write-Log "Could not evaluate sheet '$($sheet.Name)'" -Level Warning
                    }
                }
            }
        }

        Write-Log "Found $($sourceSheets.Count) source sheets to consolidate" -Level Info

        # Copy headers from first valid sheet
        if ($null -ne $firstValidSheet) {
            try {
                $sourceCols = $firstValidSheet.UsedRange.Columns.Count
                $headerRange = $firstValidSheet.Range("A1", $firstValidSheet.Cells.Item(1, $sourceCols))
                $headerValues = $headerRange.Value2
                $targetRange = $sourceDataSheet.Range("A1", $sourceDataSheet.Cells.Item(1, $sourceCols))
                $targetRange.Value2 = $headerValues
                Write-Log "Headers copied successfully" -Level Info
                Clear-ComObject $headerRange
                Clear-ComObject $targetRange
            } catch {
                throw "Failed to copy headers: $($_.Exception.Message)"
            }
        }

        # Copy data rows from all source sheets
        $destRow = 2
        foreach ($sourceSheet in $sourceSheets) {
            Write-Log "Copying data from: $($sourceSheet.Name)" -Level Info

            try {
                $sourceRange = $sourceSheet.UsedRange
                $sourceRows = $sourceRange.Rows.Count

                if ($sourceRows -gt 1) {
                    $sourceCols = $sourceRange.Columns.Count
                    $dataRange = $sourceSheet.Range("A2", $sourceSheet.Cells.Item($sourceRows, $sourceCols))
                    $dataValues = $dataRange.Value2
                    $targetRowCount = $dataRange.Rows.Count
                    $targetRange = $sourceDataSheet.Range($sourceDataSheet.Cells.Item($destRow, 1), $sourceDataSheet.Cells.Item($destRow + $targetRowCount - 1, $sourceCols))
                    $targetRange.Value2 = $dataValues
                    $destRow += $targetRowCount
                    Clear-ComObject $dataRange
                    Clear-ComObject $targetRange
                }
                Clear-ComObject $sourceRange
            } catch {
                Write-Log "Failed to copy data from '$($sourceSheet.Name)': $($_.Exception.Message)" -Level Warning
            }
        }

        Write-Log "Data consolidation complete" -Level Info

        # Release source sheet references
        foreach ($sheet in $sourceSheets) {
            Clear-ComObject $sheet
        }
        Clear-ComObject $firstValidSheet

        # --- 3. Create Pivot Table ---
        Write-Log "Creating Pivot Table..." -Level Info

        # Delete existing pivot sheet if present
        $existingPivotSheet = $null
        foreach ($sheet in $workbook.Worksheets) {
            if ($sheet.Name -eq $script:Config.PivotSheetName) {
                $existingPivotSheet = $sheet
                break
            }
        }

        if ($null -ne $existingPivotSheet) {
            Write-Log "Deleting existing '$($script:Config.PivotSheetName)' sheet..." -Level Info
            $existingPivotSheet.Delete()
            Clear-ComObject $existingPivotSheet
        }

        # Find Company sheet to place pivot after
        $companySheet = $null
        foreach ($sheet in $workbook.Worksheets) {
            if ($sheet.Name -eq $script:Config.SheetToExcludeFormatting) {
                $companySheet = $sheet
                break
            }
        }

        # Add pivot sheet
        if ($null -ne $companySheet) {
            $pivotSheet = $workbook.Worksheets.Add([System.Reflection.Missing]::Value, $companySheet)
            Write-Log "Added Pivot Table sheet after '$($companySheet.Name)'" -Level Info
        } else {
            $lastSheet2 = $workbook.Worksheets[$workbook.Worksheets.Count]
            $pivotSheet = $workbook.Worksheets.Add([System.Reflection.Missing]::Value, $lastSheet2)
            Write-Log "Added Pivot Table sheet at end" -Level Info
            Clear-ComObject $lastSheet2
        }

        $pivotSheet.Name = $script:Config.PivotSheetName
        $pivotSheet.Tab.ColorIndex = 6  # Yellow

        Clear-ComObject $companySheet

        # Create pivot table
        $pivotSourceRange = $sourceDataSheet.UsedRange
        $sourceRowCount = $pivotSourceRange.Rows.Count
        Write-Log "Source Data has $sourceRowCount rows" -Level Info

        if ($sourceRowCount -le 1) {
            Write-Log "No data rows for Pivot Table" -Level Warning
        } else {
            $xlRowField = 1
            $xlDataField = 4
            $xlMax = -4136
            $xlSum = -4157

            $pivotCache = $workbook.PivotCaches().Create(1, $pivotSourceRange)
            $pivotTable = $pivotCache.CreatePivotTable($pivotSheet.Range("A3"), "VulnPivotTable")
            Write-Log "Pivot Table object created" -Level Info

            # Configure pivot table fields
            $rowFieldsToAdd = @("Remediation Type", "Product", "Host Name", "Fix", "IP", "Evidence Path", "Evidence Version")

            Write-Log "Configuring Pivot Table fields..." -Level Info
            foreach ($fieldName in $rowFieldsToAdd) {
                try {
                    $ptField = $pivotTable.PivotFields($fieldName)
                    $ptField.Orientation = $xlRowField
                    Clear-ComObject $ptField
                } catch {
                    Write-Log "Could not add row field '$fieldName'" -Level Warning
                }
            }

            # Add value fields
            $dataField1 = $null
            $dataField2 = $null

            try {
                $sourceField1 = $pivotTable.PivotFields("EPSS Score")
                $dataField1 = $pivotTable.AddDataField($sourceField1)
                $dataField1.Function = $xlMax
                $dataField1.Name = "Max EPSS Score"
                Clear-ComObject $sourceField1
                Write-Log "Added Max EPSS Score field" -Level Info
            } catch {
                Write-Log "Could not add EPSS Score field" -Level Warning
            }

            try {
                $sourceField2 = $pivotTable.PivotFields("Vulnerability Count")
                $dataField2 = $pivotTable.AddDataField($sourceField2)
                $dataField2.Function = $xlSum
                $dataField2.Name = "Total Vulnerability Count"
                Clear-ComObject $sourceField2
                Write-Log "Added Total Vulnerability Count field" -Level Info
            } catch {
                Write-Log "Could not add Vulnerability Count field" -Level Warning
            }

            # Apply conditional formatting
            if ($null -ne $dataField1) {
                try {
                    $cfRange = $dataField1.DataRange
                    $cfRange.FormatConditions.Delete()
                    $cfCondition = $cfRange.FormatConditions.Add(1, 5, "$($script:Config.ConditionalFormatThreshold)")
                    $cfCondition.Interior.ColorIndex = 3  # Red
                    Write-Log "Applied conditional formatting (EPSS > $($script:Config.ConditionalFormatThreshold))" -Level Info
                    Clear-ComObject $cfCondition
                    Clear-ComObject $cfRange
                } catch {
                    Write-Log "Could not apply conditional formatting" -Level Warning
                }
            }

            # Add color key
            try {
                $keyStartCol = $pivotTable.TableRange2.Column + $pivotTable.TableRange2.Columns.Count + 1
                $keyStartRow = $pivotTable.TableRange1.Row

                $headerCell = $pivotSheet.Cells.Item($keyStartRow, $keyStartCol)
                $headerCell.Value2 = "Key"
                $headerCell.Font.Bold = $true
                Clear-ComObject $headerCell

                $keyCurrentRow = $keyStartRow + 1
                $keyData = @(
                    @{Text = "Do not touch"; BgColorIndex = 3; FontColorIndex = 2; Strikethrough = $false}
                    @{Text = "No action needed - auto updates"; BgColorIndex = 4; FontColorIndex = 2; Strikethrough = $false}
                    @{Text = "Update or patch"; BgColorIndex = 5; FontColorIndex = 2; Strikethrough = $false}
                    @{Text = "Uninstall"; BgColorIndex = 16; FontColorIndex = 1; Strikethrough = $false}
                    @{Text = "Already Remediated"; BgColorIndex = 2; FontColorIndex = 1; Strikethrough = $true}
                    @{Text = "Configuration change needed and further investigation"; BgColorIndex = 6; FontColorIndex = 1; Strikethrough = $false}
                )

                foreach ($item in $keyData) {
                    $keyCell = $pivotSheet.Cells.Item($keyCurrentRow, $keyStartCol)
                    $keyCell.Value2 = $item.Text
                    $keyCell.Interior.ColorIndex = $item.BgColorIndex
                    $keyCell.Font.ColorIndex = $item.FontColorIndex
                    $keyCell.Font.Strikethrough = $item.Strikethrough
                    Clear-ComObject $keyCell
                    $keyCurrentRow++
                }

                $keyColumn = $pivotSheet.Columns.Item($keyStartCol)
                $keyColumn.AutoFit() | Out-Null
                Clear-ComObject $keyColumn

                Write-Log "Color key added" -Level Info
            } catch {
                Write-Log "Could not add color key: $($_.Exception.Message)" -Level Warning
            }

            # Resize Column A
            try {
                $colA = $pivotSheet.Columns("A")
                $colA.ColumnWidth = 50
                Clear-ComObject $colA
                Write-Log "Column A resized to width 50" -Level Info
            } catch {
                Write-Log "Could not resize Column A" -Level Warning
            }

            Clear-ComObject $dataField1
            Clear-ComObject $dataField2
            Clear-ComObject $pivotTable
            Clear-ComObject $pivotCache
        }

        Clear-ComObject $pivotSourceRange

        # --- 4. Save and Close ---
        Write-Log "Saving workbook to: $OutputPath" -Level Info

        # Delete existing file if present
        if (Test-Path $OutputPath) {
            try {
                Remove-Item -Path $OutputPath -Force -ErrorAction Stop
                Write-Log "Removed existing file: $OutputPath" -Level Info
            } catch {
                throw "Cannot overwrite existing file '$OutputPath'. Please close it if it's open and try again."
            }
        }

        $workbook.SaveAs($OutputPath)
        $workbook.Close($false)

        Write-Log "Excel report generation complete" -Level Success

    } catch {
        Write-Log "Excel report generation failed: $($_.Exception.Message)" -Level Error
        throw $_
    } finally {
        # Cleanup COM objects
        Clear-ComObject $pivotSheet
        Clear-ComObject $sourceDataSheet
        Clear-ComObject $workbook

        if ($null -ne $excel) {
            try {
                $excel.Quit()
            } catch {
                Write-Log "Excel quit failed" -Level Warning
            }
            Clear-ComObject $excel
        }

        [System.GC]::Collect()
        [System.GC]::WaitForPendingFinalizers()
    }
}

# --- GUI Functions ---

function Show-VScanMagicGUI {
    # Initialize script-level variables for output file paths
    $script:WordReportPath = $null
    $script:ExcelReportPath = $null

    # Create main form
    $form = New-Object System.Windows.Forms.Form
    $form.Text = "$($script:Config.AppName) - Vulnerability Report Generator"
    $form.Size = New-Object System.Drawing.Size(700, 620)
    $form.StartPosition = "CenterScreen"
    $form.FormBorderStyle = "FixedDialog"
    $form.MaximizeBox = $false

    # --- Input File Section ---
    $labelInputFile = New-Object System.Windows.Forms.Label
    $labelInputFile.Location = New-Object System.Drawing.Point(20, 20)
    $labelInputFile.Size = New-Object System.Drawing.Size(150, 20)
    $labelInputFile.Text = "Input XLSX File:"
    $form.Controls.Add($labelInputFile)

    $textBoxInputFile = New-Object System.Windows.Forms.TextBox
    $textBoxInputFile.Location = New-Object System.Drawing.Point(20, 45)
    $textBoxInputFile.Size = New-Object System.Drawing.Size(520, 20)
    $textBoxInputFile.ReadOnly = $true
    $form.Controls.Add($textBoxInputFile)

    $buttonBrowseInput = New-Object System.Windows.Forms.Button
    $buttonBrowseInput.Location = New-Object System.Drawing.Point(550, 43)
    $buttonBrowseInput.Size = New-Object System.Drawing.Size(100, 25)
    $buttonBrowseInput.Text = "Browse..."
    $buttonBrowseInput.Add_Click({
        $openFileDialog = New-Object System.Windows.Forms.OpenFileDialog
        $openFileDialog.Filter = "Excel Files (*.xlsx)|*.xlsx|All Files (*.*)|*.*"
        $openFileDialog.Title = "Select Input Vulnerability Scan File"

        if ($openFileDialog.ShowDialog() -eq [System.Windows.Forms.DialogResult]::OK) {
            $textBoxInputFile.Text = $openFileDialog.FileName

            # Automatically set output directory to input file's directory
            $inputDirectory = [System.IO.Path]::GetDirectoryName($openFileDialog.FileName)
            $textBoxOutputDir.Text = $inputDirectory
        }
    })
    $form.Controls.Add($buttonBrowseInput)

    # --- Client Name ---
    $labelClientName = New-Object System.Windows.Forms.Label
    $labelClientName.Location = New-Object System.Drawing.Point(20, 85)
    $labelClientName.Size = New-Object System.Drawing.Size(150, 20)
    $labelClientName.Text = "Client Name:"
    $form.Controls.Add($labelClientName)

    $textBoxClientName = New-Object System.Windows.Forms.TextBox
    $textBoxClientName.Location = New-Object System.Drawing.Point(20, 110)
    $textBoxClientName.Size = New-Object System.Drawing.Size(300, 20)
    $form.Controls.Add($textBoxClientName)

    # --- Scan Date ---
    $labelScanDate = New-Object System.Windows.Forms.Label
    $labelScanDate.Location = New-Object System.Drawing.Point(350, 85)
    $labelScanDate.Size = New-Object System.Drawing.Size(150, 20)
    $labelScanDate.Text = "Scan Date:"
    $form.Controls.Add($labelScanDate)

    $datePickerScanDate = New-Object System.Windows.Forms.DateTimePicker
    $datePickerScanDate.Location = New-Object System.Drawing.Point(350, 110)
    $datePickerScanDate.Size = New-Object System.Drawing.Size(200, 20)
    $datePickerScanDate.Format = [System.Windows.Forms.DateTimePickerFormat]::Short
    $form.Controls.Add($datePickerScanDate)

    # --- Output Options ---
    $groupBoxOutput = New-Object System.Windows.Forms.GroupBox
    $groupBoxOutput.Location = New-Object System.Drawing.Point(20, 145)
    $groupBoxOutput.Size = New-Object System.Drawing.Size(630, 80)
    $groupBoxOutput.Text = "Output Options"
    $form.Controls.Add($groupBoxOutput)

    $checkBoxExcel = New-Object System.Windows.Forms.CheckBox
    $checkBoxExcel.Location = New-Object System.Drawing.Point(20, 25)
    $checkBoxExcel.Size = New-Object System.Drawing.Size(200, 20)
    $checkBoxExcel.Text = "Generate Excel Report"
    $checkBoxExcel.Checked = $true
    $groupBoxOutput.Controls.Add($checkBoxExcel)

    $checkBoxWord = New-Object System.Windows.Forms.CheckBox
    $checkBoxWord.Location = New-Object System.Drawing.Point(20, 50)
    $checkBoxWord.Size = New-Object System.Drawing.Size(200, 20)
    $checkBoxWord.Text = "Generate Word Report"
    $checkBoxWord.Checked = $true
    $groupBoxOutput.Controls.Add($checkBoxWord)

    # --- Output Directory ---
    $labelOutputDir = New-Object System.Windows.Forms.Label
    $labelOutputDir.Location = New-Object System.Drawing.Point(20, 240)
    $labelOutputDir.Size = New-Object System.Drawing.Size(150, 20)
    $labelOutputDir.Text = "Output Directory:"
    $form.Controls.Add($labelOutputDir)

    $textBoxOutputDir = New-Object System.Windows.Forms.TextBox
    $textBoxOutputDir.Location = New-Object System.Drawing.Point(20, 265)
    $textBoxOutputDir.Size = New-Object System.Drawing.Size(520, 20)
    $textBoxOutputDir.Text = [Environment]::GetFolderPath("Desktop")
    $form.Controls.Add($textBoxOutputDir)

    $buttonBrowseOutput = New-Object System.Windows.Forms.Button
    $buttonBrowseOutput.Location = New-Object System.Drawing.Point(550, 263)
    $buttonBrowseOutput.Size = New-Object System.Drawing.Size(100, 25)
    $buttonBrowseOutput.Text = "Browse..."
    $buttonBrowseOutput.Add_Click({
        $folderBrowser = New-Object System.Windows.Forms.FolderBrowserDialog
        $folderBrowser.Description = "Select Output Directory"
        $folderBrowser.SelectedPath = $textBoxOutputDir.Text

        if ($folderBrowser.ShowDialog() -eq [System.Windows.Forms.DialogResult]::OK) {
            $textBoxOutputDir.Text = $folderBrowser.SelectedPath
        }
    })
    $form.Controls.Add($buttonBrowseOutput)

    # --- Log Section ---
    $labelLog = New-Object System.Windows.Forms.Label
    $labelLog.Location = New-Object System.Drawing.Point(20, 305)
    $labelLog.Size = New-Object System.Drawing.Size(150, 20)
    $labelLog.Text = "Processing Log:"
    $form.Controls.Add($labelLog)

    $script:LogTextBox = New-Object System.Windows.Forms.TextBox
    $script:LogTextBox.Location = New-Object System.Drawing.Point(20, 330)
    $script:LogTextBox.Size = New-Object System.Drawing.Size(630, 160)
    $script:LogTextBox.Multiline = $true
    $script:LogTextBox.ScrollBars = "Vertical"
    $script:LogTextBox.ReadOnly = $true
    $script:LogTextBox.Font = New-Object System.Drawing.Font("Consolas", 9)
    $form.Controls.Add($script:LogTextBox)

    # --- Open Report Buttons ---
    $script:buttonOpenWord = New-Object System.Windows.Forms.Button
    $script:buttonOpenWord.Location = New-Object System.Drawing.Point(20, 500)
    $script:buttonOpenWord.Size = New-Object System.Drawing.Size(150, 25)
    $script:buttonOpenWord.Text = "Open Word Report"
    $script:buttonOpenWord.Enabled = $false
    $script:buttonOpenWord.Add_Click({
        if ($script:WordReportPath -and (Test-Path $script:WordReportPath)) {
            Start-Process $script:WordReportPath
        }
    })
    $form.Controls.Add($script:buttonOpenWord)

    $script:buttonOpenExcel = New-Object System.Windows.Forms.Button
    $script:buttonOpenExcel.Location = New-Object System.Drawing.Point(180, 500)
    $script:buttonOpenExcel.Size = New-Object System.Drawing.Size(150, 25)
    $script:buttonOpenExcel.Text = "Open Excel Report"
    $script:buttonOpenExcel.Enabled = $false
    $script:buttonOpenExcel.Add_Click({
        if ($script:ExcelReportPath -and (Test-Path $script:ExcelReportPath)) {
            Start-Process $script:ExcelReportPath
        }
    })
    $form.Controls.Add($script:buttonOpenExcel)

    # --- Action Buttons ---
    $buttonGenerate = New-Object System.Windows.Forms.Button
    $buttonGenerate.Location = New-Object System.Drawing.Point(450, 535)
    $buttonGenerate.Size = New-Object System.Drawing.Size(100, 30)
    $buttonGenerate.Text = "Generate"
    $buttonGenerate.Add_Click({
        # Validate inputs
        if ([string]::IsNullOrWhiteSpace($textBoxInputFile.Text)) {
            [System.Windows.Forms.MessageBox]::Show("Please select an input file.", "Validation Error",
                [System.Windows.Forms.MessageBoxButtons]::OK, [System.Windows.Forms.MessageBoxIcon]::Warning)
            return
        }

        if ([string]::IsNullOrWhiteSpace($textBoxClientName.Text)) {
            [System.Windows.Forms.MessageBox]::Show("Please enter a client name.", "Validation Error",
                [System.Windows.Forms.MessageBoxButtons]::OK, [System.Windows.Forms.MessageBoxIcon]::Warning)
            return
        }

        if (-not $checkBoxExcel.Checked -and -not $checkBoxWord.Checked) {
            [System.Windows.Forms.MessageBox]::Show("Please select at least one output option.", "Validation Error",
                [System.Windows.Forms.MessageBoxButtons]::OK, [System.Windows.Forms.MessageBoxIcon]::Warning)
            return
        }

        # Check if output files are locked before starting
        $baseFileName = [System.IO.Path]::GetFileNameWithoutExtension($textBoxInputFile.Text)
        $lockedFiles = @()

        if ($checkBoxWord.Checked) {
            $wordOutputPath = Join-Path $textBoxOutputDir.Text "$baseFileName`_Report.docx"
            if (Test-FileLocked -FilePath $wordOutputPath) {
                $lockedFiles += "Word Report: $wordOutputPath"
            }
        }

        if ($checkBoxExcel.Checked) {
            $excelOutputPath = Join-Path $textBoxOutputDir.Text "$baseFileName`_Processed.xlsx"
            if (Test-FileLocked -FilePath $excelOutputPath) {
                $lockedFiles += "Excel Report: $excelOutputPath"
            }
        }

        if ($lockedFiles.Count -gt 0) {
            $lockedFilesList = $lockedFiles -join "`n"
            [System.Windows.Forms.MessageBox]::Show(
                "The following output file(s) are currently open or locked:`n`n$lockedFilesList`n`nPlease close these files and try again.",
                "Files Are Locked",
                [System.Windows.Forms.MessageBoxButtons]::OK,
                [System.Windows.Forms.MessageBoxIcon]::Warning)
            return
        }

        # Disable button during processing
        $buttonGenerate.Enabled = $false
        $script:LogTextBox.Clear()

        # Disable open buttons at start
        $script:buttonOpenWord.Enabled = $false
        $script:buttonOpenExcel.Enabled = $false

        try {
            Write-Log "=== Starting VScanMagic Processing ===" -Level Info
            Write-Log "Input File: $($textBoxInputFile.Text)"
            Write-Log "Client: $($textBoxClientName.Text)"
            Write-Log "Scan Date: $($datePickerScanDate.Value.ToShortDateString())"

            # Read vulnerability data from all remediation sheets
            $vulnData = Get-VulnerabilityData -ExcelPath $textBoxInputFile.Text

            # Calculate top 10 vulnerabilities
            $top10 = Get-Top10Vulnerabilities -VulnData $vulnData

            # Generate Word report
            if ($checkBoxWord.Checked) {
                $baseFileName = [System.IO.Path]::GetFileNameWithoutExtension($textBoxInputFile.Text)
                $wordOutputPath = Join-Path $textBoxOutputDir.Text "$baseFileName`_Report.docx"

                New-WordReport -OutputPath $wordOutputPath `
                              -ClientName $textBoxClientName.Text `
                              -ScanDate $datePickerScanDate.Value.ToShortDateString() `
                              -Top10Data $top10

                # Store path and enable open button
                $script:WordReportPath = $wordOutputPath
                $script:buttonOpenWord.Enabled = $true

                Write-Log "Word report saved to: $wordOutputPath" -Level Success
            }

            # Generate Excel report
            if ($checkBoxExcel.Checked) {
                $baseFileName = [System.IO.Path]::GetFileNameWithoutExtension($textBoxInputFile.Text)
                $excelOutputPath = Join-Path $textBoxOutputDir.Text "$baseFileName`_Processed.xlsx"

                New-ExcelReport -InputPath $textBoxInputFile.Text -OutputPath $excelOutputPath

                # Store path and enable open button
                $script:ExcelReportPath = $excelOutputPath
                $script:buttonOpenExcel.Enabled = $true

                Write-Log "Excel report saved to: $excelOutputPath" -Level Success
            }

            Write-Log "=== Processing Complete ===" -Level Success

            [System.Windows.Forms.MessageBox]::Show("Report generation completed successfully!", "Success",
                [System.Windows.Forms.MessageBoxButtons]::OK, [System.Windows.Forms.MessageBoxIcon]::Information)

        } catch {
            Write-Log "Processing failed: $($_.Exception.Message)" -Level Error
            [System.Windows.Forms.MessageBox]::Show("An error occurred during processing. Check the log for details.", "Error",
                [System.Windows.Forms.MessageBoxButtons]::OK, [System.Windows.Forms.MessageBoxIcon]::Error)
        } finally {
            $buttonGenerate.Enabled = $true
        }
    })
    $form.Controls.Add($buttonGenerate)

    $buttonClose = New-Object System.Windows.Forms.Button
    $buttonClose.Location = New-Object System.Drawing.Point(560, 535)
    $buttonClose.Size = New-Object System.Drawing.Size(90, 30)
    $buttonClose.Text = "Close"
    $buttonClose.Add_Click({ $form.Close() })
    $form.Controls.Add($buttonClose)

    # Show form
    Write-Log "VScanMagic v3 initialized" -Level Info
    $form.ShowDialog() | Out-Null
}

# --- Main Execution ---
Show-VScanMagicGUI
