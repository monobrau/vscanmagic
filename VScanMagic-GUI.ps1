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

    # Color Thresholds for Risk Scores
    RiskColors = @{
        Critical = @{ Threshold = 7500; Color = 'DC143C'; Name = 'Crimson'; TextColor = 'FFFFFF' }
        High = @{ Threshold = 1000; Color = 'FF8C00'; Name = 'Dark Orange'; TextColor = 'FFFFFF' }
        Medium = @{ Threshold = 0; Color = 'FFFF00'; Name = 'Yellow'; TextColor = '000000' }
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
    ConsolidatedSheetName = "Source Data"
    PivotSheetName = "Proposed Remediations (all)"

    # Excel Path Limit
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

function Get-RiskScoreColor {
    param([double]$RiskScore)

    if ($RiskScore -ge $script:Config.RiskColors.Critical.Threshold) {
        return $script:Config.RiskColors.Critical
    } elseif ($RiskScore -ge $script:Config.RiskColors.High.Threshold) {
        return $script:Config.RiskColors.High
    } else {
        return $script:Config.RiskColors.Medium
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

function Get-VulnerabilityData {
    param(
        [string]$ExcelPath,
        [string]$SheetName = "Source Data"
    )

    Write-Log "Reading vulnerability data from Excel..."

    $excel = $null
    $workbook = $null
    $worksheet = $null
    $data = @()

    try {
        $excel = New-Object -ComObject Excel.Application
        $excel.Visible = $false
        $excel.DisplayAlerts = $false

        $workbook = $excel.Workbooks.Open($ExcelPath)

        # Try to find the sheet - be flexible with sheet name
        $worksheet = $workbook.Worksheets | Where-Object { $_.Name -eq $SheetName }

        if (-not $worksheet) {
            # Try case-insensitive match
            $worksheet = $workbook.Worksheets | Where-Object { $_.Name.ToLower() -eq $SheetName.ToLower() }
        }

        if (-not $worksheet) {
            # List available sheets
            $availableSheets = ($workbook.Worksheets | ForEach-Object { $_.Name }) -join ", "
            Write-Log "Available sheets: $availableSheets" -Level Warning
            throw "Sheet '$SheetName' not found. Available sheets: $availableSheets"
        }

        Write-Log "Found sheet: $($worksheet.Name)"

        $usedRange = $worksheet.UsedRange
        $rowCount = $usedRange.Rows.Count
        $colCount = $usedRange.Columns.Count

        Write-Log "Sheet has $rowCount rows and $colCount columns"

        # Get headers with flexible matching
        $headers = @{}
        for ($col = 1; $col -le $colCount; $col++) {
            $headerName = $worksheet.Cells.Item(1, $col).Text
            if ($headerName) {
                $headers[$headerName] = $col
            }
        }

        Write-Log "Found headers: $($headers.Keys -join ', ')"

        # Define flexible column mappings with multiple possible names
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

        Write-Log "Successfully mapped $($columnIndices.Count) columns. Reading data rows..."

        # Read data rows with flexible parsing
        for ($row = 2; $row -le $rowCount; $row++) {
            # Show progress for large datasets
            if ($row % 100 -eq 0) {
                Write-Log "Processing row $row of $rowCount..."
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
                $rowData['Host Name'] = $worksheet.Cells.Item($row, $columnIndices['HostName']).Text
            }

            # Read IP
            if ($columnIndices.ContainsKey('IP')) {
                $rowData['IP'] = $worksheet.Cells.Item($row, $columnIndices['IP']).Text
            }

            # Read Product (required)
            if ($columnIndices.ContainsKey('Product')) {
                $rowData['Product'] = $worksheet.Cells.Item($row, $columnIndices['Product']).Text
            }

            # Skip rows with no product name
            if ([string]::IsNullOrWhiteSpace($rowData['Product'])) {
                continue
            }

            # Read severity counts
            if ($columnIndices.ContainsKey('Critical')) {
                $rowData['Critical'] = Get-SafeNumericValue -Value $worksheet.Cells.Item($row, $columnIndices['Critical']).Text
            }

            if ($columnIndices.ContainsKey('High')) {
                $rowData['High'] = Get-SafeNumericValue -Value $worksheet.Cells.Item($row, $columnIndices['High']).Text
            }

            if ($columnIndices.ContainsKey('Medium')) {
                $rowData['Medium'] = Get-SafeNumericValue -Value $worksheet.Cells.Item($row, $columnIndices['Medium']).Text
            }

            if ($columnIndices.ContainsKey('Low')) {
                $rowData['Low'] = Get-SafeNumericValue -Value $worksheet.Cells.Item($row, $columnIndices['Low']).Text
            }

            # Read Vulnerability Count
            if ($columnIndices.ContainsKey('VulnCount')) {
                $rowData['Vulnerability Count'] = Get-SafeNumericValue -Value $worksheet.Cells.Item($row, $columnIndices['VulnCount']).Text
            } else {
                # Calculate from severity counts if not provided
                $rowData['Vulnerability Count'] = $rowData['Critical'] + $rowData['High'] + $rowData['Medium'] + $rowData['Low']
            }

            # Read EPSS Score
            if ($columnIndices.ContainsKey('EPSS')) {
                $rowData['EPSS Score'] = Get-SafeDoubleValue -Value $worksheet.Cells.Item($row, $columnIndices['EPSS']).Text
            }

            # Only add rows that have at least one vulnerability
            if ($rowData['Vulnerability Count'] -gt 0) {
                $data += [PSCustomObject]$rowData
            }
        }

        Write-Log "Successfully read $($data.Count) vulnerability records" -Level Success

        return $data

    } catch {
        Write-Log "Error reading Excel data: $($_.Exception.Message)" -Level Error
        throw
    } finally {
        if ($worksheet) { Clear-ComObject $worksheet }
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

        # --- Title Page ---
        Write-Log "Creating title page..."
        $selection = $word.Selection

        $selection.Font.Size = 28
        $selection.Font.Bold = $true
        $selection.ParagraphFormat.Alignment = 1  # Center
        $selection.TypeText("Vulnerability Assessment Report")
        $selection.TypeParagraph()
        $selection.TypeParagraph()

        $selection.Font.Size = 18
        $selection.Font.Bold = $false
        $selection.TypeText($ClientName)
        $selection.TypeParagraph()
        $selection.TypeParagraph()

        $selection.Font.Size = 14
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
        $selection.Font.Size = 16
        $selection.Font.Bold = $true
        $selection.TypeText("Executive Summary")
        $selection.TypeParagraph()
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
        $selection.Font.Size = 16
        $selection.Font.Bold = $true
        $selection.TypeText("Scoring Methodology")
        $selection.TypeParagraph()
        $selection.TypeParagraph()

        $selection.Font.Size = 11
        $selection.Font.Bold = $false
        $selection.TypeText("The Composite Risk Score is calculated using the following formula:")
        $selection.TypeParagraph()
        $selection.TypeParagraph()

        $selection.Font.Name = "Courier New"
        $selection.Font.Bold = $true
        $selection.TypeText("Risk Score = Vulnerability Count × EPSS Score × Average CVSS Equivalent")
        $selection.TypeParagraph()
        $selection.TypeParagraph()

        $selection.Font.Name = "Calibri"
        $selection.Font.Bold = $false
        $selection.TypeText("Where Average CVSS is calculated as:")
        $selection.TypeParagraph()
        $selection.TypeParagraph()

        $selection.Font.Name = "Courier New"
        $selection.TypeText("(Critical × 9.0 + High × 7.0 + Medium × 5.0 + Low × 3.0) / Total Vulnerabilities")
        $selection.TypeParagraph()
        $selection.TypeParagraph()

        # --- Risk Score Color Legend ---
        Write-Log "Creating color legend..."
        $selection.Font.Name = "Calibri"
        $selection.Font.Size = 14
        $selection.Font.Bold = $true
        $selection.TypeText("Risk Score Color Legend")
        $selection.TypeParagraph()
        $selection.TypeParagraph()

        # Create legend table
        $legendTable = $doc.Tables.Add($selection.Range, 3, 2)
        $legendTable.Borders.Enable = $true

        $legendTable.Cell(1, 1).Range.Text = "Critical"
        $legendTable.Cell(1, 1).Shading.BackgroundPatternColor = ConvertTo-HexColor -HexColor $script:Config.RiskColors.Critical.Color
        $legendTable.Cell(1, 1).Range.Font.Color = ConvertTo-HexColor -HexColor $script:Config.RiskColors.Critical.TextColor
        $legendTable.Cell(1, 2).Range.Text = "Risk Score ≥ $($script:Config.RiskColors.Critical.Threshold)"

        $legendTable.Cell(2, 1).Range.Text = "High"
        $legendTable.Cell(2, 1).Shading.BackgroundPatternColor = ConvertTo-HexColor -HexColor $script:Config.RiskColors.High.Color
        $legendTable.Cell(2, 1).Range.Font.Color = ConvertTo-HexColor -HexColor $script:Config.RiskColors.High.TextColor
        $legendTable.Cell(2, 2).Range.Text = "Risk Score ≥ $($script:Config.RiskColors.High.Threshold)"

        $legendTable.Cell(3, 1).Range.Text = "Medium"
        $legendTable.Cell(3, 1).Shading.BackgroundPatternColor = ConvertTo-HexColor -HexColor $script:Config.RiskColors.Medium.Color
        $legendTable.Cell(3, 1).Range.Font.Color = ConvertTo-HexColor -HexColor $script:Config.RiskColors.Medium.TextColor
        $legendTable.Cell(3, 2).Range.Text = "Risk Score > $($script:Config.RiskColors.Medium.Threshold)"

        $selection.EndKey(6)  # Move to end of document
        $selection.TypeParagraph()
        $selection.TypeParagraph()

        # --- Top 10 Vulnerabilities Table ---
        Write-Log "Creating top 10 vulnerabilities table..."
        $selection.Font.Size = 16
        $selection.Font.Bold = $true
        $selection.TypeText("Top 10 Vulnerabilities by Risk Score")
        $selection.TypeParagraph()
        $selection.TypeParagraph()

        # Create table
        $table = $doc.Tables.Add($selection.Range, ($Top10Data.Count + 1), 7)
        $table.Borders.Enable = $true
        $table.Style = "Grid Table 4 - Accent 1"

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

            $table.Cell($rowIndex, 1).Range.Text = $rank
            $table.Cell($rowIndex, 2).Range.Text = $item.Product
            $table.Cell($rowIndex, 3).Range.Text = $item.RiskScore.ToString("N2")
            $table.Cell($rowIndex, 4).Range.Text = $item.EPSSScore.ToString("N4")
            $table.Cell($rowIndex, 5).Range.Text = $item.AvgCVSS.ToString("N2")
            $table.Cell($rowIndex, 6).Range.Text = $item.VulnCount
            $table.Cell($rowIndex, 7).Range.Text = $item.AffectedSystems.Count

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

        $selection.EndKey(6)
        $selection.TypeParagraph()
        $selection.TypeParagraph()

        # --- Detailed Findings ---
        Write-Log "Creating detailed findings section..."
        $selection.Font.Size = 16
        $selection.Font.Bold = $true
        $selection.Font.Color = 0  # Black
        $selection.TypeText("Detailed Findings and Remediation Guidance")
        $selection.TypeParagraph()
        $selection.TypeParagraph()

        $rank = 1
        foreach ($item in $Top10Data) {
            $selection.Font.Size = 14
            $selection.Font.Bold = $true
            $selection.TypeText("$rank. $($item.Product)")
            $selection.TypeParagraph()

            $selection.Font.Size = 11
            $selection.Font.Bold = $false

            # Risk metrics
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
            $selection.TypeParagraph()

            # Affected systems list
            $selection.Font.Bold = $true
            $selection.TypeText("Affected Systems:")
            $selection.TypeParagraph()
            $selection.Font.Bold = $false

            foreach ($system in ($item.AffectedSystems | Select-Object -Unique)) {
                $selection.TypeText("  • $system")
                $selection.TypeParagraph()
            }
            $selection.TypeParagraph()

            # Remediation guidance
            $selection.Font.Bold = $true
            $selection.TypeText("Remediation Guidance:")
            $selection.TypeParagraph()
            $selection.Font.Bold = $false

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

            $selection.TypeParagraph()
            $selection.TypeParagraph()

            $rank++
        }

        # Save document
        Write-Log "Saving document to: $OutputPath"
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

# --- GUI Functions ---

function Show-VScanMagicGUI {
    # Create main form
    $form = New-Object System.Windows.Forms.Form
    $form.Text = "$($script:Config.AppName) - Vulnerability Report Generator"
    $form.Size = New-Object System.Drawing.Size(700, 640)
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

            # Populate sheet selector
            $comboBoxSheet.Items.Clear()
            try {
                $excel = New-Object -ComObject Excel.Application
                $excel.Visible = $false
                $excel.DisplayAlerts = $false
                $workbook = $excel.Workbooks.Open($openFileDialog.FileName)

                foreach ($sheet in $workbook.Worksheets) {
                    $comboBoxSheet.Items.Add($sheet.Name) | Out-Null
                }

                # Auto-select "Source Data" if it exists
                $sourceDataIndex = $comboBoxSheet.Items.IndexOf("Source Data")
                if ($sourceDataIndex -ge 0) {
                    $comboBoxSheet.SelectedIndex = $sourceDataIndex
                } else {
                    # Select first sheet by default
                    if ($comboBoxSheet.Items.Count -gt 0) {
                        $comboBoxSheet.SelectedIndex = 0
                    }
                }

                $workbook.Close($false)
                $excel.Quit()
                [System.Runtime.InteropServices.Marshal]::ReleaseComObject($workbook) | Out-Null
                [System.Runtime.InteropServices.Marshal]::ReleaseComObject($excel) | Out-Null
                [System.GC]::Collect()
                [System.GC]::WaitForPendingFinalizers()
            } catch {
                [System.Windows.Forms.MessageBox]::Show("Could not read sheets from Excel file: $($_.Exception.Message)",
                    "Warning", [System.Windows.Forms.MessageBoxButtons]::OK, [System.Windows.Forms.MessageBoxIcon]::Warning)
            }
        }
    })
    $form.Controls.Add($buttonBrowseInput)

    # --- Sheet Selector ---
    $labelSheet = New-Object System.Windows.Forms.Label
    $labelSheet.Location = New-Object System.Drawing.Point(20, 75)
    $labelSheet.Size = New-Object System.Drawing.Size(150, 20)
    $labelSheet.Text = "Source Sheet:"
    $form.Controls.Add($labelSheet)

    $comboBoxSheet = New-Object System.Windows.Forms.ComboBox
    $comboBoxSheet.Location = New-Object System.Drawing.Point(20, 100)
    $comboBoxSheet.Size = New-Object System.Drawing.Size(300, 20)
    $comboBoxSheet.DropDownStyle = [System.Windows.Forms.ComboBoxStyle]::DropDownList
    $comboBoxSheet.Items.Add("Source Data") | Out-Null
    $comboBoxSheet.SelectedIndex = 0
    $form.Controls.Add($comboBoxSheet)

    # --- Client Name ---
    $labelClientName = New-Object System.Windows.Forms.Label
    $labelClientName.Location = New-Object System.Drawing.Point(20, 135)
    $labelClientName.Size = New-Object System.Drawing.Size(150, 20)
    $labelClientName.Text = "Client Name:"
    $form.Controls.Add($labelClientName)

    $textBoxClientName = New-Object System.Windows.Forms.TextBox
    $textBoxClientName.Location = New-Object System.Drawing.Point(20, 160)
    $textBoxClientName.Size = New-Object System.Drawing.Size(300, 20)
    $form.Controls.Add($textBoxClientName)

    # --- Scan Date ---
    $labelScanDate = New-Object System.Windows.Forms.Label
    $labelScanDate.Location = New-Object System.Drawing.Point(350, 135)
    $labelScanDate.Size = New-Object System.Drawing.Size(150, 20)
    $labelScanDate.Text = "Scan Date:"
    $form.Controls.Add($labelScanDate)

    $datePickerScanDate = New-Object System.Windows.Forms.DateTimePicker
    $datePickerScanDate.Location = New-Object System.Drawing.Point(350, 160)
    $datePickerScanDate.Size = New-Object System.Drawing.Size(200, 20)
    $datePickerScanDate.Format = [System.Windows.Forms.DateTimePickerFormat]::Short
    $form.Controls.Add($datePickerScanDate)

    # --- Output Options ---
    $groupBoxOutput = New-Object System.Windows.Forms.GroupBox
    $groupBoxOutput.Location = New-Object System.Drawing.Point(20, 195)
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
    $labelOutputDir.Location = New-Object System.Drawing.Point(20, 290)
    $labelOutputDir.Size = New-Object System.Drawing.Size(150, 20)
    $labelOutputDir.Text = "Output Directory:"
    $form.Controls.Add($labelOutputDir)

    $textBoxOutputDir = New-Object System.Windows.Forms.TextBox
    $textBoxOutputDir.Location = New-Object System.Drawing.Point(20, 315)
    $textBoxOutputDir.Size = New-Object System.Drawing.Size(520, 20)
    $textBoxOutputDir.Text = [Environment]::GetFolderPath("Desktop")
    $form.Controls.Add($textBoxOutputDir)

    $buttonBrowseOutput = New-Object System.Windows.Forms.Button
    $buttonBrowseOutput.Location = New-Object System.Drawing.Point(550, 313)
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
    $labelLog.Location = New-Object System.Drawing.Point(20, 355)
    $labelLog.Size = New-Object System.Drawing.Size(150, 20)
    $labelLog.Text = "Processing Log:"
    $form.Controls.Add($labelLog)

    $script:LogTextBox = New-Object System.Windows.Forms.TextBox
    $script:LogTextBox.Location = New-Object System.Drawing.Point(20, 380)
    $script:LogTextBox.Size = New-Object System.Drawing.Size(630, 150)
    $script:LogTextBox.Multiline = $true
    $script:LogTextBox.ScrollBars = "Vertical"
    $script:LogTextBox.ReadOnly = $true
    $script:LogTextBox.Font = New-Object System.Drawing.Font("Consolas", 9)
    $form.Controls.Add($script:LogTextBox)

    # --- Action Buttons ---
    $buttonGenerate = New-Object System.Windows.Forms.Button
    $buttonGenerate.Location = New-Object System.Drawing.Point(450, 545)
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

        # Disable button during processing
        $buttonGenerate.Enabled = $false
        $script:LogTextBox.Clear()

        try {
            Write-Log "=== Starting VScanMagic Processing ===" -Level Info
            Write-Log "Input File: $($textBoxInputFile.Text)"
            Write-Log "Sheet: $($comboBoxSheet.SelectedItem)"
            Write-Log "Client: $($textBoxClientName.Text)"
            Write-Log "Scan Date: $($datePickerScanDate.Value.ToShortDateString())"

            # Read vulnerability data
            $selectedSheet = if ($comboBoxSheet.SelectedItem) { $comboBoxSheet.SelectedItem.ToString() } else { "Source Data" }
            $vulnData = Get-VulnerabilityData -ExcelPath $textBoxInputFile.Text -SheetName $selectedSheet

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

                Write-Log "Word report saved to: $wordOutputPath" -Level Success
            }

            # Generate Excel report (existing functionality would go here)
            if ($checkBoxExcel.Checked) {
                Write-Log "Excel report generation: Feature coming soon" -Level Info
                # TODO: Integrate existing Excel processing code
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
    $buttonClose.Location = New-Object System.Drawing.Point(560, 545)
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
