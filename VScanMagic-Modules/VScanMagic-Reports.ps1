# VScanMagic-Reports.ps1 - Word, Excel, Email, Time Estimate, Ticket Instructions/Notes
# Dot-sourced by VScanMagic-GUI.ps1

function Get-RiskScoreColor {
    param(
        [double]$RiskScore,
        [hashtable]$DynamicThresholds = $null
    )

    # Use dynamic thresholds if provided, otherwise use static config
    $thresholds = if ($DynamicThresholds) { $DynamicThresholds } else { $script:Config.RiskColors }

    # Validate thresholds configuration exists
    if (-not $thresholds) {
        Write-Log "ERROR: RiskColors configuration is null!" -Level Error
        return @{ Color = 'FFFF00'; TextColor = '000000'; Name = 'Unknown' }
    }

    # Check heatmap levels from highest to lowest
    if ($RiskScore -ge $thresholds.Critical.Threshold) {
        $result = $thresholds.Critical
    } elseif ($RiskScore -ge $thresholds.VeryHigh.Threshold) {
        $result = $thresholds.VeryHigh
    } elseif ($RiskScore -ge $thresholds.High.Threshold) {
        $result = $thresholds.High
    } elseif ($RiskScore -ge $thresholds.MediumHigh.Threshold) {
        $result = $thresholds.MediumHigh
    } else {
        # Everything else is Medium (yellow) - all Top 10 items need attention
        $result = $thresholds.Medium
    }

    # Validate the result has required properties
    if (-not $result) {
        Write-Log "ERROR: Get-RiskScoreColor returned null for score $RiskScore" -Level Error
        return @{ Color = 'FFFF00'; TextColor = '000000'; Name = 'Unknown' }
    }

    if (-not $result.Color -or -not $result.TextColor) {
        Write-Log "ERROR: Color object missing properties. Color=$($result.Color), TextColor=$($result.TextColor)" -Level Error
        if (-not $result.Color) { $result.Color = 'FFFF00' }
        if (-not $result.TextColor) { $result.TextColor = '000000' }
    }

    return $result
}

function ConvertTo-HexColor {
    param([string]$HexColor)

    # Convert hex string to Word color integer (BGR format)
    $r = [Convert]::ToInt32($HexColor.Substring(0, 2), 16)
    $g = [Convert]::ToInt32($HexColor.Substring(2, 2), 16)
    $b = [Convert]::ToInt32($HexColor.Substring(4, 2), 16)

    return $b * 65536 + $g * 256 + $r
}

# Word COM TypeText has a 255-character limit. Use Selection.Text for longer strings.
function Add-WordText {
    param([object]$Selection, [string]$Text)
    if ([string]::IsNullOrEmpty($Text)) { return }
    if ($Text.Length -le 255) {
        $Selection.TypeText($Text)
    } else {
        $Selection.Text = $Text
    }
}

function New-WordReport {
    param(
        [string]$OutputPath,
        [string]$ClientName,
        [string]$ScanDate,
        [array]$Top10Data,
        [array]$TimeEstimates = $null,
        [bool]$IsRMITPlus = $false,
        [array]$GeneralRecommendations = $null,
        [string]$ReportTitle = "Top Ten Vulnerabilities Report"
    )

    Write-Log "Generating Word document report..."

    # Calculate dynamic thresholds based on maximum risk score
    $maxRiskScore = ($Top10Data | Measure-Object -Property RiskScore -Maximum).Maximum
    if ($maxRiskScore -le 0) {
        $maxRiskScore = 10  # Default fallback (max theoretical: 1.0 Ã— 10.0)
    }
    Write-Log "Maximum risk score in data: $($maxRiskScore.ToString('N2'))"

    # Create proportional thresholds (as percentages of max score)
    # Yellow to Red gradient - no greens to emphasize all items need attention
    $dynamicThresholds = @{
        Critical   = @{ Threshold = $maxRiskScore * 1.00; Color = 'DC143C'; Name = 'Critical'; TextColor = 'FFFFFF'; Percent = '100%' }
        VeryHigh   = @{ Threshold = $maxRiskScore * 0.70; Color = 'FF4500'; Name = 'Very High'; TextColor = 'FFFFFF'; Percent = '70%' }
        High       = @{ Threshold = $maxRiskScore * 0.50; Color = 'FF8C00'; Name = 'High'; TextColor = 'FFFFFF'; Percent = '50%' }
        MediumHigh = @{ Threshold = $maxRiskScore * 0.30; Color = 'FFA500'; Name = 'Medium-High'; TextColor = '000000'; Percent = '30%' }
        Medium     = @{ Threshold = 0;                    Color = 'FFFF00'; Name = 'Medium'; TextColor = '000000'; Percent = '0%' }
    }
    Write-Log "Dynamic thresholds created based on max score"

    $word = $null
    $doc = $null

    try {
        Write-Log "Creating Word application..."
        $word = New-Object -ComObject Word.Application
        $word.Visible = $false
        Write-Log "Adding new document..."
        $doc = $word.Documents.Add()

        # Set document properties (optional - may fail on some systems; Word limits each to 255 chars)
        Write-Log "Setting document properties..."
        try {
            $titleProp = "$ReportTitle - $ClientName"
            if ($titleProp.Length -gt 255) { $titleProp = $titleProp.Substring(0, 255) }
            $doc.BuiltInDocumentProperties.Item("Title").Value = $titleProp
            $doc.BuiltInDocumentProperties.Item("Subject").Value = "Security Vulnerability Assessment"
            $authorVal = if ($null -eq $script:Config.Author) { "" } elseif ($script:Config.Author.Length -gt 255) { $script:Config.Author.Substring(0, 255) } else { $script:Config.Author }
            $doc.BuiltInDocumentProperties.Item("Author").Value = $authorVal
            $doc.BuiltInDocumentProperties.Item("Keywords").Value = "Vulnerability, Security, Assessment, EPSS, CVSS"
            Write-Log "Document properties set successfully"
        } catch {
            Write-Log "Warning: Could not set document properties (this is optional): $($_.Exception.Message)" -Level Warning
        }

        # Set page margins (in points: 1 inch = 72 points)
        # Default 1 inch margins for non-table pages
        Write-Log "Setting page margins..."
        $doc.PageSetup.LeftMargin = 72    # 1 inch
        $doc.PageSetup.RightMargin = 72   # 1 inch
        $doc.PageSetup.TopMargin = 72     # 1 inch
        $doc.PageSetup.BottomMargin = 72  # 1 inch

        # --- Title Page ---
        Write-Log "Creating title page..."
        $selection = $word.Selection
        Write-Log "Selection object created"

        # Add some top spacing for title page
        $selection.TypeParagraph()
        $selection.TypeParagraph()

        $selection.Font.Name = "Calibri"
        $selection.Font.Size = 32
        $selection.Font.Bold = $true
        $selection.Font.Color = 5855577  # Dark blue color
        $selection.ParagraphFormat.Alignment = 1  # Center
        $selection.TypeText($ReportTitle)
        $selection.TypeParagraph()

        # Add horizontal line
        $selection.ParagraphFormat.Borders.Item(3).LineStyle = 1  # Bottom border
        $selection.ParagraphFormat.Borders.Item(3).LineWidth = 24  # Thicker line
        $selection.ParagraphFormat.Borders.Item(3).Color = 5855577  # Dark blue
        $selection.TypeParagraph()
        $selection.ParagraphFormat.Borders.Item(3).LineStyle = 0  # Reset border
        $selection.TypeParagraph()

        $selection.Font.Size = 20
        $selection.Font.Bold = $true
        $selection.Font.Color = 0  # Black
        $selection.TypeText($ClientName)
        $selection.TypeParagraph()
        $selection.TypeParagraph()

        $selection.Font.Size = 14
        $selection.Font.Bold = $false
        $selection.TypeText("Scan Date: $ScanDate")
        $selection.TypeParagraph()

        $selection.Font.Size = 12
        $selection.TypeText("Prepared by: $($script:UserSettings.PreparedBy)")
        $selection.TypeParagraph()

        # Add company contact information if available
        if (-not [string]::IsNullOrWhiteSpace($script:UserSettings.CompanyName)) {
            $selection.TypeText($script:UserSettings.CompanyName)
            $selection.TypeParagraph()
        }
        if (-not [string]::IsNullOrWhiteSpace($script:UserSettings.CompanyAddress)) {
            $selection.TypeText($script:UserSettings.CompanyAddress)
            $selection.TypeParagraph()
        }
        if (-not [string]::IsNullOrWhiteSpace($script:UserSettings.Email)) {
            $selection.TypeText("Email: $($script:UserSettings.Email)")
            $selection.TypeParagraph()
        }
        if (-not [string]::IsNullOrWhiteSpace($script:UserSettings.PhoneNumber)) {
            $selection.TypeText("Phone: $($script:UserSettings.PhoneNumber)")
            $selection.TypeParagraph()
        }
        if (-not [string]::IsNullOrWhiteSpace($script:UserSettings.CompanyPhoneNumber)) {
            $selection.TypeText("Company Phone: $($script:UserSettings.CompanyPhoneNumber)")
            $selection.TypeParagraph()
        }

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
        $selection.TypeParagraph()

        $selection.TypeText("This report identifies the top security risks based on a ConnectSecure-aligned risk score that considers severity-weighted vulnerability counts ")
        $selection.TypeText("and EPSS (Exploit Prediction Scoring System) scores to prioritize by both impact and likelihood of exploitation. ")
        $selection.TypeText("Each finding includes specific remediation guidance appropriate for the environment.")
        $selection.TypeParagraph()

        $selection.TypeText("Please note that application vulnerabilities can be resolved either by upgrading the vulnerable application to the latest version or, ")
        $selection.TypeText("depending on the situation, by uninstalling the application if it is no longer needed.")
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

        # Add shaded box for formula
        $selection.ParagraphFormat.Shading.BackgroundPatternColor = 15329769  # Light blue/gray
        $selection.ParagraphFormat.LeftIndent = 18
        $selection.ParagraphFormat.RightIndent = 18
        $selection.ParagraphFormat.SpaceBefore = 6
        $selection.ParagraphFormat.SpaceAfter = 6

        $selection.Font.Name = "Courier New"
        $selection.Font.Bold = $true
        $selection.Font.Size = 10
        $selection.TypeText("Risk Score = (Severity Weighted Sum) x (1 + EPSS Score)")
        $selection.TypeParagraph()

        $selection.Font.Bold = $false
        $selection.TypeText("Severity Weighted Sum uses ConnectSecure Problem Category weights:")
        $selection.TypeParagraph()

        $selection.TypeText("(Critical x 0.90 + High x 0.80 + Medium x 0.50 + Low x 0.30)")
        $selection.TypeParagraph()

        $selection.TypeText("End of Life (EOL) software receives maximum weight (1.0 per vulnerability) since it no longer receives security updates.")
        $selection.TypeParagraph()

        $selection.TypeText("EPSS (Exploit Prediction Scoring System) emphasizes exploit probability - higher EPSS indicates greater likelihood of exploitation in the next 30 days.")
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

        # Validate RiskColors configuration
        if (-not $script:Config.RiskColors) {
            throw "RiskColors configuration is null or not defined"
        }

        # Create legend table with heatmap gradient (5 levels: Yellow to Red)
        Write-Log "Adding legend table..."
        $legendTable = $doc.Tables.Add($selection.Range, 5, 2)
        if (-not $legendTable) {
            throw "Failed to create legend table"
        }
        Write-Log "Legend table created successfully"
        $legendTable.Borders.Enable = $true
        $legendTable.Range.Font.Size = 10

        # Row 1: Critical (Red)
        Write-Log "Populating legend table with dynamic thresholds"
        $legendTable.Cell(1, 1).Range.Text = $dynamicThresholds.Critical.Name
        $legendTable.Cell(1, 1).Shading.BackgroundPatternColor = ConvertTo-HexColor -HexColor $dynamicThresholds.Critical.Color
        $legendTable.Cell(1, 1).Range.Font.Color = ConvertTo-HexColor -HexColor $dynamicThresholds.Critical.TextColor
        $legendTable.Cell(1, 1).Range.Font.Bold = $true
        $legendTable.Cell(1, 2).Range.Text = "Risk Score >= $($dynamicThresholds.Critical.Threshold.ToString('N2')) ($($dynamicThresholds.Critical.Percent) of max)"

        # Row 2: Very High (Orange-Red)
        $legendTable.Cell(2, 1).Range.Text = $dynamicThresholds.VeryHigh.Name
        $legendTable.Cell(2, 1).Shading.BackgroundPatternColor = ConvertTo-HexColor -HexColor $dynamicThresholds.VeryHigh.Color
        $legendTable.Cell(2, 1).Range.Font.Color = ConvertTo-HexColor -HexColor $dynamicThresholds.VeryHigh.TextColor
        $legendTable.Cell(2, 1).Range.Font.Bold = $true
        $legendTable.Cell(2, 2).Range.Text = "Risk Score >= $($dynamicThresholds.VeryHigh.Threshold.ToString('N2')) ($($dynamicThresholds.VeryHigh.Percent) of max)"

        # Row 3: High (Dark Orange)
        $legendTable.Cell(3, 1).Range.Text = $dynamicThresholds.High.Name
        $legendTable.Cell(3, 1).Shading.BackgroundPatternColor = ConvertTo-HexColor -HexColor $dynamicThresholds.High.Color
        $legendTable.Cell(3, 1).Range.Font.Color = ConvertTo-HexColor -HexColor $dynamicThresholds.High.TextColor
        $legendTable.Cell(3, 1).Range.Font.Bold = $true
        $legendTable.Cell(3, 2).Range.Text = "Risk Score >= $($dynamicThresholds.High.Threshold.ToString('N2')) ($($dynamicThresholds.High.Percent) of max)"

        # Row 4: Medium-High (Orange)
        $legendTable.Cell(4, 1).Range.Text = $dynamicThresholds.MediumHigh.Name
        $legendTable.Cell(4, 1).Shading.BackgroundPatternColor = ConvertTo-HexColor -HexColor $dynamicThresholds.MediumHigh.Color
        $legendTable.Cell(4, 1).Range.Font.Color = ConvertTo-HexColor -HexColor $dynamicThresholds.MediumHigh.TextColor
        $legendTable.Cell(4, 1).Range.Font.Bold = $true
        $legendTable.Cell(4, 2).Range.Text = "Risk Score >= $($dynamicThresholds.MediumHigh.Threshold.ToString('N2')) ($($dynamicThresholds.MediumHigh.Percent) of max)"

        # Row 5: Medium (Yellow - baseline)
        $legendTable.Cell(5, 1).Range.Text = $dynamicThresholds.Medium.Name
        $legendTable.Cell(5, 1).Shading.BackgroundPatternColor = ConvertTo-HexColor -HexColor $dynamicThresholds.Medium.Color
        $legendTable.Cell(5, 1).Range.Font.Color = ConvertTo-HexColor -HexColor $dynamicThresholds.Medium.TextColor
        $legendTable.Cell(5, 1).Range.Font.Bold = $true
        $legendTable.Cell(5, 2).Range.Text = "Risk Score >= $($dynamicThresholds.Medium.Threshold.ToString('N2')) ($($dynamicThresholds.Medium.Percent) of max)"

        # AutoFit the legend table
        $legendTable.AutoFitBehavior(1)  # 1 = wdAutoFitContent (fit to content)

        $selection.EndKey(6)  # Move to end of document

        # No page break - let content flow naturally from legend to Top 10 table
        # Set narrower margins for table page (0.5 inch)
        $currentSection = $selection.Sections.Item(1)
        $currentSection.PageSetup.LeftMargin = 36    # 0.5 inch
        $currentSection.PageSetup.RightMargin = 36   # 0.5 inch

        # --- Top 10 Vulnerabilities Table ---
        Write-Log "Creating top 10 vulnerabilities table..."
        $selection.Style = "Heading 1"
        $vulnCountText = if ($Top10Data.Count -eq 10) { "Top 10" } elseif ($Top10Data.Count -gt 0) { "Top $($Top10Data.Count)" } else { "Top Vulnerabilities" }
        $selection.TypeText("$vulnCountText Vulnerabilities by Risk Score")
        $selection.TypeParagraph()

        $selection.Style = "Normal"
        $selection.TypeParagraph()

        # Create table (8 columns when Source present for Application/Registry/Network sections)
        $hasSections = ($Top10Data | Where-Object { $_.Source -and $_.Source -ne 'Application' }).Count -gt 0
        $colCount = if ($hasSections) { 8 } else { 7 }
        $table = $doc.Tables.Add($selection.Range, ($Top10Data.Count + 1), $colCount)
        $table.Borders.Enable = $true
        $table.Style = "Grid Table 4 - Accent 1"

        # Set table font size to 9 points for better fit
        $table.Range.Font.Size = 9

        # Headers
        $headers = if ($hasSections) {
            @("Rank", "Section", "Product/System", "Risk Score", "EPSS", "Avg CVSS", "Total Vulns", "Affected Systems")
        } else {
            @("Rank", "Product/System", "Risk Score", "EPSS", "Avg CVSS", "Total Vulns", "Affected Systems")
        }
        for ($i = 0; $i -lt $headers.Count; $i++) {
            $table.Cell(1, $i + 1).Range.Text = $headers[$i]
            $table.Cell(1, $i + 1).Range.Font.Bold = $true
        }

        # Data rows (group by section for visual separation when multiple sections)
        $rank = 1
        $sectionRanks = @{}
        foreach ($item in $Top10Data) {
            $rowIndex = $rank + 1
            $sectionLabel = if ($item.Source) { $item.Source } else { 'Application' }
            if (-not $sectionRanks.ContainsKey($sectionLabel)) { $sectionRanks[$sectionLabel] = 0 }
            $sectionRanks[$sectionLabel]++
            $sectionRank = $sectionRanks[$sectionLabel]

            if ($hasSections) {
                $table.Cell($rowIndex, 1).Range.Text = $sectionRank.ToString()
                $table.Cell($rowIndex, 2).Range.Text = $sectionLabel
                $table.Cell($rowIndex, 3).Range.Text = $item.Product
                $table.Cell($rowIndex, 4).Range.Text = $item.RiskScore.ToString("N2")
                $table.Cell($rowIndex, 5).Range.Text = $item.EPSSScore.ToString("N4")
                $table.Cell($rowIndex, 6).Range.Text = $item.AvgCVSS.ToString("N2")
                $table.Cell($rowIndex, 7).Range.Text = $item.VulnCount.ToString()
                $table.Cell($rowIndex, 8).Range.Text = $item.AffectedSystems.Count.ToString()
            } else {
                $table.Cell($rowIndex, 1).Range.Text = $rank.ToString()
                $table.Cell($rowIndex, 2).Range.Text = $item.Product
                $table.Cell($rowIndex, 3).Range.Text = $item.RiskScore.ToString("N2")
                $table.Cell($rowIndex, 4).Range.Text = $item.EPSSScore.ToString("N4")
                $table.Cell($rowIndex, 5).Range.Text = $item.AvgCVSS.ToString("N2")
                $table.Cell($rowIndex, 6).Range.Text = $item.VulnCount.ToString()
                $table.Cell($rowIndex, 7).Range.Text = $item.AffectedSystems.Count.ToString()
            }

            # Apply color coding based on risk score using dynamic thresholds
            $colorInfo = Get-RiskScoreColor -RiskScore $item.RiskScore -DynamicThresholds $dynamicThresholds
            if (-not $colorInfo) {
                Write-Log "Warning: Get-RiskScoreColor returned null for score $($item.RiskScore)" -Level Warning
                # Use default color (yellow/black)
                $colorInfo = @{ Color = 'FFFF00'; TextColor = '000000' }
            }
            if (-not $colorInfo.Color) {
                Write-Log "Warning: colorInfo.Color is null for score $($item.RiskScore)" -Level Warning
                $colorInfo.Color = 'FFFF00'
            }
            if (-not $colorInfo.TextColor) {
                Write-Log "Warning: colorInfo.TextColor is null for score $($item.RiskScore)" -Level Warning
                $colorInfo.TextColor = '000000'
            }

            $bgColor = ConvertTo-HexColor -HexColor $colorInfo.Color
            $textColor = ConvertTo-HexColor -HexColor $colorInfo.TextColor

            for ($col = 1; $col -le $colCount; $col++) {
                $table.Cell($rowIndex, $col).Shading.BackgroundPatternColor = $bgColor
                $table.Cell($rowIndex, $col).Range.Font.Color = $textColor
            }

            $rank++
        }

        # Set custom column widths for better appearance
        # Column widths in points (1 inch = 72 points)
        $table.Columns(1).SetWidth(36, 0)   # Rank: 0.5 inch (narrow)
        if ($hasSections) {
            $table.Columns(2).SetWidth(72, 0)   # Section: 1 inch
            $table.Columns(3).PreferredWidthType = 3
            $table.Columns(3).PreferredWidth = 180    # Product/System
            $table.Columns(4).SetWidth(72, 0)   # Risk Score
            $table.Columns(5).SetWidth(54, 0)   # EPSS
            $table.Columns(6).SetWidth(72, 0)   # Avg CVSS
            $table.Columns(7).SetWidth(72, 0)   # Total Vulns
            $table.Columns(8).SetWidth(90, 0)   # Affected Systems
        } else {
            $table.Columns(2).PreferredWidthType = 3
            $table.Columns(2).PreferredWidth = 180    # Product/System
            $table.Columns(3).SetWidth(72, 0)   # Risk Score
            $table.Columns(4).SetWidth(54, 0)   # EPSS
            $table.Columns(5).SetWidth(72, 0)   # Avg CVSS
            $table.Columns(6).SetWidth(72, 0)   # Total Vulns
            $table.Columns(7).SetWidth(90, 0)   # Affected Systems
        }

        $selection.EndKey(6)
        $selection.TypeParagraph()
        $selection.TypeParagraph()

        # --- Vulnerability Distribution Chart ---
        Write-Log "Creating vulnerability distribution pie chart..."
        $selection.Style = "Heading 2"
        $selection.TypeText("Vulnerability Distribution by Product/System")
        $selection.TypeParagraph()

        try {
            # Create simple pie chart using older AddChart method for better compatibility
            Write-Log "Adding chart object..."
            $chartShape = $selection.InlineShapes.AddChart(5)  # 5 = xlPie
            $chart = $chartShape.Chart

            # Make pie chart area large with legend on the right
            # Set appropriate dimensions for larger pie with legend to the right
            $chartShape.Width = 540  # 7.5 inches (72 points per inch) - extra width for legend on right
            $chartShape.Height = 400  # 5.5 inches - taller to accommodate all 10 legend entries
            Write-Log "Chart dimensions set: $($chartShape.Width)w x $($chartShape.Height)h points"

            # Access chart data without calling Activate() to avoid COM disconnection
            $chartData = $chart.ChartData
            $chartWorkbook = $chartData.Workbook

            # Try to hide the Excel instance that Word creates for chart data
            try {
                $chartExcel = $chartWorkbook.Application
                $chartExcel.Visible = $false
                $chartExcel.ScreenUpdating = $false
            } catch {
                # Silently ignore if we can't hide it
            }

            $worksheet = $chartWorkbook.Worksheets(1)
            Write-Log "Chart data workbook accessed"

            # Don't clear - just overwrite the cells directly
            # Populate headers and all 10 data rows
            $worksheet.Cells.Item(1, 1) = "Product/System"
            $worksheet.Cells.Item(1, 2) = "Vulnerabilities"

            # Sort by VulnCount descending for pie chart (largest percentage first)
            # Limit to 25 items to avoid Word COM RPC failures with large datasets (150+ items)
            $maxChartItems = 25
            $sortedChartData = $Top10Data | Sort-Object -Property VulnCount -Descending
            $chartItems = if ($sortedChartData.Count -le $maxChartItems) { $sortedChartData } else { $sortedChartData | Select-Object -First $maxChartItems }
            $otherCount = if ($sortedChartData.Count -gt $maxChartItems) { ($sortedChartData | Select-Object -Skip $maxChartItems | Measure-Object -Property VulnCount -Sum).Sum } else { 0 }

            $row = 2
            foreach ($item in $chartItems) {
                $worksheet.Cells.Item($row, 1) = [string]$item.Product
                $worksheet.Cells.Item($row, 2) = [int]$item.VulnCount
                $row++
            }
            if ($otherCount -gt 0) {
                $worksheet.Cells.Item($row, 1) = "Other ($($sortedChartData.Count - $maxChartItems) more)"
                $worksheet.Cells.Item($row, 2) = [int]$otherCount
                $row++
            }
            $lastRow = $row - 1
            Write-Log "Chart data populated ($($lastRow - 1) items in rows 2-$lastRow, sorted by vulnerability count descending)"

            # Update the existing series with direct range assignment
            try {
                $series = $chart.SeriesCollection(1)
                $sheetName = $worksheet.Name

                # Use direct XValues/Values assignment (most reliable for embedded Word charts)
                $series.XValues = "='$sheetName'!`$A`$2:`$A`$$lastRow"
                $series.Values = "='$sheetName'!`$B`$2:`$B`$$lastRow"
                $series.HasDataLabels = $false
                Write-Log "Chart series configured with all $($lastRow - 1) items"
            } catch {
                Write-Log "Warning: Could not update chart series: $($_.Exception.Message)" -Level Warning
            }

            # Basic chart formatting
            $chart.HasTitle = $true
            $vulnCountText = if ($Top10Data.Count -eq 10) { "Top 10" } elseif ($Top10Data.Count -gt 0) { "Top $($Top10Data.Count)" } else { "Top Vulnerabilities" }
            $chart.ChartTitle.Text = "$vulnCountText Vulnerabilities by Count"
            $chart.HasLegend = $true
            $chart.Legend.Position = -4152  # xlLegendPositionRight - legend to the right of pie
            $chart.Legend.Font.Size = 11  # Reduced size to fit all 10 entries (1/3 smaller than 16)
            Write-Log "Chart formatting applied (legend positioned right of pie, font size 11)"

            Write-Log "Pie chart created successfully"

            # Move selection to after the chart so subsequent content appears below it
            $selection.MoveDown(5, 1)  # wdLine = 5, move 1 line down
            $selection.EndKey(6)       # wdLine = 6, move to end of line
        } catch {
            Write-Log "Warning: Could not create pie chart: $($_.Exception.Message)" -Level Warning
            # Chart creation is optional, continue with report generation
        }

        $selection.TypeParagraph()
        $selection.TypeParagraph()

        # Insert continuous section break to restore normal margins without page break
        $selection.InsertBreak(3)  # 3 = wdSectionBreakContinuous
        $detailedSection = $selection.Sections.Item(1)
        $detailedSection.PageSetup.LeftMargin = 72    # 1 inch
        $detailedSection.PageSetup.RightMargin = 72   # 1 inch

        # --- Detailed Findings ---
        Write-Log "Creating detailed findings section..."
        $selection.Style = "Heading 1"
        $selection.Font.Color = 0  # Black
        $selection.TypeText("Detailed Findings and Remediation Guidance")
        $selection.TypeParagraph()

        $selection.Style = "Normal"
        $selection.TypeParagraph()

        $timeEstimateMap = @{}
        if ($null -ne $TimeEstimates) {
            foreach ($te in $TimeEstimates) {
                if (-not [string]::IsNullOrWhiteSpace($te.Product)) { $timeEstimateMap[$te.Product] = $te }
            }
        }
        $generalRecMap = @{}
        if ($null -ne $GeneralRecommendations -and $GeneralRecommendations.Count -gt 0) {
            foreach ($rec in $GeneralRecommendations) {
                if (-not [string]::IsNullOrWhiteSpace($rec.Product)) { $generalRecMap[$rec.Product] = $rec }
            }
        }
        $rank = 1
        foreach ($item in $Top10Data) {
            # Vulnerability title with special notes
            $selection.Style = "Heading 2"
            $title = "$rank. $($item.Product)"

            $lookupKey = Get-TimeEstimateGroupKey -ProductName $item.Product
            $timeEstimate = if ($timeEstimateMap.ContainsKey($lookupKey)) { $timeEstimateMap[$lookupKey] } else { $null }

            # Add modifier text or product-type suffix (Top N report uses full modifier for client benefit: ticket generated, approval needed)
            if ($null -ne $timeEstimate -and $IsRMITPlus) {
                $afterHours = $timeEstimate.AfterHours
                $ticketGenerated = $timeEstimate.TicketGenerated
                $thirdParty = $timeEstimate.ThirdParty
                $autoTicketGenerated = $thirdParty -and $afterHours
                $isTicketGenerated = $ticketGenerated -or $autoTicketGenerated

                $modifierText = Get-ModifierText -AfterHours $afterHours -TicketGenerated $isTicketGenerated -ThirdParty $thirdParty
                if (-not [string]::IsNullOrWhiteSpace($modifierText)) {
                    $title += $modifierText
                }
                if ($afterHours) {
                    $title = "After Hours - $title"
                }
            } else {
                $title += Get-ProductTypeSuffix -ProductName $item.Product -IsRMITPlus $IsRMITPlus
            }

            Add-WordText -Selection $selection -Text $title
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

            # Display systems as comma-separated list with indent (same uniqueness as Ticket Instructions: HostName, IP, Username)
            $selection.ParagraphFormat.LeftIndent = 36
            $uniqueSystems = $item.AffectedSystems | Select-Object HostName, IP, Username, LastPingTime -Unique
            $systemsList = ($uniqueSystems | ForEach-Object {
                $systemLine = if ($_.HostName) { $_.HostName } else { $_.IP }
                if (-not [string]::IsNullOrWhiteSpace($_.Username)) { $systemLine += " ($($_.Username))" }
                if (-not [string]::IsNullOrWhiteSpace($_.IP)) { $systemLine += " - $($_.IP)" }
                if (-not [string]::IsNullOrWhiteSpace($_.LastPingTime)) { $systemLine += " (last seen $($_.LastPingTime))" }
                $systemLine
            } | Where-Object { -not [string]::IsNullOrWhiteSpace($_) }) -join ", "
            Add-WordText -Selection $selection -Text $systemsList
            $selection.TypeParagraph()
            $selection.ParagraphFormat.LeftIndent = 0
            $selection.TypeParagraph()

            # Remediation guidance
            $selection.Font.Bold = $true
            $selection.TypeText("Remediation Guidance:")
            $selection.TypeParagraph()
            $selection.Font.Bold = $false

            $selection.ParagraphFormat.LeftIndent = 36

            # Get remediation guidance from configurable rules
            $remediationText = Get-RemediationGuidance -ProductName $item.Product -OutputType 'Word'
            Add-WordText -Selection $selection -Text $remediationText

            # ConnectSecure Solution/Fix when available
            if ($item.Fix -and -not [string]::IsNullOrWhiteSpace($item.Fix)) {
                $selection.ParagraphFormat.LeftIndent = 0
                $selection.TypeParagraph()
                $selection.Font.Bold = $true
                $selection.TypeText("ConnectSecure Solution:")
                $selection.TypeParagraph()
                $selection.Font.Bold = $false
                $selection.ParagraphFormat.LeftIndent = 36
                Add-WordText -Selection $selection -Text (ConvertTo-ReadableFixText -RawFix $item.Fix)
                $selection.ParagraphFormat.LeftIndent = 0
                $selection.TypeParagraph()
            }

            $selection.ParagraphFormat.LeftIndent = 0  # Reset indent
            $selection.TypeParagraph()

            $matchingRec = if ($generalRecMap.ContainsKey($item.Product)) { $generalRecMap[$item.Product] } else { $null }
            if ($null -ne $matchingRec -and -not [string]::IsNullOrWhiteSpace($matchingRec.Recommendations)) {
                $selection.Font.Bold = $true
                $selection.TypeText("General Recommendations:")
                $selection.TypeParagraph()
                $selection.Font.Bold = $false

                $selection.ParagraphFormat.LeftIndent = 36
                Add-WordText -Selection $selection -Text $matchingRec.Recommendations
                $selection.ParagraphFormat.LeftIndent = 0  # Reset indent
                $selection.TypeParagraph()
            }

            # Add spacing between items (except after the last one)
            if ($rank -lt $Top10Data.Count) {
                $selection.TypeParagraph()
            }

            $rank++
        }

        # Save document
        Write-Log "Saving document to: $OutputPath"

        # Workaround: Word SaveAs fails with "String is longer than 255 characters" when path exceeds ~255 chars,
        # or on OneDrive/sync folders. Save to temp first, then copy to final destination.
        $pathToSave = $OutputPath
        $wordSaveTempPath = $null
        if ($OutputPath.Length -gt 255 -or $OutputPath -match 'OneDrive|iCloud|Dropbox|Google Drive|Box\.com') {
            $tempDir = [System.IO.Path]::GetTempPath()
            $baseName = [System.IO.Path]::GetFileName($OutputPath)
            if ([string]::IsNullOrEmpty($baseName)) { $baseName = "VScanMagic_Word_Report.docx" }
            $wordSaveTempPath = Join-Path $tempDir ("VScanMagic_Word_" + [Guid]::NewGuid().ToString("N") + "_" + $baseName)
            $pathToSave = [System.IO.Path]::GetFullPath($wordSaveTempPath)
            Write-Log "Saving to temp first (path length or cloud sync workaround): $pathToSave" -Level Info
        }

        # Delete existing file if present (Word SaveAs can be finicky with overwriting)
        if (Test-Path -LiteralPath $pathToSave) {
            try {
                Remove-Item -LiteralPath $pathToSave -Force -ErrorAction Stop
                Write-Log "Removed existing Word report file"
            } catch {
                Write-Log "Warning: Could not delete existing file, attempting SaveAs anyway: $($_.Exception.Message)" -Level Warning
            }
        }

        $doc.SaveAs([ref]$pathToSave, [ref]16)  # 16 = wdFormatDocumentDefault (.docx)

        # Copy from temp to final destination if workaround was used
        if ($wordSaveTempPath -and (Test-Path -LiteralPath $wordSaveTempPath)) {
            try {
                if (Test-Path -LiteralPath $OutputPath) { Remove-Item -LiteralPath $OutputPath -Force -ErrorAction SilentlyContinue }
                Copy-Item -LiteralPath $wordSaveTempPath -Destination $OutputPath -Force
                Remove-Item -LiteralPath $wordSaveTempPath -Force -ErrorAction SilentlyContinue
                Write-Log "Copied Word report to final destination" -Level Info
            } catch {
                Write-Log "Warning: Could not copy Word report to final destination: $($_.Exception.Message)" -Level Warning
            }
        }

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

# Auto-resize downloaded XLSX reports: process smaller workbooks first, All Vulnerabilities last (largest / OneDrive-sensitive).
function Invoke-AutoResizeDownloadedXlsx {
    param(
        [Parameter(Mandatory = $false)]
        [AllowNull()]
        [array]$Succeeded,
        [Parameter(Mandatory = $true)]
        [scriptblock]$OutputPathResolver
    )
    if (-not $Succeeded) { return }
    $ordered = @(
        $Succeeded |
            Where-Object { $_.Ext -eq 'xlsx' } |
            Sort-Object { if ($_.Type -eq 'all-vulnerabilities') { 1 } else { 0 } }
    )
    foreach ($r in $ordered) {
        $p = & $OutputPathResolver $r
        if (Test-Path -LiteralPath $p) {
            Invoke-AutoResizeExcelColumns -ExcelPath $p
        }
    }
}

# Auto-resize columns on Excel workbook. Excludes Company and Proposed Remediations (all) sheets.
function Invoke-AutoResizeExcelColumns {
    param([string]$ExcelPath)
    if (-not $ExcelPath -or -not (Test-Path -LiteralPath $ExcelPath)) { return }
    $excel = $null
    $workbook = $null
    $pathToOpen = $ExcelPath
    $tempCopyPath = $null
    try {
        # Workaround: copy to temp when path is in OneDrive/sync folder (Excel COM can fail to open)
        if ($ExcelPath -match 'OneDrive|iCloud|Dropbox|Google Drive|Box\.com') {
            $tempDir = [System.IO.Path]::GetTempPath()
            $baseName = [System.IO.Path]::GetFileName($ExcelPath)
            $tempCopyPath = Join-Path $tempDir ("VScanMagic_Resize_" + [Guid]::NewGuid().ToString("N") + "_" + $baseName)
            Copy-Item -LiteralPath $ExcelPath -Destination $tempCopyPath -Force
            $pathToOpen = [System.IO.Path]::GetFullPath($tempCopyPath)
            Start-Sleep -Milliseconds 450
        }
        $opened = Open-ExcelWorkbookWithRetry -Path $pathToOpen -ReadOnly $false -MaxAttempts 5
        $excel = $opened.ExcelApp
        $workbook = $opened.Workbook
        $excludeSheets = @('Company', 'Proposed Remediations (all)')
        foreach ($ws in $workbook.Worksheets) {
            $name = [string]$ws.Name
            if ($name -notin $excludeSheets) {
                try {
                    $ws.UsedRange.Columns.AutoFit() | Out-Null
                } catch {
                    Write-Log "Auto-resize skipped sheet '$name': $($_.Exception.Message)" -Level Warning
                }
            }
        }
        $workbook.Save()
        $workbook.Close($false)
        if ($tempCopyPath -and (Test-Path -LiteralPath $tempCopyPath)) {
            Copy-Item -LiteralPath $tempCopyPath -Destination $ExcelPath -Force
            Remove-Item -LiteralPath $tempCopyPath -Force -ErrorAction SilentlyContinue
        }
        Write-Log "Auto-resized columns in $([System.IO.Path]::GetFileName($ExcelPath))" -Level Info
    } catch {
        Write-Log "Auto-resize failed for $ExcelPath : $($_.Exception.Message)" -Level Warning
    } finally {
        if ($workbook) { try { $workbook.Close($false) } catch {}; Clear-ComObject $workbook }
        if ($excel) { try { $excel.Quit() } catch {}; Clear-ComObject $excel }
        if ($tempCopyPath -and (Test-Path -LiteralPath $tempCopyPath)) { Remove-Item -LiteralPath $tempCopyPath -Force -ErrorAction SilentlyContinue }
        [System.GC]::Collect()
        [System.GC]::WaitForPendingFinalizers()
        Start-Sleep -Milliseconds 350
    }
}

# Helper: convert any value to string for Excel text cells (avoids Double-to-String COM cast errors)
function ConvertTo-SafeExcelString {
    param([object]$Value)
    if ($null -eq $Value -or $Value -is [DBNull]) { return '' }
    if ($Value -is [string]) { return $Value }
    if ($Value -is [double] -or $Value -is [float] -or $Value -is [decimal]) { return $Value.ToString([System.Globalization.CultureInfo]::InvariantCulture) }
    if ($Value -is [int] -or $Value -is [long]) { return $Value.ToString() }
    try { return [System.Convert]::ToString($Value) } catch {}
    try { return $Value.ToString() } catch {}
    return ''
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
        # Ensure Desktop folder exists (Excel COM can fail with "Unable to get the Open property"
        # when running in contexts without proper profile - e.g. RDP, packaged exe, scheduled task)
        $desktopPath = if ([Environment]::Is64BitProcess) {
            "C:\Windows\System32\config\systemprofile\Desktop"
        } else {
            "C:\Windows\SysWOW64\config\systemprofile\Desktop"
        }
        if (-not (Test-Path $desktopPath)) {
            try {
                New-Item -ItemType Directory -Path $desktopPath -Force -ErrorAction Stop | Out-Null
                Write-Log "Created Excel automation profile folder: $desktopPath" -Level Info
            } catch {
                Write-Log "Could not create profile Desktop (non-fatal): $($_.Exception.Message)" -Level Warning
            }
        }

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
        # Workaround for OneDrive/sync folder locks - copy to temp before opening.
        # Excel COM can fail with "Unable to get the Open property" on cloud-synced paths
        # or when exe bitness (32/64) doesn't match Excel. Build with -x64 to match Office.
        $pathToOpen = $InputPath
        $tempPath = $null  # Used by OneDrive workaround (input) and finally cleanup
        $saveTempPath = $null  # Used by OneDrive workaround (output save) and finally cleanup
        if ($InputPath -match 'OneDrive|iCloud|Dropbox|Google Drive|Box\.com') {
            $tempDir = [System.IO.Path]::GetTempPath()
            $baseName = [System.IO.Path]::GetFileName($InputPath)
            if ([string]::IsNullOrEmpty($baseName)) { $baseName = "vuln_report.xlsx" }
            $tempPath = Join-Path $tempDir ("VScanMagic_" + [Guid]::NewGuid().ToString("N") + "_" + $baseName)
            Copy-Item -LiteralPath $InputPath -Destination $tempPath -Force
            $pathToOpen = $tempPath
            Write-Log "Copied to temp (OneDrive workaround): $tempPath" -Level Info
        }

        # Ensure canonical path (helps Excel COM with long/UNC paths)
        if (Test-Path -LiteralPath $pathToOpen) {
            $pathToOpen = [System.IO.Path]::GetFullPath($pathToOpen)
        }

        Write-Log "Opening input workbook..." -Level Info
        if (Test-FileLocked $pathToOpen) {
            if ($tempPath) { Remove-Item -LiteralPath $tempPath -Force -ErrorAction SilentlyContinue }
            throw "The file is in use by another process. Please close it and try again."
        }
        # Use single-arg Open to avoid COM parameter binding issues; ReadOnly via optional params can trigger "Unable to get the Open property"
        try {
            $workbook = $excel.Workbooks.Open($pathToOpen)
        } catch {
            if (-not $tempPath) {
                $tempDir = [System.IO.Path]::GetTempPath()
                $baseName = [System.IO.Path]::GetFileName($InputPath)
                if ([string]::IsNullOrEmpty($baseName)) { $baseName = "vuln_report.xlsx" }
                $tempPath = Join-Path $tempDir ("VScanMagic_" + [Guid]::NewGuid().ToString("N") + "_" + $baseName)
                Copy-Item -LiteralPath $InputPath -Destination $tempPath -Force
                $pathToOpen = [System.IO.Path]::GetFullPath($tempPath)
                Write-Log "Open failed, retrying from temp copy: $($_.Exception.Message)" -Level Warning
                $workbook = $excel.Workbooks.Open($pathToOpen)
            } else { throw }
        }
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

        # Check if this is a full list format file
        $isFullListFormat = Test-IsFullListFormat -Workbook $workbook

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

        if ($isFullListFormat) {
            Write-Log "Processing as full vulnerability list format for Excel report..." -Level Info
            
            # Find vulnerability severity sheets
            $fullListSheetPatterns = @(
                "*Critical Vulnerabilities*",
                "*High Vulnerabilities*",
                "*Medium Vulnerabilities*",
                "*Low Vulnerabilities*"
            )
            
            $sourceSheets = @()
            foreach ($sheet in $workbook.Worksheets) {
                $sheetName = $sheet.Name
                foreach ($pattern in $fullListSheetPatterns) {
                    if ($sheetName -like $pattern) {
                        Write-Log "Found vulnerability list sheet: $sheetName"
                        $sourceSheets += $sheet
                        break
                    }
                }
            }
            
            if ($sourceSheets.Count -eq 0) {
                throw "No vulnerability list sheets found for consolidation"
            }
            
            # Read all vulnerabilities and aggregate them
            $allVulnerabilities = @()
            foreach ($sheet in $sourceSheets) {
                Write-Log "Reading vulnerabilities from: $($sheet.Name)" -Level Info
                $usedRange = $sheet.UsedRange
                $rowCount = $usedRange.Rows.Count
                
                if ($rowCount -le 1) { continue }
                
                # Get headers
                $headers = @{}
                $colCount = $usedRange.Columns.Count
                for ($col = 1; $col -le $colCount; $col++) {
                    $headerName = $sheet.Cells.Item(1, $col).Text
                    if ($headerName) {
                        $headers[$headerName] = $col
                    }
                }
                
                # Find column indices
                $hostNameCol = $null
                $ipCol = $null
                $productCol = $null
                $severityCol = $null
                $epssCol = $null
                $fixCol = $null
                $evidencePathCol = $null
                $evidenceVersionCol = $null
                
                foreach ($header in $headers.Keys) {
                    if ($header -match '^(Host Name|Hostname)$') { $hostNameCol = $headers[$header] }
                    if ($header -match '^(IP|IP Address)$') { $ipCol = $headers[$header] }
                    if ($header -match '^(Software Name|Product|Software)$') { $productCol = $headers[$header] }
                    if ($header -eq 'Severity') { $severityCol = $headers[$header] }
                    if ($header -match '^(EPSS Score|EPSS)$') { $epssCol = $headers[$header] }
                    if ($header -match '^(Solution|Fix)$') { $fixCol = $headers[$header] }
                    if ($header -match '^(Evidence Path)$') { $evidencePathCol = $headers[$header] }
                    if ($header -match '^(Evidence Version)$') { $evidenceVersionCol = $headers[$header] }
                }
                
                # Read vulnerabilities (use ToString() to avoid Int32-to-String cast errors when Excel stores text as numbers)
                $rangeValues = $usedRange.Value2
                $safeStr = { param($v) if ($null -eq $v -or ($v -is [DBNull])) { '' } else { $v.ToString().Trim() } }
                for ($row = 2; $row -le $rowCount; $row++) {
                    $hostName = if ($hostNameCol) { & $safeStr $rangeValues[$row, $hostNameCol] } else { '' }
                    $ip = if ($ipCol) { & $safeStr $rangeValues[$row, $ipCol] } else { '' }
                    $product = if ($productCol) { & $safeStr $rangeValues[$row, $productCol] } else { '' }
                    $product = $product -replace "^\[|'|\]$", "" -replace "^'|'$", ""
                    $severity = if ($severityCol) { & $safeStr $rangeValues[$row, $severityCol] } else { '' }
                    $epssScore = if ($epssCol) { Get-SafeDoubleValue -Value (& $safeStr $rangeValues[$row, $epssCol]) } else { 0.0 }
                    $fix = if ($fixCol) { & $safeStr $rangeValues[$row, $fixCol] } else { '' }
                    $evidencePath = if ($evidencePathCol) { & $safeStr $rangeValues[$row, $evidencePathCol] } else { '' }
                    $evidenceVersion = if ($evidenceVersionCol) { & $safeStr $rangeValues[$row, $evidenceVersionCol] } else { '' }
                    
                    if (-not [string]::IsNullOrWhiteSpace($product) -and -not [string]::IsNullOrWhiteSpace($severity)) {
                        $allVulnerabilities += [PSCustomObject]@{
                            'Host Name' = $hostName
                            'IP' = $ip
                            'Product' = $product
                            'Severity' = $severity
                            'EPSS Score' = $epssScore
                            'Fix' = $fix
                            'Evidence Path' = $evidencePath
                            'Evidence Version' = $evidenceVersion
                        }
                    }
                }
                Clear-ComObject $usedRange
            }
            
            Write-Log "Read $($allVulnerabilities.Count) vulnerabilities from full list format" -Level Info
            
            # Aggregate by Host/Product (single-pass severity counts and field extraction)
            $aggregatedData = $allVulnerabilities | Group-Object -Property {
                "$($_.'Host Name')|$($_.IP)|$($_.Product)"
            } | ForEach-Object {
                $group = $_
                $firstItem = $group.Group[0]
                $counts = Get-SeverityCounts -Items $group.Group -EPSSProp 'EPSS Score'
                $vulnCount = $counts.Critical + $counts.High + $counts.Medium + $counts.Low
                $fix = $null; $evidencePath = $null; $evidenceVersion = $null
                foreach ($it in $group.Group) {
                    if (-not $fix -and -not [string]::IsNullOrWhiteSpace($it.Fix)) { $fix = $it.Fix }
                    if (-not $evidencePath -and -not [string]::IsNullOrWhiteSpace($it.'Evidence Path')) { $evidencePath = $it.'Evidence Path' }
                    if (-not $evidenceVersion -and -not [string]::IsNullOrWhiteSpace($it.'Evidence Version')) { $evidenceVersion = $it.'Evidence Version' }
                    if ($fix -and $evidencePath -and $evidenceVersion) { break }
                }
                
                [PSCustomObject]@{
                    'Remediation Type' = ''  # Not available in full list format
                    'Product' = $firstItem.Product
                    'Host Name' = $firstItem.'Host Name'
                    'Fix' = if ($fix) { $fix } else { '' }
                    'IP' = $firstItem.IP
                    'Evidence Path' = if ($evidencePath) { $evidencePath } else { '' }
                    'Evidence Version' = if ($evidenceVersion) { $evidenceVersion } else { '' }
                    'Critical' = $counts.Critical
                    'High' = $counts.High
                    'Medium' = $counts.Medium
                    'Low' = $counts.Low
                    'Vulnerability Count' = $vulnCount
                    'EPSS Score' = $counts.MaxEPSS
                }
            }
            
            Write-Log "Aggregated to $($aggregatedData.Count) Host/Product combinations" -Level Info
            
            # Write headers to Source Data sheet
            try {
                $headers = @('Remediation Type', 'Product', 'Host Name', 'Fix', 'IP', 'Evidence Path', 'Evidence Version', 'Critical', 'High', 'Medium', 'Low', 'Vulnerability Count', 'EPSS Score')
                for ($col = 1; $col -le $headers.Count; $col++) {
                    $sourceDataSheet.Cells.Item(1, $col).Value2 = $headers[$col - 1]
                }
                Write-Log "Headers written" -Level Info
            } catch { throw "[FullListFormat-WriteHeaders] $($_.Exception.Message)" }
            
            # Write aggregated data - each column in try/catch to identify failing column
            $row = 2
            $rowNum = 0
            $colNames = @('Remediation Type','Product','Host Name','Fix','IP','Evidence Path','Evidence Version','Critical','High','Medium','Low','Vulnerability Count','EPSS Score')
            $totalRows = $aggregatedData.Count
            foreach ($item in $aggregatedData) {
                $rowNum++
                if ($rowNum % 100 -eq 0) {
                    [System.Windows.Forms.Application]::DoEvents()
                    if ($rowNum % 500 -eq 0) { Write-Log "Writing row $rowNum of $totalRows..." -Level Info }
                }
                for ($col = 1; $col -le 13; $col++) {
                    try {
                        if ($col -le 7) {
                            $val = $item.($colNames[$col-1]); $s = ConvertTo-SafeExcelString $val; $esc = $s -replace '"','""'
                            $cell = $sourceDataSheet.Cells.Item($row, $col); $cell.Formula = '="' + $esc + '"'; Clear-ComObject $cell
                        } else {
                            if ($col -eq 8) { $sourceDataSheet.Cells.Item($row, $col).Value2 = [int]($item.Critical) }
                            elseif ($col -eq 9) { $sourceDataSheet.Cells.Item($row, $col).Value2 = [int]($item.High) }
                            elseif ($col -eq 10) { $sourceDataSheet.Cells.Item($row, $col).Value2 = [int]($item.Medium) }
                            elseif ($col -eq 11) { $sourceDataSheet.Cells.Item($row, $col).Value2 = [int]($item.Low) }
                            elseif ($col -eq 12) { $sourceDataSheet.Cells.Item($row, $col).Value2 = [int]($item.'Vulnerability Count') }
                            elseif ($col -eq 13) {
                                $ev = $item.'EPSS Score'
                                $numVal = if ($null -ne $ev) { [double]$ev } else { 0.0 }
                                $cell = $sourceDataSheet.Cells.Item($row, $col)
                                $cell.NumberFormat = '0.00'
                                $cell.Formula = '=' + $numVal.ToString([System.Globalization.CultureInfo]::InvariantCulture)
                                Clear-ComObject $cell
                            }
                        }
                    } catch { throw "[FullListFormat-WriteData row $rowNum col $($colNames[$col-1])] $($_.Exception.Message)" }
                }
                $row++
            }
            
            Write-Log "Data consolidation complete for full list format" -Level Info
            
            # Release source sheet references
            foreach ($sheet in $sourceSheets) {
                Clear-ComObject $sheet
            }
        } else {
            # Original aggregated format processing
            Write-Log "Processing as aggregated format..." -Level Info
            
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
                    $sourceRow = $firstValidSheet.Rows(1)
                    $targetRow = $sourceDataSheet.Rows(1)

                    # Copy column by column to avoid type casting issues (ensure headers are strings)
                    for ($col = 1; $col -le $sourceCols; $col++) {
                        $val = $firstValidSheet.Cells.Item(1, $col).Value2
                        $sourceDataSheet.Cells.Item(1, $col).Value2 = ConvertTo-SafeExcelString $val
                    }

                    Write-Log "Headers copied successfully ($sourceCols columns)" -Level Info
                    Clear-ComObject $sourceRow
                    Clear-ComObject $targetRow
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
                    $sourceCols = $sourceRange.Columns.Count

                    if ($sourceRows -gt 1) {
                        # Use Excel's Copy/PasteSpecial to avoid PowerShell type casting
                        $sourceDataRange = $sourceSheet.Range($sourceSheet.Cells.Item(2, 1), $sourceSheet.Cells.Item($sourceRows, $sourceCols))
                        $targetCell = $sourceDataSheet.Cells.Item($destRow, 1)

                        # Copy and paste values only (not formulas/formatting)
                        $sourceDataRange.Copy()
                        $targetCell.PasteSpecial(-4163)  # xlPasteValues = -4163
                        $excel.Application.CutCopyMode = $false  # Clear clipboard

                        $rowsCopied = $sourceRows - 1
                        $destRow += $rowsCopied
                        Write-Log "  Copied $rowsCopied data rows" -Level Info
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
        }

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

        # Create pivot table (use string address for SourceData to avoid type mismatch/cast errors)
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

            [System.Windows.Forms.Application]::DoEvents()
            $pivotCache = $workbook.PivotCaches().Create(1, $pivotSourceRange)
            $pivotTable = $pivotCache.CreatePivotTable($pivotSheet.Range("A3"), "VulnPivotTable")
            Write-Log "Pivot Table object created" -Level Info
            [System.Windows.Forms.Application]::DoEvents()

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
                    $thresholdStr = ConvertTo-SafeExcelString $script:Config.ConditionalFormatThreshold
                    $cfCondition = $cfRange.FormatConditions.Add(1, 5, $thresholdStr)
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
                    @{Text = "Do not touch"; BgColorIndex = 46; FontColorIndex = 2; Strikethrough = $false}
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

            if ($null -ne $dataField1) { Clear-ComObject $dataField1 }
            if ($null -ne $dataField2) { Clear-ComObject $dataField2 }
            Clear-ComObject $pivotTable
            Clear-ComObject $pivotCache
        }

        Clear-ComObject $pivotSourceRange

        # --- 4. Save and Close ---
        [System.Windows.Forms.Application]::DoEvents()
        Write-Log "Saving workbook to: $OutputPath" -Level Info

        # Workaround: Excel COM SaveAs fails on OneDrive/sync folders ("Unable to get the SaveAs property").
        # Save to temp first, then copy to final destination.
        $pathToSave = $OutputPath
        if ($OutputPath -match 'OneDrive|iCloud|Dropbox|Google Drive|Box\.com') {
            $tempDir = [System.IO.Path]::GetTempPath()
            $baseName = [System.IO.Path]::GetFileName($OutputPath)
            if ([string]::IsNullOrEmpty($baseName)) { $baseName = "VScanMagic_report.xlsx" }
            $saveTempPath = Join-Path $tempDir ("VScanMagic_Save_" + [Guid]::NewGuid().ToString("N") + "_" + $baseName)
            $pathToSave = [System.IO.Path]::GetFullPath($saveTempPath)
            Write-Log "Saving to temp first (OneDrive workaround): $pathToSave" -Level Info
        }

        # Delete existing file if present (Excel SaveAs can be finicky with overwriting)
        if (Test-Path -LiteralPath $pathToSave) {
            try {
                Remove-Item -LiteralPath $pathToSave -Force -ErrorAction Stop
                Write-Log "Removed existing Excel report file"
            } catch {
                Write-Log "Warning: Could not delete existing file, attempting SaveAs anyway: $($_.Exception.Message)" -Level Warning
            }
        }

        $workbook.SaveAs($pathToSave)
        $workbook.Close($false)

        # Copy from temp to final destination if OneDrive workaround was used
        if ($saveTempPath -and (Test-Path -LiteralPath $saveTempPath)) {
            try {
                if (Test-Path $OutputPath) { Remove-Item -LiteralPath $OutputPath -Force -ErrorAction SilentlyContinue }
                Copy-Item -LiteralPath $saveTempPath -Destination $OutputPath -Force
                Remove-Item -LiteralPath $saveTempPath -Force -ErrorAction SilentlyContinue
                Write-Log "Copied Excel report to final destination" -Level Info
            } catch {
                Write-Log "Failed to copy Excel report to final destination: $($_.Exception.Message)" -Level Error
                throw
            }
        }

        Write-Log "Excel report generation complete" -Level Success

    } catch {
        Write-Log "Excel report generation failed: $($_.Exception.Message)" -Level Error
        throw $_
    } finally {
        # Cleanup temp copies if created for OneDrive workaround (input open + output save)
        if ($tempPath -and (Test-Path -LiteralPath $tempPath)) {
            try { Remove-Item -LiteralPath $tempPath -Force -ErrorAction SilentlyContinue } catch {}
        }
        if ($saveTempPath -and (Test-Path -LiteralPath $saveTempPath)) {
            try { Remove-Item -LiteralPath $saveTempPath -Force -ErrorAction SilentlyContinue } catch {}
        }
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

# --- Report Generation Helper Functions ---

function Get-CurrentQuarter {
    $month = (Get-Date).Month
    return [Math]::Ceiling($month / 3)
}

function Get-TimeOfDayGreeting {
    $hour = (Get-Date).Hour
    if ($hour -lt 12) {
        return "morning"
    } elseif ($hour -lt 17) {
        return "afternoon"
    } else {
        return "evening"
    }
}

function Open-EmailDraftInOutlook {
    <#
    .SYNOPSIS
    Opens the email template from the given output folder in Outlook as a draft (avoids signature-at-top when opening .eml).
    #>
    param([string]$OutputFolderPath)
    if ([string]::IsNullOrWhiteSpace($OutputFolderPath) -or -not (Test-Path $OutputFolderPath)) { return $false }
    $emlFiles = @(Get-ChildItem -Path $OutputFolderPath -Filter "*Email Template*.eml" -ErrorAction SilentlyContinue)
    $miscPath = Join-Path $OutputFolderPath "Misc"
    if ((Test-Path $miscPath) -and $emlFiles.Count -eq 0) {
        $emlFiles = @(Get-ChildItem -Path $miscPath -Filter "*Email Template*.eml" -ErrorAction SilentlyContinue)
    }
    if ($emlFiles.Count -eq 0) { return $false }
    $emlPath = $emlFiles | Sort-Object LastWriteTime -Descending | Select-Object -First 1
    $content = Get-Content -Path $emlPath.FullName -Raw -Encoding UTF8
    $headerEnd = $content.IndexOf("`r`n`r`n")
    if ($headerEnd -lt 0) { $headerEnd = $content.IndexOf("`n`n") }
    $headers = if ($headerEnd -ge 0) { $content.Substring(0, $headerEnd) } else { $content }
    $bodyEncoded = if ($headerEnd -ge 0) { $content.Substring($headerEnd).Trim() } else { "" }
    $subject = ""
    foreach ($line in ($headers -split "`r?`n")) {
        if ($line -match "^Subject:\s*(.+)$") { $subject = $matches[1].Trim(); break }
    }
    $body = ""
    if ($bodyEncoded) {
        try {
            $bodyBytes = [Convert]::FromBase64String(($bodyEncoded -replace "`r`n", ""))
            $body = [System.Text.Encoding]::UTF8.GetString($bodyBytes)
        } catch { $body = $bodyEncoded }
    }
    try {
        $ol = New-Object -ComObject Outlook.Application
        $mail = $ol.CreateItem(0)
        $mail.Subject = $subject
        $mail.Body = $body
        $mail.Display()
        [System.Runtime.Interopservices.Marshal]::ReleaseComObject($mail) | Out-Null
        [System.Runtime.Interopservices.Marshal]::ReleaseComObject($ol) | Out-Null
        return $true
    } catch {
        return $false
    }
}

function Format-EmailTemplateSpacing {
    <# Collapse multiple spaces to single; collapse mid-paragraph line breaks; preserve paragraph breaks (double newline). #>
    param([string]$Text)
    if ([string]::IsNullOrWhiteSpace($Text)) { return $Text }
    $paragraphs = $Text -split "`r?`n`r?`n"
    $result = ($paragraphs | ForEach-Object {
        $para = ($_ -split "`r?`n") -join " "
        $para = $para -replace '[ \t]+', ' '
        $para.Trim()
    }) -join "`r`n`r`n"
    return $result.Trim()
}

function New-EmailTemplate {
    param(
        [string]$OutputPath,
        [bool]$IsRMITPlus = $false,
        [switch]$PassThru,
        [string]$FilterTopN = $null
    )

    try {
        Write-Log "Generating email template..."

        if ($null -eq $script:Templates) { Load-Templates }

        $year = (Get-Date).Year
        $quarter = Get-CurrentQuarter
        $greeting = Get-TimeOfDayGreeting

        # Build client-type-specific note
        if ($IsRMITPlus) {
            $noteText = "Note: Remediation tickets have been generated for items covered under your RMIT+ agreement. Third-party items not covered under the agreement will not be remediated unless we discuss them and a quote has been generated. To schedule a discussion, please use the scheduling link above."
        } else {
            $noteText = "Note: We will not generate remediation tickets without your approval. To schedule a discussion of these findings, please use the scheduling link above."
        }

        $topNLabel = if ([string]::IsNullOrWhiteSpace($FilterTopN)) { $script:FilterTopN } else { $FilterTopN }
        if ([string]::IsNullOrWhiteSpace($topNLabel)) { $topNLabel = "10" }
        $topNLabel = if ($topNLabel -eq "All") { "Top" } elseif ($topNLabel -eq "10") { "Top Ten" } elseif (-not [string]::IsNullOrWhiteSpace($topNLabel)) { "Top $topNLabel" } else { "Top Ten" }
        $bodyTemplate = $script:Templates.EmailTemplate.Body
        $emailContent = $bodyTemplate -replace '\{Year\}', $year -replace '\{Quarter\}', $quarter -replace '\{Greeting\}', $greeting -replace '\{NoteText\}', $noteText -replace '\{PreparedBy\}', $script:UserSettings.PreparedBy -replace '\{TopNLabel\}', $topNLabel
        $emailContent = Format-EmailTemplateSpacing -Text $emailContent

        if ($OutputPath) {
            $emailContent | Out-File -FilePath $OutputPath -Encoding UTF8
        }

        # Generate .eml file for one-click open in default email client
        $emlPath = if ($OutputPath) { [System.IO.Path]::ChangeExtension($OutputPath, ".eml") } else { $null }
        $lines = $emailContent -split "`r?`n"
        $subject = if ($lines.Count -gt 0 -and $lines[0] -match '^Subject:\s*(.+)$') { $matches[1].Trim() } else { "$year Q$quarter Vulnerability Scan Follow Up" }
        $body = if ($lines.Count -gt 1) { ($lines[1..($lines.Count - 1)] -join "`r`n").Trim() } else { $emailContent }
        $body = $body -replace "`r`n", "`n" -replace "`n", "`r`n"  # Normalize to CRLF for email
        $dateRfc = (Get-Date).ToString("r")
        $bodyBytes = [System.Text.Encoding]::UTF8.GetBytes($body)
        $bodyBase64 = [Convert]::ToBase64String($bodyBytes)
        $chunks = for ($i = 0; $i -lt $bodyBase64.Length; $i += 76) { $bodyBase64.Substring($i, [Math]::Min(76, $bodyBase64.Length - $i)) }
        $bodyBase64Wrapped = $chunks -join "`r`n"
        $emlLines = @(
            "From: ",
            "To: ",
            "Subject: $subject",
            "Date: $dateRfc",
            "X-Unsent: 1",
            "MIME-Version: 1.0",
            "Content-Type: text/plain; charset=utf-8",
            "Content-Transfer-Encoding: base64",
            "",
            $bodyBase64Wrapped
        )
        if ($emlPath) {
            $emlLines -join "`r`n" | Out-File -FilePath $emlPath -Encoding ASCII
            Write-Log "Email template (.eml) saved to: $emlPath" -Level Success

            # Create shortcut (.lnk) to open .eml in default email client
            $shortcutPath = Join-Path ([System.IO.Path]::GetDirectoryName($OutputPath)) ([System.IO.Path]::GetFileNameWithoutExtension($OutputPath) + " - Open in Email.lnk")
            try {
                $ws = New-Object -ComObject WScript.Shell
                $sc = $ws.CreateShortcut($shortcutPath)
                $sc.TargetPath = $emlPath
                $sc.WorkingDirectory = [System.IO.Path]::GetDirectoryName($emlPath)
                $sc.Description = "Open vulnerability scan follow-up email in default email client"
                $sc.Save()
                [System.Runtime.Interopservices.Marshal]::ReleaseComObject($ws) | Out-Null
                Write-Log "Shortcut created: $shortcutPath" -Level Success
            } catch {
                Write-Log "Could not create shortcut: $($_.Exception.Message)" -Level Warning
            }
        }

        if ($PassThru) { return $emailContent }

    } catch {
        Write-Log "Error generating email template: $($_.Exception.Message)" -Level Error
    }
}

