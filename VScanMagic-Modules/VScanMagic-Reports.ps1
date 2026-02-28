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
        $maxRiskScore = 10  # Default fallback (max theoretical: 1.0 × 10.0)
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

        # Set document properties (optional - may fail on some systems)
        Write-Log "Setting document properties..."
        try {
            $doc.BuiltInDocumentProperties.Item("Title").Value = "$ReportTitle - $ClientName"
            $doc.BuiltInDocumentProperties.Item("Subject").Value = "Security Vulnerability Assessment"
            $doc.BuiltInDocumentProperties.Item("Author").Value = $script:Config.Author
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

            for ($col = 1; $col -le 7; $col++) {
                $table.Cell($rowIndex, $col).Shading.BackgroundPatternColor = $bgColor
                $table.Cell($rowIndex, $col).Range.Font.Color = $textColor
            }

            $rank++
        }

        # Set custom column widths for better appearance
        # Column widths in points (1 inch = 72 points)
        $table.Columns(1).SetWidth(36, 0)   # Rank: 0.5 inch (narrow)
        $table.Columns(2).PreferredWidthType = 3  # wdPreferredWidthPoints (3, not 2)
        $table.Columns(2).PreferredWidth = 180    # Product/System: 2.5 inches max

        # Auto-fit other columns - use table AutoFitBehavior after setting column 1 & 2
        $table.Columns(3).SetWidth(72, 0)   # Risk Score: 1 inch
        $table.Columns(4).SetWidth(54, 0)   # EPSS: 0.75 inch
        $table.Columns(5).SetWidth(72, 0)   # Avg CVSS: 1 inch
        $table.Columns(6).SetWidth(72, 0)   # Total Vulns: 1 inch
        $table.Columns(7).SetWidth(90, 0)   # Affected Systems: 1.25 inch

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
            $sortedChartData = $Top10Data | Sort-Object -Property VulnCount -Descending

            $row = 2
            foreach ($item in $sortedChartData) {
                $worksheet.Cells.Item($row, 1) = [string]$item.Product
                $worksheet.Cells.Item($row, 2) = [int]$item.VulnCount
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

            $timeEstimate = if ($timeEstimateMap.ContainsKey($item.Product)) { $timeEstimateMap[$item.Product] } else { $null }

            # Add modifier text based on checkbox states (only if time estimates are available)
            if ($null -ne $timeEstimate) {
                $afterHours = $timeEstimate.AfterHours
                $ticketGenerated = if ($IsRMITPlus) { $timeEstimate.TicketGenerated } else { $false }
                $thirdParty = if ($IsRMITPlus) { $timeEstimate.ThirdParty } else { $false }

                $modifierText = Get-ModifierText -AfterHours $afterHours -TicketGenerated $ticketGenerated -ThirdParty $thirdParty
                if (-not [string]::IsNullOrWhiteSpace($modifierText)) {
                    $title += $modifierText
                }
            }

            # Add End of Life note for Windows 10 (only if no checkbox suffix was added)
            if ($item.Product -like "*Windows 10*" -and $null -eq $timeEstimate) {
                $title += " - Windows 10 is End of Life"
            }

            # Add RMIT+ note for Microsoft applications (not OS) - only for RMIT+ clients (only if no checkbox suffix was added)
            $isMicrosoftApp = Test-IsMicrosoftApplication -ProductName $item.Product
            if ($isMicrosoftApp -and $IsRMITPlus -and $null -eq $timeEstimate) {
                $title += " - RMIT+ ticketed"
            }

            # Add after-hours ticket note for VMware products - only for RMIT+ clients (only if no checkbox suffix was added)
            $isVMwareProduct = Test-IsVMwareProduct -ProductName $item.Product
            if ($isVMwareProduct -and $IsRMITPlus -and $null -eq $timeEstimate) {
                $title += " - RMIT+ after-hours ticket created if we maintain this"
            }

            # Add auto-update note for Chrome/Firefox
            $isAutoUpdating = Test-IsAutoUpdatingSoftware -ProductName $item.Product
            if ($isAutoUpdating) {
                $title += " - This software updates automatically"
            }

            $selection.TypeText($title)
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
            # Use hostname or IP as identifier; format as "hostname (username)" if username present
            $selection.ParagraphFormat.LeftIndent = 36
            $systemsList = ($item.AffectedSystems | ForEach-Object {
                $display = if ($_.HostName) { $_.HostName } else { $_.IP }
                $username = $_.Username
                if (-not [string]::IsNullOrWhiteSpace($username)) {
                    "$display ($username)"
                } else {
                    $display
                }
            } | Where-Object { -not [string]::IsNullOrWhiteSpace($_) } | Select-Object -Unique) -join ", "
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

            # Get remediation guidance from configurable rules
            $remediationText = Get-RemediationGuidance -ProductName $item.Product -OutputType 'Word'
            $selection.TypeText($remediationText)

            $selection.ParagraphFormat.LeftIndent = 0  # Reset indent
            $selection.TypeParagraph()

            $matchingRec = if ($generalRecMap.ContainsKey($item.Product)) { $generalRecMap[$item.Product] } else { $null }
            if ($null -ne $matchingRec -and -not [string]::IsNullOrWhiteSpace($matchingRec.Recommendations)) {
                $selection.Font.Bold = $true
                $selection.TypeText("General Recommendations:")
                $selection.TypeParagraph()
                $selection.Font.Bold = $false

                $selection.ParagraphFormat.LeftIndent = 36
                $selection.TypeText($matchingRec.Recommendations)
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

        # Delete existing file if present (Word SaveAs can be finicky with overwriting)
        if (Test-Path $OutputPath) {
            try {
                Remove-Item -Path $OutputPath -Force -ErrorAction Stop
                Write-Log "Removed existing Word report file"
            } catch {
                Write-Log "Warning: Could not delete existing file, attempting SaveAs anyway: $($_.Exception.Message)" -Level Warning
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
            $headers = @('Remediation Type', 'Product', 'Host Name', 'Fix', 'IP', 'Evidence Path', 'Evidence Version', 'Critical', 'High', 'Medium', 'Low', 'Vulnerability Count', 'EPSS Score')
            for ($col = 1; $col -le $headers.Count; $col++) {
                $sourceDataSheet.Cells.Item(1, $col).Value2 = $headers[$col - 1]
            }
            
            # Write aggregated data (ensure text columns are strings to avoid Int32-to-String cast in Pivot Table)
            $row = 2
            foreach ($item in $aggregatedData) {
                $sourceDataSheet.Cells.Item($row, 1).Value2 = [string]($item.'Remediation Type')
                $sourceDataSheet.Cells.Item($row, 2).Value2 = [string]($item.Product)
                $sourceDataSheet.Cells.Item($row, 3).Value2 = [string]($item.'Host Name')
                $sourceDataSheet.Cells.Item($row, 4).Value2 = [string]($item.Fix)
                $sourceDataSheet.Cells.Item($row, 5).Value2 = [string]($item.IP)
                $sourceDataSheet.Cells.Item($row, 6).Value2 = [string]($item.'Evidence Path')
                $sourceDataSheet.Cells.Item($row, 7).Value2 = [string]($item.'Evidence Version')
                $sourceDataSheet.Cells.Item($row, 8).Value2 = [int]($item.Critical)
                $sourceDataSheet.Cells.Item($row, 9).Value2 = [int]($item.High)
                $sourceDataSheet.Cells.Item($row, 10).Value2 = [int]($item.Medium)
                $sourceDataSheet.Cells.Item($row, 11).Value2 = [int]($item.Low)
                $sourceDataSheet.Cells.Item($row, 12).Value2 = [int]($item.'Vulnerability Count')
                $epssVal = $item.'EPSS Score'; $sourceDataSheet.Cells.Item($row, 13).Value2 = if ($null -ne $epssVal) { [double]$epssVal } else { 0.0 }
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

                    # Copy column by column to avoid type casting issues
                    for ($col = 1; $col -le $sourceCols; $col++) {
                        $sourceDataSheet.Cells.Item(1, $col).Value2 = $firstValidSheet.Cells.Item(1, $col).Value2
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

            Clear-ComObject $dataField1
            Clear-ComObject $dataField2
            Clear-ComObject $pivotTable
            Clear-ComObject $pivotCache
        }

        Clear-ComObject $pivotSourceRange

        # --- 4. Save and Close ---
        Write-Log "Saving workbook to: $OutputPath" -Level Info

        # Delete existing file if present (Excel SaveAs can be finicky with overwriting)
        if (Test-Path $OutputPath) {
            try {
                Remove-Item -Path $OutputPath -Force -ErrorAction Stop
                Write-Log "Removed existing Excel report file"
            } catch {
                Write-Log "Warning: Could not delete existing file, attempting SaveAs anyway: $($_.Exception.Message)" -Level Warning
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

function New-EmailTemplate {
    param(
        [string]$OutputPath,
        [bool]$IsRMITPlus = $false
    )

    try {
        Write-Log "Generating email template..."

        $year = (Get-Date).Year
        $quarter = Get-CurrentQuarter
        $greeting = Get-TimeOfDayGreeting

        # Build client-type-specific note
        if ($IsRMITPlus) {
            $noteText = "Note: Remediation tickets have been generated for items that are covered under your RMIT+ agreement. A TimeZest meeting request has been sent to discuss the 3rd party items that are not covered under the RMIT+ agreement. Those 3rd party items will not be remediated unless they are discussed and a quote has been generated.`n`nIf you would like to discuss the report further, please use the meeting request to schedule an appointment directly from my availability via a Teams Meeting."
        } else {
            $noteText = "Note: We will not generate any tickets without your approval. A TimeZest meeting request has been sent to discuss the remediation of these vulnerabilities. If you would like to discuss the report further, please use the meeting request to schedule an appointment directly from my availability via a Teams Meeting."
        }

        $emailContent = @"
Subject: $year Q$quarter Vulnerability Scan Follow Up

Good $greeting,

We are pleased to inform you that your quarterly vulnerability scan report has been completed and added to your client folder.

The main list of items I recommend remediating can be found here:
<link to top ten report from onedrive>
*Note that not all vulnerabilities may be feasible to remediate depending on business need.

You can access and view the full reports using the link below:
<onedrive link to folder containing reports>

In this folder you will also find:

Pending Remediation EPSS Score Report: This report classifies vulnerabilities by the "EPSS Score." This is a measure of the likelihood that an attacker will exploit a particular vulnerability within 30 days. The scale ranges from 0 to 1.0, with 1.0 being the most critical.

All Vulnerabilities Report: This spreadsheet contains a list of all vulnerabilities (including internal and external) that were detected, ranging from critical to low

Executive Summary Report: A high-level overview of your security "grade" as well as some information about your network

External Scan: Any detected vulnerabilities or services that are exposed to the outside Internet

$noteText

We appreciate your commitment to security, as addressing these vulnerabilities is essential for maintaining the ongoing protection of your systems.


Sincerely,

$($script:UserSettings.PreparedBy)
"@

        $emailContent | Out-File -FilePath $OutputPath -Encoding UTF8
        Write-Log "Email template saved to: $OutputPath" -Level Success

    } catch {
        Write-Log "Error generating email template: $($_.Exception.Message)" -Level Error
    }
}

function Show-GeneralRecommendationsDialog {
    param(
        [array]$Top10Data
    )

    # Load recommendations if not already loaded
    if ($null -eq $script:GeneralRecommendations) {
        Load-GeneralRecommendations
    }

    # Create dialog
    $recDialog = New-Object System.Windows.Forms.Form
    $recDialog.Text = "General Recommendations"
    $recDialog.Size = New-Object System.Drawing.Size(1000, 600)
    $recDialog.StartPosition = "CenterParent"
    $recDialog.FormBorderStyle = "FixedDialog"
    $recDialog.MaximizeBox = $false
    $recDialog.MinimizeBox = $false

    # Create DataGridView
    $dataGridView = New-Object System.Windows.Forms.DataGridView
    $dataGridView.Location = New-Object System.Drawing.Point(20, 20)
    $dataGridView.Size = New-Object System.Drawing.Size(950, 450)
    $dataGridView.AutoSizeColumnsMode = [System.Windows.Forms.DataGridViewAutoSizeColumnsMode]::Fill
    $dataGridView.SelectionMode = [System.Windows.Forms.DataGridViewSelectionMode]::FullRowSelect
    $dataGridView.MultiSelect = $false
    $dataGridView.AllowUserToAddRows = $false
    $recDialog.Controls.Add($dataGridView)

    # Add columns
    $colProduct = New-Object System.Windows.Forms.DataGridViewTextBoxColumn
    $colProduct.Name = "Product"
    $colProduct.HeaderText = "Product"
    $colProduct.Width = 250
    $colProduct.ReadOnly = $true
    $dataGridView.Columns.Add($colProduct) | Out-Null

    $colRecommendations = New-Object System.Windows.Forms.DataGridViewTextBoxColumn
    $colRecommendations.Name = "Recommendations"
    $colRecommendations.HeaderText = "General Recommendations"
    $colRecommendations.Width = 700
    $colRecommendations.DefaultCellStyle.WrapMode = [System.Windows.Forms.DataGridViewTriState]::True
    $dataGridView.Columns.Add($colRecommendations) | Out-Null

    # Set row height for multi-line text
    $dataGridView.AutoSizeRowsMode = [System.Windows.Forms.DataGridViewAutoSizeRowsMode]::AllCells
    $dataGridView.DefaultCellStyle.WrapMode = [System.Windows.Forms.DataGridViewTriState]::True

    # Populate grid with top 10 data and pre-populate from saved recommendations
    foreach ($item in $Top10Data) {
        $row = $dataGridView.Rows.Add()
        $dataGridView.Rows[$row].Cells["Product"].Value = $item.Product
        
        # Try to find matching recommendation using pattern matching
        $matchingRec = $null
        foreach ($rec in $script:GeneralRecommendations) {
            if ($item.Product -like $rec.Product) {
                $matchingRec = $rec
                break
            }
        }
        
        $recommendationsText = if ($matchingRec) { $matchingRec.Recommendations } else { "" }
        $dataGridView.Rows[$row].Cells["Recommendations"].Value = $recommendationsText
    }

    # Buttons
    $y = 480

    $btnLoadDefaults = New-Object System.Windows.Forms.Button
    $btnLoadDefaults.Location = New-Object System.Drawing.Point(20, $y)
    $btnLoadDefaults.Size = New-Object System.Drawing.Size(120, 30)
    $btnLoadDefaults.Text = "Load Defaults"
    $btnLoadDefaults.Add_Click({
        # Reload recommendations from disk
        Load-GeneralRecommendations
        # Refresh grid
        foreach ($row in $dataGridView.Rows) {
            if ($row.IsNewRow) { continue }
            $product = $row.Cells["Product"].Value
            $matchingRec = $null
            foreach ($rec in $script:GeneralRecommendations) {
                if ($product -like $rec.Product) {
                    $matchingRec = $rec
                    break
                }
            }
            $row.Cells["Recommendations"].Value = if ($matchingRec) { $matchingRec.Recommendations } else { "" }
        }
    })
    $recDialog.Controls.Add($btnLoadDefaults)

    $btnCancel = New-Object System.Windows.Forms.Button
    $btnCancel.Location = New-Object System.Drawing.Point(800, $y)
    $btnCancel.Size = New-Object System.Drawing.Size(90, 30)
    $btnCancel.Text = "Cancel"
    $btnCancel.DialogResult = [System.Windows.Forms.DialogResult]::Cancel
    $recDialog.Controls.Add($btnCancel)

    $btnOK = New-Object System.Windows.Forms.Button
    $btnOK.Location = New-Object System.Drawing.Point(890, $y)
    $btnOK.Size = New-Object System.Drawing.Size(90, 30)
    $btnOK.Text = "OK"
    $btnOK.DialogResult = [System.Windows.Forms.DialogResult]::OK
    $recDialog.Controls.Add($btnOK)

    # Set default button
    $recDialog.AcceptButton = $btnOK
    $recDialog.CancelButton = $btnCancel

    if ($recDialog.ShowDialog() -eq [System.Windows.Forms.DialogResult]::OK) {
        $recommendations = @()
        foreach ($row in $dataGridView.Rows) {
            if ($row.IsNewRow) { continue }
            $product = $row.Cells["Product"].Value
            $recText = [string]$row.Cells["Recommendations"].Value
            
            if (-not [string]::IsNullOrWhiteSpace($recText)) {
                $recommendations += [PSCustomObject]@{
                    Product = $product
                    Recommendations = $recText.Trim()
                }
            }
        }
        return $recommendations
    }

    return $null
}

function Show-HostnameReviewDialog {
    param(
        [array]$Top10Data
    )

    # Create dialog
    $hostDialog = New-Object System.Windows.Forms.Form
    $hostDialog.Text = "Review Hostnames - Select Systems to Include"
    $hostDialog.Size = New-Object System.Drawing.Size(1100, 700)
    $hostDialog.StartPosition = "CenterParent"
    $hostDialog.FormBorderStyle = "FixedDialog"
    $hostDialog.MaximizeBox = $false
    $hostDialog.MinimizeBox = $false

    # Create TabControl
    $tabControl = New-Object System.Windows.Forms.TabControl
    $tabControl.Location = New-Object System.Drawing.Point(20, 20)
    $tabControl.Size = New-Object System.Drawing.Size(1050, 550)
    $hostDialog.Controls.Add($tabControl)

    # Create a copy of Top10Data to modify
    $filteredData = @()
    foreach ($item in $Top10Data) {
        $filteredData += [PSCustomObject]@{
            Product = $item.Product
            Critical = $item.Critical
            High = $item.High
            Medium = $item.Medium
            Low = $item.Low
            VulnCount = $item.VulnCount
            EPSSScore = $item.EPSSScore
            AvgCVSS = $item.AvgCVSS
            RiskScore = $item.RiskScore
            AffectedSystems = $item.AffectedSystems | ForEach-Object { $_ }  # Clone array
        }
    }

    # Create tab for each vulnerability item
    $tabDataGridViews = @()
    for ($i = 0; $i -lt $filteredData.Count; $i++) {
        $item = $filteredData[$i]
        $tabPage = New-Object System.Windows.Forms.TabPage
        $tabPage.Text = "$($i + 1). $($item.Product)"
        $tabPage.UseVisualStyleBackColor = $true
        $tabControl.TabPages.Add($tabPage)

        # Create DataGridView for this tab
        $dataGridView = New-Object System.Windows.Forms.DataGridView
        $dataGridView.Location = New-Object System.Drawing.Point(10, 10)
        $dataGridView.Size = New-Object System.Drawing.Size(1020, 480)
        $dataGridView.AutoSizeColumnsMode = [System.Windows.Forms.DataGridViewAutoSizeColumnsMode]::Fill
        $dataGridView.SelectionMode = [System.Windows.Forms.DataGridViewSelectionMode]::FullRowSelect
        $dataGridView.MultiSelect = $false
        $dataGridView.AllowUserToAddRows = $false
        $dataGridView.ReadOnly = $false
        $tabPage.Controls.Add($dataGridView)

        # Add columns
        $colInclude = New-Object System.Windows.Forms.DataGridViewCheckBoxColumn
        $colInclude.Name = "Include"
        $colInclude.HeaderText = "Include"
        $colInclude.Width = 60
        $dataGridView.Columns.Add($colInclude) | Out-Null

        $colHostname = New-Object System.Windows.Forms.DataGridViewTextBoxColumn
        $colHostname.Name = "Hostname"
        $colHostname.HeaderText = "Hostname"
        $colHostname.Width = 200
        $colHostname.ReadOnly = $true
        $dataGridView.Columns.Add($colHostname) | Out-Null

        $colIP = New-Object System.Windows.Forms.DataGridViewTextBoxColumn
        $colIP.Name = "IP"
        $colIP.HeaderText = "IP Address"
        $colIP.Width = 150
        $colIP.ReadOnly = $true
        $dataGridView.Columns.Add($colIP) | Out-Null

        $colUsername = New-Object System.Windows.Forms.DataGridViewTextBoxColumn
        $colUsername.Name = "Username"
        $colUsername.HeaderText = "Username"
        $colUsername.Width = 150
        $colUsername.ReadOnly = $true
        $dataGridView.Columns.Add($colUsername) | Out-Null

        $colVulnCount = New-Object System.Windows.Forms.DataGridViewTextBoxColumn
        $colVulnCount.Name = "VulnCount"
        $colVulnCount.HeaderText = "Vulnerability Count"
        $colVulnCount.Width = 150
        $colVulnCount.ReadOnly = $true
        $colVulnCount.ValueType = [int]
        $dataGridView.Columns.Add($colVulnCount) | Out-Null

        # Populate grid with hostnames (or IP fallback), sorted by vulnerability count descending
        $sortedSystems = $item.AffectedSystems | Sort-Object -Property VulnCount -Descending
        foreach ($sys in $sortedSystems) {
            $row = $dataGridView.Rows.Add()
            $dataGridView.Rows[$row].Cells["Include"].Value = $true  # Default checked
            $dataGridView.Rows[$row].Cells["Hostname"].Value = if ($sys.HostName) { $sys.HostName } else { $sys.IP }
            $dataGridView.Rows[$row].Cells["IP"].Value = $sys.IP
            $dataGridView.Rows[$row].Cells["Username"].Value = $sys.Username
            $dataGridView.Rows[$row].Cells["VulnCount"].Value = $sys.VulnCount
        }

        # Store reference to DataGridView for later access
        $tabDataGridViews += $dataGridView
    }

    # Summary label
    $lblSummary = New-Object System.Windows.Forms.Label
    $lblSummary.Location = New-Object System.Drawing.Point(20, 580)
    $lblSummary.Size = New-Object System.Drawing.Size(600, 20)
    $lblSummary.Text = ""
    $hostDialog.Controls.Add($lblSummary)

    # Function to update summary
    $updateSummary = {
        $totalSelected = 0
        $totalHostnames = 0
        foreach ($dgv in $tabDataGridViews) {
            foreach ($row in $dgv.Rows) {
                if ($row.IsNewRow) { continue }
                $totalHostnames++
                if ([bool]$row.Cells["Include"].Value) {
                    $totalSelected++
                }
            }
        }
        $lblSummary.Text = "Selected: $totalSelected of $totalHostnames hostnames"
    }

    # Add event handler to update summary when checkboxes change
    foreach ($dgv in $tabDataGridViews) {
        $dgv.Add_CellValueChanged($updateSummary)
    }

    # Initial summary update
    & $updateSummary

    # Buttons
    $y = 610

    $btnSelectAll = New-Object System.Windows.Forms.Button
    $btnSelectAll.Location = New-Object System.Drawing.Point(20, $y)
    $btnSelectAll.Size = New-Object System.Drawing.Size(100, 30)
    $btnSelectAll.Text = "Select All"
    $btnSelectAll.Add_Click({
        foreach ($dgv in $tabDataGridViews) {
            foreach ($row in $dgv.Rows) {
                if ($row.IsNewRow) { continue }
                $row.Cells["Include"].Value = $true
            }
        }
        & $updateSummary
    })
    $hostDialog.Controls.Add($btnSelectAll)

    $btnDeselectAll = New-Object System.Windows.Forms.Button
    $btnDeselectAll.Location = New-Object System.Drawing.Point(130, $y)
    $btnDeselectAll.Size = New-Object System.Drawing.Size(100, 30)
    $btnDeselectAll.Text = "Deselect All"
    $btnDeselectAll.Add_Click({
        foreach ($dgv in $tabDataGridViews) {
            foreach ($row in $dgv.Rows) {
                if ($row.IsNewRow) { continue }
                $row.Cells["Include"].Value = $false
            }
        }
        & $updateSummary
    })
    $hostDialog.Controls.Add($btnDeselectAll)

    $btnCancel = New-Object System.Windows.Forms.Button
    $btnCancel.Location = New-Object System.Drawing.Point(900, $y)
    $btnCancel.Size = New-Object System.Drawing.Size(90, 30)
    $btnCancel.Text = "Cancel"
    $btnCancel.DialogResult = [System.Windows.Forms.DialogResult]::Cancel
    $hostDialog.Controls.Add($btnCancel)

    $btnOK = New-Object System.Windows.Forms.Button
    $btnOK.Location = New-Object System.Drawing.Point(990, $y)
    $btnOK.Size = New-Object System.Drawing.Size(90, 30)
    $btnOK.Text = "OK"
    $btnOK.DialogResult = [System.Windows.Forms.DialogResult]::OK
    $hostDialog.Controls.Add($btnOK)

    # Set default button
    $hostDialog.AcceptButton = $btnOK
    $hostDialog.CancelButton = $btnCancel

    if ($hostDialog.ShowDialog() -eq [System.Windows.Forms.DialogResult]::OK) {
        # Filter AffectedSystems based on selections
        for ($i = 0; $i -lt $filteredData.Count; $i++) {
            $dgv = $tabDataGridViews[$i]
            $filteredSystems = @()
            
            foreach ($row in $dgv.Rows) {
                if ($row.IsNewRow) { continue }
                if ([bool]$row.Cells["Include"].Value) {
                    $hostname = $row.Cells["Hostname"].Value
                    $ip = $row.Cells["IP"].Value
                    $username = $row.Cells["Username"].Value
                    $vulnCount = [int]$row.Cells["VulnCount"].Value
                    
                    $filteredSystems += [PSCustomObject]@{
                        HostName = $hostname
                        IP = $ip
                        Username = $username
                        VulnCount = $vulnCount
                    }
                }
            }
            
            $filteredData[$i].AffectedSystems = $filteredSystems
            # Note: Count property is read-only and automatically reflects array length
        }
        
        return $filteredData
    }

    return $null
}

function Show-TimeEstimateEntryDialog {
    param(
        [array]$Top10Data,
        [bool]$IsRMITPlus
    )

    # Create dialog
    $timeDialog = New-Object System.Windows.Forms.Form
    $timeDialog.Text = "Enter Time Estimates"
    $timeDialog.Size = New-Object System.Drawing.Size(1050, 600)
    $timeDialog.StartPosition = "CenterParent"
    $timeDialog.FormBorderStyle = "FixedDialog"
    $timeDialog.MaximizeBox = $false
    $timeDialog.MinimizeBox = $false

    # Create DataGridView
    $dataGridView = New-Object System.Windows.Forms.DataGridView
    $dataGridView.Location = New-Object System.Drawing.Point(20, 20)
    $dataGridView.Size = New-Object System.Drawing.Size(950, 450)
    $dataGridView.AutoSizeColumnsMode = [System.Windows.Forms.DataGridViewAutoSizeColumnsMode]::Fill
    $dataGridView.SelectionMode = [System.Windows.Forms.DataGridViewSelectionMode]::FullRowSelect
    $dataGridView.MultiSelect = $false
    $dataGridView.AllowUserToAddRows = $false
    $timeDialog.Controls.Add($dataGridView)

    # Add columns
    $colProduct = New-Object System.Windows.Forms.DataGridViewTextBoxColumn
    $colProduct.Name = "Product"
    $colProduct.HeaderText = "Product"
    $colProduct.Width = 250
    $colProduct.ReadOnly = $true
    $dataGridView.Columns.Add($colProduct) | Out-Null

    $colHostnames = New-Object System.Windows.Forms.DataGridViewTextBoxColumn
    $colHostnames.Name = "Hostnames"
    $colHostnames.HeaderText = "Hostnames"
    $colHostnames.Width = 80
    $colHostnames.ReadOnly = $true
    $dataGridView.Columns.Add($colHostnames) | Out-Null

    $colTimeEstimate = New-Object System.Windows.Forms.DataGridViewTextBoxColumn
    $colTimeEstimate.Name = "TimeEstimate"
    $colTimeEstimate.HeaderText = "Time Estimate (hours)"
    $colTimeEstimate.Width = 150
    $dataGridView.Columns.Add($colTimeEstimate) | Out-Null

    $colAfterHours = New-Object System.Windows.Forms.DataGridViewCheckBoxColumn
    $colAfterHours.Name = "AfterHours"
    $colAfterHours.HeaderText = "After Hours"
    $colAfterHours.Width = 100
    $dataGridView.Columns.Add($colAfterHours) | Out-Null

    if ($IsRMITPlus) {
        $colThirdParty = New-Object System.Windows.Forms.DataGridViewCheckBoxColumn
        $colThirdParty.Name = "ThirdParty"
        $colThirdParty.HeaderText = "3rd Party"
        $colThirdParty.Width = 100
        $dataGridView.Columns.Add($colThirdParty) | Out-Null

        $colTicketGenerated = New-Object System.Windows.Forms.DataGridViewCheckBoxColumn
        $colTicketGenerated.Name = "TicketGenerated"
        $colTicketGenerated.HeaderText = "Ticket Generated"
        $colTicketGenerated.Width = 120
        $dataGridView.Columns.Add($colTicketGenerated) | Out-Null
    }

    # Populate grid with top 10 data
    foreach ($item in $Top10Data) {
        $row = $dataGridView.Rows.Add()
        $dataGridView.Rows[$row].Cells["Product"].Value = $item.Product
        $hostnameCount = if ($item.AffectedSystems) { $item.AffectedSystems.Count } else { 0 }
        $dataGridView.Rows[$row].Cells["Hostnames"].Value = $hostnameCount
        $dataGridView.Rows[$row].Cells["TimeEstimate"].Value = ""
        $dataGridView.Rows[$row].Cells["AfterHours"].Value = $false
        if ($IsRMITPlus) {
            # Default 3rd party status based on covered software list
            $isThirdPartyDefault = Test-IsCoveredSoftware -ProductName $item.Product
            $dataGridView.Rows[$row].Cells["ThirdParty"].Value = $isThirdPartyDefault
            $dataGridView.Rows[$row].Cells["TicketGenerated"].Value = $false
        }
    }

    # Buttons
    $y = 480

    $btnOK = New-Object System.Windows.Forms.Button
    $btnOK.Location = New-Object System.Drawing.Point(800, $y)
    $btnOK.Size = New-Object System.Drawing.Size(90, 30)
    $btnOK.Text = "OK"
    $btnOK.DialogResult = [System.Windows.Forms.DialogResult]::OK
    $timeDialog.Controls.Add($btnOK)
    $timeDialog.AcceptButton = $btnOK

    $btnCancel = New-Object System.Windows.Forms.Button
    $btnCancel.Location = New-Object System.Drawing.Point(890, $y)
    $btnCancel.Size = New-Object System.Drawing.Size(90, 30)
    $btnCancel.Text = "Cancel"
    $btnCancel.DialogResult = [System.Windows.Forms.DialogResult]::Cancel
    $timeDialog.Controls.Add($btnCancel)
    $timeDialog.CancelButton = $btnCancel

    # Validation on OK
    $btnOK.Add_Click({
        # Validate all time estimates are filled
        foreach ($row in $dataGridView.Rows) {
            if ($row.IsNewRow) { continue }
            $timeEstimate = [string]$row.Cells["TimeEstimate"].Value
            if ([string]::IsNullOrWhiteSpace($timeEstimate)) {
                [System.Windows.Forms.MessageBox]::Show("Please enter a time estimate for all items.", "Validation Error",
                    [System.Windows.Forms.MessageBoxButtons]::OK, [System.Windows.Forms.MessageBoxIcon]::Warning)
                $timeDialog.DialogResult = [System.Windows.Forms.DialogResult]::None
                return
            }
            # Validate it's a valid number
            $timeValue = 0
            if (-not [double]::TryParse($timeEstimate, [ref]$timeValue) -or $timeValue -lt 0) {
                [System.Windows.Forms.MessageBox]::Show("Time estimate must be a valid positive number (hours).", "Validation Error",
                    [System.Windows.Forms.MessageBoxButtons]::OK, [System.Windows.Forms.MessageBoxIcon]::Warning)
                $timeDialog.DialogResult = [System.Windows.Forms.DialogResult]::None
                return
            }
        }
    })

    $result = $timeDialog.ShowDialog()

    if ($result -eq [System.Windows.Forms.DialogResult]::OK) {
        # Build result array
        $timeEstimates = @()
        foreach ($row in $dataGridView.Rows) {
            if ($row.IsNewRow) { continue }
            $timeEstimates += [PSCustomObject]@{
                Product = $row.Cells["Product"].Value
                TimeEstimate = [double]$row.Cells["TimeEstimate"].Value
                AfterHours = [bool]$row.Cells["AfterHours"].Value
                ThirdParty = if ($IsRMITPlus) { [bool]$row.Cells["ThirdParty"].Value } else { $false }
                TicketGenerated = if ($IsRMITPlus) { [bool]$row.Cells["TicketGenerated"].Value } else { $false }
            }
        }
        return $timeEstimates
    } else {
        return $null
    }
}

function New-TimeEstimate {
    param(
        [string]$OutputPath,
        [array]$Top10Data,
        [array]$TimeEstimates,
        [bool]$IsRMITPlus,
        [array]$GeneralRecommendations = $null
    )

    try {
        $sb = New-Object System.Text.StringBuilder

        [void]$sb.AppendLine("=".PadRight(100, '='))
        [void]$sb.AppendLine("TIME ESTIMATE FOR VULNERABILITY REMEDIATION")
        [void]$sb.AppendLine("=".PadRight(100, '='))
        [void]$sb.AppendLine()

        [void]$sb.AppendLine("Vulnerability Time Estimates:")
        [void]$sb.AppendLine()

        $totalCovered = 0.0
        $totalRequiringApproval = 0.0
        $grandTotal = 0.0

        # Load covered software list if needed
        if ($null -eq $script:CoveredSoftware -or $script:CoveredSoftware.Count -eq 0) {
            Load-CoveredSoftware
        }

        $generalRecMap = @{}
        if ($null -ne $GeneralRecommendations -and $GeneralRecommendations.Count -gt 0) {
            foreach ($rec in $GeneralRecommendations) {
                if (-not [string]::IsNullOrWhiteSpace($rec.Product)) { $generalRecMap[$rec.Product] = $rec }
            }
        }
        for ($i = 0; $i -lt $Top10Data.Count; $i++) {
            $item = $Top10Data[$i]
            $timeEstimate = $TimeEstimates[$i]

            [void]$sb.AppendLine("$($i + 1). $($item.Product)")
            
            # Add hostnames/hosts CSV list (use hostname or IP so we include all systems)
            if ($item.AffectedSystems -and $item.AffectedSystems.Count -gt 0) {
                $identifiers = $item.AffectedSystems | ForEach-Object { if ($_.HostName) { $_.HostName } else { $_.IP } } | Where-Object { -not [string]::IsNullOrWhiteSpace($_) } | Select-Object -Unique
                if ($identifiers.Count -gt 0) {
                    $identifiersList = $identifiers -join ", "
                    [void]$sb.AppendLine("   Affected Hostnames: $identifiersList")
                }
            }
            
            # Add remediation guidance
            $remediationGuidance = Get-RemediationGuidance -ProductName $item.Product -OutputType 'Ticket'
            if (-not [string]::IsNullOrWhiteSpace($remediationGuidance)) {
                [void]$sb.AppendLine("   Remediation Guidance:")
                # Split by line breaks and indent each line
                $guidanceLines = $remediationGuidance -split "`r?`n"
                foreach ($line in $guidanceLines) {
                    if (-not [string]::IsNullOrWhiteSpace($line)) {
                        [void]$sb.AppendLine("     $line")
                    }
                }
            }
            
            $matchingRec = if ($generalRecMap.ContainsKey($item.Product)) { $generalRecMap[$item.Product] } else { $null }
            if ($null -ne $matchingRec -and -not [string]::IsNullOrWhiteSpace($matchingRec.Recommendations)) {
                [void]$sb.AppendLine("   General Recommendations:")
                $recLines = $matchingRec.Recommendations -split "`r?`n"
                foreach ($line in $recLines) {
                    if (-not [string]::IsNullOrWhiteSpace($line)) {
                        [void]$sb.AppendLine("     $line")
                    }
                }
            }

            if ($IsRMITPlus) {
                # Use the 3rd party status from the dialog checkbox
                $isThirdParty = $timeEstimate.ThirdParty
                
                # Auto-mark as ticket generated if 3rd party software AND after-hours
                # (user can't select ticket generated until after report is run)
                $autoTicketGenerated = $isThirdParty -and $timeEstimate.AfterHours
                $isTicketGenerated = $timeEstimate.TicketGenerated -or $autoTicketGenerated
                
                # Check if it requires approval (after-hours OR 3rd party software)
                $requiresApproval = $timeEstimate.AfterHours -or $isThirdParty
                
                # If ticket is already generated (manually or auto), it's covered regardless
                if ($isTicketGenerated) {
                    if ($autoTicketGenerated) {
                        [void]$sb.AppendLine("   After Hours: Yes")
                        [void]$sb.AppendLine("   Ticket Generated: Yes (Covered by Agreement - Auto-generated)")
                    } else {
                        [void]$sb.AppendLine("   Ticket Generated: Yes (Covered by Agreement)")
                    }
                    [void]$sb.AppendLine("   Estimated Time: $($timeEstimate.TimeEstimate) hours - A remediation ticket has already been generated")
                    [void]$sb.AppendLine("   Status: Covered by Agreement")
                    $totalCovered += $timeEstimate.TimeEstimate  # Add time to covered total
                } elseif (-not $requiresApproval) {
                    # Covered by agreement (non-3rd party software, not after-hours, ticket not generated)
                    [void]$sb.AppendLine("   Estimated Time: $($timeEstimate.TimeEstimate) hours - A remediation ticket has already been generated")
                    [void]$sb.AppendLine("   Status: Covered by Agreement")
                    $totalCovered += $timeEstimate.TimeEstimate  # Add time to covered total
                } else {
                    # Requires approval (3rd party software OR after-hours)
                    if ($timeEstimate.AfterHours) {
                        [void]$sb.AppendLine("   After Hours: Yes")
                        [void]$sb.AppendLine("   Estimated Time: N/A - A remediation ticket has already been generated")
                    } else {
                        [void]$sb.AppendLine("   Estimated Time: $($timeEstimate.TimeEstimate) hours")
                    }
                    
                    [void]$sb.AppendLine("   Status: Requires Approval")
                    if ($timeEstimate.AfterHours) {
                        # After-hours items don't count toward totals (ticket already generated)
                        $totalRequiringApproval += 0
                    } else {
                        # 3rd party software items require approval and count toward totals
                        $totalRequiringApproval += $timeEstimate.TimeEstimate
                        $grandTotal += $timeEstimate.TimeEstimate
                    }
                }
            } else {
                # RMIT/CMIT (hourly billing) - include all items in totals
                if ($timeEstimate.AfterHours) {
                    [void]$sb.AppendLine("   After Hours: Yes")
                    [void]$sb.AppendLine("   Estimated Time: $($timeEstimate.TimeEstimate) hours")
                } else {
                    [void]$sb.AppendLine("   Estimated Time: $($timeEstimate.TimeEstimate) hours")
                }
                # All items count toward grand total for RMIT/CMIT clients
                $grandTotal += $timeEstimate.TimeEstimate
            }
            [void]$sb.AppendLine()
        }

        [void]$sb.AppendLine()
        [void]$sb.AppendLine("=".PadRight(100, '='))
        [void]$sb.AppendLine("SUMMARY")
        [void]$sb.AppendLine("=".PadRight(100, '='))
        [void]$sb.AppendLine()

        if ($IsRMITPlus) {
            [void]$sb.AppendLine("Total Covered by Agreement: $totalCovered hours")
            [void]$sb.AppendLine("Total Requiring Approval: $totalRequiringApproval hours")
            [void]$sb.AppendLine()
        }

        [void]$sb.AppendLine("Grand Total: $grandTotal hours")
        
        if (-not $IsRMITPlus) {
            [void]$sb.AppendLine()
            [void]$sb.AppendLine("Note: We will not begin remediation without your prior approval.")
        }

        $sb.ToString() | Out-File -FilePath $OutputPath -Encoding UTF8
        Write-Log "Time estimate saved to: $OutputPath" -Level Success

    } catch {
        Write-Log "Error generating time estimate: $($_.Exception.Message)" -Level Error
        throw
    }
}

function New-TicketInstructions {
    param(
        [string]$OutputPath,
        [array]$TopTenData,
        [array]$TimeEstimates = $null,
        [bool]$IsRMITPlus = $false,
        [array]$GeneralRecommendations = $null
    )

    try {
        Write-Log "Generating ticket instructions..."

        $sb = New-Object System.Text.StringBuilder
        [void]$sb.AppendLine("=".PadRight(100, '='))
        [void]$sb.AppendLine("TOP 10 VULNERABILITIES - TICKET INSTRUCTIONS")
        [void]$sb.AppendLine("Generated: $(Get-Date -Format 'yyyy-MM-dd HH:mm:ss')")
        [void]$sb.AppendLine("=".PadRight(100, '='))
        [void]$sb.AppendLine()

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
        for ($i = 0; $i -lt $TopTenData.Count; $i++) {
            $item = $TopTenData[$i]
            $num = $i + 1

            $timeEstimate = if ($timeEstimateMap.ContainsKey($item.Product)) { $timeEstimateMap[$item.Product] } else { $null }

            # Generate ticket subject based on product type
            $ticketSubject = "Vulnerability Remediation - "
            if ($item.Product -like "*Windows Server 2012*" -or $item.Product -like "*end-of-life*" -or $item.Product -like "*out of support*") {
                $ticketSubject += "$($item.Product) - End of Support Migration Required"
            } elseif ($item.Product -like "*Windows 10*") {
                $ticketSubject += "$($item.Product) - Windows 10 is End of Life"
            } elseif ($item.Product -like "*Windows Server*") {
                $ticketSubject += "$($item.Product) - Updates Required"
            } elseif ($item.Product -like "*Windows*") {
                $ticketSubject += "$($item.Product) - Patch Management Required"
            } elseif ($item.Product -like "*printer*" -or $item.Product -like "*Ripple20*") {
                $ticketSubject += "$($item.Product) - Firmware Update Required"
            } elseif ($item.Product -like "*Microsoft Teams*") {
                $ticketSubject += "$($item.Product) - Application Update Required"
            } elseif ((Test-IsMicrosoftApplication -ProductName $item.Product) -and $IsRMITPlus) {
                $ticketSubject += "$($item.Product) - RMIT+ ticketed"
            } elseif ((Test-IsVMwareProduct -ProductName $item.Product) -and $IsRMITPlus) {
                $ticketSubject += "$($item.Product) - RMIT+ after-hours ticket created if we maintain this"
            } elseif (Test-IsAutoUpdatingSoftware -ProductName $item.Product) {
                $ticketSubject += "$($item.Product) - This software updates automatically"
            } else {
                $ticketSubject += "$($item.Product) - Update Required"
            }

            # Prepend "After Hours - " if after hours AND ticket generated
            if ($null -ne $timeEstimate) {
                $afterHours = $timeEstimate.AfterHours
                $ticketGenerated = if ($IsRMITPlus) { $timeEstimate.TicketGenerated } else { $false }
                if ($afterHours -and $ticketGenerated) {
                    $ticketSubject = "After Hours - " + $ticketSubject
                }
            }

            [void]$sb.AppendLine()
            [void]$sb.AppendLine("-".PadRight(100, '-'))
            [void]$sb.AppendLine("VULNERABILITY #$num")
            [void]$sb.AppendLine("-".PadRight(100, '-'))
            [void]$sb.AppendLine()
            [void]$sb.AppendLine("TICKET SUBJECT:")
            [void]$sb.AppendLine($ticketSubject)
            [void]$sb.AppendLine()
            [void]$sb.AppendLine(("Product/System:".PadRight(25)) + $item.Product)
            [void]$sb.AppendLine(("Risk Score:".PadRight(25)) + $item.RiskScore.ToString('N2'))
            [void]$sb.AppendLine(("EPSS Score:".PadRight(25)) + $item.EPSSScore.ToString('N4'))
            [void]$sb.AppendLine(("Average CVSS:".PadRight(25)) + $item.AvgCVSS.ToString('N2'))
            [void]$sb.AppendLine(("Total Vulnerabilities:".PadRight(25)) + $item.VulnCount)
            [void]$sb.AppendLine(("Affected Systems Count:".PadRight(25)) + $item.AffectedSystems.Count)
            [void]$sb.AppendLine()
            [void]$sb.AppendLine("NOTE: This remediation can go to any available technician.")
            [void]$sb.AppendLine()
            [void]$sb.AppendLine("Affected Systems:")
            # Group by hostname, IP, and username to get unique systems, then format as "hostname (username) - IP" or "hostname - IP"
            $uniqueSystems = $item.AffectedSystems | Select-Object HostName, IP, Username -Unique
            foreach ($sys in $uniqueSystems) {
                $hostname = $sys.HostName
                $ip = $sys.IP
                $username = $sys.Username
                # Use hostname or IP as primary identifier so we list all systems
                $systemLine = if ($hostname) { $hostname } else { $ip }
                if (-not [string]::IsNullOrWhiteSpace($username)) {
                    $systemLine += " ($username)"
                }
                if (-not [string]::IsNullOrWhiteSpace($ip)) {
                    $systemLine += " - $ip"
                }
                [void]$sb.AppendLine("  - $systemLine")
            }
            [void]$sb.AppendLine()
            [void]$sb.AppendLine("Remediation Instructions:")

            # Get remediation guidance from configurable rules
            $remediationText = Get-RemediationGuidance -ProductName $item.Product -OutputType 'Ticket'
            [void]$sb.AppendLine($remediationText)
            [void]$sb.AppendLine()

            # Add General Recommendations if available (use pre-built map for O(1) lookup)
            $matchingRec = if ($generalRecMap.ContainsKey($item.Product)) { $generalRecMap[$item.Product] } else { $null }
            if ($null -ne $matchingRec -and -not [string]::IsNullOrWhiteSpace($matchingRec.Recommendations)) {
                [void]$sb.AppendLine("General Recommendations:")
                [void]$sb.AppendLine($matchingRec.Recommendations)
                [void]$sb.AppendLine()
            }
        }

        [void]$sb.AppendLine()
        [void]$sb.AppendLine("=".PadRight(100, '='))
        [void]$sb.AppendLine("END OF TICKET INSTRUCTIONS")
        [void]$sb.AppendLine("=".PadRight(100, '='))

        $sb.ToString() | Out-File -FilePath $OutputPath -Encoding UTF8
        Write-Log "Ticket instructions saved to: $OutputPath" -Level Success

    } catch {
        Write-Log "Error generating ticket instructions: $($_.Exception.Message)" -Level Error
    }
}

function New-TicketNotes {
    param(
        [array]$Top10Data = $null,
        [array]$TimeEstimates = $null,
        [string]$OutputPath = $null,
        [bool]$IsRMITPlus = $false
    )

    # Use script variables if Top10Data not provided (backward compatibility)
    if ($null -eq $Top10Data) {
        $Top10Data = $script:CurrentTop10Data
        $TimeEstimates = $script:CurrentTimeEstimates
        $IsRMITPlus = $script:IsRMITPlus
    } elseif ($null -eq $TimeEstimates) {
        $TimeEstimates = $script:CurrentTimeEstimates
    }

    # Build steps performed list with ticket creation lines inserted in the correct position
    $stepsBeforeTickets = @"
- Examined lightweight agents
- Verified probe setup
- Checked agent/probe count compared to other systems
- Examined credential mappings
- Examined external assets
- Checked nmap interface on probe
- Verified deprecated item list
- Created all reports
- Assessed reports
- Produced top ten vulnerabilities docx report
"@

    # Collect ticket creation lines for vulnerabilities with tickets generated
    $ticketLines = @()
    if ($null -ne $Top10Data -and $null -ne $TimeEstimates -and $TimeEstimates.Count -gt 0) {
        $timeByProduct = @{}
        foreach ($te in $TimeEstimates) { if (-not [string]::IsNullOrWhiteSpace($te.Product)) { $timeByProduct[$te.Product] = $te } }
        foreach ($item in $Top10Data) {
            $timeEstimate = if ($timeByProduct.ContainsKey($item.Product)) { $timeByProduct[$item.Product] } else { $null }
            if ($null -ne $timeEstimate) {
                $ticketGenerated = if ($IsRMITPlus) { $timeEstimate.TicketGenerated } else { $false }
                if ($ticketGenerated) {
                    $ticketLines += "- Ticket created for $($item.Product)"
                }
            }
        }
    }

    $stepsAfterTickets = @"
- Sent secure email with reports to contact
- Sent TimeZest meeting request
"@

    # Combine steps with ticket lines inserted in the middle
    if ($ticketLines.Count -gt 0) {
        $stepsText = $stepsBeforeTickets + "`n" + ($ticketLines -join "`n") + "`n" + $stepsAfterTickets
    } else {
        $stepsText = $stepsBeforeTickets + "`n" + $stepsAfterTickets
    }

    # Build full ticket notes (no markdown formatting)
    $result = @"
Steps performed

$stepsText

Is the task resolved?

Yes - completed

Next step(s)

TimeZest meeting request has been sent. Please select a time to meet if you would like to discuss this further.
"@

    # Save to file if output path provided, otherwise copy to clipboard
    if (-not [string]::IsNullOrWhiteSpace($OutputPath)) {
        try {
            $result | Out-File -FilePath $OutputPath -Encoding UTF8
            $script:TicketNotesPath = $OutputPath
            [System.Windows.Forms.MessageBox]::Show("Ticket notes saved to:`n$OutputPath", "Success", [System.Windows.Forms.MessageBoxButtons]::OK, [System.Windows.Forms.MessageBoxIcon]::Information)
        } catch {
            [System.Windows.Forms.MessageBox]::Show("Failed to save ticket notes: $($_.Exception.Message)", "Error", [System.Windows.Forms.MessageBoxButtons]::OK, [System.Windows.Forms.MessageBoxIcon]::Error)
        }
    } else {
        # Copy to clipboard (fallback for when called from button without file path)
        try {
            [System.Windows.Forms.Clipboard]::SetText($result)
            [System.Windows.Forms.MessageBox]::Show("Ticket notes copied to clipboard!", "Success", [System.Windows.Forms.MessageBoxButtons]::OK, [System.Windows.Forms.MessageBoxIcon]::Information)
        } catch {
            [System.Windows.Forms.MessageBox]::Show("Failed to copy to clipboard: $($_.Exception.Message)", "Error", [System.Windows.Forms.MessageBoxButtons]::OK, [System.Windows.Forms.MessageBoxIcon]::Error)
        }
    }
}

# --- GUI Functions ---

