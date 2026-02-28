# VScanMagic-Dialogs.ps1 - UI dialogs (Filters, Output Options, Remediation Rules, etc.)
# Dot-sourced by VScanMagic-GUI.ps1
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

function Show-RemediationRulesDialog {
    # Load rules if not already loaded
    if ($null -eq $script:RemediationRules -or $script:RemediationRules.Count -eq 0) {
        Load-RemediationRules
    }

    # Create main dialog
    $rulesForm = New-Object System.Windows.Forms.Form
    $rulesForm.Text = "Remediation Rules Editor"
    $rulesForm.Size = New-Object System.Drawing.Size(900, 600)
    $rulesForm.StartPosition = "CenterParent"
    $rulesForm.FormBorderStyle = "FixedDialog"
    $rulesForm.MaximizeBox = $false
    $rulesForm.MinimizeBox = $false

    # Create DataGridView
    $dataGridView = New-Object System.Windows.Forms.DataGridView
    $dataGridView.Location = New-Object System.Drawing.Point(20, 20)
    $dataGridView.Size = New-Object System.Drawing.Size(840, 450)
    $dataGridView.AutoSizeColumnsMode = [System.Windows.Forms.DataGridViewAutoSizeColumnsMode]::Fill
    $dataGridView.SelectionMode = [System.Windows.Forms.DataGridViewSelectionMode]::FullRowSelect
    $dataGridView.MultiSelect = $false
    $dataGridView.ReadOnly = $true
    $dataGridView.AllowUserToAddRows = $false
    $rulesForm.Controls.Add($dataGridView)

    # Add columns
    $colPattern = New-Object System.Windows.Forms.DataGridViewTextBoxColumn
    $colPattern.Name = "Pattern"
    $colPattern.HeaderText = "Product Pattern"
    $colPattern.Width = 200
    $dataGridView.Columns.Add($colPattern) | Out-Null

    $colWordPreview = New-Object System.Windows.Forms.DataGridViewTextBoxColumn
    $colWordPreview.Name = "WordPreview"
    $colWordPreview.HeaderText = "Word Text Preview"
    $colWordPreview.Width = 300
    $dataGridView.Columns.Add($colWordPreview) | Out-Null

    $colTicketPreview = New-Object System.Windows.Forms.DataGridViewTextBoxColumn
    $colTicketPreview.Name = "TicketPreview"
    $colTicketPreview.HeaderText = "Ticket Text Preview"
    $colTicketPreview.Width = 300
    $dataGridView.Columns.Add($colTicketPreview) | Out-Null

    $colIsDefault = New-Object System.Windows.Forms.DataGridViewCheckBoxColumn
    $colIsDefault.Name = "IsDefault"
    $colIsDefault.HeaderText = "Default"
    $colIsDefault.Width = 60
    $dataGridView.Columns.Add($colIsDefault) | Out-Null

    # Function to refresh grid
    function Refresh-RulesGrid {
        $dataGridView.Rows.Clear()
        foreach ($rule in $script:RemediationRules) {
            $wordPreview = if ($rule.WordText.Length -gt 50) { $rule.WordText.Substring(0, 47) + "..." } else { $rule.WordText }
            $ticketPreview = if ($rule.TicketText.Length -gt 50) { $rule.TicketText.Substring(0, 47) + "..." } else { $rule.TicketText }
            $null = $dataGridView.Rows.Add($rule.Pattern, $wordPreview, $ticketPreview, $rule.IsDefault)
        }
    }

    # Function to show edit dialog
    function Show-EditRuleDialog {
        param(
            [int]$RuleIndex = -1
        )

        $isNew = $RuleIndex -lt 0
        $rule = if ($isNew) {
            @{
                Pattern = ""
                WordText = ""
                TicketText = ""
                IsDefault = $false
            }
        } else {
            $script:RemediationRules[$RuleIndex]
        }

        $editForm = New-Object System.Windows.Forms.Form
        $editForm.Text = if ($isNew) { "Add New Rule" } else { "Edit Rule" }
        $editForm.Size = New-Object System.Drawing.Size(700, 500)
        $editForm.StartPosition = "CenterParent"
        $editForm.FormBorderStyle = "FixedDialog"
        $editForm.MaximizeBox = $false
        $editForm.MinimizeBox = $false

        $y = 20

        # Pattern label and textbox
        $lblPattern = New-Object System.Windows.Forms.Label
        $lblPattern.Location = New-Object System.Drawing.Point(20, $y)
        $lblPattern.Size = New-Object System.Drawing.Size(200, 20)
        $lblPattern.Text = "Product Pattern (wildcard):"
        $editForm.Controls.Add($lblPattern)

        $txtPattern = New-Object System.Windows.Forms.TextBox
        $txtPattern.Location = New-Object System.Drawing.Point(20, ($y + 25))
        $txtPattern.Size = New-Object System.Drawing.Size(640, 20)
        $txtPattern.Text = $rule.Pattern
        $txtPattern.Enabled = -not $rule.IsDefault
        $editForm.Controls.Add($txtPattern)
        $y += 60

        # Word text label and textbox
        $lblWordText = New-Object System.Windows.Forms.Label
        $lblWordText.Location = New-Object System.Drawing.Point(20, $y)
        $lblWordText.Size = New-Object System.Drawing.Size(200, 20)
        $lblWordText.Text = "Word Report Remediation Text:"
        $editForm.Controls.Add($lblWordText)

        $txtWordText = New-Object System.Windows.Forms.TextBox
        $txtWordText.Location = New-Object System.Drawing.Point(20, ($y + 25))
        $txtWordText.Size = New-Object System.Drawing.Size(640, 120)
        $txtWordText.Multiline = $true
        $txtWordText.ScrollBars = "Vertical"
        $txtWordText.Text = $rule.WordText
        $editForm.Controls.Add($txtWordText)
        $y += 160

        # Ticket text label and textbox
        $lblTicketText = New-Object System.Windows.Forms.Label
        $lblTicketText.Location = New-Object System.Drawing.Point(20, $y)
        $lblTicketText.Size = New-Object System.Drawing.Size(200, 20)
        $lblTicketText.Text = "Ticket Instructions Remediation Text:"
        $editForm.Controls.Add($lblTicketText)

        $txtTicketText = New-Object System.Windows.Forms.TextBox
        $txtTicketText.Location = New-Object System.Drawing.Point(20, ($y + 25))
        $txtTicketText.Size = New-Object System.Drawing.Size(640, 120)
        $txtTicketText.Multiline = $true
        $txtTicketText.ScrollBars = "Vertical"
        $txtTicketText.Text = $rule.TicketText
        $editForm.Controls.Add($txtTicketText)
        $y += 160

        # IsDefault checkbox (only for new rules or if editing default)
        $chkIsDefault = New-Object System.Windows.Forms.CheckBox
        $chkIsDefault.Location = New-Object System.Drawing.Point(20, $y)
        $chkIsDefault.Size = New-Object System.Drawing.Size(300, 20)
        $chkIsDefault.Text = "This is the default rule (applies when no patterns match)"
        $chkIsDefault.Checked = $rule.IsDefault
        $chkIsDefault.Enabled = $isNew -or $rule.IsDefault
        $editForm.Controls.Add($chkIsDefault)
        $y += 40

        # Save button
        $btnSave = New-Object System.Windows.Forms.Button
        $btnSave.Location = New-Object System.Drawing.Point(480, $y)
        $btnSave.Size = New-Object System.Drawing.Size(90, 30)
        $btnSave.Text = "Save"
        $btnSave.Add_Click({
            if ([string]::IsNullOrWhiteSpace($txtPattern.Text) -and -not $chkIsDefault.Checked) {
                [System.Windows.Forms.MessageBox]::Show("Pattern cannot be empty (unless this is the default rule).", "Validation Error",
                    [System.Windows.Forms.MessageBoxButtons]::OK, [System.Windows.Forms.MessageBoxIcon]::Warning)
                return
            }

            if ([string]::IsNullOrWhiteSpace($txtWordText.Text)) {
                [System.Windows.Forms.MessageBox]::Show("Word remediation text cannot be empty.", "Validation Error",
                    [System.Windows.Forms.MessageBoxButtons]::OK, [System.Windows.Forms.MessageBoxIcon]::Warning)
                return
            }

            if ([string]::IsNullOrWhiteSpace($txtTicketText.Text)) {
                [System.Windows.Forms.MessageBox]::Show("Ticket remediation text cannot be empty.", "Validation Error",
                    [System.Windows.Forms.MessageBoxButtons]::OK, [System.Windows.Forms.MessageBoxIcon]::Warning)
                return
            }

            # If setting as default, ensure only one default exists
            if ($chkIsDefault.Checked) {
                for ($i = 0; $i -lt $script:RemediationRules.Count; $i++) {
                    if ($script:RemediationRules[$i].IsDefault -and ($isNew -or $i -ne $RuleIndex)) {
                        $script:RemediationRules[$i].IsDefault = $false
                    }
                }
            }

            if ($isNew) {
                $newRule = @{
                    Pattern = if ($chkIsDefault.Checked) { "*" } else { $txtPattern.Text }
                    WordText = $txtWordText.Text
                    TicketText = $txtTicketText.Text
                    IsDefault = $chkIsDefault.Checked
                }
                $script:RemediationRules += $newRule
            } else {
                # Update the rule at the specific index
                $script:RemediationRules[$RuleIndex].Pattern = if ($chkIsDefault.Checked) { "*" } else { $txtPattern.Text }
                $script:RemediationRules[$RuleIndex].WordText = $txtWordText.Text
                $script:RemediationRules[$RuleIndex].TicketText = $txtTicketText.Text
                $script:RemediationRules[$RuleIndex].IsDefault = $chkIsDefault.Checked
            }

            Refresh-RulesGrid
            $editForm.DialogResult = [System.Windows.Forms.DialogResult]::OK
            $editForm.Close()
        })
        $editForm.Controls.Add($btnSave)

        # Cancel button
        $btnCancel = New-Object System.Windows.Forms.Button
        $btnCancel.Location = New-Object System.Drawing.Point(580, $y)
        $btnCancel.Size = New-Object System.Drawing.Size(90, 30)
        $btnCancel.Text = "Cancel"
        $btnCancel.Add_Click({
            $editForm.DialogResult = [System.Windows.Forms.DialogResult]::Cancel
            $editForm.Close()
        })
        $editForm.Controls.Add($btnCancel)

        $editForm.ShowDialog() | Out-Null
    }

    # Populate grid
    Refresh-RulesGrid

    # Buttons
    $y = 480

    $btnAdd = New-Object System.Windows.Forms.Button
    $btnAdd.Location = New-Object System.Drawing.Point(20, $y)
    $btnAdd.Size = New-Object System.Drawing.Size(90, 30)
    $btnAdd.Text = "Add"
    $btnAdd.Add_Click({
        Show-EditRuleDialog -RuleIndex -1
    })
    $rulesForm.Controls.Add($btnAdd)

    $btnEdit = New-Object System.Windows.Forms.Button
    $btnEdit.Location = New-Object System.Drawing.Point(120, $y)
    $btnEdit.Size = New-Object System.Drawing.Size(90, 30)
    $btnEdit.Text = "Edit"
    $btnEdit.Add_Click({
        if ($dataGridView.SelectedRows.Count -gt 0) {
            $selectedIndex = $dataGridView.SelectedRows[0].Index
            Show-EditRuleDialog -RuleIndex $selectedIndex
        } else {
            [System.Windows.Forms.MessageBox]::Show("Please select a rule to edit.", "No Selection",
                [System.Windows.Forms.MessageBoxButtons]::OK, [System.Windows.Forms.MessageBoxIcon]::Information)
        }
    })
    $rulesForm.Controls.Add($btnEdit)

    $btnDelete = New-Object System.Windows.Forms.Button
    $btnDelete.Location = New-Object System.Drawing.Point(220, $y)
    $btnDelete.Size = New-Object System.Drawing.Size(90, 30)
    $btnDelete.Text = "Delete"
    $btnDelete.Add_Click({
        if ($dataGridView.SelectedRows.Count -gt 0) {
            $selectedIndex = $dataGridView.SelectedRows[0].Index
            $selectedRule = $script:RemediationRules[$selectedIndex]
            
            if ($selectedRule.IsDefault) {
                [System.Windows.Forms.MessageBox]::Show("Cannot delete the default rule. You must have at least one default rule.", "Cannot Delete",
                    [System.Windows.Forms.MessageBoxButtons]::OK, [System.Windows.Forms.MessageBoxIcon]::Warning)
                return
            }

            $result = [System.Windows.Forms.MessageBox]::Show("Are you sure you want to delete this rule?", "Confirm Delete",
                [System.Windows.Forms.MessageBoxButtons]::YesNo, [System.Windows.Forms.MessageBoxIcon]::Question)

            if ($result -eq [System.Windows.Forms.DialogResult]::Yes) {
                # Remove rule at the selected index
                $newRules = @()
                for ($i = 0; $i -lt $script:RemediationRules.Count; $i++) {
                    if ($i -ne $selectedIndex) {
                        $newRules += $script:RemediationRules[$i]
                    }
                }
                $script:RemediationRules = $newRules
                Refresh-RulesGrid
            }
        } else {
            [System.Windows.Forms.MessageBox]::Show("Please select a rule to delete.", "No Selection",
                [System.Windows.Forms.MessageBoxButtons]::OK, [System.Windows.Forms.MessageBoxIcon]::Information)
        }
    })
    $rulesForm.Controls.Add($btnDelete)

    $btnSave = New-Object System.Windows.Forms.Button
    $btnSave.Location = New-Object System.Drawing.Point(680, $y)
    $btnSave.Size = New-Object System.Drawing.Size(90, 30)
    $btnSave.Text = "Save"
    $btnSave.Add_Click({
        if (Save-RemediationRules) {
            [System.Windows.Forms.MessageBox]::Show("Remediation rules saved successfully!", "Success",
                [System.Windows.Forms.MessageBoxButtons]::OK, [System.Windows.Forms.MessageBoxIcon]::Information)
            $rulesForm.DialogResult = [System.Windows.Forms.DialogResult]::OK
            $rulesForm.Close()
        } else {
            [System.Windows.Forms.MessageBox]::Show("Failed to save remediation rules.", "Error",
                [System.Windows.Forms.MessageBoxButtons]::OK, [System.Windows.Forms.MessageBoxIcon]::Error)
        }
    })
    $rulesForm.Controls.Add($btnSave)

    $btnCancel = New-Object System.Windows.Forms.Button
    $btnCancel.Location = New-Object System.Drawing.Point(780, $y)
    $btnCancel.Size = New-Object System.Drawing.Size(90, 30)
    $btnCancel.Text = "Cancel"
    $btnCancel.Add_Click({
        # Reload rules to discard changes
        Load-RemediationRules
        $rulesForm.DialogResult = [System.Windows.Forms.DialogResult]::Cancel
        $rulesForm.Close()
    })
    $rulesForm.Controls.Add($btnCancel)

    $rulesForm.ShowDialog() | Out-Null
}

function Show-ConnectSecureSettingsDialog {
    # ConnectSecure API credentials settings window
    $dlg = New-Object System.Windows.Forms.Form
    $dlg.Text = "ConnectSecure API Settings"
    $dlg.Size = New-Object System.Drawing.Size(520, 340)
    $dlg.StartPosition = "CenterParent"
    $dlg.FormBorderStyle = "FixedDialog"
    $dlg.MaximizeBox = $false
    $dlg.MinimizeBox = $false

    $y = 20
    $lblBaseUrl = New-Object System.Windows.Forms.Label
    $lblBaseUrl.Location = New-Object System.Drawing.Point(20, $y)
    $lblBaseUrl.Size = New-Object System.Drawing.Size(120, 20)
    $lblBaseUrl.Text = "Base URL:"
    $dlg.Controls.Add($lblBaseUrl)
    $txtBaseUrl = New-Object System.Windows.Forms.TextBox
    $txtBaseUrl.Location = New-Object System.Drawing.Point(140, $y)
    $txtBaseUrl.Size = New-Object System.Drawing.Size(340, 20)
    $txtBaseUrl.Text = "https://pod104.myconnectsecure.com"
    $dlg.Controls.Add($txtBaseUrl)
    $y += 35

    $lblTenant = New-Object System.Windows.Forms.Label
    $lblTenant.Location = New-Object System.Drawing.Point(20, $y)
    $lblTenant.Size = New-Object System.Drawing.Size(120, 20)
    $lblTenant.Text = "Tenant Name:"
    $dlg.Controls.Add($lblTenant)
    $txtTenant = New-Object System.Windows.Forms.TextBox
    $txtTenant.Location = New-Object System.Drawing.Point(140, $y)
    $txtTenant.Size = New-Object System.Drawing.Size(200, 20)
    $txtTenant.Text = "river-run"
    $dlg.Controls.Add($txtTenant)
    $y += 35

    $lblClientId = New-Object System.Windows.Forms.Label
    $lblClientId.Location = New-Object System.Drawing.Point(20, $y)
    $lblClientId.Size = New-Object System.Drawing.Size(120, 20)
    $lblClientId.Text = "Client ID:"
    $dlg.Controls.Add($lblClientId)
    $txtClientId = New-Object System.Windows.Forms.TextBox
    $txtClientId.Location = New-Object System.Drawing.Point(140, $y)
    $txtClientId.Size = New-Object System.Drawing.Size(200, 20)
    $dlg.Controls.Add($txtClientId)
    $y += 35

    $lblClientSecret = New-Object System.Windows.Forms.Label
    $lblClientSecret.Location = New-Object System.Drawing.Point(20, $y)
    $lblClientSecret.Size = New-Object System.Drawing.Size(120, 20)
    $lblClientSecret.Text = "Client Secret:"
    $dlg.Controls.Add($lblClientSecret)
    $txtClientSecret = New-Object System.Windows.Forms.TextBox
    $txtClientSecret.Location = New-Object System.Drawing.Point(140, $y)
    $txtClientSecret.Size = New-Object System.Drawing.Size(200, 20)
    $txtClientSecret.PasswordChar = '*'
    $dlg.Controls.Add($txtClientSecret)
    $y += 40

    $btnLoad = New-Object System.Windows.Forms.Button
    $btnLoad.Location = New-Object System.Drawing.Point(20, $y)
    $btnLoad.Size = New-Object System.Drawing.Size(90, 28)
    $btnLoad.Text = "Load Saved"
    $btnLoad.Add_Click({
        $saved = Load-ConnectSecureCredentials
        if ($saved) {
            $txtBaseUrl.Text = $saved.BaseUrl
            $txtTenant.Text = $saved.TenantName
            $txtClientId.Text = $saved.ClientId
            $txtClientSecret.Text = $saved.ClientSecret
            [System.Windows.Forms.MessageBox]::Show("Credentials loaded.", "Loaded", [System.Windows.Forms.MessageBoxButtons]::OK, [System.Windows.Forms.MessageBoxIcon]::Information)
        } else {
            [System.Windows.Forms.MessageBox]::Show("No saved credentials found.", "Not Found", [System.Windows.Forms.MessageBoxButtons]::OK, [System.Windows.Forms.MessageBoxIcon]::Information)
        }
    })
    $dlg.Controls.Add($btnLoad)

    $btnSave = New-Object System.Windows.Forms.Button
    $btnSave.Location = New-Object System.Drawing.Point(120, $y)
    $btnSave.Size = New-Object System.Drawing.Size(90, 28)
    $btnSave.Text = "Save"
    $btnSave.BackColor = [System.Drawing.Color]::FromArgb(0, 150, 136)
    $btnSave.ForeColor = [System.Drawing.Color]::White
    $btnSave.FlatStyle = [System.Windows.Forms.FlatStyle]::Flat
    $btnSave.FlatAppearance.BorderSize = 0
    $btnSave.Add_Click({
        if ([string]::IsNullOrWhiteSpace($txtBaseUrl.Text) -or [string]::IsNullOrWhiteSpace($txtTenant.Text) -or [string]::IsNullOrWhiteSpace($txtClientId.Text) -or [string]::IsNullOrWhiteSpace($txtClientSecret.Text)) {
            [System.Windows.Forms.MessageBox]::Show("Please fill in all fields before saving.", "Validation", [System.Windows.Forms.MessageBoxButtons]::OK, [System.Windows.Forms.MessageBoxIcon]::Warning)
            return
        }
        if (Save-ConnectSecureCredentials -BaseUrl $txtBaseUrl.Text.Trim() -TenantName $txtTenant.Text.Trim() -ClientId $txtClientId.Text.Trim() -ClientSecret $txtClientSecret.Text.Trim() -CompanyId 0) {
            [System.Windows.Forms.MessageBox]::Show("Credentials saved.", "Saved", [System.Windows.Forms.MessageBoxButtons]::OK, [System.Windows.Forms.MessageBoxIcon]::Information)
        } else {
            [System.Windows.Forms.MessageBox]::Show("Failed to save.", "Error", [System.Windows.Forms.MessageBoxButtons]::OK, [System.Windows.Forms.MessageBoxIcon]::Error)
        }
    })
    $dlg.Controls.Add($btnSave)

    $btnClear = New-Object System.Windows.Forms.Button
    $btnClear.Location = New-Object System.Drawing.Point(220, $y)
    $btnClear.Size = New-Object System.Drawing.Size(90, 28)
    $btnClear.Text = "Clear"
    $btnClear.Add_Click({
        $txtBaseUrl.Text = "https://pod104.myconnectsecure.com"
        $txtTenant.Text = "river-run"
        $txtClientId.Text = ""
        $txtClientSecret.Text = ""
    })
    $dlg.Controls.Add($btnClear)

    $btnTest = New-Object System.Windows.Forms.Button
    $btnTest.Location = New-Object System.Drawing.Point(320, $y)
    $btnTest.Size = New-Object System.Drawing.Size(110, 28)
    $btnTest.Text = "Test Credentials"
    $btnTest.BackColor = [System.Drawing.Color]::FromArgb(33, 150, 243)
    $btnTest.ForeColor = [System.Drawing.Color]::White
    $btnTest.FlatStyle = [System.Windows.Forms.FlatStyle]::Flat
    $btnTest.FlatAppearance.BorderSize = 0
    $btnTest.Add_Click({
        if ([string]::IsNullOrWhiteSpace($txtBaseUrl.Text) -or [string]::IsNullOrWhiteSpace($txtTenant.Text) -or [string]::IsNullOrWhiteSpace($txtClientId.Text) -or [string]::IsNullOrWhiteSpace($txtClientSecret.Text)) {
            [System.Windows.Forms.MessageBox]::Show("Please fill in all fields first.", "Validation", [System.Windows.Forms.MessageBoxButtons]::OK, [System.Windows.Forms.MessageBoxIcon]::Warning)
            return
        }
        $btnTest.Enabled = $false
        $btnTest.Text = "Testing..."
        $dlg.Refresh()
        try {
            $connected = Connect-ConnectSecureAPI -BaseUrl $txtBaseUrl.Text -TenantName $txtTenant.Text -ClientId $txtClientId.Text -ClientSecret $txtClientSecret.Text
            if ($connected) {
                [System.Windows.Forms.MessageBox]::Show("Credentials are valid!", "Success", [System.Windows.Forms.MessageBoxButtons]::OK, [System.Windows.Forms.MessageBoxIcon]::Information)
            } else {
                [System.Windows.Forms.MessageBox]::Show("Authentication failed. Check Base URL, Tenant, Client ID, and Client Secret.", "Failed", [System.Windows.Forms.MessageBoxButtons]::OK, [System.Windows.Forms.MessageBoxIcon]::Warning)
            }
        } catch {
            [System.Windows.Forms.MessageBox]::Show("Error: $($_.Exception.Message)", "Error", [System.Windows.Forms.MessageBoxButtons]::OK, [System.Windows.Forms.MessageBoxIcon]::Error)
        } finally {
            $btnTest.Enabled = $true
            $btnTest.Text = "Test Credentials"
        }
    })
    $dlg.Controls.Add($btnTest)
    $y += 45

    $btnClose = New-Object System.Windows.Forms.Button
    $btnClose.Location = New-Object System.Drawing.Point(380, $y)
    $btnClose.Size = New-Object System.Drawing.Size(100, 30)
    $btnClose.Text = "Close"
    $btnClose.DialogResult = [System.Windows.Forms.DialogResult]::OK
    $dlg.Controls.Add($btnClose)
    $dlg.AcceptButton = $btnClose

    # Load saved credentials on open
    $saved = Load-ConnectSecureCredentials
    if ($saved) {
        $txtBaseUrl.Text = $saved.BaseUrl
        $txtTenant.Text = $saved.TenantName
        $txtClientId.Text = $saved.ClientId
        $txtClientSecret.Text = $saved.ClientSecret
    }

    $dlg.ShowDialog() | Out-Null
}

function Show-FiltersDialog {
    $dlg = New-Object System.Windows.Forms.Form
    $dlg.Text = "Report Filters"
    $dlg.Size = New-Object System.Drawing.Size(480, 220)
    $dlg.StartPosition = "CenterParent"
    $dlg.FormBorderStyle = "FixedDialog"
    $dlg.MaximizeBox = $false
    $dlg.MinimizeBox = $false

    $y = 20
    $lblMinEPSS = New-Object System.Windows.Forms.Label
    $lblMinEPSS.Location = New-Object System.Drawing.Point(20, $y)
    $lblMinEPSS.Size = New-Object System.Drawing.Size(150, 20)
    $lblMinEPSS.Text = "Minimum EPSS Score:"
    $dlg.Controls.Add($lblMinEPSS)
    $numMinEPSS = New-Object System.Windows.Forms.NumericUpDown
    $numMinEPSS.Location = New-Object System.Drawing.Point(180, ($y - 2))
    $numMinEPSS.Size = New-Object System.Drawing.Size(80, 20)
    $numMinEPSS.Minimum = 0
    $numMinEPSS.Maximum = 1
    $numMinEPSS.DecimalPlaces = 3
    $numMinEPSS.Increment = 0.001
    $numMinEPSS.Value = $script:FilterMinEPSS
    $dlg.Controls.Add($numMinEPSS)
    $y += 40

    $lblSeverity = New-Object System.Windows.Forms.Label
    $lblSeverity.Location = New-Object System.Drawing.Point(20, $y)
    $lblSeverity.Size = New-Object System.Drawing.Size(100, 20)
    $lblSeverity.Text = "Include Severities:"
    $dlg.Controls.Add($lblSeverity)
    $chkCritical = New-Object System.Windows.Forms.CheckBox
    $chkCritical.Location = New-Object System.Drawing.Point(135, $y)
    $chkCritical.Size = New-Object System.Drawing.Size(65, 20)
    $chkCritical.Text = "Critical"
    $chkCritical.Checked = $script:FilterIncludeCritical
    $dlg.Controls.Add($chkCritical)
    $chkHigh = New-Object System.Windows.Forms.CheckBox
    $chkHigh.Location = New-Object System.Drawing.Point(205, $y)
    $chkHigh.Size = New-Object System.Drawing.Size(55, 20)
    $chkHigh.Text = "High"
    $chkHigh.Checked = $script:FilterIncludeHigh
    $dlg.Controls.Add($chkHigh)
    $chkMedium = New-Object System.Windows.Forms.CheckBox
    $chkMedium.Location = New-Object System.Drawing.Point(265, $y)
    $chkMedium.Size = New-Object System.Drawing.Size(70, 20)
    $chkMedium.Text = "Medium"
    $chkMedium.Checked = $script:FilterIncludeMedium
    $dlg.Controls.Add($chkMedium)
    $chkLow = New-Object System.Windows.Forms.CheckBox
    $chkLow.Location = New-Object System.Drawing.Point(340, $y)
    $chkLow.Size = New-Object System.Drawing.Size(55, 20)
    $chkLow.Text = "Low"
    $chkLow.Checked = $script:FilterIncludeLow
    $dlg.Controls.Add($chkLow)
    $y += 35

    $lblTopN = New-Object System.Windows.Forms.Label
    $lblTopN.Location = New-Object System.Drawing.Point(20, $y)
    $lblTopN.Size = New-Object System.Drawing.Size(80, 20)
    $lblTopN.Text = "Top N:"
    $dlg.Controls.Add($lblTopN)
    $comboTopN = New-Object System.Windows.Forms.ComboBox
    $comboTopN.Location = New-Object System.Drawing.Point(100, ($y - 2))
    $comboTopN.Size = New-Object System.Drawing.Size(100, 25)
    $comboTopN.DropDownStyle = [System.Windows.Forms.ComboBoxStyle]::DropDownList
    @("10", "20", "50", "100", "All") | ForEach-Object { [void]$comboTopN.Items.Add($_) }
    $idx = 0; foreach ($i in 0..4) { if ($comboTopN.Items[$i] -eq $script:FilterTopN) { $idx = $i; break } }; $comboTopN.SelectedIndex = $idx
    $dlg.Controls.Add($comboTopN)
    $y += 45

    $btnOK = New-Object System.Windows.Forms.Button
    $btnOK.Location = New-Object System.Drawing.Point(280, $y)
    $btnOK.Size = New-Object System.Drawing.Size(90, 28)
    $btnOK.Text = "OK"
    $btnOK.DialogResult = [System.Windows.Forms.DialogResult]::OK
    $btnOK.Add_Click({
        $script:FilterMinEPSS = [double]$numMinEPSS.Value
        $script:FilterIncludeCritical = $chkCritical.Checked
        $script:FilterIncludeHigh = $chkHigh.Checked
        $script:FilterIncludeMedium = $chkMedium.Checked
        $script:FilterIncludeLow = $chkLow.Checked
        $script:FilterTopN = $comboTopN.SelectedItem.ToString()
    })
    $dlg.Controls.Add($btnOK)
    $btnCancel = New-Object System.Windows.Forms.Button
    $btnCancel.Location = New-Object System.Drawing.Point(375, $y)
    $btnCancel.Size = New-Object System.Drawing.Size(90, 28)
    $btnCancel.Text = "Cancel"
    $btnCancel.DialogResult = [System.Windows.Forms.DialogResult]::Cancel
    $dlg.Controls.Add($btnCancel)
    $dlg.AcceptButton = $btnOK
    $dlg.CancelButton = $btnCancel

    if ($dlg.ShowDialog() -eq [System.Windows.Forms.DialogResult]::OK) {
        $script:FilterMinEPSS = [double]$numMinEPSS.Value
        $script:FilterIncludeCritical = $chkCritical.Checked
        $script:FilterIncludeHigh = $chkHigh.Checked
        $script:FilterIncludeMedium = $chkMedium.Checked
        $script:FilterIncludeLow = $chkLow.Checked
        $script:FilterTopN = $comboTopN.SelectedItem.ToString()
    }
}

function Show-OutputOptionsDialog {
    $dlg = New-Object System.Windows.Forms.Form
    $dlg.Text = "Output Options"
    $dlg.Size = New-Object System.Drawing.Size(420, 250)
    $dlg.StartPosition = "CenterParent"
    $dlg.FormBorderStyle = "FixedDialog"
    $dlg.MaximizeBox = $false
    $dlg.MinimizeBox = $false

    $chkExcel = New-Object System.Windows.Forms.CheckBox
    $chkExcel.Location = New-Object System.Drawing.Point(20, 25)
    $chkExcel.Size = New-Object System.Drawing.Size(360, 20)
    $chkExcel.Text = "Generate Pending EPSS Report (Excel)"
    $chkExcel.Checked = $script:OutputExcel
    $dlg.Controls.Add($chkExcel)

    $chkWord = New-Object System.Windows.Forms.CheckBox
    $chkWord.Location = New-Object System.Drawing.Point(20, 50)
    $chkWord.Size = New-Object System.Drawing.Size(360, 20)
    $chkWord.Text = "Generate Top Vulnerabilities Report (Word)"
    $chkWord.Checked = $script:OutputWord
    $dlg.Controls.Add($chkWord)

    $chkEmail = New-Object System.Windows.Forms.CheckBox
    $chkEmail.Location = New-Object System.Drawing.Point(20, 75)
    $chkEmail.Size = New-Object System.Drawing.Size(360, 20)
    $chkEmail.Text = "Generate Email Template (Text)"
    $chkEmail.Checked = $script:OutputEmailTemplate
    $dlg.Controls.Add($chkEmail)

    $chkTicket = New-Object System.Windows.Forms.CheckBox
    $chkTicket.Location = New-Object System.Drawing.Point(20, 100)
    $chkTicket.Size = New-Object System.Drawing.Size(360, 20)
    $chkTicket.Text = "Generate Ticket Instructions (Text)"
    $chkTicket.Checked = $script:OutputTicketInstructions
    $dlg.Controls.Add($chkTicket)

    $chkTime = New-Object System.Windows.Forms.CheckBox
    $chkTime.Location = New-Object System.Drawing.Point(20, 125)
    $chkTime.Size = New-Object System.Drawing.Size(360, 20)
    $chkTime.Text = "Generate Time Estimate (Text)"
    $chkTime.Checked = $script:OutputTimeEstimate
    $dlg.Controls.Add($chkTime)

    $btnOK = New-Object System.Windows.Forms.Button
    $btnOK.Location = New-Object System.Drawing.Point(210, 165)
    $btnOK.Size = New-Object System.Drawing.Size(90, 28)
    $btnOK.Text = "OK"
    $btnOK.DialogResult = [System.Windows.Forms.DialogResult]::OK
    $btnOK.Add_Click({
        $script:OutputExcel = $chkExcel.Checked
        $script:OutputWord = $chkWord.Checked
        $script:OutputEmailTemplate = $chkEmail.Checked
        $script:OutputTicketInstructions = $chkTicket.Checked
        $script:OutputTimeEstimate = $chkTime.Checked
    })
    $dlg.Controls.Add($btnOK)
    $btnCancel = New-Object System.Windows.Forms.Button
    $btnCancel.Location = New-Object System.Drawing.Point(305, 165)
    $btnCancel.Size = New-Object System.Drawing.Size(90, 28)
    $btnCancel.Text = "Cancel"
    $btnCancel.DialogResult = [System.Windows.Forms.DialogResult]::Cancel
    $dlg.Controls.Add($btnCancel)
    $dlg.AcceptButton = $btnOK
    $dlg.CancelButton = $btnCancel

    if ($dlg.ShowDialog() -eq [System.Windows.Forms.DialogResult]::OK) {
        $script:OutputExcel = $chkExcel.Checked
        $script:OutputWord = $chkWord.Checked
        $script:OutputEmailTemplate = $chkEmail.Checked
        $script:OutputTicketInstructions = $chkTicket.Checked
        $script:OutputTimeEstimate = $chkTime.Checked
    }
}

function Show-SettingsDialog {
    $settingsForm = New-Object System.Windows.Forms.Form
    $settingsForm.Text = "User Settings"
    $settingsForm.Size = New-Object System.Drawing.Size(500, 500)
    $settingsForm.StartPosition = "CenterParent"
    $settingsForm.FormBorderStyle = "FixedDialog"
    $settingsForm.MaximizeBox = $false
    $settingsForm.MinimizeBox = $false

    $y = 20

    # Prepared By
    $lblPreparedBy = New-Object System.Windows.Forms.Label
    $lblPreparedBy.Location = New-Object System.Drawing.Point(20, $y)
    $lblPreparedBy.Size = New-Object System.Drawing.Size(150, 20)
    $lblPreparedBy.Text = "Prepared By:"
    $settingsForm.Controls.Add($lblPreparedBy)

    $txtPreparedBy = New-Object System.Windows.Forms.TextBox
    $txtPreparedBy.Location = New-Object System.Drawing.Point(180, $y)
    $txtPreparedBy.Size = New-Object System.Drawing.Size(280, 20)
    $txtPreparedBy.Text = $script:UserSettings.PreparedBy
    $settingsForm.Controls.Add($txtPreparedBy)
    $y += 35

    # Company Name
    $lblCompanyName = New-Object System.Windows.Forms.Label
    $lblCompanyName.Location = New-Object System.Drawing.Point(20, $y)
    $lblCompanyName.Size = New-Object System.Drawing.Size(150, 20)
    $lblCompanyName.Text = "Company Name:"
    $settingsForm.Controls.Add($lblCompanyName)

    $txtCompanyName = New-Object System.Windows.Forms.TextBox
    $txtCompanyName.Location = New-Object System.Drawing.Point(180, $y)
    $txtCompanyName.Size = New-Object System.Drawing.Size(280, 20)
    $txtCompanyName.Text = $script:UserSettings.CompanyName
    $settingsForm.Controls.Add($txtCompanyName)
    $y += 35

    # Company Address
    $lblCompanyAddress = New-Object System.Windows.Forms.Label
    $lblCompanyAddress.Location = New-Object System.Drawing.Point(20, $y)
    $lblCompanyAddress.Size = New-Object System.Drawing.Size(150, 20)
    $lblCompanyAddress.Text = "Company Address:"
    $settingsForm.Controls.Add($lblCompanyAddress)

    $txtCompanyAddress = New-Object System.Windows.Forms.TextBox
    $txtCompanyAddress.Location = New-Object System.Drawing.Point(180, $y)
    $txtCompanyAddress.Size = New-Object System.Drawing.Size(280, 40)
    $txtCompanyAddress.Multiline = $true
    $txtCompanyAddress.Text = $script:UserSettings.CompanyAddress
    $settingsForm.Controls.Add($txtCompanyAddress)
    $y += 55

    # Email
    $lblEmail = New-Object System.Windows.Forms.Label
    $lblEmail.Location = New-Object System.Drawing.Point(20, $y)
    $lblEmail.Size = New-Object System.Drawing.Size(150, 20)
    $lblEmail.Text = "Email:"
    $settingsForm.Controls.Add($lblEmail)

    $txtEmail = New-Object System.Windows.Forms.TextBox
    $txtEmail.Location = New-Object System.Drawing.Point(180, $y)
    $txtEmail.Size = New-Object System.Drawing.Size(280, 20)
    $txtEmail.Text = $script:UserSettings.Email
    $settingsForm.Controls.Add($txtEmail)
    $y += 35

    # Phone Number
    $lblPhoneNumber = New-Object System.Windows.Forms.Label
    $lblPhoneNumber.Location = New-Object System.Drawing.Point(20, $y)
    $lblPhoneNumber.Size = New-Object System.Drawing.Size(150, 20)
    $lblPhoneNumber.Text = "Phone Number:"
    $settingsForm.Controls.Add($lblPhoneNumber)

    $txtPhoneNumber = New-Object System.Windows.Forms.TextBox
    $txtPhoneNumber.Location = New-Object System.Drawing.Point(180, $y)
    $txtPhoneNumber.Size = New-Object System.Drawing.Size(280, 20)
    $txtPhoneNumber.Text = $script:UserSettings.PhoneNumber
    $settingsForm.Controls.Add($txtPhoneNumber)
    $y += 35

    # Company Phone Number
    $lblCompanyPhoneNumber = New-Object System.Windows.Forms.Label
    $lblCompanyPhoneNumber.Location = New-Object System.Drawing.Point(20, $y)
    $lblCompanyPhoneNumber.Size = New-Object System.Drawing.Size(150, 20)
    $lblCompanyPhoneNumber.Text = "Company Phone:"
    $settingsForm.Controls.Add($lblCompanyPhoneNumber)

    $txtCompanyPhoneNumber = New-Object System.Windows.Forms.TextBox
    $txtCompanyPhoneNumber.Location = New-Object System.Drawing.Point(180, $y)
    $txtCompanyPhoneNumber.Size = New-Object System.Drawing.Size(280, 20)
    $txtCompanyPhoneNumber.Text = $script:UserSettings.CompanyPhoneNumber
    $settingsForm.Controls.Add($txtCompanyPhoneNumber)
    $y += 35

    # Settings Directory
    $lblSettingsDirectory = New-Object System.Windows.Forms.Label
    $lblSettingsDirectory.Location = New-Object System.Drawing.Point(20, $y)
    $lblSettingsDirectory.Size = New-Object System.Drawing.Size(150, 20)
    $lblSettingsDirectory.Text = "Settings Directory:"
    $settingsForm.Controls.Add($lblSettingsDirectory)

    $txtSettingsDirectory = New-Object System.Windows.Forms.TextBox
    $txtSettingsDirectory.Location = New-Object System.Drawing.Point(180, $y)
    $txtSettingsDirectory.Size = New-Object System.Drawing.Size(200, 20)
    $txtSettingsDirectory.ReadOnly = $true
    $displayDir = if ([string]::IsNullOrEmpty($script:UserSettings.SettingsDirectory)) {
        Join-Path $env:LOCALAPPDATA "VScanMagic"
    } else {
        $script:UserSettings.SettingsDirectory
    }
    $txtSettingsDirectory.Text = $displayDir
    $settingsForm.Controls.Add($txtSettingsDirectory)

    $btnBrowseSettingsDir = New-Object System.Windows.Forms.Button
    $btnBrowseSettingsDir.Location = New-Object System.Drawing.Point(390, ($y - 2))
    $btnBrowseSettingsDir.Size = New-Object System.Drawing.Size(70, 25)
    $btnBrowseSettingsDir.Text = "Browse..."
    $btnBrowseSettingsDir.Add_Click({
        $folderBrowser = New-Object System.Windows.Forms.FolderBrowserDialog
        $folderBrowser.Description = "Select directory for settings and rules configuration files"
        $folderBrowser.ShowNewFolderButton = $true
        if (-not [string]::IsNullOrEmpty($txtSettingsDirectory.Text)) {
            $folderBrowser.SelectedPath = $txtSettingsDirectory.Text
        }
        if ($folderBrowser.ShowDialog() -eq [System.Windows.Forms.DialogResult]::OK) {
            $txtSettingsDirectory.Text = $folderBrowser.SelectedPath
        }
    })
    $settingsForm.Controls.Add($btnBrowseSettingsDir)
    $y += 35

    # Reset to Default button
    $btnResetDir = New-Object System.Windows.Forms.Button
    $btnResetDir.Location = New-Object System.Drawing.Point(180, $y)
    $btnResetDir.Size = New-Object System.Drawing.Size(150, 25)
    $btnResetDir.Text = "Reset to Default"
    $btnResetDir.Add_Click({
        $txtSettingsDirectory.Text = Join-Path $env:LOCALAPPDATA "VScanMagic"
    })
    $settingsForm.Controls.Add($btnResetDir)
    $y += 50

    # Save Button
    $btnSave = New-Object System.Windows.Forms.Button
    $btnSave.Location = New-Object System.Drawing.Point(280, $y)
    $btnSave.Size = New-Object System.Drawing.Size(90, 30)
    $btnSave.Text = "Save"
    $btnSave.Add_Click({
        $script:UserSettings.PreparedBy = $txtPreparedBy.Text
        $script:UserSettings.CompanyName = $txtCompanyName.Text
        $script:UserSettings.CompanyAddress = $txtCompanyAddress.Text
        $script:UserSettings.Email = $txtEmail.Text
        $script:UserSettings.PhoneNumber = $txtPhoneNumber.Text
        $script:UserSettings.CompanyPhoneNumber = $txtCompanyPhoneNumber.Text
        
        # Handle settings directory
        $selectedDir = $txtSettingsDirectory.Text.Trim()
        $defaultDir = Join-Path $env:LOCALAPPDATA "VScanMagic"
        if ($selectedDir -eq $defaultDir -or [string]::IsNullOrEmpty($selectedDir)) {
            $script:UserSettings.SettingsDirectory = ""
        } else {
            if (-not (Test-Path $selectedDir)) {
                try {
                    New-Item -Path $selectedDir -ItemType Directory -Force | Out-Null
                } catch {
                    [System.Windows.Forms.MessageBox]::Show("Could not create directory: $($_.Exception.Message)", "Error",
                        [System.Windows.Forms.MessageBoxButtons]::OK, [System.Windows.Forms.MessageBoxIcon]::Error)
                    return
                }
            }
            $script:UserSettings.SettingsDirectory = $selectedDir
        }
        
        # Update paths and migrate files if directory changed
        $oldSettingsPath = $script:SettingsPath
        Update-SettingsPaths
        
        # Migrate files if directory changed
        if ($oldSettingsPath -ne $script:SettingsPath -and (Test-Path $oldSettingsPath)) {
            try {
                $oldDir = [System.IO.Path]::GetDirectoryName($oldSettingsPath)
                $oldRulesPath = Join-Path $oldDir "VScanMagic_RemediationRules.json"
                $oldCoveredPath = Join-Path $oldDir "VScanMagic_CoveredSoftware.json"
                $oldRecommendationsPath = Join-Path $oldDir "VScanMagic_GeneralRecommendations.json"
                
                if ((Test-Path $oldRulesPath) -and -not (Test-Path $script:RemediationRulesPath)) {
                    Copy-Item -Path $oldRulesPath -Destination $script:RemediationRulesPath -Force
                }
                if ((Test-Path $oldCoveredPath) -and -not (Test-Path $script:CoveredSoftwarePath)) {
                    Copy-Item -Path $oldCoveredPath -Destination $script:CoveredSoftwarePath -Force
                }
                if ((Test-Path $oldRecommendationsPath) -and -not (Test-Path $script:GeneralRecommendationsPath)) {
                    Copy-Item -Path $oldRecommendationsPath -Destination $script:GeneralRecommendationsPath -Force
                }
                
                # Reload rules and covered software from new location
                Load-RemediationRules
                Load-CoveredSoftware
                Load-GeneralRecommendations
            } catch {
                Write-Warning "Could not migrate files: $($_.Exception.Message)"
            }
        }

        if (Save-UserSettings) {
            [System.Windows.Forms.MessageBox]::Show("Settings saved successfully!`n`nSettings directory: $script:SettingsDirectory", "Success",
                [System.Windows.Forms.MessageBoxButtons]::OK, [System.Windows.Forms.MessageBoxIcon]::Information)
            $settingsForm.DialogResult = [System.Windows.Forms.DialogResult]::OK
            $settingsForm.Close()
        } else {
            [System.Windows.Forms.MessageBox]::Show("Failed to save settings.", "Error",
                [System.Windows.Forms.MessageBoxButtons]::OK, [System.Windows.Forms.MessageBoxIcon]::Error)
        }
    })
    $settingsForm.Controls.Add($btnSave)

    # Cancel Button
    $btnCancel = New-Object System.Windows.Forms.Button
    $btnCancel.Location = New-Object System.Drawing.Point(380, $y)
    $btnCancel.Size = New-Object System.Drawing.Size(90, 30)
    $btnCancel.Text = "Cancel"
    $btnCancel.Add_Click({
        $settingsForm.DialogResult = [System.Windows.Forms.DialogResult]::Cancel
        $settingsForm.Close()
    })
    $settingsForm.Controls.Add($btnCancel)

    $settingsForm.ShowDialog() | Out-Null
}

function Show-VScanMagicHelpDialog {
    $helpForm = New-Object System.Windows.Forms.Form
    $helpForm.Text = "VScanMagic Help"
    $helpForm.Size = New-Object System.Drawing.Size(620, 580)
    $helpForm.StartPosition = "CenterParent"
    $helpForm.FormBorderStyle = "FixedDialog"
    $helpForm.MaximizeBox = $false

    $tabControl = New-Object System.Windows.Forms.TabControl
    $tabControl.Location = New-Object System.Drawing.Point(10, 10)
    $tabControl.Size = New-Object System.Drawing.Size(588, 500)

    # Tab 1: API Setup
    $tabApi = New-Object System.Windows.Forms.TabPage
    $tabApi.Text = "API Setup"
    $txtApi = New-Object System.Windows.Forms.RichTextBox
    $txtApi.Location = New-Object System.Drawing.Point(8, 8)
    $txtApi.Size = New-Object System.Drawing.Size(565, 440)
    $txtApi.ReadOnly = $true
    $txtApi.Font = New-Object System.Drawing.Font("Segoe UI", 9.5)
    $txtApi.BackColor = [System.Drawing.Color]::White
    $txtApi.Text = @"
GETTING YOUR CONNECTSECURE API CREDENTIALS
------------------------------------------

You need four values: Base URL, Tenant Name, Client ID, and Client Secret.

STEP 1: Log into ConnectSecure
   - Open your ConnectSecure portal in a web browser.
   - The URL is often https://pod{number}.myconnectsecure.com

STEP 2: Open the API Key page
   - Go to: Global > Settings > Users
   - For an existing user: Click the three-dot menu (...) next to your user -> "API Key"
   - For a new API user: Click "Add", create a user, then use Action > API Key

STEP 3: Copy each value
   - Base URL
     Format: https://pod{number}.myconnectsecure.com
     Example: https://pod104.myconnectsecure.com
     Tip: Copy from your browser address bar (without /tenant/... path)
     Or: Check Global Settings > Users > API Documentation for the URL.

   - Tenant Name
     The name you use to log in. Example: river-run
     Often appears in the ConnectSecure portal or API Key page.

   - Client ID
     A long alphanumeric string (e.g., UUID format).
     Copy it exactly - no leading/trailing spaces.

   - Client Secret
     Keep this confidential. Copy it exactly.
     Do not add line breaks or extra spaces when pasting.

CONFIGURING VScanMagic
----------------------

1. Click "API Settings" (or Settings > API in the main form).

2. Paste Base URL, Tenant Name, Client ID, and Client Secret into the matching fields.

3. Click "Test Connection" to verify your credentials.

4. Click "Save" when the test succeeds.

5. Click "Refresh List" in section 1 to load your company list.

TROUBLESHOOTING
   - Auth fails: Re-copy each value; avoid extra spaces or line breaks.
   - "Failed to authorize": Double-check Tenant Name, Client ID, and Client Secret.
   - Base URL: Ensure it starts with https:// and has no trailing slash.
"@
    $tabApi.Controls.Add($txtApi)
    $tabControl.TabPages.Add($tabApi)

    # Tab 2: Workflow
    $tabWorkflow = New-Object System.Windows.Forms.TabPage
    $tabWorkflow.Text = "Workflow"
    $txtWorkflow = New-Object System.Windows.Forms.RichTextBox
    $txtWorkflow.Location = New-Object System.Drawing.Point(8, 8)
    $txtWorkflow.Size = New-Object System.Drawing.Size(565, 440)
    $txtWorkflow.ReadOnly = $true
    $txtWorkflow.Font = New-Object System.Drawing.Font("Segoe UI", 9.5)
    $txtWorkflow.BackColor = [System.Drawing.Color]::White
    $txtWorkflow.Text = @"
VScanMagic WORKFLOW WALKTHROUGH
------------------------------

OPTION A: DOWNLOAD & PROCESS (Recommended)
--------------------------------------------

1. Configure API: Click API Settings, enter credentials, Test, Save.
   (See "API Setup" tab if you need help getting credentials.)

2. Load companies: Click "Refresh List" to load your ConnectSecure companies.
   Select one or more companies (or "All Companies").

3. Set Scan Date: Choose the date for the vulnerability reports.

4. Choose reports: Check the reports to download:
   - Pending EPSS - required for full VScanMagic processing
   - All Vulnerabilities, Suppressed, External, Executive Summary - optional

5. Output Directory: Choose where files will be saved.

6. Report Filters: (Button) Set severity (Critical/High/Medium/Low), Top N count, EPSS threshold.

7. Output Options: (Button) Choose outputs: Excel, Word, Email Template, Ticket Instructions, Time Estimate.

8. Click "Generate":
   - Downloads reports from ConnectSecure for each selected company
   - Runs through General Recommendations and Hostname Review dialogs
   - Generates your selected outputs (Word report, Excel, etc.)

OPTION B: PROCESS FROM FILE
---------------------------

1. In section 2, click "Browse..." and select a previously downloaded Pending EPSS report (XLSX).

2. Enter Client name and Scan Date (sometimes auto-filled from filename).

3. Set Output Options and Report Filters if needed.

4. Click "Generate" to process the file and create outputs.

QUICK ACTIONS
-------------

- Download Standard Reports Only: Downloads All Vulnerabilities, Suppressed, External Scan, and Executive Summary - no VScanMagic processing. Use for archival or quick export.

- Report Filters: Adjust severity filters, Top N, EPSS threshold.

- Output Options: Choose which reports to generate (default: all checked).

- Remediation Rules: Manage custom remediation wording.

- Settings: Access API Settings, Output Options, and Report Filters.
"@
    $tabWorkflow.Controls.Add($txtWorkflow)
    $tabControl.TabPages.Add($tabWorkflow)

    $helpForm.Controls.Add($tabControl)

    $btnClose = New-Object System.Windows.Forms.Button
    $btnClose.Location = New-Object System.Drawing.Point(498, 518)
    $btnClose.Size = New-Object System.Drawing.Size(100, 28)
    $btnClose.Text = "Close"
    $btnClose.DialogResult = [System.Windows.Forms.DialogResult]::OK
    $helpForm.AcceptButton = $btnClose
    $helpForm.Controls.Add($btnClose)

    $helpForm.ShowDialog() | Out-Null
}
