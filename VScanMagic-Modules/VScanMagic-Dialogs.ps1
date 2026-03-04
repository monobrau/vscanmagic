# VScanMagic-Dialogs.ps1 - UI dialogs (Filters, Output Options, Remediation Rules, etc.)
# Dot-sourced by VScanMagic-GUI.ps1
function Show-GeneralRecommendationsDialog {
    param(
        [array]$Top10Data
    )

    if (-not $Top10Data -or $Top10Data.Count -eq 0) {
        [System.Windows.Forms.MessageBox]::Show(
            "No vulnerabilities in this report to add recommendations for. The General Recommendations dialog requires vulnerability data from the Top N list.`n`nIf you expected to see items, check your filters (EPSS, severity, Top N count) or verify the source Excel file has vulnerability data.",
            "No Data",
            [System.Windows.Forms.MessageBoxButtons]::OK,
            [System.Windows.Forms.MessageBoxIcon]::Information
        )
        return @()
    }

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
    $dataGridView.MultiSelect = $true
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
        if ($null -eq $item) { continue }
        $productName = if ($item.Product) { [string]$item.Product } else { "(Unknown Product)" }
        $row = $dataGridView.Rows.Add()
        $dataGridView.Rows[$row].Cells["Product"].Value = $productName
        $dataGridView.Rows[$row].Tag = $item  # Store item for CVE/context when using AI improve
        
        # Try to find matching recommendation using pattern matching
        $matchingRec = $null
        foreach ($rec in $script:GeneralRecommendations) {
            if ($productName -like $rec.Product) {
                $matchingRec = $rec
                break
            }
        }
        
        try {
            $recommendationsText = if ($matchingRec) { $matchingRec.Recommendations } elseif ($item.Fix -and -not [string]::IsNullOrWhiteSpace($item.Fix)) { ConvertTo-ReadableFixText -RawFix $item.Fix } else { Get-RemediationGuidance -ProductName $productName -OutputType 'Word' }
        } catch {
            $recommendationsText = "Unable to load recommendation: $($_.Exception.Message)"
        }
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
            $row.Cells["Recommendations"].Value = if ($matchingRec) { $matchingRec.Recommendations } else { Get-RemediationGuidance -ProductName $product -OutputType 'Word' }
        }
    })
    $recDialog.Controls.Add($btnLoadDefaults)

    $btnImproveSelected = New-Object System.Windows.Forms.Button
    $btnImproveSelected.Location = New-Object System.Drawing.Point(150, $y)
    $btnImproveSelected.Size = New-Object System.Drawing.Size(140, 30)
    $btnImproveSelected.Text = "Improve Selected"
    $btnImproveSelected.Add_Click({
        $sel = $dataGridView.SelectedRows
        if (-not $sel -or $sel.Count -eq 0) {
            [System.Windows.Forms.MessageBox]::Show("Select one or more rows first.", "Improve with AI", [System.Windows.Forms.MessageBoxButtons]::OK, [System.Windows.Forms.MessageBoxIcon]::Information)
            return
        }
        $rowsToImprove = @($sel | Where-Object { -not $_.IsNewRow -and -not [string]::IsNullOrWhiteSpace([string]$_.Cells["Recommendations"].Value) })
        if ($rowsToImprove.Count -eq 0) {
            [System.Windows.Forms.MessageBox]::Show("Selected rows have no text. Enter text in the Recommendations cell first.", "Improve with AI", [System.Windows.Forms.MessageBoxButtons]::OK, [System.Windows.Forms.MessageBoxIcon]::Information)
            return
        }
        $recDialog.Cursor = [System.Windows.Forms.Cursors]::WaitCursor
        try {
            $idx = 0
            foreach ($row in $rowsToImprove) {
                $currentText = [string]$row.Cells["Recommendations"].Value
                $productName = [string]$row.Cells["Product"].Value
                $cveIds = if ($row.Tag -and $row.Tag.CveIds) { [string]$row.Tag.CveIds } else { "" }
                if (-not [string]::IsNullOrWhiteSpace($currentText)) {
                    $improved = Invoke-AIImproveRemediationText -Text $currentText -ProductName $productName -CveIdList $cveIds
                    $row.Cells["Recommendations"].Value = $improved
                    $idx++
                    if ($idx -lt $rowsToImprove.Count) {
                        Start-Sleep -Seconds 2
                        [System.Windows.Forms.Application]::DoEvents()
                    }
                }
            }
        } finally {
            $recDialog.Cursor = [System.Windows.Forms.Cursors]::Default
        }
    })
    $aiConfigured = Test-AIApiKeyConfigured
    $btnImproveSelected.Enabled = $aiConfigured
    $recDialog.Controls.Add($btnImproveSelected)

    $btnImproveAll = New-Object System.Windows.Forms.Button
    $btnImproveAll.Location = New-Object System.Drawing.Point(300, $y)
    $btnImproveAll.Size = New-Object System.Drawing.Size(120, 30)
    $btnImproveAll.Text = "Improve All"
    $btnImproveAll.Add_Click({
        $rowsToImprove = @()
        foreach ($row in $dataGridView.Rows) {
            if ($row.IsNewRow) { continue }
            $currentText = [string]$row.Cells["Recommendations"].Value
            if (-not [string]::IsNullOrWhiteSpace($currentText)) { $rowsToImprove += $row }
        }
        if ($rowsToImprove.Count -eq 0) {
            [System.Windows.Forms.MessageBox]::Show("No rows have text to improve. Enter text in the Recommendations cells first.", "Improve with AI", [System.Windows.Forms.MessageBoxButtons]::OK, [System.Windows.Forms.MessageBoxIcon]::Information)
            return
        }
        $recDialog.Cursor = [System.Windows.Forms.Cursors]::WaitCursor
        try {
            $idx = 0
            foreach ($row in $rowsToImprove) {
                $currentText = [string]$row.Cells["Recommendations"].Value
                $productName = [string]$row.Cells["Product"].Value
                $cveIds = if ($row.Tag -and $row.Tag.CveIds) { [string]$row.Tag.CveIds } else { "" }
                $improved = Invoke-AIImproveRemediationText -Text $currentText -ProductName $productName -CveIdList $cveIds
                $row.Cells["Recommendations"].Value = $improved
                $idx++
                if ($idx -lt $rowsToImprove.Count) {
                    Start-Sleep -Seconds 2
                    [System.Windows.Forms.Application]::DoEvents()
                }
            }
        } finally {
            $recDialog.Cursor = [System.Windows.Forms.Cursors]::Default
        }
    })
    $btnImproveAll.Enabled = $aiConfigured
    $recDialog.Controls.Add($btnImproveAll)
    if (-not $aiConfigured) {
        $ttAI = New-Object System.Windows.Forms.ToolTip
        $ttAI.SetToolTip($btnImproveSelected, "Configure AI API Keys in Settings to enable.")
        $ttAI.SetToolTip($btnImproveAll, "Configure AI API Keys in Settings to enable.")
    }

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

function Show-SetClientTypesDialog {
    param(
        [array]$CompaniesToProcess,
        [bool]$DefaultIsRMITPlus = $false
    )
    if (-not $CompaniesToProcess -or $CompaniesToProcess.Count -eq 0) { return @{} }
    $form = New-Object System.Windows.Forms.Form
    $form.Text = "Set Client Type (RMIT+) for Each Company"
    $form.Size = New-Object System.Drawing.Size(420, 380)
    $form.StartPosition = "CenterParent"
    $form.FormBorderStyle = "FixedDialog"
    $form.MaximizeBox = $false
    $form.MinimizeBox = $false
    $form.MinimumSize = $form.Size
    $lbl = New-Object System.Windows.Forms.Label
    $lbl.Location = New-Object System.Drawing.Point(20, 12)
    $lbl.Size = New-Object System.Drawing.Size(360, 36)
    $lbl.Text = "Set RMIT+ for each client before processing. RMIT+ clients have ticketing and remediation coverage under the agreement."
    $lbl.AutoSize = $false
    $lbl.MaximumSize = New-Object System.Drawing.Size(360, 50)
    $form.Controls.Add($lbl)
    $panelDgv = New-Object System.Windows.Forms.Panel
    $panelDgv.Location = New-Object System.Drawing.Point(20, 55)
    $panelDgv.Size = New-Object System.Drawing.Size(360, 220)
    $panelDgv.BorderStyle = [System.Windows.Forms.BorderStyle]::FixedSingle
    $dgv = New-Object System.Windows.Forms.DataGridView
    $dgv.Dock = [System.Windows.Forms.DockStyle]::Fill
    $dgv.AutoSizeColumnsMode = [System.Windows.Forms.DataGridViewAutoSizeColumnsMode]::Fill
    $dgv.AllowUserToAddRows = $false
    $dgv.AllowUserToDeleteRows = $false
    $dgv.ReadOnly = $false
    $dgv.SelectionMode = [System.Windows.Forms.DataGridViewSelectionMode]::FullRowSelect
    $colName = New-Object System.Windows.Forms.DataGridViewTextBoxColumn
    $colName.Name = "ClientName"
    $colName.HeaderText = "Client"
    $colName.ReadOnly = $true
    $colName.FillWeight = 70
    $dgv.Columns.Add($colName) | Out-Null
    $colRMIT = New-Object System.Windows.Forms.DataGridViewCheckBoxColumn
    $colRMIT.Name = "IsRMITPlus"
    $colRMIT.HeaderText = "RMIT+"
    $colRMIT.FillWeight = 30
    $dgv.Columns.Add($colRMIT) | Out-Null
    foreach ($c in $CompaniesToProcess) {
        $clientName = if ($c.ClientName) { $c.ClientName } else { ($c.Company.DisplayName -replace '\s*\(ID:\s*\d+\)\s*$', '').Trim() }
        $clientName = if ([string]::IsNullOrWhiteSpace($clientName)) { $c.Company.DisplayName } else { $clientName }
        $defaultVal = if ($c.IsRMITPlus -ne $null) { $c.IsRMITPlus } else { $DefaultIsRMITPlus }
        $row = $dgv.Rows.Add($clientName, $defaultVal)
        $dgv.Rows[$row].Tag = $c
    }
    $panelDgv.Controls.Add($dgv)
    $form.Controls.Add($panelDgv)
    $btnSetAll = New-Object System.Windows.Forms.Button
    $btnSetAll.Location = New-Object System.Drawing.Point(20, 285)
    $btnSetAll.Size = New-Object System.Drawing.Size(80, 26)
    $btnSetAll.Text = "All RMIT+"
    $btnSetAll.Add_Click({ foreach ($r in $dgv.Rows) { if (-not $r.IsNewRow) { $r.Cells["IsRMITPlus"].Value = $true } } })
    $form.Controls.Add($btnSetAll)
    $btnSetNone = New-Object System.Windows.Forms.Button
    $btnSetNone.Location = New-Object System.Drawing.Point(105, 285)
    $btnSetNone.Size = New-Object System.Drawing.Size(80, 26)
    $btnSetNone.Text = "All RMIT"
    $btnSetNone.Add_Click({ foreach ($r in $dgv.Rows) { if (-not $r.IsNewRow) { $r.Cells["IsRMITPlus"].Value = $false } } })
    $form.Controls.Add($btnSetNone)
    $btnCancel = New-Object System.Windows.Forms.Button
    $btnCancel.Location = New-Object System.Drawing.Point(210, 285)
    $btnCancel.Size = New-Object System.Drawing.Size(80, 26)
    $btnCancel.Text = "Cancel"
    $btnCancel.DialogResult = [System.Windows.Forms.DialogResult]::Cancel
    $form.Controls.Add($btnCancel)
    $btnOK = New-Object System.Windows.Forms.Button
    $btnOK.Location = New-Object System.Drawing.Point(300, 285)
    $btnOK.Size = New-Object System.Drawing.Size(80, 26)
    $btnOK.Text = "OK"
    $btnOK.DialogResult = [System.Windows.Forms.DialogResult]::OK
    $form.Controls.Add($btnOK)
    $form.AcceptButton = $btnOK
    $form.CancelButton = $btnCancel
    if ($form.ShowDialog() -ne [System.Windows.Forms.DialogResult]::OK) { return $null }
    $result = @{}
    $idx = 0
    foreach ($c in $CompaniesToProcess) {
        $val = $false
        if ($idx -lt $dgv.Rows.Count -and -not $dgv.Rows[$idx].IsNewRow) {
            $val = [bool]$dgv.Rows[$idx].Cells["IsRMITPlus"].Value
        }
        $key = if ($c.Company -and $c.Company.Id -ne $null) { $c.Company.Id } else { $idx }
        $result[$key] = $val
        $idx++
    }
    return $result
}

function Show-HostnameReviewDialog {
    param(
        [array]$Top10Data,
        [int]$CompanyId = 0
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
        $colUsername.ReadOnly = $false
        $dataGridView.Columns.Add($colUsername) | Out-Null

        $colVulnCount = New-Object System.Windows.Forms.DataGridViewTextBoxColumn
        $colVulnCount.Name = "VulnCount"
        $colVulnCount.HeaderText = "Vulnerability Count"
        $colVulnCount.Width = 150
        $colVulnCount.ReadOnly = $true
        $colVulnCount.ValueType = [int]
        $dataGridView.Columns.Add($colVulnCount) | Out-Null

        # Populate grid with hostnames (or IP fallback), sorted by vulnerability count descending
        # For Windows 11 O/S: below threshold = unselected, above = selected
        $threshold = if ($null -ne $script:UserSettings.HostnameReviewWindows11Threshold) { [int]$script:UserSettings.HostnameReviewWindows11Threshold } else { 350 }
        $isWindows11 = $item.Product -like "*Windows 11*"
        $sortedSystems = $item.AffectedSystems | Sort-Object -Property VulnCount -Descending
        foreach ($sys in $sortedSystems) {
            $row = $dataGridView.Rows.Add()
            $defaultInclude = if ($isWindows11) { $sys.VulnCount -ge $threshold } else { $true }
            $dataGridView.Rows[$row].Cells["Include"].Value = $defaultInclude
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

    $btnLookupConnectSecure = New-Object System.Windows.Forms.Button
    $btnLookupConnectSecure.Location = New-Object System.Drawing.Point(240, $y)
    $btnLookupConnectSecure.Size = New-Object System.Drawing.Size(170, 30)
    $btnLookupConnectSecure.Text = "Lookup from ConnectSecure"
    $btnLookupConnectSecure.Enabled = ($CompanyId -gt 0)
    $btnLookupConnectSecure.Add_Click({
        if ($CompanyId -le 0) { return }
        $creds = Load-ConnectSecureCredentials
        if (-not $creds -or [string]::IsNullOrWhiteSpace($creds.BaseUrl)) {
            [System.Windows.Forms.MessageBox]::Show("Configure ConnectSecure API in Settings first (Settings > API Settings...).", "Not Configured", [System.Windows.Forms.MessageBoxButtons]::OK, [System.Windows.Forms.MessageBoxIcon]::Information)
            return
        }
        $connected = Connect-ConnectSecureAPI -BaseUrl $creds.BaseUrl -TenantName $creds.TenantName -ClientId $creds.ClientId -ClientSecret $creds.ClientSecret
        if (-not $connected) {
            [System.Windows.Forms.MessageBox]::Show("Failed to connect to ConnectSecure. Check API Settings.", "Connection Error", [System.Windows.Forms.MessageBoxButtons]::OK, [System.Windows.Forms.MessageBoxIcon]::Warning)
            return
        }
        $allHostnames = @()
        foreach ($dgv in $tabDataGridViews) {
            foreach ($row in $dgv.Rows) {
                if ($row.IsNewRow) { continue }
                $hn = [string]$row.Cells["Hostname"].Value
                if (-not [string]::IsNullOrWhiteSpace($hn)) { $allHostnames += $hn }
            }
        }
        $allHostnames = $allHostnames | Select-Object -Unique
        if ($allHostnames.Count -eq 0) {
            [System.Windows.Forms.MessageBox]::Show("No hostnames to look up.", "Lookup", [System.Windows.Forms.MessageBoxButtons]::OK, [System.Windows.Forms.MessageBoxIcon]::Information)
            return
        }
        $hostDialog.Cursor = [System.Windows.Forms.Cursors]::WaitCursor
        try {
            $userMap = Get-ConnectSecureUsernamesByHostname -CompanyId $CompanyId -Hostnames $allHostnames
            $filled = 0
            foreach ($dgv in $tabDataGridViews) {
                foreach ($row in $dgv.Rows) {
                    if ($row.IsNewRow) { continue }
                    $hn = [string]$row.Cells["Hostname"].Value
                    $hnKey = $hn.Trim()
                    $currentUser = [string]$row.Cells["Username"].Value
                    if ($hnKey -and $userMap.ContainsKey($hnKey) -and -not [string]::IsNullOrWhiteSpace($userMap[$hnKey])) {
                        if ([string]::IsNullOrWhiteSpace($currentUser)) {
                            $row.Cells["Username"].Value = $userMap[$hnKey]
                            $filled++
                        }
                    }
                }
            }
            [System.Windows.Forms.MessageBox]::Show("Lookup complete. Filled $filled username(s) from ConnectSecure.", "Lookup Complete", [System.Windows.Forms.MessageBoxButtons]::OK, [System.Windows.Forms.MessageBoxIcon]::Information)
        } finally {
            $hostDialog.Cursor = [System.Windows.Forms.Cursors]::Default
        }
    })
    $hostDialog.Controls.Add($btnLookupConnectSecure)

    $btnLookupConnectWise = New-Object System.Windows.Forms.Button
    $btnLookupConnectWise.Location = New-Object System.Drawing.Point(420, $y)
    $btnLookupConnectWise.Size = New-Object System.Drawing.Size(160, 30)
    $btnLookupConnectWise.Text = "Lookup from ConnectWise"
    $btnLookupConnectWise.Add_Click({
        $creds = Load-ConnectWiseAutomateCredentials
        if (-not $creds -or [string]::IsNullOrWhiteSpace($creds.BaseUrl)) {
            [System.Windows.Forms.MessageBox]::Show("Configure ConnectWise Automate in Settings first (Settings > ConnectWise Automate...).", "Not Configured", [System.Windows.Forms.MessageBoxButtons]::OK, [System.Windows.Forms.MessageBoxIcon]::Information)
            return
        }
        $allHostnames = @()
        foreach ($dgv in $tabDataGridViews) {
            foreach ($row in $dgv.Rows) {
                if ($row.IsNewRow) { continue }
                $hn = [string]$row.Cells["Hostname"].Value
                if (-not [string]::IsNullOrWhiteSpace($hn)) { $allHostnames += $hn }
            }
        }
        $allHostnames = $allHostnames | Select-Object -Unique
        if ($allHostnames.Count -eq 0) {
            [System.Windows.Forms.MessageBox]::Show("No hostnames to look up.", "Lookup", [System.Windows.Forms.MessageBoxButtons]::OK, [System.Windows.Forms.MessageBoxIcon]::Information)
            return
        }
        $hostDialog.Cursor = [System.Windows.Forms.Cursors]::WaitCursor
        try {
            $userMap = Get-ConnectWiseUsernamesByHostname -Hostnames $allHostnames
            $filled = 0
            foreach ($dgv in $tabDataGridViews) {
                foreach ($row in $dgv.Rows) {
                    if ($row.IsNewRow) { continue }
                    $hn = [string]$row.Cells["Hostname"].Value
                    $currentUser = [string]$row.Cells["Username"].Value
                    if ($hn -and $userMap.ContainsKey($hn) -and -not [string]::IsNullOrWhiteSpace($userMap[$hn])) {
                        if ([string]::IsNullOrWhiteSpace($currentUser)) {
                            $row.Cells["Username"].Value = $userMap[$hn]
                            $filled++
                        }
                    }
                }
            }
            [System.Windows.Forms.MessageBox]::Show("Lookup complete. Filled $filled username(s) from ConnectWise Automate.", "Lookup Complete", [System.Windows.Forms.MessageBoxButtons]::OK, [System.Windows.Forms.MessageBoxIcon]::Information)
        } finally {
            $hostDialog.Cursor = [System.Windows.Forms.Cursors]::Default
        }
    })
    $hostDialog.Controls.Add($btnLookupConnectWise)

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

    # Auto-lookup usernames from ConnectSecure when setting is enabled
    $hostDialog.Add_Shown({
        if ($CompanyId -le 0) { return }
        if (-not $script:UserSettings.HostnameReviewAutoLookupConnectSecure) { return }
        $creds = Load-ConnectSecureCredentials
        if (-not $creds -or [string]::IsNullOrWhiteSpace($creds.BaseUrl)) { return }
        $connected = Connect-ConnectSecureAPI -BaseUrl $creds.BaseUrl -TenantName $creds.TenantName -ClientId $creds.ClientId -ClientSecret $creds.ClientSecret
        if (-not $connected) { return }
        $allHostnames = @()
        foreach ($dgv in $tabDataGridViews) {
            foreach ($row in $dgv.Rows) {
                if ($row.IsNewRow) { continue }
                $hn = [string]$row.Cells["Hostname"].Value
                if (-not [string]::IsNullOrWhiteSpace($hn)) { $allHostnames += $hn }
            }
        }
        $allHostnames = $allHostnames | Select-Object -Unique
        if ($allHostnames.Count -eq 0) { return }
        $hostDialog.Cursor = [System.Windows.Forms.Cursors]::WaitCursor
        try {
            $userMap = Get-ConnectSecureUsernamesByHostname -CompanyId $CompanyId -Hostnames $allHostnames
            $filled = 0
            foreach ($dgv in $tabDataGridViews) {
                foreach ($row in $dgv.Rows) {
                    if ($row.IsNewRow) { continue }
                    $hn = [string]$row.Cells["Hostname"].Value
                    $hnKey = $hn.Trim()
                    $currentUser = [string]$row.Cells["Username"].Value
                    if ($hnKey -and $userMap.ContainsKey($hnKey) -and -not [string]::IsNullOrWhiteSpace($userMap[$hnKey])) {
                        if ([string]::IsNullOrWhiteSpace($currentUser)) {
                            $row.Cells["Username"].Value = $userMap[$hnKey]
                            $filled++
                        }
                    }
                }
            }
            if ($filled -gt 0) {
                [System.Windows.Forms.MessageBox]::Show("Auto-lookup complete. Filled $filled username(s) from ConnectSecure.", "Lookup Complete", [System.Windows.Forms.MessageBoxButtons]::OK, [System.Windows.Forms.MessageBoxIcon]::Information)
            }
        } finally {
            $hostDialog.Cursor = [System.Windows.Forms.Cursors]::Default
        }
    })

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
            # Default 3rd party status: first-party vendors (SonicWall, Fortinet, Microsoft, HP, Duo) are covered; all others are 3rd party
            $isThirdPartyDefault = -not (Test-IsFirstPartyVendor -ProductName $item.Product)
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

    # Validation on OK (empty defaults to 0)
    $btnOK.Add_Click({
        foreach ($row in $dataGridView.Rows) {
            if ($row.IsNewRow) { continue }
            $timeEstimate = [string]$row.Cells["TimeEstimate"].Value
            if ([string]::IsNullOrWhiteSpace($timeEstimate)) { continue }  # Empty = 0, no validation needed
            # Validate entered value is a valid non-negative number
            $timeValue = 0
            if (-not [double]::TryParse($timeEstimate, [ref]$timeValue) -or $timeValue -lt 0) {
                [System.Windows.Forms.MessageBox]::Show("Time estimate must be a valid non-negative number (hours). Empty values default to 0.", "Validation Error",
                    [System.Windows.Forms.MessageBoxButtons]::OK, [System.Windows.Forms.MessageBoxIcon]::Warning)
                $timeDialog.DialogResult = [System.Windows.Forms.DialogResult]::None
                return
            }
        }
    })

    $result = $timeDialog.ShowDialog()

    if ($result -eq [System.Windows.Forms.DialogResult]::OK) {
        # Build result array (empty time estimate defaults to 0)
        $timeEstimates = @()
        foreach ($row in $dataGridView.Rows) {
            if ($row.IsNewRow) { continue }
            $rawTime = [string]$row.Cells["TimeEstimate"].Value
            $timeVal = 0.0
            if (-not [string]::IsNullOrWhiteSpace($rawTime)) { [double]::TryParse($rawTime, [ref]$timeVal) | Out-Null }
            $timeEstimates += [PSCustomObject]@{
                Product = $row.Cells["Product"].Value
                TimeEstimate = $timeVal
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
        [array]$GeneralRecommendations = $null,
        [switch]$PassThru
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

        $content = $sb.ToString()
        if ($OutputPath) {
            $content | Out-File -FilePath $OutputPath -Encoding UTF8
        }
        if ($PassThru) { return $content }

    } catch {
        Write-Log "Error generating time estimate: $($_.Exception.Message)" -Level Error
        throw
    }
}

function Format-TicketInstructionSpacing {
    param([string]$Text)
    if ([string]::IsNullOrWhiteSpace($Text)) { return $Text }
    $t = $Text -replace '[ \t]+', ' ' -replace '(\r?\n){2,}', "`r`n`r`n"
    return $t.Trim()
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

        $count = if ($TopTenData) { $TopTenData.Count } else { 0 }
        $headerTitle = if ($count -eq 10) { "TOP 10 VULNERABILITIES" } elseif ($count -gt 0) { "TOP $count VULNERABILITIES" } else { "VULNERABILITY REMEDIATIONS" }

        $sb = New-Object System.Text.StringBuilder
        [void]$sb.AppendLine("=".PadRight(100, '='))
        [void]$sb.AppendLine("$headerTitle - TICKET INSTRUCTIONS")
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
            } elseif ($item.Product -like "*Windows 11*") {
                $ticketSubject += "$($item.Product) - Updates Required"
            } elseif ($item.Product -like "*Windows Server*") {
                $ticketSubject += "$($item.Product) - Updates Required"
            } elseif ($item.Product -like "*Windows*") {
                $ticketSubject += "$($item.Product) - Patch Management Required"
            } elseif ($item.Product -like "*printer*" -or $item.Product -like "*Ripple20*") {
                $ticketSubject += "$($item.Product) - Firmware Update Required"
            } elseif ($item.Product -like "*Microsoft Teams*") {
                $ticketSubject += "$($item.Product) - Application Update Required"
            } elseif ((Test-IsMicrosoftApplication -ProductName $item.Product) -and $IsRMITPlus) {
                $ticketSubject += "$($item.Product) - Updates Required"
            } elseif ((Test-IsVMwareProduct -ProductName $item.Product) -and $IsRMITPlus) {
                $ticketSubject += "$($item.Product) - Update Required"
            } elseif (Test-IsAutoUpdatingSoftware -ProductName $item.Product) {
                $ticketSubject += "$($item.Product) - This software updates automatically"
            } else {
                $ticketSubject += "$($item.Product) - Update Required"
            }

            # Append modifier text for subject (no ticket-generated text; subject IS the ticket)
            if ($null -ne $timeEstimate -and $IsRMITPlus) {
                $afterHours = $timeEstimate.AfterHours
                $ticketGenerated = $timeEstimate.TicketGenerated
                $thirdParty = $timeEstimate.ThirdParty
                $autoTicketGenerated = $thirdParty -and $afterHours
                $isTicketGenerated = $ticketGenerated -or $autoTicketGenerated
                $modifierText = Get-ModifierTextForSubject -AfterHours $afterHours -TicketGenerated $isTicketGenerated -ThirdParty $thirdParty
                if (-not [string]::IsNullOrWhiteSpace($modifierText)) {
                    $ticketSubject += $modifierText
                }
                if ($afterHours) {
                    $ticketSubject = "After Hours - $ticketSubject"
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

            # ConnectSecure Solution/Fix when available
            if ($item.Fix -and -not [string]::IsNullOrWhiteSpace($item.Fix)) {
                [void]$sb.AppendLine()
                [void]$sb.AppendLine("ConnectSecure Solution:")
                [void]$sb.AppendLine((ConvertTo-ReadableFixText -RawFix $item.Fix))
            }
            [void]$sb.AppendLine()

            # Add General Recommendations if available (use pre-built map for O(1) lookup)
            $matchingRec = if ($generalRecMap.ContainsKey($item.Product)) { $generalRecMap[$item.Product] } else { $null }
            if ($null -ne $matchingRec -and -not [string]::IsNullOrWhiteSpace($matchingRec.Recommendations)) {
                [void]$sb.AppendLine("General Recommendations:")
                [void]$sb.AppendLine($matchingRec.Recommendations)
                [void]$sb.AppendLine()
            }

            [void]$sb.AppendLine("Uninstalling the software or removing/replacing the device is also a valid form of remediation when updating or patching is not feasible; the vulnerability will show as remediated on the next scan.")
            [void]$sb.AppendLine()
            [void]$sb.AppendLine("Sometimes it will not be possible to remediate the vulnerability for business or technical reasons. Other times it will be a false positive detection. In the event of either case please reach out to someone on the vulnerability scan team with your findings and we can suppress the vulnerability so it doesn't come up on future scans or remediations.")
            [void]$sb.AppendLine()
        }

        [void]$sb.AppendLine()
        [void]$sb.AppendLine("=".PadRight(100, '='))
        [void]$sb.AppendLine("END OF TICKET INSTRUCTIONS")
        [void]$sb.AppendLine("=".PadRight(100, '='))

        $output = Format-TicketInstructionSpacing -Text $sb.ToString()
        $output | Out-File -FilePath $OutputPath -Encoding UTF8

    } catch {
        Write-Log "Error generating ticket instructions: $($_.Exception.Message)" -Level Error
    }
}

function New-TicketInstructionsHtml {
    param(
        [string]$OutputPath,
        [array]$TopTenData,
        [array]$TimeEstimates = $null,
        [bool]$IsRMITPlus = $false,
        [array]$GeneralRecommendations = $null,
        [switch]$PassThru
    )

    try {
        Write-Log "Generating ticket instructions (HTML)..."

        $count = if ($TopTenData) { $TopTenData.Count } else { 0 }
        $headerTitle = if ($count -eq 10) { "TOP 10 VULNERABILITIES" } elseif ($count -gt 0) { "TOP $count VULNERABILITIES" } else { "VULNERABILITY REMEDIATIONS" }

        $escapeHtml = { param([string]$s) if ([string]::IsNullOrEmpty($s)) { return "" }; [System.Net.WebUtility]::HtmlEncode($s) }
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

        $sections = [System.Collections.ArrayList]::new()

        for ($i = 0; $i -lt $TopTenData.Count; $i++) {
            $item = $TopTenData[$i]
            $num = $i + 1
            $sectionId = "vuln-$num"
            $timeEstimate = if ($timeEstimateMap.ContainsKey($item.Product)) { $timeEstimateMap[$item.Product] } else { $null }

            $ticketSubject = "Vulnerability Remediation - "
            if ($item.Product -like "*Windows Server 2012*" -or $item.Product -like "*end-of-life*" -or $item.Product -like "*out of support*") {
                $ticketSubject += "$($item.Product) - End of Support Migration Required"
            } elseif ($item.Product -like "*Windows 10*") {
                $ticketSubject += "$($item.Product) - Windows 10 is End of Life"
            } elseif ($item.Product -like "*Windows 11*") {
                $ticketSubject += "$($item.Product) - Updates Required"
            } elseif ($item.Product -like "*Windows Server*") {
                $ticketSubject += "$($item.Product) - Updates Required"
            } elseif ($item.Product -like "*Windows*") {
                $ticketSubject += "$($item.Product) - Patch Management Required"
            } elseif ($item.Product -like "*printer*" -or $item.Product -like "*Ripple20*") {
                $ticketSubject += "$($item.Product) - Firmware Update Required"
            } elseif ($item.Product -like "*Microsoft Teams*") {
                $ticketSubject += "$($item.Product) - Application Update Required"
            } elseif ((Test-IsMicrosoftApplication -ProductName $item.Product) -and $IsRMITPlus) {
                $ticketSubject += "$($item.Product) - Updates Required"
            } elseif ((Test-IsVMwareProduct -ProductName $item.Product) -and $IsRMITPlus) {
                $ticketSubject += "$($item.Product) - Update Required"
            } elseif (Test-IsAutoUpdatingSoftware -ProductName $item.Product) {
                $ticketSubject += "$($item.Product) - This software updates automatically"
            } else {
                $ticketSubject += "$($item.Product) - Update Required"
            }
            # Append modifier text for subject (no ticket-generated text; subject IS the ticket)
            if ($null -ne $timeEstimate -and $IsRMITPlus) {
                $afterHours = $timeEstimate.AfterHours
                $ticketGenerated = $timeEstimate.TicketGenerated
                $thirdParty = $timeEstimate.ThirdParty
                $autoTicketGenerated = $thirdParty -and $afterHours
                $isTicketGenerated = $ticketGenerated -or $autoTicketGenerated
                $modifierText = Get-ModifierTextForSubject -AfterHours $afterHours -TicketGenerated $isTicketGenerated -ThirdParty $thirdParty
                if (-not [string]::IsNullOrWhiteSpace($modifierText)) {
                    $ticketSubject += $modifierText
                }
                if ($afterHours) {
                    $ticketSubject = "After Hours - $ticketSubject"
                }
            }

            $productDisplay = & $escapeHtml $item.Product
            # Only color red when 3rd party checkbox is explicitly checked (no product-based fallback)
            $isThirdParty = ($null -ne $timeEstimate -and $IsRMITPlus -and $timeEstimate.ThirdParty)
            if ($isThirdParty) {
                $productDisplay = "<span class=`"third-party`">$productDisplay</span>"
            }

            $sectionBody = @"
$ticketSubject

Product/System:          $($item.Product)
Risk Score:              $($item.RiskScore.ToString('N2'))
EPSS Score:              $($item.EPSSScore.ToString('N4'))
Average CVSS:            $($item.AvgCVSS.ToString('N2'))
Total Vulnerabilities:   $($item.VulnCount)
Affected Systems Count:  $($item.AffectedSystems.Count)

NOTE: This remediation can go to any available technician.

Affected Systems:
"@
            $uniqueSystems = $item.AffectedSystems | Select-Object HostName, IP, Username -Unique
            foreach ($sys in $uniqueSystems) {
                $hostname = $sys.HostName
                $ip = $sys.IP
                $username = $sys.Username
                $systemLine = if ($hostname) { $hostname } else { $ip }
                if (-not [string]::IsNullOrWhiteSpace($username)) { $systemLine += " ($username)" }
                if (-not [string]::IsNullOrWhiteSpace($ip)) { $systemLine += " - $ip" }
                $sectionBody += "`n  - $systemLine"
            }
            $sectionBody += "`n`nRemediation Instructions:`n"
            $remediationText = Get-RemediationGuidance -ProductName $item.Product -OutputType 'Ticket'
            $sectionBody += $remediationText
            if ($item.Fix -and -not [string]::IsNullOrWhiteSpace($item.Fix)) {
                $sectionBody += "`n`nConnectSecure Solution:`n"
                $sectionBody += (ConvertTo-ReadableFixText -RawFix $item.Fix)
            }
            $matchingRec = if ($generalRecMap.ContainsKey($item.Product)) { $generalRecMap[$item.Product] } else { $null }
            if ($null -ne $matchingRec -and -not [string]::IsNullOrWhiteSpace($matchingRec.Recommendations)) {
                $sectionBody += "`n`nGeneral Recommendations:`n"
                $sectionBody += $matchingRec.Recommendations
            }
            $sectionBody += "`n`nUninstalling the software or removing/replacing the device is also a valid form of remediation when updating or patching is not feasible; the vulnerability will show as remediated on the next scan.`n`n"
            $sectionBody += "Sometimes it will not be possible to remediate the vulnerability for business or technical reasons. Other times it will be a false positive detection. In the event of either case please reach out to someone on the vulnerability scan team with your findings and we can suppress the vulnerability so it doesn't come up on future scans or remediations."

            $sectionBody = Format-TicketInstructionSpacing -Text $sectionBody
            $sectionBodyEscaped = & $escapeHtml $sectionBody
            $sectionBodyEscaped = $sectionBodyEscaped -replace "`r?`n", "<br>`n"
            if ($isThirdParty) {
                $productEscaped = & $escapeHtml $item.Product
                $sectionBodyEscaped = $sectionBodyEscaped -replace [regex]::Escape($productEscaped), $productDisplay
            }
            $subjectEscaped = & $escapeHtml $ticketSubject
            $subjectEscaped = $subjectEscaped -replace "`r?`n", " "
            $subjectDisplay = $subjectEscaped
            if ($isThirdParty) {
                $productEscaped = & $escapeHtml $item.Product
                $subjectDisplay = $subjectDisplay -replace [regex]::Escape($productEscaped), $productDisplay
            }

            $sectionHtml = @"
<section id="$sectionId" class="vuln-section collapsed">
  <div class="section-header">
    <h2>Vulnerability #$num</h2>
    <div class="subject-line">$subjectDisplay</div>
    <div class="section-actions">
      <button type="button" onclick="copySubject(this)" data-subject="$subjectEscaped">Copy Subject</button>
      <button type="button" onclick="copySection('$sectionId')">Copy Section</button>
      <button type="button" class="toggle-btn" onclick="toggleSection('$sectionId')">▸ Show details</button>
    </div>
  </div>
  <div class="section-content">
    <pre class="section-text">$sectionBodyEscaped</pre>
  </div>
</section>
"@
            $null = $sections.Add($sectionHtml)
        }

        $sectionsHtml = ($sections -join "`n")

        $html = @"
<!DOCTYPE html>
<html lang="en">
<head>
  <meta charset="UTF-8">
  <meta name="viewport" content="width=device-width, initial-scale=1">
  <title>Ticket Instructions - $(Get-Date -Format 'yyyy-MM-dd HH:mm')</title>
  <style>
    body { font-family: Segoe UI, Arial, sans-serif; margin: 0; padding: 20px; background: #f5f5f5; }
    .vuln-section { background: #fff; padding: 20px; margin-bottom: 24px; border-radius: 8px; box-shadow: 0 2px 8px rgba(0,0,0,0.08); }
    .vuln-section h2 { margin: 0 0 12px 0; font-size: 18px; color: #333; border-bottom: 1px solid #ddd; padding-bottom: 8px; }
    .subject-line { font-size: 13px; color: #333; margin: 8px 0 12px 0; }
    .section-actions { margin-bottom: 0; }
    .section-actions button { margin-right: 8px; padding: 8px 16px; cursor: pointer; background: #0066cc; color: #fff; border: none; border-radius: 4px; font-size: 13px; }
    .section-actions button:hover { background: #0052a3; }
    .section-actions .toggle-btn { background: #6c757d; }
    .section-actions .toggle-btn:hover { background: #5a6268; }
    .section-content { margin-top: 16px; padding-top: 16px; border-top: 1px solid #eee; }
    .vuln-section.collapsed .section-content { display: none; }
    .section-text { white-space: pre-wrap; font-family: Consolas, monospace; font-size: 13px; line-height: 1.5; margin: 0; }
    .third-party { color: #c00; font-weight: bold; }
    .header { margin-bottom: 20px; }
    .header h1 { margin: 0; font-size: 20px; }
    .header .meta { color: #666; font-size: 12px; margin-top: 4px; }
    .key { margin-top: 12px; padding: 8px 12px; background: #fff3cd; border-radius: 4px; font-size: 12px; }
    .key .third-party { color: #c00; font-weight: bold; }
  </style>
</head>
<body>
  <div class="header">
    <h1>$headerTitle - TICKET INSTRUCTIONS</h1>
    <div class="meta">Generated: $(Get-Date -Format 'yyyy-MM-dd HH:mm:ss')</div>
    <div class="key"><span class="third-party">Red</span> = Third-party software; requires authorization/quoting</div>
  </div>
  $sectionsHtml
  <script>
    function copySubject(btn) {
      var text = btn.getAttribute('data-subject');
      if (!text) return;
      navigator.clipboard.writeText(text).catch(function() { prompt('Copy this:', text); });
    }
    function copySection(id) {
      var el = document.getElementById(id);
      var pre = el ? el.querySelector('.section-text') : null;
      if (!pre) return;
      var text = pre.innerText || pre.textContent;
      navigator.clipboard.writeText(text).catch(function() { prompt('Copy this:', text); });
    }
    function toggleSection(id) {
      var el = document.getElementById(id);
      var btn = el ? el.querySelector('.toggle-btn') : null;
      if (!el || !btn) return;
      if (el.classList.contains('collapsed')) {
        el.classList.remove('collapsed');
        btn.textContent = '▾ Hide details';
      } else {
        el.classList.add('collapsed');
        btn.textContent = '▸ Show details';
      }
    }
  </script>
</body>
</html>
"@

        if ($OutputPath) {
            $html | Out-File -FilePath $OutputPath -Encoding UTF8
            Write-Log "Ticket instructions (HTML) saved to: $OutputPath" -Level Success
        }
        if ($PassThru) { return $html }

    } catch {
        Write-Log "Error generating ticket instructions (HTML): $($_.Exception.Message)" -Level Error
    }
}

function New-CombinedReportHtml {
    param(
        [string]$OutputPath,
        [array]$TopTenData,
        [array]$TimeEstimates = $null,
        [bool]$IsRMITPlus = $false,
        [array]$GeneralRecommendations = $null,
        [bool]$IncludeTicketInstructions = $true,
        [bool]$IncludeEmailTemplate = $false,
        [bool]$IncludeTimeEstimate = $false,
        [string]$FilterTopN = $null,
        [string]$CompanyName = $null
    )

    try {
        Write-Log "Generating combined report (HTML)..."

        $tabButtons = [System.Collections.ArrayList]::new()
        $tabPanels = [System.Collections.ArrayList]::new()
        $firstTab = $true

        # Tab 1: Ticket Instructions
        if ($IncludeTicketInstructions -and $TopTenData -and $TopTenData.Count -gt 0) {
            $ticketHtml = New-TicketInstructionsHtml -OutputPath $null -TopTenData $TopTenData -TimeEstimates $TimeEstimates -IsRMITPlus $IsRMITPlus -GeneralRecommendations $GeneralRecommendations -PassThru
            if ($ticketHtml -match '(?s)<body[^>]*>(.*)</body>') {
                $ticketBody = $matches[1].Trim()
            } else {
                $ticketBody = $ticketHtml
            }
            $activeClass = if ($firstTab) { ' active' } else { '' }
            $null = $tabButtons.Add("<button class=`"tab-btn$activeClass`" data-tab=`"ticket`">Ticket Instructions</button>")
            $null = $tabPanels.Add("<div id=`"panel-ticket`" class=`"tab-panel$activeClass`">$ticketBody</div>")
            $firstTab = $false
        }

        # Tab 2: Email Template
        if ($IncludeEmailTemplate) {
            $emailContent = New-EmailTemplate -OutputPath $null -IsRMITPlus $IsRMITPlus -FilterTopN $FilterTopN -PassThru
            $firstLine = ($emailContent -split "`r?`n")[0]
            $emailSubject = $firstLine -replace '^Subject:\s*', ''
            $emailSubjectEscaped = [System.Net.WebUtility]::HtmlEncode($emailSubject)
            $emailEscaped = [System.Net.WebUtility]::HtmlEncode($emailContent)
            $activeClass = if ($firstTab) { ' active' } else { '' }
            $null = $tabButtons.Add("<button class=`"tab-btn$activeClass`" data-tab=`"email`">Email Template</button>")
            $emailPanelHtml = "<div id=`"panel-email`" class=`"tab-panel$activeClass`" data-email-subject=`"$emailSubjectEscaped`"><div class=`"tab-actions`"><button type=`"button`" class=`"copy-btn`" onclick=`"copyEmailSubject(this)`">Copy Subject</button><button type=`"button`" class=`"copy-btn`" onclick=`"copyEmailBody()`">Copy Body</button></div><pre id=`"email-content`" class=`"tab-pre`">$emailEscaped</pre></div>"
            $null = $tabPanels.Add($emailPanelHtml)
            $firstTab = $false
        }

        # Tab 3: Time Estimate
        if ($IncludeTimeEstimate -and $TopTenData -and $TopTenData.Count -gt 0 -and $null -ne $TimeEstimates) {
            $timeContent = New-TimeEstimate -OutputPath $null -Top10Data $TopTenData -TimeEstimates $TimeEstimates -IsRMITPlus $IsRMITPlus -GeneralRecommendations $GeneralRecommendations -PassThru
            $timeEscaped = [System.Net.WebUtility]::HtmlEncode($timeContent)
            $activeClass = if ($firstTab) { ' active' } else { '' }
            $null = $tabButtons.Add("<button class=`"tab-btn$activeClass`" data-tab=`"time`">Time Estimate</button>")
            $timePanelHtml = "<div id=`"panel-time`" class=`"tab-panel$activeClass`"><div class=`"tab-actions`"><button type=`"button`" class=`"copy-btn`" onclick=`"copyTimeEstimate()`">Copy Time Estimate</button></div><pre id=`"time-estimate-content`" class=`"tab-pre`">$timeEscaped</pre></div>"
            $null = $tabPanels.Add($timePanelHtml)
            $firstTab = $false
        }

        # Tab 4: Ticket Notes (always include when we have Top10Data)
        if ($TopTenData -and $TopTenData.Count -gt 0) {
            $notesContent = New-TicketNotes -Top10Data $TopTenData -TimeEstimates $TimeEstimates -IsRMITPlus $IsRMITPlus -FilterTopN $FilterTopN -PassThru
            $notesEscaped = [System.Net.WebUtility]::HtmlEncode($notesContent)
            $activeClass = if ($firstTab) { ' active' } else { '' }
            $null = $tabButtons.Add("<button class=`"tab-btn$activeClass`" data-tab=`"notes`">Ticket Notes</button>")
            $notesPanelHtml = "<div id=`"panel-notes`" class=`"tab-panel$activeClass`"><div class=`"tab-actions`"><button type=`"button`" class=`"copy-btn`" onclick=`"copyTicketNotes()`">Copy Ticket Notes</button></div><pre id=`"ticket-notes-content`" class=`"tab-pre`">$notesEscaped</pre></div>"
            $null = $tabPanels.Add($notesPanelHtml)
        }

        $tabsHtml = $tabButtons -join "`n    "
        $panelsHtml = $tabPanels -join "`n  "

        $displayCompanyName = if ([string]::IsNullOrWhiteSpace($CompanyName)) {
            if ($script:UserSettings -and -not [string]::IsNullOrWhiteSpace($script:UserSettings.CompanyName)) { $script:UserSettings.CompanyName }
            else { "Client" }
        } else { $CompanyName }
        $companyNameEscaped = [System.Net.WebUtility]::HtmlEncode($displayCompanyName)

        $html = @"
<!DOCTYPE html>
<html lang="en">
<head>
  <meta charset="UTF-8">
  <meta name="viewport" content="width=device-width, initial-scale=1">
  <title>Vulnerability Report - $companyNameEscaped - $(Get-Date -Format 'yyyy-MM-dd HH:mm')</title>
  <style>
    body { font-family: Segoe UI, Arial, sans-serif; margin: 0; padding: 20px; background: #f5f5f5; }
    .report-header { margin-bottom: 20px; padding: 16px 20px; background: #fff; border-radius: 8px; box-shadow: 0 2px 8px rgba(0,0,0,0.08); }
    .report-header .company-name { font-size: 24px; font-weight: bold; color: #1a1a1a; margin: 0; }
    .report-header .report-meta { color: #666; font-size: 13px; margin-top: 6px; }
    .tab-bar { display: flex; gap: 4px; margin-bottom: 16px; border-bottom: 2px solid #ddd; }
    .tab-btn { padding: 10px 20px; cursor: pointer; background: #e9ecef; border: none; border-radius: 4px 4px 0 0; font-size: 14px; }
    .tab-btn:hover { background: #dee2e6; }
    .tab-btn.active { background: #fff; border: 1px solid #ddd; border-bottom: 2px solid #fff; margin-bottom: -2px; font-weight: bold; }
    .tab-panel { display: none; }
    .tab-panel.active { display: block; }
    .tab-actions { margin-bottom: 12px; }
    .tab-actions .copy-btn { margin-right: 8px; padding: 8px 16px; cursor: pointer; background: #0066cc; color: #fff; border: none; border-radius: 4px; font-size: 13px; }
    .tab-actions .copy-btn:hover { background: #0052a3; }
    .tab-pre { white-space: pre-wrap; font-family: Consolas, monospace; font-size: 13px; line-height: 1.5; margin: 0; padding: 20px; background: #fff; border-radius: 8px; }
    .vuln-section { background: #fff; padding: 20px; margin-bottom: 24px; border-radius: 8px; box-shadow: 0 2px 8px rgba(0,0,0,0.08); }
    .vuln-section h2 { margin: 0 0 12px 0; font-size: 18px; color: #333; border-bottom: 1px solid #ddd; padding-bottom: 8px; }
    .subject-line { font-size: 13px; color: #333; margin: 8px 0 12px 0; }
    .section-actions { margin-bottom: 0; }
    .section-actions button { margin-right: 8px; padding: 8px 16px; cursor: pointer; background: #0066cc; color: #fff; border: none; border-radius: 4px; font-size: 13px; }
    .section-actions button:hover { background: #0052a3; }
    .section-actions .toggle-btn { background: #6c757d; }
    .section-actions .toggle-btn:hover { background: #5a6268; }
    .section-content { margin-top: 16px; padding-top: 16px; border-top: 1px solid #eee; }
    .vuln-section.collapsed .section-content { display: none; }
    .section-text { white-space: pre-wrap; font-family: Consolas, monospace; font-size: 13px; line-height: 1.5; margin: 0; }
    .third-party { color: #c00; font-weight: bold; }
    .header { margin-bottom: 20px; }
    .header h1 { margin: 0; font-size: 20px; }
    .header .meta { color: #666; font-size: 12px; margin-top: 4px; }
    .key { margin-top: 12px; padding: 8px 12px; background: #fff3cd; border-radius: 4px; font-size: 12px; }
    .key .third-party { color: #c00; font-weight: bold; }
  </style>
</head>
<body>
  <div class="report-header">
    <h1 class="company-name">$companyNameEscaped</h1>
    <div class="report-meta">Vulnerability Report - $(Get-Date -Format 'MMMM d, yyyy')</div>
  </div>
  <div class="tab-bar">
    $tabsHtml
  </div>
  $panelsHtml
  <script>
    document.querySelectorAll('.tab-btn').forEach(function(btn) {
      btn.addEventListener('click', function() {
        var tab = this.getAttribute('data-tab');
        document.querySelectorAll('.tab-btn').forEach(function(b) { b.classList.remove('active'); });
        document.querySelectorAll('.tab-panel').forEach(function(p) { p.classList.remove('active'); });
        this.classList.add('active');
        var panel = document.getElementById('panel-' + tab);
        if (panel) panel.classList.add('active');
      });
    });
    function copySubject(btn) {
      var text = btn.getAttribute('data-subject');
      if (!text) return;
      navigator.clipboard.writeText(text).catch(function() { prompt('Copy this:', text); });
    }
    function copySection(id) {
      var el = document.getElementById(id);
      var pre = el ? el.querySelector('.section-text') : null;
      if (!pre) return;
      var text = pre.innerText || pre.textContent;
      navigator.clipboard.writeText(text).catch(function() { prompt('Copy this:', text); });
    }
    function toggleSection(id) {
      var el = document.getElementById(id);
      var btn = el ? el.querySelector('.toggle-btn') : null;
      if (!el || !btn) return;
      if (el.classList.contains('collapsed')) {
        el.classList.remove('collapsed');
        btn.textContent = '▾ Hide details';
      } else {
        el.classList.add('collapsed');
        btn.textContent = '▸ Show details';
      }
    }
    function copyEmailSubject(btn) {
      var panel = document.getElementById('panel-email');
      var subject = panel ? panel.getAttribute('data-email-subject') : '';
      if (subject) navigator.clipboard.writeText(subject).catch(function() { prompt('Copy this:', subject); });
    }
    function copyEmailBody() {
      var pre = document.getElementById('email-content');
      if (!pre) return;
      var full = pre.innerText || pre.textContent;
      var bodyStart = full.indexOf('\n\n');
      var body = bodyStart >= 0 ? full.substring(bodyStart + 2) : full;
      navigator.clipboard.writeText(body).catch(function() { prompt('Copy this:', body); });
    }
    function copyTimeEstimate() {
      var pre = document.getElementById('time-estimate-content');
      if (!pre) return;
      var text = pre.innerText || pre.textContent;
      navigator.clipboard.writeText(text).catch(function() { prompt('Copy this:', text); });
    }
    function copyTicketNotes() {
      var pre = document.getElementById('ticket-notes-content');
      if (!pre) return;
      var text = pre.innerText || pre.textContent;
      navigator.clipboard.writeText(text).catch(function() { prompt('Copy this:', text); });
    }
  </script>
</body>
</html>
"@

        if (-not [string]::IsNullOrWhiteSpace($OutputPath)) {
            $html | Out-File -FilePath $OutputPath -Encoding UTF8
            Write-Log "Combined report (HTML) saved to: $OutputPath" -Level Success
        }

    } catch {
        Write-Log "Error generating combined report (HTML): $($_.Exception.Message)" -Level Error
    }
}

function New-TicketNotes {
    param(
        [array]$Top10Data = $null,
        [array]$TimeEstimates = $null,
        [string]$OutputPath = $null,
        [bool]$IsRMITPlus = $false,
        [switch]$PassThru,
        [string]$FilterTopN = $null
    )

    # Use script variables if Top10Data not provided (backward compatibility)
    if ($null -eq $Top10Data) {
        $Top10Data = $script:CurrentTop10Data
        $TimeEstimates = $script:CurrentTimeEstimates
        $IsRMITPlus = $script:IsRMITPlus
        if ([string]::IsNullOrWhiteSpace($FilterTopN)) { $FilterTopN = $script:FilterTopN }
    } elseif ([string]::IsNullOrWhiteSpace($FilterTopN)) {
        $FilterTopN = $script:FilterTopN
    }
    if ([string]::IsNullOrWhiteSpace($FilterTopN)) { $FilterTopN = "10" }
    $topNLabel = if ($FilterTopN -eq "All") { "Top" } elseif ($FilterTopN -eq "10") { "Top Ten" } elseif (-not [string]::IsNullOrWhiteSpace($FilterTopN)) { "Top $FilterTopN" } else { "Top Ten" }
    $reportStepLine = "Produced $topNLabel vulnerabilities docx report"

    if ($null -eq $script:Templates) { Load-Templates }

    $stepsBeforeTickets = $script:Templates.TicketNotes.StepsBeforeTickets -replace '\{ReportStepLine\}', $reportStepLine

    # Collect ticket creation lines for vulnerabilities with tickets generated (incl. auto: 3rd party + after hours)
    $ticketLines = @()
    if ($null -ne $Top10Data -and $null -ne $TimeEstimates -and $TimeEstimates.Count -gt 0) {
        $timeByProduct = @{}
        foreach ($te in $TimeEstimates) { if (-not [string]::IsNullOrWhiteSpace($te.Product)) { $timeByProduct[$te.Product] = $te } }
        foreach ($item in $Top10Data) {
            $timeEstimate = if ($timeByProduct.ContainsKey($item.Product)) { $timeByProduct[$item.Product] } else { $null }
            if ($null -ne $timeEstimate -and $IsRMITPlus) {
                $autoTicketGenerated = $timeEstimate.ThirdParty -and $timeEstimate.AfterHours
                $isTicketGenerated = $timeEstimate.TicketGenerated -or $autoTicketGenerated
                if ($isTicketGenerated) {
                    $ticketLines += "- Ticket created for $($item.Product)"
                }
            }
        }
    }

    $stepsAfterTickets = $script:Templates.TicketNotes.StepsAfterTickets

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

$($script:Templates.TicketNotes.ResolvedQuestion)

$($script:Templates.TicketNotes.ResolvedAnswer)

$($script:Templates.TicketNotes.NextStepsQuestion)

$($script:Templates.TicketNotes.NextStepsText)
"@

    if ($PassThru) { return $result }

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

    # Double-click to edit
    $dataGridView.Add_CellDoubleClick({
        param($sender, $e)
        if ($e.RowIndex -ge 0) {
            Show-EditRuleDialog -RuleIndex $e.RowIndex
        }
    })

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

    $btnReset = New-Object System.Windows.Forms.Button
    $btnReset.Location = New-Object System.Drawing.Point(320, $y)
    $btnReset.Size = New-Object System.Drawing.Size(100, 30)
    $btnReset.Text = "Reset to Defaults"
    $btnReset.Add_Click({
        $result = [System.Windows.Forms.MessageBox]::Show("Replace all rules with app defaults? Unsaved changes will be lost.", "Reset to Defaults",
            [System.Windows.Forms.MessageBoxButtons]::YesNo, [System.Windows.Forms.MessageBoxIcon]::Question)
        if ($result -eq [System.Windows.Forms.DialogResult]::Yes) {
            $script:RemediationRules = Get-DefaultRemediationRules
            Refresh-RulesGrid
        }
    })
    $rulesForm.Controls.Add($btnReset)

    $btnExport = New-Object System.Windows.Forms.Button
    $btnExport.Location = New-Object System.Drawing.Point(430, $y)
    $btnExport.Size = New-Object System.Drawing.Size(75, 30)
    $btnExport.Text = "Export..."
    $btnExport.Add_Click({
        $saveDlg = New-Object System.Windows.Forms.SaveFileDialog
        $saveDlg.Filter = "JSON (*.json)|*.json|All Files (*.*)|*.*"
        $saveDlg.Title = "Export Remediation Rules"
        $saveDlg.FileName = "VScanMagic_RemediationRules.json"
        if ($saveDlg.ShowDialog() -eq [System.Windows.Forms.DialogResult]::OK) {
            try {
                $script:RemediationRules | ConvertTo-Json -Depth 10 | Set-Content -Path $saveDlg.FileName -Encoding UTF8
                [System.Windows.Forms.MessageBox]::Show("Rules exported to:`n$($saveDlg.FileName)", "Export Complete",
                    [System.Windows.Forms.MessageBoxButtons]::OK, [System.Windows.Forms.MessageBoxIcon]::Information)
            } catch {
                [System.Windows.Forms.MessageBox]::Show("Export failed: $($_.Exception.Message)", "Error",
                    [System.Windows.Forms.MessageBoxButtons]::OK, [System.Windows.Forms.MessageBoxIcon]::Error)
            }
        }
    })
    $rulesForm.Controls.Add($btnExport)

    $btnImport = New-Object System.Windows.Forms.Button
    $btnImport.Location = New-Object System.Drawing.Point(515, $y)
    $btnImport.Size = New-Object System.Drawing.Size(75, 30)
    $btnImport.Text = "Import..."
    $btnImport.Add_Click({
        $dlg = New-Object System.Windows.Forms.OpenFileDialog
        $dlg.Filter = "JSON (*.json)|*.json|All Files (*.*)|*.*"
        $dlg.Title = "Import Remediation Rules"
        if ($dlg.ShowDialog() -eq [System.Windows.Forms.DialogResult]::OK) {
            try {
                $json = Get-Content $dlg.FileName -Raw -Encoding UTF8 | ConvertFrom-Json
                $arr = if ($json -is [Array]) { @($json) } else { @($json) }
                $imported = @()
                foreach ($r in $arr) {
                    $imported += @{ Pattern = $r.Pattern; WordText = $r.WordText; TicketText = $r.TicketText; IsDefault = $r.IsDefault }
                }
                if ($imported.Count -gt 0) {
                    $result = [System.Windows.Forms.MessageBox]::Show("Replace current rules with imported rules ($($imported.Count) rules)?", "Import",
                        [System.Windows.Forms.MessageBoxButtons]::YesNo, [System.Windows.Forms.MessageBoxIcon]::Question)
                    if ($result -eq [System.Windows.Forms.DialogResult]::Yes) {
                        $script:RemediationRules = $imported
                        Refresh-RulesGrid
                        [System.Windows.Forms.MessageBox]::Show("Imported $($imported.Count) rules. Click Save to persist.", "Import Complete",
                            [System.Windows.Forms.MessageBoxButtons]::OK, [System.Windows.Forms.MessageBoxIcon]::Information)
                    }
                } else {
                    [System.Windows.Forms.MessageBox]::Show("No valid rules found in file.", "Import",
                        [System.Windows.Forms.MessageBoxButtons]::OK, [System.Windows.Forms.MessageBoxIcon]::Warning)
                }
            } catch {
                [System.Windows.Forms.MessageBox]::Show("Import failed: $($_.Exception.Message)", "Error",
                    [System.Windows.Forms.MessageBoxButtons]::OK, [System.Windows.Forms.MessageBoxIcon]::Error)
            }
        }
    })
    $rulesForm.Controls.Add($btnImport)

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
    $txtBaseUrl.Text = "https://pod0.myconnectsecure.com"
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
        $txtBaseUrl.Text = "https://pod0.myconnectsecure.com"
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

function Show-ConnectWiseAutomateSettingsDialog {
    $dlg = New-Object System.Windows.Forms.Form
    $dlg.Text = "ConnectWise Automate - Username Lookup"
    $dlg.Size = New-Object System.Drawing.Size(480, 260)
    $dlg.StartPosition = "CenterParent"
    $dlg.FormBorderStyle = "FixedDialog"
    $dlg.MaximizeBox = $false
    $dlg.MinimizeBox = $false

    $y = 20
    $lblHint = New-Object System.Windows.Forms.Label
    $lblHint.Location = New-Object System.Drawing.Point(20, $y)
    $lblHint.Size = New-Object System.Drawing.Size(430, 32)
    $lblHint.Text = "Configure ConnectWise Automate to look up Last Logged In User for hostnames in the Review Hostnames dialog. Uses AutomateAPI module or REST API."
    $lblHint.AutoSize = $false
    $lblHint.MaximumSize = New-Object System.Drawing.Size(430, 0)
    $dlg.Controls.Add($lblHint)
    $y += 40

    $lblBaseUrl = New-Object System.Windows.Forms.Label
    $lblBaseUrl.Location = New-Object System.Drawing.Point(20, $y)
    $lblBaseUrl.Size = New-Object System.Drawing.Size(100, 20)
    $lblBaseUrl.Text = "Automate URL:"
    $dlg.Controls.Add($lblBaseUrl)
    $txtBaseUrl = New-Object System.Windows.Forms.TextBox
    $txtBaseUrl.Location = New-Object System.Drawing.Point(130, $y)
    $txtBaseUrl.Size = New-Object System.Drawing.Size(310, 20)
    $txtBaseUrl.Text = "https://your-automate-server"
    $dlg.Controls.Add($txtBaseUrl)
    $y += 30

    $lblUser = New-Object System.Windows.Forms.Label
    $lblUser.Location = New-Object System.Drawing.Point(20, $y)
    $lblUser.Size = New-Object System.Drawing.Size(100, 20)
    $lblUser.Text = "Username:"
    $dlg.Controls.Add($lblUser)
    $txtUser = New-Object System.Windows.Forms.TextBox
    $txtUser.Location = New-Object System.Drawing.Point(130, $y)
    $txtUser.Size = New-Object System.Drawing.Size(200, 20)
    $dlg.Controls.Add($txtUser)
    $y += 30

    $lblPass = New-Object System.Windows.Forms.Label
    $lblPass.Location = New-Object System.Drawing.Point(20, $y)
    $lblPass.Size = New-Object System.Drawing.Size(100, 20)
    $lblPass.Text = "Password:"
    $dlg.Controls.Add($lblPass)
    $txtPass = New-Object System.Windows.Forms.TextBox
    $txtPass.Location = New-Object System.Drawing.Point(130, $y)
    $txtPass.Size = New-Object System.Drawing.Size(200, 20)
    $txtPass.PasswordChar = '*'
    $dlg.Controls.Add($txtPass)
    $y += 35

    $chkUseModule = New-Object System.Windows.Forms.CheckBox
    $chkUseModule.Location = New-Object System.Drawing.Point(20, $y)
    $chkUseModule.Size = New-Object System.Drawing.Size(400, 20)
    $chkUseModule.Text = "Use AutomateAPI PowerShell module (Install-Module AutomateAPI) if available"
    $chkUseModule.Checked = $true
    $dlg.Controls.Add($chkUseModule)
    $y += 35

    $btnLoad = New-Object System.Windows.Forms.Button
    $btnLoad.Location = New-Object System.Drawing.Point(20, $y)
    $btnLoad.Size = New-Object System.Drawing.Size(80, 28)
    $btnLoad.Text = "Load Saved"
    $btnLoad.Add_Click({
        $saved = Load-ConnectWiseAutomateCredentials
        if ($saved) {
            $txtBaseUrl.Text = $saved.BaseUrl
            $txtUser.Text = $saved.Username
            $txtPass.Text = $saved.Password
            $chkUseModule.Checked = $saved.UseAutomateModule
        }
    })
    $dlg.Controls.Add($btnLoad)

    $btnSave = New-Object System.Windows.Forms.Button
    $btnSave.Location = New-Object System.Drawing.Point(110, $y)
    $btnSave.Size = New-Object System.Drawing.Size(80, 28)
    $btnSave.Text = "Save"
    $btnSave.Add_Click({
        if (Save-ConnectWiseAutomateCredentials -BaseUrl $txtBaseUrl.Text.Trim() -Username $txtUser.Text -Password $txtPass.Text -UseAutomateModule $chkUseModule.Checked) {
            [System.Windows.Forms.MessageBox]::Show("ConnectWise Automate settings saved.", "Saved", [System.Windows.Forms.MessageBoxButtons]::OK, [System.Windows.Forms.MessageBoxIcon]::Information)
        }
    })
    $dlg.Controls.Add($btnSave)

    $btnClose = New-Object System.Windows.Forms.Button
    $btnClose.Location = New-Object System.Drawing.Point(360, $y)
    $btnClose.Size = New-Object System.Drawing.Size(80, 28)
    $btnClose.Text = "Close"
    $btnClose.DialogResult = [System.Windows.Forms.DialogResult]::OK
    $dlg.Controls.Add($btnClose)

    $saved = Load-ConnectWiseAutomateCredentials
    if ($saved) {
        $txtBaseUrl.Text = $saved.BaseUrl
        $txtUser.Text = $saved.Username
        $txtPass.Text = $saved.Password
        $chkUseModule.Checked = $saved.UseAutomateModule
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
    $dlg.Size = New-Object System.Drawing.Size(420, 290)
    $dlg.StartPosition = "CenterParent"
    $dlg.FormBorderStyle = "FixedDialog"
    $dlg.MaximizeBox = $false
    $dlg.MinimizeBox = $false

    $chkExcel = New-Object System.Windows.Forms.CheckBox
    $chkExcel.Location = New-Object System.Drawing.Point(20, 25)
    $chkExcel.Size = New-Object System.Drawing.Size(360, 20)
    $chkExcel.Text = "Generate Excel Report (consolidated pivot)"
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
    $chkTicket.Text = "Generate Ticket Instructions (Text + HTML)"
    $chkTicket.Checked = $script:OutputTicketInstructions
    $dlg.Controls.Add($chkTicket)

    $chkTime = New-Object System.Windows.Forms.CheckBox
    $chkTime.Location = New-Object System.Drawing.Point(20, 125)
    $chkTime.Size = New-Object System.Drawing.Size(360, 20)
    $chkTime.Text = "Generate Time Estimate (Text)"
    $chkTime.Checked = $script:OutputTimeEstimate
    $dlg.Controls.Add($chkTime)

    $chkAutoLookup = New-Object System.Windows.Forms.CheckBox
    $chkAutoLookup.Location = New-Object System.Drawing.Point(20, 155)
    $chkAutoLookup.Size = New-Object System.Drawing.Size(360, 20)
    $chkAutoLookup.Text = "Hostname Review - Look up usernames from ConnectSecure by default"
    $chkAutoLookup.Checked = if ($script:UserSettings.HostnameReviewAutoLookupConnectSecure) { $true } else { $false }
    $dlg.Controls.Add($chkAutoLookup)

    $chkAutoResize = New-Object System.Windows.Forms.CheckBox
    $chkAutoResize.Location = New-Object System.Drawing.Point(20, 178)
    $chkAutoResize.Size = New-Object System.Drawing.Size(360, 20)
    $chkAutoResize.Text = "Downloaded reports - Auto-resize Excel columns (excludes Company, Proposed Remediations)"
    $chkAutoResize.Checked = if ($script:UserSettings.DownloadAutoResizeColumns) { $true } else { $false }
    $dlg.Controls.Add($chkAutoResize)

    $lblWin11Thresh = New-Object System.Windows.Forms.Label
    $lblWin11Thresh.Location = New-Object System.Drawing.Point(20, 203)
    $lblWin11Thresh.Size = New-Object System.Drawing.Size(280, 20)
    $lblWin11Thresh.Text = "Hostname Review - Windows 11 O/S vuln threshold (below = unselected):"
    $dlg.Controls.Add($lblWin11Thresh)
    $numWin11Thresh = New-Object System.Windows.Forms.NumericUpDown
    $numWin11Thresh.Location = New-Object System.Drawing.Point(300, 201)
    $numWin11Thresh.Size = New-Object System.Drawing.Size(80, 22)
    $numWin11Thresh.Minimum = 0
    $numWin11Thresh.Maximum = 9999
    $numWin11Thresh.Value = $script:UserSettings.HostnameReviewWindows11Threshold
    $dlg.Controls.Add($numWin11Thresh)

    $btnOK = New-Object System.Windows.Forms.Button
    $btnOK.Location = New-Object System.Drawing.Point(210, 232)
    $btnOK.Size = New-Object System.Drawing.Size(90, 28)
    $btnOK.Text = "OK"
    $btnOK.DialogResult = [System.Windows.Forms.DialogResult]::OK
    $btnOK.Add_Click({
        $script:OutputExcel = $chkExcel.Checked
        $script:OutputWord = $chkWord.Checked
        $script:OutputEmailTemplate = $chkEmail.Checked
        $script:OutputTicketInstructions = $chkTicket.Checked
        $script:OutputTimeEstimate = $chkTime.Checked
        $script:UserSettings.HostnameReviewAutoLookupConnectSecure = $chkAutoLookup.Checked
        $script:UserSettings.DownloadAutoResizeColumns = $chkAutoResize.Checked
        $script:UserSettings.HostnameReviewWindows11Threshold = [int]$numWin11Thresh.Value
        Save-UserSettings | Out-Null
    })
    $dlg.Controls.Add($btnOK)
    $btnCancel = New-Object System.Windows.Forms.Button
    $btnCancel.Location = New-Object System.Drawing.Point(305, 232)
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
        $script:UserSettings.HostnameReviewAutoLookupConnectSecure = $chkAutoLookup.Checked
        $script:UserSettings.DownloadAutoResizeColumns = $chkAutoResize.Checked
        $script:UserSettings.HostnameReviewWindows11Threshold = [int]$numWin11Thresh.Value
        Save-UserSettings | Out-Null
    }
}

function Show-TemplatesDialog {
    if ($null -eq $script:Templates) { Load-Templates }

    $dlg = New-Object System.Windows.Forms.Form
    $dlg.Text = "Customize Templates"
    $dlg.Size = New-Object System.Drawing.Size(750, 550)
    $dlg.StartPosition = "CenterParent"
    $dlg.FormBorderStyle = "FixedDialog"
    $dlg.MaximizeBox = $false

    $tabControl = New-Object System.Windows.Forms.TabControl
    $tabControl.Location = New-Object System.Drawing.Point(12, 12)
    $tabControl.Size = New-Object System.Drawing.Size(710, 460)

    # --- Email Template Tab ---
    $tabEmail = New-Object System.Windows.Forms.TabPage
    $tabEmail.Text = "Email Template"
    $lblEmailHint = New-Object System.Windows.Forms.Label
    $lblEmailHint.Location = New-Object System.Drawing.Point(8, 8)
    $lblEmailHint.Size = New-Object System.Drawing.Size(650, 32)
    $lblEmailHint.Text = "Placeholders: {Year}, {Quarter}, {Greeting}, {NoteText}, {PreparedBy}, {TopNLabel}"
    $lblEmailHint.ForeColor = [System.Drawing.Color]::Gray
    $lblEmailHint.AutoSize = $false
    $tabEmail.Controls.Add($lblEmailHint)
    $txtEmail = New-Object System.Windows.Forms.TextBox
    $txtEmail.Location = New-Object System.Drawing.Point(8, 44)
    $txtEmail.Size = New-Object System.Drawing.Size(680, 360)
    $txtEmail.Multiline = $true
    $txtEmail.ScrollBars = "Vertical"
    $txtEmail.Font = New-Object System.Drawing.Font("Consolas", 9)
    $txtEmail.Text = $script:Templates.EmailTemplate.Body
    $tabEmail.Controls.Add($txtEmail)
    $tabControl.TabPages.Add($tabEmail)

    # --- Ticket Notes Tab ---
    $tabNotes = New-Object System.Windows.Forms.TabPage
    $tabNotes.Text = "Ticket Notes"
    $notesPanel = New-Object System.Windows.Forms.Panel
    $notesPanel.Location = New-Object System.Drawing.Point(0, 0)
    $notesPanel.Size = New-Object System.Drawing.Size(694, 420)
    $notesPanel.AutoScroll = $true
    $notesPanel.BorderStyle = [System.Windows.Forms.BorderStyle]::None
    $notesY = 8
    $notesGap = 6
    $notesLabelHeight = 18
    $lblNotesHint = New-Object System.Windows.Forms.Label
    $lblNotesHint.Location = New-Object System.Drawing.Point(8, $notesY)
    $lblNotesHint.Size = New-Object System.Drawing.Size(680, 40)
    $lblNotesHint.Text = "Placeholders: {ReportStepLine} (in steps before tickets). Ticket creation lines are inserted between Steps Before and Steps After."
    $lblNotesHint.ForeColor = [System.Drawing.Color]::Gray
    $lblNotesHint.AutoSize = $false
    $notesPanel.Controls.Add($lblNotesHint)
    $notesY += 40 + $notesGap
    $lblStepsBefore = New-Object System.Windows.Forms.Label
    $lblStepsBefore.Location = New-Object System.Drawing.Point(8, $notesY)
    $lblStepsBefore.Size = New-Object System.Drawing.Size(680, $notesLabelHeight)
    $lblStepsBefore.Text = "Steps Before Tickets:"
    $lblStepsBefore.AutoSize = $false
    $notesPanel.Controls.Add($lblStepsBefore)
    $notesY += $notesLabelHeight + $notesGap
    $txtStepsBefore = New-Object System.Windows.Forms.TextBox
    $txtStepsBefore.Location = New-Object System.Drawing.Point(8, $notesY)
    $txtStepsBefore.Size = New-Object System.Drawing.Size(680, 90)
    $txtStepsBefore.Multiline = $true
    $txtStepsBefore.ScrollBars = "Vertical"
    $txtStepsBefore.Font = New-Object System.Drawing.Font("Consolas", 9)
    $txtStepsBefore.Text = $script:Templates.TicketNotes.StepsBeforeTickets
    $notesPanel.Controls.Add($txtStepsBefore)
    $notesY += 90 + $notesGap
    $lblStepsAfter = New-Object System.Windows.Forms.Label
    $lblStepsAfter.Location = New-Object System.Drawing.Point(8, $notesY)
    $lblStepsAfter.Size = New-Object System.Drawing.Size(680, $notesLabelHeight)
    $lblStepsAfter.Text = "Steps After Tickets:"
    $lblStepsAfter.AutoSize = $false
    $notesPanel.Controls.Add($lblStepsAfter)
    $notesY += $notesLabelHeight + $notesGap
    $txtStepsAfter = New-Object System.Windows.Forms.TextBox
    $txtStepsAfter.Location = New-Object System.Drawing.Point(8, $notesY)
    $txtStepsAfter.Size = New-Object System.Drawing.Size(680, 55)
    $txtStepsAfter.Multiline = $true
    $txtStepsAfter.ScrollBars = "Vertical"
    $txtStepsAfter.Font = New-Object System.Drawing.Font("Consolas", 9)
    $txtStepsAfter.Text = $script:Templates.TicketNotes.StepsAfterTickets
    $notesPanel.Controls.Add($txtStepsAfter)
    $notesY += 55 + $notesGap
    $lblResolved = New-Object System.Windows.Forms.Label
    $lblResolved.Location = New-Object System.Drawing.Point(8, $notesY)
    $lblResolved.Size = New-Object System.Drawing.Size(680, $notesLabelHeight)
    $lblResolved.Text = "Resolved (Question / Answer):"
    $lblResolved.AutoSize = $false
    $notesPanel.Controls.Add($lblResolved)
    $notesY += $notesLabelHeight + $notesGap
    $txtResolvedQ = New-Object System.Windows.Forms.TextBox
    $txtResolvedQ.Location = New-Object System.Drawing.Point(8, $notesY)
    $txtResolvedQ.Size = New-Object System.Drawing.Size(330, 22)
    $txtResolvedQ.Text = $script:Templates.TicketNotes.ResolvedQuestion
    $notesPanel.Controls.Add($txtResolvedQ)
    $txtResolvedA = New-Object System.Windows.Forms.TextBox
    $txtResolvedA.Location = New-Object System.Drawing.Point(348, $notesY)
    $txtResolvedA.Size = New-Object System.Drawing.Size(340, 22)
    $txtResolvedA.Text = $script:Templates.TicketNotes.ResolvedAnswer
    $notesPanel.Controls.Add($txtResolvedA)
    $notesY += 22 + $notesGap
    $lblNextSteps = New-Object System.Windows.Forms.Label
    $lblNextSteps.Location = New-Object System.Drawing.Point(8, $notesY)
    $lblNextSteps.Size = New-Object System.Drawing.Size(680, $notesLabelHeight)
    $lblNextSteps.Text = "Next Steps (Question / Text):"
    $lblNextSteps.AutoSize = $false
    $notesPanel.Controls.Add($lblNextSteps)
    $notesY += $notesLabelHeight + $notesGap
    $txtNextStepsQ = New-Object System.Windows.Forms.TextBox
    $txtNextStepsQ.Location = New-Object System.Drawing.Point(8, $notesY)
    $txtNextStepsQ.Size = New-Object System.Drawing.Size(330, 22)
    $txtNextStepsQ.Text = $script:Templates.TicketNotes.NextStepsQuestion
    $notesPanel.Controls.Add($txtNextStepsQ)
    $notesY += 22 + $notesGap
    $txtNextStepsT = New-Object System.Windows.Forms.TextBox
    $txtNextStepsT.Location = New-Object System.Drawing.Point(8, $notesY)
    $txtNextStepsT.Size = New-Object System.Drawing.Size(680, 50)
    $txtNextStepsT.Multiline = $true
    $txtNextStepsT.Text = $script:Templates.TicketNotes.NextStepsText
    $notesPanel.Controls.Add($txtNextStepsT)
    $tabNotes.Controls.Add($notesPanel)
    $tabControl.TabPages.Add($tabNotes)

    $dlg.Controls.Add($tabControl)

    $btnSave = New-Object System.Windows.Forms.Button
    $btnSave.Location = New-Object System.Drawing.Point(12, 482)
    $btnSave.Size = New-Object System.Drawing.Size(100, 28)
    $btnSave.Text = "Save"
    $btnSave.Add_Click({
        $script:Templates.EmailTemplate.Body = $txtEmail.Text
        $script:Templates.TicketNotes.StepsBeforeTickets = $txtStepsBefore.Text
        $script:Templates.TicketNotes.StepsAfterTickets = $txtStepsAfter.Text
        $script:Templates.TicketNotes.ResolvedQuestion = $txtResolvedQ.Text
        $script:Templates.TicketNotes.ResolvedAnswer = $txtResolvedA.Text
        $script:Templates.TicketNotes.NextStepsQuestion = $txtNextStepsQ.Text
        $script:Templates.TicketNotes.NextStepsText = $txtNextStepsT.Text
        if (Save-Templates) {
            [System.Windows.Forms.MessageBox]::Show("Templates saved successfully.", "Saved", [System.Windows.Forms.MessageBoxButtons]::OK, [System.Windows.Forms.MessageBoxIcon]::Information)
        } else {
            [System.Windows.Forms.MessageBox]::Show("Failed to save templates.", "Error", [System.Windows.Forms.MessageBoxButtons]::OK, [System.Windows.Forms.MessageBoxIcon]::Warning)
        }
    })
    $dlg.Controls.Add($btnSave)

    $btnLoadDefaults = New-Object System.Windows.Forms.Button
    $btnLoadDefaults.Location = New-Object System.Drawing.Point(120, 482)
    $btnLoadDefaults.Size = New-Object System.Drawing.Size(120, 28)
    $btnLoadDefaults.Text = "Load Defaults"
    $btnLoadDefaults.Add_Click({
        $defaults = Get-DefaultTemplates
        $txtEmail.Text = $defaults.EmailTemplate.Body
        $txtStepsBefore.Text = $defaults.TicketNotes.StepsBeforeTickets
        $txtStepsAfter.Text = $defaults.TicketNotes.StepsAfterTickets
        $txtResolvedQ.Text = $defaults.TicketNotes.ResolvedQuestion
        $txtResolvedA.Text = $defaults.TicketNotes.ResolvedAnswer
        $txtNextStepsQ.Text = $defaults.TicketNotes.NextStepsQuestion
        $txtNextStepsT.Text = $defaults.TicketNotes.NextStepsText
    })
    $dlg.Controls.Add($btnLoadDefaults)

    $btnClose = New-Object System.Windows.Forms.Button
    $btnClose.Location = New-Object System.Drawing.Point(622, 482)
    $btnClose.Size = New-Object System.Drawing.Size(100, 28)
    $btnClose.Text = "Close"
    $btnClose.DialogResult = [System.Windows.Forms.DialogResult]::OK
    $btnClose.Add_Click({ $dlg.Close() })
    $dlg.Controls.Add($btnClose)

    $dlg.ShowDialog() | Out-Null
}

function Show-AIApiKeysDialog {
    $dlg = New-Object System.Windows.Forms.Form
    $dlg.Text = "AI API Keys"
    $dlg.Size = New-Object System.Drawing.Size(480, 220)
    $dlg.StartPosition = "CenterParent"
    $dlg.FormBorderStyle = "FixedDialog"
    $dlg.MaximizeBox = $false

    $y = 20
    $lblHint = New-Object System.Windows.Forms.Label
    $lblHint.Location = New-Object System.Drawing.Point(20, $y)
    $lblHint.Size = New-Object System.Drawing.Size(420, 48)
    $lblHint.Text = "Store API keys for AI-assisted report generation. Future use: email templates, ticket notes, remediation instructions, remediation guidance, and time estimate guidance (factors + human input). Keys are saved locally."
    $lblHint.ForeColor = [System.Drawing.Color]::Gray
    $lblHint.AutoSize = $false
    $lblHint.Font = New-Object System.Drawing.Font($lblHint.Font.FontFamily, 8.5)
    $dlg.Controls.Add($lblHint)
    $y += 52

    $providers = @(
        @{ Name = "Microsoft Copilot"; Key = "AIApiKeyCopilot" }
        @{ Name = "OpenAI (ChatGPT)"; Key = "AIApiKeyChatGPT" }
        @{ Name = "Anthropic (Claude)"; Key = "AIApiKeyClaude" }
    )
    $textBoxes = @{}
    foreach ($p in $providers) {
        $lbl = New-Object System.Windows.Forms.Label
        $lbl.Location = New-Object System.Drawing.Point(20, $y)
        $lbl.Size = New-Object System.Drawing.Size(140, 20)
        $lbl.Text = "$($p.Name):"
        $dlg.Controls.Add($lbl)
        $txt = New-Object System.Windows.Forms.TextBox
        $txt.Location = New-Object System.Drawing.Point(165, $y)
        $txt.Size = New-Object System.Drawing.Size(280, 22)
        $txt.PasswordChar = [char]'*'
        $txt.UseSystemPasswordChar = $true
        $val = $script:UserSettings[$p.Key]; $txt.Text = if ($val) { $val } else { "" }
        $dlg.Controls.Add($txt)
        $textBoxes[$p.Key] = $txt
        $y += 32
    }

    $btnSave = New-Object System.Windows.Forms.Button
    $btnSave.Location = New-Object System.Drawing.Point(165, $y)
    $btnSave.Size = New-Object System.Drawing.Size(90, 28)
    $btnSave.Text = "Save"
    $btnSave.Add_Click({
        $script:UserSettings.AIApiKeyCopilot = $textBoxes["AIApiKeyCopilot"].Text
        $script:UserSettings.AIApiKeyChatGPT = $textBoxes["AIApiKeyChatGPT"].Text
        $script:UserSettings.AIApiKeyClaude = $textBoxes["AIApiKeyClaude"].Text
        if (Save-UserSettings) {
            [System.Windows.Forms.MessageBox]::Show("AI API keys saved.", "Saved", [System.Windows.Forms.MessageBoxButtons]::OK, [System.Windows.Forms.MessageBoxIcon]::Information)
            $dlg.DialogResult = [System.Windows.Forms.DialogResult]::OK
        }
    })
    $dlg.Controls.Add($btnSave)

    $btnCancel = New-Object System.Windows.Forms.Button
    $btnCancel.Location = New-Object System.Drawing.Point(265, $y)
    $btnCancel.Size = New-Object System.Drawing.Size(80, 28)
    $btnCancel.Text = "Cancel"
    $btnCancel.DialogResult = [System.Windows.Forms.DialogResult]::Cancel
    $dlg.Controls.Add($btnCancel)

    $dlg.CancelButton = $btnCancel
    $dlg.AcceptButton = $btnSave
    $dlg.ShowDialog() | Out-Null
}

function Show-CompanyReviewDialog {
    <#
    .SYNOPSIS
    Displays 7 tenant configuration checks plus IP ranges, external assets, subnet issues, and offline agents.
    #>
    param([int]$CompanyId, [string]$CompanyName = '')

    if ($CompanyId -le 0) {
        [System.Windows.Forms.MessageBox]::Show("Select a single company (not All Companies) to run Company Review.", "Company Review", [System.Windows.Forms.MessageBoxButtons]::OK, [System.Windows.Forms.MessageBoxIcon]::Information)
        return
    }

    $creds = Load-ConnectSecureCredentials
    if (-not $creds -or [string]::IsNullOrWhiteSpace($creds.BaseUrl) -or [string]::IsNullOrWhiteSpace($creds.ClientId) -or [string]::IsNullOrWhiteSpace($creds.ClientSecret)) {
        [System.Windows.Forms.MessageBox]::Show("Please configure API credentials first. Click Settings > API Settings.", "Credentials Required", [System.Windows.Forms.MessageBoxButtons]::OK, [System.Windows.Forms.MessageBoxIcon]::Warning)
        return
    }

    $connected = Connect-ConnectSecureAPI -BaseUrl $creds.BaseUrl -TenantName $creds.TenantName -ClientId $creds.ClientId -ClientSecret $creds.ClientSecret
    if (-not $connected) {
        [System.Windows.Forms.MessageBox]::Show("Failed to connect to ConnectSecure. Check API Settings.", "Connection Error", [System.Windows.Forms.MessageBoxButtons]::OK, [System.Windows.Forms.MessageBoxIcon]::Error)
        return
    }

    try {
        $data = Get-ConnectSecureCompanyReviewData -CompanyId $CompanyId -CompanyName $CompanyName
    } catch {
        [System.Windows.Forms.MessageBox]::Show("Failed to fetch Company Review data: $($_.Exception.Message)", "Error", [System.Windows.Forms.MessageBoxButtons]::OK, [System.Windows.Forms.MessageBoxIcon]::Error)
        return
    }

    $dlg = New-Object System.Windows.Forms.Form
    $dlg.Text = "Company Review - $CompanyName"
    $dlg.Size = New-Object System.Drawing.Size(620, 780)
    $dlg.StartPosition = "CenterParent"
    $dlg.FormBorderStyle = "FixedDialog"
    $dlg.MaximizeBox = $false

    $panel = New-Object System.Windows.Forms.Panel
    $panel.Location = New-Object System.Drawing.Point(0, 0)
    $panel.Size = New-Object System.Drawing.Size(600, 690)
    $panel.AutoScroll = $true
    $panel.BorderStyle = [System.Windows.Forms.BorderStyle]::FixedSingle
    $dlg.Controls.Add($panel)

    $y = 12
    $fontB = New-Object System.Drawing.Font("Segoe UI", 9, [System.Drawing.FontStyle]::Bold)

    # --- 7 configuration checks ---
    $extCount = if ($data.ExternalAssets) { $data.ExternalAssets.Count } else { 0 }
    $off7 = 0; $off14 = 0; $off30 = 0
    if ($null -ne $data.AgentsOffline7PlusDays) { $off7 = [int]$data.AgentsOffline7PlusDays }
    if ($null -ne $data.AgentsOffline14PlusDays) { $off14 = [int]$data.AgentsOffline14PlusDays }
    if ($null -ne $data.AgentsOffline30PlusDays) { $off30 = [int]$data.AgentsOffline30PlusDays }
    $checks = @(
        @{ L = "1. Lightweight agents"; V = [string]$data.AgentCount; Ok = ($data.AgentCount -gt 0) }
        @{ L = "2. Probes w/ creds + networks"; V = [string]$data.ProbesWithBoth; Ok = ($data.ProbesWithBoth -gt 0) }
        @{ L = "3. External scan targets"; V = [string]$extCount; Ok = ($extCount -gt 0) }
        @{ L = "4. Offline (7d / 14d / 30+d)"; V = "$off7 / $off14 / $off30"; Ok = ($off30 -eq 0) }
        @{ L = "5. Firewall integration"; V = if ($data.FirewallActive) { $cnt = if ($data.FirewallCount -gt 0) { $data.FirewallCount } else { 0 }; $types = if ($data.FirewallType) { $data.FirewallType } else { "Unknown" }; "$cnt firewall(s): $types" } else { "Not configured" }; Ok = $data.FirewallActive }
        @{ L = "6. Last internal scan"; V = if ($data.LastInternalScan) { $data.LastInternalScan } else { "None" }; Ok = [bool]$data.LastInternalScan }
        @{ L = "7. Last external scan"; V = if ($data.LastExternalScan) { $data.LastExternalScan } else { "None" }; Ok = [bool]$data.LastExternalScan }
    )

    foreach ($c in $checks) {
        $lbl = New-Object System.Windows.Forms.Label
        $lbl.Location = New-Object System.Drawing.Point(12, $y)
        $lbl.AutoSize = $true
        $lbl.Text = $c.L
        $lbl.Font = $fontB
        $panel.Controls.Add($lbl)
        $val = New-Object System.Windows.Forms.Label
        $val.Location = New-Object System.Drawing.Point(220, $y)
        $val.Size = New-Object System.Drawing.Size(340, 20)
        $val.Text = $c.V
        $val.ForeColor = if ($c.Ok) { [System.Drawing.Color]::DarkGreen } else { [System.Drawing.Color]::DarkOrange }
        $panel.Controls.Add($val)
        $y += 22
    }
    $y += 8

    # --- IP Ranges / Probe Subnets ---
    $lblSubnets = New-Object System.Windows.Forms.Label
    $lblSubnets.Location = New-Object System.Drawing.Point(12, $y)
    $lblSubnets.Size = New-Object System.Drawing.Size(550, 18)
    $lblSubnets.Text = "IP Ranges / Probe Subnets:"
    $lblSubnets.Font = $fontB
    $panel.Controls.Add($lblSubnets)
    $y += 20

    $subnetLines = @()
    if ($data.ProbesSubnets -and $data.ProbesSubnets.Count -gt 0) {
        foreach ($s in $data.ProbesSubnets) { $subnetLines += [string]$s }
    }
    if ($data.ScanTargets -and $data.ScanTargets.Count -gt 0) {
        foreach ($t in $data.ScanTargets) {
            $ts = [string]$t
            if ($subnetLines -notcontains $ts) { $subnetLines += $ts }
        }
    }
    $subnetText = if ($subnetLines.Count -gt 0) { ($subnetLines | Sort-Object) -join "`r`n" } else { "(none)" }

    $txtSubnets = New-Object System.Windows.Forms.TextBox
    $txtSubnets.Location = New-Object System.Drawing.Point(12, $y)
    $txtSubnets.Size = New-Object System.Drawing.Size(550, 150)
    $txtSubnets.Multiline = $true
    $txtSubnets.ReadOnly = $true
    $txtSubnets.ScrollBars = [System.Windows.Forms.ScrollBars]::Vertical
    $txtSubnets.Text = $subnetText
    $txtSubnets.Font = New-Object System.Drawing.Font("Consolas", 8.5)
    $panel.Controls.Add($txtSubnets)
    $y += 156

    # --- External Assets (name + address) ---
    $lblExt = New-Object System.Windows.Forms.Label
    $lblExt.Location = New-Object System.Drawing.Point(12, $y)
    $lblExt.Size = New-Object System.Drawing.Size(550, 18)
    $lblExt.Text = "External Scan Targets (Name / Address):"
    $lblExt.Font = $fontB
    $panel.Controls.Add($lblExt)
    $y += 20

    $extLines = @()
    if ($data.ExternalAssets -and $data.ExternalAssets.Count -gt 0) {
        foreach ($ea in $data.ExternalAssets) {
            $n = if ($ea.Name) { $ea.Name } else { "(unnamed)" }
            $a = if ($ea.Address) { $ea.Address } else { "" }
            $extLines += "$n : $a"
        }
    }
    $extText = if ($extLines.Count -gt 0) { $extLines -join "`r`n" } else { "(none)" }

    $txtExt = New-Object System.Windows.Forms.TextBox
    $txtExt.Location = New-Object System.Drawing.Point(12, $y)
    $txtExt.Size = New-Object System.Drawing.Size(550, 150)
    $txtExt.Multiline = $true
    $txtExt.ReadOnly = $true
    $txtExt.ScrollBars = [System.Windows.Forms.ScrollBars]::Vertical
    $txtExt.Text = $extText
    $txtExt.Font = New-Object System.Drawing.Font("Consolas", 8.5)
    $panel.Controls.Add($txtExt)
    $y += 156

    # --- Probe Agent(s) Nmap Interface (IP / Port) ---
    $lblProbeNmap = New-Object System.Windows.Forms.Label
    $lblProbeNmap.Location = New-Object System.Drawing.Point(12, $y)
    $lblProbeNmap.Size = New-Object System.Drawing.Size(550, 18)
    $lblProbeNmap.Text = "Probe Agent(s) Nmap Interface (IP / Port):"
    $lblProbeNmap.Font = $fontB
    $panel.Controls.Add($lblProbeNmap)
    $y += 20

    $probeNmapLines = @()
    if ($data.ProbeAgentsNmapInfo -and $data.ProbeAgentsNmapInfo.Count -gt 0) {
        foreach ($p in $data.ProbeAgentsNmapInfo) {
            $line = "$($p.HostName) - IP: $($p.IP)"
            if ($p.NmapInterface -and $p.NmapInterface -ne '(not set)') {
                $line += " | Nmap Interface: $($p.NmapInterface)"
            } else {
                $line += " | Nmap Interface: (not set)"
            }
            if ($p.Port) {
                $line += " | Port: $($p.Port)"
            }
            $probeNmapLines += $line
        }
    }
    $probeNmapText = if ($probeNmapLines.Count -gt 0) { $probeNmapLines -join "`r`n" } else { "(none - no probe agents)" }

    $txtProbeNmap = New-Object System.Windows.Forms.TextBox
    $txtProbeNmap.Location = New-Object System.Drawing.Point(12, $y)
    $txtProbeNmap.Size = New-Object System.Drawing.Size(550, 80)
    $txtProbeNmap.Multiline = $true
    $txtProbeNmap.ReadOnly = $true
    $txtProbeNmap.ScrollBars = [System.Windows.Forms.ScrollBars]::Vertical
    $txtProbeNmap.Text = $probeNmapText
    $txtProbeNmap.Font = New-Object System.Drawing.Font("Consolas", 8.5)
    $panel.Controls.Add($txtProbeNmap)
    $y += 86

    # --- Subnet Issues ---
    if ($data.SubnetIssues -and $data.SubnetIssues.Count -gt 0) {
        $lblIssues = New-Object System.Windows.Forms.Label
        $lblIssues.Location = New-Object System.Drawing.Point(12, $y)
        $lblIssues.Size = New-Object System.Drawing.Size(550, 18)
        $lblIssues.Text = "Subnet Configuration Issues:"
        $lblIssues.Font = $fontB
        $lblIssues.ForeColor = [System.Drawing.Color]::DarkRed
        $panel.Controls.Add($lblIssues)
        $y += 20
        $issueText = ($data.SubnetIssues | ForEach-Object { [string]$_ }) -join "`r`n"
        $txtIssues = New-Object System.Windows.Forms.TextBox
        $txtIssues.Location = New-Object System.Drawing.Point(12, $y)
        $txtIssues.Size = New-Object System.Drawing.Size(550, 40)
        $txtIssues.Multiline = $true
        $txtIssues.ReadOnly = $true
        $txtIssues.Text = $issueText
        $txtIssues.ForeColor = [System.Drawing.Color]::DarkRed
        $txtIssues.Font = New-Object System.Drawing.Font("Consolas", 8.5)
        $panel.Controls.Add($txtIssues)
        $y += 46
    }

    # --- Offline agent names ---
    if ($data.AgentsOffline30PlusNames -and $data.AgentsOffline30PlusNames.Count -gt 0) {
        $offCount = $data.AgentsOffline30PlusNames.Count
        $lblOff = New-Object System.Windows.Forms.Label
        $lblOff.Location = New-Object System.Drawing.Point(12, $y)
        $lblOff.Size = New-Object System.Drawing.Size(550, 18)
        $lblOff.Text = "Offline 30+ days ($offCount):"
        $lblOff.Font = $fontB
        $lblOff.ForeColor = [System.Drawing.Color]::DarkOrange
        $panel.Controls.Add($lblOff)
        $y += 20
        $offText = ($data.AgentsOffline30PlusNames | ForEach-Object { [string]$_ }) -join "`r`n"
        $txtOff = New-Object System.Windows.Forms.TextBox
        $txtOff.Location = New-Object System.Drawing.Point(12, $y)
        $txtOff.Size = New-Object System.Drawing.Size(550, [Math]::Min(200, 40 + ($offCount * 16)))
        $txtOff.Multiline = $true
        $txtOff.ReadOnly = $true
        $txtOff.ScrollBars = [System.Windows.Forms.ScrollBars]::Vertical
        $txtOff.Text = $offText
        $txtOff.Font = New-Object System.Drawing.Font("Consolas", 8.5)
        $panel.Controls.Add($txtOff)
        $offH = [Math]::Min(200, 40 + ($offCount * 16))
        $y += $offH + 8
    }

    # --- Quick Wins ---
    if ($data.QuickWins -and $data.QuickWins.Count -gt 0) {
        $y += 4
        $lblQ = New-Object System.Windows.Forms.Label
        $lblQ.Location = New-Object System.Drawing.Point(12, $y)
        $lblQ.Size = New-Object System.Drawing.Size(550, 18)
        $lblQ.Text = "Recommendations:"
        $lblQ.Font = $fontB
        $panel.Controls.Add($lblQ)
        $y += 22
        foreach ($qw in $data.QuickWins) {
            $l = New-Object System.Windows.Forms.Label
            $l.Location = New-Object System.Drawing.Point(16, $y)
            $l.Size = New-Object System.Drawing.Size(540, 16)
            $l.Text = "• $qw"
            $l.AutoSize = $false
            $panel.Controls.Add($l)
            $y += 18
        }
    }

    $btnClose = New-Object System.Windows.Forms.Button
    $btnClose.Anchor = [System.Windows.Forms.AnchorStyles]::Bottom
    $btnClose.Location = New-Object System.Drawing.Point(250, 698)
    $btnClose.Size = New-Object System.Drawing.Size(100, 28)
    $btnClose.Text = "Close"
    $btnClose.DialogResult = [System.Windows.Forms.DialogResult]::OK
    $dlg.Controls.Add($btnClose)
    $dlg.AcceptButton = $btnClose

    $dlg.ShowDialog() | Out-Null
}

function Show-ProcessingSummaryDialog {
    param([array]$ProcessedOutputs = @())
    if (-not $ProcessedOutputs -or $ProcessedOutputs.Count -eq 0) { return }
    $dlg = New-Object System.Windows.Forms.Form
    $dlg.Text = "Processing Summary - Output Folders"
    $dlg.Size = New-Object System.Drawing.Size(700, 380)
    $dlg.StartPosition = "CenterParent"
    $dlg.FormBorderStyle = "FixedDialog"
    $dlg.MaximizeBox = $false

    $lblHint = New-Object System.Windows.Forms.Label
    $lblHint.Location = New-Object System.Drawing.Point(20, 12)
    $lblHint.Size = New-Object System.Drawing.Size(640, 32)
    $lblHint.Text = "Reports saved to the following quarter folders. Double-click a row or click Open to open the folder in Explorer."
    $lblHint.AutoSize = $false
    $lblHint.MaximumSize = New-Object System.Drawing.Size(640, 0)
    $dlg.Controls.Add($lblHint)

    $listView = New-Object System.Windows.Forms.ListView
    $listView.Location = New-Object System.Drawing.Point(20, 50)
    $listView.Size = New-Object System.Drawing.Size(640, 240)
    $listView.View = [System.Windows.Forms.View]::Details
    $listView.FullRowSelect = $true
    $listView.GridLines = $true
    $listView.Columns.Add("Company", 140) | Out-Null
    $listView.Columns.Add("Output folder (year-quarter)", 480) | Out-Null
    foreach ($item in $ProcessedOutputs) {
        $row = New-Object System.Windows.Forms.ListViewItem($item.CompanyName)
        $row.SubItems.Add($item.OutputPath) | Out-Null
        $row.Tag = $item.OutputPath
        if (Test-Path $item.OutputPath) {
            $listView.Items.Add($row) | Out-Null
        }
    }
    $dlg.Controls.Add($listView)

    $openSelected = {
        $sel = $listView.SelectedItems
        if ($sel -and $sel.Count -gt 0 -and $sel[0].Tag) {
            $path = [System.IO.Path]::GetFullPath($sel[0].Tag)
            if (Test-Path -LiteralPath $path -PathType Container) {
                Invoke-Item -LiteralPath $path
            }
        }
    }
    $listView.Add_DoubleClick($openSelected)

    $btnOpen = New-Object System.Windows.Forms.Button
    $btnOpen.Location = New-Object System.Drawing.Point(20, 300)
    $btnOpen.Size = New-Object System.Drawing.Size(100, 28)
    $btnOpen.Text = "Open folder..."
    $btnOpen.Add_Click($openSelected)
    $dlg.Controls.Add($btnOpen)

    $btnOpenTicketHtml = New-Object System.Windows.Forms.Button
    $btnOpenTicketHtml.Location = New-Object System.Drawing.Point(130, 300)
    $btnOpenTicketHtml.Size = New-Object System.Drawing.Size(160, 28)
    $btnOpenTicketHtml.Text = "Open Report"
    $ttTicket = New-Object System.Windows.Forms.ToolTip
    $ttTicket.SetToolTip($btnOpenTicketHtml, "Opens the Report (Ticket Instructions, Email Template, Time Estimate) with tabs.")
    $btnOpenTicketHtml.Add_Click({
        $sel = $listView.SelectedItems
        if ($sel -and $sel.Count -gt 0 -and $sel[0].Tag) {
            $path = [System.IO.Path]::GetFullPath($sel[0].Tag)
            $miscPath = Join-Path $path "Misc"
            $searchPaths = @($path)
            if (Test-Path -LiteralPath $miscPath -PathType Container) { $searchPaths = @($miscPath, $path) }
            $htmlFile = $null
            foreach ($dir in $searchPaths) {
                $htmlFile = Get-ChildItem -LiteralPath $dir -Filter "*.html" -ErrorAction SilentlyContinue | Where-Object { $_.Name -like "*Report*" -or $_.Name -like "*Ticket Instructions*" } | Sort-Object LastWriteTime -Descending | Select-Object -First 1
                if ($htmlFile) { break }
            }
            if ($htmlFile -and (Test-Path -LiteralPath $htmlFile.FullName)) {
                Invoke-Item -LiteralPath $htmlFile.FullName
            } else {
                [System.Windows.Forms.MessageBox]::Show("No Report file found in this folder or Misc subfolder.", "Not Found", [System.Windows.Forms.MessageBoxButtons]::OK, [System.Windows.Forms.MessageBoxIcon]::Information)
            }
        }
    })
    $dlg.Controls.Add($btnOpenTicketHtml)

    $btnOpenInOutlook = New-Object System.Windows.Forms.Button
    $btnOpenInOutlook.Location = New-Object System.Drawing.Point(300, 300)
    $btnOpenInOutlook.Size = New-Object System.Drawing.Size(200, 28)
    $btnOpenInOutlook.Text = "Open email (Classic Outlook)"
    $ttOutlook = New-Object System.Windows.Forms.ToolTip
    $ttOutlook.SetToolTip($btnOpenInOutlook, "Opens email as draft in Classic Outlook. New Outlook does not support this - use the .eml file or shortcut in the folder instead.")
    $btnOpenInOutlook.Add_Click({
        $sel = $listView.SelectedItems
        if ($sel -and $sel.Count -gt 0 -and $sel[0].Tag) {
            $path = [System.IO.Path]::GetFullPath($sel[0].Tag)
            if (Open-EmailDraftInOutlook -OutputFolderPath $path) {
                # Success
            } else {
                [System.Windows.Forms.MessageBox]::Show("This requires Classic Outlook (not New Outlook).`n`nUse the .eml file or ""Open in Email"" shortcut in the folder instead - double-click to open the draft.", "Open in Outlook", [System.Windows.Forms.MessageBoxButtons]::OK, [System.Windows.Forms.MessageBoxIcon]::Information)
            }
        }
    })
    $dlg.Controls.Add($btnOpenInOutlook)

    $btnClose = New-Object System.Windows.Forms.Button
    $btnClose.Location = New-Object System.Drawing.Point(510, 300)
    $btnClose.Size = New-Object System.Drawing.Size(100, 28)
    $btnClose.Text = "Close"
    $btnClose.DialogResult = [System.Windows.Forms.DialogResult]::OK
    $dlg.Controls.Add($btnClose)

    $dlg.AcceptButton = $btnClose
    $dlg.ShowDialog() | Out-Null
}

function Show-ReportFolderHistoryDialog {
    $entries = @(Get-ReportFolderHistory)
    if ($entries.Count -eq 0) {
        [System.Windows.Forms.MessageBox]::Show("No report folder history yet. History is built as you process reports.", "Report Folder History", [System.Windows.Forms.MessageBoxButtons]::OK, [System.Windows.Forms.MessageBoxIcon]::Information)
        return
    }
    $dlg = New-Object System.Windows.Forms.Form
    $dlg.Text = "Report Folder History"
    $dlg.Size = New-Object System.Drawing.Size(720, 420)
    $dlg.StartPosition = "CenterParent"
    $dlg.FormBorderStyle = "FixedDialog"
    $dlg.MaximizeBox = $false

    $lblHint = New-Object System.Windows.Forms.Label
    $lblHint.Location = New-Object System.Drawing.Point(20, 12)
    $lblHint.Size = New-Object System.Drawing.Size(660, 32)
    $lblHint.Text = "Previously processed report folders. Double-click a row or click Open to open the folder in Explorer. Use Remove to clear an entry from history."
    $lblHint.AutoSize = $false
    $lblHint.MaximumSize = New-Object System.Drawing.Size(660, 0)
    $dlg.Controls.Add($lblHint)

    $listView = New-Object System.Windows.Forms.ListView
    $listView.Location = New-Object System.Drawing.Point(20, 50)
    $listView.Size = New-Object System.Drawing.Size(660, 270)
    $listView.View = [System.Windows.Forms.View]::Details
    $listView.FullRowSelect = $true
    $listView.GridLines = $true
    $listView.Columns.Add("Company", 130) | Out-Null
    $listView.Columns.Add("Processed", 90) | Out-Null
    $listView.Columns.Add("Output folder", 420) | Out-Null
    foreach ($item in $entries) {
        $row = New-Object System.Windows.Forms.ListViewItem($item.CompanyName)
        $row.SubItems.Add($item.ProcessedAt) | Out-Null
        $row.SubItems.Add($item.OutputPath) | Out-Null
        $row.Tag = $item
        $listView.Items.Add($row) | Out-Null
    }
    $dlg.Controls.Add($listView)

    $openSelected = {
        $sel = $listView.SelectedItems
        if ($sel -and $sel.Count -gt 0 -and $sel[0].Tag -and $sel[0].Tag.OutputPath) {
            $path = [System.IO.Path]::GetFullPath($sel[0].Tag.OutputPath)
            if (Test-Path -LiteralPath $path -PathType Container) {
                Invoke-Item -LiteralPath $path
            }
        }
    }
    $listView.Add_DoubleClick($openSelected)

    $btnOpen = New-Object System.Windows.Forms.Button
    $btnOpen.Location = New-Object System.Drawing.Point(20, 330)
    $btnOpen.Size = New-Object System.Drawing.Size(100, 28)
    $btnOpen.Text = "Open folder..."
    $btnOpen.Add_Click($openSelected)
    $dlg.Controls.Add($btnOpen)

    $btnRemove = New-Object System.Windows.Forms.Button
    $btnRemove.Location = New-Object System.Drawing.Point(130, 330)
    $btnRemove.Size = New-Object System.Drawing.Size(100, 28)
    $btnRemove.Text = "Remove from history"
    $btnRemove.Add_Click({
        $sel = $listView.SelectedItems
        if (-not $sel -or $sel.Count -eq 0) { return }
        $toRemove = $sel[0].Tag.OutputPath
        $newEntries = @(Get-ReportFolderHistory | Where-Object { $_.OutputPath -ne $toRemove })
        $obj = @{ Entries = @($newEntries) }
        Set-JsonFile -Path $script:ReportFolderHistoryPath -Object $obj -Depth 3 | Out-Null
        $listView.Items.Remove($sel[0])
    })
    $dlg.Controls.Add($btnRemove)

    $btnClearAll = New-Object System.Windows.Forms.Button
    $btnClearAll.Location = New-Object System.Drawing.Point(240, 330)
    $btnClearAll.Size = New-Object System.Drawing.Size(120, 28)
    $btnClearAll.Text = "Clear all history"
    $btnClearAll.Add_Click({
        if ([System.Windows.Forms.MessageBox]::Show("Clear all report folder history?", "Confirm", [System.Windows.Forms.MessageBoxButtons]::YesNo, [System.Windows.Forms.MessageBoxIcon]::Question) -eq [System.Windows.Forms.DialogResult]::Yes) {
            $obj = @{ Entries = @() }
            Set-JsonFile -Path $script:ReportFolderHistoryPath -Object $obj -Depth 3 | Out-Null
            $listView.Items.Clear()
        }
    })
    $dlg.Controls.Add($btnClearAll)

    $btnClose = New-Object System.Windows.Forms.Button
    $btnClose.Location = New-Object System.Drawing.Point(580, 330)
    $btnClose.Size = New-Object System.Drawing.Size(100, 28)
    $btnClose.Text = "Close"
    $btnClose.DialogResult = [System.Windows.Forms.DialogResult]::OK
    $dlg.Controls.Add($btnClose)

    $dlg.AcceptButton = $btnClose
    $dlg.ShowDialog() | Out-Null
}

function Show-CompanyFolderMappingDialog {
    if ([string]::IsNullOrWhiteSpace($script:UserSettings.ReportsBasePath) -or -not (Test-Path $script:UserSettings.ReportsBasePath)) {
        [System.Windows.Forms.MessageBox]::Show("Reports Base Path must be set in Settings first.", "Folder Mappings", [System.Windows.Forms.MessageBoxButtons]::OK, [System.Windows.Forms.MessageBoxIcon]::Information)
        return
    }
    Load-CompanyFolderMap | Out-Null
    $base = $script:UserSettings.ReportsBasePath.Trim()
    $companyNames = @{}
    $creds = Load-ConnectSecureCredentials
    if ($creds) {
        $cached = Load-ConnectSecureCompaniesCache -BaseUrl $creds.BaseUrl -TenantName $creds.TenantName
        if ($cached) {
            foreach ($c in $cached) { $companyNames["$($c.Id)"] = ($c.DisplayName -replace '\s*\(ID:\s*\d+\)\s*$', '').Trim() }
        }
    }
    $companyNames["0"] = "Global"

    $dlg = New-Object System.Windows.Forms.Form
    $dlg.Text = "Company Folder Mappings"
    $dlg.Size = New-Object System.Drawing.Size(620, 420)
    $dlg.StartPosition = "CenterParent"
    $dlg.FormBorderStyle = "FixedDialog"
    $dlg.MaximizeBox = $false

    $lblHint = New-Object System.Windows.Forms.Label
    $lblHint.Location = New-Object System.Drawing.Point(20, 12)
    $lblHint.Size = New-Object System.Drawing.Size(560, 32)
    $lblHint.Text = "Map ConnectSecure companies to folder paths under the Reports Base Path. Select a mapping and click Edit or Remove."
    $lblHint.AutoSize = $false
    $lblHint.MaximumSize = New-Object System.Drawing.Size(560, 0)
    $dlg.Controls.Add($lblHint)

    $listView = New-Object System.Windows.Forms.ListView
    $listView.Location = New-Object System.Drawing.Point(20, 50)
    $listView.Size = New-Object System.Drawing.Size(560, 260)
    $listView.View = [System.Windows.Forms.View]::Details
    $listView.FullRowSelect = $true
    $listView.GridLines = $true
    $listView.Columns.Add("Company", 120) | Out-Null
    $listView.Columns.Add("Folder path", 420) | Out-Null
    $dlg.Controls.Add($listView)

    # Normalize stored paths to include Network Documentation\Vulnerability Scans (fixes legacy mappings)
    $needsSave = $false
    foreach ($key in @($script:CompanyFolderMap.Keys)) {
        $current = $script:CompanyFolderMap[$key]
        $normalized = Resolve-VulnerabilityScansSubpath -FolderName $current
        if ($current -ne $normalized) {
            $script:CompanyFolderMap[$key] = $normalized
            $needsSave = $true
        }
    }
    if ($needsSave) { Save-CompanyFolderMap | Out-Null }

    $script:RefreshMappings = {
        $listView.Items.Clear()
        $keys = @($script:CompanyFolderMap.Keys | Sort-Object { $_ })
        foreach ($key in $keys) {
            $name = if ($companyNames.ContainsKey($key)) { $companyNames[$key] } else { "Company $key" }
            $path = Resolve-VulnerabilityScansSubpath -FolderName $script:CompanyFolderMap[$key]
            $item = New-Object System.Windows.Forms.ListViewItem($name)
            $item.SubItems.Add($path) | Out-Null
            $item.Tag = $key
            $listView.Items.Add($item) | Out-Null
        }
    }
    & $script:RefreshMappings

    $btnEdit = New-Object System.Windows.Forms.Button
    $btnEdit.Location = New-Object System.Drawing.Point(20, 320)
    $btnEdit.Size = New-Object System.Drawing.Size(100, 28)
    $btnEdit.Text = "Edit..."
    $btnEdit.Add_Click({
        $sel = $listView.SelectedItems
        if (-not $sel -or $sel.Count -eq 0) {
            [System.Windows.Forms.MessageBox]::Show("Please select a mapping to edit.", "Edit", [System.Windows.Forms.MessageBoxButtons]::OK, [System.Windows.Forms.MessageBoxIcon]::Information)
            return
        }
        $key = $sel[0].Tag
        Add-Type -AssemblyName System.Windows.Forms -ErrorAction SilentlyContinue
        $folderBrowser = New-Object System.Windows.Forms.FolderBrowserDialog
        $folderBrowser.Description = "Select the Vulnerability Scans folder for this company (e.g. ...\Company\Network Documentation\Vulnerability Scans). The YYYY - QN folder will be created inside."
        $folderBrowser.SelectedPath = $base
        $currentPath = Join-Path $base $script:CompanyFolderMap[$key]
        if (Test-Path $currentPath) { $folderBrowser.SelectedPath = $currentPath }
        if ($folderBrowser.ShowDialog() -eq [System.Windows.Forms.DialogResult]::OK) {
            $selectedFull = [System.IO.Path]::GetFullPath($folderBrowser.SelectedPath)
            if ($selectedFull.StartsWith($base, [StringComparison]::OrdinalIgnoreCase)) {
                $rel = $selectedFull.Substring($base.Length).TrimStart([char]'\', [char]'/')
                $folderName = if ([string]::IsNullOrWhiteSpace($rel)) { [System.IO.Path]::GetFileName($selectedFull) } else { $rel }
            } else {
                $folderName = [System.IO.Path]::GetFileName($selectedFull)
            }
            $folderName = Resolve-VulnerabilityScansSubpath -FolderName $folderName
            $script:CompanyFolderMap[$key] = $folderName
            Save-CompanyFolderMap | Out-Null
            & $script:RefreshMappings
        }
    })
    $dlg.Controls.Add($btnEdit)

    $btnRemove = New-Object System.Windows.Forms.Button
    $btnRemove.Location = New-Object System.Drawing.Point(130, 320)
    $btnRemove.Size = New-Object System.Drawing.Size(100, 28)
    $btnRemove.Text = "Remove"
    $btnRemove.Add_Click({
        $sel = $listView.SelectedItems
        if (-not $sel -or $sel.Count -eq 0) {
            [System.Windows.Forms.MessageBox]::Show("Please select a mapping to remove.", "Remove", [System.Windows.Forms.MessageBoxButtons]::OK, [System.Windows.Forms.MessageBoxIcon]::Information)
            return
        }
        $key = $sel[0].Tag
        $script:CompanyFolderMap.Remove($key) | Out-Null
        Save-CompanyFolderMap | Out-Null
        & $script:RefreshMappings
    })
    $dlg.Controls.Add($btnRemove)

    $btnClose = New-Object System.Windows.Forms.Button
    $btnClose.Location = New-Object System.Drawing.Point(480, 320)
    $btnClose.Size = New-Object System.Drawing.Size(100, 28)
    $btnClose.Text = "Close"
    $btnClose.DialogResult = [System.Windows.Forms.DialogResult]::OK
    $btnClose.Add_Click({ $dlg.Close() })
    $dlg.Controls.Add($btnClose)

    $dlg.CancelButton = $btnClose
    $dlg.ShowDialog() | Out-Null
}

function Show-SettingsDialog {
    $settingsForm = New-Object System.Windows.Forms.Form
    $settingsForm.Text = "User Settings"
    $settingsForm.Size = New-Object System.Drawing.Size(500, 550)
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

    # Reports Base Path
    $lblReportsBasePath = New-Object System.Windows.Forms.Label
    $lblReportsBasePath.Location = New-Object System.Drawing.Point(20, $y)
    $lblReportsBasePath.Size = New-Object System.Drawing.Size(150, 20)
    $lblReportsBasePath.Text = "Reports Base Path:"
    $settingsForm.Controls.Add($lblReportsBasePath)

    $txtReportsBasePath = New-Object System.Windows.Forms.TextBox
    $txtReportsBasePath.Location = New-Object System.Drawing.Point(180, $y)
    $txtReportsBasePath.Size = New-Object System.Drawing.Size(200, 20)
    $txtReportsBasePath.Text = if ($script:UserSettings.ReportsBasePath) { $script:UserSettings.ReportsBasePath } else { "" }
    $settingsForm.Controls.Add($txtReportsBasePath)

    $btnBrowseReportsBase = New-Object System.Windows.Forms.Button
    $btnBrowseReportsBase.Location = New-Object System.Drawing.Point(390, ($y - 2))
    $btnBrowseReportsBase.Size = New-Object System.Drawing.Size(70, 25)
    $btnBrowseReportsBase.Text = "Browse..."
    $btnBrowseReportsBase.Add_Click({
        $folderBrowser = New-Object System.Windows.Forms.FolderBrowserDialog
        $folderBrowser.Description = "Select base folder for client output (e.g. OneDrive root). When mapping companies, you can select subfolders at any depth (e.g. ...\General\Accurate Metal\Network Documentation\Vulnerability Scans)."
        $folderBrowser.ShowNewFolderButton = $true
        $startPath = $txtReportsBasePath.Text.Trim()
        if ($startPath -and (Test-Path $startPath)) { $folderBrowser.SelectedPath = $startPath }
        elseif ($script:UserSettings.LastOutputDirectory -and (Test-Path $script:UserSettings.LastOutputDirectory)) { $folderBrowser.SelectedPath = $script:UserSettings.LastOutputDirectory }
        if ($folderBrowser.ShowDialog() -eq [System.Windows.Forms.DialogResult]::OK) {
            $txtReportsBasePath.Text = $folderBrowser.SelectedPath
        }
    })
    $settingsForm.Controls.Add($btnBrowseReportsBase)

    $btnEditMappings = New-Object System.Windows.Forms.Button
    $btnEditMappings.Location = New-Object System.Drawing.Point(180, ($y + 2))
    $btnEditMappings.Size = New-Object System.Drawing.Size(140, 22)
    $btnEditMappings.Text = "Edit folder mappings..."
    $btnEditMappings.Add_Click({ Show-CompanyFolderMappingDialog })
    $settingsForm.Controls.Add($btnEditMappings)
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
        $folderBrowser.Description = "Select directory for settings and rules configuration files. You can use a cloud folder (OneDrive, Google Drive) or network share."
        $folderBrowser.ShowNewFolderButton = $true
        if (-not [string]::IsNullOrEmpty($txtSettingsDirectory.Text)) {
            $folderBrowser.SelectedPath = $txtSettingsDirectory.Text
        }
        if ($folderBrowser.ShowDialog() -eq [System.Windows.Forms.DialogResult]::OK) {
            $txtSettingsDirectory.Text = $folderBrowser.SelectedPath
        }
    })
    $settingsForm.Controls.Add($btnBrowseSettingsDir)

    # Quick paths for cloud folders (show only if path exists)
    $oneDriveOrg = Get-ChildItem -Path $env:USERPROFILE -Filter "OneDrive - *" -Directory -ErrorAction SilentlyContinue | Select-Object -First 1
    $cloudPaths = @(
        @{ Name = "OneDrive"; Path = (Join-Path $env:USERPROFILE "OneDrive") },
        @{ Name = "OneDrive (org)"; Path = $(if ($oneDriveOrg) { $oneDriveOrg.FullName } else { $null }) },
        @{ Name = "Google Drive"; Path = (Join-Path $env:USERPROFILE "Google Drive") },
        @{ Name = "My Drive"; Path = (Join-Path $env:USERPROFILE "My Drive") },
        @{ Name = "Dropbox"; Path = (Join-Path $env:USERPROFILE "Dropbox") }
    )
    $cloudX = 470
    foreach ($cp in $cloudPaths) {
        if ($cp.Path -and (Test-Path $cp.Path)) {
            $btn = New-Object System.Windows.Forms.Button
            $btn.Location = New-Object System.Drawing.Point($cloudX, ($y - 2))
            $btn.Size = New-Object System.Drawing.Size(75, 25)
            $btn.Text = $cp.Name
            $btn.Font = New-Object System.Drawing.Font($btn.Font.FontFamily, 8)
            $path = $cp.Path
            $btn.Add_Click({ $txtSettingsDirectory.Text = $path })
            $settingsForm.Controls.Add($btn)
            $cloudX += 80
        }
    }
    $y += 35

    # Help text for cloud storage
    $lblCloudHint = New-Object System.Windows.Forms.Label
    $lblCloudHint.Location = New-Object System.Drawing.Point(180, $y)
    $lblCloudHint.Size = New-Object System.Drawing.Size(400, 16)
    $lblCloudHint.Text = "You can use a cloud folder (OneDrive, Google Drive) or network share."
    $lblCloudHint.ForeColor = [System.Drawing.Color]::Gray
    $lblCloudHint.Font = New-Object System.Drawing.Font($lblCloudHint.Font.FontFamily, 8)
    $settingsForm.Controls.Add($lblCloudHint)
    $y += 22

    # Reset to Default button
    $btnResetDir = New-Object System.Windows.Forms.Button
    $btnResetDir.Location = New-Object System.Drawing.Point(180, $y)
    $btnResetDir.Size = New-Object System.Drawing.Size(150, 25)
    $btnResetDir.Text = "Reset to Default"
    $btnResetDir.Add_Click({
        $txtSettingsDirectory.Text = Join-Path $env:LOCALAPPDATA "VScanMagic"
    })
    $settingsForm.Controls.Add($btnResetDir)
    $y += 40

    # AI API Keys (for future expansion)
    $btnAIApiKeys = New-Object System.Windows.Forms.Button
    $btnAIApiKeys.Location = New-Object System.Drawing.Point(20, $y)
    $btnAIApiKeys.Size = New-Object System.Drawing.Size(120, 25)
    $btnAIApiKeys.Text = "AI API Keys..."
    $btnAIApiKeys.Add_Click({ Show-AIApiKeysDialog })
    $settingsForm.Controls.Add($btnAIApiKeys)

    $btnConnectWise = New-Object System.Windows.Forms.Button
    $btnConnectWise.Location = New-Object System.Drawing.Point(150, $y)
    $btnConnectWise.Size = New-Object System.Drawing.Size(160, 25)
    $btnConnectWise.Text = "ConnectWise Automate..."
    $btnConnectWise.Add_Click({ Show-ConnectWiseAutomateSettingsDialog })
    $settingsForm.Controls.Add($btnConnectWise)
    $lblAIApiHint = New-Object System.Windows.Forms.Label
    $lblAIApiHint.Location = New-Object System.Drawing.Point(150, ($y + 4))
    $lblAIApiHint.Size = New-Object System.Drawing.Size(320, 18)
    $lblAIApiHint.Text = "Copilot, ChatGPT, Claude - for future AI features"
    $lblAIApiHint.ForeColor = [System.Drawing.Color]::Gray
    $lblAIApiHint.Font = New-Object System.Drawing.Font($lblAIApiHint.Font.FontFamily, 8.5)
    $settingsForm.Controls.Add($lblAIApiHint)
    $y += 40

    # Backup / Restore Settings (All, Shared, or User scope)
    $lblBackupScope = New-Object System.Windows.Forms.Label
    $lblBackupScope.Location = New-Object System.Drawing.Point(20, $y + 6)
    $lblBackupScope.Size = New-Object System.Drawing.Size(50, 18)
    $lblBackupScope.Text = "Scope:"
    $settingsForm.Controls.Add($lblBackupScope)

    $cmbBackupScope = New-Object System.Windows.Forms.ComboBox
    $cmbBackupScope.Location = New-Object System.Drawing.Point(70, ($y + 2))
    $cmbBackupScope.Size = New-Object System.Drawing.Size(90, 21)
    $cmbBackupScope.DropDownStyle = "DropDownList"
    [void]$cmbBackupScope.Items.AddRange(@("All", "Shared only", "User only"))
    $cmbBackupScope.SelectedIndex = 0
    $settingsForm.Controls.Add($cmbBackupScope)

    $btnBackup = New-Object System.Windows.Forms.Button
    $btnBackup.Location = New-Object System.Drawing.Point(170, $y)
    $btnBackup.Size = New-Object System.Drawing.Size(90, 28)
    $btnBackup.Text = "Backup..."
    $btnBackup.Add_Click({
        $scope = switch ($cmbBackupScope.SelectedIndex) { 0 { "All" } 1 { "Shared" } default { "User" } }
        $defaultName = "VScanMagic_Settings_Backup$(if($scope -ne 'All'){"_$scope"})_$(Get-Date -Format 'yyyy-MM-dd_HHmmss').zip"
        $saveDlg = New-Object System.Windows.Forms.SaveFileDialog
        $saveDlg.Filter = "ZIP Archive (*.zip)|*.zip|All Files (*.*)|*.*"
        $saveDlg.Title = "Save Settings Backup ($scope)"
        $saveDlg.FileName = $defaultName
        $initDir = if ($script:UserSettings.SettingsDirectory -and (Test-Path $script:UserSettings.SettingsDirectory)) { $script:UserSettings.SettingsDirectory } else { [Environment]::GetFolderPath("Desktop") }
        $saveDlg.InitialDirectory = $initDir
        if ($saveDlg.ShowDialog() -eq [System.Windows.Forms.DialogResult]::OK) {
            $outPath = Backup-Settings -OutputPath $saveDlg.FileName -Scope $scope
        } else { $outPath = $null }
        if ($outPath) {
            [System.Windows.Forms.MessageBox]::Show("Settings backed up to:`n$outPath", "Backup Complete",
                [System.Windows.Forms.MessageBoxButtons]::OK, [System.Windows.Forms.MessageBoxIcon]::Information)
        } else {
            [System.Windows.Forms.MessageBox]::Show("No settings files found to backup.", "Backup",
                [System.Windows.Forms.MessageBoxButtons]::OK, [System.Windows.Forms.MessageBoxIcon]::Warning)
        }
    })
    $settingsForm.Controls.Add($btnBackup)

    $btnRestore = New-Object System.Windows.Forms.Button
    $btnRestore.Location = New-Object System.Drawing.Point(270, $y)
    $btnRestore.Size = New-Object System.Drawing.Size(90, 28)
    $btnRestore.Text = "Restore..."
    $btnRestore.Add_Click({
        $scope = switch ($cmbBackupScope.SelectedIndex) { 0 { "All" } 1 { "Shared" } default { "User" } }
        $dlg = New-Object System.Windows.Forms.OpenFileDialog
        $dlg.Filter = "Backup ZIP (*.zip)|*.zip|All Files (*.*)|*.*"
        $dlg.Title = "Select VScanMagic Settings Backup (will restore $scope)"
        if ($dlg.ShowDialog() -eq [System.Windows.Forms.DialogResult]::OK) {
            $result = Restore-Settings -BackupPath $dlg.FileName -Scope $scope
            if ($result) {
                [System.Windows.Forms.MessageBox]::Show("Settings restored successfully. Restart the application for full effect.", "Restore Complete",
                    [System.Windows.Forms.MessageBoxButtons]::OK, [System.Windows.Forms.MessageBoxIcon]::Information)
                $settingsForm.DialogResult = [System.Windows.Forms.DialogResult]::OK
                $settingsForm.Close()
            } else {
                [System.Windows.Forms.MessageBox]::Show("Restore failed or backup contained no valid files for scope '$scope'.", "Restore Failed",
                    [System.Windows.Forms.MessageBoxButtons]::OK, [System.Windows.Forms.MessageBoxIcon]::Error)
            }
        }
    })
    $settingsForm.Controls.Add($btnRestore)
    $y += 40

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
        $script:UserSettings.ReportsBasePath = $txtReportsBasePath.Text.Trim()
        
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
                $oldCompanyFolderMapPath = Join-Path $oldDir "VScanMagic_CompanyFolderMap.json"
                
                if ((Test-Path $oldRulesPath) -and -not (Test-Path $script:RemediationRulesPath)) {
                    Copy-Item -Path $oldRulesPath -Destination $script:RemediationRulesPath -Force
                }
                if ((Test-Path $oldCoveredPath) -and -not (Test-Path $script:CoveredSoftwarePath)) {
                    Copy-Item -Path $oldCoveredPath -Destination $script:CoveredSoftwarePath -Force
                }
                if ((Test-Path $oldRecommendationsPath) -and -not (Test-Path $script:GeneralRecommendationsPath)) {
                    Copy-Item -Path $oldRecommendationsPath -Destination $script:GeneralRecommendationsPath -Force
                }
                if ((Test-Path $oldCompanyFolderMapPath) -and -not (Test-Path $script:CompanyFolderMapPath)) {
                    Copy-Item -Path $oldCompanyFolderMapPath -Destination $script:CompanyFolderMapPath -Force
                }
                
                # Reload rules and covered software from new location
                Load-RemediationRules
                Load-CoveredSoftware
                Load-GeneralRecommendations
                Load-CompanyFolderMap
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

function Show-VScanMagicOverviewHelpDialog {
    # General overview - how the software works (broader than API-specific help)
    $helpForm = New-Object System.Windows.Forms.Form
    $helpForm.Text = "VScanMagic - Overview"
    $helpForm.Size = New-Object System.Drawing.Size(610, 640)
    $helpForm.StartPosition = "CenterParent"
    $helpForm.FormBorderStyle = "FixedDialog"
    $helpForm.MaximizeBox = $false

    $txt = New-Object System.Windows.Forms.RichTextBox
    $txt.Location = New-Object System.Drawing.Point(12, 12)
    $txt.Size = New-Object System.Drawing.Size(575, 545)
    $txt.ReadOnly = $true
    $txt.Font = New-Object System.Drawing.Font("Segoe UI", 9.5)
    $txt.BackColor = [System.Drawing.Color]::White
    $txt.Text = @"
WHAT IS VScanMagic?
-------------------

VScanMagic turns vulnerability scan data into client-ready reports. You give it
scan results (from ConnectSecure or an Excel file), and it produces Word
documents, Excel spreadsheets, email templates, and more - with risk scores
and remediation guidance already calculated.

HOW TO USE IT (SIMPLE)
----------------------

In short: pick your data, choose your outputs, then click the main
"Download and Generate Reports" button (green, bottom-left).

- If you use ConnectSecure: Enter your API credentials once, pick a company
  and scan date, check which reports you want, and click Download & Generate
  Reports. VScanMagic downloads the data, processes it, and saves the files
  you selected.

- If you have a file: Click Browse, select your All Vulnerabilities Excel
  file, enter the client name and scan date, choose outputs, and click
  Download and Generate Reports.

EXAMPLE WORKFLOW
----------------

Scenario: Generate a quarterly vulnerability report for "Acme Corp" from
ConnectSecure.

1. API Settings: Enter Base URL, Tenant, Client ID, Client Secret. Test and Save.
2. Refresh List: Load your companies, select "Acme Corp".
3. Scan Date: Pick the date of the scan (e.g., last week).
4. Reports: Ensure "All Vulnerabilities Report" is checked (required).
   Optionally check Pending EPSS, Executive Summary, etc.
5. Output Directory: Choose where to save the files (e.g., Desktop).
6. Output Options: Check Word Report, Excel, Email Template - whatever you need.
7. Download and Generate Reports: Click the main green button. VScanMagic will:
   - Download reports from ConnectSecure
   - Show a General Recommendations dialog (edit or keep defaults)
   - Show a Hostname Review dialog (filter if needed)
   - Generate your Word doc, Excel file, and other outputs

Done. Your reports are in the output folder.

KEY CONCEPTS
------------

- Risk Score: Combines severity (Critical/High/Medium/Low) and EPSS.
  Higher score = fix it sooner.

- Top 10: The highest-risk items. You can change the count in Report Filters.

- All Vulnerabilities: The report VScanMagic uses for risk scoring. Required
  when downloading from ConnectSecure.

NEED MORE DETAIL?
-----------------

- API credentials and ConnectSecure setup: Click "API Help" (next to API Settings).
- Full workflow options: Click "API Help", then see the Workflow tab.
"@
    $helpForm.Controls.Add($txt)

    $btnClose = New-Object System.Windows.Forms.Button
    $btnClose.Location = New-Object System.Drawing.Point(488, 565)
    $btnClose.Size = New-Object System.Drawing.Size(100, 28)
    $btnClose.Text = "Close"
    $btnClose.DialogResult = [System.Windows.Forms.DialogResult]::OK
    $helpForm.AcceptButton = $btnClose
    $helpForm.Controls.Add($btnClose)

    $helpForm.ShowDialog() | Out-Null
}

function Show-ConnectSecureApiHelpDialog {
    # API-specific: credentials, workflow (shown when user needs ConnectSecure help)
    $helpForm = New-Object System.Windows.Forms.Form
    $helpForm.Text = "ConnectSecure API Help"
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

STEP 1: Find the Base URL
   - Log into your ConnectSecure portal in a web browser.
   - Click the user profile icon, then click "API Documentation".
   - The page URL will look like: https://pod0.myconnectsecure.com/apidocs/
   - The Base URL is the part before /apidocs/ (e.g., https://pod0.myconnectsecure.com)

STEP 2: Open the API Key page
   - Go to: Global > Settings > Users
   - For an existing user: Click the three-dot menu (...) next to your user -> "API Key"
   - For a new API user: Click "Add", create a user, then use Action > API Key

STEP 3: Copy each value
   - Base URL
     Use the URL from Step 1 (from Profile > API Documentation, before /apidocs/).
     Example: https://pod0.myconnectsecure.com

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
   - All Vulnerabilities - required for full VScanMagic processing
   - Pending EPSS, Suppressed, External, Executive Summary - optional

5. Output Directory: Choose where files will be saved.

6. Report Filters: (Button) Set severity (Critical/High/Medium/Low), Top N count, EPSS threshold.

7. Output Options: (Button) Choose outputs: Excel, Word, Email Template, Ticket Instructions, Time Estimate.

8. Click "Download and Generate Reports" (main button):
   - Downloads reports from ConnectSecure for each selected company
   - Runs through General Recommendations and Hostname Review dialogs
   - Generates your selected outputs (Word report, Excel, etc.)

OPTION B: PROCESS FROM FILE
---------------------------

1. In section 2, click "Browse..." and select a previously downloaded All Vulnerabilities report (XLSX).

2. Enter Client name and Scan Date (sometimes auto-filled from filename).

3. Set Output Options and Report Filters if needed.

4. Click "Download and Generate Reports" to process the file and create outputs.

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

function Show-DownloadCustomReportDialog {
    <#
    .SYNOPSIS
    Shows a dialog to select standard reports and export formats for download only (no processing).
    Dynamically loads report list from ConnectSecure API when connected; falls back to the 5 known types when not.
    .PARAMETER CompanyId
    First selected company ID - use for API (isGlobal=false). 0 = global context.
    .PARAMETER GlobalReports
    When true, fetches global reports (isGlobal=true) - Report Builder templates, no company scope.
    .OUTPUTS
    Array of @{ ReportId (optional); Type (optional); Name; Ext } for selected report+format combinations, or $null if cancelled.
    #>
    param([int]$CompanyId = 0, [switch]$GlobalReports = $false)
    # Fallback when API not connected or returns no reports
    $fallbackDefs = @(
        @{ Type = 'all-vulnerabilities'; Name = 'All Vulnerabilities Report'; Formats = @('xlsx','docx','pdf') }
        @{ Type = 'suppressed-vulnerabilities'; Name = 'Suppressed Vulnerabilities'; Formats = @('xlsx','docx','pdf') }
        @{ Type = 'external-vulnerabilities'; Name = 'External Scan'; Formats = @('xlsx','docx','pdf') }
        @{ Type = 'executive-summary'; Name = 'Executive Summary Report'; Formats = @('xlsx','docx','pdf') }
        @{ Type = 'pending-epss'; Name = 'Pending Remediation EPSS Score Reports'; Formats = @('xlsx','docx','pdf') }
    )

    $reportDefs = @()
    try {
        # Ensure API connected so we can fetch full report list (ConnectSecure-API refreshes token if needed)
        $creds = Load-ConnectSecureCredentials
        if ($creds -and -not [string]::IsNullOrWhiteSpace($creds.ClientId)) {
            $null = Connect-ConnectSecureAPI -BaseUrl $creds.BaseUrl -TenantName $creds.TenantName -ClientId $creds.ClientId -ClientSecret $creds.ClientSecret
        }
        $apiReports = if ($GlobalReports) {
            Get-ConnectSecureStandardReports -IsGlobal $true -UseGlobalOnly
        } else {
            Get-ConnectSecureStandardReports -CompanyId $CompanyId
        }
        if ($apiReports -and $apiReports.Count -gt 0) {
            # Group by _categoryDisplay, collect report entries (id, reportType) per category
            $byCategory = @{}
            foreach ($r in $apiReports) {
                $cat = if ($r._categoryDisplay) { $r._categoryDisplay } else { $r._category }
                if (-not $cat) { $cat = 'Unknown' }
                $fmt = if ($r.reportType) { $r.reportType.ToString().ToLower() } else { 'xlsx' }
                if ($fmt -notin @('xlsx','docx','pdf')) { continue }
                if (-not $byCategory[$cat]) { $byCategory[$cat] = @{} }
                if (-not $byCategory[$cat][$fmt]) {
                    $byCategory[$cat][$fmt] = @{ ReportId = $r.id; Name = $cat; Ext = $fmt }
                }
            }
            foreach ($key in ($byCategory.Keys | Sort-Object)) {
                $fmts = $byCategory[$key]
                $reportDefs += @{ ReportIdMap = $fmts; Name = $key }
            }
        }
    } catch { }
    if ($reportDefs.Count -eq 0) {
        foreach ($def in $fallbackDefs) {
            $reportDefs += @{ Type = $def.Type; Name = $def.Name; Formats = $def.Formats }
        }
    }

    $form = New-Object System.Windows.Forms.Form
    $form.Text = if ($GlobalReports) { 'Download Global Reports' } else { 'Download Standard Report (Company)' }
    $form.Size = New-Object System.Drawing.Size(480, 580)
    $form.StartPosition = 'CenterParent'
    $form.FormBorderStyle = 'FixedDialog'
    $form.MaximizeBox = $false
    $form.MinimizeBox = $false

    $lbl = New-Object System.Windows.Forms.Label
    $lbl.Location = New-Object System.Drawing.Point(12, 10)
    $lbl.Size = New-Object System.Drawing.Size(440, 24)
    $lbl.Text = if ($GlobalReports) { 'Select global report template(s) and format(s) to download:' } else { 'Select report(s) and format(s) to download (no processing):' }
    $form.Controls.Add($lbl)

    $flow = New-Object System.Windows.Forms.FlowLayoutPanel
    $flow.Location = New-Object System.Drawing.Point(12, 40)
    $flow.Size = New-Object System.Drawing.Size(440, 430)
    $flow.FlowDirection = [System.Windows.Forms.FlowDirection]::TopDown
    $flow.AutoScroll = $true
    $flow.WrapContents = $false
    $flow.BorderStyle = [System.Windows.Forms.BorderStyle]::FixedSingle

    $checkboxes = @{}
    $formatsOrder = @('xlsx','docx','pdf')
    foreach ($def in $reportDefs) {
        $grp = New-Object System.Windows.Forms.GroupBox
        $grp.Text = $def.Name
        $grp.Size = New-Object System.Drawing.Size(420, 52)
        $grp.Margin = New-Object System.Windows.Forms.Padding(0, 4, 0, 4)
        $x = 8; $y = 20
        if ($def.ReportIdMap) {
            foreach ($ext in $formatsOrder) {
                $entry = $def.ReportIdMap[$ext]
                if (-not $entry) { continue }
                $displayFmt = $ext.ToUpper()
                $chk = New-Object System.Windows.Forms.CheckBox
                $chk.Text = $displayFmt
                $chk.Location = New-Object System.Drawing.Point($x, $y)
                $chk.Size = New-Object System.Drawing.Size(60, 24)
                $chk.Tag = @{ ReportId = $entry.ReportId; Name = $entry.Name; Ext = $entry.Ext }
                $grp.Controls.Add($chk)
                $checkboxes["$($def.Name)-$ext"] = $chk
                $x += 70
            }
        } else {
            foreach ($fmt in $def.Formats) {
                $ext = $fmt
                $displayFmt = $fmt.ToUpper()
                $chk = New-Object System.Windows.Forms.CheckBox
                $chk.Text = $displayFmt
                $chk.Location = New-Object System.Drawing.Point($x, $y)
                $chk.Size = New-Object System.Drawing.Size(60, 24)
                $chk.Tag = @{ Type = $def.Type; Name = $def.Name; Ext = $ext }
                $grp.Controls.Add($chk)
                $checkboxes["$($def.Type)-$ext"] = $chk
                $x += 70
            }
        }
        $flow.Controls.Add($grp)
    }

    $form.Controls.Add($flow)

    $btnSelectAll = New-Object System.Windows.Forms.Button
    $btnSelectAll.Location = New-Object System.Drawing.Point(12, 478)
    $btnSelectAll.Size = New-Object System.Drawing.Size(90, 28)
    $btnSelectAll.Text = 'Select All'
    $btnSelectAll.Add_Click({
        foreach ($chk in $checkboxes.Values) { $chk.Checked = $true }
    })
    $form.Controls.Add($btnSelectAll)

    $btnClearAll = New-Object System.Windows.Forms.Button
    $btnClearAll.Location = New-Object System.Drawing.Point(108, 478)
    $btnClearAll.Size = New-Object System.Drawing.Size(90, 28)
    $btnClearAll.Text = 'Clear All'
    $btnClearAll.Add_Click({
        foreach ($chk in $checkboxes.Values) { $chk.Checked = $false }
    })
    $form.Controls.Add($btnClearAll)

    $btnOk = New-Object System.Windows.Forms.Button
    $btnOk.Location = New-Object System.Drawing.Point(290, 478)
    $btnOk.Size = New-Object System.Drawing.Size(90, 28)
    $btnOk.Text = 'Download'
    $btnOk.DialogResult = [System.Windows.Forms.DialogResult]::OK
    $form.AcceptButton = $btnOk

    $btnCancel = New-Object System.Windows.Forms.Button
    $btnCancel.Location = New-Object System.Drawing.Point(392, 478)
    $btnCancel.Size = New-Object System.Drawing.Size(90, 28)
    $btnCancel.Text = 'Cancel'
    $btnCancel.DialogResult = [System.Windows.Forms.DialogResult]::Cancel
    $form.CancelButton = $btnCancel

    $form.Controls.Add($btnOk)
    $form.Controls.Add($btnCancel)

    if ($form.ShowDialog() -ne [System.Windows.Forms.DialogResult]::OK) { return $null }

    $selected = @()
    foreach ($key in $checkboxes.Keys) {
        $chk = $checkboxes[$key]
        if ($chk.Checked) {
            $tag = $chk.Tag
            $item = @{ Name = $tag.Name; Ext = $tag.Ext }
            if ($tag.ReportId) { $item.ReportId = $tag.ReportId }
            else { $item.Type = $tag.Type }
            $selected += $item
        }
    }
    return $selected
}
