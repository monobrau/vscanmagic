# VScanMagic-Form.ps1 - Main form and event handlers
# Dot-sourced by VScanMagic-GUI.ps1
function Show-VScanMagicGUI {
    # Initialize script-level variables for output file paths
    $script:WordReportPath = $null
    $script:ExcelReportPath = $null
    $script:EmailTemplatePath = $null
    $script:TicketInstructionsPath = $null
    $script:TimeEstimatePath = $null
    $script:TicketNotesPath = $null
    $script:IsRMITPlus = $false
    $script:CurrentTop10Data = $null
    $script:CurrentTimeEstimates = $null

    # Load user settings from disk (this also initializes and updates paths)
    Load-UserSettings
    
    # Ensure paths are updated after loading settings
    Update-SettingsPaths
    
    # Load company folder mapping (for structured output paths)
    Load-CompanyFolderMap
    
    # Load remediation rules from disk
    Load-RemediationRules
    
    # Load covered software list from disk
    Load-CoveredSoftware
    
    # Load general recommendations from disk
    Load-GeneralRecommendations

    # Load templates (email, ticket notes)
    Load-Templates

    # Load ConnectSecure API (used by Download section)
    $scriptDir = $script:ScriptDirectory
    if ([string]::IsNullOrWhiteSpace($scriptDir)) { $scriptDir = if ($PSScriptRoot) { $PSScriptRoot } else { (Get-Location).Path } }
    $connectSecureScriptPath = Join-Path $scriptDir "ConnectSecure-API.ps1"
    if (Test-Path $connectSecureScriptPath) {
        try { . $connectSecureScriptPath } catch { Write-Log "Could not load ConnectSecure API: $($_.Exception.Message)" -Level Warning }
    }

    # Create main form
    $form = New-Object System.Windows.Forms.Form
    $form.Text = "$($script:Config.AppName) - Vulnerability Report Generator"
    $form.Size = New-Object System.Drawing.Size(750, 945)
    $form.StartPosition = "CenterScreen"
    $form.FormBorderStyle = "FixedDialog"
    $form.MaximizeBox = $false
    $form.AutoScroll = $true
    $form.AutoScrollMinSize = New-Object System.Drawing.Size(720, 878)

    # --- Report Filters & Output Options (buttons, like API Settings) ---
    $btnReportFilters = New-Object System.Windows.Forms.Button
    $btnReportFilters.Location = New-Object System.Drawing.Point(20, 20)
    $btnReportFilters.Size = New-Object System.Drawing.Size(110, 24)
    $btnReportFilters.Text = "Report Filters"
    $btnReportFilters.BackColor = [System.Drawing.Color]::FromArgb(0, 120, 215)
    $btnReportFilters.ForeColor = [System.Drawing.Color]::White
    $btnReportFilters.FlatStyle = [System.Windows.Forms.FlatStyle]::Flat
    $btnReportFilters.FlatAppearance.BorderSize = 0
    $btnReportFilters.Add_Click({ Show-FiltersDialog })
    $form.Controls.Add($btnReportFilters)

    $lblFiltersHint = New-Object System.Windows.Forms.Label
    $lblFiltersHint.Location = New-Object System.Drawing.Point(138, 24)
    $lblFiltersHint.Size = New-Object System.Drawing.Size(180, 18)
    $lblFiltersHint.Text = "Min EPSS, Severities, Top N"
    $lblFiltersHint.ForeColor = [System.Drawing.Color]::Gray
    $lblFiltersHint.Font = New-Object System.Drawing.Font($lblFiltersHint.Font.FontFamily, 8.5)
    $form.Controls.Add($lblFiltersHint)

    $btnOutputOptions = New-Object System.Windows.Forms.Button
    $btnOutputOptions.Location = New-Object System.Drawing.Point(20, 52)
    $btnOutputOptions.Size = New-Object System.Drawing.Size(110, 24)
    $btnOutputOptions.Text = "Output Options"
    $btnOutputOptions.BackColor = [System.Drawing.Color]::FromArgb(94, 53, 177)
    $btnOutputOptions.ForeColor = [System.Drawing.Color]::White
    $btnOutputOptions.FlatStyle = [System.Windows.Forms.FlatStyle]::Flat
    $btnOutputOptions.FlatAppearance.BorderSize = 0
    $btnOutputOptions.Add_Click({ Show-OutputOptionsDialog })
    $form.Controls.Add($btnOutputOptions)

    $btnTemplates = New-Object System.Windows.Forms.Button
    $btnTemplates.Location = New-Object System.Drawing.Point(330, 52)
    $btnTemplates.Size = New-Object System.Drawing.Size(90, 24)
    $btnTemplates.Text = "Templates"
    $btnTemplates.BackColor = [System.Drawing.Color]::FromArgb(94, 53, 177)
    $btnTemplates.ForeColor = [System.Drawing.Color]::White
    $btnTemplates.FlatStyle = [System.Windows.Forms.FlatStyle]::Flat
    $btnTemplates.FlatAppearance.BorderSize = 0
    $btnTemplates.Add_Click({ Show-TemplatesDialog })
    $form.Controls.Add($btnTemplates)
    $ttTemplates = New-Object System.Windows.Forms.ToolTip
    $ttTemplates.SetToolTip($btnTemplates, "Customize email template and ticket notes. Changes are saved to your settings folder.")

    $lblOutputHint = New-Object System.Windows.Forms.Label
    $lblOutputHint.Location = New-Object System.Drawing.Point(430, 56)
    $lblOutputHint.Size = New-Object System.Drawing.Size(280, 18)
    $lblOutputHint.Text = "Excel, Word, Email, Ticket, Time Estimate"
    $lblOutputHint.ForeColor = [System.Drawing.Color]::Gray
    $lblOutputHint.Font = New-Object System.Drawing.Font($lblOutputHint.Font.FontFamily, 8.5)
    $form.Controls.Add($lblOutputHint)

    $chkBulkProcessing = New-Object System.Windows.Forms.CheckBox
    $chkBulkProcessing.Location = New-Object System.Drawing.Point(20, 84)
    $chkBulkProcessing.Size = New-Object System.Drawing.Size(420, 20)
    $chkBulkProcessing.Text = "Skip follow-up dialogs (bulk processing)"
    $chkBulkProcessing.Checked = $false
    $chkBulkProcessing.ForeColor = [System.Drawing.Color]::Gray
    $chkBulkProcessing.Font = New-Object System.Drawing.Font($chkBulkProcessing.Font.FontFamily, 8.5)
    $toolTipBulk = New-Object System.Windows.Forms.ToolTip
    $toolTipBulk.SetToolTip($chkBulkProcessing, "When checked and 2+ companies are selected: skips General Recommendations, Hostname Review, Time Estimate dialogs, and completion/error popups. Uses defaults as if OK was clicked.")
    $form.Controls.Add($chkBulkProcessing)

    # --- 1. Download from ConnectSecure (inline) ---
    $groupBoxDownload = New-Object System.Windows.Forms.GroupBox
    $groupBoxDownload.Location = New-Object System.Drawing.Point(20, 118)
    $groupBoxDownload.Size = New-Object System.Drawing.Size(680, 315)
    $groupBoxDownload.Text = "1. Download from ConnectSecure"
    $form.Controls.Add($groupBoxDownload)

    $dlgY = 22
    $btnApiSettings = New-Object System.Windows.Forms.Button
    $btnApiSettings.Location = New-Object System.Drawing.Point(20, $dlgY)
    $btnApiSettings.Size = New-Object System.Drawing.Size(100, 24)
    $btnApiSettings.Text = "API Settings"
    $btnApiSettings.BackColor = [System.Drawing.Color]::FromArgb(94, 53, 177)
    $btnApiSettings.ForeColor = [System.Drawing.Color]::White
    $btnApiSettings.FlatStyle = [System.Windows.Forms.FlatStyle]::Flat
    $btnApiSettings.FlatAppearance.BorderSize = 0
    $btnApiSettings.Add_Click({ Show-ConnectSecureSettingsDialog })
    $groupBoxDownload.Controls.Add($btnApiSettings)
    $btnApiHelp = New-Object System.Windows.Forms.Button
    $btnApiHelp.Location = New-Object System.Drawing.Point(128, $dlgY)
    $btnApiHelp.Size = New-Object System.Drawing.Size(80, 24)
    $btnApiHelp.Text = "API Help"
    $btnApiHelp.BackColor = [System.Drawing.Color]::FromArgb(255, 167, 38)
    $btnApiHelp.ForeColor = [System.Drawing.Color]::White
    $btnApiHelp.FlatStyle = [System.Windows.Forms.FlatStyle]::Flat
    $btnApiHelp.FlatAppearance.BorderSize = 0
    $btnApiHelp.Add_Click({ Show-ConnectSecureApiHelpDialog })
    $groupBoxDownload.Controls.Add($btnApiHelp)
    $lblApiSettingsHint = New-Object System.Windows.Forms.Label
    $lblApiSettingsHint.Location = New-Object System.Drawing.Point(215, ($dlgY + 4))
    $lblApiSettingsHint.Size = New-Object System.Drawing.Size(380, 18)
    $lblApiSettingsHint.Text = "Base URL, Tenant, Client ID, Client Secret"
    $lblApiSettingsHint.ForeColor = [System.Drawing.Color]::Gray
    $lblApiSettingsHint.Font = New-Object System.Drawing.Font($lblApiSettingsHint.Font.FontFamily, 8.5)
    $groupBoxDownload.Controls.Add($lblApiSettingsHint)
    $dlgY += 32

    $lblCompany = New-Object System.Windows.Forms.Label
    $lblCompany.Location = New-Object System.Drawing.Point(20, ($dlgY + 2))
    $lblCompany.Size = New-Object System.Drawing.Size(65, 18)
    $lblCompany.Text = "Company:"
    $groupBoxDownload.Controls.Add($lblCompany)
    $checkedListCompany = New-Object System.Windows.Forms.CheckedListBox
    $checkedListCompany.Location = New-Object System.Drawing.Point(88, $dlgY)
    $checkedListCompany.Size = New-Object System.Drawing.Size(250, 85)
    $checkedListCompany.CheckOnClick = $true
    $checkedListCompany.DisplayMember = "DisplayName"
    $checkedListCompany.IntegralHeight = $false
    $groupBoxDownload.Controls.Add($checkedListCompany)
    $btnRefreshCompanies = New-Object System.Windows.Forms.Button
    $btnRefreshCompanies.Location = New-Object System.Drawing.Point(348, $dlgY)
    $btnRefreshCompanies.Size = New-Object System.Drawing.Size(80, 22)
    $btnRefreshCompanies.Text = "Refresh List"
    $btnRefreshCompanies.BackColor = [System.Drawing.Color]::FromArgb(0, 120, 215)
    $btnRefreshCompanies.ForeColor = [System.Drawing.Color]::White
    $btnRefreshCompanies.FlatStyle = [System.Windows.Forms.FlatStyle]::Flat
    $btnRefreshCompanies.FlatAppearance.BorderSize = 0
    $groupBoxDownload.Controls.Add($btnRefreshCompanies)
    $btnSelectAllCompanies = New-Object System.Windows.Forms.Button
    $btnSelectAllCompanies.Location = New-Object System.Drawing.Point(348, ($dlgY + 26))
    $btnSelectAllCompanies.Size = New-Object System.Drawing.Size(60, 22)
    $btnSelectAllCompanies.Text = "Select All"
    $btnSelectAllCompanies.Add_Click({
        for ($i = 0; $i -lt $checkedListCompany.Items.Count; $i++) {
            $item = $checkedListCompany.Items[$i]
            if ($item.Id -eq 0 -and $checkedListCompany.CheckedItems.Count -gt 0) {
                $hasOthers = $checkedListCompany.CheckedItems | Where-Object { $_.Id -ne 0 } | Select-Object -First 1
                if ($hasOthers) { continue }
            }
            $checkedListCompany.SetItemChecked($i, $true)
        }
    })
    $groupBoxDownload.Controls.Add($btnSelectAllCompanies)
    $btnClearAllCompanies = New-Object System.Windows.Forms.Button
    $btnClearAllCompanies.Location = New-Object System.Drawing.Point(412, ($dlgY + 26))
    $btnClearAllCompanies.Size = New-Object System.Drawing.Size(60, 22)
    $btnClearAllCompanies.Text = "Clear All"
    $btnClearAllCompanies.Add_Click({
        for ($i = 0; $i -lt $checkedListCompany.Items.Count; $i++) {
            $checkedListCompany.SetItemChecked($i, $false)
        }
    })
    $groupBoxDownload.Controls.Add($btnClearAllCompanies)

    $btnCompanyReview = New-Object System.Windows.Forms.Button
    $btnCompanyReview.Location = New-Object System.Drawing.Point(348, ($dlgY + 52))
    $btnCompanyReview.Size = New-Object System.Drawing.Size(124, 22)
    $btnCompanyReview.Text = "Company Review"
    $btnCompanyReview.BackColor = [System.Drawing.Color]::FromArgb(103, 58, 183)
    $btnCompanyReview.ForeColor = [System.Drawing.Color]::White
    $btnCompanyReview.FlatStyle = [System.Windows.Forms.FlatStyle]::Flat
    $btnCompanyReview.FlatAppearance.BorderSize = 0
    $btnCompanyReview.Add_Click({
        $checked = @($checkedListCompany.CheckedItems)
        if ($checked.Count -ne 1 -or $checked[0].Id -eq 0) {
            [System.Windows.Forms.MessageBox]::Show("Select exactly one company (not All Companies) to run Company Review.", "Company Review", [System.Windows.Forms.MessageBoxButtons]::OK, [System.Windows.Forms.MessageBoxIcon]::Information)
            return
        }
        $company = $checked[0]
        $clientName = ($company.DisplayName -replace '\s*\(ID:\s*\d+\)\s*$', '').Trim()
        if ([string]::IsNullOrWhiteSpace($clientName)) { $clientName = $company.DisplayName }
        Show-CompanyReviewDialog -CompanyId $company.Id -CompanyName $clientName
    })
    $groupBoxDownload.Controls.Add($btnCompanyReview)

    $checkedListCompany.Add_ItemCheck({
        param($sender, $e)
        if ($e.NewValue -eq [System.Windows.Forms.CheckState]::Checked) {
            $item = $sender.Items[$e.Index]
            if ($item.Id -eq 0) {
                for ($i = 0; $i -lt $sender.Items.Count; $i++) {
                    if ($i -ne $e.Index) { $sender.SetItemChecked($i, $false) }
                }
            } else {
                $allIdx = 0
                for ($i = 0; $i -lt $sender.Items.Count; $i++) {
                    if ($sender.Items[$i].Id -eq 0) { $allIdx = $i; break }
                }
                $sender.SetItemChecked($allIdx, $false)
            }
        }
    })
    $dlgY += 95
    $lblScanDateDlg = New-Object System.Windows.Forms.Label
    $lblScanDateDlg.Location = New-Object System.Drawing.Point(20, ($dlgY + 2))
    $lblScanDateDlg.Size = New-Object System.Drawing.Size(65, 18)
    $lblScanDateDlg.Text = "Scan Date:"
    $groupBoxDownload.Controls.Add($lblScanDateDlg)
    $datePickerDownloadScanDate = New-Object System.Windows.Forms.DateTimePicker
    $datePickerDownloadScanDate.Location = New-Object System.Drawing.Point(88, $dlgY)
    $datePickerDownloadScanDate.Size = New-Object System.Drawing.Size(100, 20)
    $datePickerDownloadScanDate.Format = [System.Windows.Forms.DateTimePickerFormat]::Short
    $datePickerDownloadScanDate.Value = Get-Date
    $groupBoxDownload.Controls.Add($datePickerDownloadScanDate)
    $checkBoxRMITPlus = New-Object System.Windows.Forms.CheckBox
    $checkBoxRMITPlus.Location = New-Object System.Drawing.Point(210, ($dlgY + 2))
    $checkBoxRMITPlus.Size = New-Object System.Drawing.Size(75, 18)
    $checkBoxRMITPlus.Text = "RMIT+?"
    $checkBoxRMITPlus.Checked = $false
    $groupBoxDownload.Controls.Add($checkBoxRMITPlus)
    $dlgY += 28

    $chkAllVulnerabilities = New-Object System.Windows.Forms.CheckBox
    $chkAllVulnerabilities.Location = New-Object System.Drawing.Point(20, $dlgY)
    $chkAllVulnerabilities.Size = New-Object System.Drawing.Size(310, 18)
    $chkAllVulnerabilities.Text = "All Vulnerabilities Report (XLSX)"
    $chkAllVulnerabilities.Checked = $true
    $groupBoxDownload.Controls.Add($chkAllVulnerabilities)
    $chkExecutiveSummary = New-Object System.Windows.Forms.CheckBox
    $chkExecutiveSummary.Location = New-Object System.Drawing.Point(350, $dlgY)
    $chkExecutiveSummary.Size = New-Object System.Drawing.Size(310, 18)
    $chkExecutiveSummary.Text = "Executive Summary Report (DOCX)"
    $chkExecutiveSummary.Checked = $true
    $groupBoxDownload.Controls.Add($chkExecutiveSummary)
    $dlgY += 20
    $chkSuppressedVulnerabilities = New-Object System.Windows.Forms.CheckBox
    $chkSuppressedVulnerabilities.Location = New-Object System.Drawing.Point(20, $dlgY)
    $chkSuppressedVulnerabilities.Size = New-Object System.Drawing.Size(310, 18)
    $chkSuppressedVulnerabilities.Text = "Suppressed Vulnerabilities (XLSX)"
    $chkSuppressedVulnerabilities.Checked = $true
    $groupBoxDownload.Controls.Add($chkSuppressedVulnerabilities)
    $chkPendingEPSS = New-Object System.Windows.Forms.CheckBox
    $chkPendingEPSS.Location = New-Object System.Drawing.Point(350, $dlgY)
    $chkPendingEPSS.Size = New-Object System.Drawing.Size(310, 18)
    $chkPendingEPSS.Text = "Pending Remediation EPSS Score Reports (XLSX)"
    $chkPendingEPSS.Checked = $true
    $groupBoxDownload.Controls.Add($chkPendingEPSS)
    $dlgY += 20
    $chkExternalVulnerabilities = New-Object System.Windows.Forms.CheckBox
    $chkExternalVulnerabilities.Location = New-Object System.Drawing.Point(20, $dlgY)
    $chkExternalVulnerabilities.Size = New-Object System.Drawing.Size(310, 18)
    $chkExternalVulnerabilities.Text = "External Scan (XLSX)"
    $chkExternalVulnerabilities.Checked = $true
    $groupBoxDownload.Controls.Add($chkExternalVulnerabilities)
    $dlgY += 26

    $lblDownloadProgress = New-Object System.Windows.Forms.Label
    $lblDownloadProgress.Location = New-Object System.Drawing.Point(20, $dlgY)
    $lblDownloadProgress.Size = New-Object System.Drawing.Size(620, 18)
    $lblDownloadProgress.Text = ""
    $lblDownloadProgress.ForeColor = [System.Drawing.Color]::Blue
    $groupBoxDownload.Controls.Add($lblDownloadProgress)
    $dlgY += 28

    $btnDownloadStandardOnly = New-Object System.Windows.Forms.Button
    $btnDownloadStandardOnly.Location = New-Object System.Drawing.Point(20, $dlgY)
    $btnDownloadStandardOnly.Size = New-Object System.Drawing.Size(240, 28)
    $btnDownloadStandardOnly.Text = "Download Standard Reports Only"
    $btnDownloadStandardOnly.BackColor = [System.Drawing.Color]::FromArgb(46, 125, 50)
    $btnDownloadStandardOnly.ForeColor = [System.Drawing.Color]::White
    $btnDownloadStandardOnly.FlatStyle = [System.Windows.Forms.FlatStyle]::Flat
    $btnDownloadStandardOnly.FlatAppearance.BorderSize = 0
    $btnDownloadStandardOnly.Add_Click({
        $creds = Load-ConnectSecureCredentials
        if (-not $creds -or [string]::IsNullOrWhiteSpace($creds.BaseUrl) -or [string]::IsNullOrWhiteSpace($creds.ClientId) -or [string]::IsNullOrWhiteSpace($creds.ClientSecret)) {
            [System.Windows.Forms.MessageBox]::Show("Please configure API credentials first. Click 'Settings' then 'API Settings'.", "Credentials Required", [System.Windows.Forms.MessageBoxButtons]::OK, [System.Windows.Forms.MessageBoxIcon]::Warning)
            return
        }
        $checkedCompanies = @($checkedListCompany.CheckedItems)
        if ($checkedCompanies.Count -eq 0) {
            [System.Windows.Forms.MessageBox]::Show("Please select at least one company.", "Validation", [System.Windows.Forms.MessageBoxButtons]::OK, [System.Windows.Forms.MessageBoxIcon]::Warning)
            return
        }
        if ([string]::IsNullOrWhiteSpace($script:UserSettings.ReportsBasePath)) {
            $downloadFolder = $textBoxOutputDir.Text
            if (-not (Test-Path $downloadFolder)) {
                [System.Windows.Forms.MessageBox]::Show("Output directory does not exist. Please select a valid directory.", "Validation", [System.Windows.Forms.MessageBoxButtons]::OK, [System.Windows.Forms.MessageBoxIcon]::Warning)
                return
            }
        }
        $btnDownloadStandardOnly.Enabled = $false
        $lblDownloadProgress.Text = "Connecting..."
        $form.Refresh()

        $onProgress = { param($m) $lblDownloadProgress.Text = $m; $form.Refresh(); [System.Windows.Forms.Application]::DoEvents() }
        $connected = Connect-ConnectSecureAPI -BaseUrl $creds.BaseUrl -TenantName $creds.TenantName -ClientId $creds.ClientId -ClientSecret $creds.ClientSecret
        if (-not $connected) {
            [System.Windows.Forms.MessageBox]::Show("Authentication failed. Check API Settings.", "Authentication Error", [System.Windows.Forms.MessageBoxButtons]::OK, [System.Windows.Forms.MessageBoxIcon]::Error)
            $btnDownloadStandardOnly.Enabled = $true
            return
        }

        $successCount = 0
        $failCount = 0
        $maxRetries = 3
        $standardReports = @(
            @{ Type = "all-vulnerabilities"; Name = "All Vulnerabilities Report"; Ext = "xlsx" }
            @{ Type = "suppressed-vulnerabilities"; Name = "Suppressed Vulnerabilities"; Ext = "xlsx" }
            @{ Type = "external-vulnerabilities"; Name = "External Scan"; Ext = "xlsx" }
            @{ Type = "executive-summary"; Name = "Executive Summary Report"; Ext = "docx" }
            @{ Type = "pending-epss"; Name = "Pending Remediation EPSS Score Reports"; Ext = "xlsx" }
        )
        foreach ($company in $checkedCompanies) {
            $clientName = ($company.DisplayName -replace '\s*\(ID:\s*\d+\)\s*$', '').Trim()
            if ([string]::IsNullOrWhiteSpace($clientName)) { $clientName = $company.DisplayName }

            $downloadFolder = Resolve-ClientOutputPath -CompanyId $company.Id -CompanyName $clientName -ScanDate ($datePickerDownloadScanDate.Value.ToString("MM/dd/yyyy")) -FallbackPath $textBoxOutputDir.Text -ForceManual:$chkReselectFolder.Checked
            if (-not $downloadFolder) {
                Write-Log "Skipped $clientName - folder selection cancelled" -Level Warning
                continue
            }
            $useMiscForEpss = -not [string]::IsNullOrWhiteSpace($script:UserSettings.ReportsBasePath)
            $miscDir = if ($useMiscForEpss) { Join-Path $downloadFolder "Misc" } else { $downloadFolder }
            $outputPathScript = { param($r) $targetDir = if ($useMiscForEpss -and $r.Type -eq 'pending-epss') { $miscDir } else { $downloadFolder }; Get-SafeDownloadPath -TargetDir $targetDir -ClientName $clientName -ReportName $r.Name -Ext $r.Ext -Timestamp $timestamp }

            $batchResult = $null
            $lastErr = $null
            for ($attempt = 1; $attempt -le $maxRetries; $attempt++) {
                try {
                    $retryPart = if ($attempt -gt 1) { " (retry $attempt/$maxRetries)" } else { "" }
                    $lblDownloadProgress.Text = "Downloading for $clientName...$retryPart"
                    $form.Refresh()
                    [System.Windows.Forms.Application]::DoEvents()
                    $timestamp = (Get-Date -Format "yyyy-MM-dd_HH-mm-ss.fff") + "_" + [Guid]::NewGuid().ToString("N").Substring(0, 8)
                    $batchResult = Invoke-ConnectSecureReportsBatch -Reports $standardReports -OutputPathTemplate $outputPathScript -CompanyId $company.Id -ClientName $clientName -ScanDate ($datePickerDownloadScanDate.Value.ToString("MM/dd/yyyy")) -SkipPostDownloadTopX -OnProgress $onProgress
                    if ($script:UserSettings.DownloadAutoResizeColumns -and $batchResult.Succeeded) {
                        foreach ($r in $batchResult.Succeeded) { if ($r.Ext -eq 'xlsx') { $p = & $outputPathScript $r; if (Test-Path -LiteralPath $p) { Invoke-AutoResizeExcelColumns -ExcelPath $p } } }
                    }
                    break
                } catch {
                    $lastErr = $_
                    $isRetryable = $_.Exception.Message -match 'timeout|timed out|connection|404|Unable to connect|reset'
                    if (-not $isRetryable -or $attempt -eq $maxRetries) { break }
                    Start-Sleep -Seconds ([Math]::Min(3 + $attempt * 2, 10))
                }
            }
            if ($batchResult) {
                $successCount += if ($batchResult.Succeeded) { $batchResult.Succeeded.Count } else { 0 }
                $failCount += if ($batchResult.Failed) { $batchResult.Failed.Count } else { 0 }
                Write-Log "Downloaded $($batchResult.Succeeded.Count) standard reports for $clientName" -Level Success
            } else {
                Write-Log "Download failed for $clientName : $($lastErr.Exception.Message)" -Level Error
                $failCount += $standardReports.Count
            }
        }

        $lblDownloadProgress.Text = "Download complete."
        Write-Log "Standard reports download complete. Succeeded: $successCount, Failed: $failCount" -Level Info
        if ($textBoxOutputDir.Text -and (Test-Path $textBoxOutputDir.Text)) {
            $script:UserSettings.LastOutputDirectory = $textBoxOutputDir.Text
            Save-UserSettings | Out-Null
        }
        [System.Windows.Forms.MessageBox]::Show("Standard reports download complete.`nSucceeded: $successCount | Failed: $failCount", "Download Complete", [System.Windows.Forms.MessageBoxButtons]::OK, [System.Windows.Forms.MessageBoxIcon]::Information)
        $btnDownloadStandardOnly.Enabled = $true
    })
    $groupBoxDownload.Controls.Add($btnDownloadStandardOnly)

    $btnDownloadCustom = New-Object System.Windows.Forms.Button
    $btnDownloadCustom.Location = New-Object System.Drawing.Point(270, $dlgY)
    $btnDownloadCustom.Size = New-Object System.Drawing.Size(150, 28)
    $btnDownloadCustom.Text = "Download Custom..."
    $btnDownloadCustom.BackColor = [System.Drawing.Color]::FromArgb(46, 125, 50)
    $btnDownloadCustom.ForeColor = [System.Drawing.Color]::White
    $btnDownloadCustom.FlatStyle = [System.Windows.Forms.FlatStyle]::Flat
    $btnDownloadCustom.FlatAppearance.BorderSize = 0
    $btnDownloadCustom.Add_Click({
        $creds = Load-ConnectSecureCredentials
        if (-not $creds -or [string]::IsNullOrWhiteSpace($creds.BaseUrl) -or [string]::IsNullOrWhiteSpace($creds.ClientId) -or [string]::IsNullOrWhiteSpace($creds.ClientSecret)) {
            [System.Windows.Forms.MessageBox]::Show("Please configure API credentials first. Click 'Settings' then 'API Settings'.", "Credentials Required", [System.Windows.Forms.MessageBoxButtons]::OK, [System.Windows.Forms.MessageBoxIcon]::Warning)
            return
        }
        $checkedCompanies = @($checkedListCompany.CheckedItems)
        if ($checkedCompanies.Count -eq 0) {
            [System.Windows.Forms.MessageBox]::Show("Please select at least one company.", "Validation", [System.Windows.Forms.MessageBoxButtons]::OK, [System.Windows.Forms.MessageBoxIcon]::Warning)
            return
        }
        if ([string]::IsNullOrWhiteSpace($script:UserSettings.ReportsBasePath)) {
            $downloadFolder = $textBoxOutputDir.Text
            if (-not (Test-Path $downloadFolder)) {
                [System.Windows.Forms.MessageBox]::Show("Output directory does not exist. Please select a valid directory.", "Validation", [System.Windows.Forms.MessageBoxButtons]::OK, [System.Windows.Forms.MessageBoxIcon]::Warning)
                return
            }
        }
        $firstCompanyId = if ($checkedCompanies.Count -gt 0 -and $checkedCompanies[0].Id) { $checkedCompanies[0].Id } else { 0 }
        $reports = Show-DownloadCustomReportDialog -CompanyId $firstCompanyId
        if (-not $reports -or $reports.Count -eq 0) { return }
        $btnDownloadCustom.Enabled = $false
        $lblDownloadProgress.Text = "Connecting..."
        $form.Refresh()
        [System.Windows.Forms.Application]::DoEvents()
        $onProgress = { param($m) $lblDownloadProgress.Text = $m; $form.Refresh(); [System.Windows.Forms.Application]::DoEvents() }
        $connected = Connect-ConnectSecureAPI -BaseUrl $creds.BaseUrl -TenantName $creds.TenantName -ClientId $creds.ClientId -ClientSecret $creds.ClientSecret
        if (-not $connected) {
            [System.Windows.Forms.MessageBox]::Show("Authentication failed. Check API Settings.", "Authentication Error", [System.Windows.Forms.MessageBoxButtons]::OK, [System.Windows.Forms.MessageBoxIcon]::Error)
            $btnDownloadCustom.Enabled = $true
            return
        }
        $successCount = 0
        $failCount = 0
        $maxRetries = 3
        $scanDate = $datePickerDownloadScanDate.Value.ToString("MM/dd/yyyy")
        foreach ($company in $checkedCompanies) {
            $clientName = ($company.DisplayName -replace '\s*\(ID:\s*\d+\)\s*$', '').Trim()
            if ([string]::IsNullOrWhiteSpace($clientName)) { $clientName = $company.DisplayName }

            $downloadFolder = Resolve-ClientOutputPath -CompanyId $company.Id -CompanyName $clientName -ScanDate $scanDate -FallbackPath $textBoxOutputDir.Text -ForceManual:$chkReselectFolder.Checked
            if (-not $downloadFolder) {
                Write-Log "Skipped $clientName - folder selection cancelled" -Level Warning
                continue
            }
            $useMiscForEpss = -not [string]::IsNullOrWhiteSpace($script:UserSettings.ReportsBasePath)
            $miscDir = if ($useMiscForEpss) { Join-Path $downloadFolder "Misc" } else { $downloadFolder }
            $outputPathScript = { param($r) $targetDir = if ($useMiscForEpss -and $r.Type -eq 'pending-epss') { $miscDir } else { $downloadFolder }; Get-SafeDownloadPath -TargetDir $targetDir -ClientName $clientName -ReportName $r.Name -Ext $r.Ext -Timestamp $timestamp }

            $batchResult = $null
            for ($attempt = 1; $attempt -le $maxRetries; $attempt++) {
                try {
                    $retryPart = if ($attempt -gt 1) { " (retry $attempt/$maxRetries)" } else { "" }
                    $lblDownloadProgress.Text = "Downloading for $clientName...$retryPart"
                    $form.Refresh()
                    [System.Windows.Forms.Application]::DoEvents()
                    $timestamp = (Get-Date -Format "yyyy-MM-dd_HH-mm-ss.fff") + "_" + [Guid]::NewGuid().ToString("N").Substring(0, 8)
                    $batchResult = Invoke-ConnectSecureReportsBatch -Reports $reports -OutputPathTemplate $outputPathScript -CompanyId $company.Id -ClientName $clientName -ScanDate $scanDate -SkipPostDownloadTopX -OnProgress $onProgress
                    if ($script:UserSettings.DownloadAutoResizeColumns -and $batchResult.Succeeded) {
                        foreach ($r in $batchResult.Succeeded) { if ($r.Ext -eq 'xlsx') { $p = & $outputPathScript $r; if (Test-Path -LiteralPath $p) { Invoke-AutoResizeExcelColumns -ExcelPath $p } } }
                    }
                    break
                } catch {
                    $isRetryable = $_.Exception.Message -match 'timeout|timed out|connection|404|Unable to connect|reset'
                    if (-not $isRetryable -or $attempt -eq $maxRetries) { break }
                    Start-Sleep -Seconds ([Math]::Min(3 + $attempt * 2, 10))
                }
            }
            if ($batchResult) {
                $successCount += if ($batchResult.Succeeded) { $batchResult.Succeeded.Count } else { 0 }
                $failCount += if ($batchResult.Failed) { $batchResult.Failed.Count } else { 0 }
                Write-Log "Downloaded $($batchResult.Succeeded.Count) custom reports for $clientName" -Level Success
            } else {
                $failCount += $reports.Count
            }
        }
        $lblDownloadProgress.Text = "Download complete."
        Write-Log "Custom reports download complete. Succeeded: $successCount, Failed: $failCount" -Level Info
        if ($textBoxOutputDir.Text -and (Test-Path $textBoxOutputDir.Text)) {
            $script:UserSettings.LastOutputDirectory = $textBoxOutputDir.Text
            Save-UserSettings | Out-Null
        }
        [System.Windows.Forms.MessageBox]::Show("Download complete.`nSucceeded: $successCount | Failed: $failCount", "Download Complete", [System.Windows.Forms.MessageBoxButtons]::OK, [System.Windows.Forms.MessageBoxIcon]::Information)
        $btnDownloadCustom.Enabled = $true
    })
    $groupBoxDownload.Controls.Add($btnDownloadCustom)

    $btnDownloadGlobal = New-Object System.Windows.Forms.Button
    $btnDownloadGlobal.Location = New-Object System.Drawing.Point(430, $dlgY)
    $btnDownloadGlobal.Size = New-Object System.Drawing.Size(140, 28)
    $btnDownloadGlobal.Text = "Download Global..."
    $btnDownloadGlobal.BackColor = [System.Drawing.Color]::FromArgb(46, 125, 50)
    $btnDownloadGlobal.ForeColor = [System.Drawing.Color]::White
    $btnDownloadGlobal.FlatStyle = [System.Windows.Forms.FlatStyle]::Flat
    $btnDownloadGlobal.FlatAppearance.BorderSize = 0
    $btnDownloadGlobal.Add_Click({
        $creds = Load-ConnectSecureCredentials
        if (-not $creds -or [string]::IsNullOrWhiteSpace($creds.BaseUrl) -or [string]::IsNullOrWhiteSpace($creds.ClientId) -or [string]::IsNullOrWhiteSpace($creds.ClientSecret)) {
            [System.Windows.Forms.MessageBox]::Show("Please configure API credentials first. Click 'Settings' then 'API Settings'.", "Credentials Required", [System.Windows.Forms.MessageBoxButtons]::OK, [System.Windows.Forms.MessageBoxIcon]::Warning)
            return
        }
        if ([string]::IsNullOrWhiteSpace($script:UserSettings.ReportsBasePath)) {
            $downloadFolder = $textBoxOutputDir.Text
            if (-not (Test-Path $downloadFolder)) {
                [System.Windows.Forms.MessageBox]::Show("Output directory does not exist. Please select a valid directory.", "Validation", [System.Windows.Forms.MessageBoxButtons]::OK, [System.Windows.Forms.MessageBoxIcon]::Warning)
                return
            }
        }
        $reports = Show-DownloadCustomReportDialog -GlobalReports
        if (-not $reports -or $reports.Count -eq 0) { return }
        $downloadFolder = Resolve-ClientOutputPath -CompanyId 0 -CompanyName "Global" -ScanDate ($datePickerDownloadScanDate.Value.ToString("MM/dd/yyyy")) -FallbackPath $textBoxOutputDir.Text -ForceManual:$chkReselectFolder.Checked
        if (-not $downloadFolder) {
            [System.Windows.Forms.MessageBox]::Show("Folder selection was cancelled.", "Cancelled", [System.Windows.Forms.MessageBoxButtons]::OK, [System.Windows.Forms.MessageBoxIcon]::Information)
            return
        }
        $useMiscForEpss = -not [string]::IsNullOrWhiteSpace($script:UserSettings.ReportsBasePath)
        $miscDir = if ($useMiscForEpss) { Join-Path $downloadFolder "Misc" } else { $downloadFolder }
        $outputPathScript = { param($r) $targetDir = if ($useMiscForEpss -and $r.Type -eq 'pending-epss') { $miscDir } else { $downloadFolder }; Get-SafeDownloadPath -TargetDir $targetDir -ClientName "Global" -ReportName $r.Name -Ext $r.Ext -Timestamp $timestamp }
        $btnDownloadGlobal.Enabled = $false
        $lblDownloadProgress.Text = "Connecting..."
        $form.Refresh()
        [System.Windows.Forms.Application]::DoEvents()
        $onProgress = { param($m) $lblDownloadProgress.Text = $m; $form.Refresh(); [System.Windows.Forms.Application]::DoEvents() }
        $connected = Connect-ConnectSecureAPI -BaseUrl $creds.BaseUrl -TenantName $creds.TenantName -ClientId $creds.ClientId -ClientSecret $creds.ClientSecret
        if (-not $connected) {
            [System.Windows.Forms.MessageBox]::Show("Authentication failed. Check API Settings.", "Authentication Error", [System.Windows.Forms.MessageBoxButtons]::OK, [System.Windows.Forms.MessageBoxIcon]::Error)
            $btnDownloadGlobal.Enabled = $true
            return
        }
        $scanDate = $datePickerDownloadScanDate.Value.ToString("MM/dd/yyyy")
        $timestamp = (Get-Date -Format "yyyy-MM-dd_HH-mm-ss.fff") + "_" + [Guid]::NewGuid().ToString("N").Substring(0, 8)
        try {
            $lblDownloadProgress.Text = "Downloading global reports..."
            $form.Refresh()
            $batchResult = Invoke-ConnectSecureReportsBatch -Reports $reports -OutputPathTemplate $outputPathScript -CompanyId 0 -ClientName "Global" -ScanDate $scanDate -SkipPostDownloadTopX -OnProgress $onProgress
            if ($script:UserSettings.DownloadAutoResizeColumns -and $batchResult.Succeeded) {
                foreach ($r in $batchResult.Succeeded) { if ($r.Ext -eq 'xlsx') { $p = & $outputPathScript $r; if (Test-Path -LiteralPath $p) { Invoke-AutoResizeExcelColumns -ExcelPath $p } } }
            }
            $successCount = if ($batchResult.Succeeded) { $batchResult.Succeeded.Count } else { 0 }
            $failCount = if ($batchResult.Failed) { $batchResult.Failed.Count } else { 0 }
            [System.Windows.Forms.MessageBox]::Show("Download complete.`nSucceeded: $successCount | Failed: $failCount", "Download Complete", [System.Windows.Forms.MessageBoxButtons]::OK, [System.Windows.Forms.MessageBoxIcon]::Information)
        } catch {
            [System.Windows.Forms.MessageBox]::Show("Download failed: $($_.Exception.Message)", "Error", [System.Windows.Forms.MessageBoxButtons]::OK, [System.Windows.Forms.MessageBoxIcon]::Error)
        }
        $lblDownloadProgress.Text = "Download complete."
        if ($textBoxOutputDir.Text -and (Test-Path $textBoxOutputDir.Text)) {
            $script:UserSettings.LastOutputDirectory = $textBoxOutputDir.Text
            Save-UserSettings | Out-Null
        }
        $btnDownloadGlobal.Enabled = $true
    })
    $groupBoxDownload.Controls.Add($btnDownloadGlobal)

    # Initialize checkedListCompany with All Companies (from saved credentials + cache)
    $script:PopulateCheckedListCompany = {
        $checkedListCompany.Items.Clear()
        $checkedListCompany.Items.Add([PSCustomObject]@{ Id = 0; DisplayName = "All Companies" }) | Out-Null
        $savedCreds = Load-ConnectSecureCredentials
        if ($savedCreds) {
            $cached = Load-ConnectSecureCompaniesCache -BaseUrl $savedCreds.BaseUrl -TenantName $savedCreds.TenantName
            if ($cached) {
                foreach ($c in ($cached | Sort-Object { $_.DisplayName })) {
                    $opt = [PSCustomObject]@{ Id = $c.Id; DisplayName = $c.DisplayName }
                    $checked = ($opt.Id -eq $savedCreds.CompanyId)
                    $idx = $checkedListCompany.Items.Add($opt)
                    if ($checked) { $checkedListCompany.SetItemChecked($idx, $true) }
                }
            }
        }
    }
    & $script:PopulateCheckedListCompany

    $btnRefreshCompanies.Add_Click({
        $creds = Load-ConnectSecureCredentials
        if (-not $creds -or [string]::IsNullOrWhiteSpace($creds.BaseUrl) -or [string]::IsNullOrWhiteSpace($creds.ClientId) -or [string]::IsNullOrWhiteSpace($creds.ClientSecret)) {
            [System.Windows.Forms.MessageBox]::Show("Please configure API credentials first. Click 'API Settings' to enter Base URL, Tenant, Client ID, and Client Secret.", "Credentials Required", [System.Windows.Forms.MessageBoxButtons]::OK, [System.Windows.Forms.MessageBoxIcon]::Warning)
            return
        }
        $btnRefreshCompanies.Enabled = $false
        $btnRefreshCompanies.Text = "Loading..."
        $form.Refresh()
        try {
            $connected = Connect-ConnectSecureAPI -BaseUrl $creds.BaseUrl -TenantName $creds.TenantName -ClientId $creds.ClientId -ClientSecret $creds.ClientSecret
            if (-not $connected) {
                [System.Windows.Forms.MessageBox]::Show("Failed to authenticate. Check your API Settings.", "Authentication Failed", [System.Windows.Forms.MessageBoxButtons]::OK, [System.Windows.Forms.MessageBoxIcon]::Error)
                return
            }
            $companies = Get-ConnectSecureCompanies -FetchAll
            $checkedIds = @($checkedListCompany.CheckedItems | ForEach-Object { $_.Id })
            $checkedListCompany.Items.Clear()
            $checkedListCompany.Items.Add([PSCustomObject]@{ Id = 0; DisplayName = "All Companies" }) | Out-Null
            $toCache = [System.Collections.ArrayList]::new()
            $idx = 1
            foreach ($company in $companies) {
                $info = Get-ConnectSecureCompanyDisplayInfo -Company $company
                $idStr = if ($null -ne $info.Id -and $info.Id -ne "") { $info.Id.ToString() } else { $null }
                $companyName = if ([string]::IsNullOrWhiteSpace($info.Name)) { $null } else { $info.Name }
                $displayName = if ($companyName) { if ($idStr) { "$companyName (ID: $idStr)" } else { $companyName } } elseif ($idStr) { "Company (ID: $idStr)" } else { "Company $idx" }
                $opt = [PSCustomObject]@{ Id = $info.Id; DisplayName = $displayName }
                [void]$toCache.Add($opt)
                $idx++
            }
            foreach ($o in ($toCache | Sort-Object { $_.DisplayName })) {
                $i = $checkedListCompany.Items.Add($o)
                if ($o.Id -in $checkedIds) { $checkedListCompany.SetItemChecked($i, $true) }
            }
            Save-ConnectSecureCompaniesCache -BaseUrl $creds.BaseUrl -TenantName ($creds.TenantName.Trim()) -Companies @($toCache)
            [System.Windows.Forms.MessageBox]::Show("Loaded $($companies.Count) companies.", "Success", [System.Windows.Forms.MessageBoxButtons]::OK, [System.Windows.Forms.MessageBoxIcon]::Information)
        } catch {
            [System.Windows.Forms.MessageBox]::Show("Error: $($_.Exception.Message)", "Error", [System.Windows.Forms.MessageBoxButtons]::OK, [System.Windows.Forms.MessageBoxIcon]::Error)
        } finally {
            $btnRefreshCompanies.Enabled = $true
            $btnRefreshCompanies.Text = "Refresh List"
        }
    })

    # --- 2. Process from file ---
    $groupBoxManual = New-Object System.Windows.Forms.GroupBox
    $groupBoxManual.Location = New-Object System.Drawing.Point(20, 443)
    $groupBoxManual.Size = New-Object System.Drawing.Size(680, 95)
    $groupBoxManual.Text = "2. Or process a previously downloaded file"
    $form.Controls.Add($groupBoxManual)

    $labelInputFile = New-Object System.Windows.Forms.Label
    $labelInputFile.Location = New-Object System.Drawing.Point(20, 18)
    $labelInputFile.Size = New-Object System.Drawing.Size(180, 18)
    $labelInputFile.Text = "All Vulnerabilities Report (XLSX):"
    $groupBoxManual.Controls.Add($labelInputFile)

    $textBoxInputFile = New-Object System.Windows.Forms.TextBox
    $textBoxInputFile.Location = New-Object System.Drawing.Point(20, 38)
    $textBoxInputFile.Size = New-Object System.Drawing.Size(470, 20)
    $textBoxInputFile.ReadOnly = $true
    $groupBoxManual.Controls.Add($textBoxInputFile)

    $buttonBrowseInput = New-Object System.Windows.Forms.Button
    $buttonBrowseInput.Location = New-Object System.Drawing.Point(500, 36)
    $buttonBrowseInput.Size = New-Object System.Drawing.Size(80, 25)
    $buttonBrowseInput.Text = "Browse..."
    $buttonBrowseInput.Add_Click({
        $openFileDialog = New-Object System.Windows.Forms.OpenFileDialog
        $openFileDialog.Filter = "Excel Files (*.xlsx)|*.xlsx|All Files (*.*)|*.*"
        $openFileDialog.Title = "Select All Vulnerabilities Report"

        if ($openFileDialog.ShowDialog() -eq [System.Windows.Forms.DialogResult]::OK) {
            $textBoxInputFile.Text = $openFileDialog.FileName

            # Automatically set output directory to input file's directory
            $inputDirectory = [System.IO.Path]::GetDirectoryName($openFileDialog.FileName)
            if ($textBoxOutputDir) { $textBoxOutputDir.Text = $inputDirectory }

            # Extract company name from filename
            # Try multiple patterns to detect company name
            $fileName = [System.IO.Path]::GetFileNameWithoutExtension($openFileDialog.FileName)
            Write-Log "Attempting to extract company name from filename: $fileName"

            $companyName = $null
            # Pattern 1: "...Reports-{CompanyName}_{timestamp}" or "...Report-{CompanyName}_..." or "...Reports-{CompanyName} "
            # Captures company name with spaces until underscore, end of string, or space before date/timestamp
            if ($fileName -match 'Reports?[-_]\s*([^_]+?)(?:_\d|$|\s+\d)') {
                $rawName = $matches[1].Trim()
                # Skip if it's a report-related keyword
                if ($rawName -notmatch '^(Pending|EPSS|Report|Reports?|Vulnerability|Security)$') {
                    $companyName = $rawName
                    Write-Log "Matched Pattern 1 (Reports-Company_): $companyName"
                } else {
                    Write-Log "Pattern 1 matched but result was report keyword: $rawName" -Level Warning
                }
            }
            # Pattern 2: "{CompanyName}-Reports" or "{CompanyName}_Reports" or "{CompanyName} Reports"
            # Captures company name with spaces before Reports/Report
            if (-not $companyName -and $fileName -match '^(.+?)[\s_-]+Reports?') {
                $rawName = $matches[1].Trim()
                if ($rawName -notmatch '^(Pending|EPSS|Report|Reports?|Vulnerability|Security)$' -and $rawName.Length -gt 0) {
                    $companyName = $rawName
                    Write-Log "Matched Pattern 2 (Company-Reports): $companyName"
                } else {
                    Write-Log "Pattern 2 matched but result was report keyword: $rawName" -Level Warning
                }
            }
            # Pattern 3: Extract text before first underscore or hyphen (preserving spaces)
            # But exclude if it contains report-related keywords
            if (-not $companyName -and $fileName -match '^([^_-]+)') {
                $rawName = $matches[1].Trim()
                if ($rawName -notmatch '(Pending|EPSS|Report|Reports?|Vulnerability|Security)' -and $rawName.Length -gt 0) {
                    $companyName = $rawName
                    Write-Log "Matched Pattern 3 (first segment, not keyword): $companyName"
                } else {
                    Write-Log "Pattern 3 matched but result contained report keyword: $rawName" -Level Warning
                }
            }

            if ($companyName) {
                $textBoxClientName.Text = $companyName
                Write-Log "Company name set to: $companyName"
            } else {
                Write-Log "Could not extract company name from filename" -Level Warning
            }
        }
    })
    $groupBoxManual.Controls.Add($buttonBrowseInput)

    $labelClientName = New-Object System.Windows.Forms.Label
    $labelClientName.Location = New-Object System.Drawing.Point(20, 68)
    $labelClientName.Size = New-Object System.Drawing.Size(55, 18)
    $labelClientName.Text = "Client:"
    $groupBoxManual.Controls.Add($labelClientName)

    $textBoxClientName = New-Object System.Windows.Forms.TextBox
    $textBoxClientName.Location = New-Object System.Drawing.Point(75, 65)
    $textBoxClientName.Size = New-Object System.Drawing.Size(140, 20)
    $groupBoxManual.Controls.Add($textBoxClientName)

    $labelScanDate = New-Object System.Windows.Forms.Label
    $labelScanDate.Location = New-Object System.Drawing.Point(218, 68)
    $labelScanDate.Size = New-Object System.Drawing.Size(62, 18)
    $labelScanDate.Text = "Scan Date:"
    $groupBoxManual.Controls.Add($labelScanDate)

    $datePickerScanDate = New-Object System.Windows.Forms.DateTimePicker
    $datePickerScanDate.Location = New-Object System.Drawing.Point(283, 65)
    $datePickerScanDate.Size = New-Object System.Drawing.Size(130, 20)
    $datePickerScanDate.Format = [System.Windows.Forms.DateTimePickerFormat]::Short
    $groupBoxManual.Controls.Add($datePickerScanDate)

    # --- Output Directory ---
    $labelOutputDir = New-Object System.Windows.Forms.Label
    $labelOutputDir.Location = New-Object System.Drawing.Point(20, 553)
    $labelOutputDir.Size = New-Object System.Drawing.Size(150, 20)
    $labelOutputDir.Text = "Output Directory:"
    $form.Controls.Add($labelOutputDir)

    $textBoxOutputDir = New-Object System.Windows.Forms.TextBox
    $textBoxOutputDir.Location = New-Object System.Drawing.Point(20, 578)
    $textBoxOutputDir.Size = New-Object System.Drawing.Size(570, 20)
    $textBoxOutputDir.Text = if ($script:UserSettings.LastOutputDirectory -and (Test-Path $script:UserSettings.LastOutputDirectory)) { $script:UserSettings.LastOutputDirectory } else { [Environment]::GetFolderPath("Desktop") }
    $form.Controls.Add($textBoxOutputDir)

    $buttonBrowseOutput = New-Object System.Windows.Forms.Button
    $buttonBrowseOutput.Location = New-Object System.Drawing.Point(600, 548)
    $buttonBrowseOutput.Size = New-Object System.Drawing.Size(80, 25)
    $buttonBrowseOutput.Text = "Browse..."
    $buttonBrowseOutput.Add_Click({
        $folderBrowser = New-Object System.Windows.Forms.FolderBrowserDialog
        $folderBrowser.Description = "Select Output Directory"
        $folderBrowser.SelectedPath = $textBoxOutputDir.Text

        if ($folderBrowser.ShowDialog() -eq [System.Windows.Forms.DialogResult]::OK) {
            $textBoxOutputDir.Text = $folderBrowser.SelectedPath
            $script:UserSettings.LastOutputDirectory = $folderBrowser.SelectedPath
            Save-UserSettings | Out-Null
        }
    })
    $form.Controls.Add($buttonBrowseOutput)

    $toolTipOutput = New-Object System.Windows.Forms.ToolTip
    $lblStructuredPathHint = New-Object System.Windows.Forms.Label
    $lblStructuredPathHint.Location = New-Object System.Drawing.Point(20, 576)
    $lblStructuredPathHint.Size = New-Object System.Drawing.Size(560, 18)
    $lblStructuredPathHint.AutoEllipsis = $true
    $lblStructuredPathHint.ForeColor = [System.Drawing.Color]::Gray
    $lblStructuredPathHint.Font = New-Object System.Drawing.Font($lblStructuredPathHint.Font.FontFamily, 8.5)
    $form.Controls.Add($lblStructuredPathHint)

    $chkReselectFolder = New-Object System.Windows.Forms.CheckBox
    $chkReselectFolder.Location = New-Object System.Drawing.Point(20, 596)
    $chkReselectFolder.Size = New-Object System.Drawing.Size(220, 20)
    $chkReselectFolder.Text = "Re-select folder(s) for this run"
    $chkReselectFolder.ForeColor = [System.Drawing.Color]::Gray
    $form.Controls.Add($chkReselectFolder)

    $btnEditMappings = New-Object System.Windows.Forms.Button
    $btnEditMappings.Location = New-Object System.Drawing.Point(250, 594)
    $btnEditMappings.Size = New-Object System.Drawing.Size(210, 24)
    $btnEditMappings.Text = "Edit Output Folder Mappings"
    $btnEditMappings.Add_Click({ Show-CompanyFolderMappingDialog })
    $form.Controls.Add($btnEditMappings)

    $btnReportHistory = New-Object System.Windows.Forms.Button
    $btnReportHistory.Location = New-Object System.Drawing.Point(470, 594)
    $btnReportHistory.Size = New-Object System.Drawing.Size(130, 24)
    $btnReportHistory.Text = "Report Folder History"
    $btnReportHistory.Add_Click({ Show-ReportFolderHistoryDialog })
    $form.Controls.Add($btnReportHistory)

    $script:UpdateOutputDirUI = {
        $structuredOffset = 0
        if (-not [string]::IsNullOrWhiteSpace($script:UserSettings.ReportsBasePath) -and (Test-Path $script:UserSettings.ReportsBasePath)) {
            $labelOutputDir.Text = "Output Directory:"
            $labelOutputDir.Visible = $true
            $textBoxOutputDir.Visible = $true
            $buttonBrowseOutput.Visible = $true
            $lblStructuredPathHint.Text = "Leave empty for structured paths under base, or enter/browse a path to use a one-off location"
            $lblStructuredPathHint.Location = New-Object System.Drawing.Point(20, 600)
            $lblStructuredPathHint.Visible = $true
            $toolTipOutput.SetToolTip($lblStructuredPathHint, $script:UserSettings.ReportsBasePath)
            $chkReselectFolder.Location = New-Object System.Drawing.Point(20, 620)
            $chkReselectFolder.Visible = $true
            $btnEditMappings.Location = New-Object System.Drawing.Point(250, 618)
            $btnEditMappings.Visible = $true
            $btnReportHistory.Location = New-Object System.Drawing.Point(470, 618)
            $btnReportHistory.Visible = $true
            $structuredOffset = 24
        } else {
            $lblStructuredPathHint.Location = New-Object System.Drawing.Point(20, 576)
            $chkReselectFolder.Location = New-Object System.Drawing.Point(20, 596)
            $btnEditMappings.Location = New-Object System.Drawing.Point(250, 594)
            $btnReportHistory.Location = New-Object System.Drawing.Point(470, 594)
            $labelOutputDir.Text = "Output Directory:"
            $labelOutputDir.Visible = $true
            $textBoxOutputDir.Visible = $true
            $buttonBrowseOutput.Visible = $true
            $lblStructuredPathHint.Visible = $false
            $chkReselectFolder.Visible = $false
            $btnEditMappings.Visible = $false
            $btnReportHistory.Visible = $true
            $toolTipOutput.SetToolTip($lblStructuredPathHint, $null)
        }
        if ($script:StatusLabel) {
            $baseY = 620 + $structuredOffset   # 8px padding below Output (checkbox/button row)
            $script:StatusLabel.Location = New-Object System.Drawing.Point(20, $baseY)
            $script:ProgressBar.Location = New-Object System.Drawing.Point(20, ($baseY + 25))
            $labelLog.Location = New-Object System.Drawing.Point(20, ($baseY + 50))
            $script:LogTextBox.Location = New-Object System.Drawing.Point(20, ($baseY + 75))
            $labelTicketNotes.Location = New-Object System.Drawing.Point(20, ($baseY + 160))
            $buttonCopyTicketNotes.Location = New-Object System.Drawing.Point(20, ($baseY + 185))
            $script:buttonOpenTicketNotes.Location = New-Object System.Drawing.Point(160, ($baseY + 185))
            $panelBottomButtons.Location = New-Object System.Drawing.Point(20, ($baseY + 220))
        }
    }

    # --- Progress Section ---
    $script:StatusLabel = New-Object System.Windows.Forms.Label
    $script:StatusLabel.Location = New-Object System.Drawing.Point(20, 620)
    $script:StatusLabel.Size = New-Object System.Drawing.Size(660, 20)
    $script:StatusLabel.Text = "Ready"
    $script:StatusLabel.Visible = $false
    $form.Controls.Add($script:StatusLabel)

    $script:ProgressBar = New-Object System.Windows.Forms.ProgressBar
    $script:ProgressBar.Location = New-Object System.Drawing.Point(20, 625)
    $script:ProgressBar.Size = New-Object System.Drawing.Size(660, 20)
    $script:ProgressBar.Style = 'Marquee'
    $script:ProgressBar.MarqueeAnimationSpeed = 30
    $script:ProgressBar.Visible = $false
    $form.Controls.Add($script:ProgressBar)

    # --- Log Section ---
    $labelLog = New-Object System.Windows.Forms.Label
    $labelLog.Location = New-Object System.Drawing.Point(20, 630)
    $labelLog.Size = New-Object System.Drawing.Size(150, 20)
    $labelLog.Text = "Processing Log:"
    $form.Controls.Add($labelLog)

    $script:LogTextBox = New-Object System.Windows.Forms.TextBox
    $script:LogTextBox.Location = New-Object System.Drawing.Point(20, 665)
    $script:LogTextBox.Size = New-Object System.Drawing.Size(660, 80)
    $script:LogTextBox.Multiline = $true
    $script:LogTextBox.ScrollBars = "Vertical"
    $script:LogTextBox.ReadOnly = $true
    $script:LogTextBox.Font = New-Object System.Drawing.Font("Consolas", 9)
    $form.Controls.Add($script:LogTextBox)

    # --- Ticket Notes Section Label ---
    $labelTicketNotes = New-Object System.Windows.Forms.Label
    $labelTicketNotes.Location = New-Object System.Drawing.Point(20, 780)
    $labelTicketNotes.Size = New-Object System.Drawing.Size(200, 20)
    $labelTicketNotes.Text = "Ticket Notes:"
    $form.Controls.Add($labelTicketNotes)

    # --- Ticket Notes Buttons ---
    $buttonCopyTicketNotes = New-Object System.Windows.Forms.Button
    $buttonCopyTicketNotes.Location = New-Object System.Drawing.Point(20, 805)
    $buttonCopyTicketNotes.Size = New-Object System.Drawing.Size(130, 25)
    $buttonCopyTicketNotes.Text = "Copy to Clipboard"
    $buttonCopyTicketNotes.Add_Click({
        New-TicketNotes -Top10Data $script:CurrentTop10Data -TimeEstimates $script:CurrentTimeEstimates -IsRMITPlus $script:IsRMITPlus
    })
    $form.Controls.Add($buttonCopyTicketNotes)

    $script:buttonOpenTicketNotes = New-Object System.Windows.Forms.Button
    $script:buttonOpenTicketNotes.Location = New-Object System.Drawing.Point(160, 805)
    $script:buttonOpenTicketNotes.Size = New-Object System.Drawing.Size(130, 25)
    $script:buttonOpenTicketNotes.Text = "View Ticket Notes"
    $script:buttonOpenTicketNotes.Enabled = $false
    $script:buttonOpenTicketNotes.Add_Click({
        if ($script:TicketNotesPath -and (Test-Path $script:TicketNotesPath)) {
            Start-Process $script:TicketNotesPath
        }
    })
    $form.Controls.Add($script:buttonOpenTicketNotes)

    # --- Action Buttons (Bottom row: Generate = main; Remediation Rules | Settings | Help | Close) ---
    $panelBottomButtons = New-Object System.Windows.Forms.Panel
    $panelBottomButtons.Location = New-Object System.Drawing.Point(20, 840)
    $panelBottomButtons.Size = New-Object System.Drawing.Size(680, 35)
    $form.Controls.Add($panelBottomButtons)

    & $script:UpdateOutputDirUI

    $buttonGenerate = New-Object System.Windows.Forms.Button
    $buttonGenerate.Location = New-Object System.Drawing.Point(0, 0)
    $buttonGenerate.Size = New-Object System.Drawing.Size(220, 30)
    $panelBottomButtons.Controls.Add($buttonGenerate)

    $buttonRemediationRules = New-Object System.Windows.Forms.Button
    $buttonRemediationRules.Location = New-Object System.Drawing.Point(230, 0)
    $buttonRemediationRules.Size = New-Object System.Drawing.Size(140, 30)
    $buttonRemediationRules.Text = "Remediation Rules"
    $buttonRemediationRules.Add_Click({
        Show-RemediationRulesDialog
    })
    $panelBottomButtons.Controls.Add($buttonRemediationRules)

    $buttonSettings = New-Object System.Windows.Forms.Button
    $buttonSettings.Location = New-Object System.Drawing.Point(380, 0)
    $buttonSettings.Size = New-Object System.Drawing.Size(100, 30)
    $buttonSettings.Text = "Settings"
    $buttonSettings.BackColor = [System.Drawing.Color]::FromArgb(94, 53, 177)
    $buttonSettings.ForeColor = [System.Drawing.Color]::White
    $buttonSettings.FlatStyle = [System.Windows.Forms.FlatStyle]::Flat
    $buttonSettings.FlatAppearance.BorderSize = 0
    $buttonSettings.Add_Click({
        Show-SettingsDialog
        if ($script:UpdateOutputDirUI) { & $script:UpdateOutputDirUI }
    })
    $panelBottomButtons.Controls.Add($buttonSettings)

    $buttonHelp = New-Object System.Windows.Forms.Button
    $buttonHelp.Location = New-Object System.Drawing.Point(490, 0)
    $buttonHelp.Size = New-Object System.Drawing.Size(80, 30)
    $buttonHelp.Text = "Help"
    $buttonHelp.BackColor = [System.Drawing.Color]::FromArgb(255, 167, 38)
    $buttonHelp.ForeColor = [System.Drawing.Color]::White
    $buttonHelp.FlatStyle = [System.Windows.Forms.FlatStyle]::Flat
    $buttonHelp.FlatAppearance.BorderSize = 0
    $buttonHelp.Add_Click({ Show-VScanMagicOverviewHelpDialog })
    $panelBottomButtons.Controls.Add($buttonHelp)

    $buttonGenerate.Text = "Download and Generate Reports"
    $buttonGenerate.BackColor = [System.Drawing.Color]::FromArgb(46, 125, 50)  # Green - Primary action (download + generate)
    $buttonGenerate.ForeColor = [System.Drawing.Color]::White
    $buttonGenerate.FlatStyle = [System.Windows.Forms.FlatStyle]::Flat
    $buttonGenerate.FlatAppearance.BorderSize = 0
    $toolTipOutput.SetToolTip($buttonGenerate, "Main action: Downloads reports from ConnectSecure (if configured) and/or processes files to generate Word, Excel, and other outputs.")
    $buttonGenerate.Add_Click({
        # Validate: need either (a) input file selected, or (b) API creds + at least one company checked + Pending EPSS for download
        $hasFile = -not [string]::IsNullOrWhiteSpace($textBoxInputFile.Text)
        $creds = Load-ConnectSecureCredentials
        $canDownload = $creds -and -not [string]::IsNullOrWhiteSpace($creds.BaseUrl) -and -not [string]::IsNullOrWhiteSpace($creds.ClientId) -and -not [string]::IsNullOrWhiteSpace($creds.ClientSecret) -and $chkAllVulnerabilities.Checked
        $checkedCompanies = @($checkedListCompany.CheckedItems)
        $hasCompanies = $checkedCompanies.Count -gt 0

        if (-not $hasFile) {
            if (-not $canDownload -or -not $hasCompanies) {
                [System.Windows.Forms.MessageBox]::Show("Please select an input file (Browse) OR configure API credentials, select one or more companies, and ensure All Vulnerabilities Report is checked to download first.", "Validation Error",
                    [System.Windows.Forms.MessageBoxButtons]::OK, [System.Windows.Forms.MessageBoxIcon]::Warning)
                return
            }
        } elseif ([string]::IsNullOrWhiteSpace($textBoxClientName.Text)) {
            [System.Windows.Forms.MessageBox]::Show("Please enter a client name.", "Validation Error",
                [System.Windows.Forms.MessageBoxButtons]::OK, [System.Windows.Forms.MessageBoxIcon]::Warning)
            return
        }

        if (-not $script:OutputExcel -and -not $script:OutputWord -and -not $script:OutputEmailTemplate -and -not $script:OutputTicketInstructions -and -not $script:OutputTimeEstimate) {
            [System.Windows.Forms.MessageBox]::Show("Please select at least one output option.", "Validation Error",
                [System.Windows.Forms.MessageBoxButtons]::OK, [System.Windows.Forms.MessageBoxIcon]::Warning)
            return
        }

        # Capture client type selection
        $script:IsRMITPlus = $checkBoxRMITPlus.Checked

        # Disable button during processing
        $buttonGenerate.Enabled = $false
        $script:LogTextBox.Clear()

        # Disable open buttons at start
        $script:buttonOpenTicketNotes.Enabled = $false

        $isAsyncPath = $false
        try {
            $companiesToProcess = if ($hasFile) {
                $resolvedPath = Resolve-ClientOutputPath -CompanyId 0 -CompanyName $textBoxClientName.Text -ScanDate ($datePickerScanDate.Value.ToShortDateString()) -FallbackPath $textBoxOutputDir.Text -ForceManual:$chkReselectFolder.Checked
                if (-not $resolvedPath) {
                    [System.Windows.Forms.MessageBox]::Show("Folder selection was cancelled.", "Cancelled", [System.Windows.Forms.MessageBoxButtons]::OK, [System.Windows.Forms.MessageBoxIcon]::Information)
                    $buttonGenerate.Enabled = $true
                    return
                }
                @([PSCustomObject]@{ Id = 0; DisplayName = $textBoxClientName.Text; InputPath = $textBoxInputFile.Text; OutputDir = $resolvedPath })
            } else {
                @($checkedCompanies)
            }

            $downloadFolder = $textBoxOutputDir.Text
            $reports = @()
            if ($chkAllVulnerabilities.Checked) { $reports += @{ Type = "all-vulnerabilities"; Name = "All Vulnerabilities Report"; Ext = "xlsx" } }
            if ($chkSuppressedVulnerabilities.Checked) { $reports += @{ Type = "suppressed-vulnerabilities"; Name = "Suppressed Vulnerabilities"; Ext = "xlsx" } }
            if ($chkExternalVulnerabilities.Checked) { $reports += @{ Type = "external-vulnerabilities"; Name = "External Scan"; Ext = "xlsx" } }
            if ($chkExecutiveSummary.Checked) { $reports += @{ Type = "executive-summary"; Name = "Executive Summary Report"; Ext = "docx" } }
            if ($chkPendingEPSS.Checked) { $reports += @{ Type = "pending-epss"; Name = "Pending Remediation EPSS Score Reports"; Ext = "xlsx" } }
            $topCount = if ($script:FilterTopN -eq "All") { 500 } else { [int]$script:FilterTopN }

            if (-not $hasFile) {
                if ([string]::IsNullOrWhiteSpace($script:UserSettings.ReportsBasePath) -and -not (Test-Path $downloadFolder)) {
                    [System.Windows.Forms.MessageBox]::Show("Output directory does not exist. Please select a valid directory.", "Validation Error",
                        [System.Windows.Forms.MessageBoxButtons]::OK, [System.Windows.Forms.MessageBoxIcon]::Warning)
                    $buttonGenerate.Enabled = $true
                    return
                }
                $isAsyncPath = $true
                $lblDownloadProgress.Text = "Connecting..."
                $form.Refresh()
                [System.Windows.Forms.Application]::DoEvents()

                $connected = Connect-ConnectSecureAPI -BaseUrl $creds.BaseUrl -TenantName $creds.TenantName -ClientId $creds.ClientId -ClientSecret $creds.ClientSecret
                if (-not $connected) {
                    [System.Windows.Forms.MessageBox]::Show("Authentication failed. Check API Settings.", "Authentication Error", [System.Windows.Forms.MessageBoxButtons]::OK, [System.Windows.Forms.MessageBoxIcon]::Error)
                    $buttonGenerate.Enabled = $true
                    return
                }

                $companiesData = [System.Collections.ArrayList]::new()
                $scanDate = $datePickerDownloadScanDate.Value.ToString("MM/dd/yyyy")
                $maxRetries = 3
                $onProgress = { param($m) $lblDownloadProgress.Text = $m; $form.Refresh(); [System.Windows.Forms.Application]::DoEvents() }
                foreach ($company in $checkedCompanies) {
                    $clientName = ($company.DisplayName -replace '\s*\(ID:\s*\d+\)\s*$', '').Trim()
                    if ([string]::IsNullOrWhiteSpace($clientName)) { $clientName = $company.DisplayName }

                    $downloadFolder = Resolve-ClientOutputPath -CompanyId $company.Id -CompanyName $clientName -ScanDate $scanDate -FallbackPath $textBoxOutputDir.Text -ForceManual:$chkReselectFolder.Checked
                    if (-not $downloadFolder) {
                        Write-Log "Skipped $clientName - folder selection cancelled" -Level Warning
                        continue
                    }
                    $useMiscForEpss = -not [string]::IsNullOrWhiteSpace($script:UserSettings.ReportsBasePath)
                    $miscDir = if ($useMiscForEpss) { Join-Path $downloadFolder "Misc" } else { $downloadFolder }
                    $outputPathScript = { param($r) $targetDir = if ($useMiscForEpss -and $r.Type -eq 'pending-epss') { $miscDir } else { $downloadFolder }; Get-SafeDownloadPath -TargetDir $targetDir -ClientName $clientName -ReportName $r.Name -Ext $r.Ext -Timestamp $timestamp }

                    $batchResult = $null
                    $lastErr = $null
                    for ($attempt = 1; $attempt -le $maxRetries; $attempt++) {
                        try {
                            $retryPart = if ($attempt -gt 1) { " (retry $attempt/$maxRetries)" } else { "" }
                            $lblDownloadProgress.Text = "Downloading for $clientName...$retryPart"
                            $form.Refresh()
                            [System.Windows.Forms.Application]::DoEvents()
                            $timestamp = (Get-Date -Format "yyyy-MM-dd_HH-mm-ss.fff") + "_" + [Guid]::NewGuid().ToString("N").Substring(0, 8)
                            $batchResult = Invoke-ConnectSecureReportsBatch -Reports $reports -OutputPathTemplate $outputPathScript -CompanyId $company.Id -ClientName $clientName -ScanDate $scanDate -TopCount $topCount -MinEPSS $script:FilterMinEPSS -IncludeCritical $script:FilterIncludeCritical -IncludeHigh $script:FilterIncludeHigh -IncludeMedium $script:FilterIncludeMedium -IncludeLow $script:FilterIncludeLow -SkipPostDownloadTopX -OnProgress $onProgress
                            if ($script:UserSettings.DownloadAutoResizeColumns -and $batchResult.Succeeded) {
                                foreach ($r in $batchResult.Succeeded) { if ($r.Ext -eq 'xlsx') { $p = & $outputPathScript $r; if (Test-Path -LiteralPath $p) { Invoke-AutoResizeExcelColumns -ExcelPath $p } } }
                            }
                            break
                        } catch {
                            $lastErr = $_
                            $isRetryable = $_.Exception.Message -match 'timeout|timed out|connection|404|Unable to connect|reset'
                            if (-not $isRetryable -or $attempt -eq $maxRetries) { break }
                            Start-Sleep -Seconds ([Math]::Min(3 + $attempt * 2, 10))
                        }
                    }
                    if (-not $batchResult) { continue }
                    $inputReport = $batchResult.Succeeded | Where-Object { $_.Type -eq "all-vulnerabilities" } | Select-Object -First 1
                    $inputPath = if ($inputReport) { & $outputPathScript $inputReport } else { $null }
                    if ($inputPath -and (Test-Path $inputPath)) {
                        $null = $companiesData.Add(@{ Company = $company; InputPath = $inputPath; ClientName = $clientName; ScanDate = $scanDate; OutputDir = $downloadFolder })
                    }
                }

                if ($companiesData.Count -eq 0) {
                    [System.Windows.Forms.MessageBox]::Show("No companies were downloaded successfully. Check that All Vulnerabilities report is selected and that the reports download correctly (try Download Standard Reports first).", "Download Error", [System.Windows.Forms.MessageBoxButtons]::OK, [System.Windows.Forms.MessageBoxIcon]::Warning)
                    $buttonGenerate.Enabled = $true
                    return
                }

                $lblDownloadProgress.Text = "Download complete. Processing..."
                $form.Refresh()
                [System.Windows.Forms.Application]::DoEvents()
                $companiesToProcess = @($companiesData)
                # Set RMIT+ per client when processing multiple companies
                if ($companiesToProcess.Count -gt 1) {
                    $clientTypes = Show-SetClientTypesDialog -CompaniesToProcess $companiesToProcess -DefaultIsRMITPlus $checkBoxRMITPlus.Checked
                    if ($null -eq $clientTypes) {
                        Write-Log "Client type selection cancelled." -Level Warning
                        $buttonGenerate.Enabled = $true
                        return
                    }
                    foreach ($companyData in $companiesToProcess) {
                        $companyId = if ($companyData.Company -and $companyData.Company.Id -ne $null) { $companyData.Company.Id } else { 0 }
                        $companyData["IsRMITPlus"] = if ($clientTypes.ContainsKey($companyId)) { $clientTypes[$companyId] } else { $checkBoxRMITPlus.Checked }
                    }
                } else {
                    foreach ($companyData in $companiesToProcess) {
                        $companyData["IsRMITPlus"] = $checkBoxRMITPlus.Checked
                    }
                }
                # Only skip dialogs when 2+ companies (true bulk); single-company always shows dialogs
                $skipDialogsForBulk = $chkBulkProcessing.Checked -and $companiesData.Count -gt 1
                $processedOutputs = [System.Collections.ArrayList]::new()
                try {
                        foreach ($companyData in $companiesToProcess) {
                            [System.Windows.Forms.Application]::DoEvents()
                            $inputPath = $companyData.InputPath
                            $clientName = $companyData.ClientName
                            $scanDate = $companyData.ScanDate
                            $outputDir = if ($companyData.OutputDir) { $companyData.OutputDir } else { $textBoxOutputDir.Text }
                            $isRMITPlus = if ($companyData.IsRMITPlus -ne $null) { $companyData.IsRMITPlus } else { $script:IsRMITPlus }
                            $companyId = if ($companyData.Company -and $companyData.Company.Id -ne $null) { $companyData.Company.Id } else { if ($companyData.Id -ne $null) { $companyData.Id } else { 0 } }

            Write-Log "=== Processing client: $clientName ===" -Level Info
            Write-Log "Input File: $inputPath"
            Write-Log "Client: $clientName"
            Write-Log "Client Type: $(if ($isRMITPlus) { 'RMIT+' } else { 'RMIT/CMIT' })"
            Write-Log "Scan Date: $scanDate"

            # Reusable values for this client (avoids repeated logic)
            $companyName = if ([string]::IsNullOrWhiteSpace($clientName)) { "Client" } else { $clientName }
            $reportTimestamp = Get-Date -Format "yyyy-MM-dd_HHmmss"
            $useMiscForText = -not [string]::IsNullOrWhiteSpace($script:UserSettings.ReportsBasePath)
            $miscDir = if ($useMiscForText) { Join-Path $outputDir "Misc" } else { $outputDir }
            if ($useMiscForText -and -not (Test-Path $miscDir)) { New-Item -ItemType Directory -Path $miscDir -Force | Out-Null }
            $textOutputDir = $miscDir

            # Read vulnerability data from all remediation sheets
            Update-Progress -Status "Reading vulnerability data from Excel file..." -Show $true
            $vulnData = Get-VulnerabilityData -ExcelPath $inputPath

            # Calculate top 10 vulnerabilities
            Update-Progress -Status "Calculating top vulnerabilities..." -Show $true
            $minEPSS = $script:FilterMinEPSS
            $vulnCountSelection = $script:FilterTopN
            $vulnCount = if ($vulnCountSelection -eq "All") { 0 } else { [int]$vulnCountSelection }
            $reportTitle = if ($vulnCountSelection -eq "All") { "Top Vulnerabilities Report" } elseif ($vulnCountSelection -eq "10") { "Top Ten Vulnerabilities Report" } else { "Top $vulnCountSelection Vulnerabilities Report" }
            
            $top10 = Get-Top10Vulnerabilities -VulnData $vulnData `
                                            -MinEPSS $minEPSS `
                                            -IncludeCritical $script:FilterIncludeCritical `
                                            -IncludeHigh $script:FilterIncludeHigh `
                                            -IncludeMedium $script:FilterIncludeMedium `
                                            -IncludeLow $script:FilterIncludeLow `
                                            -Count $vulnCount

            # Store time estimates and general recommendations for use in reports
            $timeEstimates = $null
            $generalRecommendations = @()

            # Generate Excel report
            if ($script:OutputExcel) {
                Update-Progress -Status "Generating Excel Report..." -Show $true
                $excelOutputPath = Get-SafeReportOutputPath -TargetDir $outputDir -CompanyName $companyName -ReportSuffix " Vulnerability Report_$reportTimestamp" -Ext "xlsx"

                # Allow previous Excel instance (from Get-VulnerabilityData) to fully release before starting report generation
                [System.GC]::Collect()
                [System.GC]::WaitForPendingFinalizers()

                Invoke-OperationWithRetry -OperationName "Excel Report Generation" -Operation {
                    New-ExcelReport -InputPath $inputPath -OutputPath $excelOutputPath
                }

                $script:ExcelReportPath = $excelOutputPath
                Write-Log "Excel Report saved to: $excelOutputPath" -Level Success
            }

            # Generate Email Template
            if ($script:OutputEmailTemplate) {
                Update-Progress -Status "Generating Email Template..." -Show $true
                $emailOutputPath = Get-SafeReportOutputPath -TargetDir $textOutputDir -CompanyName $companyName -ReportSuffix " Email Template_$reportTimestamp" -Ext "txt"

                New-EmailTemplate -OutputPath $emailOutputPath -IsRMITPlus $isRMITPlus -FilterTopN $script:FilterTopN

                $script:EmailTemplatePath = $emailOutputPath
                Write-Log "Email Template saved to: $emailOutputPath" -Level Success
            }

            # General Recommendations, Hostname Review, Time Estimate (or skip dialogs in bulk mode)
            $skipDialogs = $skipDialogsForBulk
            $hasTop10Data = $top10 -and $top10.Count -gt 0
            if (-not $hasTop10Data -and -not $skipDialogs) {
                [System.Windows.Forms.MessageBox]::Show(
                    "No vulnerabilities matched your filters (EPSS, severity, Top N). The following dialogs will be skipped: General Recommendations, Hostname Review, Time Estimate.`n`nExcel and Email Template (if selected) have been generated. Check your filter settings or source data if you expected vulnerability items.",
                    "No Vulnerability Data",
                    [System.Windows.Forms.MessageBoxButtons]::OK,
                    [System.Windows.Forms.MessageBoxIcon]::Information
                )
            }
            if (-not $skipDialogs -and $hasTop10Data) {
                Update-Progress -Status "Entering General Recommendations..." -Show $true
                $generalRecommendations = Show-GeneralRecommendationsDialog -Top10Data $top10
                if ($null -eq $generalRecommendations) {
                    Write-Log "General Recommendations dialog cancelled by user." -Level Warning
                    $generalRecommendations = @()
                }
            } else {
                Load-GeneralRecommendations | Out-Null
                if ($null -eq $script:GeneralRecommendations) { $script:GeneralRecommendations = @() }
                $generalRecommendations = @()
                foreach ($item in $top10) {
                    $prod = if ($null -ne $item -and $null -ne $item.Product) { [string]$item.Product } else { '' }
                    $matchingRec = $null
                    foreach ($rec in @($script:GeneralRecommendations)) {
                        if ($rec -and $rec.Product -and $prod -like $rec.Product) { $matchingRec = $rec; break }
                    }
                    if ($matchingRec -and -not [string]::IsNullOrWhiteSpace($matchingRec.Recommendations)) {
                        $generalRecommendations += [PSCustomObject]@{ Product = $prod; Recommendations = $matchingRec.Recommendations }
                    }
                }
                [System.Windows.Forms.Application]::DoEvents()
            }

            if (-not $skipDialogs -and $hasTop10Data) {
                Update-Progress -Status "Reviewing Hostnames..." -Show $true
                $filteredTop10 = Show-HostnameReviewDialog -Top10Data $top10 -CompanyId $companyId
                if ($null -eq $filteredTop10) {
                    Write-Log "Hostname Review dialog cancelled by user. Using original data." -Level Warning
                    $filteredTop10 = $top10
                } else {
                    $top10 = $filteredTop10
                }
            } else {
                $filteredTop10 = $top10
            }
            
            $script:CurrentTop10Data = $top10

            if ($script:OutputTimeEstimate) {
                if (-not $skipDialogs -and $hasTop10Data) {
                    Update-Progress -Status "Generating Time Estimate..." -Show $true
                    $timeEstimates = Show-TimeEstimateEntryDialog -Top10Data $top10 -IsRMITPlus $isRMITPlus
                } elseif ($skipDialogs -and $hasTop10Data) {
                    $timeEstimates = foreach ($item in $top10) {
                        $prod = if ($null -ne $item -and $null -ne $item.Product) { [string]$item.Product } else { '' }
                        [PSCustomObject]@{
                            Product = $prod
                            TimeEstimate = 0.0
                            AfterHours = $false
                            ThirdParty = if ($isRMITPlus) { -not (Test-IsFirstPartyVendor -ProductName $prod) } else { $false }
                            TicketGenerated = $false
                        }
                    }
                    [System.Windows.Forms.Application]::DoEvents()
                } else {
                    $timeEstimates = $null
                }
                
                if ($null -ne $timeEstimates -and $timeEstimates.Count -gt 0) {
                    $timeEstimateOutputPath = Get-SafeReportOutputPath -TargetDir $textOutputDir -CompanyName $companyName -ReportSuffix " Time Estimate_$reportTimestamp" -Ext "txt"

                    New-TimeEstimate -OutputPath $timeEstimateOutputPath -Top10Data $top10 -TimeEstimates $timeEstimates -IsRMITPlus $isRMITPlus -GeneralRecommendations $generalRecommendations

                    $script:TimeEstimatePath = $timeEstimateOutputPath
                    # Store TimeEstimates in script variable for ticket notes
                    $script:CurrentTimeEstimates = $timeEstimates

                    Write-Log "Time Estimate saved to: $timeEstimateOutputPath" -Level Success
                } else {
                    Write-Log "Time Estimate generation cancelled by user." -Level Warning
                    $script:CurrentTimeEstimates = $null
                }
            } else {
                # Store empty TimeEstimates if time estimate not generated
                $script:CurrentTimeEstimates = $null
            }

            # Generate Word report (after time estimate dialog so it can reflect checkbox states)
            if ($script:OutputWord) {
                # Only generate Word report if time estimate was not requested, or if it was requested and completed successfully
                # (if time estimate was requested but cancelled, skip Word report)
                if (-not $script:OutputTimeEstimate -or $null -ne $timeEstimates) {
                    Update-Progress -Status "Generating $reportTitle (Word)..." -Show $true
                    $wordOutputPath = Get-SafeReportOutputPath -TargetDir $outputDir -CompanyName $companyName -ReportSuffix " $reportTitle _$reportTimestamp" -Ext "docx"

                    Invoke-OperationWithRetry -OperationName "Word Report Generation" -Operation {
                        New-WordReport -OutputPath $wordOutputPath `
                                      -ClientName $clientName `
                                      -ScanDate $scanDate `
                                      -Top10Data $top10 `
                                      -TimeEstimates $timeEstimates `
                                      -IsRMITPlus $isRMITPlus `
                                      -GeneralRecommendations $generalRecommendations `
                                      -ReportTitle $reportTitle
                    }

                    $script:WordReportPath = $wordOutputPath
                    Write-Log "$reportTitle saved to: $wordOutputPath" -Level Success
                } else {
                    Write-Log "Word report generation skipped because time estimate was cancelled." -Level Warning
                }
            }

            # Generate combined report (Ticket Instructions + Email Template + Time Estimate in tabs)
            if ($script:OutputTicketInstructions -or $script:OutputEmailTemplate -or $script:OutputTimeEstimate) {
                Update-Progress -Status "Generating Report (HTML)..." -Show $true
                $reportHtmlPath = Get-SafeReportOutputPath -TargetDir $textOutputDir -CompanyName $companyName -ReportSuffix " Report_$reportTimestamp" -Ext "html"

                New-CombinedReportHtml -OutputPath $reportHtmlPath -TopTenData $top10 -TimeEstimates $timeEstimates -IsRMITPlus $isRMITPlus -GeneralRecommendations $generalRecommendations `
                    -IncludeTicketInstructions $script:OutputTicketInstructions `
                    -IncludeEmailTemplate $script:OutputEmailTemplate `
                    -IncludeTimeEstimate ($script:OutputTimeEstimate -and $null -ne $timeEstimates) `
                    -FilterTopN $script:FilterTopN `
                    -CompanyName $companyName

                $script:TicketInstructionsPath = $reportHtmlPath
                $script:TicketInstructionsHtmlPath = $reportHtmlPath
                Write-Log "Report saved to: $reportHtmlPath" -Level Success
            }

            # Auto-generate ticket notes file if we have data
            if ($null -ne $script:CurrentTop10Data) {
                $ticketNotesOutputPath = Get-SafeReportOutputPath -TargetDir $textOutputDir -CompanyName $companyName -ReportSuffix " Ticket Notes_$reportTimestamp" -Ext "txt"
                
                New-TicketNotes -Top10Data $script:CurrentTop10Data -TimeEstimates $script:CurrentTimeEstimates -OutputPath $ticketNotesOutputPath -IsRMITPlus $isRMITPlus -FilterTopN $script:FilterTopN
                
                if ($script:TicketNotesPath) {
                    $script:buttonOpenTicketNotes.Enabled = $true
                }
            }

            # Hide progress bar for this client
            Update-Progress -Status "Complete" -Show $false

            Write-Log "=== Processing Complete for $clientName ===" -Level Success

            $null = $processedOutputs.Add([PSCustomObject]@{ CompanyName = $clientName; OutputPath = $outputDir })
            Add-ToReportFolderHistory -CompanyName $clientName -OutputPath $outputDir

            }  # end foreach ($companyData in $companiesToProcess)

            Write-Log "=== All Processing Complete ===" -Level Success

            if ($textBoxOutputDir.Text -and (Test-Path $textBoxOutputDir.Text)) {
                $script:UserSettings.LastOutputDirectory = $textBoxOutputDir.Text
                Save-UserSettings | Out-Null
            }

            Show-ProcessingSummaryDialog -ProcessedOutputs @($processedOutputs)

            if (-not $skipDialogsForBulk) {
                [System.Windows.Forms.MessageBox]::Show("Report generation completed successfully!", "Success",
                    [System.Windows.Forms.MessageBoxButtons]::OK, [System.Windows.Forms.MessageBoxIcon]::Information)
            }

                } catch {
                        Update-Progress -Status "Error" -Show $false
                        Write-Log "Processing failed: $($_.Exception.Message)" -Level Error
                        if (-not $skipDialogsForBulk) {
                            [System.Windows.Forms.MessageBox]::Show("An error occurred during processing. Check the log for details.", "Error",
                                [System.Windows.Forms.MessageBoxButtons]::OK, [System.Windows.Forms.MessageBoxIcon]::Error)
                        }
                } finally {
                        $buttonGenerate.Enabled = $true
                }
                return
            }

            # Sync path: $hasFile - process from input file
            $processedOutputsSync = [System.Collections.ArrayList]::new()
            foreach ($company in $companiesToProcess) {
                $clientName = ($company.DisplayName -replace '\s*\(ID:\s*\d+\)\s*$', '').Trim()
                if ([string]::IsNullOrWhiteSpace($clientName)) { $clientName = $company.DisplayName }
                $scanDate = $datePickerScanDate.Value.ToShortDateString()
                $inputPath = $company.InputPath
                $outputDir = if ($company.OutputDir) { $company.OutputDir } else { $textBoxOutputDir.Text }
                $isRMITPlus = $script:IsRMITPlus
                $companyId = if ($company.Company -and $company.Company.Id -ne $null) { $company.Company.Id } else { if ($company.Id -ne $null) { $company.Id } else { 0 } }

            Write-Log "=== Processing client: $clientName ===" -Level Info
            Write-Log "Input File: $inputPath"
            Write-Log "Client: $clientName"
            Write-Log "Client Type: $(if ($isRMITPlus) { 'RMIT+' } else { 'RMIT/CMIT' })"
            Write-Log "Scan Date: $scanDate"

            $companyName = if ([string]::IsNullOrWhiteSpace($clientName)) { "Client" } else { $clientName }
            $reportTimestamp = Get-Date -Format "yyyy-MM-dd_HHmmss"
            $useMiscForText = -not [string]::IsNullOrWhiteSpace($script:UserSettings.ReportsBasePath)
            $miscDir = if ($useMiscForText) { Join-Path $outputDir "Misc" } else { $outputDir }
            if ($useMiscForText -and -not (Test-Path $miscDir)) { New-Item -ItemType Directory -Path $miscDir -Force | Out-Null }
            $textOutputDir = $miscDir

            Update-Progress -Status "Reading vulnerability data from Excel file..." -Show $true
            $vulnData = Get-VulnerabilityData -ExcelPath $inputPath

            Update-Progress -Status "Calculating top vulnerabilities..." -Show $true
            $minEPSS = $script:FilterMinEPSS
            $vulnCountSelection = $script:FilterTopN
            $vulnCount = if ($vulnCountSelection -eq "All") { 0 } else { [int]$vulnCountSelection }
            $reportTitle = if ($vulnCountSelection -eq "All") { "Top Vulnerabilities Report" } elseif ($vulnCountSelection -eq "10") { "Top Ten Vulnerabilities Report" } else { "Top $vulnCountSelection Vulnerabilities Report" }
            
            $top10 = Get-Top10Vulnerabilities -VulnData $vulnData `
                                            -MinEPSS $minEPSS `
                                            -IncludeCritical $script:FilterIncludeCritical `
                                            -IncludeHigh $script:FilterIncludeHigh `
                                            -IncludeMedium $script:FilterIncludeMedium `
                                            -IncludeLow $script:FilterIncludeLow `
                                            -Count $vulnCount

            $timeEstimates = $null
            $generalRecommendations = @()

            if ($script:OutputExcel) {
                Update-Progress -Status "Generating Excel Report..." -Show $true
                $excelOutputPath = Get-SafeReportOutputPath -TargetDir $outputDir -CompanyName $companyName -ReportSuffix " Vulnerability Report_$reportTimestamp" -Ext "xlsx"

                [System.GC]::Collect()
                [System.GC]::WaitForPendingFinalizers()

                Invoke-OperationWithRetry -OperationName "Excel Report Generation" -Operation {
                    New-ExcelReport -InputPath $inputPath -OutputPath $excelOutputPath
                }

                $script:ExcelReportPath = $excelOutputPath
                Write-Log "Excel Report saved to: $excelOutputPath" -Level Success
            }

            if ($script:OutputEmailTemplate) {
                Update-Progress -Status "Generating Email Template..." -Show $true
                $emailOutputPath = Get-SafeReportOutputPath -TargetDir $textOutputDir -CompanyName $companyName -ReportSuffix " Email Template_$reportTimestamp" -Ext "txt"

                New-EmailTemplate -OutputPath $emailOutputPath -IsRMITPlus $isRMITPlus -FilterTopN $script:FilterTopN

                $script:EmailTemplatePath = $emailOutputPath
                Write-Log "Email Template saved to: $emailOutputPath" -Level Success
            }

            $skipDialogs = $chkBulkProcessing.Checked
            $hasTop10Data = $top10 -and $top10.Count -gt 0
            if (-not $hasTop10Data -and -not $skipDialogs) {
                [System.Windows.Forms.MessageBox]::Show(
                    "No vulnerabilities matched your filters (EPSS, severity, Top N). The following dialogs will be skipped: General Recommendations, Hostname Review, Time Estimate.`n`nExcel and Email Template (if selected) have been generated. Check your filter settings or source data if you expected vulnerability items.",
                    "No Vulnerability Data",
                    [System.Windows.Forms.MessageBoxButtons]::OK,
                    [System.Windows.Forms.MessageBoxIcon]::Information
                )
            }
            if (-not $skipDialogs -and $hasTop10Data) {
                Update-Progress -Status "Entering General Recommendations..." -Show $true
                $generalRecommendations = Show-GeneralRecommendationsDialog -Top10Data $top10
                if ($null -eq $generalRecommendations) {
                    Write-Log "General Recommendations dialog cancelled by user." -Level Warning
                    $generalRecommendations = @()
                }
            } else {
                Load-GeneralRecommendations | Out-Null
                if ($null -eq $script:GeneralRecommendations) { $script:GeneralRecommendations = @() }
                $generalRecommendations = @()
                foreach ($item in $top10) {
                    $prod = if ($null -ne $item -and $null -ne $item.Product) { [string]$item.Product } else { '' }
                    $matchingRec = $null
                    foreach ($rec in @($script:GeneralRecommendations)) {
                        if ($rec -and $rec.Product -and $prod -like $rec.Product) { $matchingRec = $rec; break }
                    }
                    if ($matchingRec -and -not [string]::IsNullOrWhiteSpace($matchingRec.Recommendations)) {
                        $generalRecommendations += [PSCustomObject]@{ Product = $prod; Recommendations = $matchingRec.Recommendations }
                    }
                }
                [System.Windows.Forms.Application]::DoEvents()
            }

            if (-not $skipDialogs -and $hasTop10Data) {
                Update-Progress -Status "Reviewing Hostnames..." -Show $true
                $filteredTop10 = Show-HostnameReviewDialog -Top10Data $top10 -CompanyId $companyId
                if ($null -eq $filteredTop10) {
                    Write-Log "Hostname Review dialog cancelled by user. Using original data." -Level Warning
                    $filteredTop10 = $top10
                } else {
                    $top10 = $filteredTop10
                }
            } else {
                $filteredTop10 = $top10
            }
            
            $script:CurrentTop10Data = $top10

            if ($script:OutputTimeEstimate) {
                if (-not $skipDialogs -and $hasTop10Data) {
                    Update-Progress -Status "Generating Time Estimate..." -Show $true
                    $timeEstimates = Show-TimeEstimateEntryDialog -Top10Data $top10 -IsRMITPlus $isRMITPlus
                } elseif ($skipDialogs -and $hasTop10Data) {
                    $timeEstimates = foreach ($item in $top10) {
                        $prod = if ($null -ne $item -and $null -ne $item.Product) { [string]$item.Product } else { '' }
                        [PSCustomObject]@{
                            Product = $prod
                            TimeEstimate = 0.0
                            AfterHours = $false
                            ThirdParty = if ($isRMITPlus) { -not (Test-IsFirstPartyVendor -ProductName $prod) } else { $false }
                            TicketGenerated = $false
                        }
                    }
                    [System.Windows.Forms.Application]::DoEvents()
                } else {
                    $timeEstimates = $null
                }
                
                if ($null -ne $timeEstimates -and $timeEstimates.Count -gt 0) {
                    $timeEstimateOutputPath = Get-SafeReportOutputPath -TargetDir $textOutputDir -CompanyName $companyName -ReportSuffix " Time Estimate_$reportTimestamp" -Ext "txt"

                    New-TimeEstimate -OutputPath $timeEstimateOutputPath -Top10Data $top10 -TimeEstimates $timeEstimates -IsRMITPlus $isRMITPlus -GeneralRecommendations $generalRecommendations

                    $script:TimeEstimatePath = $timeEstimateOutputPath
                    
                    $script:CurrentTimeEstimates = $timeEstimates

                    Write-Log "Time Estimate saved to: $timeEstimateOutputPath" -Level Success
                } else {
                    Write-Log "Time Estimate generation cancelled by user." -Level Warning
                    $script:CurrentTimeEstimates = $null
                }
            } else {
                $script:CurrentTimeEstimates = $null
            }

            if ($script:OutputWord) {
                if (-not $script:OutputTimeEstimate -or $null -ne $timeEstimates) {
                    Update-Progress -Status "Generating $reportTitle (Word)..." -Show $true
                    $wordOutputPath = Get-SafeReportOutputPath -TargetDir $outputDir -CompanyName $companyName -ReportSuffix " $reportTitle _$reportTimestamp" -Ext "docx"

                    Invoke-OperationWithRetry -OperationName "Word Report Generation" -Operation {
                        New-WordReport -OutputPath $wordOutputPath `
                                      -ClientName $clientName `
                                      -ScanDate $scanDate `
                                      -Top10Data $top10 `
                                      -TimeEstimates $timeEstimates `
                                      -IsRMITPlus $isRMITPlus `
                                      -GeneralRecommendations $generalRecommendations `
                                      -ReportTitle $reportTitle
                    }

                    $script:WordReportPath = $wordOutputPath
                    Write-Log "$reportTitle saved to: $wordOutputPath" -Level Success
                } else {
                    Write-Log "Word report generation skipped because time estimate was cancelled." -Level Warning
                }
            }

            if ($script:OutputTicketInstructions -or $script:OutputEmailTemplate -or $script:OutputTimeEstimate) {
                Update-Progress -Status "Generating Report (HTML)..." -Show $true
                $reportHtmlPath = Get-SafeReportOutputPath -TargetDir $textOutputDir -CompanyName $companyName -ReportSuffix " Report_$reportTimestamp" -Ext "html"

                New-CombinedReportHtml -OutputPath $reportHtmlPath -TopTenData $top10 -TimeEstimates $timeEstimates -IsRMITPlus $isRMITPlus -GeneralRecommendations $generalRecommendations `
                    -IncludeTicketInstructions $script:OutputTicketInstructions `
                    -IncludeEmailTemplate $script:OutputEmailTemplate `
                    -IncludeTimeEstimate ($script:OutputTimeEstimate -and $null -ne $timeEstimates) `
                    -FilterTopN $script:FilterTopN `
                    -CompanyName $companyName

                $script:TicketInstructionsPath = $reportHtmlPath
                $script:TicketInstructionsHtmlPath = $reportHtmlPath
                Write-Log "Report saved to: $reportHtmlPath" -Level Success
            }

            if ($null -ne $script:CurrentTop10Data) {
                $ticketNotesOutputPath = Get-SafeReportOutputPath -TargetDir $textOutputDir -CompanyName $companyName -ReportSuffix " Ticket Notes_$reportTimestamp" -Ext "txt"
                
                New-TicketNotes -Top10Data $script:CurrentTop10Data -TimeEstimates $script:CurrentTimeEstimates -OutputPath $ticketNotesOutputPath -IsRMITPlus $isRMITPlus -FilterTopN $script:FilterTopN
                
                if ($script:TicketNotesPath) {
                    $script:buttonOpenTicketNotes.Enabled = $true
                }
            }

            Update-Progress -Status "Complete" -Show $false

            Write-Log "=== Processing Complete for $clientName ===" -Level Success

            $null = $processedOutputsSync.Add([PSCustomObject]@{ CompanyName = $clientName; OutputPath = $outputDir })
            Add-ToReportFolderHistory -CompanyName $clientName -OutputPath $outputDir

            }  # end foreach ($company in $companiesToProcess) - sync path

            Write-Log "=== All Processing Complete ===" -Level Success

            if ($textBoxOutputDir.Text -and (Test-Path $textBoxOutputDir.Text)) {
                $script:UserSettings.LastOutputDirectory = $textBoxOutputDir.Text
                Save-UserSettings | Out-Null
            }

            Show-ProcessingSummaryDialog -ProcessedOutputs @($processedOutputsSync)

            [System.Windows.Forms.MessageBox]::Show("Report generation completed successfully!", "Success",
                [System.Windows.Forms.MessageBoxButtons]::OK, [System.Windows.Forms.MessageBoxIcon]::Information)

        } catch {
            # Hide progress bar on error
            Update-Progress -Status "Error" -Show $false

            Write-Log "Processing failed: $($_.Exception.Message)" -Level Error
            [System.Windows.Forms.MessageBox]::Show("An error occurred during processing. Check the log for details.", "Error",
                [System.Windows.Forms.MessageBoxButtons]::OK, [System.Windows.Forms.MessageBoxIcon]::Error)
        } finally {
            if (-not $isAsyncPath) { $buttonGenerate.Enabled = $true }
        }
    })

    $buttonClose = New-Object System.Windows.Forms.Button
    $buttonClose.Location = New-Object System.Drawing.Point(580, 0)
    $buttonClose.Size = New-Object System.Drawing.Size(100, 30)
    $buttonClose.Text = "Close"
    $buttonClose.BackColor = [System.Drawing.Color]::FromArgb(128, 128, 128)  # Gray - Secondary action
    $buttonClose.ForeColor = [System.Drawing.Color]::White
    $buttonClose.FlatStyle = [System.Windows.Forms.FlatStyle]::Flat
    $buttonClose.FlatAppearance.BorderSize = 0
    $buttonClose.Add_Click({ $form.Close() })
    $panelBottomButtons.Controls.Add($buttonClose)

    # Show form
    Write-Log "VScanMagic $($script:Config.Version) initialized" -Level Info
    $form.ShowDialog() | Out-Null
}
