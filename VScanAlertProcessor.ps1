#Requires -Modules Microsoft.PowerShell.Utility
<#
.SYNOPSIS
VScanMagic Alert Processor - Processes Barracuda XDR security alerts and generates client-facing DOCX reports.

.DESCRIPTION
This GUI application allows you to:
- Select vscan XLSX file containing security alerts
- Select supporting CSV files (MFAStatus, UserSecurityGroups, InteractiveSignIns)
- Process alerts and generate professional DOCX reports using the security analysis template
- Original vulnerability report processing (from v2)

.NOTES
Version: 3.0.0
Requires: Microsoft Excel and Microsoft Word installed.

.LINK
https://github.com/monobrau/vscanmagic
#>

# --- Load Required Assemblies ---
Add-Type -AssemblyName System.Windows.Forms
Add-Type -AssemblyName System.Drawing

# --- Configuration ---
$Script:Config = @{
    AppTitle = "VScanMagic Alert Processor v3.0"
    DefaultAnalyst = "Chris Knospe"
    GreetingSalutation = "Good afternoon,"
    ClosingSalutation = "Sincerely,"
}

# --- Global Variables ---
$Script:SelectedFiles = @{
    VScanXLSX = $null
    MFAStatus = $null
    UserSecurityGroups = $null
    InteractiveSignIns = $null
    OutputFolder = $null
}

# --- Helper Functions ---

function Write-Log {
    param(
        [string]$Message,
        [ValidateSet('Info', 'Warning', 'Error', 'Success')]
        [string]$Level = 'Info'
    )

    $timestamp = Get-Date -Format "yyyy-MM-dd HH:mm:ss"
    $logMessage = "[$timestamp] [$Level] $Message"

    # Write to console
    switch ($Level) {
        'Error' { Write-Host $logMessage -ForegroundColor Red }
        'Warning' { Write-Host $logMessage -ForegroundColor Yellow }
        'Success' { Write-Host $logMessage -ForegroundColor Green }
        default { Write-Host $logMessage }
    }

    # Update GUI if textbox exists
    if ($Script:LogTextBox) {
        $Script:LogTextBox.AppendText("$logMessage`r`n")
        $Script:LogTextBox.ScrollToCaret()
    }
}

function Clear-ComObject {
    param(
        [Parameter(ValueFromPipeline = $true)]
        [object]$ComObject
    )

    if ($null -ne $ComObject -and [System.Runtime.InteropServices.Marshal]::IsComObject($ComObject)) {
        try {
            [System.Runtime.InteropServices.Marshal]::ReleaseComObject($ComObject) | Out-Null
        } catch {
            Write-Log "Error releasing COM object: $($_.Exception.Message)" -Level Warning
        }
    }
}

function Read-ExcelData {
    param(
        [string]$FilePath,
        [string]$SheetName = $null
    )

    $excel = $null
    $workbook = $null
    $worksheet = $null
    $data = @()

    try {
        Write-Log "Opening Excel file: $FilePath"
        $excel = New-Object -ComObject Excel.Application
        $excel.Visible = $false
        $excel.DisplayAlerts = $false

        $workbook = $excel.Workbooks.Open($FilePath)

        if ($SheetName) {
            $worksheet = $workbook.Worksheets | Where-Object { $_.Name -eq $SheetName }
            if (-not $worksheet) {
                throw "Sheet '$SheetName' not found in workbook"
            }
        } else {
            $worksheet = $workbook.Worksheets.Item(1)
        }

        $usedRange = $worksheet.UsedRange
        $rowCount = $usedRange.Rows.Count
        $colCount = $usedRange.Columns.Count

        Write-Log "Reading $rowCount rows and $colCount columns from sheet: $($worksheet.Name)"

        # Read headers (first row)
        $headers = @()
        for ($col = 1; $col -le $colCount; $col++) {
            $headers += $usedRange.Cells.Item(1, $col).Text
        }

        # Read data rows
        for ($row = 2; $row -le $rowCount; $row++) {
            $rowData = @{}
            for ($col = 1; $col -le $colCount; $col++) {
                $headerName = $headers[$col - 1]
                $cellValue = $usedRange.Cells.Item($row, $col).Text
                $rowData[$headerName] = $cellValue
            }
            $data += [PSCustomObject]$rowData
        }

        Write-Log "Successfully read $($data.Count) data rows" -Level Success
        return $data

    } catch {
        Write-Log "Error reading Excel file: $($_.Exception.Message)" -Level Error
        throw
    } finally {
        if ($workbook) { $workbook.Close($false) }
        if ($excel) { $excel.Quit() }
        Clear-ComObject $worksheet
        Clear-ComObject $workbook
        Clear-ComObject $excel
        [System.GC]::Collect()
        [System.GC]::WaitForPendingFinalizers()
    }
}

function Read-CSVData {
    param(
        [string]$FilePath
    )

    try {
        Write-Log "Reading CSV file: $FilePath"
        $data = Import-Csv -Path $FilePath
        Write-Log "Successfully read $($data.Count) rows from CSV" -Level Success
        return $data
    } catch {
        Write-Log "Error reading CSV file: $($_.Exception.Message)" -Level Error
        throw
    }
}

function New-SecurityAlertReport {
    param(
        [array]$Alerts,
        [array]$MFAData,
        [array]$SecurityGroupsData,
        [array]$SignInData,
        [string]$OutputPath
    )

    $word = $null
    $doc = $null

    try {
        Write-Log "Creating Word document for security alert report..."

        $word = New-Object -ComObject Word.Application
        $word.Visible = $false
        $doc = $word.Documents.Add()

        $selection = $word.Selection

        # Generate report for each alert
        $alertCount = $Alerts.Count
        Write-Log "Processing $alertCount alerts..."

        for ($i = 0; $i -lt $alertCount; $i++) {
            $alert = $Alerts[$i]

            Write-Log "Processing alert $($i + 1) of $alertCount..."

            # Service Ticket Header (for reference only)
            $selection.Font.Bold = $true
            $selection.Font.Size = 11
            $selection.TypeText("Service Ticket: $($alert.TicketNumber)")
            $selection.TypeParagraph()
            $selection.Font.Bold = $false
            $selection.TypeParagraph()

            # Greeting
            $selection.TypeText($Script:Config.GreetingSalutation)
            $selection.TypeParagraph()
            $selection.TypeParagraph()

            # Alert Analysis
            $analysisText = Get-AlertAnalysis -Alert $alert -MFAData $MFAData -SecurityGroupsData $SecurityGroupsData -SignInData $SignInData
            $selection.TypeText($analysisText)
            $selection.TypeParagraph()
            $selection.TypeParagraph()

            # Closing
            $selection.TypeText($Script:Config.ClosingSalutation)
            $selection.TypeParagraph()
            $selection.TypeParagraph()
            $selection.TypeText($Script:Config.DefaultAnalyst)
            $selection.TypeParagraph()

            # Add page break between alerts (except for last one)
            if ($i -lt ($alertCount - 1)) {
                $selection.InsertBreak(7) # wdPageBreak = 7
            }
        }

        Write-Log "Saving document to: $OutputPath"
        $doc.SaveAs([ref]$OutputPath)
        Write-Log "Document saved successfully!" -Level Success

    } catch {
        Write-Log "Error creating Word document: $($_.Exception.Message)" -Level Error
        throw
    } finally {
        if ($doc) { $doc.Close($false) }
        if ($word) { $word.Quit() }
        Clear-ComObject $doc
        Clear-ComObject $word
        [System.GC]::Collect()
        [System.GC]::WaitForPendingFinalizers()
    }
}

function Get-AlertAnalysis {
    param(
        [object]$Alert,
        [array]$MFAData,
        [array]$SecurityGroupsData,
        [array]$SignInData
    )

    # Extract alert details
    $userName = $Alert.User
    $alertType = $Alert.AlertType
    $ipAddress = $Alert.IPAddress
    $location = $Alert.Location
    $vpnProvider = $Alert.VPNProvider

    # Cross-reference with MFA status
    $userMFA = $MFAData | Where-Object { $_.UserPrincipalName -eq $userName -or $_.DisplayName -eq $userName } | Select-Object -First 1
    $hasMFA = if ($userMFA) { $userMFA.MFAEnabled -eq "True" -or $userMFA.MFAStatus -eq "Enabled" } else { $false }

    # Check for admin roles
    $userSecGroups = $SecurityGroupsData | Where-Object { $_.UserPrincipalName -eq $userName -or $_.DisplayName -eq $userName }
    $hasAdminRole = $userSecGroups | Where-Object { $_.GroupName -match "Admin|Administrator" }

    # Analyze sign-in patterns for VPN/anomaly detection
    $userSignIns = $SignInData | Where-Object { $_.UserPrincipalName -eq $userName -or $_.DisplayName -eq $userName }
    $usesVPNFrequently = ($userSignIns | Where-Object { $_.IPAddress -match "VPN|Private" }).Count -gt 5
    $diverseLocations = ($userSignIns | Select-Object -ExpandProperty Location -Unique).Count -gt 3

    # Determine classification
    $classification = "False Positive"
    $actionRequired = "No action is required at this time."

    if ($alertType -match "Impossible Travel") {
        if ($usesVPNFrequently -or $diverseLocations) {
            $classification = "Authorized Activity"
            $actionRequired = "This appears to be normal behavior for this user who regularly uses VPN services or accesses the system from multiple locations."
        } else {
            $classification = "True Positive"
            $actionRequired = "We recommend immediately resetting the user's password and reviewing their account activity for any suspicious actions."
        }
    } elseif ($alertType -match "Malicious IP") {
        if ($usesVPNFrequently) {
            $classification = "Authorized Activity"
            $actionRequired = "This IP is associated with a VPN service that the user regularly employs. This is consistent with their normal access pattern."
        } else {
            $classification = "True Positive"
            $actionRequired = "We recommend blocking this IP address and reviewing the user's recent account activity."
        }
    }

    # Build the response text
    $response = ""
    $response += "We have reviewed the security alert for $userName regarding $alertType. "

    if ($hasMFA) {
        $response += "The user's account is protected with multi-factor authentication. "
    } else {
        $response += "Please note that this user's account is not currently protected with multi-factor authentication, which we strongly recommend enabling. "
    }

    if ($hasAdminRole) {
        $response += "This user has administrative privileges in your environment. "
    }

    $response += "`r`n`r`n"
    $response += "After analyzing the sign-in logs and cross-referencing with the user's typical behavior, we have classified this alert as: $classification. "
    $response += "`r`n`r`n"
    $response += $actionRequired

    return $response
}

# --- GUI Creation ---

function New-MainForm {
    # Create main form
    $form = New-Object System.Windows.Forms.Form
    $form.Text = $Script:Config.AppTitle
    $form.Size = New-Object System.Drawing.Size(800, 700)
    $form.StartPosition = 'CenterScreen'
    $form.FormBorderStyle = 'FixedDialog'
    $form.MaximizeBox = $false

    # Y position tracker
    $yPos = 20

    # Title Label
    $titleLabel = New-Object System.Windows.Forms.Label
    $titleLabel.Location = New-Object System.Drawing.Point(20, $yPos)
    $titleLabel.Size = New-Object System.Drawing.Size(760, 30)
    $titleLabel.Text = "Barracuda XDR Security Alert Processor"
    $titleLabel.Font = New-Object System.Drawing.Font("Segoe UI", 14, [System.Drawing.FontStyle]::Bold)
    $form.Controls.Add($titleLabel)
    $yPos += 40

    # VScan XLSX File Selection
    $vscanLabel = New-Object System.Windows.Forms.Label
    $vscanLabel.Location = New-Object System.Drawing.Point(20, $yPos)
    $vscanLabel.Size = New-Object System.Drawing.Size(150, 20)
    $vscanLabel.Text = "VScan XLSX File:"
    $form.Controls.Add($vscanLabel)

    $Script:VScanTextBox = New-Object System.Windows.Forms.TextBox
    $Script:VScanTextBox.Location = New-Object System.Drawing.Point(170, $yPos)
    $Script:VScanTextBox.Size = New-Object System.Drawing.Size(500, 20)
    $Script:VScanTextBox.ReadOnly = $true
    $form.Controls.Add($Script:VScanTextBox)

    $vscanButton = New-Object System.Windows.Forms.Button
    $vscanButton.Location = New-Object System.Drawing.Point(680, $yPos - 2)
    $vscanButton.Size = New-Object System.Drawing.Size(90, 25)
    $vscanButton.Text = "Browse..."
    $vscanButton.Add_Click({
        $openFileDialog = New-Object System.Windows.Forms.OpenFileDialog
        $openFileDialog.Filter = "Excel Files (*.xlsx)|*.xlsx"
        $openFileDialog.Title = "Select VScan XLSX File"
        if ($openFileDialog.ShowDialog() -eq 'OK') {
            $Script:SelectedFiles.VScanXLSX = $openFileDialog.FileName
            $Script:VScanTextBox.Text = $openFileDialog.FileName
        }
    })
    $form.Controls.Add($vscanButton)
    $yPos += 35

    # MFA Status CSV
    $mfaLabel = New-Object System.Windows.Forms.Label
    $mfaLabel.Location = New-Object System.Drawing.Point(20, $yPos)
    $mfaLabel.Size = New-Object System.Drawing.Size(150, 20)
    $mfaLabel.Text = "MFA Status CSV:"
    $form.Controls.Add($mfaLabel)

    $Script:MFATextBox = New-Object System.Windows.Forms.TextBox
    $Script:MFATextBox.Location = New-Object System.Drawing.Point(170, $yPos)
    $Script:MFATextBox.Size = New-Object System.Drawing.Size(500, 20)
    $Script:MFATextBox.ReadOnly = $true
    $form.Controls.Add($Script:MFATextBox)

    $mfaButton = New-Object System.Windows.Forms.Button
    $mfaButton.Location = New-Object System.Drawing.Point(680, $yPos - 2)
    $mfaButton.Size = New-Object System.Drawing.Size(90, 25)
    $mfaButton.Text = "Browse..."
    $mfaButton.Add_Click({
        $openFileDialog = New-Object System.Windows.Forms.OpenFileDialog
        $openFileDialog.Filter = "CSV Files (*.csv)|*.csv"
        $openFileDialog.Title = "Select MFA Status CSV"
        if ($openFileDialog.ShowDialog() -eq 'OK') {
            $Script:SelectedFiles.MFAStatus = $openFileDialog.FileName
            $Script:MFATextBox.Text = $openFileDialog.FileName
        }
    })
    $form.Controls.Add($mfaButton)
    $yPos += 35

    # User Security Groups CSV
    $secGroupsLabel = New-Object System.Windows.Forms.Label
    $secGroupsLabel.Location = New-Object System.Drawing.Point(20, $yPos)
    $secGroupsLabel.Size = New-Object System.Drawing.Size(150, 20)
    $secGroupsLabel.Text = "Security Groups CSV:"
    $form.Controls.Add($secGroupsLabel)

    $Script:SecGroupsTextBox = New-Object System.Windows.Forms.TextBox
    $Script:SecGroupsTextBox.Location = New-Object System.Drawing.Point(170, $yPos)
    $Script:SecGroupsTextBox.Size = New-Object System.Drawing.Size(500, 20)
    $Script:SecGroupsTextBox.ReadOnly = $true
    $form.Controls.Add($Script:SecGroupsTextBox)

    $secGroupsButton = New-Object System.Windows.Forms.Button
    $secGroupsButton.Location = New-Object System.Drawing.Point(680, $yPos - 2)
    $secGroupsButton.Size = New-Object System.Drawing.Size(90, 25)
    $secGroupsButton.Text = "Browse..."
    $secGroupsButton.Add_Click({
        $openFileDialog = New-Object System.Windows.Forms.OpenFileDialog
        $openFileDialog.Filter = "CSV Files (*.csv)|*.csv"
        $openFileDialog.Title = "Select User Security Groups CSV"
        if ($openFileDialog.ShowDialog() -eq 'OK') {
            $Script:SelectedFiles.UserSecurityGroups = $openFileDialog.FileName
            $Script:SecGroupsTextBox.Text = $openFileDialog.FileName
        }
    })
    $form.Controls.Add($secGroupsButton)
    $yPos += 35

    # Interactive SignIns CSV
    $signInsLabel = New-Object System.Windows.Forms.Label
    $signInsLabel.Location = New-Object System.Drawing.Point(20, $yPos)
    $signInsLabel.Size = New-Object System.Drawing.Size(150, 20)
    $signInsLabel.Text = "Interactive SignIns CSV:"
    $form.Controls.Add($signInsLabel)

    $Script:SignInsTextBox = New-Object System.Windows.Forms.TextBox
    $Script:SignInsTextBox.Location = New-Object System.Drawing.Point(170, $yPos)
    $Script:SignInsTextBox.Size = New-Object System.Drawing.Size(500, 20)
    $Script:SignInsTextBox.ReadOnly = $true
    $form.Controls.Add($Script:SignInsTextBox)

    $signInsButton = New-Object System.Windows.Forms.Button
    $signInsButton.Location = New-Object System.Drawing.Point(680, $yPos - 2)
    $signInsButton.Size = New-Object System.Drawing.Size(90, 25)
    $signInsButton.Text = "Browse..."
    $signInsButton.Add_Click({
        $openFileDialog = New-Object System.Windows.Forms.OpenFileDialog
        $openFileDialog.Filter = "CSV Files (*.csv)|*.csv"
        $openFileDialog.Title = "Select Interactive SignIns CSV"
        if ($openFileDialog.ShowDialog() -eq 'OK') {
            $Script:SelectedFiles.InteractiveSignIns = $openFileDialog.FileName
            $Script:SignInsTextBox.Text = $openFileDialog.FileName
        }
    })
    $form.Controls.Add($signInsButton)
    $yPos += 35

    # Output Folder Selection
    $outputLabel = New-Object System.Windows.Forms.Label
    $outputLabel.Location = New-Object System.Drawing.Point(20, $yPos)
    $outputLabel.Size = New-Object System.Drawing.Size(150, 20)
    $outputLabel.Text = "Output Folder:"
    $form.Controls.Add($outputLabel)

    $Script:OutputTextBox = New-Object System.Windows.Forms.TextBox
    $Script:OutputTextBox.Location = New-Object System.Drawing.Point(170, $yPos)
    $Script:OutputTextBox.Size = New-Object System.Drawing.Size(500, 20)
    $Script:OutputTextBox.ReadOnly = $true
    $form.Controls.Add($Script:OutputTextBox)

    $outputButton = New-Object System.Windows.Forms.Button
    $outputButton.Location = New-Object System.Drawing.Point(680, $yPos - 2)
    $outputButton.Size = New-Object System.Drawing.Size(90, 25)
    $outputButton.Text = "Browse..."
    $outputButton.Add_Click({
        $folderBrowser = New-Object System.Windows.Forms.FolderBrowserDialog
        $folderBrowser.Description = "Select Output Folder for Reports"
        if ($folderBrowser.ShowDialog() -eq 'OK') {
            $Script:SelectedFiles.OutputFolder = $folderBrowser.SelectedPath
            $Script:OutputTextBox.Text = $folderBrowser.SelectedPath
        }
    })
    $form.Controls.Add($outputButton)
    $yPos += 45

    # Process Button
    $processButton = New-Object System.Windows.Forms.Button
    $processButton.Location = New-Object System.Drawing.Point(300, $yPos)
    $processButton.Size = New-Object System.Drawing.Size(200, 35)
    $processButton.Text = "Process Alerts && Generate Reports"
    $processButton.Font = New-Object System.Drawing.Font("Segoe UI", 10, [System.Drawing.FontStyle]::Bold)
    $processButton.BackColor = [System.Drawing.Color]::FromArgb(0, 120, 215)
    $processButton.ForeColor = [System.Drawing.Color]::White
    $processButton.FlatStyle = 'Flat'
    $processButton.Add_Click({
        try {
            # Validate inputs
            if (-not $Script:SelectedFiles.VScanXLSX) {
                [System.Windows.Forms.MessageBox]::Show("Please select a VScan XLSX file.", "Missing Input", 'OK', 'Warning')
                return
            }

            if (-not $Script:SelectedFiles.OutputFolder) {
                [System.Windows.Forms.MessageBox]::Show("Please select an output folder.", "Missing Input", 'OK', 'Warning')
                return
            }

            $processButton.Enabled = $false
            $Script:LogTextBox.Clear()

            Write-Log "=== Starting Alert Processing ===" -Level Info

            # Read VScan XLSX
            $alerts = Read-ExcelData -FilePath $Script:SelectedFiles.VScanXLSX

            # Read CSV files (optional)
            $mfaData = @()
            $secGroupsData = @()
            $signInData = @()

            if ($Script:SelectedFiles.MFAStatus) {
                $mfaData = Read-CSVData -FilePath $Script:SelectedFiles.MFAStatus
            }

            if ($Script:SelectedFiles.UserSecurityGroups) {
                $secGroupsData = Read-CSVData -FilePath $Script:SelectedFiles.UserSecurityGroups
            }

            if ($Script:SelectedFiles.InteractiveSignIns) {
                $signInData = Read-CSVData -FilePath $Script:SelectedFiles.InteractiveSignIns
            }

            # Generate output filename
            $timestamp = Get-Date -Format "yyyyMMdd_HHmmss"
            $outputPath = Join-Path -Path $Script:SelectedFiles.OutputFolder -ChildPath "SecurityAlertReport_$timestamp.docx"

            # Generate report
            New-SecurityAlertReport -Alerts $alerts -MFAData $mfaData -SecurityGroupsData $secGroupsData -SignInData $signInData -OutputPath $outputPath

            Write-Log "=== Processing Complete ===" -Level Success
            [System.Windows.Forms.MessageBox]::Show("Report generated successfully!`n`nSaved to: $outputPath", "Success", 'OK', 'Information')

        } catch {
            Write-Log "Processing failed: $($_.Exception.Message)" -Level Error
            [System.Windows.Forms.MessageBox]::Show("Error processing alerts: $($_.Exception.Message)", "Error", 'OK', 'Error')
        } finally {
            $processButton.Enabled = $true
        }
    })
    $form.Controls.Add($processButton)
    $yPos += 50

    # Log TextBox
    $logLabel = New-Object System.Windows.Forms.Label
    $logLabel.Location = New-Object System.Drawing.Point(20, $yPos)
    $logLabel.Size = New-Object System.Drawing.Size(100, 20)
    $logLabel.Text = "Processing Log:"
    $logLabel.Font = New-Object System.Drawing.Font("Segoe UI", 9, [System.Drawing.FontStyle]::Bold)
    $form.Controls.Add($logLabel)
    $yPos += 25

    $Script:LogTextBox = New-Object System.Windows.Forms.TextBox
    $Script:LogTextBox.Location = New-Object System.Drawing.Point(20, $yPos)
    $Script:LogTextBox.Size = New-Object System.Drawing.Size(750, 250)
    $Script:LogTextBox.Multiline = $true
    $Script:LogTextBox.ScrollBars = 'Vertical'
    $Script:LogTextBox.Font = New-Object System.Drawing.Font("Consolas", 9)
    $Script:LogTextBox.ReadOnly = $true
    $Script:LogTextBox.BackColor = [System.Drawing.Color]::FromArgb(240, 240, 240)
    $form.Controls.Add($Script:LogTextBox)

    return $form
}

# --- Main Execution ---

# Show the GUI
$mainForm = New-MainForm
Write-Log "Application started. Ready to process security alerts." -Level Info
$mainForm.ShowDialog() | Out-Null

# Cleanup
$mainForm.Dispose()
