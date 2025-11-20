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
Version: 3.1.2
Requires: Microsoft Excel and Microsoft Word installed.
Author: River Run MSP
#>

# --- Add Required Assemblies ---
Add-Type -AssemblyName System.Windows.Forms
Add-Type -AssemblyName System.Drawing

# --- Configuration ---
$script:Config = @{
    AppName = "VScanMagic v3"
    Version = "3.1.2"
    Author = "River Run MSP"

    # Risk Score Calculation
    CVSSEquivalent = @{
        Critical = 9.0
        High = 7.0
        Medium = 5.0
        Low = 3.0
    }

    # Heatmap Color Thresholds for Risk Scores (Yellow to Red gradient - no greens)
    # Risk Score = EPSS × Average CVSS (max theoretical: 1.0 × 10.0 = 10.0)
    RiskColors = [ordered]@{
        Critical   = @{ Threshold = 8.0;  Color = 'DC143C'; Name = 'Critical'; TextColor = 'FFFFFF' }      # Crimson red
        VeryHigh   = @{ Threshold = 6.0;  Color = 'FF4500'; Name = 'Very High'; TextColor = 'FFFFFF' }     # Orange-red
        High       = @{ Threshold = 4.0;  Color = 'FF8C00'; Name = 'High'; TextColor = 'FFFFFF' }          # Dark orange
        MediumHigh = @{ Threshold = 2.0;  Color = 'FFA500'; Name = 'Medium-High'; TextColor = '000000' }   # Orange
        Medium     = @{ Threshold = 0;    Color = 'FFFF00'; Name = 'Medium'; TextColor = '000000' }        # Yellow (baseline)
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

# --- User Settings Persistence ---
# Handle both script and EXE execution
if ([string]::IsNullOrEmpty($PSScriptRoot)) {
    # Running as EXE - use executable directory
    try {
        $exePath = [System.Diagnostics.Process]::GetCurrentProcess().MainModule.FileName
        $exeDir = [System.IO.Path]::GetDirectoryName($exePath)
        if ([string]::IsNullOrEmpty($exeDir)) {
            # Fallback to current directory
            $exeDir = Get-Location | Select-Object -ExpandProperty Path
        }
        $script:SettingsPath = Join-Path $exeDir "VScanMagic_Settings.json"
    } catch {
        # Fallback to current directory
        $script:SettingsPath = Join-Path (Get-Location | Select-Object -ExpandProperty Path) "VScanMagic_Settings.json"
    }
} else {
    # Running as script
    $script:SettingsPath = Join-Path $PSScriptRoot "VScanMagic_Settings.json"
}
$script:UserSettings = @{
    PreparedBy = "River Run MSP"
    CompanyName = ""
    CompanyAddress = ""
    Email = ""
    PhoneNumber = ""
    CompanyPhoneNumber = ""
}

function Load-UserSettings {
    if (-not [string]::IsNullOrEmpty($script:SettingsPath) -and (Test-Path $script:SettingsPath)) {
        try {
            $json = Get-Content $script:SettingsPath -Raw | ConvertFrom-Json
            $script:UserSettings.PreparedBy = $json.PreparedBy
            $script:UserSettings.CompanyName = $json.CompanyName
            $script:UserSettings.CompanyAddress = $json.CompanyAddress
            $script:UserSettings.Email = $json.Email
            $script:UserSettings.PhoneNumber = $json.PhoneNumber
            $script:UserSettings.CompanyPhoneNumber = $json.CompanyPhoneNumber
            Write-Host "User settings loaded from $script:SettingsPath"
        } catch {
            Write-Warning "Could not load settings: $($_.Exception.Message)"
        }
    }
}

function Save-UserSettings {
    if ([string]::IsNullOrEmpty($script:SettingsPath)) {
        Write-Warning "Settings path is not set. Cannot save settings."
        return $false
    }
    try {
        $script:UserSettings | ConvertTo-Json | Set-Content $script:SettingsPath -Encoding UTF8
        Write-Host "User settings saved to $script:SettingsPath"
        return $true
    } catch {
        Write-Warning "Could not save settings: $($_.Exception.Message)"
        return $false
    }
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

function Update-Progress {
    param(
        [Parameter(Mandatory=$true)]
        [string]$Status,
        [Parameter(Mandatory=$false)]
        [bool]$Show = $true
    )

    if ($script:ProgressBar) {
        if ($Show) {
            $script:ProgressBar.Visible = $true
            $script:ProgressBar.Style = 'Marquee'
            $script:ProgressBar.MarqueeAnimationSpeed = 30
        } else {
            $script:ProgressBar.Visible = $false
        }
    }

    if ($script:StatusLabel) {
        if ($Show) {
            $script:StatusLabel.Text = $Status
            $script:StatusLabel.Visible = $true
        } else {
            $script:StatusLabel.Visible = $false
        }
    }

    # Force GUI update
    [System.Windows.Forms.Application]::DoEvents()
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

function Invoke-OperationWithRetry {
    param(
        [Parameter(Mandatory=$true)]
        [scriptblock]$Operation,
        [Parameter(Mandatory=$false)]
        [string]$OperationName = "Operation",
        [Parameter(Mandatory=$false)]
        [int]$MaxRetries = 3,
        [Parameter(Mandatory=$false)]
        [int]$DelaySeconds = 2
    )

    for ($i = 1; $i -le $MaxRetries; $i++) {
        try {
            Write-Log "Attempting $OperationName (attempt $i of $MaxRetries)..." -Level Info
            $result = & $Operation
            Write-Log "$OperationName completed successfully" -Level Success
            return $result
        } catch {
            $errorMessage = $_.Exception.Message

            if ($i -eq $MaxRetries) {
                Write-Log "$OperationName failed after $MaxRetries attempts: $errorMessage" -Level Error
                throw
            } else {
                Write-Log "$OperationName failed (attempt $i): $errorMessage. Retrying in $DelaySeconds seconds..." -Level Warning
                Start-Sleep -Seconds $DelaySeconds
            }
        }
    }
}

function Test-FileLocked {
    param([string]$FilePath)

    if (-not (Test-Path $FilePath)) {
        return $false
    }

    try {
        # Try to open with delete access (less restrictive than ReadWrite)
        # This will only fail if the file is actually locked by another process
        $fileStream = [System.IO.File]::Open($FilePath, [System.IO.FileMode]::Open, [System.IO.FileAccess]::Read, [System.IO.FileShare]::None)
        $fileStream.Close()
        $fileStream.Dispose()
        return $false
    } catch [System.IO.IOException] {
        # Check if it's specifically a sharing violation
        if ($_.Exception.Message -like "*being used by another process*" -or
            $_.Exception.Message -like "*locked*" -or
            $_.Exception.HResult -eq 0x80070020) {
            return $true
        }
        # Other IO errors (permissions, etc.) - not a lock issue
        return $false
    } catch {
        # Other exceptions (permissions, etc.) - not a lock issue
        return $false
    }
}

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

    $usedRange = $Worksheet.UsedRange
    $rowCount = $usedRange.Rows.Count

    if ($rowCount -le 1) {
        return @()
    }

    Write-Log "  Reading $rowCount rows into memory (bulk read)..."

    # PERFORMANCE OPTIMIZATION: Read entire range into memory with single COM call
    # This is 10-100x faster than reading cells individually
    $rangeValues = $usedRange.Value2

    if ($null -eq $rangeValues) {
        return @()
    }

    Write-Log "  Processing data in memory..."

    # Use ArrayList for better performance than array append
    $data = [System.Collections.ArrayList]::new()

    # Process rows in memory (no COM calls)
    for ($row = 2; $row -le $rowCount; $row++) {
        # Show progress for large datasets
        if ($row % 500 -eq 0) {
            Write-Log "  Processed $row of $rowCount rows..."
        }

        # Get values from 2D array (row, column) - fast, no COM calls
        $product = ''
        if ($columnIndices.ContainsKey('Product')) {
            $product = [string]$rangeValues[$row, $columnIndices['Product']]
        }

        # Skip rows with no product name
        if ([string]::IsNullOrWhiteSpace($product)) {
            continue
        }

        # Build row data from in-memory array
        $hostName = ''
        if ($columnIndices.ContainsKey('HostName')) {
            $hostName = [string]$rangeValues[$row, $columnIndices['HostName']]
        }

        $ip = ''
        if ($columnIndices.ContainsKey('IP')) {
            $ip = [string]$rangeValues[$row, $columnIndices['IP']]
        }

        $username = ''
        if ($columnIndices.ContainsKey('Username')) {
            $username = [string]$rangeValues[$row, $columnIndices['Username']]
        }

        $critical = 0
        if ($columnIndices.ContainsKey('Critical')) {
            $critical = Get-SafeNumericValue -Value ([string]$rangeValues[$row, $columnIndices['Critical']])
        }

        $high = 0
        if ($columnIndices.ContainsKey('High')) {
            $high = Get-SafeNumericValue -Value ([string]$rangeValues[$row, $columnIndices['High']])
        }

        $medium = 0
        if ($columnIndices.ContainsKey('Medium')) {
            $medium = Get-SafeNumericValue -Value ([string]$rangeValues[$row, $columnIndices['Medium']])
        }

        $low = 0
        if ($columnIndices.ContainsKey('Low')) {
            $low = Get-SafeNumericValue -Value ([string]$rangeValues[$row, $columnIndices['Low']])
        }

        $vulnCount = 0
        if ($columnIndices.ContainsKey('VulnCount')) {
            $vulnCount = Get-SafeNumericValue -Value ([string]$rangeValues[$row, $columnIndices['VulnCount']])
        } else {
            # Calculate from severity counts if not provided
            $vulnCount = $critical + $high + $medium + $low
        }

        $epssScore = 0.0
        if ($columnIndices.ContainsKey('EPSS')) {
            $epssScore = Get-SafeDoubleValue -Value ([string]$rangeValues[$row, $columnIndices['EPSS']])
        }

        # Only add rows that have at least one vulnerability
        if ($vulnCount -gt 0) {
            $null = $data.Add([PSCustomObject]@{
                'Host Name' = $hostName
                'IP' = $ip
                'Username' = $username
                'Product' = $product
                'Critical' = $critical
                'High' = $high
                'Medium' = $medium
                'Low' = $low
                'Vulnerability Count' = $vulnCount
                'EPSS Score' = $epssScore
            })
        }
    }

    Write-Log "  Completed processing $($data.Count) vulnerability records"
    return $data.ToArray()
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
            'Username' = @('Username', 'User Name', 'User', 'Account', 'Login', 'Login Name')
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

function Test-IsMicrosoftApplication {
    param([string]$ProductName)

    if ([string]::IsNullOrWhiteSpace($ProductName)) {
        return $false
    }

    # List of Microsoft application patterns (not OS components)
    $microsoftAppPatterns = @(
        'Microsoft Office',
        'Microsoft 365',
        'Microsoft Teams',
        'Microsoft Edge',
        'Microsoft OneDrive',
        'Microsoft Outlook',
        'Microsoft Word',
        'Microsoft Excel',
        'Microsoft PowerPoint',
        'Microsoft Access',
        'Microsoft Publisher',
        'Microsoft Visio',
        'Microsoft Project',
        'Microsoft SharePoint',
        'Skype for Business',
        'Microsoft SQL Server Management Studio',
        'Microsoft Visual Studio Code',
        'Microsoft .NET Framework',
        'Microsoft .NET Core',
        'Microsoft .NET Runtime'
    )

    foreach ($pattern in $microsoftAppPatterns) {
        if ($ProductName -like "*$pattern*") {
            return $true
        }
    }

    return $false
}

function Test-IsVMwareProduct {
    param([string]$ProductName)

    if ([string]::IsNullOrWhiteSpace($ProductName)) {
        return $false
    }

    # List of VMware product patterns
    $vmwarePatterns = @(
        'VMware',
        'VMWare',
        'vSphere',
        'vCenter',
        'ESXi',
        'VMware Tools',
        'VMware Workstation',
        'VMware Player',
        'VMware Horizon',
        'vRealize',
        'vCloud',
        'NSX'
    )

    foreach ($pattern in $vmwarePatterns) {
        if ($ProductName -like "*$pattern*") {
            return $true
        }
    }

    return $false
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
        [int]$VulnCount,  # Not used in calculation, kept for compatibility
        [double]$EPSSScore,
        [double]$AvgCVSS
    )

    # Simplified risk score: EPSS × Average CVSS
    # Max theoretical score: 1.0 × 10.0 = 10.0
    return [Math]::Round($EPSSScore * $AvgCVSS, 2)
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

            # Add affected systems (store objects with hostname, IP, and username)
            foreach ($item in $group.Group) {
                $existing.AffectedSystems += [PSCustomObject]@{
                    HostName = $item.'Host Name'
                    IP = $item.'IP'
                    Username = $item.'Username'
                }
            }
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

            # Create affected systems array with hostname, IP, and username
            $affectedSystems = @()
            foreach ($item in $group.Group) {
                $affectedSystems += [PSCustomObject]@{
                    HostName = $item.'Host Name'
                    IP = $item.'IP'
                    Username = $item.'Username'
                }
            }

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
                AffectedSystems = $affectedSystems
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
            $doc.BuiltInDocumentProperties.Item("Title").Value = "Top Ten Vulnerabilities Report - $ClientName"
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
        $selection.TypeText("Top Ten Vulnerabilities Report")
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
        $selection.TypeText("The organization utilizes ConnectWise Automate for patch management. WSUS is not currently in use.")
        $selection.TypeParagraph()

        $selection.TypeText("This report identifies the top 10 security risks based on a composite risk score that considers vulnerability count, ")
        $selection.TypeText("EPSS (Exploit Prediction Scoring System) scores, and CVSS severity ratings. ")
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
        $selection.TypeText("Risk Score = EPSS Score x Average CVSS")
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
            $chart.ChartTitle.Text = "Top 10 Vulnerabilities by Count"
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

        $rank = 1
        foreach ($item in $Top10Data) {
            # Vulnerability title with special notes
            $selection.Style = "Heading 2"
            $title = "$rank. $($item.Product)"

            # Add End of Life note for Windows 10
            if ($item.Product -like "*Windows 10*") {
                $title += " - Windows 10 is End of Life"
            }

            # Add RMIT+ note for Microsoft applications (not OS) - only for RMIT+ clients
            $isMicrosoftApp = Test-IsMicrosoftApplication -ProductName $item.Product
            if ($isMicrosoftApp -and $script:IsRMITPlus) {
                $title += " - RMIT+ ticketed"
            }

            # Add after-hours ticket note for VMware products - only for RMIT+ clients
            $isVMwareProduct = Test-IsVMwareProduct -ProductName $item.Product
            if ($isVMwareProduct -and $script:IsRMITPlus) {
                $title += " - RMIT+ after-hours ticket created if we maintain this"
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
            # Format as "hostname (username)" if username is present, otherwise just "hostname"
            $selection.ParagraphFormat.LeftIndent = 36
            $systemsList = ($item.AffectedSystems | ForEach-Object {
                $hostname = $_.HostName
                $username = $_.Username
                if (-not [string]::IsNullOrWhiteSpace($username)) {
                    "$hostname ($username)"
                } else {
                    $hostname
                }
            } | Select-Object -Unique) -join ", "
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
            } elseif ($item.Product -like "*Windows 10*") {
                $selection.TypeText("Windows 10 reached End of Life on October 14, 2025, and is no longer supported by Microsoft unless you have extended support licensing. ")
                $selection.TypeText("If Windows Updates are functional and no extension licensing is in place, there is nothing further to be done other than considering an upgrade to Windows 11 or retiring the machine. ")
                $selection.TypeText("For systems with extension licensing, continue to verify Windows Update status through ConnectWise Automate.")
            } elseif ($item.Product -like "*Windows*") {
                $selection.TypeText("Windows patch inconsistencies should be investigated via ConnectWise Automate. ")
                $selection.TypeText("Systems with lower vulnerability counts may indicate that patching is working correctly and awaiting the latest patch cycles. ")
                $selection.TypeText("For systems with high vulnerability counts, verify Windows Update status and investigate any potential issues preventing patch installation.")
            } elseif ($item.Product -like "*printer*" -or $item.Product -like "*Ripple20*") {
                $selection.TypeText("Network printers and IoT devices require manual firmware updates via manufacturer-provided tools and interfaces. ")
                $selection.TypeText("Consult the manufacturer's documentation for firmware update procedures.")
            } elseif ($item.Product -like "*Microsoft Teams*") {
                $selection.TypeText("Microsoft Teams can be updated via RMM script deployed through ConnectWise Automate. ")
                $selection.TypeText("This can sometimes be remediated by cleaning up the unused user profile installed versions using: ")
                $selection.TypeText("Select Scripts > RR - Custom > RR - Custom - R-Security Remediation > R-Security - Teams Classic Cleanup Remediation in RMM.")
            } else {
                $selection.TypeText("This application should be updated to the latest version. ")
                $selection.TypeText("If available via ConnectWise Automate/RMM, deploy updates using the patch management system. ")
                $selection.TypeText("Otherwise, manual updates may be required on affected systems.")
            }

            $selection.ParagraphFormat.LeftIndent = 0  # Reset indent
            $selection.TypeParagraph()

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
        [string]$OutputPath
    )

    try {
        Write-Log "Generating email template..."

        $year = (Get-Date).Year
        $quarter = Get-CurrentQuarter
        $greeting = Get-TimeOfDayGreeting

        $emailContent = @"
Subject: $year Q$quarter Vulnerability Scan Follow Up

Good $greeting,

We are pleased to inform you that your quarterly vulnerability scan report has been completed and added to your client folder.

The main list of items I recommend remediating can be found here:
<link to top ten report from onedrive>
*Note that not all vulnerabilities may be feasible to remediate depending on business need.

You can access and view the full reports using the link below:
<onedrive link to folder containing reports>

In this folder you will find:

Pending Remediation EPSS Score Report: This report classifies vulnerabilities by the "EPSS Score." This is a measure of the likelihood that an attacker will exploit a particular vulnerability within 30 days. The scale ranges from 0 to 1.0, with 1.0 being the most critical. <--- look in "Proposed Remediations (all)" for items that require your attention.

All Vulnerabilities Report: This spreadsheet contains a list of all vulnerabilities (including internal and external) that were detected, ranging from critical to low

Executive Summary Report: A high-level overview of your security "grade" as well as some information about your network

External Scan: Any detected vulnerabilities or services that are exposed to the outside Internet

Note: If you would like to discuss the report further, please contact your TSL or vCIO and we will coordinate a meeting. We will not begin remediation without your approval.

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

function New-TicketInstructions {
    param(
        [string]$OutputPath,
        [array]$TopTenData
    )

    try {
        Write-Log "Generating ticket instructions..."

        $sb = New-Object System.Text.StringBuilder
        [void]$sb.AppendLine("=".PadRight(100, '='))
        [void]$sb.AppendLine("TOP 10 VULNERABILITIES - TICKET INSTRUCTIONS")
        [void]$sb.AppendLine("Generated: $(Get-Date -Format 'yyyy-MM-dd HH:mm:ss')")
        [void]$sb.AppendLine("=".PadRight(100, '='))
        [void]$sb.AppendLine()

        for ($i = 0; $i -lt $TopTenData.Count; $i++) {
            $item = $TopTenData[$i]
            $num = $i + 1

            # Generate ticket subject based on product type
            $ticketSubject = "Vulnerability Scan - "
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
            } elseif ((Test-IsMicrosoftApplication -ProductName $item.Product) -and $script:IsRMITPlus) {
                $ticketSubject += "$($item.Product) - RMIT+ ticketed"
            } elseif ((Test-IsVMwareProduct -ProductName $item.Product) -and $script:IsRMITPlus) {
                $ticketSubject += "$($item.Product) - RMIT+ after-hours ticket created if we maintain this"
            } else {
                $ticketSubject += "$($item.Product) - Update Required"
            }

            [void]$sb.AppendLine()
            [void]$sb.AppendLine("-".PadRight(100, '-'))
            [void]$sb.AppendLine("VULNERABILITY #$num")
            [void]$sb.AppendLine("-".PadRight(100, '-'))
            [void]$sb.AppendLine()
            [void]$sb.AppendLine("TICKET SUBJECT:")
            [void]$sb.AppendLine("  $ticketSubject")
            [void]$sb.AppendLine()
            [void]$sb.AppendLine("Product/System:          $($item.Product)")
            [void]$sb.AppendLine("Risk Score:              $($item.RiskScore.ToString('N2'))")
            [void]$sb.AppendLine("EPSS Score:              $($item.EPSSScore.ToString('N4'))")
            [void]$sb.AppendLine("Average CVSS:            $($item.AvgCVSS.ToString('N2'))")
            [void]$sb.AppendLine("Total Vulnerabilities:   $($item.VulnCount)")
            [void]$sb.AppendLine("Affected Systems Count:  $($item.AffectedSystems.Count)")
            [void]$sb.AppendLine()
            [void]$sb.AppendLine("Affected Systems:")
            # Group by hostname and IP to get unique systems, then format as "hostname - IP"
            $uniqueSystems = $item.AffectedSystems | Select-Object HostName, IP -Unique
            foreach ($sys in $uniqueSystems) {
                $hostname = $sys.HostName
                $ip = $sys.IP
                if (-not [string]::IsNullOrWhiteSpace($ip)) {
                    [void]$sb.AppendLine("  - $hostname - $ip")
                } else {
                    [void]$sb.AppendLine("  - $hostname")
                }
            }
            [void]$sb.AppendLine()
            [void]$sb.AppendLine("Remediation Instructions:")

            # Determine remediation type
            if ($item.Product -like "*Windows Server 2012*" -or $item.Product -like "*end-of-life*" -or $item.Product -like "*out of support*") {
                [void]$sb.AppendLine("  - This end-of-support operating system represents an infrastructure project")
                [void]$sb.AppendLine("  - Consider planning a migration to a supported operating system version")
            } elseif ($item.Product -like "*Windows 10*") {
                [void]$sb.AppendLine("  - Windows 10 reached End of Life on October 14, 2025")
                [void]$sb.AppendLine("  - No longer supported unless you have extended support licensing")
                [void]$sb.AppendLine("  - If Windows Updates are functional and no extension licensing in place:")
                [void]$sb.AppendLine("    * Nothing to be done other than considering upgrade to Windows 11 or retiring machine")
                [void]$sb.AppendLine("  - For systems with extension licensing:")
                [void]$sb.AppendLine("    * Continue to verify Windows Update status through ConnectWise Automate")
            } elseif ($item.Product -like "*Windows*") {
                [void]$sb.AppendLine("  - Investigate via ConnectWise Automate")
                [void]$sb.AppendLine("  - Verify Windows Update status on affected systems")
                [void]$sb.AppendLine("  - Check for any issues preventing patch installation")
            } elseif ($item.Product -like "*printer*" -or $item.Product -like "*Ripple20*") {
                [void]$sb.AppendLine("  - Requires manual firmware updates via manufacturer tools")
                [void]$sb.AppendLine("  - Consult manufacturer documentation for update procedures")
            } elseif ($item.Product -like "*Microsoft Teams*") {
                [void]$sb.AppendLine("  - Update via RMM script deployed through ConnectWise Automate")
                [void]$sb.AppendLine("  - Can sometimes be remediated by cleaning up unused user profile installed versions")
                [void]$sb.AppendLine("  - Script path: Select Scripts > RR - Custom > RR - Custom - R-Security Remediation > R-Security - Teams Classic Cleanup Remediation in RMM")
            } else {
                [void]$sb.AppendLine("  - Update to latest version")
                [void]$sb.AppendLine("  - Deploy via ConnectWise Automate/RMM if available")
                [void]$sb.AppendLine("  - Otherwise, manual updates required on affected systems")
            }
            [void]$sb.AppendLine()
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
    # Task variations
    $taskVariations = @(
        "Set up and pull vulnerability scans and verify for accuracy. Send out encrypted email with reports and recommendations.",
        "Configured and executed vulnerability scans, verified accuracy, and sent encrypted email with reports and recommendations.",
        "Pulled vulnerability scans, verified for accuracy, and delivered encrypted email containing reports and recommendations.",
        "Set up vulnerability scans and verified accuracy. Transmitted encrypted email with comprehensive reports and recommendations.",
        "Executed vulnerability scans with accuracy verification. Sent secure encrypted email with reports and recommendations to contact.",
        "Configured vulnerability scans, performed accuracy verification, and delivered encrypted email with reports and recommendations.",
        "Set up and ran vulnerability scans with verification. Sent encrypted email containing reports and recommendations.",
        "Initiated vulnerability scans, verified accuracy of results, and sent encrypted email with detailed reports and recommendations."
    )

    # Steps performed variations
    $stepVariations = @(
        ,@("Reviewed lightweight agents", "Checked lightweight agents", "Examined lightweight agents", "Verified lightweight agent configuration"),
        ,@("Reviewed probe", "Checked probe configuration", "Examined probe settings", "Verified probe setup"),
        ,@("Compared agent/probe count vs other systems", "Verified agent/probe count against other systems", "Checked agent/probe count compared to other systems", "Analyzed agent/probe count relative to other systems"),
        ,@("Reviewed mapped credentials", "Checked mapped credentials", "Verified mapped credentials", "Examined credential mappings"),
        ,@("Reviewed external assets", "Checked external assets", "Examined external assets", "Verified external asset inventory"),
        ,@("Reviewed nmap interface on probe", "Checked nmap interface on probe", "Examined probe nmap interface", "Verified nmap interface configuration on probe"),
        ,@("Reviewed deprecated items", "Checked deprecated items", "Examined deprecated items", "Verified deprecated item list"),
        ,@("Generated all reports", "Created all reports", "Produced all reports", "Compiled all reports"),
        ,@("Analyzed reports", "Reviewed reports", "Examined reports", "Assessed reports"),
        ,@("Made a top ten vulnerabilities docx file", "Created top ten vulnerabilities docx file", "Generated top ten vulnerabilities docx file", "Produced top ten vulnerabilities docx report"),
        ,@("Sent encrypted email to contact with reports", "Delivered encrypted email with reports to contact", "Transmitted encrypted email containing reports to contact", "Sent secure email with reports to contact")
    )

    # Task resolved variations
    $taskResolvedVariations = @(
        "Task complete",
        "Completed",
        "Task completed",
        "Yes - task complete",
        "Yes - completed"
    )

    # Select random task
    $task = $taskVariations | Get-Random

    # Generate random steps - select ONE random variation from each step category
    $steps = New-Object System.Collections.ArrayList
    foreach ($varSet in $stepVariations) {
        if ($varSet -is [Array] -and $varSet.Count -gt 0) {
            # Use Get-Random directly on the array to select ONE element
            $selectedStep = $varSet | Get-Random
            # Ensure it's a string and not an array
            $selectedStep = [string]$selectedStep
            if ($selectedStep -and $selectedStep.Trim().Length -gt 0) {
                [void]$steps.Add("- " + $selectedStep.Trim())
            }
        }
    }
    $stepsText = $steps -join "`r`n"

    # Select random task resolved
    $taskResolved = $taskResolvedVariations | Get-Random

    # Build full ticket notes
    $result = @"
**Task**

$task


**Steps performed**

$stepsText


**Is the task resolved?**

$taskResolved


**Next step(s)**



**Special note or recommendation(s)**


"@

    # Copy to clipboard
    try {
        [System.Windows.Forms.Clipboard]::SetText($result)
        [System.Windows.Forms.MessageBox]::Show("Ticket notes copied to clipboard!", "Success", [System.Windows.Forms.MessageBoxButtons]::OK, [System.Windows.Forms.MessageBoxIcon]::Information)
    } catch {
        [System.Windows.Forms.MessageBox]::Show("Failed to copy to clipboard: $($_.Exception.Message)", "Error", [System.Windows.Forms.MessageBoxButtons]::OK, [System.Windows.Forms.MessageBoxIcon]::Error)
    }
}

# --- GUI Functions ---

function Show-SettingsDialog {
    $settingsForm = New-Object System.Windows.Forms.Form
    $settingsForm.Text = "User Settings"
    $settingsForm.Size = New-Object System.Drawing.Size(500, 400)
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

        if (Save-UserSettings) {
            [System.Windows.Forms.MessageBox]::Show("Settings saved successfully!", "Success",
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

function Show-VScanMagicGUI {
    # Initialize script-level variables for output file paths
    $script:WordReportPath = $null
    $script:ExcelReportPath = $null
    $script:EmailTemplatePath = $null
    $script:TicketInstructionsPath = $null
    $script:IsRMITPlus = $false

    # Load user settings from disk
    Load-UserSettings

    # Create main form
    $form = New-Object System.Windows.Forms.Form
    $form.Text = "$($script:Config.AppName) - Vulnerability Report Generator"
    $form.Size = New-Object System.Drawing.Size(700, 720)
    $form.StartPosition = "CenterScreen"
    $form.FormBorderStyle = "FixedDialog"
    $form.MaximizeBox = $false

    # --- Ticket Notes Button (Top Right) ---
    $buttonTicketNotes = New-Object System.Windows.Forms.Button
    $buttonTicketNotes.Location = New-Object System.Drawing.Point(450, 10)
    $buttonTicketNotes.Size = New-Object System.Drawing.Size(120, 25)
    $buttonTicketNotes.Text = "Ticket Notes"
    $buttonTicketNotes.Add_Click({
        New-TicketNotes
    })
    $form.Controls.Add($buttonTicketNotes)

    # --- Settings Button (Top Right) ---
    $buttonSettings = New-Object System.Windows.Forms.Button
    $buttonSettings.Location = New-Object System.Drawing.Point(580, 10)
    $buttonSettings.Size = New-Object System.Drawing.Size(90, 25)
    $buttonSettings.Text = "Settings"
    $buttonSettings.Add_Click({
        Show-SettingsDialog
    })
    $form.Controls.Add($buttonSettings)

    # --- Input File Section ---
    $labelInputFile = New-Object System.Windows.Forms.Label
    $labelInputFile.Location = New-Object System.Drawing.Point(20, 20)
    $labelInputFile.Size = New-Object System.Drawing.Size(200, 20)
    $labelInputFile.Text = "Pending EPSS Report (XLSX):"
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
        $openFileDialog.Title = "Select Pending EPSS Report"

        if ($openFileDialog.ShowDialog() -eq [System.Windows.Forms.DialogResult]::OK) {
            $textBoxInputFile.Text = $openFileDialog.FileName

            # Automatically set output directory to input file's directory
            $inputDirectory = [System.IO.Path]::GetDirectoryName($openFileDialog.FileName)
            $textBoxOutputDir.Text = $inputDirectory

            # Extract company name from filename
            # Try multiple patterns to detect company name
            $fileName = [System.IO.Path]::GetFileNameWithoutExtension($openFileDialog.FileName)
            Write-Log "Attempting to extract company name from filename: $fileName"

            $companyName = $null
            # Pattern 1: "...Reports-{CompanyName}_{timestamp}" or "...Report-{CompanyName}_..."
            if ($fileName -match 'Reports?-([^\s_-]+)(?:_|$)') {
                $rawName = $matches[1]
                # Skip if it's a report-related keyword
                if ($rawName -notmatch '^(Pending|EPSS|Report|Reports?|Vulnerability|Security)$') {
                    $companyName = $rawName
                    Write-Log "Matched Pattern 1 (Reports-Company_): $companyName"
                } else {
                    Write-Log "Pattern 1 matched but result was report keyword: $rawName" -Level Warning
                }
            }
            # Pattern 2: "{CompanyName}-Reports" or "{CompanyName}_Reports" (but not report keywords)
            if (-not $companyName -and $fileName -match '^([^\s_-]+)[-_]Reports?') {
                $rawName = $matches[1]
                if ($rawName -notmatch '^(Pending|EPSS|Report|Reports?|Vulnerability|Security)$') {
                    $companyName = $rawName
                    Write-Log "Matched Pattern 2 (Company-Reports): $companyName"
                } else {
                    Write-Log "Pattern 2 matched but result was report keyword: $rawName" -Level Warning
                }
            }
            # Pattern 3: Any text before first delimiter (space, underscore, hyphen), but exclude report-related keywords
            if (-not $companyName -and $fileName -match '^([^\s_-]+)') {
                $rawName = $matches[1]
                if ($rawName -notmatch '(Pending|EPSS|Report|Reports?|Vulnerability|Security)') {
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

    # --- Client Type ---
    $groupBoxClientType = New-Object System.Windows.Forms.GroupBox
    $groupBoxClientType.Location = New-Object System.Drawing.Point(570, 85)
    $groupBoxClientType.Size = New-Object System.Drawing.Size(100, 75)
    $groupBoxClientType.Text = "Client Type"
    $form.Controls.Add($groupBoxClientType)

    $radioRMITCMIT = New-Object System.Windows.Forms.RadioButton
    $radioRMITCMIT.Location = New-Object System.Drawing.Point(10, 20)
    $radioRMITCMIT.Size = New-Object System.Drawing.Size(80, 20)
    $radioRMITCMIT.Text = "RMIT/CMIT"
    $radioRMITCMIT.Checked = $true
    $groupBoxClientType.Controls.Add($radioRMITCMIT)

    $radioRMITPlus = New-Object System.Windows.Forms.RadioButton
    $radioRMITPlus.Location = New-Object System.Drawing.Point(10, 45)
    $radioRMITPlus.Size = New-Object System.Drawing.Size(80, 20)
    $radioRMITPlus.Text = "RMIT+"
    $radioRMITPlus.Checked = $false
    $groupBoxClientType.Controls.Add($radioRMITPlus)

    # --- Output Options ---
    $groupBoxOutput = New-Object System.Windows.Forms.GroupBox
    $groupBoxOutput.Location = New-Object System.Drawing.Point(20, 170)
    $groupBoxOutput.Size = New-Object System.Drawing.Size(630, 210)
    $groupBoxOutput.Text = "Output Options"
    $form.Controls.Add($groupBoxOutput)

    $checkBoxExcel = New-Object System.Windows.Forms.CheckBox
    $checkBoxExcel.Location = New-Object System.Drawing.Point(20, 25)
    $checkBoxExcel.Size = New-Object System.Drawing.Size(250, 20)
    $checkBoxExcel.Text = "Generate Pending EPSS Report (Excel)"
    $checkBoxExcel.Checked = $true
    $groupBoxOutput.Controls.Add($checkBoxExcel)

    $checkBoxWord = New-Object System.Windows.Forms.CheckBox
    $checkBoxWord.Location = New-Object System.Drawing.Point(20, 50)
    $checkBoxWord.Size = New-Object System.Drawing.Size(300, 20)
    $checkBoxWord.Text = "Generate Top Ten Vulnerabilities Report (Word)"
    $checkBoxWord.Checked = $true
    $groupBoxOutput.Controls.Add($checkBoxWord)

    $checkBoxEmailTemplate = New-Object System.Windows.Forms.CheckBox
    $checkBoxEmailTemplate.Location = New-Object System.Drawing.Point(20, 75)
    $checkBoxEmailTemplate.Size = New-Object System.Drawing.Size(300, 20)
    $checkBoxEmailTemplate.Text = "Generate Email Template (Text)"
    $checkBoxEmailTemplate.Checked = $false
    $groupBoxOutput.Controls.Add($checkBoxEmailTemplate)

    $checkBoxTicketInstructions = New-Object System.Windows.Forms.CheckBox
    $checkBoxTicketInstructions.Location = New-Object System.Drawing.Point(20, 100)
    $checkBoxTicketInstructions.Size = New-Object System.Drawing.Size(300, 20)
    $checkBoxTicketInstructions.Text = "Generate Ticket Instructions (Text)"
    $checkBoxTicketInstructions.Checked = $false
    $groupBoxOutput.Controls.Add($checkBoxTicketInstructions)

    # --- Output Directory ---
    $labelOutputDir = New-Object System.Windows.Forms.Label
    $labelOutputDir.Location = New-Object System.Drawing.Point(20, 395)
    $labelOutputDir.Size = New-Object System.Drawing.Size(150, 20)
    $labelOutputDir.Text = "Output Directory:"
    $form.Controls.Add($labelOutputDir)

    $textBoxOutputDir = New-Object System.Windows.Forms.TextBox
    $textBoxOutputDir.Location = New-Object System.Drawing.Point(20, 420)
    $textBoxOutputDir.Size = New-Object System.Drawing.Size(520, 20)
    $textBoxOutputDir.Text = [Environment]::GetFolderPath("Desktop")
    $form.Controls.Add($textBoxOutputDir)

    $buttonBrowseOutput = New-Object System.Windows.Forms.Button
    $buttonBrowseOutput.Location = New-Object System.Drawing.Point(550, 418)
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

    # --- Progress Section ---
    $script:StatusLabel = New-Object System.Windows.Forms.Label
    $script:StatusLabel.Location = New-Object System.Drawing.Point(20, 455)
    $script:StatusLabel.Size = New-Object System.Drawing.Size(630, 20)
    $script:StatusLabel.Text = "Ready"
    $script:StatusLabel.Visible = $false
    $form.Controls.Add($script:StatusLabel)

    $script:ProgressBar = New-Object System.Windows.Forms.ProgressBar
    $script:ProgressBar.Location = New-Object System.Drawing.Point(20, 480)
    $script:ProgressBar.Size = New-Object System.Drawing.Size(630, 20)
    $script:ProgressBar.Style = 'Marquee'
    $script:ProgressBar.MarqueeAnimationSpeed = 30
    $script:ProgressBar.Visible = $false
    $form.Controls.Add($script:ProgressBar)

    # --- Log Section ---
    $labelLog = New-Object System.Windows.Forms.Label
    $labelLog.Location = New-Object System.Drawing.Point(20, 505)
    $labelLog.Size = New-Object System.Drawing.Size(150, 20)
    $labelLog.Text = "Processing Log:"
    $form.Controls.Add($labelLog)

    $script:LogTextBox = New-Object System.Windows.Forms.TextBox
    $script:LogTextBox.Location = New-Object System.Drawing.Point(20, 530)
    $script:LogTextBox.Size = New-Object System.Drawing.Size(630, 50)
    $script:LogTextBox.Multiline = $true
    $script:LogTextBox.ScrollBars = "Vertical"
    $script:LogTextBox.ReadOnly = $true
    $script:LogTextBox.Font = New-Object System.Drawing.Font("Consolas", 9)
    $form.Controls.Add($script:LogTextBox)

    # --- Open Report Buttons ---
    $script:buttonOpenWord = New-Object System.Windows.Forms.Button
    $script:buttonOpenWord.Location = New-Object System.Drawing.Point(20, 590)
    $script:buttonOpenWord.Size = New-Object System.Drawing.Size(150, 25)
    $script:buttonOpenWord.Text = "Open Top Ten"
    $script:buttonOpenWord.Enabled = $false
    $script:buttonOpenWord.Add_Click({
        if ($script:WordReportPath -and (Test-Path $script:WordReportPath)) {
            Start-Process $script:WordReportPath
        }
    })
    $form.Controls.Add($script:buttonOpenWord)

    $script:buttonOpenExcel = New-Object System.Windows.Forms.Button
    $script:buttonOpenExcel.Location = New-Object System.Drawing.Point(180, 590)
    $script:buttonOpenExcel.Size = New-Object System.Drawing.Size(150, 25)
    $script:buttonOpenExcel.Text = "Open EPSS Report"
    $script:buttonOpenExcel.Enabled = $false
    $script:buttonOpenExcel.Add_Click({
        if ($script:ExcelReportPath -and (Test-Path $script:ExcelReportPath)) {
            Start-Process $script:ExcelReportPath
        }
    })
    $form.Controls.Add($script:buttonOpenExcel)

    $script:buttonOpenEmail = New-Object System.Windows.Forms.Button
    $script:buttonOpenEmail.Location = New-Object System.Drawing.Point(340, 590)
    $script:buttonOpenEmail.Size = New-Object System.Drawing.Size(150, 25)
    $script:buttonOpenEmail.Text = "Open Email Template"
    $script:buttonOpenEmail.Enabled = $false
    $script:buttonOpenEmail.Add_Click({
        if ($script:EmailTemplatePath -and (Test-Path $script:EmailTemplatePath)) {
            Start-Process $script:EmailTemplatePath
        }
    })
    $form.Controls.Add($script:buttonOpenEmail)

    $script:buttonOpenTicket = New-Object System.Windows.Forms.Button
    $script:buttonOpenTicket.Location = New-Object System.Drawing.Point(500, 590)
    $script:buttonOpenTicket.Size = New-Object System.Drawing.Size(150, 25)
    $script:buttonOpenTicket.Text = "Open Ticket Instr."
    $script:buttonOpenTicket.Enabled = $false
    $script:buttonOpenTicket.Add_Click({
        if ($script:TicketInstructionsPath -and (Test-Path $script:TicketInstructionsPath)) {
            Start-Process $script:TicketInstructionsPath
        }
    })
    $form.Controls.Add($script:buttonOpenTicket)

    # --- Action Buttons ---
    $buttonGenerate = New-Object System.Windows.Forms.Button
    $buttonGenerate.Location = New-Object System.Drawing.Point(450, 630)
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

        if (-not $checkBoxExcel.Checked -and -not $checkBoxWord.Checked -and -not $checkBoxEmailTemplate.Checked -and -not $checkBoxTicketInstructions.Checked) {
            [System.Windows.Forms.MessageBox]::Show("Please select at least one output option.", "Validation Error",
                [System.Windows.Forms.MessageBoxButtons]::OK, [System.Windows.Forms.MessageBoxIcon]::Warning)
            return
        }

        # Capture client type selection
        $script:IsRMITPlus = $radioRMITPlus.Checked

        # Disable button during processing
        $buttonGenerate.Enabled = $false
        $script:LogTextBox.Clear()

        # Disable open buttons at start
        $script:buttonOpenWord.Enabled = $false
        $script:buttonOpenExcel.Enabled = $false
        $script:buttonOpenEmail.Enabled = $false
        $script:buttonOpenTicket.Enabled = $false

        try {
            Write-Log "=== Starting VScanMagic Processing ===" -Level Info
            Write-Log "Input File: $($textBoxInputFile.Text)"
            Write-Log "Client: $($textBoxClientName.Text)"
            Write-Log "Client Type: $(if ($script:IsRMITPlus) { 'RMIT+' } else { 'RMIT/CMIT' })"
            Write-Log "Scan Date: $($datePickerScanDate.Value.ToShortDateString())"

            # Read vulnerability data from all remediation sheets
            Update-Progress -Status "Reading vulnerability data from Excel file..." -Show $true
            $vulnData = Get-VulnerabilityData -ExcelPath $textBoxInputFile.Text

            # Calculate top 10 vulnerabilities
            Update-Progress -Status "Calculating top 10 vulnerabilities..." -Show $true
            $top10 = Get-Top10Vulnerabilities -VulnData $vulnData

            # Generate Word report
            if ($checkBoxWord.Checked) {
                Update-Progress -Status "Generating Top Ten Vulnerabilities Report (Word)..." -Show $true
                $companyName = $textBoxClientName.Text
                if ([string]::IsNullOrWhiteSpace($companyName)) {
                    $companyName = "Client"
                }
                # Add timestamp to filename to avoid duplicate name conflicts
                $timestamp = Get-Date -Format "yyyy-MM-dd_HHmmss"
                $wordOutputPath = Join-Path $textBoxOutputDir.Text "$companyName Top Ten Vulnerabilities Report_$timestamp.docx"

                Invoke-OperationWithRetry -OperationName "Word Report Generation" -Operation {
                    New-WordReport -OutputPath $wordOutputPath `
                                  -ClientName $textBoxClientName.Text `
                                  -ScanDate $datePickerScanDate.Value.ToShortDateString() `
                                  -Top10Data $top10
                }

                # Store path and enable open button
                $script:WordReportPath = $wordOutputPath
                $script:buttonOpenWord.Enabled = $true

                Write-Log "Top Ten Vulnerabilities Report saved to: $wordOutputPath" -Level Success
            }

            # Generate Excel report
            if ($checkBoxExcel.Checked) {
                Update-Progress -Status "Generating Pending EPSS Report (Excel)..." -Show $true
                $companyName = $textBoxClientName.Text
                if ([string]::IsNullOrWhiteSpace($companyName)) {
                    $companyName = "Client"
                }
                # Add timestamp to filename to avoid duplicate name conflicts
                $timestamp = Get-Date -Format "yyyy-MM-dd_HHmmss"
                $excelOutputPath = Join-Path $textBoxOutputDir.Text "$companyName Pending EPSS Report_$timestamp.xlsx"

                Invoke-OperationWithRetry -OperationName "Excel Report Generation" -Operation {
                    New-ExcelReport -InputPath $textBoxInputFile.Text -OutputPath $excelOutputPath
                }

                # Store path and enable open button
                $script:ExcelReportPath = $excelOutputPath
                $script:buttonOpenExcel.Enabled = $true

                Write-Log "Pending EPSS Report saved to: $excelOutputPath" -Level Success
            }

            # Generate Email Template
            if ($checkBoxEmailTemplate.Checked) {
                Update-Progress -Status "Generating Email Template..." -Show $true
                $companyName = $textBoxClientName.Text
                if ([string]::IsNullOrWhiteSpace($companyName)) {
                    $companyName = "Client"
                }
                # Add timestamp to filename
                $timestamp = Get-Date -Format "yyyy-MM-dd_HHmmss"
                $emailOutputPath = Join-Path $textBoxOutputDir.Text "$companyName Email Template_$timestamp.txt"

                New-EmailTemplate -OutputPath $emailOutputPath

                # Store path and enable open button
                $script:EmailTemplatePath = $emailOutputPath
                $script:buttonOpenEmail.Enabled = $true

                Write-Log "Email Template saved to: $emailOutputPath" -Level Success
            }

            # Generate Ticket Instructions
            if ($checkBoxTicketInstructions.Checked) {
                Update-Progress -Status "Generating Ticket Instructions..." -Show $true
                $companyName = $textBoxClientName.Text
                if ([string]::IsNullOrWhiteSpace($companyName)) {
                    $companyName = "Client"
                }
                # Add timestamp to filename
                $timestamp = Get-Date -Format "yyyy-MM-dd_HHmmss"
                $ticketOutputPath = Join-Path $textBoxOutputDir.Text "$companyName Ticket Instructions_$timestamp.txt"

                New-TicketInstructions -OutputPath $ticketOutputPath -TopTenData $top10

                # Store path and enable open button
                $script:TicketInstructionsPath = $ticketOutputPath
                $script:buttonOpenTicket.Enabled = $true

                Write-Log "Ticket Instructions saved to: $ticketOutputPath" -Level Success
            }

            # Hide progress bar
            Update-Progress -Status "Complete" -Show $false

            Write-Log "=== Processing Complete ===" -Level Success

            [System.Windows.Forms.MessageBox]::Show("Report generation completed successfully!", "Success",
                [System.Windows.Forms.MessageBoxButtons]::OK, [System.Windows.Forms.MessageBoxIcon]::Information)

        } catch {
            # Hide progress bar on error
            Update-Progress -Status "Error" -Show $false

            Write-Log "Processing failed: $($_.Exception.Message)" -Level Error
            [System.Windows.Forms.MessageBox]::Show("An error occurred during processing. Check the log for details.", "Error",
                [System.Windows.Forms.MessageBoxButtons]::OK, [System.Windows.Forms.MessageBoxIcon]::Error)
        } finally {
            $buttonGenerate.Enabled = $true
        }
    })
    $form.Controls.Add($buttonGenerate)

    $buttonClose = New-Object System.Windows.Forms.Button
    $buttonClose.Location = New-Object System.Drawing.Point(560, 630)
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
