# VScanMagic-Core.ps1 - Config, settings, persistence, helpers
# Dot-sourced by VScanMagic-GUI.ps1 (do not run directly)

# --- Configuration ---
$script:Config = @{
    AppName = "VScanMagic v4"
    Version = "4.0.1"
    Author = "River Run MSP"

    # Risk Score Calculation - ConnectSecure-aligned methodology
    # Severity weights from ConnectSecure Problem Category Weightage (External Scan / Asset Risk)
    SeverityWeights = @{
        Critical = 0.90
        High = 0.80
        Medium = 0.50
        Low = 0.30
    }
    # CVSS equivalents for Average CVSS display (used for reporting, not primary risk calculation)
    CVSSEquivalent = @{
        Critical = 9.0
        High = 7.0
        Medium = 5.0
        Low = 3.0
    }

    # Heatmap Color Thresholds for Risk Scores (Yellow to Red gradient - no greens)
    # Risk Score uses dynamic thresholds based on max score in dataset
    RiskColors = [ordered]@{
        Critical   = @{ Threshold = 8.0;  Color = 'DC143C'; Name = 'Critical'; TextColor = 'FFFFFF' }      # Crimson red
        VeryHigh   = @{ Threshold = 6.0;  Color = 'FF4500'; Name = 'Very High'; TextColor = 'FFFFFF' }     # Orange-red
        High       = @{ Threshold = 4.0;  Color = 'FF8C00'; Name = 'High'; TextColor = 'FFFFFF' }          # Dark orange
        MediumHigh = @{ Threshold = 2.0;  Color = 'FFA500'; Name = 'Medium-High'; TextColor = '000000' }   # Orange
        Medium     = @{ Threshold = 0;    Color = 'FFFF00'; Name = 'Medium'; TextColor = '000000' }        # Yellow (baseline)
    }

    # Products to Filter Out (EOL is no longer filtered - it gets max risk weight instead)
    FilteredProducts = @(
    )

    # End of Life product patterns - ConnectSecure treats EOL as maximum risk (weight 1.0)
    EOLProductPatterns = @(
        'OS-OUT-OF-SUPPORT',
        'OS-OUT-OF-ACTIVE-SUPPORT',
        'OS-OUT-OF-SECURITY-SUPPORT',
        'END OF LIFE',
        'End of Life',
        'end-of-life',
        'out of support',
        'Out of Support'
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
# Use Windows LocalAppData for settings (standard Windows application data location)
# This resolves to: C:\Users\<Username>\AppData\Local\VScanMagic\
# Can be overridden via UserSettings.SettingsDirectory
$script:SettingsDirectory = Join-Path $env:LOCALAPPDATA "VScanMagic"
$script:SettingsPath = Join-Path $script:SettingsDirectory "VScanMagic_Settings.json"
$script:RemediationRulesPath = Join-Path $script:SettingsDirectory "VScanMagic_RemediationRules.json"
$script:RemediationRules = $null
$script:CoveredSoftwarePath = Join-Path $script:SettingsDirectory "VScanMagic_CoveredSoftware.json"
$script:CoveredSoftware = $null
$script:GeneralRecommendationsPath = Join-Path $script:SettingsDirectory "VScanMagic_GeneralRecommendations.json"
$script:GeneralRecommendations = $null
$script:ConnectSecureCredentialsPath = Join-Path $script:SettingsDirectory "ConnectSecure-Credentials.json"
$script:ConnectSecureCompaniesCachePath = Join-Path $script:SettingsDirectory "ConnectSecure-Companies-Cache.json"

# Report Filters and Output Options (modified via dialogs)
$script:FilterMinEPSS = 0
$script:FilterIncludeCritical = $true
$script:FilterIncludeHigh = $true
$script:FilterIncludeMedium = $true
$script:FilterIncludeLow = $true
$script:FilterTopN = "10"
$script:OutputExcel = $true
$script:OutputWord = $true
$script:OutputEmailTemplate = $true
$script:OutputTicketInstructions = $true
$script:OutputTimeEstimate = $true

# ScriptDirectory is set by the bootstrap (VScanMagic-GUI.ps1) before dot-sourcing

# Migration: Check for old settings file in script/exe directory
$oldSettingsPath = $null
if (-not [string]::IsNullOrEmpty($script:ScriptDirectory)) {
    $oldSettingsPath = Join-Path $script:ScriptDirectory "VScanMagic_Settings.json"
}

# Migrate old settings file if it exists and new location doesn't have settings yet
if ($oldSettingsPath -and (Test-Path $oldSettingsPath) -and -not (Test-Path $script:SettingsPath)) {
    try {
        Copy-Item -Path $oldSettingsPath -Destination $script:SettingsPath -Force
        Write-Host "Migrated settings from $oldSettingsPath to $script:SettingsPath"
    } catch {
        Write-Warning "Could not migrate old settings: $($_.Exception.Message)"
    }
}
$script:UserSettings = @{
    PreparedBy = "River Run MSP"
    CompanyName = ""
    CompanyAddress = ""
    Email = ""
    PhoneNumber = ""
    CompanyPhoneNumber = ""
    SettingsDirectory = ""  # Empty = use default (LOCALAPPDATA\VScanMagic)
    LastOutputDirectory = ""  # Last-used output directory for reports
}

function Ensure-SettingsDirectory {
    param([string]$Path = $script:SettingsDirectory)
    if ([string]::IsNullOrEmpty($Path)) { return $false }
    if (Test-Path $Path) { return $false }
    try { New-Item -Path $Path -ItemType Directory -Force | Out-Null; return $true } catch { return $false }
}

function Get-JsonFile {
    param([Parameter(Mandatory=$true)][string]$Path)
    if (-not (Test-Path $Path)) { return $null }
    try {
        return Get-Content $Path -Raw -Encoding UTF8 | ConvertFrom-Json
    } catch {
        Write-Warning "Could not load JSON from $Path : $($_.Exception.Message)"
        return $null
    }
}

function Set-JsonFile {
    param([Parameter(Mandatory=$true)][string]$Path, [Parameter(Mandatory=$true)]$Object, [int]$Depth = 4)
    $dir = [System.IO.Path]::GetDirectoryName($Path)
    Ensure-SettingsDirectory -Path $dir
    try {
        $Object | ConvertTo-Json -Depth $Depth | Set-Content $Path -Encoding UTF8 -Force
        return $true
    } catch {
        Write-Warning "Could not save JSON to $Path : $($_.Exception.Message)"
        return $false
    }
}

function Update-SettingsPaths {
    # Update paths based on custom directory if set, otherwise use default
    if (-not [string]::IsNullOrEmpty($script:UserSettings.SettingsDirectory) -and 
        (Test-Path $script:UserSettings.SettingsDirectory)) {
        $script:SettingsDirectory = $script:UserSettings.SettingsDirectory
    } else {
        # Use default location
        $script:SettingsDirectory = Join-Path $env:LOCALAPPDATA "VScanMagic"
    }
    
    # Update all path variables
    $script:SettingsPath = Join-Path $script:SettingsDirectory "VScanMagic_Settings.json"
    $script:RemediationRulesPath = Join-Path $script:SettingsDirectory "VScanMagic_RemediationRules.json"
    $script:CoveredSoftwarePath = Join-Path $script:SettingsDirectory "VScanMagic_CoveredSoftware.json"
    $script:GeneralRecommendationsPath = Join-Path $script:SettingsDirectory "VScanMagic_GeneralRecommendations.json"
    
    if (Ensure-SettingsDirectory -Path $script:SettingsDirectory) { Write-Host "Created settings directory: $script:SettingsDirectory" }
}

function Load-UserSettings {
    $defaultSettingsPath = Join-Path (Join-Path $env:LOCALAPPDATA "VScanMagic") "VScanMagic_Settings.json"
    $json = Get-JsonFile -Path $defaultSettingsPath
    if ($json -and $json.SettingsDirectory) { $script:UserSettings.SettingsDirectory = $json.SettingsDirectory }
    Update-SettingsPaths

    $json = Get-JsonFile -Path $script:SettingsPath
    if ($json) {
        $script:UserSettings.PreparedBy = if ($json.PreparedBy) { $json.PreparedBy } else { "River Run MSP" }
        $script:UserSettings.CompanyName = if ($json.CompanyName) { $json.CompanyName } else { "" }
        $script:UserSettings.CompanyAddress = if ($json.CompanyAddress) { $json.CompanyAddress } else { "" }
        $script:UserSettings.Email = if ($json.Email) { $json.Email } else { "" }
        $script:UserSettings.PhoneNumber = if ($json.PhoneNumber) { $json.PhoneNumber } else { "" }
        $script:UserSettings.CompanyPhoneNumber = if ($json.CompanyPhoneNumber) { $json.CompanyPhoneNumber } else { "" }
        if ($json.SettingsDirectory) { $script:UserSettings.SettingsDirectory = $json.SettingsDirectory }
        if ($json.LastOutputDirectory -and (Test-Path $json.LastOutputDirectory)) { $script:UserSettings.LastOutputDirectory = $json.LastOutputDirectory }
        Write-Host "User settings loaded from $script:SettingsPath"
    }
}

function Save-UserSettings {
    Update-SettingsPaths
    if ([string]::IsNullOrEmpty($script:SettingsPath)) {
        Write-Warning "Settings path is not set. Cannot save settings."
        return $false
    }
    if (Set-JsonFile -Path $script:SettingsPath -Object $script:UserSettings) {
        Write-Host "User settings saved to $script:SettingsPath"
        return $true
    }
    return $false
}

# --- ConnectSecure Credentials Persistence ---

function Load-ConnectSecureCredentials {
    if ([string]::IsNullOrEmpty($script:ConnectSecureCredentialsPath)) { return $null }
    $json = Get-JsonFile -Path $script:ConnectSecureCredentialsPath
    if (-not $json) { return $null }
    return @{
        BaseUrl = if ($json.BaseUrl) { $json.BaseUrl } else { "" }
        TenantName = if ($json.TenantName) { $json.TenantName } else { "" }
        ClientId = if ($json.ClientId) { $json.ClientId } else { "" }
        ClientSecret = if ($json.ClientSecret) { $json.ClientSecret } else { "" }
        CompanyId = if ($json.CompanyId) { [int]$json.CompanyId } else { 0 }
    }
}

function Save-ConnectSecureCredentials {
    param([string]$BaseUrl, [string]$TenantName, [string]$ClientId, [string]$ClientSecret, [int]$CompanyId = 0)
    if ([string]::IsNullOrEmpty($script:ConnectSecureCredentialsPath)) {
        Write-Warning "Credentials path is not set. Cannot save credentials."
        return $false
    }
    $credentials = @{ BaseUrl = $BaseUrl; TenantName = $TenantName; ClientId = $ClientId; ClientSecret = $ClientSecret; CompanyId = $CompanyId }
    if (Set-JsonFile -Path $script:ConnectSecureCredentialsPath -Object $credentials) {
        Write-Log "ConnectSecure credentials saved to $script:ConnectSecureCredentialsPath" -Level Success
        return $true
    }
    return $false
}

function Save-ConnectSecureCompaniesCache {
    param([string]$BaseUrl, [string]$TenantName, [array]$Companies)
    if ([string]::IsNullOrEmpty($script:ConnectSecureCompaniesCachePath) -or -not $Companies) { return $false }
    $cache = @{ BaseUrl = $BaseUrl; TenantName = $TenantName; CachedAt = (Get-Date -Format "o"); Companies = $Companies }
    return Set-JsonFile -Path $script:ConnectSecureCompaniesCachePath -Object $cache -Depth 5
}

function Load-ConnectSecureCompaniesCache {
    param([string]$BaseUrl, [string]$TenantName)
    if ([string]::IsNullOrEmpty($script:ConnectSecureCompaniesCachePath)) { return $null }
    $cache = Get-JsonFile -Path $script:ConnectSecureCompaniesCachePath
    if (-not $cache -or -not $cache.Companies) { return $null }
    $baseMatch = ($cache.BaseUrl -replace '/$','') -eq ($BaseUrl -replace '/$','')
    if ($baseMatch -and $cache.TenantName -eq $TenantName) { return $cache.Companies }
    return $null
}

# --- Remediation Rules Persistence ---

function Get-DefaultRemediationRules {
    return @(
        @{
            Pattern = "*Windows Server 2012*"
            WordText = "This end-of-support operating system represents an infrastructure project beyond the scope of quarterly vulnerability remediation. Consider planning a migration to a supported operating system version."
            TicketText = "- This end-of-support operating system represents an infrastructure project`r`n  - Consider planning a migration to a supported operating system version"
            IsDefault = $false
        },
        @{
            Pattern = "*end-of-life*"
            WordText = "This end-of-support operating system represents an infrastructure project beyond the scope of quarterly vulnerability remediation. Consider planning a migration to a supported operating system version."
            TicketText = "- This end-of-support operating system represents an infrastructure project`r`n  - Consider planning a migration to a supported operating system version"
            IsDefault = $false
        },
        @{
            Pattern = "*out of support*"
            WordText = "This end-of-support operating system represents an infrastructure project beyond the scope of quarterly vulnerability remediation. Consider planning a migration to a supported operating system version."
            TicketText = "- This end-of-support operating system represents an infrastructure project`r`n  - Consider planning a migration to a supported operating system version"
            IsDefault = $false
        },
        @{
            Pattern = "*Windows 10*"
            WordText = "Windows 10 reached End of Life on October 14, 2025, and is no longer supported by Microsoft unless you have extended support licensing. If Windows Updates are functional and no extension licensing is in place, there is nothing further to be done other than considering an upgrade to Windows 11 or retiring the machine. For systems with extension licensing, continue to verify Windows Update status through ConnectWise Automate."
            TicketText = "- Windows 10 reached End of Life on October 14, 2025`r`n  - No longer supported unless you have extended support licensing`r`n  - If Windows Updates are functional and no extension licensing in place:`r`n    * Nothing to be done other than considering upgrade to Windows 11 or retiring machine`r`n  - For systems with extension licensing:`r`n    * Continue to verify Windows Update status through ConnectWise Automate"
            IsDefault = $false
        },
        @{
            Pattern = "*Windows*"
            WordText = "Windows patch inconsistencies should be investigated via ConnectWise Automate. Systems with lower vulnerability counts may indicate that patching is working correctly and awaiting the latest patch cycles. For systems with high vulnerability counts, verify Windows Update status and investigate any potential issues preventing patch installation."
            TicketText = "- Investigate via ConnectWise Automate`r`n  - Verify Windows Update status on affected systems`r`n  - Check for any issues preventing patch installation"
            IsDefault = $false
        },
        @{
            Pattern = "*printer*"
            WordText = "Network printers and IoT devices require manual firmware updates via manufacturer-provided tools and interfaces. Consult the manufacturer's documentation for firmware update procedures."
            TicketText = "- Requires manual firmware updates via manufacturer tools`r`n  - Consult manufacturer documentation for update procedures"
            IsDefault = $false
        },
        @{
            Pattern = "*Ripple20*"
            WordText = "Network printers and IoT devices require manual firmware updates via manufacturer-provided tools and interfaces. Consult the manufacturer's documentation for firmware update procedures."
            TicketText = "- Requires manual firmware updates via manufacturer tools`r`n  - Consult manufacturer documentation for update procedures"
            IsDefault = $false
        },
        @{
            Pattern = "*Microsoft Teams*"
            WordText = "Microsoft Teams can be updated via RMM script deployed through ConnectWise Automate. This can be remediated by cleaning up unused user profile installed versions using: Select Scripts > RR - Custom > RR - Custom - R-Security Remediation > R-Security - Teams Classic Cleanup Remediation in RMM."
            TicketText = "- Update via RMM script deployed through ConnectWise Automate`r`n  - Can be remediated by cleaning up unused user profile installed versions`r`n  - Script path: Select Scripts > RR - Custom > RR - Custom - R-Security Remediation > R-Security - Teams Classic Cleanup Remediation in RMM"
            IsDefault = $false
        },
        @{
            Pattern = "*"
            WordText = "This application should be updated to the latest version. If available via ConnectWise Automate/RMM or scripting, deploy updates using the patch management system or scripts. Otherwise, manual updates may be required on affected systems."
            TicketText = "- Update to latest version`r`n  - Deploy via ConnectWise Automate/RMM or scripting if available`r`n  - Otherwise, manual updates required on affected systems"
            IsDefault = $true
        }
    )
}

function Load-RemediationRules {
    $script:CachedRemediationRulesForGuidance = $null  # invalidate cache
    $json = Get-JsonFile -Path $script:RemediationRulesPath
    if ($json -and $json.Count -gt 0) {
        $script:RemediationRules = @()
        foreach ($rule in $json) {
            $script:RemediationRules += @{ Pattern = $rule.Pattern; WordText = $rule.WordText; TicketText = $rule.TicketText; IsDefault = $rule.IsDefault }
        }
        Write-Host "Remediation rules loaded from $script:RemediationRulesPath"
    } else {
        $script:RemediationRules = Get-DefaultRemediationRules
        if (-not [string]::IsNullOrEmpty($script:RemediationRulesPath)) { Save-RemediationRules }
    }
}

function Save-RemediationRules {
    if ([string]::IsNullOrEmpty($script:RemediationRulesPath)) {
        Write-Warning "Remediation rules path is not set. Cannot save rules."
        return $false
    }
    $script:CachedRemediationRulesForGuidance = $null  # invalidate cache
    if (Set-JsonFile -Path $script:RemediationRulesPath -Object $script:RemediationRules -Depth 10) {
        Write-Host "Remediation rules saved to $script:RemediationRulesPath"
        return $true
    }
    return $false
}

# --- Covered Software Persistence ---

function Get-DefaultCoveredSoftware {
    return @(
        @{ Pattern = "*Microsoft*"; IsPattern = $true; Override = $false },
        @{ Pattern = "*Adobe*"; IsPattern = $true; Override = $false },
        @{ Pattern = "*Google Chrome*"; IsPattern = $true; Override = $false },
        @{ Pattern = "*Mozilla Firefox*"; IsPattern = $true; Override = $false }
    )
}

function Load-CoveredSoftware {
    if (-not [string]::IsNullOrEmpty($script:CoveredSoftwarePath) -and (Test-Path $script:CoveredSoftwarePath)) {
        try {
            $json = Get-Content $script:CoveredSoftwarePath -Raw | ConvertFrom-Json
            $script:CoveredSoftware = @()
            foreach ($item in $json) {
                $script:CoveredSoftware += @{
                    Pattern = $item.Pattern
                    IsPattern = $item.IsPattern
                    Override = if ($null -ne $item.Override) { $item.Override } else { $false }
                }
            }
            Write-Host "Covered software list loaded from $script:CoveredSoftwarePath"
        } catch {
            Write-Warning "Could not load covered software list: $($_.Exception.Message). Using defaults."
            $script:CoveredSoftware = Get-DefaultCoveredSoftware
            Save-CoveredSoftware
        }
    } else {
        # Initialize with default list
        $script:CoveredSoftware = Get-DefaultCoveredSoftware
        Save-CoveredSoftware
    }
}

function Save-CoveredSoftware {
    if ([string]::IsNullOrEmpty($script:CoveredSoftwarePath)) {
        Write-Warning "Covered software path is not set. Cannot save list."
        return $false
    }
    try {
        # Ensure settings directory exists before saving
        $settingsDir = [System.IO.Path]::GetDirectoryName($script:CoveredSoftwarePath)
        if (-not (Test-Path $settingsDir)) {
            New-Item -Path $settingsDir -ItemType Directory -Force | Out-Null
        }

        $script:CoveredSoftware | ConvertTo-Json -Depth 10 | Set-Content $script:CoveredSoftwarePath -Encoding UTF8
        Write-Host "Covered software list saved to $script:CoveredSoftwarePath"
        return $true
    } catch {
        Write-Warning "Could not save covered software list: $($_.Exception.Message)"
        return $false
    }
}

# --- General Recommendations Persistence ---

function Load-GeneralRecommendations {
    if (-not [string]::IsNullOrEmpty($script:GeneralRecommendationsPath) -and (Test-Path $script:GeneralRecommendationsPath)) {
        try {
            $json = Get-Content $script:GeneralRecommendationsPath -Raw | ConvertFrom-Json
            $script:GeneralRecommendations = @()
            foreach ($rec in $json) {
                $script:GeneralRecommendations += @{
                    Product = $rec.Product
                    Recommendations = $rec.Recommendations
                }
            }
            Write-Host "General recommendations loaded from $script:GeneralRecommendationsPath"
        } catch {
            Write-Warning "Could not load general recommendations: $($_.Exception.Message). Using empty list."
            $script:GeneralRecommendations = @()
        }
    } else {
        # Initialize with empty list
        $script:GeneralRecommendations = @()
        Save-GeneralRecommendations
    }
}

function Save-GeneralRecommendations {
    if ([string]::IsNullOrEmpty($script:GeneralRecommendationsPath)) {
        Write-Warning "General recommendations path is not set. Cannot save recommendations."
        return $false
    }
    try {
        # Ensure settings directory exists before saving
        $settingsDir = [System.IO.Path]::GetDirectoryName($script:GeneralRecommendationsPath)
        if (-not (Test-Path $settingsDir)) {
            New-Item -Path $settingsDir -ItemType Directory -Force | Out-Null
        }

        $script:GeneralRecommendations | ConvertTo-Json -Depth 10 | Set-Content $script:GeneralRecommendationsPath -Encoding UTF8
        Write-Host "General recommendations saved to $script:GeneralRecommendationsPath"
        return $true
    } catch {
        Write-Warning "Could not save general recommendations: $($_.Exception.Message)"
        return $false
    }
}

function Get-ModifierText {
    param(
        [bool]$AfterHours,
        [bool]$TicketGenerated,
        [bool]$ThirdParty
    )

    # Handle all combinations with proper English grammar
    if ($AfterHours -and $TicketGenerated -and $ThirdParty) {
        return " - After-hours ticket generated for 3rd party application"
    } elseif ($AfterHours -and $TicketGenerated) {
        return " - After-hours ticket generated"
    } elseif ($TicketGenerated -and $ThirdParty) {
        return " - Ticket generated for 3rd party application"
    } elseif ($AfterHours -and $ThirdParty) {
        return " - After-hours work required for 3rd party application, approval needed"
    } elseif ($TicketGenerated) {
        return " - Ticket generated"
    } elseif ($AfterHours) {
        return " - After-hours work required"
    } elseif ($ThirdParty) {
        return " - 3rd party application, approval needed"
    }

    return ""
}

function Test-IsCoveredSoftware {
    param(
        [string]$ProductName
    )

    if ($null -eq $script:CoveredSoftware -or $script:CoveredSoftware.Count -eq 0) {
        Load-CoveredSoftware
    }

    foreach ($item in $script:CoveredSoftware) {
        if ($item.Override) {
            continue  # Skip overridden items
        }

        if ($item.IsPattern) {
            if ($ProductName -like $item.Pattern) {
                return $true
            }
        } else {
            if ($ProductName -eq $item.Pattern) {
                return $true
            }
        }
    }

    return $false
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
