# VScanMagic-Core.ps1 - Config, settings, persistence, helpers
# Dot-sourced by VScanMagic-GUI.ps1 (do not run directly)

# --- Configuration ---
$script:Config = @{
    AppName = "VScanMagic v4"
    Version = "4.0.10"
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

    # Synthetic EPSS for items without EPSS (e.g. Network vulns) - used for ranking and MinEPSS filter
    SyntheticEPSSForNoEPSS = 0.1

    # Top N report: always include at least this many Network vulns regardless of score (ensures network findings are visible)
    MinNetworkVulnsInTopN = 2

    # AI batch: max items per API call (avoids rate limits); delay in seconds between chunks
    AIBatchChunkSize = 2
    AIBatchChunkDelaySeconds = 20
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
$script:ConnectWiseAutomateCredentialsPath = Join-Path $script:SettingsDirectory "ConnectWise-Automate-Credentials.json"
$script:CompanyFolderMapPath = Join-Path $script:SettingsDirectory "VScanMagic_CompanyFolderMap.json"
$script:ReportFolderHistoryPath = Join-Path $script:SettingsDirectory "VScanMagic_ReportFolderHistory.json"
$script:TemplatesPath = Join-Path $script:SettingsDirectory "VScanMagic_Templates.json"
$script:Templates = $null

# Report Filters and Output Options (modified via dialogs)
$script:FilterMinEPSS = 0
$script:FilterIncludeCritical = $true
$script:FilterIncludeHigh = $true
$script:FilterIncludeMedium = $false
$script:FilterIncludeLow = $false
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
    ReportsBasePath = ""  # Base folder for client output; when set, uses [Base]\[Folder]\[Year] - [QN]\
    HostnameReviewWindows11Threshold = 350  # VulnCount threshold: below = unselected, above = selected for Windows 11 O/S tabs
    HostnameReviewAutoLookupConnectSecure = $true  # When true, automatically lookup usernames from ConnectSecure when Hostname Review opens (CompanyId > 0)
    HostnameReviewAutoLookupLastPing = $true  # When true, lookup last ping time from ConnectSecure for Top N report and ticket instructions (CompanyId > 0)
    DownloadAutoResizeColumns = $true  # When true, auto-resize columns on downloaded Excel reports (excludes Company and Proposed Remediations (all) sheets)
    # AI API Keys (future: email, ticket notes, remediation, time estimate guidance)
    AIApiKeyCopilot = ""
    AIApiKeyChatGPT = ""
    AIApiKeyClaude = ""
    # Report Filters (persisted so Top N selection is remembered)
    FilterTopN = "10"
    FilterMinEPSS = 0
    FilterIncludeCritical = $true
    FilterIncludeHigh = $true
    FilterIncludeMedium = $false
    FilterIncludeLow = $false
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
    $script:ConnectSecureCredentialsPath = Join-Path $script:SettingsDirectory "ConnectSecure-Credentials.json"
    $script:ConnectSecureCompaniesCachePath = Join-Path $script:SettingsDirectory "ConnectSecure-Companies-Cache.json"
    $script:CompanyFolderMapPath = Join-Path $script:SettingsDirectory "VScanMagic_CompanyFolderMap.json"
    $script:ReportFolderHistoryPath = Join-Path $script:SettingsDirectory "VScanMagic_ReportFolderHistory.json"
    $script:ConnectWiseAutomateCredentialsPath = Join-Path $script:SettingsDirectory "ConnectWise-Automate-Credentials.json"
    $script:TemplatesPath = Join-Path $script:SettingsDirectory "VScanMagic_Templates.json"
    
    if (Ensure-SettingsDirectory -Path $script:SettingsDirectory) { Write-Host "Created settings directory: $script:SettingsDirectory" }
}

# Shared files: remediation rules, templates, etc. - safe to share between users
$script:BackupSharedFiles = @(
    "VScanMagic_RemediationRules.json",
    "VScanMagic_Templates.json",
    "VScanMagic_GeneralRecommendations.json",
    "VScanMagic_CoveredSoftware.json",
    "VScanMagic_CompanyFolderMap.json"
)
# User-specific files: credentials, paths, API keys - personal, not for sharing
$script:BackupUserFiles = @(
    "VScanMagic_Settings.json",
    "VScanMagic_ReportFolderHistory.json",
    "ConnectSecure-Credentials.json",
    "ConnectSecure-Companies-Cache.json",
    "ConnectWise-Automate-Credentials.json"
)

function Backup-Settings {
    param(
        [string]$OutputPath = $null,
        [ValidateSet("All", "Shared", "User")]
        [string]$Scope = "All"
    )
    if (-not (Test-Path $script:SettingsDirectory)) {
        return $null
    }
    # Ensure remediation rules are on disk before backup (defaults may be in-memory only)
    Load-RemediationRules | Out-Null
    Save-RemediationRules | Out-Null

    $settingsFiles = switch ($Scope) {
        "Shared" { $script:BackupSharedFiles }
        "User"   { $script:BackupUserFiles }
        default  { $script:BackupSharedFiles + $script:BackupUserFiles }
    }

    $timestamp = Get-Date -Format "yyyy-MM-dd_HHmmss"
    $scopeSuffix = if ($Scope -eq "All") { "" } else { "_$Scope" }
    $defaultName = "VScanMagic_Settings_Backup$scopeSuffix`_$timestamp.zip"
    if ([string]::IsNullOrWhiteSpace($OutputPath)) {
        $OutputPath = Join-Path ([Environment]::GetFolderPath("Desktop")) $defaultName
    } elseif ([System.IO.Directory]::Exists($OutputPath)) {
        $OutputPath = Join-Path $OutputPath $defaultName
    }
    $tempDir = Join-Path $env:TEMP "VScanMagic_Backup_$([Guid]::NewGuid().ToString('N').Substring(0,8))"
    try {
        New-Item -Path $tempDir -ItemType Directory -Force | Out-Null
        $copied = 0
        $copiedList = @()
        foreach ($f in $settingsFiles) {
            $src = Join-Path $script:SettingsDirectory $f
            if (Test-Path $src) {
                Copy-Item -Path $src -Destination (Join-Path $tempDir $f) -Force
                $copied++
                $copiedList += $f
            }
        }
        if ($copied -eq 0) {
            Write-Warning "No settings files found to backup."
            return $null
        }
        # Write manifest
        $manifest = @"
VScanMagic Settings Backup
Scope: $Scope
Timestamp: $(Get-Date -Format "yyyy-MM-dd HH:mm:ss")
Files included:
$($copiedList -join "`r`n")
"@
        $manifest | Set-Content -Path (Join-Path $tempDir "backup-manifest.txt") -Encoding UTF8
        if (Test-Path $OutputPath) { Remove-Item $OutputPath -Force }
        Compress-Archive -Path (Join-Path $tempDir "*") -DestinationPath $OutputPath -Force
        return $OutputPath
    } finally {
        if (Test-Path $tempDir) { Remove-Item $tempDir -Recurse -Force -ErrorAction SilentlyContinue }
    }
}

function Restore-Settings {
    param(
        [string]$BackupPath,
        [ValidateSet("All", "Shared", "User")]
        [string]$Scope = "All"
    )
    if (-not $BackupPath -or -not (Test-Path $BackupPath)) { return $false }
    $ext = [System.IO.Path]::GetExtension($BackupPath)
    if ($ext -ne ".zip") {
        Write-Warning "Backup file must be a .zip file."
        return $false
    }
    try {
        Ensure-SettingsDirectory -Path $script:SettingsDirectory | Out-Null
        $tempDir = Join-Path $env:TEMP "VScanMagic_Restore_$([Guid]::NewGuid().ToString('N').Substring(0,8))"
        Expand-Archive -Path $BackupPath -DestinationPath $tempDir -Force
        $allowedFiles = switch ($Scope) {
            "Shared" { $script:BackupSharedFiles }
            "User"   { $script:BackupUserFiles }
            default  { $null }  # All = no filter
        }
        $restored = 0
        Get-ChildItem -Path $tempDir -Filter "*.json" | ForEach-Object {
            if ($null -eq $allowedFiles -or $allowedFiles -contains $_.Name) {
                $dest = Join-Path $script:SettingsDirectory $_.Name
                Copy-Item -Path $_.FullName -Destination $dest -Force
                $restored++
            }
        }
        Remove-Item $tempDir -Recurse -Force -ErrorAction SilentlyContinue
        if ($restored -gt 0) {
            if ($Scope -eq "All" -or $Scope -eq "User") {
                Load-UserSettings | Out-Null
                Update-SettingsPaths
            }
            if ($Scope -eq "All" -or $Scope -eq "Shared") {
                Load-RemediationRules
                Load-CoveredSoftware
                Load-GeneralRecommendations
                Load-CompanyFolderMap
                Load-Templates
            }
            return $true
        }
        return $false
    } catch {
        Write-Warning "Restore failed: $($_.Exception.Message)"
        return $false
    }
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
        if ($null -ne $json.ReportsBasePath) { $script:UserSettings.ReportsBasePath = $json.ReportsBasePath } else { $script:UserSettings.ReportsBasePath = "" }
        if ($null -ne $json.HostnameReviewWindows11Threshold -and $json.HostnameReviewWindows11Threshold -ge 0) { $script:UserSettings.HostnameReviewWindows11Threshold = [int]$json.HostnameReviewWindows11Threshold } else { $script:UserSettings.HostnameReviewWindows11Threshold = 350 }
        if ($null -ne $json.HostnameReviewAutoLookupConnectSecure) { $script:UserSettings.HostnameReviewAutoLookupConnectSecure = [bool]$json.HostnameReviewAutoLookupConnectSecure } else { $script:UserSettings.HostnameReviewAutoLookupConnectSecure = $true }
        if ($null -ne $json.HostnameReviewAutoLookupLastPing) { $script:UserSettings.HostnameReviewAutoLookupLastPing = [bool]$json.HostnameReviewAutoLookupLastPing } else { $script:UserSettings.HostnameReviewAutoLookupLastPing = $true }
        if ($null -ne $json.DownloadAutoResizeColumns) { $script:UserSettings.DownloadAutoResizeColumns = [bool]$json.DownloadAutoResizeColumns } else { $script:UserSettings.DownloadAutoResizeColumns = $true }
        if ($null -ne $json.AIApiKeyCopilot) { $script:UserSettings.AIApiKeyCopilot = $json.AIApiKeyCopilot } else { $script:UserSettings.AIApiKeyCopilot = "" }
        if ($null -ne $json.AIApiKeyChatGPT) { $script:UserSettings.AIApiKeyChatGPT = $json.AIApiKeyChatGPT } else { $script:UserSettings.AIApiKeyChatGPT = "" }
        if ($null -ne $json.AIApiKeyClaude) { $script:UserSettings.AIApiKeyClaude = $json.AIApiKeyClaude } else { $script:UserSettings.AIApiKeyClaude = "" }
        if ($null -ne $json.FilterTopN -and $json.FilterTopN -in @('10','20','50','100','All')) { $script:UserSettings.FilterTopN = $json.FilterTopN } else { $script:UserSettings.FilterTopN = "10" }
        if ($null -ne $json.FilterMinEPSS) { $script:UserSettings.FilterMinEPSS = [double]$json.FilterMinEPSS } else { $script:UserSettings.FilterMinEPSS = 0 }
        if ($null -ne $json.FilterIncludeCritical) { $script:UserSettings.FilterIncludeCritical = [bool]$json.FilterIncludeCritical } else { $script:UserSettings.FilterIncludeCritical = $true }
        if ($null -ne $json.FilterIncludeHigh) { $script:UserSettings.FilterIncludeHigh = [bool]$json.FilterIncludeHigh } else { $script:UserSettings.FilterIncludeHigh = $true }
        if ($null -ne $json.FilterIncludeMedium) { $script:UserSettings.FilterIncludeMedium = [bool]$json.FilterIncludeMedium } else { $script:UserSettings.FilterIncludeMedium = $false }
        if ($null -ne $json.FilterIncludeLow) { $script:UserSettings.FilterIncludeLow = [bool]$json.FilterIncludeLow } else { $script:UserSettings.FilterIncludeLow = $false }
        Write-Host "User settings loaded from $script:SettingsPath"
    }
    # Sync filter script variables from UserSettings (used by Get-Top10Vulnerabilities etc.)
    $script:FilterTopN = $script:UserSettings.FilterTopN
    $script:FilterMinEPSS = $script:UserSettings.FilterMinEPSS
    $script:FilterIncludeCritical = $script:UserSettings.FilterIncludeCritical
    $script:FilterIncludeHigh = $script:UserSettings.FilterIncludeHigh
    $script:FilterIncludeMedium = $script:UserSettings.FilterIncludeMedium
    $script:FilterIncludeLow = $script:UserSettings.FilterIncludeLow
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
    # Trim all values - JSON/save can introduce trailing spaces, newlines, or BOM
    $baseUrl = if ($json.BaseUrl) { [string]$json.BaseUrl.Trim() -replace "`r`n|`r|`n", "" } else { "" }
    $tenantName = if ($json.TenantName) { [string]$json.TenantName.Trim() -replace "`r`n|`r|`n", "" } else { "" }
    $clientId = if ($json.ClientId) { [string]$json.ClientId.Trim() -replace "`r`n|`r|`n", "" } else { "" }
    $clientSecret = if ($json.ClientSecret) { [string]$json.ClientSecret.Trim() -replace "`r`n|`r|`n", "" } else { "" }
    return @{
        BaseUrl = $baseUrl
        TenantName = $tenantName
        ClientId = $clientId
        ClientSecret = $clientSecret
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

# --- ConnectWise Automate (username lookup for Hostname Review) ---

function Load-ConnectWiseAutomateCredentials {
    if ([string]::IsNullOrEmpty($script:ConnectWiseAutomateCredentialsPath)) { return $null }
    $json = Get-JsonFile -Path $script:ConnectWiseAutomateCredentialsPath
    if (-not $json) { return $null }
    return @{
        BaseUrl = if ($json.BaseUrl) { ($json.BaseUrl -replace '/$','') } else { "" }
        Username = if ($json.Username) { $json.Username } else { "" }
        Password = if ($json.Password) { $json.Password } else { "" }
        UseAutomateModule = if ($null -ne $json.UseAutomateModule) { [bool]$json.UseAutomateModule } else { $true }
    }
}

function Save-ConnectWiseAutomateCredentials {
    param([string]$BaseUrl, [string]$Username, [string]$Password, [bool]$UseAutomateModule = $true)
    if ([string]::IsNullOrEmpty($script:ConnectWiseAutomateCredentialsPath)) { return $false }
    $creds = @{ BaseUrl = $BaseUrl; Username = $Username; Password = $Password; UseAutomateModule = $UseAutomateModule }
    return Set-JsonFile -Path $script:ConnectWiseAutomateCredentialsPath -Object $creds
}

function Get-ConnectWiseUsernamesByHostname {
    param([string[]]$Hostnames)
    if (-not $Hostnames -or $Hostnames.Count -eq 0) { return @{} }
    $hostnames = $Hostnames | Where-Object { -not [string]::IsNullOrWhiteSpace($_) } | Select-Object -Unique
    if ($hostnames.Count -eq 0) { return @{} }
    $creds = Load-ConnectWiseAutomateCredentials
    if (-not $creds -or [string]::IsNullOrWhiteSpace($creds.BaseUrl)) {
        Write-Log "ConnectWise Automate not configured. Configure in Settings > ConnectWise Automate." -Level Warning
        return @{}
    }
    $result = @{}
    foreach ($h in $hostnames) { $result[$h] = "" }
    if ($creds.UseAutomateModule) {
        try {
            $mod = Get-Module -ListAvailable -Name AutomateAPI
            if ($mod) {
                Import-Module AutomateAPI -ErrorAction Stop
                $securePass = ConvertTo-SecureString $creds.Password -AsPlainText -Force
                $cred = New-Object PSCredential ($creds.Username, $securePass)
                Connect-AutomateApi -Server $creds.BaseUrl -Credential $cred -ErrorAction Stop | Out-Null
                foreach ($hostname in $hostnames) {
                    try {
                        $escaped = ($hostname -replace "'","''")
                        $computers = Get-AutomateComputer -Condition "(Name eq '$escaped')" -ErrorAction SilentlyContinue
                        if ($computers -and $computers.Count -gt 0) {
                            $comp = $computers[0]
                            $user = $comp.LastLoggedInUser; if (-not $user) { $user = $comp.LastLoggedInUserName }; if (-not $user) { $user = $comp.LoggedInUser }
                            if ($user -and -not [string]::IsNullOrWhiteSpace([string]$user)) { $result[$hostname] = [string]$user }
                        }
                    } catch { }
                }
                Disconnect-AutomateApi -ErrorAction SilentlyContinue | Out-Null
                return $result
            }
        } catch {
            Write-Log "AutomateAPI module lookup failed: $($_.Exception.Message)" -Level Warning
        }
    }
    try {
        $auth = [Convert]::ToBase64String([Text.Encoding]::ASCII.GetBytes("$($creds.Username):$($creds.Password)"))
        $headers = @{ "Authorization" = "Basic $auth"; "Content-Type" = "application/json" }
        $base = $creds.BaseUrl -replace '/$',''
        foreach ($hostname in $hostnames) {
            try {
                $escaped = ($hostname -replace "'","''")
                $uri = "$base/cwa/api/v1/Computers?`$filter=Name eq '$escaped'"
                $resp = Invoke-RestMethod -Uri $uri -Headers $headers -Method Get -TimeoutSec 15 -ErrorAction Stop
                if ($resp -and $resp.value -and $resp.value.Count -gt 0) {
                    $comp = $resp.value[0]
                    $user = $comp.LastLoggedInUser; if (-not $user) { $user = $comp.LastLoggedInUserName }; if (-not $user) { $user = $comp.LoggedInUser }
                    if ($user -and -not [string]::IsNullOrWhiteSpace([string]$user)) { $result[$hostname] = [string]$user }
                }
            } catch { }
        }
    } catch {
        Write-Log "ConnectWise Automate REST API failed: $($_.Exception.Message)" -Level Warning
    }
    return $result
}

# --- Company Folder Mapping (for structured output paths) ---

$script:CompanyFolderMap = @{}

function Load-CompanyFolderMap {
    # Try to use MemberberryIntegration module if available
    if (Get-Command Get-VScanMagicClientData -ErrorAction SilentlyContinue) {
        try {
            $clientData = Get-VScanMagicClientData
            if ($clientData -and $clientData.CompanyFolderMap) {
                $script:CompanyFolderMap = $clientData.CompanyFolderMap
                $storageInfo = Get-VScanMagicStorageInfo
                Write-Host "Company folder mappings loaded from shared storage: $($storageInfo.DataPath)" -ForegroundColor Green
                return
            }
        } catch {
            Write-Warning "Could not load company folder mappings from shared storage: $($_.Exception.Message). Falling back to local storage."
        }
    }
    
    # Fallback to local storage (backward compatibility)
    if ([string]::IsNullOrEmpty($script:CompanyFolderMapPath)) { $script:CompanyFolderMap = @{}; return }
    $json = Get-JsonFile -Path $script:CompanyFolderMapPath
    $script:CompanyFolderMap = @{}
    if ($json -and $json.PSObject.Properties) {
        foreach ($p in $json.PSObject.Properties) {
            $script:CompanyFolderMap[$p.Name] = $p.Value
        }
    }
}

function Save-CompanyFolderMap {
    # Try to use MemberberryIntegration module if available
    if (Get-Command Save-VScanMagicClientData -ErrorAction SilentlyContinue) {
        try {
            # Load existing client data to preserve RMIT Plus settings
            $existingData = Get-VScanMagicClientData
            $rmitPlusSettings = if ($existingData -and $existingData.RMITPlusSettings) { $existingData.RMITPlusSettings } else { @{} }
            
            if (Save-VScanMagicClientData -RMITPlusSettings $rmitPlusSettings -CompanyFolderMap $script:CompanyFolderMap) {
                $storageInfo = Get-VScanMagicStorageInfo
                Write-Host "Company folder mappings saved to shared storage: $($storageInfo.DataPath)" -ForegroundColor Green
                return $true
            }
        } catch {
            Write-Warning "Could not save company folder mappings to shared storage: $($_.Exception.Message). Falling back to local storage."
        }
    }
    
    # Fallback to local storage (backward compatibility)
    if ([string]::IsNullOrEmpty($script:CompanyFolderMapPath)) { return $false }
    return Set-JsonFile -Path $script:CompanyFolderMapPath -Object $script:CompanyFolderMap -Depth 2
}

# --- Report Folder History (persistent list of processed output folders) ---

$script:ReportFolderHistoryMaxEntries = 100

function Get-ReportFolderHistory {
    if ([string]::IsNullOrEmpty($script:ReportFolderHistoryPath)) { return @() }
    $json = Get-JsonFile -Path $script:ReportFolderHistoryPath
    if (-not $json -or -not $json.Entries) { return @() }
    return @($json.Entries | ForEach-Object { [PSCustomObject]$_ })
}

function Add-ToReportFolderHistory {
    param(
        [string]$CompanyName,
        [string]$OutputPath
    )
    if ([string]::IsNullOrWhiteSpace($OutputPath) -or -not (Test-Path $OutputPath)) { return }
    $pathNorm = [System.IO.Path]::GetFullPath($OutputPath)
    # Don't add one-off paths (outside Reports Base Path) to history
    $base = $script:UserSettings.ReportsBasePath
    if ($base -and (Test-Path $base)) {
        $baseNorm = [System.IO.Path]::GetFullPath($base.Trim())
        if (-not $pathNorm.StartsWith($baseNorm, [StringComparison]::OrdinalIgnoreCase)) {
            return
        }
    }
    $existing = @(Get-ReportFolderHistory | Where-Object { $_.OutputPath -ne $pathNorm })
    $now = Get-Date -Format "yyyy-MM-dd HH:mm"
    $newEntry = [PSCustomObject]@{ CompanyName = $CompanyName; OutputPath = $pathNorm; ProcessedAt = $now }
    $maxRest = [Math]::Max(0, $script:ReportFolderHistoryMaxEntries - 1)
    $takeCount = [Math]::Min($existing.Count, $maxRest)
    $rest = if ($takeCount -le 0) { @() } else { $existing[0..($takeCount - 1)] }
    $entries = @($newEntry) + @($rest)
    $obj = @{ Entries = @($entries) }
    Set-JsonFile -Path $script:ReportFolderHistoryPath -Object $obj -Depth 3 | Out-Null
}

function Get-QuarterFromDate {
    param([string]$ScanDate)
    if ([string]::IsNullOrWhiteSpace($ScanDate)) { return (Get-Date).Year.ToString() + " - Q" + [Math]::Ceiling((Get-Date).Month / 3) }
    try {
        $d = [DateTime]::Parse($ScanDate)
        $q = [Math]::Ceiling($d.Month / 3)
        return "$($d.Year) - Q$q"
    } catch {
        return (Get-Date).Year.ToString() + " - Q" + [Math]::Ceiling((Get-Date).Month / 3)
    }
}

function Resolve-VulnerabilityScansSubpath {
    param([string]$FolderName)
    if ([string]::IsNullOrWhiteSpace($FolderName)) { return $FolderName }
    if ($FolderName -like "*Vulnerability Scans*") { return $FolderName }
    return Join-Path $FolderName "Network Documentation\Vulnerability Scans"
}

function Get-ReportsPathPartial {
    <#
    .SYNOPSIS
    Returns a partial path for ticket instructions: company name (if provided) + Network Documentation onwards.
    Technicians use this to locate the reports folder (e.g. ClientName\Network Documentation\Vulnerability Scans\2026 - Q1).
    #>
    param(
        [string]$FullOutputPath,
        [string]$CompanyName = $null
    )
    if ([string]::IsNullOrWhiteSpace($FullOutputPath)) { return $null }
    $path = [System.IO.Path]::GetFullPath($FullOutputPath.Trim())
    $ndIdx = $path.IndexOf("Network Documentation", [StringComparison]::OrdinalIgnoreCase)
    $partial = $null
    if ($ndIdx -ge 0) {
        $partial = $path.Substring($ndIdx).Replace([System.IO.Path]::DirectorySeparatorChar, '\')
    } else {
        # Fallback: use last path segment (year-quarter folder, e.g. 2026 - Q1)
        $parts = $path.Split([System.IO.Path]::DirectorySeparatorChar, [StringSplitOptions]::RemoveEmptyEntries)
        if ($parts.Count -ge 1) { $partial = $parts[-1] }
    }
    if ([string]::IsNullOrWhiteSpace($partial)) { return $null }
    if (-not [string]::IsNullOrWhiteSpace($CompanyName)) {
        $partial = "$($CompanyName.Trim())\$partial"
    }
    return $partial
}

function Resolve-QuarterFolderName {
    param(
        [string]$ClientPath,
        [string]$ScanDate
    )
    $yearQuarter = Get-QuarterFromDate -ScanDate $ScanDate
    $outPath = Join-Path $ClientPath $yearQuarter
    if (-not (Test-Path $outPath)) { return $yearQuarter }
    try {
        $d = [DateTime]::Parse($ScanDate)
        $q = [Math]::Ceiling($d.Month / 3)
        $dateStr = $d.ToString("yyyy-MM-dd")
        return "$($d.Year) - Q$q $dateStr"
    } catch {
        $d = Get-Date
        $q = [Math]::Ceiling($d.Month / 3)
        return "$($d.Year) - Q$q $($d.ToString('yyyy-MM-dd'))"
    }
}

function Resolve-ClientOutputPath {
    param(
        [int]$CompanyId,
        [string]$CompanyName,
        [string]$ScanDate,
        [switch]$ForceManual,
        [string]$FallbackPath = ""  # When set, use instead of LastOutputDirectory (e.g. current Output Directory from form)
    )
    # Manual path (folder picker) only when: (1) Re-select folder is checked, or (2) mapping/auto-match cannot determine path.
    $base = $script:UserSettings.ReportsBasePath
    if ([string]::IsNullOrWhiteSpace($base) -or -not (Test-Path $base)) {
        $fallback = if ($FallbackPath -and (Test-Path $FallbackPath)) { $FallbackPath } else { $script:UserSettings.LastOutputDirectory }
        if ([string]::IsNullOrWhiteSpace($fallback) -or -not (Test-Path $fallback)) {
            $fallback = [Environment]::GetFolderPath("Desktop")
        }
        Write-Log "Using flat output path (Reports Base Path not set or invalid): $fallback" -Level Info
        return $fallback
    }
    $base = [System.IO.Path]::GetFullPath($base.Trim())
    $key = "$CompanyId"
    if ($script:CompanyFolderMap.Count -eq 0) { Load-CompanyFolderMap }
    $folderName = $null
    if (-not $ForceManual -and $script:CompanyFolderMap.ContainsKey($key)) {
        $folderName = $script:CompanyFolderMap[$key]
        $folderName = Resolve-VulnerabilityScansSubpath -FolderName $folderName
        if ($script:CompanyFolderMap[$key] -ne $folderName) {
            $script:CompanyFolderMap[$key] = $folderName
            Save-CompanyFolderMap | Out-Null
        }
        $clientPath = Join-Path $base $folderName
        if (Test-Path $clientPath) {
            $quarterFolder = Resolve-QuarterFolderName -ClientPath $clientPath -ScanDate $ScanDate
            $outPath = Join-Path $clientPath $quarterFolder
            if (-not (Test-Path $outPath)) { New-Item -ItemType Directory -Path $outPath -Force | Out-Null }
            $miscPath = Join-Path $outPath "Misc"
            if (-not (Test-Path $miscPath)) { New-Item -ItemType Directory -Path $miscPath -Force | Out-Null }
            Write-Log "Output path (mapped): $outPath" -Level Info
            return $outPath
        }
    }
    if (-not $ForceManual) {
        $subfolders = @()
        try {
            $subfolders = @(Get-ChildItem -Path $base -Directory -ErrorAction SilentlyContinue | ForEach-Object { $_.Name })
        } catch { }
        $norm = $CompanyName.Trim() -replace '\s*(Inc|LLC|Corp|ltd)\.?\s*$', '' -replace '\s+', ' '
        $normLower = $norm.ToLowerInvariant()
        $bestMatch = $null
        $bestScore = -1
        foreach ($f in $subfolders) {
            $fLower = $f.ToLowerInvariant()
            $score = 0
            if ($normLower -eq $fLower) { $score = 100 }
            elseif ($normLower.Length -gt 0 -and ($fLower -like "*$normLower*" -or $normLower -like "*$fLower*")) { $score = 50 }
            elseif ($normLower.Length -ge 3 -and $fLower.Length -ge 3) {
                $nPre = $normLower.Substring(0, [Math]::Min(5, $normLower.Length))
                $fPre = $fLower.Substring(0, [Math]::Min(5, $fLower.Length))
                if ($fLower.StartsWith($nPre) -or $normLower.StartsWith($fPre)) { $score = 25 }
            }
            if ($score -gt $bestScore) { $bestScore = $score; $bestMatch = $f }
            elseif ($score -eq $bestScore -and $bestMatch) { $bestMatch = $null }
        }
        if ($bestMatch) {
            $folderName = Resolve-VulnerabilityScansSubpath -FolderName $bestMatch
            $script:CompanyFolderMap[$key] = $folderName
            Save-CompanyFolderMap | Out-Null
            $clientPath = Join-Path $base $folderName
            $quarterFolder = Resolve-QuarterFolderName -ClientPath $clientPath -ScanDate $ScanDate
            $outPath = Join-Path $clientPath $quarterFolder
            if (-not (Test-Path $outPath)) { New-Item -ItemType Directory -Path $outPath -Force | Out-Null }
            $miscPath = Join-Path $outPath "Misc"
            if (-not (Test-Path $miscPath)) { New-Item -ItemType Directory -Path $miscPath -Force | Out-Null }
            Write-Log "Output path (auto-matched): $outPath" -Level Info
            return $outPath
        }
    }
    Add-Type -AssemblyName System.Windows.Forms -ErrorAction SilentlyContinue
    $dialog = New-Object System.Windows.Forms.FolderBrowserDialog
    $dialog.Description = "Select the vulnerability scan directory where the year-quarter folder (e.g. 2026 - Q1) will be created. Choose under the base path or any other location (one-off)."
    $dialog.SelectedPath = if ($FallbackPath -and (Test-Path $FallbackPath)) { $FallbackPath } else { $base }
    $dialog.ShowNewFolderButton = $true
    if ($dialog.ShowDialog() -ne [System.Windows.Forms.DialogResult]::OK) { return $null }
    $selected = $dialog.SelectedPath
    $selectedFull = [System.IO.Path]::GetFullPath($selected)
    $baseTrimmed = $base.TrimEnd([char]'\', [char]'/')
    $isInsideBase = $selectedFull.StartsWith($baseTrimmed, [StringComparison]::OrdinalIgnoreCase) -and
        ($selectedFull.Length -eq $baseTrimmed.Length -or $selectedFull[$baseTrimmed.Length] -in '\', '/')
    if (-not $isInsideBase) {
        # User selected a path outside base (one-off) - use it directly, don't map to base
        if (-not (Test-Path $selectedFull)) { New-Item -ItemType Directory -Path $selectedFull -Force | Out-Null }
        $miscPath = Join-Path $selectedFull "Misc"
        if (-not (Test-Path $miscPath)) { New-Item -ItemType Directory -Path $miscPath -Force | Out-Null }
        Write-Log "Output path (selected, one-off): $selectedFull" -Level Info
        return $selectedFull
    }
    $rel = $selectedFull.Substring($baseTrimmed.Length).TrimStart([char]'\', [char]'/')
    $folderName = if ([string]::IsNullOrWhiteSpace($rel)) { [System.IO.Path]::GetFileName($selectedFull) } else { $rel }
    if ([string]::IsNullOrWhiteSpace($folderName)) { $folderName = "Client" }
    $folderName = Resolve-VulnerabilityScansSubpath -FolderName $folderName
    $script:CompanyFolderMap[$key] = $folderName
    Save-CompanyFolderMap | Out-Null
    $clientPath = Join-Path $base $folderName
    $quarterFolder = Resolve-QuarterFolderName -ClientPath $clientPath -ScanDate $ScanDate
    $outPath = Join-Path $clientPath $quarterFolder
    if (-not (Test-Path $outPath)) { New-Item -ItemType Directory -Path $outPath -Force | Out-Null }
    $miscPath = Join-Path $outPath "Misc"
    if (-not (Test-Path $miscPath)) { New-Item -ItemType Directory -Path $miscPath -Force | Out-Null }
    Write-Log "Output path (selected): $outPath" -Level Info
    return $outPath
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
            Pattern = "*ripple20*"
            WordText = "Ripple20 vulnerabilities affect the Treck TCP/IP stack in millions of IoT and embedded devices. Remediation requires firmware updates from the device manufacturer. Identify affected devices via network scanning; check vendor security advisories (e.g., Intel, HP, Cisco, Schneider Electric). Where patching is not immediately possible, implement network segmentation and firewall rules to restrict access to vulnerable devices."
            TicketText = "- Ripple20 affects Treck TCP/IP stack in IoT/embedded devices`r`n  - Update firmware from device manufacturer`r`n  - Check vendor security advisories for patches`r`n  - Where patching not possible: network segmentation and firewall restrictions"
            IsDefault = $false
        },
        @{
            Pattern = "*rippl20*"
            WordText = "Ripple20 vulnerabilities affect the Treck TCP/IP stack in millions of IoT and embedded devices. Remediation requires firmware updates from the device manufacturer. Identify affected devices via network scanning; check vendor security advisories (e.g., Intel, HP, Cisco, Schneider Electric). Where patching is not immediately possible, implement network segmentation and firewall rules to restrict access to vulnerable devices."
            TicketText = "- Ripple20 affects Treck TCP/IP stack in IoT/embedded devices`r`n  - Update firmware from device manufacturer`r`n  - Check vendor security advisories for patches`r`n  - Where patching not possible: network segmentation and firewall restrictions"
            IsDefault = $false
        },
        @{
            Pattern = "*Ripple20*"
            WordText = "Ripple20 vulnerabilities affect the Treck TCP/IP stack in millions of IoT and embedded devices. Remediation requires firmware updates from the device manufacturer. Identify affected devices via network scanning; check vendor security advisories (e.g., Intel, HP, Cisco, Schneider Electric). Where patching is not immediately possible, implement network segmentation and firewall rules to restrict access to vulnerable devices."
            TicketText = "- Ripple20 affects Treck TCP/IP stack in IoT/embedded devices`r`n  - Update firmware from device manufacturer`r`n  - Check vendor security advisories for patches`r`n  - Where patching not possible: network segmentation and firewall restrictions"
            IsDefault = $false
        },
        @{
            Pattern = "*Microsoft Teams*"
            WordText = "Microsoft Teams can be updated via RMM script deployed through ConnectWise Automate. This can be remediated by cleaning up unused user profile installed versions using: Select Scripts > RR - Custom > RR - Custom - R-Security Remediation > R-Security - Teams Classic Cleanup Remediation in RMM."
            TicketText = "- Update via RMM script deployed through ConnectWise Automate`r`n  - Can be remediated by cleaning up unused user profile installed versions`r`n  - Script path: Select Scripts > RR - Custom > RR - Custom - R-Security Remediation > R-Security - Teams Classic Cleanup Remediation in RMM"
            IsDefault = $false
        },
        @{
            Pattern = "*VMware*"
            WordText = "VMware ESXi and vSphere updates are released by Broadcom. Remediation requires downloading patches from the Broadcom Support Portal and applying via vSphere Lifecycle Manager (recommended) or by manually installing the offline bundle or ISO. Host reboot is required; migrate or shut down VMs before patching."
            TicketText = "- Download patches from Broadcom Support Portal`r`n  - Apply via vSphere Lifecycle Manager (recommended) or manual offline bundle/ISO`r`n  - Host reboot required; migrate or shut down VMs before patching"
            IsDefault = $false
        },
        @{
            Pattern = "*vSphere*"
            WordText = "VMware ESXi and vSphere updates are released by Broadcom. Remediation requires downloading patches from the Broadcom Support Portal and applying via vSphere Lifecycle Manager (recommended) or by manually installing the offline bundle or ISO. Host reboot is required; migrate or shut down VMs before patching."
            TicketText = "- Download patches from Broadcom Support Portal`r`n  - Apply via vSphere Lifecycle Manager (recommended) or manual offline bundle/ISO`r`n  - Host reboot required; migrate or shut down VMs before patching"
            IsDefault = $false
        },
        @{
            Pattern = "*ESXi*"
            WordText = "VMware ESXi and vSphere updates are released by Broadcom. Remediation requires downloading patches from the Broadcom Support Portal and applying via vSphere Lifecycle Manager (recommended) or by manually installing the offline bundle or ISO. Host reboot is required; migrate or shut down VMs before patching."
            TicketText = "- Download patches from Broadcom Support Portal`r`n  - Apply via vSphere Lifecycle Manager (recommended) or manual offline bundle/ISO`r`n  - Host reboot required; migrate or shut down VMs before patching"
            IsDefault = $false
        },
        @{
            Pattern = "*vCenter*"
            WordText = "VMware vCenter Server updates are released by Broadcom. Remediation requires downloading the patch ISO from the Broadcom Support Portal and applying via the vCenter Lifecycle Manager plug-in, GUI installer, or Virtual Appliance Management Interface (VAMI). Plan for maintenance; vCenter services restart during patching."
            TicketText = "- Download patch ISO from Broadcom Support Portal`r`n  - Apply via vCenter Lifecycle Manager, GUI installer, or VAMI`r`n  - Plan for maintenance; vCenter services restart during patching"
            IsDefault = $false
        },
        @{
            Pattern = "*FortiGate*"
            WordText = "FortiGate firewall firmware updates are available from the Fortinet Support Portal. For managed devices, use FortiCloud Fabric Manager (Management > Firmware) to download, backup, and install firmware. If automatic upgrade fails, apply manually via the web interface (System > Firmware). Backup the configuration before updating. Plan a maintenance window; the device reboots during the update."
            TicketText = "- Download firmware from Fortinet Support Portal`r`n  - Prefer FortiCloud Fabric Manager (Management > Firmware) for managed devices`r`n  - If automatic upgrade fails, apply manually via web interface (System > Firmware)`r`n  - Backup configuration before updating`r`n  - Plan maintenance window; device reboots during update"
            IsDefault = $false
        },
        @{
            Pattern = "*Fortinet*"
            WordText = "FortiGate firewall firmware updates are available from the Fortinet Support Portal. For managed devices, use FortiCloud Fabric Manager (Management > Firmware) to download, backup, and install firmware. If automatic upgrade fails, apply manually via the web interface (System > Firmware). Backup the configuration before updating. Plan a maintenance window; the device reboots during the update."
            TicketText = "- Download firmware from Fortinet Support Portal`r`n  - Prefer FortiCloud Fabric Manager (Management > Firmware) for managed devices`r`n  - If automatic upgrade fails, apply manually via web interface (System > Firmware)`r`n  - Backup configuration before updating`r`n  - Plan maintenance window; device reboots during update"
            IsDefault = $false
        },
        @{
            Pattern = "*SonicWall*"
            WordText = "SonicWall firewall firmware updates are available from the MySonicWall portal. Download the firmware for your appliance model, then apply via the management interface (System > Settings > Firmware). Backup the configuration before updating. Plan a maintenance window; the device reboots during the update."
            TicketText = "- Download firmware from MySonicWall portal`r`n  - Apply via management interface (System > Settings > Firmware)`r`n  - Backup configuration before updating`r`n  - Plan maintenance window; device reboots during update"
            IsDefault = $false
        },
        @{
            Pattern = "*Microsoft .NET Framework*"
            WordText = "Legacy .NET: .NET Framework (1.0 through 4.8) cannot be upgraded in-place to modern .NET (5, 6, 7, 8, 9). There is no direct upgrade path; attempting to switch runtimes without migration will likely break the application. Modern .NET is a different runtime with different APIs - some Framework APIs do not exist or behave differently. The older the Framework version (e.g. 3.5, 4.0), the more painful the migration. Migration is a real project: retarget the application to .NET 8 (LTS) or later, update deprecated or incompatible APIs, and test thoroughly. For MSP/client conversations: the main reason to migrate is that Framework versions go out of support and stop receiving security patches - not because end users will notice any difference."
            TicketText = "- .NET Framework cannot be upgraded to modern .NET (5, 6, 7, 8, 9) without migration`r`n  - No direct upgrade path; switching runtimes without migration will likely break the app`r`n  - Modern .NET is a different runtime; many APIs differ or are missing`r`n  - Older Framework versions (3.5, 4.0) are more painful to migrate than 4.x`r`n  - Migration is a real project: retarget to .NET 8 (LTS), update APIs, test thoroughly`r`n  - Main driver: security patches (Framework goes out of support) - not user experience"
            IsDefault = $false
        },
        @{
            Pattern = "*Microsoft .NET Core 2.*"
            WordText = "Legacy .NET: .NET Core 2.x is out of support. Migration to modern .NET (5/6/7/8) is required. .NET Framework apps need a full migration project; apps already on .NET Core can usually retarget with minimal code changes. The main driver is security: out-of-support versions stop receiving patches."
            TicketText = "- Legacy .NET Core 2.x: out of support`r`n  - Retarget to .NET 8 (LTS) or later`r`n  - .NET Core apps: usually minimal code changes`r`n  - Main driver: security patches; out-of-support versions receive none"
            IsDefault = $false
        },
        @{
            Pattern = "*Microsoft .NET Core 3.*"
            WordText = "Legacy .NET: .NET Core 3.x (including 3.1 LTS) has reached or is nearing end of support. Migration to modern .NET 8 (LTS) is recommended. Within the modern .NET lineage, retargeting is usually a project file edit and minimal code changes. The main driver is security: out-of-support versions stop receiving patches."
            TicketText = "- Legacy .NET Core 3.x: end of support or nearing`r`n  - Retarget to .NET 8 (LTS) with minimal code changes`r`n  - Main driver: security patches; out-of-support versions receive none`r`n  - LTS versions (6, 8, 10) get 3 years support"
            IsDefault = $false
        },
        @{
            Pattern = "*Microsoft .NET Core*"
            WordText = "Legacy .NET: .NET Core (1.x, 2.x, 3.x) is the predecessor to modern .NET (5+). Core 2.x and 3.x are out of support. Migration to .NET 8 (LTS) is recommended. .NET Framework requires a full migration project; apps on Core can usually retarget with minimal changes. The main driver is security patches."
            TicketText = "- Legacy .NET Core: out of support or nearing`r`n  - Retarget to .NET 8 (LTS) or later`r`n  - Main driver: security patches; out-of-support versions receive none`r`n  - .NET Framework migration is a real project; Core retarget is usually low friction"
            IsDefault = $false
        },
        @{
            Pattern = "*Microsoft .NET Runtime 5.*"
            WordText = "Modern .NET: .NET 5 is out of support (18-month lifecycle). Retarget to .NET 8 (LTS) with minimal code changes. Within modern .NET, upgrades are largely drop-in. The main driver is support: LTS versions get 3 years; non-LTS get 18 months. End users typically notice nothing; value is security patches and lower infrastructure costs."
            TicketText = "- Modern .NET 5: out of support`r`n  - Retarget to .NET 8 (LTS) with minimal code changes`r`n  - Main driver: security patches`r`n  - LTS versions get 3 years support; non-LTS get 18 months"
            IsDefault = $false
        },
        @{
            Pattern = "*Microsoft .NET Runtime 6.*"
            WordText = "Modern .NET: .NET 6 (LTS) maintains strong backwards compatibility. Retargeting to .NET 8 (LTS) is usually a project file edit and minimal code changes. The main reason to upgrade is support lifecycle: LTS versions get 3 years; staying current ensures security patches. End users typically notice nothing; value is security patches and lower infrastructure costs."
            TicketText = "- Modern .NET 6: largely drop-in upgradeable to .NET 8`r`n  - Retarget to .NET 8 (LTS) with minimal code changes`r`n  - Main driver: support lifecycle and security patches`r`n  - LTS versions get 3 years support"
            IsDefault = $false
        },
        @{
            Pattern = "*Microsoft .NET Runtime 7.*"
            WordText = "Modern .NET: .NET 7 (non-LTS) has an 18-month support window. Consider retargeting to .NET 8 (LTS) for 3-year support. Upgrades within modern .NET are largely drop-in. The main driver is support lifecycle and security patches. End users typically notice nothing."
            TicketText = "- Modern .NET 7: non-LTS (18 months support)`r`n  - Consider retarget to .NET 8 (LTS) for 3-year support`r`n  - Largely drop-in upgradeable`r`n  - Main driver: support lifecycle and security patches"
            IsDefault = $false
        },
        @{
            Pattern = "*Microsoft .NET Runtime 8.*"
            WordText = "Modern .NET: .NET 8 (LTS) is the current long-term support release with 3 years of support. Keep current with patch updates. If a newer LTS (e.g. .NET 10) is available, retargeting is usually minimal. End users typically notice nothing; value is security patches and lower infrastructure costs."
            TicketText = "- Modern .NET 8 (LTS): keep current with patch updates`r`n  - 3 years support`r`n  - Retarget to newer LTS when available with minimal changes`r`n  - Main driver: security patches"
            IsDefault = $false
        },
        @{
            Pattern = "*Microsoft .NET Runtime 9.*"
            WordText = "Modern .NET: .NET 9 (non-LTS) has an 18-month support window. For long-term support, consider .NET 8 (LTS) or plan for .NET 10 (LTS). Upgrades within modern .NET are largely drop-in. The main driver is support lifecycle and security patches."
            TicketText = "- Modern .NET 9: non-LTS (18 months support)`r`n  - Consider .NET 8 (LTS) or .NET 10 (LTS) for longer support`r`n  - Largely drop-in upgradeable`r`n  - Main driver: support lifecycle and security patches"
            IsDefault = $false
        },
        @{
            Pattern = "*Microsoft .NET Runtime*"
            WordText = "Modern .NET: .NET 5+ (Runtime) maintains strong backwards compatibility. Retargeting between versions is usually a project file edit and minimal code changes. The main reason to upgrade is support: LTS versions (6, 8, 10) get 3 years; non-LTS get 18 months. End users typically notice nothing; for MSP/client conversations, the value is security patches and lower infrastructure costs - not user-facing improvements."
            TicketText = "- Modern .NET: largely drop-in upgradeable between versions`r`n  - Retarget to .NET 8 (LTS) with minimal code changes`r`n  - Main driver: security patches (older versions go out of support) - not user experience`r`n  - LTS versions get 3 years support; non-LTS get 18 months"
            IsDefault = $false
        },
        @{
            Pattern = "*Microsoft .NET 5.*"
            WordText = "Modern .NET: .NET 5 is out of support (18-month lifecycle). Retarget to .NET 8 (LTS) with minimal code changes. Within modern .NET, upgrades are largely drop-in. The main driver is support and security patches."
            TicketText = "- Modern .NET 5: out of support`r`n  - Retarget to .NET 8 (LTS) with minimal code changes`r`n  - Main driver: security patches"
            IsDefault = $false
        },
        @{
            Pattern = "*Microsoft .NET 6.*"
            WordText = "Modern .NET: .NET 6 (LTS) is largely drop-in upgradeable to .NET 8 (LTS). Retargeting is usually a project file edit and minimal code changes. The main driver is support lifecycle and security patches. End users typically notice nothing."
            TicketText = "- Modern .NET 6: largely drop-in upgradeable to .NET 8`r`n  - Retarget to .NET 8 (LTS) with minimal code changes`r`n  - Main driver: support lifecycle and security patches"
            IsDefault = $false
        },
        @{
            Pattern = "*Microsoft .NET 7.*"
            WordText = "Modern .NET: .NET 7 (non-LTS) has an 18-month support window. Consider retargeting to .NET 8 (LTS) for 3-year support. Upgrades within modern .NET are largely drop-in."
            TicketText = "- Modern .NET 7: non-LTS`r`n  - Consider retarget to .NET 8 (LTS) for 3-year support`r`n  - Largely drop-in upgradeable"
            IsDefault = $false
        },
        @{
            Pattern = "*Microsoft .NET 8.*"
            WordText = "Modern .NET: .NET 8 (LTS) is the current long-term support release with 3 years of support. Keep current with patch updates. End users typically notice nothing; value is security patches and lower infrastructure costs."
            TicketText = "- Modern .NET 8 (LTS): keep current with patch updates`r`n  - 3 years support`r`n  - Main driver: security patches"
            IsDefault = $false
        },
        @{
            Pattern = "*Microsoft .NET 9.*"
            WordText = "Modern .NET: .NET 9 (non-LTS) has an 18-month support window. For long-term support, consider .NET 8 (LTS) or plan for .NET 10 (LTS). Upgrades within modern .NET are largely drop-in."
            TicketText = "- Modern .NET 9: non-LTS (18 months support)`r`n  - Consider .NET 8 or .NET 10 (LTS) for longer support`r`n  - Largely drop-in upgradeable"
            IsDefault = $false
        },
        @{
            Pattern = "*"
            WordText = "Determine what the device or software is (use Product/OS and affected hosts). Review the manufacturer's security advisories and vulnerability data for patches or firmware updates. Consider configuration mitigations (e.g. network segmentation, hardening) where patching is not immediately possible. If available via ConnectWise Automate/RMM, deploy updates via patch management or scripts; otherwise, manual updates may be required."
            TicketText = "- Determine device/software identity (Product/OS, affected hosts)`r`n  - Review manufacturer security advisories and vulnerability data`r`n  - Check for firmware updates or patches`r`n  - Consider configuration mitigations where patching not possible`r`n  - Deploy via ConnectWise Automate/RMM if available; otherwise manual updates"
            IsDefault = $true
        }
    )
}

function Load-RemediationRules {
    $script:CachedRemediationRulesForGuidance = $null  # invalidate cache
    
    # Try to use MemberberryIntegration module if available
    if (Get-Command Get-VScanMagicRemediationRules -ErrorAction SilentlyContinue) {
        try {
            $sharedRules = Get-VScanMagicRemediationRules
            if ($sharedRules -and $sharedRules.Count -gt 0) {
                $script:RemediationRules = @()
                foreach ($rule in $sharedRules) {
                    $script:RemediationRules += @{ Pattern = $rule.Pattern; WordText = $rule.WordText; TicketText = $rule.TicketText; IsDefault = $rule.IsDefault }
                }
                # Merge in any new default rules not already in shared file
                $defaults = Get-DefaultRemediationRules
                $existingPatterns = @{}
                foreach ($r in $script:RemediationRules) { $existingPatterns[$r.Pattern] = $true }
                $added = 0
                foreach ($d in $defaults) {
                    if (-not $existingPatterns.ContainsKey($d.Pattern)) {
                        $script:RemediationRules += @{ Pattern = $d.Pattern; WordText = $d.WordText; TicketText = $d.TicketText; IsDefault = $d.IsDefault }
                        $existingPatterns[$d.Pattern] = $true
                        $added++
                    }
                }
                if ($added -gt 0) {
                    Save-RemediationRules | Out-Null
                    Write-Host "Added $added new default remediation rule(s) to shared storage"
                }
                $storageInfo = Get-VScanMagicStorageInfo
                Write-Host "Remediation rules loaded from shared storage: $($storageInfo.DataPath)" -ForegroundColor Green
                return
            }
        } catch {
            Write-Warning "Could not load remediation rules from shared storage: $($_.Exception.Message). Falling back to local storage."
        }
    }
    
    # Fallback to local storage (backward compatibility)
    $json = Get-JsonFile -Path $script:RemediationRulesPath
    if ($json -and $json.Count -gt 0) {
        $script:RemediationRules = @()
        foreach ($rule in $json) {
            $script:RemediationRules += @{ Pattern = $rule.Pattern; WordText = $rule.WordText; TicketText = $rule.TicketText; IsDefault = $rule.IsDefault }
        }
        # Merge in any new default rules not already in user's file (e.g. VMware, vCenter added in newer versions)
        $defaults = Get-DefaultRemediationRules
        $existingPatterns = @{}
        foreach ($r in $script:RemediationRules) { $existingPatterns[$r.Pattern] = $true }
        $added = 0
        foreach ($d in $defaults) {
            if (-not $existingPatterns.ContainsKey($d.Pattern)) {
                $script:RemediationRules += @{ Pattern = $d.Pattern; WordText = $d.WordText; TicketText = $d.TicketText; IsDefault = $d.IsDefault }
                $existingPatterns[$d.Pattern] = $true
                $added++
            }
        }
        if ($added -gt 0 -and -not [string]::IsNullOrEmpty($script:RemediationRulesPath)) {
            Save-RemediationRules | Out-Null
            Write-Host "Added $added new default remediation rule(s)"
        }
        Write-Host "Remediation rules loaded from local storage: $script:RemediationRulesPath"
    } else {
        $script:RemediationRules = Get-DefaultRemediationRules
        if (-not [string]::IsNullOrEmpty($script:RemediationRulesPath)) { Save-RemediationRules }
    }
}

function Save-RemediationRules {
    $script:CachedRemediationRulesForGuidance = $null  # invalidate cache
    
    # Try to use MemberberryIntegration module if available
    if (Get-Command Save-VScanMagicRemediationRules -ErrorAction SilentlyContinue) {
        try {
            if (Save-VScanMagicRemediationRules -Rules $script:RemediationRules) {
                $storageInfo = Get-VScanMagicStorageInfo
                Write-Host "Remediation rules saved to shared storage: $($storageInfo.DataPath)" -ForegroundColor Green
                return $true
            }
        } catch {
            Write-Warning "Could not save remediation rules to shared storage: $($_.Exception.Message). Falling back to local storage."
        }
    }
    
    # Fallback to local storage (backward compatibility)
    if ([string]::IsNullOrEmpty($script:RemediationRulesPath)) {
        Write-Warning "Remediation rules path is not set. Cannot save rules."
        return $false
    }
    if (Set-JsonFile -Path $script:RemediationRulesPath -Object $script:RemediationRules -Depth 10) {
        Write-Host "Remediation rules saved to local storage: $script:RemediationRulesPath"
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
    # Try to use MemberberryIntegration module if available
    if (Get-Command Get-VScanMagicCoveredSoftware -ErrorAction SilentlyContinue) {
        try {
            $sharedCoveredSoftware = Get-VScanMagicCoveredSoftware
            if ($sharedCoveredSoftware -and $sharedCoveredSoftware.Count -gt 0) {
                $script:CoveredSoftware = @()
                foreach ($item in $sharedCoveredSoftware) {
                    $script:CoveredSoftware += @{
                        Pattern = $item.Pattern
                        IsPattern = $item.IsPattern
                        Override = if ($null -ne $item.Override) { $item.Override } else { $false }
                    }
                }
                $storageInfo = Get-VScanMagicStorageInfo
                Write-Host "Covered software list loaded from shared storage: $($storageInfo.DataPath)" -ForegroundColor Green
                return
            }
        } catch {
            Write-Warning "Could not load covered software from shared storage: $($_.Exception.Message). Falling back to local storage."
        }
    }
    
    # Fallback to local storage (backward compatibility)
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
            Write-Host "Covered software list loaded from local storage: $script:CoveredSoftwarePath"
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
    # Try to use MemberberryIntegration module if available
    if (Get-Command Save-VScanMagicCoveredSoftware -ErrorAction SilentlyContinue) {
        try {
            if (Save-VScanMagicCoveredSoftware -CoveredSoftware $script:CoveredSoftware) {
                $storageInfo = Get-VScanMagicStorageInfo
                Write-Host "Covered software list saved to shared storage: $($storageInfo.DataPath)" -ForegroundColor Green
                return $true
            }
        } catch {
            Write-Warning "Could not save covered software to shared storage: $($_.Exception.Message). Falling back to local storage."
        }
    }
    
    # Fallback to local storage (backward compatibility)
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
        Write-Host "Covered software list saved to local storage: $script:CoveredSoftwarePath"
        return $true
    } catch {
        Write-Warning "Could not save covered software list: $($_.Exception.Message)"
        return $false
    }
}

# --- General Recommendations Persistence ---

function Load-GeneralRecommendations {
    # Try to use MemberberryIntegration module if available
    if (Get-Command Get-VScanMagicGeneralRecommendations -ErrorAction SilentlyContinue) {
        try {
            $sharedRecommendations = Get-VScanMagicGeneralRecommendations
            if ($sharedRecommendations -and $sharedRecommendations.Count -gt 0) {
                $script:GeneralRecommendations = @()
                foreach ($rec in $sharedRecommendations) {
                    $script:GeneralRecommendations += @{
                        Product = $rec.Product
                        Recommendations = $rec.Recommendations
                    }
                }
                $storageInfo = Get-VScanMagicStorageInfo
                Write-Host "General recommendations loaded from shared storage: $($storageInfo.DataPath)" -ForegroundColor Green
                return
            }
        } catch {
            Write-Warning "Could not load general recommendations from shared storage: $($_.Exception.Message). Falling back to local storage."
        }
    }
    
    # Fallback to local storage (backward compatibility)
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
            Write-Host "General recommendations loaded from local storage: $script:GeneralRecommendationsPath"
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
    # Try to use MemberberryIntegration module if available
    if (Get-Command Save-VScanMagicGeneralRecommendations -ErrorAction SilentlyContinue) {
        try {
            if (Save-VScanMagicGeneralRecommendations -GeneralRecommendations $script:GeneralRecommendations) {
                $storageInfo = Get-VScanMagicStorageInfo
                Write-Host "General recommendations saved to shared storage: $($storageInfo.DataPath)" -ForegroundColor Green
                return $true
            }
        } catch {
            Write-Warning "Could not save general recommendations to shared storage: $($_.Exception.Message). Falling back to local storage."
        }
    }
    
    # Fallback to local storage (backward compatibility)
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
        Write-Host "General recommendations saved to local storage: $script:GeneralRecommendationsPath"
        return $true
    } catch {
        Write-Warning "Could not save general recommendations: $($_.Exception.Message)"
        return $false
    }
}

# --- Templates Persistence (Email, Ticket Notes) ---

function Get-DefaultTemplates {
    return @{
        EmailTemplate = @{
            SubjectFormat = "{Year} Q{Quarter} Vulnerability Scan Follow Up"
            Body = @"
Subject: {Year} Q{Quarter} Vulnerability Scan Follow Up

Good {Greeting},

Your quarterly vulnerability scan report has been completed and is available in your client folder.

Recommended remediation priorities ({TopNLabel}):
<link to top ten report from onedrive>

Complete report package:
<onedrive link to folder containing reports>

The folder contains the following reports:
• Pending Remediation EPSS Score Report – Classifies vulnerabilities by Exploit Prediction Scoring System (EPSS), which measures the likelihood of exploitation within 30 days (scale 0–1.0, with 1.0 being most critical).
• All Vulnerabilities Report – A comprehensive list of all detected vulnerabilities (internal and external), from critical to low severity.
• Executive Summary Report – A high-level overview of your security posture and network information.
• External Scan – Detected vulnerabilities and services exposed to the internet.
• Suppressed Vulnerabilities Report – Vulnerabilities that have been suppressed (e.g., false positives or accepted risk) and will not appear on future remediation lists.

Not all vulnerabilities may be feasible to remediate depending on business or technical constraints.

Schedule time with me
<scheduling link>

{NoteText}

We appreciate your commitment to security. Addressing these vulnerabilities is essential for maintaining the protection of your systems.

Sincerely,

{PreparedBy}
"@
        }
        TicketNotes = @{
            StepsBeforeTickets = @"
- Examined lightweight agents
- Verified probe setup
- Checked agent/probe count compared to other systems
- Examined credential mappings
- Examined external assets
- Checked nmap interface on probe
- Verified deprecated item list
- Created all reports
- Assessed reports
- {ReportStepLine}
"@
            StepsAfterTickets = @"
- Sent secure email with reports to contact
- Sent TimeZest meeting request
"@
            ResolvedQuestion = "Is the task resolved?"
            ResolvedAnswer = "Yes - completed"
            NextStepsQuestion = "Next step(s)"
            NextStepsText = "TimeZest meeting request has been sent. Please select a time to meet if you would like to discuss this further."
        }
    }
}

function Load-Templates {
    if (-not [string]::IsNullOrEmpty($script:TemplatesPath) -and (Test-Path $script:TemplatesPath)) {
        try {
            $json = Get-Content $script:TemplatesPath -Raw -Encoding UTF8 | ConvertFrom-Json
            $script:Templates = @{
                EmailTemplate = @{
                    SubjectFormat = if ($json.EmailTemplate.SubjectFormat) { $json.EmailTemplate.SubjectFormat } else { (Get-DefaultTemplates).EmailTemplate.SubjectFormat }
                    Body = if ($json.EmailTemplate.Body) { $json.EmailTemplate.Body } else { (Get-DefaultTemplates).EmailTemplate.Body }
                }
                TicketNotes = @{
                    StepsBeforeTickets = if ($json.TicketNotes.StepsBeforeTickets) { $json.TicketNotes.StepsBeforeTickets } else { (Get-DefaultTemplates).TicketNotes.StepsBeforeTickets }
                    StepsAfterTickets = if ($json.TicketNotes.StepsAfterTickets) { $json.TicketNotes.StepsAfterTickets } else { (Get-DefaultTemplates).TicketNotes.StepsAfterTickets }
                    ResolvedQuestion = if ($json.TicketNotes.ResolvedQuestion) { $json.TicketNotes.ResolvedQuestion } else { (Get-DefaultTemplates).TicketNotes.ResolvedQuestion }
                    ResolvedAnswer = if ($json.TicketNotes.ResolvedAnswer) { $json.TicketNotes.ResolvedAnswer } else { (Get-DefaultTemplates).TicketNotes.ResolvedAnswer }
                    NextStepsQuestion = if ($json.TicketNotes.NextStepsQuestion) { $json.TicketNotes.NextStepsQuestion } else { (Get-DefaultTemplates).TicketNotes.NextStepsQuestion }
                    NextStepsText = if ($json.TicketNotes.NextStepsText) { $json.TicketNotes.NextStepsText } else { (Get-DefaultTemplates).TicketNotes.NextStepsText }
                }
            }
        } catch {
            Write-Warning "Could not load templates: $($_.Exception.Message). Using defaults."
            $script:Templates = Get-DefaultTemplates
        }
    } else {
        $script:Templates = Get-DefaultTemplates
    }
}

function Save-Templates {
    if ([string]::IsNullOrEmpty($script:TemplatesPath)) { return $false }
    try {
        $settingsDir = [System.IO.Path]::GetDirectoryName($script:TemplatesPath)
        if (-not (Test-Path $settingsDir)) { New-Item -Path $settingsDir -ItemType Directory -Force | Out-Null }
        $script:Templates | ConvertTo-Json -Depth 10 | Set-Content $script:TemplatesPath -Encoding UTF8
        return $true
    } catch {
        Write-Warning "Could not save templates: $($_.Exception.Message)"
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

# Modifier text for ticket/report subject lines only: no "ticket generated" text (subject IS the ticket).
# When AfterHours, caller prepends "After Hours - " to the subject (no "ticket generated" or "after-hours work required").
function Get-ModifierTextForSubject {
    param(
        [bool]$AfterHours,
        [bool]$TicketGenerated,
        [bool]$ThirdParty
    )

    # Ticket generated: no modifier (subject IS the ticket; redundant to say "ticket generated")
    if ($TicketGenerated) {
        return ""
    }
    # After hours: caller prepends "After Hours - "; don't add "after-hours work required" (redundant)
    if ($AfterHours -and $ThirdParty) {
        return " - 3rd party application"
    }
    if ($AfterHours) {
        return ""
    }
    # Third party only (no "approval needed" - ticket creation implies approval from quoting process)
    if ($ThirdParty) {
        return " - 3rd party application"
    }

    return ""
}

# First-party vendors: always treated as first party (not 3rd party) in time estimate dialog
$script:FirstPartyVendorPatterns = @(
    '*Sonicwall*', '*SonicWall*',
    '*Fortinet*', '*FortiGate*', '*Forti*',
    '*Microsoft*',
    '*Windows 11*', '*Windows 10*', '*Windows Server*', '*Windows 8*', '*Windows 7*',
    '*HP *', '* HP *', '*HP Pro*', '*HP LaserJet*', '*HP OfficeJet*', '*Hewlett-Packard*',
    '*Duo Security*', '*Duo *',
    '*VMware*', '*vSphere*', '*VMware Tools*'
)

function Get-TimeEstimateGroupKey {
    <#
    .SYNOPSIS
    Returns a group key for time estimate lookups. Visual C++ variants and .NET variants are grouped;
    other products use the product name as-is.
    #>
    param([string]$ProductName)
    if ([string]::IsNullOrWhiteSpace($ProductName)) { return $ProductName }
    $p = $ProductName.Trim()
    if ($p -like "*Visual C++*" -or $p -like "*Visual C#*") { return "Microsoft Visual C++" }
    if ($p -like "*\.NET Framework*" -or $p -match "Microsoft \.NET Framework") { return "Microsoft .NET Framework" }
    if ($p -like "*\.NET Core*" -or $p -match "Microsoft \.NET Core") { return "Microsoft .NET Core" }
    if ($p -like "*\.NET Runtime*" -or $p -like "*\.NET 5*" -or $p -like "*\.NET 6*" -or $p -like "*\.NET 7*" -or $p -like "*\.NET 8*" -or $p -like "*\.NET 9*" -or $p -like "*\.NET 10*") { return "Microsoft .NET Runtime" }
    if ($p -match "Microsoft \.NET") { return "Microsoft .NET" }
    return $p
}

function Test-IsFirstPartyVendor {
    param(
        [string]$ProductName
    )
    if ([string]::IsNullOrWhiteSpace($ProductName)) { return $false }
    $p = $ProductName.Trim()
    foreach ($pattern in $script:FirstPartyVendorPatterns) {
        if ($p -like $pattern) { return $true }
    }
    return $false
}

# Build download path with truncated client name when path would exceed Windows MAX_PATH (~260 chars)
function Get-SafeDownloadPath {
    param(
        [string]$TargetDir,
        [string]$ClientName,
        [string]$ReportName,
        [string]$Ext,
        [string]$Timestamp
    )
    $suffix = " - $ReportName - $Timestamp.$Ext"
    $maxPathLen = 250
    $availableForName = $maxPathLen - $TargetDir.Length - 1 - $suffix.Length
    $safeClient = if ($availableForName -lt $ClientName.Length) {
        if ($availableForName -gt 5) { $ClientName.Substring(0, $availableForName) } else { $ClientName.Substring(0, [Math]::Min(5, $ClientName.Length)) }
    } else {
        $ClientName
    }
    $filename = "$safeClient$suffix"
    return Join-Path $TargetDir $filename
}

# Build report output path with truncated company name when path would exceed Windows MAX_PATH (~260 chars)
function Get-SafeReportOutputPath {
    param(
        [string]$TargetDir,
        [string]$CompanyName,
        [string]$ReportSuffix,
        [string]$Ext
    )
    $filename = "$CompanyName$ReportSuffix.$Ext"
    $fullPath = Join-Path $TargetDir $filename
    $maxPathLen = 250
    if ($fullPath.Length -le $maxPathLen) { return $fullPath }
    $suffixPart = "$ReportSuffix.$Ext"
    $availableForName = $maxPathLen - $TargetDir.Length - 1 - $suffixPart.Length
    $safeName = if ($availableForName -lt $CompanyName.Length) {
        if ($availableForName -gt 5) { $CompanyName.Substring(0, $availableForName) } else { $CompanyName.Substring(0, [Math]::Min(5, $CompanyName.Length)) }
    } else {
        $CompanyName
    }
    return Join-Path $TargetDir "$safeName$suffixPart"
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

function Test-IsPlaceholderFixText {
    <#
    .SYNOPSIS
    Returns $true if the Fix/Solution text is a placeholder (None, [None], N/A) with no real remediation.
    #>
    param([string]$RawFix)
    if ([string]::IsNullOrWhiteSpace($RawFix)) { return $true }
    $t = $RawFix.Trim()
    return $t -match '^\s*(\[?\s*)?(None|N/A|nil)\s*(\]?\s*)?$' -or $t -eq "['None']" -or $t -eq '["None"]'
}

function ConvertTo-ReadableFixText {
    <#
    .SYNOPSIS
    Cleans up raw Fix/Solution text (versions, KBs, patches) from ConnectSecure into readable format.
    Converts ['5077181']; ['148.0.0'] style into "Apply KB5077181. Update to version 148.0.0 or later."
    #>
    param([string]$RawFix)
    if ([string]::IsNullOrWhiteSpace($RawFix)) { return $RawFix }
    if (Test-IsPlaceholderFixText -RawFix $RawFix) { return '' }
    $parts = [System.Collections.ArrayList]::new()
    $matches = [regex]::Matches($RawFix, "'([^']*)'")
    foreach ($m in $matches) {
        $item = $m.Groups[1].Value.Trim()
        if ([string]::IsNullOrWhiteSpace($item) -or $item -eq "None") { continue }
        if ($item -match '^\d{6,8}$') {
            $null = $parts.Add("Apply Windows Update KB$item")
        } elseif ($item -match '^[\d\.]+$') {
            $null = $parts.Add("Update to version $item or later")
        } elseif ($item -eq "Latest Patch") {
            $null = $parts.Add("Apply the latest patch")
        } else {
            $null = $parts.Add($item)
        }
    }
    if ($parts.Count -eq 0) { return $RawFix }
    $unique = $parts | Select-Object -Unique
    $result = ($unique -join ". ").Trim()
    return $result
}

function Test-AIApiKeyConfigured {
    <#
    .SYNOPSIS
    Returns $true if any AI API key (ChatGPT, Claude, Copilot) is configured.
    #>
    $chat = $script:UserSettings.AIApiKeyChatGPT
    $claude = $script:UserSettings.AIApiKeyClaude
    $copilot = $script:UserSettings.AIApiKeyCopilot
    return (-not [string]::IsNullOrWhiteSpace($chat)) -or (-not [string]::IsNullOrWhiteSpace($claude)) -or (-not [string]::IsNullOrWhiteSpace($copilot))
}

function Invoke-AIImproveRemediationText {
    <#
    .SYNOPSIS
    Rephrases vulnerability remediation text using AI (ChatGPT, Claude, or Copilot).
    Uses first available API key. Returns original text on error.
    #>
    param(
        [string]$Text,
        [string]$ProductName = "",
        [string]$CveIdList = "",
        [string]$HostContext = ""
    )

    if ([string]::IsNullOrWhiteSpace($Text)) {
        return $Text
    }

    $cleaned = ConvertTo-ReadableFixText -RawFix $Text
    $vulnContext = ""
    $parts = @()
    if (-not [string]::IsNullOrWhiteSpace($CveIdList)) { $parts += "CVE ID(s): $CveIdList" }
    if (-not [string]::IsNullOrWhiteSpace($ProductName)) { $parts += "Product/OS: $ProductName" }
    if (-not [string]::IsNullOrWhiteSpace($HostContext)) { $parts += "Affected hosts: $HostContext" }
    if ($parts.Count -gt 0) { $vulnContext = ($parts -join "`n") + "`n`n" }

    $cveOnlyHint = ""
    $trimmed = $cleaned.Trim()
    if ($trimmed -match '^CVE-\d{4}-\d+' -and $trimmed.Length -lt 120) {
        $cveOnlyHint = "IMPORTANT: The input is only a CVE ID (or minimal text). First determine what the device or software is (use Product/OS and Affected hosts). Remediation often requires reviewing the manufacturer's security advisories and vulnerability data, checking for firmware updates, and considering configuration changes (e.g. mitigating controls, network segmentation). Provide platform-appropriate guidance: Windows Update KB for Windows hosts, firmware updates for printers/IoT/embedded devices, vendor patches for specific software. Match the fix to the host/OS type.`n`n"
    }

    $prompt = @"
You are rephrasing SPECIFIC remediation steps for ONE vulnerability. The text below contains the exact fix (e.g. KB numbers, version updates) for a particular CVE or product.

$cveOnlyHint
RULES:
- Rephrase ONLY the provided text. Do NOT add generic advice, best practices, or numbered lists.
- Do NOT output things like "Prioritize remediation", "Patch management", "Vulnerability scanning", "Documentation" - those are generic.
- Output ONLY a clear, professional rewrite of the SPECIFIC fix below. One short paragraph or a few sentences max.
- If the input is already specific (e.g. "Apply KB5077181"), just make it slightly more readable - do not expand into generic guidance.
- Output ONLY the rephrased text. No preamble, no explanation, no bullet lists of general practices.
- NEVER respond with "I don't have specific information" or similar. If CVE IDs are provided or the vulnerability is known (e.g. Ripple20/rippl20, CVE-2020-11898, CVE-2020-11910), provide the best available remediation guidance for that specific vulnerability.
- When CVE IDs are given: determine the device/software identity, review manufacturer vulnerability data and security advisories, check for firmware updates, consider configuration mitigations. Provide CVE-specific remediation (patches, KB numbers, version updates, workarounds). When input is only a CVE, use Product/OS and host context to tailor the fix (Windows vs Linux vs firmware vs IoT).
- CRITICAL: The remediation MUST match the Product/OS above. If Product is Windows 11 or Windows Server, output ONLY Windows Update KB numbers or build numbers - NEVER Chrome, Firefox, or other software version numbers. If Product is Google Chrome, output ONLY Chrome version numbers. Do NOT mix product types.
- Common vulnerabilities: rippl20-icmp = Ripple20 Treck TCP/IP ICMP flaws; remediate via firmware updates from manufacturer, network segmentation where patching not possible.

$vulnContext
Input to rephrase:
"@

    # Try ChatGPT first, then Claude, then Copilot
    $key = $script:UserSettings.AIApiKeyChatGPT
    if (-not [string]::IsNullOrWhiteSpace($key)) {
        try {
            $reqBody = @{
                model = "gpt-4o-mini"
                messages = @(
                    @{ role = "user"; content = "$prompt`n`n---`n`n$cleaned" }
                )
                max_tokens = 1024
            } | ConvertTo-Json -Depth 5
            $headers = @{
                "Authorization" = "Bearer $key"
                "Content-Type" = "application/json"
            }
            $resp = Invoke-RestMethod -Uri "https://api.openai.com/v1/chat/completions" -Method Post -Headers $headers -Body $reqBody -TimeoutSec 30
            $improved = $resp.choices[0].message.content.Trim()
            if (-not [string]::IsNullOrWhiteSpace($improved)) {
                if ($improved -match "don't have specific information|do not have specific information|don't have information|no specific information") {
                    $fallback = Get-RemediationGuidance -ProductName $ProductName -OutputType 'Word'
                    if (-not [string]::IsNullOrWhiteSpace($fallback)) { return $fallback }
                }
                return $improved
            }
        } catch {
            Write-Log "ChatGPT AI improvement failed: $($_.Exception.Message)" -Level Warning
        }
    }

    $key = $script:UserSettings.AIApiKeyClaude
    if (-not [string]::IsNullOrWhiteSpace($key)) {
        $claudeRetries = 2
        for ($attempt = 0; $attempt -le $claudeRetries; $attempt++) {
            try {
                $reqBody = @{
                    model = "claude-haiku-4-5-20251001"
                    max_tokens = 1024
                    messages = @(
                        @{ role = "user"; content = "$prompt`n`n---`n`n$cleaned" }
                    )
                } | ConvertTo-Json -Depth 5
                $headers = @{
                    "x-api-key" = $key
                    "anthropic-version" = "2023-06-01"
                    "Content-Type" = "application/json"
                }
                $resp = Invoke-RestMethod -Uri "https://api.anthropic.com/v1/messages" -Method Post -Headers $headers -Body $reqBody -TimeoutSec 30
                $content = @($resp.content)
                $improved = ''
                foreach ($block in $content) {
                    if ($block -and $block.type -eq 'text' -and $null -ne $block.text) {
                        $improved += [string]$block.text
                    }
                }
                if ([string]::IsNullOrWhiteSpace($improved) -and $content.Count -gt 0 -and $content[0] -and $null -ne $content[0].text) {
                    $improved = [string]$content[0].text
                }
                $improved = if ($improved) { $improved.Trim() } else { '' }
                if (-not [string]::IsNullOrWhiteSpace($improved)) {
                    if ($improved -match "don't have specific information|do not have specific information|don't have information|no specific information") {
                        $fallback = Get-RemediationGuidance -ProductName $ProductName -OutputType 'Word'
                        if (-not [string]::IsNullOrWhiteSpace($fallback)) { return $fallback }
                    }
                    return $improved
                }
                break
            } catch {
                $is429 = $_.Exception.Message -match '429|Too Many Requests'
                if ($is429 -and $attempt -lt $claudeRetries) {
                    $waitSec = 5 + ($attempt * 3)
                    Write-Log "Claude rate limited (429). Waiting ${waitSec}s before retry..." -Level Warning
                    Start-Sleep -Seconds $waitSec
                } else {
                    Write-Log "Claude AI improvement failed: $($_.Exception.Message)" -Level Warning
                    break
                }
            }
        }
    }

    $key = $script:UserSettings.AIApiKeyCopilot
    if (-not [string]::IsNullOrWhiteSpace($key)) {
        try {
            $reqBody = @{
                model = "gpt-4o-mini"
                messages = @(
                    @{ role = "user"; content = "$prompt`n`n---`n`n$cleaned" }
                )
                max_tokens = 1024
            } | ConvertTo-Json -Depth 5
            $headers = @{
                "Authorization" = "Bearer $key"
                "Content-Type" = "application/json"
            }
                $resp = Invoke-RestMethod -Uri "https://api.openai.com/v1/chat/completions" -Method Post -Headers $headers -Body $reqBody -TimeoutSec 30
            $improved = $resp.choices[0].message.content.Trim()
            if (-not [string]::IsNullOrWhiteSpace($improved)) {
                if ($improved -match "don't have specific information|do not have specific information|don't have information|no specific information") {
                    $fallback = Get-RemediationGuidance -ProductName $ProductName -OutputType 'Word'
                    if (-not [string]::IsNullOrWhiteSpace($fallback)) { return $fallback }
                }
                return $improved
            }
        } catch {
            Write-Log "Copilot AI improvement failed: $($_.Exception.Message)" -Level Warning
        }
    }

    Write-Log "No AI API key configured or all attempts failed. Configure AI keys in Settings." -Level Warning
    return $Text
}

function Invoke-AIImproveRemediationTextBatch {
    <#
    .SYNOPSIS
    Rephrases multiple vulnerability remediation texts via AI API calls.
    Processes in chunks to avoid rate limits; retries on 429.
    .PARAMETER Items
    Array of hashtables: @{ Text, ProductName, CveIdList }
    .OUTPUTS
    Array of improved strings (same order as input). Falls back to original text on parse/API failure.
    #>
    param(
        [Parameter(Mandatory = $true)]
        [array]$Items
    )

    if ($null -eq $Items -or $Items.Count -eq 0) {
        return @()
    }
    if ($Items.Count -eq 1) {
        $i = $Items[0]
        $hostCtx = if ($i.HostContext) { $i.HostContext } else { "" }
        return @(Invoke-AIImproveRemediationText -Text $i.Text -ProductName $i.ProductName -CveIdList $i.CveIdList -HostContext $hostCtx)
    }

    $script:AIBatch429Count = 0
    $delim = "###RESULT###"
    $chunkSize = $script:Config.AIBatchChunkSize
    $chunkDelay = $script:Config.AIBatchChunkDelaySeconds
    $rulesHeader = @"
You are rephrasing SPECIFIC remediation steps for MULTIPLE vulnerabilities. Below are N items. For EACH item, output the improved text, then the exact delimiter '$delim' (including the delimiter line).

RULES (apply to EACH item):
- Rephrase ONLY the provided text. Do NOT add generic advice, best practices, or numbered lists.
- Do NOT output things like 'Prioritize remediation', 'Patch management', 'Vulnerability scanning' - those are generic.
- Output ONLY a clear, professional rewrite of the SPECIFIC fix. One short paragraph or a few sentences max.
- If the input is already specific (e.g. 'Apply KB5077181'), just make it slightly more readable.
- NEVER respond with 'I don't have specific information'. Use CVE IDs to provide CVE-specific remediation when given.
- When investigating CVEs: determine what the device/software is (Product/OS, Affected hosts), review manufacturer security advisories and vulnerability data, check for firmware updates, consider configuration mitigations. Provide platform-appropriate remediation (Windows KB for Windows, firmware for printers/IoT/embedded, vendor patch for software).
- CRITICAL: Remediation MUST match the Product/OS. Windows 11/Server = Windows KB numbers only, never Chrome/Firefox versions. Google Chrome = Chrome versions only. Do NOT mix product types.
- Common: rippl20-icmp = Ripple20 Treck TCP/IP; remediate via firmware updates from manufacturer, network segmentation.

OUTPUT FORMAT: For each item, output ONLY the improved text (no 'ITEM 1', 'ITEM 2', or numbering), then a line with exactly: $delim

--- INPUT ITEMS ---
"@

    $tryApiChunk = {
        param($key, $provider, $prompt, $maxTokens)
        if ([string]::IsNullOrWhiteSpace($key)) { return $null }
        try {
            if ($provider -eq 'ChatGPT' -or $provider -eq 'Copilot') {
                $reqBody = @{ model = "gpt-4o-mini"; messages = @( @{ role = "user"; content = $prompt } ); max_tokens = $maxTokens } | ConvertTo-Json -Depth 5
                $headers = @{ "Authorization" = "Bearer $key"; "Content-Type" = "application/json" }
                $resp = Invoke-RestMethod -Uri "https://api.openai.com/v1/chat/completions" -Method Post -Headers $headers -Body $reqBody -TimeoutSec 90
                return $resp.choices[0].message.content.Trim()
            }
            if ($provider -eq 'Claude') {
                $reqBody = @{ model = "claude-haiku-4-5-20251001"; max_tokens = [Math]::Min(4096, $maxTokens); messages = @( @{ role = "user"; content = $prompt } ) } | ConvertTo-Json -Depth 5
                $headers = @{ "x-api-key" = $key; "anthropic-version" = "2023-06-01"; "Content-Type" = "application/json" }
                $resp = Invoke-RestMethod -Uri "https://api.anthropic.com/v1/messages" -Method Post -Headers $headers -Body $reqBody -TimeoutSec 90
                $content = @($resp.content)
                $rawText = ''
                foreach ($block in $content) {
                    if ($block -and $block.type -eq 'text' -and $null -ne $block.text) { $rawText += [string]$block.text }
                }
                if ([string]::IsNullOrWhiteSpace($rawText) -and $content.Count -gt 0 -and $content[0] -and $null -ne $content[0].text) {
                    $rawText = [string]$content[0].text
                }
                if ($rawText) { return $rawText.Trim() }
                return ''
            }
        } catch {
            $errDetail = $_.Exception.Message
            if ($_.ErrorDetails.Message) { $errDetail += " | " + $_.ErrorDetails.Message }
            Write-Log "$provider batch AI failed: $errDetail" -Level Warning
            throw
        }
        return $null
    }

    $allResults = @()
    $chunks = @()
    for ($i = 0; $i -lt $Items.Count; $i += $chunkSize) {
        $end = [Math]::Min($i + $chunkSize, $Items.Count)
        $chunks += ,@($Items[$i..($end - 1)])
    }

    for ($c = 0; $c -lt $chunks.Count; $c++) {
        $chunk = $chunks[$c]
        $sb = [System.Text.StringBuilder]::new()
        [void]$sb.AppendLine($rulesHeader)
        for ($idx = 0; $idx -lt $chunk.Count; $idx++) {
            $i = $chunk[$idx]
            $textToUse = $i.Text
            if ($textToUse.Length -le 2 -and -not [string]::IsNullOrWhiteSpace($i.ProductName)) {
                $fallback = Get-RemediationGuidance -ProductName $i.ProductName -OutputType 'Word'
                if (-not [string]::IsNullOrWhiteSpace($fallback)) { $textToUse = $fallback }
            }
            $cleaned = ConvertTo-ReadableFixText -RawFix $textToUse
            [void]$sb.AppendLine("")
            [void]$sb.AppendLine("--- ITEM $($idx + 1) ---")
            if (-not [string]::IsNullOrWhiteSpace($i.CveIdList)) { [void]$sb.AppendLine("CVE ID(s): $($i.CveIdList)") }
            if (-not [string]::IsNullOrWhiteSpace($i.ProductName)) { [void]$sb.AppendLine("Product/OS: $($i.ProductName)") }
            if (-not [string]::IsNullOrWhiteSpace($i.HostContext)) { [void]$sb.AppendLine("Affected hosts: $($i.HostContext)") }
            [void]$sb.AppendLine("Text to rephrase: $cleaned")
        }
        $chunkPrompt = $sb.ToString()
        $chunkMaxTokens = [Math]::Min(8192, 1024 + ($chunk.Count * 400))

        $raw = $null
        $claudeRetries = 1
        $script:AIBatch429Count = if ($null -eq $script:AIBatch429Count) { 0 } else { $script:AIBatch429Count }
        if ($script:AIBatch429Count -ge 3) {
            Write-Log "AI batch: too many 429s ($($script:AIBatch429Count)). Skipping remaining chunks. Using original text. Try again later or use Improve Selected for fewer items." -Level Warning
            try {
                [System.Windows.Forms.MessageBox]::Show("Rate limit exceeded (429). $($allResults.Count) items improved; remaining use original text.`n`nTry again later or use Improve Selected for fewer items.", "AI Rate Limit", [System.Windows.Forms.MessageBoxButtons]::OK, [System.Windows.Forms.MessageBoxIcon]::Warning)
            } catch { }
            $allResults += @($chunk | ForEach-Object { $_.Text })
            $remaining = $Items.Count - $allResults.Count
            for ($r = 0; $r -lt $remaining; $r++) {
                $idx = $allResults.Count + $r
                if ($idx -lt $Items.Count) { $allResults += $Items[$idx].Text }
            }
            return $allResults
        }
        foreach ($provider in @('ChatGPT','Claude','Copilot')) {
            $key = if ($provider -eq 'ChatGPT') { $script:UserSettings.AIApiKeyChatGPT } elseif ($provider -eq 'Claude') { $script:UserSettings.AIApiKeyClaude } else { $script:UserSettings.AIApiKeyCopilot }
            if ([string]::IsNullOrWhiteSpace($key)) { continue }
            for ($attempt = 0; $attempt -le $claudeRetries; $attempt++) {
                try {
                    $raw = & $tryApiChunk $key $provider $chunkPrompt $chunkMaxTokens
                    $script:AIBatch429Count = 0
                    break
                } catch {
                    $is429 = $_.Exception.Message -match '429|Too Many Requests'
                    if ($is429) {
                        $script:AIBatch429Count++
                        if ($attempt -lt $claudeRetries) {
                            $waitSec = 30
                            Write-Log "Claude batch rate limited (429). Waiting ${waitSec}s before retry ($($script:AIBatch429Count)/3 total)..." -Level Warning
                            for ($w = 0; $w -lt $waitSec; $w++) {
                                Start-Sleep -Seconds 1
                                try { [System.Windows.Forms.Application]::DoEvents() } catch { }
                            }
                        } else {
                            $raw = $null
                            break
                        }
                    } else {
                        $raw = $null
                        break
                    }
                }
            }
            if ($raw) { break }
        }

        if ([string]::IsNullOrWhiteSpace($raw)) {
            Write-Log "Batch AI chunk $($c + 1)/$($chunks.Count): all providers failed. Using original text for chunk." -Level Warning
            $allResults += @($chunk | ForEach-Object { $_.Text })
        } else {
            $parts = $raw -split [regex]::Escape($delim)
            $itemPrefix = '^\s*(?:\*\*)?\s*ITEM\s+\d+\s*:?\s*\r?\n?\s*'
            for ($idx = 0; $idx -lt $chunk.Count; $idx++) {
                $improved = if ($idx -lt $parts.Count) { $parts[$idx].Trim() } else { $null }
                if ($improved -and $improved -match $itemPrefix) { $improved = $improved -replace $itemPrefix, '' }
                if ([string]::IsNullOrWhiteSpace($improved) -or $improved -match "don't have specific information|do not have specific information") {
                    $fallback = Get-RemediationGuidance -ProductName $chunk[$idx].ProductName -OutputType 'Word'
                    $improved = if (-not [string]::IsNullOrWhiteSpace($fallback)) { $fallback } else { $chunk[$idx].Text }
                }
                $allResults += $improved
            }
        }

        if ($c -lt $chunks.Count - 1 -and $chunkDelay -gt 0) {
            for ($d = 0; $d -lt $chunkDelay; $d++) {
                Start-Sleep -Seconds 1
                try { [System.Windows.Forms.Application]::DoEvents() } catch { }
            }
        }
    }
    return $allResults
}
