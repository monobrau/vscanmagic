# VScanMagic-Core.ps1 - Config, settings, persistence, helpers
# Dot-sourced by VScanMagic-GUI.ps1 (do not run directly)

# --- Configuration ---
$script:Config = @{
    AppName = "VScanMagic v4"
    Version = "4.0.5"
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
$script:ConnectWiseAutomateCredentialsPath = Join-Path $script:SettingsDirectory "ConnectWise-Automate-Credentials.json"
$script:CompanyFolderMapPath = Join-Path $script:SettingsDirectory "VScanMagic_CompanyFolderMap.json"
$script:ReportFolderHistoryPath = Join-Path $script:SettingsDirectory "VScanMagic_ReportFolderHistory.json"
$script:TemplatesPath = Join-Path $script:SettingsDirectory "VScanMagic_Templates.json"
$script:Templates = $null

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
    ReportsBasePath = ""  # Base folder for client output; when set, uses [Base]\[Folder]\[Year] - [QN]\
    # AI API Keys (future: email, ticket notes, remediation, time estimate guidance)
    AIApiKeyCopilot = ""
    AIApiKeyChatGPT = ""
    AIApiKeyClaude = ""
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
    $script:CompanyFolderMapPath = Join-Path $script:SettingsDirectory "VScanMagic_CompanyFolderMap.json"
    $script:ReportFolderHistoryPath = Join-Path $script:SettingsDirectory "VScanMagic_ReportFolderHistory.json"
    $script:ConnectWiseAutomateCredentialsPath = Join-Path $script:SettingsDirectory "ConnectWise-Automate-Credentials.json"
    $script:TemplatesPath = Join-Path $script:SettingsDirectory "VScanMagic_Templates.json"
    
    if (Ensure-SettingsDirectory -Path $script:SettingsDirectory) { Write-Host "Created settings directory: $script:SettingsDirectory" }
}

function Backup-Settings {
    param([string]$OutputPath = $null)
    $settingsFiles = @(
        "VScanMagic_Settings.json",
        "VScanMagic_RemediationRules.json",
        "VScanMagic_CoveredSoftware.json",
        "VScanMagic_GeneralRecommendations.json",
        "VScanMagic_CompanyFolderMap.json",
        "VScanMagic_ReportFolderHistory.json",
        "VScanMagic_Templates.json",
        "ConnectSecure-Credentials.json",
        "ConnectSecure-Companies-Cache.json",
        "ConnectWise-Automate-Credentials.json"
    )
    if (-not (Test-Path $script:SettingsDirectory)) {
        return $null
    }
    $timestamp = Get-Date -Format "yyyy-MM-dd_HHmmss"
    $defaultName = "VScanMagic_Settings_Backup_$timestamp.zip"
    if ([string]::IsNullOrWhiteSpace($OutputPath)) {
        $OutputPath = Join-Path ([Environment]::GetFolderPath("Desktop")) $defaultName
    } elseif ([System.IO.Directory]::Exists($OutputPath)) {
        $OutputPath = Join-Path $OutputPath $defaultName
    }
    $tempDir = Join-Path $env:TEMP "VScanMagic_Backup_$([Guid]::NewGuid().ToString('N').Substring(0,8))"
    try {
        New-Item -Path $tempDir -ItemType Directory -Force | Out-Null
        $copied = 0
        foreach ($f in $settingsFiles) {
            $src = Join-Path $script:SettingsDirectory $f
            if (Test-Path $src) {
                Copy-Item -Path $src -Destination (Join-Path $tempDir $f) -Force
                $copied++
            }
        }
        if ($copied -eq 0) {
            Write-Warning "No settings files found to backup."
            return $null
        }
        if (Test-Path $OutputPath) { Remove-Item $OutputPath -Force }
        Compress-Archive -Path (Join-Path $tempDir "*") -DestinationPath $OutputPath -Force
        return $OutputPath
    } finally {
        if (Test-Path $tempDir) { Remove-Item $tempDir -Recurse -Force -ErrorAction SilentlyContinue }
    }
}

function Restore-Settings {
    param([string]$BackupPath)
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
        $restored = 0
        Get-ChildItem -Path $tempDir -Filter "*.json" | ForEach-Object {
            $dest = Join-Path $script:SettingsDirectory $_.Name
            Copy-Item -Path $_.FullName -Destination $dest -Force
            $restored++
        }
        Remove-Item $tempDir -Recurse -Force -ErrorAction SilentlyContinue
        if ($restored -gt 0) {
            Load-UserSettings | Out-Null
            Update-SettingsPaths
            Load-RemediationRules
            Load-CoveredSoftware
            Load-GeneralRecommendations
            Load-CompanyFolderMap
            Load-Templates
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
        if ($null -ne $json.AIApiKeyCopilot) { $script:UserSettings.AIApiKeyCopilot = $json.AIApiKeyCopilot } else { $script:UserSettings.AIApiKeyCopilot = "" }
        if ($null -ne $json.AIApiKeyChatGPT) { $script:UserSettings.AIApiKeyChatGPT = $json.AIApiKeyChatGPT } else { $script:UserSettings.AIApiKeyChatGPT = "" }
        if ($null -ne $json.AIApiKeyClaude) { $script:UserSettings.AIApiKeyClaude = $json.AIApiKeyClaude } else { $script:UserSettings.AIApiKeyClaude = "" }
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

# First-party vendors: always treated as first party (not 3rd party) in time estimate dialog
$script:FirstPartyVendorPatterns = @(
    '*Sonicwall*', '*SonicWall*',
    '*Fortinet*', '*FortiGate*', '*Forti*',
    '*Microsoft*',
    '*HP *', '* HP *', '*HP Pro*', '*HP LaserJet*', '*HP OfficeJet*', '*Hewlett-Packard*',
    '*Duo Security*', '*Duo *'
)

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

function Test-IsCoveredSoftware {
    param(
        [string]$ProductName
    )
    if ([string]::IsNullOrWhiteSpace($ProductName)) { return $false }

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

function ConvertTo-ReadableFixText {
    <#
    .SYNOPSIS
    Cleans up raw Fix/Solution text (versions, KBs, patches) from ConnectSecure into readable format.
    Converts ['5077181']; ['148.0.0'] style into "Apply KB5077181. Update to version 148.0.0 or later."
    #>
    param([string]$RawFix)
    if ([string]::IsNullOrWhiteSpace($RawFix)) { return $RawFix }
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
        [string]$CveIdList = ""
    )

    if ([string]::IsNullOrWhiteSpace($Text)) {
        return $Text
    }

    $cleaned = ConvertTo-ReadableFixText -RawFix $Text
    $vulnContext = ""
    if (-not [string]::IsNullOrWhiteSpace($ProductName) -or -not [string]::IsNullOrWhiteSpace($CveIdList)) {
        $parts = @()
        if (-not [string]::IsNullOrWhiteSpace($CveIdList)) { $parts += "CVE ID(s): $CveIdList" }
        if (-not [string]::IsNullOrWhiteSpace($ProductName)) { $parts += "Product: $ProductName" }
        $vulnContext = ($parts -join "`n") + "`n`n"
    }
    $prompt = @"
You are rephrasing SPECIFIC remediation steps for ONE vulnerability. The text below contains the exact fix (e.g. KB numbers, version updates) for a particular CVE or product.

RULES:
- Rephrase ONLY the provided text. Do NOT add generic advice, best practices, or numbered lists.
- Do NOT output things like "Prioritize remediation", "Patch management", "Vulnerability scanning", "Documentation" - those are generic.
- Output ONLY a clear, professional rewrite of the SPECIFIC fix below. One short paragraph or a few sentences max.
- If the input is already specific (e.g. "Apply KB5077181"), just make it slightly more readable - do not expand into generic guidance.
- Output ONLY the rephrased text. No preamble, no explanation, no bullet lists of general practices.
- NEVER respond with "I don't have specific information" or similar. If CVE IDs are provided or the vulnerability is known (e.g. Ripple20/rippl20, CVE-2020-11898, CVE-2020-11910), provide the best available remediation guidance for that specific vulnerability.
- When CVE IDs are given, use them to provide CVE-specific remediation (patches, KB numbers, version updates, workarounds from NVD/CERT advisories).
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
                    model = "claude-3-haiku-20240307"
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
                $improved = $resp.content[0].text.Trim()
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
