# MemberberryIntegration.psm1
# Integration module for VScanMagic to use memberberry's shared data storage
# Enables multi-user access to client data and remediation rules via memberberry's config.json

# Module-level variables
$script:MemberberryConfigPath = $null
$script:MemberberryDataPath = $null
$script:DefaultLocalAppDataPath = Join-Path $env:LOCALAPPDATA "VScanMagic"

# File names for shared storage
$script:ClientDataFileName = "vscanmagic-clients.json"
$script:RemediationRulesFileName = "vscanmagic-remediation-rules.json"
$script:CoveredSoftwareFileName = "vscanmagic-covered-software.json"
$script:GeneralRecommendationsFileName = "vscanmagic-general-recommendations.json"

<#
.SYNOPSIS
    Gets the memberberry config.json path by checking common locations
.DESCRIPTION
    Searches for memberberry's config.json in common installation locations:
    - c:\git\memberberry
    - $env:USERPROFILE\Documents\memberberry
    - Current directory (if running from memberberry folder)
#>
function Get-MemberberryConfigPath {
    if ($script:MemberberryConfigPath) {
        return $script:MemberberryConfigPath
    }
    
    $possiblePaths = @(
        "c:\git\memberberry\config.json",
        "$env:USERPROFILE\Documents\memberberry\config.json",
        "$PSScriptRoot\..\..\memberberry\config.json",
        "$PSScriptRoot\..\memberberry\config.json"
    )
    
    foreach ($path in $possiblePaths) {
        if (Test-Path $path) {
            $script:MemberberryConfigPath = $path
            return $path
        }
    }
    
    return $null
}

<#
.SYNOPSIS
    Reads memberberry's config.json to get the shared data_path
.DESCRIPTION
    Reads memberberry's config.json file and extracts the data_path value.
    Returns the data_path if found, or $null if not found or config doesn't exist.
#>
function Get-MemberberryConfig {
    $configPath = Get-MemberberryConfigPath
    if (-not $configPath) {
        return $null
    }
    
    try {
        $configContent = Get-Content $configPath -Raw | ConvertFrom-Json
        if ($configContent.data_path -and $configContent.data_path -ne "") {
            $script:MemberberryDataPath = $configContent.data_path
            return $configContent
        }
    } catch {
        Write-Warning "Could not read memberberry config: $($_.Exception.Message)"
    }
    
    return $null
}

<#
.SYNOPSIS
    Gets the shared data path from memberberry config, or falls back to local storage
.DESCRIPTION
    Attempts to read memberberry's config.json to get the shared data_path.
    If not found or config doesn't exist, returns the default LOCALAPPDATA path for backward compatibility.
#>
function Get-VScanMagicDataPath {
    if ($script:MemberberryDataPath) {
        return $script:MemberberryDataPath
    }
    
    $config = Get-MemberberryConfig
    if ($config -and $config.data_path -and $config.data_path -ne "") {
        $dataPath = $config.data_path
        # Validate path exists or can be created
        if (-not (Test-Path $dataPath)) {
            try {
                New-Item -Path $dataPath -ItemType Directory -Force | Out-Null
                Write-Host "Created shared data directory: $dataPath" -ForegroundColor Green
            } catch {
                Write-Warning "Could not create shared data directory: $dataPath. Falling back to local storage."
                return $script:DefaultLocalAppDataPath
            }
        }
        $script:MemberberryDataPath = $dataPath
        return $dataPath
    }
    
    # Fallback to local storage (backward compatibility)
    return $script:DefaultLocalAppDataPath
}

<#
.SYNOPSIS
    Loads per-client data from shared storage
.DESCRIPTION
    Loads client-specific data including RMIT Plus settings and company folder mappings
    from the shared location specified in memberberry's config.json.
    Returns a hashtable with client data.
#>
function Get-VScanMagicClientData {
    $dataPath = Get-VScanMagicDataPath
    $clientDataFile = Join-Path $dataPath $script:ClientDataFileName
    
    if (-not (Test-Path $clientDataFile)) {
        # Return empty structure if file doesn't exist
        return @{
            RMITPlusSettings = @{}
            CompanyFolderMap = @{}
        }
    }
    
    try {
        # Validate file size (max 10MB)
        $fileSize = (Get-Item $clientDataFile).Length
        if ($fileSize -gt 10MB) {
            Write-Warning "Client data file too large: $([math]::Round($fileSize / 1MB, 2)) MB. Skipping load."
            return @{
                RMITPlusSettings = @{}
                CompanyFolderMap = @{}
            }
        }
        
        $json = Get-Content $clientDataFile -Raw | ConvertFrom-Json
        
        $result = @{
            RMITPlusSettings = @{}
            CompanyFolderMap = @{}
        }
        
        # Load RMIT Plus settings
        if ($json.RMITPlusSettings -and $json.RMITPlusSettings.PSObject.Properties) {
            foreach ($prop in $json.RMITPlusSettings.PSObject.Properties) {
                $result.RMITPlusSettings[$prop.Name] = $prop.Value
            }
        }
        
        # Load company folder mappings
        if ($json.CompanyFolderMap -and $json.CompanyFolderMap.PSObject.Properties) {
            foreach ($prop in $json.CompanyFolderMap.PSObject.Properties) {
                $result.CompanyFolderMap[$prop.Name] = $prop.Value
            }
        }
        
        return $result
    } catch {
        Write-Warning "Could not load client data from $clientDataFile : $($_.Exception.Message)"
        return @{
            RMITPlusSettings = @{}
            CompanyFolderMap = @{}
        }
    }
}

<#
.SYNOPSIS
    Saves per-client data to shared storage
.DESCRIPTION
    Saves client-specific data including RMIT Plus settings and company folder mappings
    to the shared location specified in memberberry's config.json.
    Implements file locking to prevent concurrent write conflicts.
.PARAMETER RMITPlusSettings
    Hashtable of client names to RMIT Plus boolean values
.PARAMETER CompanyFolderMap
    Hashtable of company names to folder paths
#>
function Save-VScanMagicClientData {
    param(
        [hashtable]$RMITPlusSettings = @{},
        [hashtable]$CompanyFolderMap = @{}
    )
    
    $dataPath = Get-VScanMagicDataPath
    $clientDataFile = Join-Path $dataPath $script:ClientDataFileName
    
    # Ensure directory exists
    if (-not (Test-Path $dataPath)) {
        try {
            New-Item -Path $dataPath -ItemType Directory -Force | Out-Null
        } catch {
            Write-Error "Could not create data directory: $dataPath"
            return $false
        }
    }
    
    # Implement simple file locking using a lock file
    $lockFile = "$clientDataFile.lock"
    $maxRetries = 10
    $retryDelay = 200  # milliseconds
    
    for ($retry = 0; $retry -lt $maxRetries; $retry++) {
        if (Test-Path $lockFile) {
            Start-Sleep -Milliseconds $retryDelay
            continue
        }
        
        try {
            # Create lock file
            $null = New-Item -Path $lockFile -ItemType File -Force
            
            # Read existing data to merge
            $existingData = Get-VScanMagicClientData
            
            # Merge with new data
            $mergedData = @{
                RMITPlusSettings = @{}
                CompanyFolderMap = @{}
            }
            
            # Merge RMIT Plus settings
            foreach ($key in $existingData.RMITPlusSettings.Keys) {
                $mergedData.RMITPlusSettings[$key] = $existingData.RMITPlusSettings[$key]
            }
            foreach ($key in $RMITPlusSettings.Keys) {
                $mergedData.RMITPlusSettings[$key] = $RMITPlusSettings[$key]
            }
            
            # Merge company folder mappings
            foreach ($key in $existingData.CompanyFolderMap.Keys) {
                $mergedData.CompanyFolderMap[$key] = $existingData.CompanyFolderMap[$key]
            }
            foreach ($key in $CompanyFolderMap.Keys) {
                $mergedData.CompanyFolderMap[$key] = $CompanyFolderMap[$key]
            }
            
            # Save merged data
            $mergedData | ConvertTo-Json -Depth 10 | Set-Content $clientDataFile -Encoding UTF8
            
            # Remove lock file
            Remove-Item -Path $lockFile -Force -ErrorAction SilentlyContinue
            
            return $true
        } catch {
            # Remove lock file on error
            Remove-Item -Path $lockFile -Force -ErrorAction SilentlyContinue
            
            if ($retry -eq $maxRetries - 1) {
                Write-Error "Failed to save client data after $maxRetries retries: $($_.Exception.Message)"
                return $false
            }
            
            Start-Sleep -Milliseconds $retryDelay
        }
    }
    
    return $false
}

<#
.SYNOPSIS
    Loads remediation rules from shared storage
.DESCRIPTION
    Loads vulnerability remediation rules from the shared location specified in memberberry's config.json.
    Returns an array of remediation rule objects.
#>
function Get-VScanMagicRemediationRules {
    $dataPath = Get-VScanMagicDataPath
    $rulesFile = Join-Path $dataPath $script:RemediationRulesFileName
    
    if (-not (Test-Path $rulesFile)) {
        return @()
    }
    
    try {
        # Validate file size (max 10MB)
        $fileSize = (Get-Item $rulesFile).Length
        if ($fileSize -gt 10MB) {
            Write-Warning "Remediation rules file too large: $([math]::Round($fileSize / 1MB, 2)) MB. Skipping load."
            return @()
        }
        
        $json = Get-Content $rulesFile -Raw | ConvertFrom-Json
        if ($json -is [System.Array]) {
            return @($json)
        } elseif ($json -is [PSCustomObject]) {
            # Handle single object or wrapped array
            if ($json.Rules) {
                return @($json.Rules)
            }
            return @($json)
        }
        
        return @()
    } catch {
        Write-Warning "Could not load remediation rules from $rulesFile : $($_.Exception.Message)"
        return @()
    }
}

<#
.SYNOPSIS
    Saves remediation rules to shared storage
.DESCRIPTION
    Saves vulnerability remediation rules to the shared location specified in memberberry's config.json.
    Implements file locking to prevent concurrent write conflicts.
.PARAMETER Rules
    Array of remediation rule objects to save
#>
function Save-VScanMagicRemediationRules {
    param(
        [array]$Rules = @()
    )
    
    $dataPath = Get-VScanMagicDataPath
    $rulesFile = Join-Path $dataPath $script:RemediationRulesFileName
    
    # Ensure directory exists
    if (-not (Test-Path $dataPath)) {
        try {
            New-Item -Path $dataPath -ItemType Directory -Force | Out-Null
        } catch {
            Write-Error "Could not create data directory: $dataPath"
            return $false
        }
    }
    
    # Implement simple file locking using a lock file
    $lockFile = "$rulesFile.lock"
    $maxRetries = 10
    $retryDelay = 200  # milliseconds
    
    for ($retry = 0; $retry -lt $maxRetries; $retry++) {
        if (Test-Path $lockFile) {
            Start-Sleep -Milliseconds $retryDelay
            continue
        }
        
        try {
            # Create lock file
            $null = New-Item -Path $lockFile -ItemType File -Force
            
            # Validate JSON structure before saving
            $testJson = $Rules | ConvertTo-Json -Depth 10
            $null = $testJson | ConvertFrom-Json  # Validate it can be parsed back
            
            # Save rules
            $Rules | ConvertTo-Json -Depth 10 | Set-Content $rulesFile -Encoding UTF8
            
            # Remove lock file
            Remove-Item -Path $lockFile -Force -ErrorAction SilentlyContinue
            
            return $true
        } catch {
            # Remove lock file on error
            Remove-Item -Path $lockFile -Force -ErrorAction SilentlyContinue
            
            if ($retry -eq $maxRetries - 1) {
                Write-Error "Failed to save remediation rules after $maxRetries retries: $($_.Exception.Message)"
                return $false
            }
            
            Start-Sleep -Milliseconds $retryDelay
        }
    }
    
    return $false
}

<#
.SYNOPSIS
    Gets the current storage location (shared or local)
.DESCRIPTION
    Returns information about where VScanMagic data is currently being stored.
#>
function Get-VScanMagicStorageInfo {
    $dataPath = Get-VScanMagicDataPath
    $isShared = ($dataPath -ne $script:DefaultLocalAppDataPath)
    
    return @{
        DataPath = $dataPath
        IsShared = $isShared
        MemberberryConfigPath = Get-MemberberryConfigPath
    }
}

<#
.SYNOPSIS
    Loads covered software list from shared storage
.DESCRIPTION
    Loads covered software patterns from the shared location specified in memberberry's config.json.
    Returns an array of covered software objects.
#>
function Get-VScanMagicCoveredSoftware {
    $dataPath = Get-VScanMagicDataPath
    $coveredSoftwareFile = Join-Path $dataPath $script:CoveredSoftwareFileName
    
    if (-not (Test-Path $coveredSoftwareFile)) {
        return @()
    }
    
    try {
        # Validate file size (max 10MB)
        $fileSize = (Get-Item $coveredSoftwareFile).Length
        if ($fileSize -gt 10MB) {
            Write-Warning "Covered software file too large: $([math]::Round($fileSize / 1MB, 2)) MB. Skipping load."
            return @()
        }
        
        $json = Get-Content $coveredSoftwareFile -Raw | ConvertFrom-Json
        if ($json -is [System.Array]) {
            return @($json)
        } elseif ($json -is [PSCustomObject]) {
            if ($json.CoveredSoftware) {
                return @($json.CoveredSoftware)
            }
            return @($json)
        }
        
        return @()
    } catch {
        Write-Warning "Could not load covered software from $coveredSoftwareFile : $($_.Exception.Message)"
        return @()
    }
}

<#
.SYNOPSIS
    Saves covered software list to shared storage
.DESCRIPTION
    Saves covered software patterns to the shared location specified in memberberry's config.json.
    Implements file locking to prevent concurrent write conflicts.
.PARAMETER CoveredSoftware
    Array of covered software objects to save
#>
function Save-VScanMagicCoveredSoftware {
    param(
        [array]$CoveredSoftware = @()
    )
    
    $dataPath = Get-VScanMagicDataPath
    $coveredSoftwareFile = Join-Path $dataPath $script:CoveredSoftwareFileName
    
    # Ensure directory exists
    if (-not (Test-Path $dataPath)) {
        try {
            New-Item -Path $dataPath -ItemType Directory -Force | Out-Null
        } catch {
            Write-Error "Could not create data directory: $dataPath"
            return $false
        }
    }
    
    # Implement simple file locking using a lock file
    $lockFile = "$coveredSoftwareFile.lock"
    $maxRetries = 10
    $retryDelay = 200  # milliseconds
    
    for ($retry = 0; $retry -lt $maxRetries; $retry++) {
        if (Test-Path $lockFile) {
            Start-Sleep -Milliseconds $retryDelay
            continue
        }
        
        try {
            # Create lock file
            $null = New-Item -Path $lockFile -ItemType File -Force
            
            # Validate JSON structure before saving
            $testJson = $CoveredSoftware | ConvertTo-Json -Depth 10
            $null = $testJson | ConvertFrom-Json  # Validate it can be parsed back
            
            # Save covered software
            $CoveredSoftware | ConvertTo-Json -Depth 10 | Set-Content $coveredSoftwareFile -Encoding UTF8
            
            # Remove lock file
            Remove-Item -Path $lockFile -Force -ErrorAction SilentlyContinue
            
            return $true
        } catch {
            # Remove lock file on error
            Remove-Item -Path $lockFile -Force -ErrorAction SilentlyContinue
            
            if ($retry -eq $maxRetries - 1) {
                Write-Error "Failed to save covered software after $maxRetries retries: $($_.Exception.Message)"
                return $false
            }
            
            Start-Sleep -Milliseconds $retryDelay
        }
    }
    
    return $false
}

<#
.SYNOPSIS
    Loads general recommendations from shared storage
.DESCRIPTION
    Loads general recommendations from the shared location specified in memberberry's config.json.
    Returns an array of recommendation objects.
#>
function Get-VScanMagicGeneralRecommendations {
    $dataPath = Get-VScanMagicDataPath
    $generalRecommendationsFile = Join-Path $dataPath $script:GeneralRecommendationsFileName
    
    if (-not (Test-Path $generalRecommendationsFile)) {
        return @()
    }
    
    try {
        # Validate file size (max 10MB)
        $fileSize = (Get-Item $generalRecommendationsFile).Length
        if ($fileSize -gt 10MB) {
            Write-Warning "General recommendations file too large: $([math]::Round($fileSize / 1MB, 2)) MB. Skipping load."
            return @()
        }
        
        $json = Get-Content $generalRecommendationsFile -Raw | ConvertFrom-Json
        if ($json -is [System.Array]) {
            return @($json)
        } elseif ($json -is [PSCustomObject]) {
            if ($json.GeneralRecommendations) {
                return @($json.GeneralRecommendations)
            }
            return @($json)
        }
        
        return @()
    } catch {
        Write-Warning "Could not load general recommendations from $generalRecommendationsFile : $($_.Exception.Message)"
        return @()
    }
}

<#
.SYNOPSIS
    Saves general recommendations to shared storage
.DESCRIPTION
    Saves general recommendations to the shared location specified in memberberry's config.json.
    Implements file locking to prevent concurrent write conflicts.
.PARAMETER GeneralRecommendations
    Array of general recommendation objects to save
#>
function Save-VScanMagicGeneralRecommendations {
    param(
        [array]$GeneralRecommendations = @()
    )
    
    $dataPath = Get-VScanMagicDataPath
    $generalRecommendationsFile = Join-Path $dataPath $script:GeneralRecommendationsFileName
    
    # Ensure directory exists
    if (-not (Test-Path $dataPath)) {
        try {
            New-Item -Path $dataPath -ItemType Directory -Force | Out-Null
        } catch {
            Write-Error "Could not create data directory: $dataPath"
            return $false
        }
    }
    
    # Implement simple file locking using a lock file
    $lockFile = "$generalRecommendationsFile.lock"
    $maxRetries = 10
    $retryDelay = 200  # milliseconds
    
    for ($retry = 0; $retry -lt $maxRetries; $retry++) {
        if (Test-Path $lockFile) {
            Start-Sleep -Milliseconds $retryDelay
            continue
        }
        
        try {
            # Create lock file
            $null = New-Item -Path $lockFile -ItemType File -Force
            
            # Validate JSON structure before saving
            $testJson = $GeneralRecommendations | ConvertTo-Json -Depth 10
            $null = $testJson | ConvertFrom-Json  # Validate it can be parsed back
            
            # Save general recommendations
            $GeneralRecommendations | ConvertTo-Json -Depth 10 | Set-Content $generalRecommendationsFile -Encoding UTF8
            
            # Remove lock file
            Remove-Item -Path $lockFile -Force -ErrorAction SilentlyContinue
            
            return $true
        } catch {
            # Remove lock file on error
            Remove-Item -Path $lockFile -Force -ErrorAction SilentlyContinue
            
            if ($retry -eq $maxRetries - 1) {
                Write-Error "Failed to save general recommendations after $maxRetries retries: $($_.Exception.Message)"
                return $false
            }
            
            Start-Sleep -Milliseconds $retryDelay
        }
    }
    
    return $false
}

# Export module members
Export-ModuleMember -Function @(
    'Get-MemberberryConfig',
    'Get-VScanMagicDataPath',
    'Get-VScanMagicClientData',
    'Save-VScanMagicClientData',
    'Get-VScanMagicRemediationRules',
    'Save-VScanMagicRemediationRules',
    'Get-VScanMagicCoveredSoftware',
    'Save-VScanMagicCoveredSoftware',
    'Get-VScanMagicGeneralRecommendations',
    'Save-VScanMagicGeneralRecommendations',
    'Get-VScanMagicStorageInfo'
)
