<#
.SYNOPSIS
    Migrates VScanMagic data from local storage to memberberry shared storage
.DESCRIPTION
    One-time migration script to move existing VScanMagic client data and remediation rules
    from local storage (%LOCALAPPDATA%\VScanMagic\) to shared location specified in memberberry's config.json.
.EXAMPLE
    .\MigrateToMemberberryStorage.ps1
#>

param(
    [Parameter(Mandatory=$false)]
    [string]$MemberberryConfigPath = "",
    
    [Parameter(Mandatory=$false)]
    [switch]$BackupOriginal = $true,
    
    [Parameter(Mandatory=$false)]
    [switch]$DeleteOriginal = $false
)

# Import MemberberryIntegration module
$modulePath = Join-Path $PSScriptRoot "Modules\MemberberryIntegration.psm1"
if (-not (Test-Path $modulePath)) {
    Write-Error "MemberberryIntegration module not found at: $modulePath"
    exit 1
}

Import-Module $modulePath -Force

# Get local storage paths
$localAppDataPath = Join-Path $env:LOCALAPPDATA "VScanMagic"
$localRemediationRulesPath = Join-Path $localAppDataPath "VScanMagic_RemediationRules.json"
$localCompanyFolderMapPath = Join-Path $localAppDataPath "VScanMagic_CompanyFolderMap.json"

# Get shared storage path
if (-not [string]::IsNullOrWhiteSpace($MemberberryConfigPath)) {
    if (-not (Test-Path $MemberberryConfigPath)) {
        Write-Error "Specified memberberry config path does not exist: $MemberberryConfigPath"
        exit 1
    }
    # Temporarily override the config path
    $script:MemberberryConfigPath = $MemberberryConfigPath
}

$sharedDataPath = Get-VScanMagicDataPath
$storageInfo = Get-VScanMagicStorageInfo

Write-Host "`n=== VScanMagic Data Migration to Memberberry Shared Storage ===" -ForegroundColor Cyan
Write-Host "Local Storage: $localAppDataPath" -ForegroundColor Yellow
Write-Host "Shared Storage: $sharedDataPath" -ForegroundColor Green
Write-Host "Using Shared Storage: $($storageInfo.IsShared)" -ForegroundColor $(if ($storageInfo.IsShared) { "Green" } else { "Yellow" })
Write-Host ""

if (-not $storageInfo.IsShared) {
    Write-Warning "Memberberry config not found or data_path not set. Migration will copy to local storage location."
    $response = Read-Host "Continue anyway? (y/n)"
    if ($response -ne "y" -and $response -ne "Y") {
        Write-Host "Migration cancelled." -ForegroundColor Yellow
        exit 0
    }
}

# Ensure shared directory exists
if (-not (Test-Path $sharedDataPath)) {
    try {
        New-Item -Path $sharedDataPath -ItemType Directory -Force | Out-Null
        Write-Host "Created shared data directory: $sharedDataPath" -ForegroundColor Green
    } catch {
        Write-Error "Could not create shared data directory: $sharedDataPath"
        exit 1
    }
}

# Create backup directory if requested
$backupPath = $null
if ($BackupOriginal) {
    $backupPath = Join-Path $localAppDataPath "Backup_Before_Migration_$(Get-Date -Format 'yyyyMMdd_HHmmss')"
    try {
        New-Item -Path $backupPath -ItemType Directory -Force | Out-Null
        Write-Host "Created backup directory: $backupPath" -ForegroundColor Green
    } catch {
        Write-Warning "Could not create backup directory: $backupPath"
        $BackupOriginal = $false
    }
}

$migrationSuccess = $true
$filesMigrated = 0

# Migrate Remediation Rules
Write-Host "`n[1/2] Migrating Remediation Rules..." -ForegroundColor Cyan
if (Test-Path $localRemediationRulesPath) {
    try {
        $rules = Get-Content $localRemediationRulesPath -Raw | ConvertFrom-Json
        if ($rules -and $rules.Count -gt 0) {
            if (Save-VScanMagicRemediationRules -Rules $rules) {
                Write-Host "  ✓ Remediation rules migrated successfully" -ForegroundColor Green
                $filesMigrated++
                
                if ($BackupOriginal) {
                    Copy-Item -Path $localRemediationRulesPath -Destination (Join-Path $backupPath "VScanMagic_RemediationRules.json") -Force
                }
                
                if ($DeleteOriginal) {
                    Remove-Item -Path $localRemediationRulesPath -Force
                    Write-Host "  ✓ Original file deleted" -ForegroundColor Yellow
                }
            } else {
                Write-Warning "  ✗ Failed to save remediation rules to shared storage"
                $migrationSuccess = $false
            }
        } else {
            Write-Host "  - No remediation rules found in local storage" -ForegroundColor Gray
        }
    } catch {
        Write-Warning "  ✗ Error migrating remediation rules: $($_.Exception.Message)"
        $migrationSuccess = $false
    }
} else {
    Write-Host "  - No local remediation rules file found (skipping)" -ForegroundColor Gray
}

# Migrate Company Folder Map
Write-Host "`n[2/2] Migrating Company Folder Mappings..." -ForegroundColor Cyan
if (Test-Path $localCompanyFolderMapPath) {
    try {
        $companyFolderMap = Get-Content $localCompanyFolderMapPath -Raw | ConvertFrom-Json
        $folderMapHashtable = @{}
        if ($companyFolderMap -and $companyFolderMap.PSObject.Properties) {
            foreach ($prop in $companyFolderMap.PSObject.Properties) {
                $folderMapHashtable[$prop.Name] = $prop.Value
            }
        }
        
        if ($folderMapHashtable.Count -gt 0) {
            # Load existing client data to preserve RMIT Plus settings (if any)
            $existingData = Get-VScanMagicClientData
            $rmitPlusSettings = if ($existingData -and $existingData.RMITPlusSettings) { $existingData.RMITPlusSettings } else { @{} }
            
            if (Save-VScanMagicClientData -RMITPlusSettings $rmitPlusSettings -CompanyFolderMap $folderMapHashtable) {
                Write-Host "  ✓ Company folder mappings migrated successfully" -ForegroundColor Green
                $filesMigrated++
                
                if ($BackupOriginal) {
                    Copy-Item -Path $localCompanyFolderMapPath -Destination (Join-Path $backupPath "VScanMagic_CompanyFolderMap.json") -Force
                }
                
                if ($DeleteOriginal) {
                    Remove-Item -Path $localCompanyFolderMapPath -Force
                    Write-Host "  ✓ Original file deleted" -ForegroundColor Yellow
                }
            } else {
                Write-Warning "  ✗ Failed to save company folder mappings to shared storage"
                $migrationSuccess = $false
            }
        } else {
            Write-Host "  - No company folder mappings found in local storage" -ForegroundColor Gray
        }
    } catch {
        Write-Warning "  ✗ Error migrating company folder mappings: $($_.Exception.Message)"
        $migrationSuccess = $false
    }
} else {
    Write-Host "  - No local company folder map file found (skipping)" -ForegroundColor Gray
}

# Summary
Write-Host "`n=== Migration Summary ===" -ForegroundColor Cyan
if ($migrationSuccess -and $filesMigrated -gt 0) {
    Write-Host "✓ Migration completed successfully!" -ForegroundColor Green
    Write-Host "  Files migrated: $filesMigrated" -ForegroundColor Green
    if ($BackupOriginal -and $backupPath) {
        Write-Host "  Backup location: $backupPath" -ForegroundColor Yellow
    }
    if (-not $DeleteOriginal) {
        Write-Host "`nNote: Original files are still in local storage. Delete them manually if desired." -ForegroundColor Yellow
    }
} elseif ($filesMigrated -eq 0) {
    Write-Host "⚠ No files found to migrate." -ForegroundColor Yellow
} else {
    Write-Host "✗ Migration completed with errors. Please review the output above." -ForegroundColor Red
    exit 1
}

Write-Host "`nNext steps:" -ForegroundColor Cyan
Write-Host "1. Verify data in shared storage: $sharedDataPath" -ForegroundColor White
Write-Host "2. Test VScanMagic to ensure it loads data from shared storage" -ForegroundColor White
Write-Host "3. If everything works, you can delete original files from: $localAppDataPath" -ForegroundColor White
Write-Host ""
