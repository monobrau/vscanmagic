# VScanMagic Memberberry Integration Setup Guide

This guide explains how to set up VScanMagic to use Memberberry's shared storage system for client data and remediation rules.

## Prerequisites

1. **Memberberry must be installed** and configured with a shared `data_path`
2. **VScanMagic** must have the `MemberberryIntegration.psm1` module (included in Phase 1 integration)

## Step 1: Configure Memberberry

Memberberry needs to be configured with a shared network location for multi-user access.

1. Navigate to your Memberberry installation directory (typically `c:\git\memberberry`)
2. Create or edit `config.json`:
   ```json
   {
     "data_path": "\\\\server\\share\\memberberry-data"
   }
   ```
   Or use a local shared folder:
   ```json
   {
     "data_path": "C:\\Shared\\memberberry-data"
   }
   ```

3. Ensure the `data_path` directory exists and is accessible to all users who will use VScanMagic

## Step 2: Verify VScanMagic Can Find Memberberry

VScanMagic automatically searches for Memberberry's `config.json` in these locations (in order):
1. **`MEMBERBERRY_CONFIG`** — full path to `config.json` (recommended for portability)
2. **`MEMBERBERRY_HOME`** — directory containing `config.json`
3. `c:\git\memberberry\config.json`
4. `%USERPROFILE%\Documents\memberberry\config.json`
5. Relative paths from VScanMagic installation

**To verify VScanMagic found Memberberry:**
1. Launch VScanMagic GUI
2. Look for one of these messages in the console:
   - ✅ `"Memberberry integration module loaded successfully"` (Green) - Success!
   - ⚠️ `"MemberberryIntegration module not found. Using local storage."` (Yellow) - Module missing
   - ⚠️ `"Could not load MemberberryIntegration module: ..."` (Warning) - Error loading

## Step 3: Migrate Existing Data (Optional but Recommended)

If you have existing VScanMagic data in local storage (`%LOCALAPPDATA%\VScanMagic\`), migrate it to shared storage:

1. Open PowerShell in the VScanMagic directory
2. Run the migration script:
   ```powershell
   .\MigrateToMemberberryStorage.ps1
   ```

**Migration Options:**
- **Backup original files**: Creates a backup before migrating (default: enabled)
- **Delete original files**: Removes local files after successful migration (default: disabled)
- **Custom config path**: Specify memberberry config location if not auto-detected

**Example with options:**
```powershell
# Migrate with backup, keep originals
.\MigrateToMemberberryStorage.ps1 -BackupOriginal

# Migrate with backup, delete originals after migration
.\MigrateToMemberberryStorage.ps1 -BackupOriginal -DeleteOriginal

# Specify custom memberberry config path
.\MigrateToMemberberryStorage.ps1 -MemberberryConfigPath "D:\Tools\memberberry\config.json"
```

## Step 4: Verify Integration is Working

After setup, verify that VScanMagic is using shared storage:

### Check Storage Location
1. Launch VScanMagic GUI
2. Look for console messages indicating shared storage:
   - `"Remediation rules loaded from shared storage: \\server\share\memberberry-data"` (Green)
   - `"Company folder mappings loaded from shared storage: \\server\share\memberberry-data"` (Green)

### Test Data Persistence
1. **Test Remediation Rules:**
   - Open VScanMagic → Settings → Remediation Rules
   - Add or modify a rule
   - Save and close
   - Verify the file exists in: `{data_path}\vscanmagic-remediation-rules.json`

2. **Test Company Folder Mappings:**
   - Open VScanMagic → Settings → Company Folder Mappings
   - Add a mapping
   - Save
   - Verify the file exists in: `{data_path}\vscanmagic-clients.json`

3. **Test RMIT Plus Settings:**
   - When setting client types (RMIT+), settings are automatically saved
   - Verify in: `{data_path}\vscanmagic-clients.json` under `RMITPlusSettings`

## Storage Locations

### Shared Storage (when Memberberry integration is active)
- **Client Data**: `{memberberry_data_path}\vscanmagic-clients.json`
  - Contains: RMIT Plus settings, Company folder mappings
- **Remediation Rules**: `{memberberry_data_path}\vscanmagic-remediation-rules.json`
  - Contains: Product-specific remediation guidance

### Local Storage (fallback, when Memberberry not found)
- **Location**: `%LOCALAPPDATA%\VScanMagic\`
- **Files**: `VScanMagic_RemediationRules.json`, `VScanMagic_CompanyFolderMap.json`

## Troubleshooting

### Issue: "MemberberryIntegration module not found"
**Solution:**
- Verify `Modules\MemberberryIntegration.psm1` exists in VScanMagic directory
- Check file permissions (script execution policy may block module loading)

### Issue: "Could not load MemberberryIntegration module"
**Solution:**
- Check PowerShell execution policy: `Get-ExecutionPolicy`
- If restricted, run: `Set-ExecutionPolicy RemoteSigned -Scope CurrentUser`
- Check for syntax errors in the module file

### Issue: "Using local storage" (but Memberberry is configured)
**Solution:**
- Verify Memberberry's `config.json` exists and has `data_path` set
- Check that `data_path` is accessible (network path is reachable)
- Verify `data_path` directory exists (VScanMagic will create it if it has permissions)

### Issue: "Could not create shared data directory"
**Solution:**
- Check write permissions on the `data_path` location
- Ensure network share is accessible
- Try creating the directory manually: `New-Item -Path "{data_path}" -ItemType Directory`

### Issue: Migration script fails
**Solution:**
- Ensure you have read access to local VScanMagic data
- Ensure you have write access to shared storage location
- Check that Memberberry config is accessible
- Review error messages for specific file access issues

## Multi-User Setup

For multiple users to share the same data:

1. **All users** must point to the same Memberberry `data_path` in their `config.json`
2. **All users** must have read/write access to the shared network location
3. **First user** should run the migration script to populate shared storage
4. **Other users** can start using VScanMagic immediately (will read from shared storage)

**Important:** File locking is implemented to prevent data corruption when multiple users access shared files simultaneously. If you see lock files (`.lock` extension), wait a few seconds and try again.

## Verifying Storage Type

To check which storage type VScanMagic is currently using, you can run this PowerShell command:

```powershell
Import-Module ".\Modules\MemberberryIntegration.psm1" -Force
$info = Get-VScanMagicStorageInfo
Write-Host "Data Path: $($info.DataPath)"
Write-Host "Is Shared: $($info.IsShared)"
Write-Host "Memberberry Config: $($info.MemberberryConfigPath)"
```

## Rollback (if needed)

If you need to revert to local storage:

1. Remove or rename Memberberry's `config.json` (or clear the `data_path` field)
2. VScanMagic will automatically fall back to local storage
3. Your local data remains in `%LOCALAPPDATA%\VScanMagic\`

## Support

If you encounter issues:
1. Check the console output for error messages
2. Verify Memberberry configuration
3. Test network path accessibility
4. Review file permissions on shared storage location
