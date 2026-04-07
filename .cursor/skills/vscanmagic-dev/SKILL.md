---
name: vscanmagic-dev
description: Guides development for the VScanMagic PowerShell vulnerability management tool. Use when editing VScanMagic scripts, adding features, debugging ConnectSecure API issues, writing tests, or working with VScanMagic-Modules, archive scripts, or release notes.
---

# VScanMagic Development

## Test Scripts and Archive

**Always create test scripts in `archive/`.** The archive folder is gitignored.

- Test scripts: `archive/Test-*.ps1`, `archive/company-review-tests/`
- One-off utilities and debug scripts: `archive/Capture-*.ps1`, etc.
- Do not create test scripts in the project root or other tracked directories.

## Archive Script Pattern

Scripts in `archive/` that need project modules must:

1. Resolve project root: `$projectRoot = Split-Path -Parent $PSScriptRoot`
2. Dot-source Core and API (or other modules as needed):

```powershell
$corePath = Join-Path $projectRoot "VScanMagic-Modules\VScanMagic-Core.ps1"
$apiPath = Join-Path $projectRoot "ConnectSecure-API.ps1"
if (-not (Test-Path $corePath) -or -not (Test-Path $apiPath)) {
    Write-Host "VScanMagic project not found. Run from archive/ inside vscanmagic repo." -ForegroundColor Red
    exit 1
}
. $corePath
. $apiPath
```

3. Use saved credentials: `$creds = Load-ConnectSecureCredentials` then `Connect-ConnectSecureAPI` with BaseUrl, TenantName, ClientId, ClientSecret.

## Module Load Order

Main GUI loads modules in this order (do not change):

1. VScanMagic-Core.ps1
2. VScanMagic-Data.ps1
3. VScanMagic-Reports.ps1
4. VScanMagic-Dialogs.ps1
5. VScanMagic-Form.ps1

Paths: `Join-Path $script:ScriptDirectory "VScanMagic-Modules"` then `Join-Path $modulesDir "VScanMagic-Core.ps1"` etc.

## REST API (`VScanMagic-API.ps1`)

Dot-sources **`VScanMagic-ApiBootstrap.ps1`** (Core → Data → Reports only). Does **not** load Dialogs, Form, or Memberberry. If bootstrap is missing, falls back to `VScanMagic-GUI.ps1`. Sets `$script:IsApiMode = $true` and `$script:ScriptDirectory` before loading modules.

## PowerShell: Single-Item Array Unwrapping

PowerShell unwraps single-element arrays. A function returning `@($singleObject)` may be received as the object, not an array.

**Impact:** `$result -is [array]` can be false; `$result.Count` can be null.

**Mitigations:**

| Location | Approach |
|----------|----------|
| Caller | `$data = @(Get-SomeFunction)` |
| Function (single object) | `return @(,$item)` |
| Function (array) | `return ,$array` |
| Normalize in function | See below |

```powershell
$raw = Invoke-Something
if ($null -eq $raw) { return @() }
if ($raw -is [array]) { return @($raw) }
return @(,$raw)  # single object -> array of 1
```

**Relevant:** `ConnectSecure-API.ps1` Company Review functions normalize returns: `Get-ConnectSecureLightweightAssets`, `Get-ConnectSecureCompanyAgents`, `Get-ConnectSecureAgentCredentialsMapping`, `Get-ConnectSecureAgentDiscoveryMapping`, `Get-ConnectSecureDiscoverySettings`, `Get-ConnectSecureExternalScanDiscoverySettings`, `Get-ConnectSecureAssetFirewallPolicy`, `jobsView` in `Get-ConnectSecureCompanyReviewData`.

## ConnectSecure API

- Credentials: BaseUrl, TenantName, ClientId, ClientSecret (from `Load-ConnectSecureCredentials`)
- Rate limiting: `$script:ConnectSecureConfig.RateLimit`; use `Test-RateLimit`, `Wait-ForRateLimit`, `Add-RequestToHistory`
- Logging: `Write-CSApiLog -Message "..." -Level Info|Warning|Error|Success`
- Sensitive data: `Remove-SensitiveDataFromObject` before logging

## Release Notes Format

Create `RELEASE_vX.Y.Z.md` with:

```markdown
# VScanMagic vX.Y.Z Release Notes

VScanMagic vX.Y.Z - [Brief tagline].

## 📦 Installation

1. Download `VScanMagic.zip` from the release
2. Extract `VScanMagic.exe` along with `VScanMagic-Modules` and `ConnectSecure-API.ps1` to the same folder
3. Run `VScanMagic.exe`

## 📝 Changelog

- **[Feature/fix]**: Description
- Bump version to X.Y.Z
```

Update version in `VScanMagic-GUI.ps1` (header) and `BuildExeFinal.ps1` (title, version).

## Build

- EXE build: `BuildExeFinal.ps1` uses ps2exe from `Modules/ps2exe/1.0.17/`
- Input: `VScanMagic-GUI.ps1` → Output: `VScanMagic.exe`
- Requires: `VScanMagic.ico`, `VScanMagic-Modules/`, `ConnectSecure-API.ps1` in same folder as exe
