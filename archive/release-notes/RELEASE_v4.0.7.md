# VScanMagic v4.0.7 Release Notes

VScanMagic v4.0.7 - ConnectSecure username lookup, Word 255-char fix, VMware first-party.

## 📦 Installation

1. Download `VScanMagic.zip` from the release
2. Extract `VScanMagic.exe` along with `VScanMagic-Modules` and `ConnectSecure-API.ps1` to the same folder
3. Run `VScanMagic.exe`

## 📝 Changelog

- **ConnectSecure username lookup**: Hostname Review dialog can fetch usernames from ConnectSecure asset view; "Lookup from ConnectSecure" button (enabled when processing ConnectSecure download data)
- **Auto-lookup setting**: New Output Options checkbox "Look up usernames from ConnectSecure by default" (on by default) – automatically populates usernames when Hostname Review opens
- **Word report fix**: Resolved "String is longer than 255 characters" error – Add-WordText helper for long content; save-to-temp workaround for long paths and OneDrive/sync folders
- **First-party vendors**: VMware, vSphere, VMware Tools now treated as first-party (not 3rd party) in time estimate
- Bump version to 4.0.7
