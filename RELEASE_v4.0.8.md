# VScanMagic v4.0.8 Release Notes

VScanMagic v4.0.8 - Filename shortening, auto-resize Excel columns, Output Options fix, Windows/VMware first-party.

## 📦 Installation

1. Download `VScanMagic.zip` from the release
2. Extract `VScanMagic.exe` along with `VScanMagic-Modules` and `ConnectSecure-API.ps1` to the same folder
3. Run `VScanMagic.exe`

## 📝 Changelog

- **Filename shortening**: Download and generated report paths truncated when exceeding ~250 chars (Windows MAX_PATH); company name shortened to fit
- **Auto-resize Excel columns**: Downloaded Excel reports auto-resize columns (excludes Company, Proposed Remediations sheets); setting in Output Options to disable
- **Output Options dialog**: Fixed OK/Cancel button layout and form height
- **First-party vendors**: Windows 11, Windows 10, Windows Server, VMware, vSphere, VMware Tools now treated as first-party (not 3rd party)
- Bump version to 4.0.8
