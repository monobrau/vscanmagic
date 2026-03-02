# VScanMagic v4.0.4 Release Notes

VScanMagic v4.0.4 - Excel report generation fix for "Unable to get the Open property of the Workbooks class".

## 📦 Installation

1. Download `VScanMagic.zip` from the release
2. Extract `VScanMagic.exe` along with `VScanMagic-Modules` and `ConnectSecure-API.ps1` to the same folder
3. Run `VScanMagic.exe`

## 🐛 Bug Fixes

- **Excel report generation**: Fixed "Unable to get the Open property of the Workbooks class" error when generating Excel reports after reading vulnerability data. Changes:
  - Ensure system profile Desktop folder exists (required by Excel COM in some contexts)
  - Use single-argument Workbooks.Open to avoid COM parameter binding issues
  - Explicit GC before report generation to allow previous Excel instance to fully release

## 📝 Changelog

- Fix Excel COM automation error during report generation
- Bump version to 4.0.4
