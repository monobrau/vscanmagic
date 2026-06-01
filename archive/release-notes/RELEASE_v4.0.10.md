# VScanMagic v4.0.10 Release Notes

VScanMagic v4.0.10 - .NET remediation guidance, UI tweaks, Remediation Time Estimate rename.

## 📦 Installation

1. Download `VScanMagic.zip` from the release
2. Extract `VScanMagic.exe` along with `VScanMagic-Modules` and `ConnectSecure-API.ps1` to the same folder
3. Run `VScanMagic.exe`

## 📝 Changelog

- **.NET remediation rules**: Comprehensive guidance for .NET Framework (legacy) vs modern .NET (5, 6, 7, 8, 9)
  - Framework: Cannot upgrade in-place; migration required; older versions (3.5, 4.0) more painful; main driver is security patches, not user experience
  - Modern .NET: Drop-in upgradeable; LTS (3 years) vs non-LTS (18 months); MSP angle: value is security patches, not user-facing improvements
- **Show/Hide details**: ASCII symbols (+ / -) instead of Unicode arrows for better compatibility
- **Remediation Time Estimate**: Renamed from "Vulnerability Time Estimate" in txt and HTML reports
- Bump version to 4.0.10
