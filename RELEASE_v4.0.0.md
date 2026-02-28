# VScanMagic v4.0.0 Release Notes

VScanMagic v4 introduces a new major version with improved UI layout, all five standard reports in Download Standard Reports, and enhanced download reliability.

## ✨ New in v4.0.0

**Download Standard Reports – All 5 Reports:**
- ✅ **All Vulnerabilities Report (XLSX)**
- ✅ **Suppressed Vulnerabilities (XLSX)**
- ✅ **External Scan (XLSX)**
- ✅ **Executive Summary Report (DOCX)** – now included in Download Standard Reports
- ✅ **Pending Remediation EPSS Score Reports (XLSX)** – now included in Download Standard Reports

**UI Layout Improvements:**
- 📐 **Fixed section overlap** – Section 1 (Download from ConnectSecure) no longer overlaps Section 2 (Or process a previously downloaded file)
- 📐 **Download button placement** – Download Standard Reports button fully contained within Section 1
- 📐 **Adjusted spacing** – Improved vertical layout for all form sections

**Reliability & Usability:**
- 🔄 **Retry logic** – Up to 3 attempts with exponential backoff for transient download failures (timeout, connection, 404)
- 💾 **Last output directory** – Saves and prefills the last used output directory
- 📊 **Accurate counts** – Succeeded/Failed counts display correctly after downloads

## 🔄 Migration from v3.x

- ✅ No breaking changes
- ✅ Existing settings and configurations remain compatible
- ✅ Same ConnectSecure API integration

## 📦 Installation

1. Download `VScanMagic.zip` from the release
2. Extract `VScanMagic.exe` along with `VScanMagic-Modules` and `ConnectSecure-API.ps1` to the same folder
3. Run `VScanMagic.exe`

## 📝 Changelog

- Add pending-epss and executive-summary to Download Standard Reports (5 reports total)
- Fix Section 1/2 layout overlap; increase Section 1 height, shift Section 2 and lower controls
- Add retry logic (3 attempts) for ConnectSecure downloads
- Save and prefill last output directory
- Bump version to 4.0.0
