# VScanMagic v4.0.2 Release Notes

VScanMagic v4.0.2 adds structured output paths, processing summaries with folder links, report folder history, and bulk processing improvements.

## ✨ New in v4.0.2

**Structured Output Paths:**
- 📁 **Network Documentation / Vulnerability Scans** – Output paths now consistently include this subfolder (auto-appended when missing)
- 📁 **Quarter folder format** – Changed from `Q1 - 2026` to `2026 - Q1` (year first)
- 📁 **Company Folder Mappings** – Auto-normalizes legacy mappings to include the standard subpath; Edit dialog guides to select Vulnerability Scans folder

**Processing Summary & History:**
- 📋 **Processing Summary** – After report generation, a summary dialog shows links to all output folders (company + quarter path)
- 📋 **Report Folder History** – Persistent history of processed report folders (button in Output section); double-click to open in Explorer; Remove/Clear options
- 📋 **Both bulk and single-company modes** – Summary and history work in all flows

**Bulk Processing:**
- ⚡ **Skip follow-up dialogs** – When checked and 2+ companies selected: skips General Recommendations, Hostname Review, Time Estimate dialogs, and completion/error popups
- ⚡ **Uses defaults** – Saved recommendations, all hostnames, 0 hours per item when skipping

**UI Changes:**
- 🗑️ **Removed View Generated Reports** – Word, Excel, Email, Ticket, Time Estimate buttons removed (use Processing Summary and Report Folder History instead)
- 📐 **Recovered space** – Form height reduced; Processing Log spacing fixed
- 📐 **Layout fix** – Output section controls no longer overlap

## 🔄 Migration from v4.0.1

- ✅ No breaking changes
- ✅ Company Folder Mappings auto-updated to include Network Documentation\Vulnerability Scans
- ✅ Report Folder History builds as you process (new feature)

## 📦 Installation

1. Download `VScanMagic.zip` from the release
2. Extract `VScanMagic.exe` along with `VScanMagic-Modules` and `ConnectSecure-API.ps1` to the same folder
3. Run `VScanMagic.exe`

## 📝 Changelog

- Add Resolve-VulnerabilityScansSubpath for consistent Network Documentation\Vulnerability Scans in output paths
- Change quarter folder format from Q1 - 2026 to 2026 - Q1
- Add Processing Summary dialog after report generation (links to output folders)
- Add Report Folder History (persistent, accessible via button)
- Skip completion/error popups when bulk processing with Skip follow-up dialogs
- Remove View Generated Reports section; recover form space
- Fix Output/Processing Log overlap; improve layout spacing
- Auto-normalize Company Folder Mappings on load and in Edit dialog
- Bump version to 4.0.2
