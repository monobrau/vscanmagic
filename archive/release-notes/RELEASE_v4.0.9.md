# VScanMagic v4.0.9 Release Notes

VScanMagic v4.0.9 - AI improvements, workflow popups removed, Report Filters persistence, Ripple20/placeholder fix, pie chart stability.

## 📦 Installation

1. Download `VScanMagic.zip` from the release
2. Extract `VScanMagic.exe` along with `VScanMagic-Modules` and `ConnectSecure-API.ps1` to the same folder
3. Run `VScanMagic.exe`

## 📝 Changelog

- **AI Improve Selected**: Uses single-item API path (fixes single-character output bug)
- **AI Improve All**: Chunked batch processing (4 items/chunk) with 429 retry; longer delays between chunks; ITEM # prefix stripped from output
- **Report Filters persistence**: Top N, EPSS, severity filters saved and restored across sessions
- **Removed workflow popups**: Lookup Complete, Ticket Notes saved, Report generation completed – no more OK dialogs
- **Ripple20 / placeholder Fix**: `[None]` and similar placeholders now use Get-RemediationGuidance; product `ripple20-icmp` gets Ripple20 remediation
- **Pie chart stability**: Limited to 25 items (avoids Word COM RPC failures with 150+ items); "Other" slice for remainder
- **Rate limit tuning**: Smaller batch chunks, longer delays, 429 retry with 15s/25s/35s waits
- Bump version to 4.0.9
