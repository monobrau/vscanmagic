# VScanMagic v4.0.5 Release Notes

VScanMagic v4.0.5 - Email template update, RMIT/Client Type dialog fix, probe nmap filtering.

## 📦 Installation

1. Download `VScanMagic.zip` from the release
2. Extract `VScanMagic.exe` along with `VScanMagic-Modules` and `ConnectSecure-API.ps1` to the same folder
3. Run `VScanMagic.exe`

## 📝 Changelog

- **Email template**: Updated default body with new structure (report links, scheduling, RMIT+ note placement)
- **RMIT/Client Type dialog**: Fixed blank screen – DataGridView was never added to the form; now shows client list correctly during bulk download
- **Company Review probe nmap**: Fixed list showing only probe agents (not lightweight); uses `agent_discovery_credentials` endpoint with fallback to `probe_setting` filter
- Bump version to 4.0.5
