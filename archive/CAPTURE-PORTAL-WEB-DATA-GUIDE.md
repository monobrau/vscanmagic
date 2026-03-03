# Capturing ConnectSecure Portal Web Data

Capture API calls from the ConnectSecure portal to discover what hostname, username, vulnerability, and asset data the portal exposes. Use this to align VScanMagic with portal data structures (e.g. username detection for Hostname Review).

## Steps

1. **Log into ConnectSecure** in Chrome or Edge
2. **Open DevTools** (F12) → **Console** tab
3. **Paste** the contents of `Capture-PortalWebData.js` into the console and press Enter
4. You should see: `[WebData Capturer Active] Navigate to: Vulnerability reports...`
5. **Navigate** through the portal and let data load:
   - **Vulnerability reports** – All Vulnerabilities, Pending Remediation EPSS, etc.
   - **Assets** – Asset lists, host details
   - **Agents** – Lightweight agents, probe agents
   - **Company views** – Dashboards that show hostname/username
   - **Report creation** – When generating or viewing reports
   - Expand tables and details that show hostname, IP, username
6. **Export captured data**:
   - Run in console: `copy(JSON.stringify(window.__capturedWebDataCalls, null, 2))`
   - Paste into a file: `portal-web-data-capture.json`

## What to Look For

- **Username fields** – Entries with `hasUsernameField: true` show which endpoints return username data
- **Hostname fields** – `host_name`, `hostname`, `computer_name`, etc.
- **Response structure** – Field names and nesting for vulnerability/asset records
- **Endpoints** – Paths like `/r/company/agents`, `report_queries`, `lightweight_assets`

## Using the Capture

Share `portal-web-data-capture.json` (or relevant snippets) to:
- Identify endpoints that return username data
- Match field names for hostname/username mapping
- Update `ConnectSecure-API.ps1` or `VScanMagic-Data.ps1` to use the same structures
- Add username extraction to vulnerability data if the portal exposes it
