# Capturing ConnectSecure Portal Firewall API Calls

Capture the exact API calls the ConnectSecure web UI makes when viewing **firewall** data (Fortigate, Sonicwall counts, device types, etc.). Use this to align VScanMagic's Company Review firewall display with what the portal uses.

## Steps

1. **Log into ConnectSecure** in Chrome or Edge
2. **Open DevTools** (F12) → go to the **Console** tab
3. **Paste the contents** of `Capture-PortalFirewall.js` into the console and press Enter
4. You should see: `[Firewall Capturer Active] Navigate to the Firewall section...`
5. **Navigate to firewall data** in the portal:
   - Select a **specific company** (not Global) from the company selector
   - Go to the **Firewall** section (or equivalent – may be under Assets, Security, or Company view)
   - Expand any firewall device lists, interfaces, or dashboards
   - Let each view load fully
6. **Export captured data**:
   - Run in console: `copy(JSON.stringify(window.__capturedFirewallCalls, null, 2))`
   - Paste into a file: `portal-firewall-capture.json`

## What to Look For

- **Endpoints** – Does the portal use `firewall_interfaces`, `firewall_groups`, `asset_firewall_policy`, or something else for device counts?
- **Query params** – `company_id`, `condition`, or other filters?
- **Response structure** – Fields like `manufacturer`, `asset_type`, `policy_type`, `is_firewall` that indicate device vendor (Fortigate, Sonicwall) vs Windows profiles (PublicProfile, DomainProfile, StandardProfile)

## Alternative: DevTools Network Tab

1. **Open DevTools** (F12) → **Network** tab
2. Enable **Preserve log**
3. Filter by **Fetch/XHR** (or search for `firewall`)
4. Select a company and navigate to the Firewall section
5. Inspect requests – look for URLs containing `firewall`, `asset_firewall`
6. Right-click a request → **Copy** → **Copy as cURL** or inspect Response to see structure

## Using the Capture

Share `portal-firewall-capture.json` (or paste relevant snippets) so we can:
- Identify the correct endpoint(s) for firewall device count and vendor (Fortigate, Sonicwall, etc.)
- Match query parameters the portal uses for company filtering
- Update `ConnectSecure-API.ps1` (Get-ConnectSecureCompanyReviewData) and the Company Review dialog to use the same logic

**Portal finding (2026-03):** The portal uses `/r/report_queries/firewall_asset_view` with `condition=company_id=X`, `limit`, `skip`, `order_by=host_name asc`. Response includes `manufacturer` (e.g. Cisco Meraki), `is_firewall`, `name`, `company_id`. VScanMagic now uses this endpoint for firewall count and vendor types.
