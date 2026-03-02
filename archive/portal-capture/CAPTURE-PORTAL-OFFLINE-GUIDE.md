# Capturing ConnectSecure Portal Offline Agent/Asset API Calls

Capture the API calls the ConnectSecure web UI makes when viewing **offline agents** or **offline assets** data. Use this to align VScanMagic's Company Review offline counts (7d / 14d / 30+d) with what the portal uses.

## Steps

1. **Log into ConnectSecure** in Chrome or Edge
2. **Open DevTools** (F12) → go to the **Console** tab
3. **Paste the contents** of `Capture-PortalOffline.js` into the console and press Enter
4. You should see: `[Offline Capturer Active] Navigate to views showing offline agents/assets...`
5. **Navigate to offline-related views** in the portal:
   - Select a **specific company** (not Global) from the company selector
   - Go to **Agents** or **Assets** or **Company overview** – anywhere that shows offline counts or offline agent lists
   - Look for dashboards/cards showing "Offline", "Offline agents", "Offline assets", or similar
   - Expand agent lists, asset lists, or company summary views that might show offline status
   - Let each view load fully
6. **Export captured data**:
   - Run in console: `copy(JSON.stringify(window.__capturedOfflineCalls, null, 2))`
   - Paste into a file: `portal-offline-capture.json`

## What to Look For

- **Endpoints** – Does the portal use `company/agents`, `lightweight_assets`, `offline_assets`, or another endpoint for offline counts?
- **Query params** – `company_id`, `condition`, `limit`, filters for offline status?
- **Response structure** – Fields like `last_ping_time`, `lastPingTime`, `offline_assets_count`, `offline_assets`, `is_deprecated`
- **Time buckets** – Does the portal return 7d/14d/30d counts or a single offline count?

## Alternative: DevTools Network Tab

1. **Open DevTools** (F12) → **Network** tab
2. Enable **Preserve log**
3. Filter by **Fetch/XHR** (or search for `offline`, `agents`, `lightweight`, `last_ping`)
4. Select a company and navigate to Agents, Assets, or Company view
5. Inspect requests – look for URLs that might return offline data
6. Right-click a request → **Copy** → **Copy as cURL** or inspect Response

## Using the Capture

Share `portal-offline-capture.json` (or paste relevant snippets) so we can:
- Identify the correct endpoint(s) for offline agent counts
- Match query parameters and response structure
- Update `ConnectSecure-API.ps1` (Get-ConnectSecureCompanyReviewData) to use the same logic as the portal
