# Capturing ConnectSecure Portal Probe / Nmap Interface API Calls

Capture how the ConnectSecure web portal fetches and filters **probe agents** vs lightweight agents, and how it displays **nmap interface** data.

**Goal:** VScanMagic Company Review shows "Probe Agent(s) Nmap Interface" but lists lightweight agents instead of only probes. The portal knows the difference—we need to see what API calls and filters it uses.

## Steps

1. **Log into ConnectSecure** (Chrome or Edge)
2. **Open DevTools** (F12) → **Console** tab
3. **Clear previous capture** (optional): `window.__capturedProbeNmapCalls = []`
4. **Paste** the contents of `Capture-PortalProbeNmap.js` into the console and press Enter
5. You should see: `[Probe/Nmap Capturer Active] Navigate to Probes or Agents view...`
6. **Navigate to the probe/agent view:**
   - Select a **specific company** (the one where VScanMagic shows 1 probe but 38 in the list)
   - Go to **Probes** (if the portal has a Probes menu)
   - Or **Agents** → look for a filter/tab for "Probe Agents" vs "Lightweight Agents"
   - Or **Company Settings** → **Discovery** / **Credentials** / **Scan Agents**
   - Let the page load fully; expand any agent details that show nmap interface
7. **Export captured data:**
   - For probe-related calls only:
     ```
     copy(JSON.stringify(window.__capturedProbeNmapCalls, null, 2))
     ```
   - For all API calls (broader capture):
     ```
     copy(JSON.stringify(window.__capturedAllApiCalls, null, 2))
     ```
8. **Save** the pasted JSON to `portal-probe-nmap-capture.json`

## What to Look For

### 1. Agents endpoint – how does the portal filter?

- **Path:** `/r/company/agents` or similar
- **queryParams:** Does it send `condition=agent_type='probe'` or `agent_type=probe`?
- **firstRecordSample:** What does `agent_type` / `agentType` look like for probes vs lightweight?
  - e.g. `"probe"`, `"scan_probe"`, `"lightweight"`, `null`?

### 2. Nmap interface and probe_setting

- For each agent in the response, does `firstRecordSample` include:
  - `nmap_interface` / `nmapInterface`
  - `probe_setting` / `probeSetting` (object with listen_port, etc.)

### 3. Separate probe-only endpoint?

- Does the portal call a **different endpoint** for probes (e.g. `/r/company/probes` or `/r/report_queries/probe_agents`) instead of filtering `/r/company/agents`?

### 4. Credential / discovery mappings

- `agent_credentials_mapping` and `agent_discoverysettings_mapping`
- Do these use `agent_id`? Are they already filtered to probe agents only?

## Share the Capture

After capturing, paste or share:

1. The `/r/company/agents` (or equivalent) call—full `queryParams` and `firstRecordSample` of the first 1–2 agents
2. Any probe-specific endpoints
3. `agent_type` values seen in the responses

This will show how to align VScanMagic’s Probe Nmap list with the portal’s probe-only view.
