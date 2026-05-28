# ConnectSecure Patch Management API

Reference for VScanMagic patching and job status. Discovered via CS portal Network capture (Stream Jog / company 17624, pod104, May 2026).

**Portal UI:** Patch Management → Application Patching / Patch Jobs  
**API host:** Often `https://api.myconnectsecure.com` (may differ from portal URL).

---

## Trigger patch (write)

**Endpoint:** `POST /w/company/patch_now`

**Headers:** `Authorization: Bearer …`, `X-USER-ID: …`

**Example — immediate application patch (matches CS portal):**

```json
{
  "companies": [17624],
  "patch_when": "now",
  "type": "application",
  "included_application": ["Mozilla Firefox"],
  "agents_id": [402780, 5996928],
  "assets": [24079295, 293858126],
  "from_versions": {
    "24079295": "150.0.1",
    "293858126": "150.0.3"
  },
  "excludecompany": [],
  "exclude_tags": [],
  "execluded_application": [],
  "include_tags": []
}
```

**Success response:**

```json
{ "status": true, "message": "Message sent for patch update" }
```

**Notes:**

- `from_versions` keys are **asset IDs** (strings), not agent IDs.
- Portal sends empty exclude/tag arrays; VScanMagic omits them when empty (harmless).
- Portal may pre-flight with `POST /r/asset/assets_status` (`condition: id IN (…)`); not required for `patch_now` to return success.

**VScanMagic:** `ConnectSecurePatchService.BuildPatchPayload` / `PatchApplicationsNowAsync`.

---

## Patch job list (read) — use this, not `jobs_view`

**Endpoint:** `GET /r/report_queries/patch_jobview`

**Example query:**

```
condition=company_id=17624&skip=0&limit=100&order_by=created desc
```

**Do not use** `/r/company/jobs_view` for patch job status — that endpoint lists scans, reports, and suppressions, not patch jobs.

**VScanMagic:** `ConnectSecurePatchService.FetchConnectSecurePatchJobsAsync` → `PatchJobCorrelationHelper.ParsePatchJobViewRow`.

**UI query:** `PatchJobListQuery` supports `DaysBack` (default 7), `Page` / `PageSize` (default 15), and `LocalOnly` (VScanMagic-triggered patches only).

### Response row shape (sample)

| Field | Example | Purpose |
|-------|---------|---------|
| `job_id` / `patch_id` | `ac71ebf5-fcbe-4b57-9995-ff02ae21435a` | CS job UUID; link local activity |
| `job_status` / `status` | `Initiated`, `Success`, `Failed`, `Pending` | Overall job state |
| `product_name` | `Mozilla Firefox` | Software patched |
| `type` | `Application`, `OS`, `linux patch` | Patch type |
| `created` / `updated` | ISO datetime | Sort / correlate by time |
| `msg` | `[0, 0, 2]` | `[success, failed, pending]` counts |
| `patch_job_details` | object keyed by asset ID | Per-host status |

### `patch_job_details` (per asset)

```json
{
  "24079295": {
    "from_version": "150.0.1",
    "to_version": "151.0.1",
    "host_name": "Roswell.dorks.lan",
    "status": "Pending",
    "status_msg": "Initiated"
  }
}
```

---

## Application patching catalog (read)

Product list (Software Name, Fix Version, Assets, Action) uses report_queries remediation/vulnerability endpoints — not `patch_jobview`.

Example row object when drilling into a product (Action click):

- `software_name`, `fix`, `company_ids`, `asset_ids`, `host_names`, `software_type`

**VScanMagic:** company-scoped `get_remediation?condition=company_id=X` + product-scoped asset details.

---

## Related endpoints

| Endpoint | Use |
|----------|-----|
| `POST /w/company/reset_agents` | Inventory refresh (`message: update_agent`); does **not** install patches |
| `GET /r/company/get_patch_settings?company_id=X` | Pre-flight: patch management enabled for company |
| `GET /r/company/scheduler` | Scheduled patch definitions (not the Patch Jobs history table) |
| `GET /r/company/jobs_view` | Scan/report jobs only — **not** patch jobs |

---

## Debugging in browser

Capture scripts (gitignored): `archive/CsPortal-CapturePatchRequests.js`, `archive/CsPortal-CapturePatchJobsLoad.js`

Probe script: `archive/Test-PatchJobsEndpoints.ps1 -UseSavedCredentials -CompanyId 17624`

Compare portal vs VScanMagic payload: `archive/Compare-CsPortalPatchRequest.ps1`
