# Developer Notes

Notes for future development and debugging.

---

## PowerShell: Single-Item Array Unwrapping

**Behavior:** PowerShell automatically unwraps single-element arrays when a function returns them. A function that returns `@($singleObject)` may be received by the caller as the single object itself, not as an array.

**Impact:**
- `$result -is [array]` can be `$false` when you expect an array
- `$result.Count` can be `$null` (single objects don't have `.Count` by default)
- `foreach ($x in $result)` still works (iterates once), but code that checks `$result.Count -eq 0` or assumes `$result -is [array]` can fail

**When it happens:**
- Function returns `@($item)` or `$array` where the array has exactly one element
- API/JSON deserialization returns `{ "data": [ {...} ] }` — some responses may deserialize the single-element array as a single object

**Mitigations:**
1. **At the caller:** Wrap the result to force array: `$data = @(Get-SomeFunction)`
2. **In the function:** Return explicitly so it doesn't unwrap:
   - For single object: `return @(,$item)` (comma creates single-element array)
   - For array: `return ,$array` (comma prevents unwrapping)
3. **Normalize in the function:** Always return a proper array:
   ```powershell
   $raw = Invoke-Something
   if ($null -eq $raw) { return @() }
   if ($raw -is [array]) { return @($raw) }
   return @(,$raw)  # single object -> array of 1
   ```

**Relevant code:** `ConnectSecure-API.ps1` — All Company Review data functions now normalize single-object returns:

- `Get-ConnectSecureLightweightAssets` — `@(,$raw)` when single object
- `Get-ConnectSecureCompanyAgents` — `@(,$raw)` when single object
- `Get-ConnectSecureCompanyCredentials` — `@(,$raw)` when single object
- `Get-ConnectSecureAgentCredentialsMapping` — `@(,$raw)` when single object
- `Get-ConnectSecureAgentDiscoveryMapping` — `@(,$raw)` when single object
- `Get-ConnectSecureDiscoverySettings` — `@(,$raw)` in both report and fallback paths
- `Get-ConnectSecureExternalScanDiscoverySettings` — `@(,$raw)` when single object
- `Get-ConnectSecureAssetFirewallPolicy` — `@(,$raw)` when single object
- `jobsView` in `Get-ConnectSecureCompanyReviewData` — `@(,$jobsData)` when single job
