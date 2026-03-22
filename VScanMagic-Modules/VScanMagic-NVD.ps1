# VScanMagic-NVD.ps1 - NVD API 2.0 helpers for remediation text (dot-sourced)
# Requires: PowerShell 5.1+

$script:NvdCache = @{}  # CVE ID (uppercase) -> remediation string
$script:NvdLastRequestUtc = $null

function Wait-NvdRateLimit {
    param([bool]$HasApiKey)
    $minMs = if ($HasApiKey) { 600 } else { 6000 }
    if ($null -eq $script:NvdLastRequestUtc) { return }
    $elapsed = ([datetime]::UtcNow - $script:NvdLastRequestUtc).TotalMilliseconds
    if ($elapsed -lt $minMs) {
        $wait = [int]([Math]::Ceiling($minMs - $elapsed))
        Start-Sleep -Milliseconds $wait
    }
}

function Split-NvdCveIds {
    param([string]$Text)
    if ([string]::IsNullOrWhiteSpace($Text)) { return @() }
    $matches = [regex]::Matches($Text, 'CVE-\d{4}-\d+', [System.Text.RegularExpressions.RegexOptions]::IgnoreCase)
    $ids = [System.Collections.Generic.HashSet[string]]::new([StringComparer]::OrdinalIgnoreCase)
    foreach ($m in $matches) { $null = $ids.Add($m.Value.ToUpperInvariant()) }
    return @($ids)
}

function Build-NvdRemediationText {
    param([object]$CveItem)
    if ($null -eq $CveItem) { return '' }
    $cve = $CveItem.cve
    if ($null -eq $cve) { return '' }

    $desc = ''
    if ($cve.descriptions) {
        $en = @($cve.descriptions | Where-Object { $_.lang -eq 'en' })
        if ($en.Count -gt 0) { $desc = [string]$en[0].value }
    }

    $refUrls = [System.Collections.Generic.List[string]]::new()
    $seenUrls = [System.Collections.Generic.HashSet[string]]::new([StringComparer]::OrdinalIgnoreCase)
    if ($cve.references) {
        foreach ($r in $cve.references) {
            $url = [string]$r.url
            if ([string]::IsNullOrWhiteSpace($url)) { continue }
            $src = [string]$r.source
            if ($url -match 'patch|advisory|security|bulletin|vendor|mitre|nist\.gov' -or
                $src -match 'patch|advisory|security|vendor') {
                if ($seenUrls.Add($url)) { $refUrls.Add($url) | Out-Null }
            }
        }
        if ($refUrls.Count -eq 0) {
            foreach ($r in $cve.references) {
                $url = [string]$r.url
                if (-not [string]::IsNullOrWhiteSpace($url) -and $seenUrls.Add($url)) { $refUrls.Add($url) | Out-Null }
            }
        }
    }

    $parts = @()
    if ($refUrls.Count -gt 0) {
        $take = [Math]::Min(3, $refUrls.Count)
        for ($i = 0; $i -lt $take; $i++) { $parts += $refUrls[$i] }
    }
    if (-not [string]::IsNullOrWhiteSpace($desc)) {
        $maxDesc = 400
        if ($desc.Length -gt $maxDesc) { $desc = $desc.Substring(0, $maxDesc).TrimEnd() + '…' }
        if ($parts.Count -gt 0) { $parts += $desc } else { $parts = @($desc) }
    }

    $out = ($parts -join ' | ')
    if ($out.Length -gt 500) { $out = $out.Substring(0, 500).TrimEnd() + '…' }
    return $out
}

<#
.SYNOPSIS
    Fetches a short remediation-oriented summary from the NVD CVE 2.0 API.
.DESCRIPTION
    Uses https://services.nvd.nist.gov/rest/json/cves/2.0?cveId=...
    Rate limits: ~5 req/30s without API key, ~50 req/30s with key. Results are cached per process.
#>
function Get-NVDRemediationForCve {
    param(
        [Parameter(Mandatory = $true)]
        [string]$CveId,
        [string]$ApiKey = ''
    )

    $normalized = ($CveId -replace '^\s+|\s+$', '').ToUpperInvariant()
    if ($normalized -notmatch '^CVE-\d{4}-\d+$') {
        Write-Warning "Get-NVDRemediationForCve: invalid CVE id '$CveId'"
        return ''
    }

    if ($script:NvdCache.ContainsKey($normalized)) {
        return [string]$script:NvdCache[$normalized]
    }

    $hasKey = -not [string]::IsNullOrWhiteSpace($ApiKey)
    Wait-NvdRateLimit -HasApiKey:$hasKey

    $base = 'https://services.nvd.nist.gov/rest/json/cves/2.0'
    $uri = "$base`?cveId=$([uri]::EscapeDataString($normalized))"
    $headers = @{
        'Accept' = 'application/json'
    }
    if ($hasKey) {
        $headers['apiKey'] = $ApiKey.Trim()
    }

    try {
        $params = @{
            Uri             = $uri
            Method          = 'Get'
            Headers         = $headers
            UseBasicParsing = $true
            TimeoutSec      = 60
            ErrorAction     = 'Stop'
        }
        $resp = Invoke-RestMethod @params
        $script:NvdLastRequestUtc = [datetime]::UtcNow

        $text = ''
        if ($resp.vulnerabilities -and $resp.vulnerabilities.Count -gt 0) {
            $text = Build-NvdRemediationText -CveItem $resp.vulnerabilities[0]
        }

        $script:NvdCache[$normalized] = $text
        return $text
    } catch {
        Write-Warning "Get-NVDRemediationForCve: NVD request failed for $normalized : $($_.Exception.Message)"
        $script:NvdCache[$normalized] = ''
        $script:NvdLastRequestUtc = [datetime]::UtcNow
        return ''
    }
}

<#
.SYNOPSIS
    Returns combined remediation text for multiple CVE IDs (deduped, cached).
#>
function Get-NVDRemediationForCveList {
    param(
        [string]$CveListText,
        [string]$ApiKey = ''
    )

    $ids = Split-NvdCveIds -Text $CveListText
    if ($ids.Count -eq 0) { return '' }

    $chunks = @()
    foreach ($id in $ids) {
        Write-Host "  Fetching NVD for $id..." -ForegroundColor DarkGray
        $t = Get-NVDRemediationForCve -CveId $id -ApiKey $ApiKey
        if (-not [string]::IsNullOrWhiteSpace($t)) { $chunks += $t }
    }

    $combined = ($chunks -join ' | ')
    if ($combined.Length -gt 500) { $combined = $combined.Substring(0, 500).TrimEnd() + '…' }
    return $combined
}
