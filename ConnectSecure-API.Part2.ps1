# ConnectSecure-API.Part2.ps1 — Dot-sourced by ConnectSecure-API.ps1 (do not run directly)
# Company review data, asset resolution, Excel export helpers, ConvertTo-ConnectSecureToVulnData.
function Invoke-ConnectSecureCompanyReviewRequest {
    param([string]$Endpoint, [hashtable]$QueryParams = @{})
    try {
        $response = Invoke-ConnectSecureRequest -Endpoint $Endpoint -QueryParameters $QueryParams
        switch ($response -is [array]) { $true { return @($response) } }
        $data = $response.data
        # API returns data as string "Failed to retrieve data from the server...." when condition param causes error
        switch ($data -is [string] -and $data -match 'Failed|error|Error') { $true { return @() } }
        switch ($data -is [array]) { $true { return @($data) } }
        switch ($data -and $data -isnot [string]) { $true { return @(,$data) } }
        return @()
    } catch {
        Write-CSApiLog "Company review endpoint $Endpoint failed: $($_.Exception.Message)" -Level Warning
        return $null
    }
}

function Get-ConnectSecureCompanyStats {
    param([int]$CompanyId = 0)
    $qp = if ($CompanyId -gt 0) { @{ condition = "company_id=$CompanyId"; limit = 100; skip = 0 } } else { @{ limit = 1000; skip = 0 } }
    $data = Invoke-ConnectSecureCompanyReviewRequest -Endpoint $script:CompanyReviewEndpoints.CompanyStats -QueryParams $qp
    switch ($null -eq $data) { $true { return @{} } }
    $rows = switch ($data -is [array] -and $data.Count -gt 0) { $true { $data } default { @($data) } }
    switch (-not $rows -or $rows.Count -eq 0) { $true { return @{} } }
    $sorted = $rows | Sort-Object -Property @{ Expression = { $_.date }; Descending = $true }, @{ Expression = { $_.ad_last_scan_time }; Descending = $true }
    return $sorted[0]
}

function Get-ConnectSecureLightweightAssets {
    param([int]$CompanyId)
    switch ($CompanyId -le 0) { $true { return @() } }
    $qp = @{ company_id = $CompanyId; limit = 5000; skip = 0 }
    $raw = Invoke-ConnectSecureCompanyReviewRequest -Endpoint $script:CompanyReviewEndpoints.LightweightAssets -QueryParams $qp
    if ($null -eq $raw) { return @() }
    if ($raw -is [array]) { return @($raw) }
    return @(,$raw)
}

function Get-ConnectSecureCompanyAgents {
    param([int]$CompanyId)
    switch ($CompanyId -le 0) { $true { return @() } }
    $qp = @{ condition = "company_id=$CompanyId"; limit = 5000; skip = 0 }
    $raw = Invoke-ConnectSecureCompanyReviewRequest -Endpoint $script:CompanyReviewEndpoints.Agents -QueryParams $qp
    if ($null -eq $raw) { return @() }
    if ($raw -is [array]) { return @($raw) }
    return @(,$raw)
}

function Get-ConnectSecureUsernamesByHostname {
    <#
    .SYNOPSIS
    Returns hostname -> logged_in_user map from ConnectSecure asset_view. Use for Hostname Review username lookup.
    #>
    param(
        [int]$CompanyId,
        [string[]]$Hostnames
    )
    if ($CompanyId -le 0 -or -not $Hostnames -or $Hostnames.Count -eq 0) { return @{} }
    $result = @{}
    foreach ($h in $Hostnames) { if (-not [string]::IsNullOrWhiteSpace($h)) { $result[$h.Trim()] = "" } }
    if ($result.Count -eq 0) { return @{} }
    try {
        $assetToUser = @{}
        $skip = 0
        $limit = 500
        do {
            $qp = @{ condition = "company_id=$CompanyId"; limit = $limit; skip = $skip; order_by = "host_name asc" }
            $data = Invoke-ConnectSecureCompanyReviewRequest -Endpoint $script:CompanyReviewEndpoints.AssetView -QueryParams $qp
            if (-not $data -or ($data -is [array] -and $data.Count -eq 0)) { break }
            $rows = if ($data -is [array]) { $data } else { @(,$data) }
            foreach ($r in $rows) {
                $hn = $r.host_name; if (-not $hn) { $hn = $r.hostname }; if (-not $hn) { $hn = $r.name }; if (-not $hn) { $hn = $r.'Host Name' }
                $user = $r.logged_in_user; if (-not $user) { $user = $r.logged_in_user_name }; if (-not $user) { $user = $r.'Logged In User' }
                if ([string]::IsNullOrWhiteSpace($user)) { continue }
                $userStr = [string]$user.Trim()
                $hnStr = [string]$hn
                if (-not $hnStr) { continue }
                $hnNorm = $hnStr.Trim().ToLowerInvariant()
                $shortName = if ($hnNorm -match '^([^.]+)\.') { $Matches[1] } else { $hnNorm }
                $assetToUser[$hnNorm] = $userStr
                $assetToUser[$shortName] = $userStr
                $nameVal = $r.name; if ($nameVal) { $assetToUser[[string]$nameVal.Trim().ToLowerInvariant()] = $userStr }
            }
            $skip += $rows.Count
            if ($rows.Count -lt $limit) { break }
        } while ($true)
        foreach ($key in @($result.Keys)) {
            $keyNorm = $key.Trim().ToLowerInvariant()
            $shortKey = if ($keyNorm -match '^([^.]+)\.') { $Matches[1] } else { $keyNorm }
            if ($assetToUser.ContainsKey($keyNorm)) { $result[$key] = $assetToUser[$keyNorm] }
            elseif ($assetToUser.ContainsKey($shortKey)) { $result[$key] = $assetToUser[$shortKey] }
        }
    } catch {
        Write-CSApiLog "Get-ConnectSecureUsernamesByHostname failed: $($_.Exception.Message)" -Level Warning
    }
    return $result
}

function Get-ConnectSecureLastPingByHostname {
    <#
    .SYNOPSIS
    Returns hostname/IP -> last_ping_time map from ConnectSecure agents. Use for Top N report and ticket instructions.
    Same last_ping_time used for company review "agent last online" display.
    #>
    param(
        [int]$CompanyId,
        [string[]]$Hostnames,
        [string[]]$IPs = @()
    )
    if ($CompanyId -le 0) { return @{} }
    $result = @{}
    foreach ($h in $Hostnames) { if (-not [string]::IsNullOrWhiteSpace($h)) { $result[$h.Trim()] = "" } }
    foreach ($ip in $IPs) { if (-not [string]::IsNullOrWhiteSpace($ip)) { $result[$ip.Trim()] = "" } }
    if ($result.Count -eq 0) { return @{} }
    try {
        $agents = Get-ConnectSecureCompanyAgents -CompanyId $CompanyId
        if (-not $agents -or $agents.Count -eq 0) { return $result }
        $hostToPing = @{}
        $ipToPing = @{}
        foreach ($a in $agents) {
            $lp = $a.last_ping_time; switch (-not $lp) { $true { $lp = $a.lastPingTime } }
            switch (-not $lp) { $true { $lp = $a.'Last Ping Time' } }
            switch (-not $lp) { $true { continue } }
            try {
                $dt = [DateTime]::Parse($lp)
                $formatted = $dt.ToLocalTime().ToString('yyyy-MM-dd HH:mm')
            } catch { continue }
            $hn = $a.host_name; switch (-not $hn) { $true { $hn = $a.'Host Name' } }
            switch (-not $hn) { $true { $hn = $a.hostname } }; switch (-not $hn) { $true { $hn = $a.name } }
            $ip = $a.ip; switch (-not $ip) { $true { $ip = $a.IP } }
            if (-not [string]::IsNullOrWhiteSpace($hn)) {
                $hnNorm = [string]$hn.Trim().ToLowerInvariant()
                $shortName = if ($hnNorm -match '^([^.]+)\.') { $Matches[1] } else { $hnNorm }
                if (-not $hostToPing.ContainsKey($hnNorm) -or [string]::Compare($formatted, $hostToPing[$hnNorm]) -gt 0) { $hostToPing[$hnNorm] = $formatted }
                if (-not $hostToPing.ContainsKey($shortName) -or [string]::Compare($formatted, $hostToPing[$shortName]) -gt 0) { $hostToPing[$shortName] = $formatted }
            }
            if (-not [string]::IsNullOrWhiteSpace($ip)) {
                $ipStr = [string]$ip.Trim()
                if (-not $ipToPing.ContainsKey($ipStr) -or [string]::Compare($formatted, $ipToPing[$ipStr]) -gt 0) { $ipToPing[$ipStr] = $formatted }
            }
        }
        foreach ($key in @($result.Keys)) {
            $keyNorm = $key.Trim().ToLowerInvariant()
            $shortKey = if ($keyNorm -match '^([^.]+)\.') { $Matches[1] } else { $keyNorm }
            $pingStr = ""
            if ($hostToPing.ContainsKey($keyNorm)) { $pingStr = $hostToPing[$keyNorm] }
            elseif ($hostToPing.ContainsKey($shortKey)) { $pingStr = $hostToPing[$shortKey] }
            elseif ($ipToPing.ContainsKey($key)) { $pingStr = $ipToPing[$key] }
            elseif ($ipToPing.ContainsKey($key.Trim())) { $pingStr = $ipToPing[$key.Trim()] }
            $result[$key] = $pingStr
        }
    } catch {
        Write-CSApiLog "Get-ConnectSecureLastPingByHostname failed: $($_.Exception.Message)" -Level Warning
    }
    return $result
}

function Get-ConnectSecureAgentCredentialsMapping {
    param([int]$CompanyId)
    switch ($CompanyId -le 0) { $true { return @() } }
    $qp = @{ condition = "company_id=$CompanyId"; limit = 5000; skip = 0 }
    $raw = Invoke-ConnectSecureCompanyReviewRequest -Endpoint $script:CompanyReviewEndpoints.AgentCredentialsMapping -QueryParams $qp
    if ($null -eq $raw) { return @() }
    if ($raw -is [array]) { return @($raw) }
    return @(,$raw)
}

function Get-ConnectSecureDiscoverySettings {
    param([int]$CompanyId)
    switch ($CompanyId -le 0) { $true { return @() } }
    $qp = @{ condition = "company_id=$CompanyId"; limit = 500; skip = 0; order_by = 'updated desc' }
    try {
        $raw = Invoke-ConnectSecureCompanyReviewRequest -Endpoint $script:CompanyReviewEndpoints.DiscoverySettingsReport -QueryParams $qp
        if ($null -ne $raw) {
            $data = if ($raw -is [array]) { @($raw) } else { @(,$raw) }
            if ($data.Count -gt 0) { return $data }
        }
    } catch { }
    $qp2 = @{ limit = 2000; skip = 0 }
    $raw = Invoke-ConnectSecureCompanyReviewRequest -Endpoint $script:CompanyReviewEndpoints.DiscoverySettings -QueryParams $qp2
    if ($null -eq $raw) { return @() }
    $data = if ($raw -is [array]) { @($raw) } else { @(,$raw) }
    return @($data | Where-Object { ($_.company_id -eq $CompanyId) -or ($_.companyId -eq $CompanyId) })
}

function Get-ConnectSecureExternalScanDiscoverySettings {
    <#
    .SYNOPSIS
    Returns external scan discovery_settings for a company. Server-side filtered via condition.
    Used for Company Review section 3 (External assets).
    #>
    param([int]$CompanyId)
    switch ($CompanyId -le 0) { $true { return @() } }
    $qp = @{ condition = "company_id=$CompanyId and discovery_settings_type='externalscan'"; limit = 500; skip = 0; order_by = 'updated desc' }
    try {
        $raw = Invoke-ConnectSecureCompanyReviewRequest -Endpoint $script:CompanyReviewEndpoints.DiscoverySettingsReport -QueryParams $qp
        if ($null -eq $raw) { return @() }
        if ($raw -is [array]) { return @($raw) }
        return @(,$raw)
    } catch {
        Write-CSApiLog "Get-ConnectSecureExternalScanDiscoverySettings catch: $($_.Exception.Message)" -Level Warning
    }
    return @()
}

function Get-ConnectSecureAgentDiscoveryMapping {
    param([int]$CompanyId)
    switch ($CompanyId -le 0) { $true { return @() } }
    $qp = @{ condition = "company_id=$CompanyId"; limit = 5000; skip = 0 }
    $raw = Invoke-ConnectSecureCompanyReviewRequest -Endpoint $script:CompanyReviewEndpoints.AgentDiscoveryMapping -QueryParams $qp
    if ($null -eq $raw) { return @() }
    if ($raw -is [array]) { return @($raw) }
    return @(,$raw)
}

function Get-ConnectSecureProbeAgents {
    <#
    .SYNOPSIS
    Returns probe agents only (agent_type='PROBE'). Portal uses /r/company/agent_discovery_credentials
    with server-side filter - aligns with ConnectSecure portal probe view.
    #>
    param([int]$CompanyId)
    switch ($CompanyId -le 0) { $true { return @() } }
    $cond = "company_id=$CompanyId and agent_type='PROBE' and is_deprecated=FALSE and is_retired=FALSE"
    $qp = @{ condition = $cond; skip = 0; limit = 100; order_by = 'host_name asc' }
    try {
        $raw = Invoke-ConnectSecureCompanyReviewRequest -Endpoint $script:CompanyReviewEndpoints.AgentDiscoveryCredentials -QueryParams $qp
        if ($null -eq $raw) { return @() }
        if ($raw -is [array]) { return @($raw) }
        return @(,$raw)
    } catch {
        Write-CSApiLog "Get-ConnectSecureProbeAgents failed: $($_.Exception.Message)" -Level Warning
        return @()
    }
}

function Get-ConnectSecureAssetFirewallPolicy {
    param([int]$CompanyId)
    switch ($CompanyId -le 0) { $true { return @() } }
    $qp = @{ condition = "company_id=$CompanyId"; limit = 2000; skip = 0 }
    $raw = Invoke-ConnectSecureCompanyReviewRequest -Endpoint $script:CompanyReviewEndpoints.AssetFirewallPolicy -QueryParams $qp
    if ($null -eq $raw) { return @() }
    if ($raw -is [array]) { return @($raw) }
    return @(,$raw)
}

function Get-ConnectSecureCompanyReviewData {
    <#
    .SYNONOPSIS
    Fetches data for Company Review dialog: agents, probes, assets, company stats.
    Uses company_stats and assets; extends when swagger endpoints are available.
    #>
    param([int]$CompanyId, [string]$CompanyName = '')

    $result = @{
        CompanyId = $CompanyId
        CompanyName = $CompanyName
        AgentCount = 0
        Agents = @()
        Probes = @()
        ProbesWithCredentials = 0
        ProbesWithNetworks = 0
        ProbesWithBoth = 0
        ProbesSubnets = [System.Collections.ArrayList]::new()
        SubnetIssues = [System.Collections.ArrayList]::new()
        ScanTargets = [System.Collections.ArrayList]::new()
        ExternalAssets = [System.Collections.ArrayList]::new()
        OldestOfflineDays = $null
        AgentsOffline7PlusDays = 0
        AgentsOffline14PlusDays = 0
        AgentsOffline30PlusDays = 0
        AgentsOffline30PlusNames = [System.Collections.ArrayList]::new()
        FirewallActive = $false
        FirewallCount = 0
        FirewallType = ''
        LastInternalScan = $null
        LastExternalScan = $null
        QuickWins = [System.Collections.ArrayList]::new()
        ProbeAgentsNmapInfo = [System.Collections.ArrayList]::new()
    }

    switch ($CompanyId -le 0) { $true { return $result } }

    $cidStr = [string]$CompanyId

    # 1. Lightweight agent count - /r/report_queries/lightweight_assets
    $lwAssets = Get-ConnectSecureLightweightAssets -CompanyId $CompanyId
    switch ($lwAssets -is [array]) { $true { $result.AgentCount = $lwAssets.Count } }

    # 2. Agents (for probe check, oldest offline)
    $agents = Get-ConnectSecureCompanyAgents -CompanyId $CompanyId
    switch (-not $agents) { $true { $agents = @() } }
    $result.Agents = @($agents)

    $needFallback = $result.AgentCount -eq 0 -and $agents.Count -gt 0
    switch ($needFallback) {
        $true {
            $lwAgents = $agents | Where-Object { $_.agent_type -and [string]$_.agent_type -match 'lightweight' }
            $ac = ($lwAgents | Measure-Object).Count
            $result.AgentCount = switch ($ac -gt 0) { $true { $ac } default { $agents.Count } }
        }
    }

    # 3. Probes with credentials and networks
    $credMappings = Get-ConnectSecureAgentCredentialsMapping -CompanyId $CompanyId
    $discMappings = Get-ConnectSecureAgentDiscoveryMapping -CompanyId $CompanyId
    switch (-not $credMappings) { $true { $credMappings = @() } }
    switch (-not $discMappings) { $true { $discMappings = @() } }
    $agentIdsWithCreds = @{}
    foreach ($m in $credMappings) { $aid = $m.agent_id; switch ($null -ne $aid) { $true { $agentIdsWithCreds[[int]$aid] = $true } } }
    $agentIdsWithNetworks = @{}
    foreach ($m in $discMappings) { $aid = $m.agent_id; switch ($null -ne $aid) { $true { $agentIdsWithNetworks[[int]$aid] = $true } } }
    $probesWithBoth = 0
    foreach ($a in $agents) {
        $aid = $a.id
        switch ($null -eq $aid) { $true { continue } }
        $aidInt = [int]$aid
        switch ($agentIdsWithCreds[$aidInt] -and $agentIdsWithNetworks[$aidInt]) { $true { $probesWithBoth++ } }
    }
    $result.ProbesWithCredentials = ($agentIdsWithCreds.Keys | Measure-Object).Count
    $result.ProbesWithNetworks = ($agentIdsWithNetworks.Keys | Measure-Object).Count
    $result.ProbesWithBoth = $probesWithBoth

    # 3b. Probe agents nmap interface - use portal endpoint agent_discovery_credentials with agent_type='PROBE'
    $probeAgents = Get-ConnectSecureProbeAgents -CompanyId $CompanyId
    if (-not $probeAgents -or $probeAgents.Count -eq 0) {
        # Fallback: filter agents by probe_setting (only probe agents have it; lightweight do not)
        $probeAgents = @($agents | Where-Object {
            $ps = $_.probe_setting; if (-not $ps) { $ps = $_.probeSetting }
            $ps -and $ps -is [System.Management.Automation.PSObject]
        })
    }
    foreach ($a in $probeAgents) {
        $hostName = $a.host_name; switch (-not $hostName) { $true { $hostName = $a.'Host Name' } }
        switch (-not $hostName) { $true { $hostName = $a.agent_name } }
        switch (-not $hostName) { $true { $hostName = $a.hostname } }
        switch (-not $hostName) { $true { $hostName = $a.name } }
        switch (-not $hostName) { $true { $hostName = "(unnamed)" } }
        $ip = $a.ip; switch (-not $ip) { $true { $ip = $a.agent_ip } }
        $nmapIf = $a.nmap_interface; switch (-not $nmapIf) { $true { $nmapIf = $a.nmapInterface } }
        $port = $null
        $ps = $a.probe_setting; switch (-not $ps) { $true { $ps = $a.probeSetting } }
        if ($ps) {
            $port = $ps.listen_port; switch (-not $port) { $true { $port = $ps.port } }
            switch (-not $port) { $true { $port = $ps.nmap_port } }
        }
        [void]$result.ProbeAgentsNmapInfo.Add(@{
            HostName = [string]$hostName
            IP = if ($ip) { [string]$ip } else { "(none)" }
            NmapInterface = if ($nmapIf) { [string]$nmapIf } else { "(not set)" }
            Port = if ($null -ne $port -and [string]$port -ne '') { [string]$port } else { $null }
        })
    }

    # 4. External subnet config - discovery_settings with address (CIDR) and target_ip (external IPs to scan)
    $discoverySettings = Get-ConnectSecureDiscoverySettings -CompanyId $CompanyId
    $dsIdToAddr = @{}
    switch ($discoverySettings) { { $_ } {
        foreach ($ds in $discoverySettings) {
            $dsId = $ds.id; switch ($null -ne $dsId) { $true { $dsIdToAddr[[int]$dsId] = $ds.address } }
            $addr = $ds.address
            $targets = @()
            switch ($addr) { { $_ } {
                $addrStr = ([string]$addr).Trim()
                switch ($addrStr -match '^(\d+\.\d+\.\d+\.\d+)(/(\d+))?$') {
                    $true { $targets += $addrStr }
                    default {
                        $parts = $addrStr -split '[,\s;]+' | Where-Object { $_ -match '^\d+\.\d+\.\d+\.\d+(\/\d+)?$' }
                        foreach ($a in $parts) { $targets += $a.Trim() }
                    }
                }
            } }
            $tipVal = $ds.target_ip; switch (-not $tipVal) { $true { $tipVal = $ds.target_ips } }
            switch (-not $tipVal) { $true { $tipVal = $ds.targetIp } }
            switch ($tipVal) { { $_ } {
                $tip = $tipVal
                switch ($tip -is [array]) { $true { foreach ($x in $tip) { $targets += ([string]$x).Trim() } } default { foreach ($x in (([string]$tip) -split '[,\s;]+' | Where-Object { $_ })) { $targets += ([string]$x).Trim() } } }
            } }
            foreach ($t in $targets) { $tStr = ([string]$t -replace '\s+', '').Trim(); switch ($tStr) { { $_ } { [void]$result.ScanTargets.Add($tStr) } } }
            switch ($addr -match '^(\d+\.\d+\.\d+\.\d+)/(\d+)$') { $true {
                $targetsForValidation = @($addr)
                switch ($tipVal) { { $_ } {
                    $tip = $tipVal
                    switch ($tip -is [array]) { $true { foreach ($x in $tip) { $targetsForValidation += [string]$x } } default { foreach ($x in (([string]$tip) -split '[,\s;]+' | Where-Object { $_ })) { $targetsForValidation += $x } } }
                } }
                $issues = Test-ExternalSubnetConfig -Targets $targetsForValidation -Cidr $addr
                switch ($issues) { { $_ } { foreach ($i in $issues) { [void]$result.SubnetIssues.Add($i) } } }
            } }
        }
    } }
    foreach ($m in $discMappings) {
        $dsId = $m.discovery_settings_id; switch (-not $dsId) { $true { $dsId = $m.discoverysettings_id } }
        switch ($null -ne $dsId -and $dsIdToAddr.ContainsKey([int]$dsId)) { $true {
            $a = $dsIdToAddr[[int]$dsId]
            switch ($a -and -not $result.ProbesSubnets.Contains($a)) { $true { [void]$result.ProbesSubnets.Add($a) } }
        } }
    }

    # 4b. External assets (external scan endpoints) - from discovery_settings externalscan
    $extScanDs = @(Get-ConnectSecureExternalScanDiscoverySettings -CompanyId $CompanyId)
    switch ($extScanDs.Count -eq 0) { $true {
        $allDs = Get-ConnectSecureDiscoverySettings -CompanyId $CompanyId
        $extScanDs = @($allDs | Where-Object {
            $t = $_.discovery_settings_type; switch (-not $t) { $true { $t = $_.type } }
            $t -and [string]$t -match 'external|externalscan'
        })
    } }
    foreach ($ds in $extScanDs) {
        $dsName = $ds.discovery_settings_name; switch (-not $dsName) { $true { $dsName = $ds.name } }
        $dsAddr = $ds.address; switch (-not $dsAddr) { $true { $dsAddr = $ds.target_ip } }
        switch (-not $dsAddr) { $true { $dsAddr = $ds.targetIp } }
        switch ($dsAddr -is [array]) { $true { $dsAddr = ($dsAddr | Where-Object { $_ }) -join ', ' } }
        switch ($dsAddr) { { $_ } {
            $addrStr = ([string]$dsAddr).Trim()
            [void]$result.ExternalAssets.Add(@{ Name = [string]$dsName; Address = $addrStr })
            $parts = $addrStr -split '[,\s;]+' | Where-Object { $_ -match '^\d+\.\d+\.\d+\.\d+(-|\/|$)' }
            foreach ($p in $parts) { $t = ([string]$p -replace '\s+', '').Trim(); switch ($t) { { $_ } { [void]$result.ScanTargets.Add($t) } } }
        } }
    }

    # 5. Agents/Probes offline 7+, 14+, 30+ days - based on last_ping_time of agents OR probes
    # Exclude deprecated agents (already auto-archived); counts show only active agents that need attention
    $now = Get-Date
    $oldestDays = $null
    $offline7Plus = 0
    $offline14Plus = 0
    $offline30Plus = 0
    $offlineSource = @($agents | Where-Object {
        $dep = $_.is_deprecated; if ($null -eq $dep) { $dep = $_.isDeprecated }
        -not $dep
    })
    foreach ($a in $offlineSource) {
        $lp = $a.last_ping_time; switch (-not $lp) { $true { $lp = $a.lastPingTime } }
        switch (-not $lp) { $true { $lp = $a.'Last Ping Time' } }
        switch (-not $lp) { $true { continue } }
        try {
            $dt = [DateTime]::Parse($lp)
            $days = ($now - $dt).TotalDays
            switch ($null -eq $oldestDays -or $days -gt $oldestDays) { $true { $oldestDays = [int][Math]::Floor($days) } }
            if ($days -ge 7) { $offline7Plus++ }
            if ($days -ge 14) { $offline14Plus++ }
            if ($days -ge 30) {
                $offline30Plus++
                $an = $a.host_name; switch (-not $an) { $true { $an = $a.'Host Name' } }
                switch (-not $an) { $true { $an = $a.agent_name } }
                switch (-not $an) { $true { $an = $a.hostname } }
                switch (-not $an) { $true { $an = $a.name } }
                switch (-not $an) { $true { $an = $a.computer_name } }
                switch (-not $an) { $true { $an = $a.display_name } }
                switch (-not $an) { $true { $an = $a.asset_name } }
                switch (-not $an) { $true { $ip = $a.ip; if ($ip) { $an = "IP: $ip" } } }
                switch (-not $an) { $true { $aid = $a.id; if ($null -ne $aid) { $an = "Agent #$aid" } } }
                [void]$result.AgentsOffline30PlusNames.Add([string](if ($an) { $an } else { "Unknown" }))
            }
        } catch { }
    }
    $result.OldestOfflineDays = $oldestDays
    $result.AgentsOffline7PlusDays = $offline7Plus
    $result.AgentsOffline14PlusDays = $offline14Plus
    $result.AgentsOffline30PlusDays = $offline30Plus

    # 6. Firewall - managed devices (Cisco Meraki, Fortigate, Sonicwall, etc.) per client
    # Portal uses /r/report_queries/firewall_asset_view with condition=company_id=X (per capture)
    $fwAssets = @()
    try {
        $qp = @{ condition = "company_id=$CompanyId"; limit = 500; skip = 0; order_by = "host_name asc" }
        $raw = Invoke-ConnectSecureCompanyReviewRequest -Endpoint $script:CompanyReviewEndpoints.FirewallAssetView -QueryParams $qp
        $fwAssets = if ($null -eq $raw) { @() } elseif ($raw -is [array]) { @($raw) } else { @(,$raw) }
    } catch { }

    $result.FirewallCount = $fwAssets.Count
    $manufacturers = @{}
    foreach ($a in $fwAssets) {
        $isFw = $a.is_firewall; if ($null -eq $isFw) { $isFw = $a.isFirewall }
        if (-not $isFw) { continue }
        $mfr = $a.manufacturer; if ($null -eq $mfr) { $mfr = $a.asset_type }
        if ($mfr) { $manufacturers[[string]$mfr] = $true }
    }
    if ($manufacturers.Count -eq 0 -and $result.FirewallCount -gt 0) {
        $manufacturers['Managed firewall'] = $true
    }

    $result.FirewallActive = $result.FirewallCount -gt 0
    $result.FirewallType = ($manufacturers.Keys | Sort-Object) -join ', '

    # 7. Last scan dates - internal from probe agents; external from jobs_view (server-side, no assets fetch)
    $lastInternal = $null
    $probeAgentIds = @{}
    foreach ($aid in $agentIdsWithCreds.Keys) { switch ($agentIdsWithNetworks[$aid]) { $true { $probeAgentIds[$aid] = $true } } }
    foreach ($a in $agents) {
        $aid = $a.id; switch ($null -eq $aid) { $true { continue } }
        switch (-not $probeAgentIds[[int]$aid]) { $true { continue } }
        $lst = $a.last_scanned_time; switch (-not $lst) { $true { $lst = $a.lastScannedTime } }
        switch ($lst) { { $_ } {
            try { $dt = [DateTime]::Parse($lst); switch ($null -eq $lastInternal -or $dt -gt $lastInternal) { $true { $lastInternal = $dt } } } catch { }
        } }
    }
    $result.LastInternalScan = switch ($lastInternal) { { $_ } { $lastInternal.ToLocalTime().ToString('yyyy-MM-dd HH:mm') } default { $null } }
    $stats = Get-ConnectSecureCompanyStats -CompanyId $CompanyId
    $hasStats = $stats -and $stats.PSObject.Properties.Count -gt 0
    switch (-not $result.LastInternalScan -and $hasStats) { $true {
        $result.LastInternalScan = $stats.ad_last_scan_time
        switch (-not $result.LastInternalScan) { $true { $result.LastInternalScan = $stats.date } }
    } }
    $lastExternal = $null
    $jobsView = @()
    try {
        $qp = @{ condition = "company_id=$CompanyId and type='External Scan'"; limit = 50; skip = 0; order_by = 'updated desc' }
        $jobsData = Invoke-ConnectSecureCompanyReviewRequest -Endpoint $script:CompanyReviewEndpoints.JobsView -QueryParams $qp
        $jobsView = if ($null -eq $jobsData) { @() } elseif ($jobsData -is [array]) { @($jobsData) } else { @(,$jobsData) }
    } catch { }
    foreach ($j in $jobsView) {
        $u = $j.updated; switch (-not $u) { $true { $u = $j.created } }
        switch ($u) { { $_ } {
            try { $dt = [DateTime]::Parse($u); switch ($null -eq $lastExternal -or $dt -gt $lastExternal) { $true { $lastExternal = $dt } } } catch { }
        } }
    }
    $result.LastExternalScan = switch ($lastExternal) { { $_ } { $lastExternal.ToLocalTime().ToString('yyyy-MM-dd HH:mm') } default { $null } }
    switch (-not $result.LastExternalScan -and $hasStats) { $true {
        $result.LastExternalScan = $stats.external_last_scan_time
        switch (-not $result.LastExternalScan) { $true { $result.LastExternalScan = $stats.date } }
        switch (-not $result.LastExternalScan) { $true { $result.LastExternalScan = $stats.updated } }
        switch ($result.LastExternalScan) { { $_ } {
            try { $dt = [DateTime]::Parse($result.LastExternalScan); $result.LastExternalScan = $dt.ToLocalTime().ToString('yyyy-MM-dd HH:mm') } catch { }
        } }
    } }

    # 8. Quick wins
    switch ($result.AgentCount -eq 0) { $true { [void]$result.QuickWins.Add('Add lightweight agents to enable internal scanning') } }
    switch ($result.ProbesWithBoth -eq 0) { $true { [void]$result.QuickWins.Add('Map credentials and discovery networks to at least one probe agent') } }
    $probesWithNmap = ($result.ProbeAgentsNmapInfo | Where-Object { $_.NmapInterface -and $_.NmapInterface -ne '(not set)' }).Count
    switch ($result.ProbesWithBoth -gt 0 -and $probesWithNmap -eq 0) { $true { [void]$result.QuickWins.Add('Configure nmap interface on probe agent(s) for scanning') } }
    switch ($result.SubnetIssues.Count -gt 0) { $true { [void]$result.QuickWins.Add('Exclude network and broadcast addresses from external scan targets') } }
    switch ($result.AgentsOffline30PlusDays -gt 0) { $true { [void]$result.QuickWins.Add('Investigate agents/probes offline more than 30 days; reinstall or remove') } }
    switch (-not $result.FirewallActive) { $true { [void]$result.QuickWins.Add('Configure firewall integration for visibility') } }
    switch (-not $result.LastInternalScan -and -not $result.LastExternalScan) { $true { [void]$result.QuickWins.Add('Run internal and external scans to populate vulnerability data') } }

    return $result
}

function Get-ConnectSecureCompanyAssetIds {
    <#
    .SYNOPSIS
    Returns the set of asset IDs that belong to the given company. Used for filtering asset_wise data when rows lack company_ids.
    #>
    param([int]$CompanyId, [array]$Assets = $null)
    if ($CompanyId -le 0) { return @{} }
    $assets = if ($null -ne $Assets) { $Assets } else {
        try { Get-ConnectSecureAssets -Limit 5000 -FetchAll:$true } catch { return @{} }
    }
    if (-not $assets -or $assets.Count -eq 0) { return @{} }
    $cidStr = [string]$CompanyId
    $companyAssets = $assets | Where-Object { Test-RowMatchesCompanyId -Row $_ -CompanyIdStr $cidStr }
    $idSet = @{}
    foreach ($a in $companyAssets) {
        $id = $null
        if ($null -ne $a.id) { $id = [string]$a.id }
        elseif ($null -ne $a.asset_id) { $id = [string]$a.asset_id }
        elseif ($null -ne $a._id) { $id = [string]$a._id }
        if ($id) { $idSet[$id] = $true }
    }
    return $idSet
}

function Test-AssetWiseRowMatchesCompany {
    param([object]$Row, [hashtable]$CompanyAssetIds)
    if (-not $CompanyAssetIds -or $CompanyAssetIds.Count -eq 0) { return $false }
    $obj = if ($null -ne $Row._source) { $Row._source } else { $Row }
    if ($null -eq $obj) { return $false }
    $aidVal = $obj.asset_ids
    if ($null -eq $aidVal) { return $false }
    $ids = if ($aidVal -is [array]) { $aidVal | ForEach-Object { [string]$_ } } elseif ($aidVal -is [string]) { ($aidVal -split ';') | ForEach-Object { $_.Trim() } | Where-Object { $_ } } else { @([string]$aidVal) }
    foreach ($id in $ids) {
        if ($CompanyAssetIds.ContainsKey($id)) { return $true }
    }
    return $false
}

function Get-ConnectSecureAssetWiseVulnerabilities {
    <#
    .SYNOPSIS
    Fetches asset-wise vulnerabilities (one row per host+vuln) from /r/report_queries/asset_wise_vulnerabilities.
    Use for 13-column export: CVE ID, Severity, CVSS, EPSS, Asset Name, OS, IP, Description, Application Name, Product Name, First Seen, Last Seen, Owner.
    .PARAMETER Assets
    Optional. Pre-fetched assets for company filtering. If provided, avoids a separate asset fetch when filtering by company.
    #>
    param(
        [int]$CompanyId = 0,
        [int]$Limit = 5000,
        [int]$Skip = 0,
        [string]$Filter = "",
        [string]$Sort = 'severity.keyword:desc',
        [switch]$FetchAll,
        [array]$Assets = $null
    )
    $endpoint = '/r/report_queries/asset_wise_vulnerabilities'
    $doFetchAll = if ($FetchAll) { $true } else { $false }
    Write-CSApiLog ('Fetching asset-wise vulnerabilities (CompanyId: ' + $CompanyId + ', Limit: ' + $Limit + ', Skip: ' + $Skip + ')...') -Level Info
    $allData = @()
    $currentSkip = $Skip
    $pageSize = $Limit
    $maxPages = 200
    do {
        $queryParams = @{ limit = $pageSize; skip = $currentSkip; sort = $Sort }
        if ($CompanyId -gt 0) {
            $queryParams.company_id = $CompanyId
            if ($script:UseConditionForCompanyFilter) {
                $queryParams.condition = 'company_ids:' + $CompanyId
            }
        }
        if (-not [string]::IsNullOrWhiteSpace($Filter)) { $queryParams.filter = $Filter }
        try {
            $response = Invoke-ConnectSecureRequest -Endpoint $endpoint -QueryParameters $queryParams
            if ($response.status -and $response.data) {
                $allData += $response.data
                $pageNum = [Math]::Floor($currentSkip / $pageSize) + 1
                Write-CSApiLog ('Page ' + $pageNum + ': retrieved ' + $response.data.Count + ' (Total: ' + $allData.Count + ')') -Level Info
                $maxRec = $script:VulnMaxRecords
                if ($maxRec -gt 0 -and $allData.Count -ge $maxRec) {
                    Write-CSApiLog ('Reached record limit (' + $maxRec + ').') -Level Warning
                    break
                }
                if ($doFetchAll -and $response.data.Count -eq $pageSize -and $pageNum -lt $maxPages) {
                    $currentSkip += $pageSize
                } else { break }
            } else { break }
        } catch {
            Write-CSApiLog ('Error fetching asset-wise vulnerabilities: ' + $_.Exception.Message) -Level Error
            throw
        }
    } while ($doFetchAll)
    # Client-side filter: API returns tenant-wide data (asset_wise ignores company_id per swagger)
    # Try company_ids on row first; if that fails (0 rows or no change), filter by asset IDs belonging to company
    if ($CompanyId -gt 0 -and $allData.Count -gt 0) {
        $before = $allData.Count
        $cidStr = [string]$CompanyId
        $byCompanyIds = $allData | Where-Object { Test-RowMatchesCompanyId -Row $_ -CompanyIdStr $cidStr }
        if ($byCompanyIds.Count -gt 0 -and $byCompanyIds.Count -lt $before) {
            $allData = $byCompanyIds
            Write-CSApiLog ('Filtered asset-wise by company_ids to company ' + $CompanyId + ': ' + $before + ' -> ' + $allData.Count + ' rows') -Level Info
        } else {
            # Rows often lack company_ids; filter by asset_ids belonging to company
            $companyAssetIds = Get-ConnectSecureCompanyAssetIds -CompanyId $CompanyId -Assets $Assets
            if ($companyAssetIds.Count -gt 0) {
                $filtered = $allData | Where-Object { Test-AssetWiseRowMatchesCompany -Row $_ -CompanyAssetIds $companyAssetIds }
                if ($filtered.Count -gt 0) {
                    $allData = $filtered
                    Write-CSApiLog ('Filtered asset-wise by asset IDs for company ' + $CompanyId + ': ' + $before + ' -> ' + $allData.Count + ' rows') -Level Info
                } else {
                    Write-CSApiLog ('Asset filter would remove all rows (asset ID format may differ); keeping unfiltered data') -Level Warning
                }
            } else {
                Write-CSApiLog ('No assets found for company ' + $CompanyId + '; keeping unfiltered data') -Level Warning
            }
        }
    }
    Write-CSApiLog ('Total asset-wise vulnerabilities retrieved: ' + $allData.Count) -Level Success
    return $allData
}

function Get-ConnectSecureAssetDetailMap {
    <#
    .SYNOPSIS
    Builds hashtable asset_id -> @{os_name; os_version; asset_owner} for 13-column OS/Owner enrichment.
    #>
    param([array]$Assets)
    $map = @{}
    if (-not $Assets -or $Assets.Count -eq 0) { return $map }
    foreach ($a in $Assets) {
        $id = $null
        if ($null -ne $a.id) { $id = $a.id }
        elseif ($null -ne $a.asset_id) { $id = $a.asset_id }
        if ($null -eq $id) { continue }
        $key = [string]$id
        if ($map.ContainsKey($key)) { continue }
        $osName = ''
        if ($null -ne $a.os_name -and -not [string]::IsNullOrWhiteSpace([string]$a.os_name)) { $osName = [string]$a.os_name }
        elseif ($null -ne $a.os_full_name -and -not [string]::IsNullOrWhiteSpace([string]$a.os_full_name)) { $osName = [string]$a.os_full_name }
        $osVer = if ($null -ne $a.os_version -and -not [string]::IsNullOrWhiteSpace([string]$a.os_version)) { [string]$a.os_version } else { '' }
        $owner = if ($null -ne $a.asset_owner -and -not [string]::IsNullOrWhiteSpace([string]$a.asset_owner)) { [string]$a.asset_owner } else { '' }
        $map[$key] = @{ os_name = $osName; os_version = $osVer; asset_owner = $owner }
    }
    return $map
}

function ConvertTo-ConnectSecure13ColumnFormat {
    <#
    .SYNOPSIS
    Converts asset_wise vulnerability data to 13-column format with OS/Owner enrichment.
    #>
    param(
        [array]$AssetWiseData,
        [hashtable]$AssetDetailMap
    )
    if (-not $AssetWiseData -or $AssetWiseData.Count -eq 0) { return @() }
    $result = @()
    foreach ($v in $AssetWiseData) {
        $aid = $v.asset_ids
        $assetId = ''
        if ($null -ne $aid) {
            if ($aid -is [array] -and $aid.Count -gt 0) { $assetId = [string]$aid[0] }
            elseif ($aid -is [string]) { $assetId = ($aid -split ';')[0].Trim() }
            else { $assetId = [string]$aid }
        }
        $osVal = ''; $ownerVal = ''
        if ($assetId -and $AssetDetailMap -and $AssetDetailMap.ContainsKey($assetId)) {
            $d = $AssetDetailMap[$assetId]
            $osVal = $d.os_name
            if ($d.os_version) { $osVal = ($osVal + ' ' + $d.os_version).Trim() }
            $ownerVal = $d.asset_owner
        }
        $app = $v.software_name
        $appStr = if ($null -eq $app) { '' } elseif ($app -is [array]) { ($app | ForEach-Object { [string]$_ }) -join '; ' } else { [string]$app }
        $fixVal = Get-ConnectSecureVulnField -Item $v -Paths @('solution', 'fix', 'remediation', '_source.solution', '_source.fix', '_source.remediation')
        if ($null -eq $fixVal) { $fixVal = Get-ConnectSecureVulnFieldDeep -Item $v -FieldNames @('Solution', 'solution', 'fix', 'remediation') }
        $fixStr = if ($null -ne $fixVal -and -not [string]::IsNullOrWhiteSpace([string]$fixVal)) { [string]$fixVal } else { '' }
        $result += [PSCustomObject]@{
            'Source'           = 'Application'
            'CVE ID'           = if ($v.problem_name) { [string]$v.problem_name } else { '' }
            'Severity'         = if ($v.severity) { [string]$v.severity } else { '' }
            'CVSS Score'       = if ($null -ne $v.base_score) { $v.base_score } else { '' }
            'EPSS Score'       = if ($null -ne $v.epss_score) { $v.epss_score } else { '' }
            'Asset Name'       = if ($v.host_name) { [string]$v.host_name } else { '' }
            'OS'               = $osVal
            'IP Address'       = if ($v.ip) { [string]$v.ip } else { '' }
            'Description'      = if ($v.description) { [string]$v.description } else { '' }
            'Application Name' = $appStr
            'Product Name'     = $appStr
            'First Seen'       = if ($v.discovered) { [string]$v.discovered } else { '' }
            'Last Seen'        = if ($v.last_discovered_time) { [string]$v.last_discovered_time } else { '' }
            'Owner'            = $ownerVal
            'Fix'              = $fixStr
        }
    }
    return $result
}

function ConvertFrom-RegistryProblemsFormat {
    <#
    .SYNOPSIS
    Maps registry_problems_remediation rows to 13-column format for Excel export.
    Portal format: name, affected_assets, asset_ids, base_score, description, epss_score, fix, severity, url, vul_id.
    No host_name/ip per row; resolves asset_ids to host names via AssetMap.
    #>
    param([array]$Data, [hashtable]$AssetDetailMap = $null, [hashtable]$AssetMap = $null)
    if (-not $Data -or $Data.Count -eq 0) { return @() }
    $result = @()
    foreach ($v in $Data) {
        $obj = if ($null -ne $v._source) { $v._source } else { $v }
        $assetIds = $obj.asset_ids
        $hostNames = @()
        if ($assetIds -and $AssetMap) {
            $ids = if ($assetIds -is [array]) { @($assetIds) } else { @([string]$assetIds) }
            foreach ($aid in $ids) {
                $k = [string]$aid
                if ($AssetMap.ContainsKey($k)) { $hostNames += $AssetMap[$k] } else { $hostNames += $k }
            }
        }
        $assetNameStr = if ($hostNames.Count -gt 0) { ($hostNames -join '; ').Trim() } else { '' }
        if (-not $assetNameStr -and $obj.affected_assets) { $assetNameStr = "($($obj.affected_assets) affected)" }
        $osVal = ''; $ownerVal = ''
        if ($AssetDetailMap -and $assetIds) {
            $firstId = if ($assetIds -is [array] -and $assetIds.Count -gt 0) { [string]$assetIds[0] } else { [string]$assetIds }
            if ($AssetDetailMap.ContainsKey($firstId)) {
                $d = $AssetDetailMap[$firstId]
                $osVal = $d.os_name; if ($d.os_version) { $osVal = ($osVal + ' ' + $d.os_version).Trim() }
                $ownerVal = $d.asset_owner
            }
        }
        $nameVal = if ($obj.name) { [string]$obj.name } else { '' }
        $lastSeenVal = if ($obj.remediated_on) { [string]$obj.remediated_on } else { '' }
        $result += [PSCustomObject]@{
            'Source'           = 'Registry'
            'CVE ID'           = $nameVal
            'Severity'         = if ($obj.severity) { [string]$obj.severity } else { '' }
            'CVSS Score'       = if ($null -ne $obj.base_score) { $obj.base_score } else { '' }
            'EPSS Score'       = if ($null -ne $obj.epss_score) { $obj.epss_score } else { '' }
            'Asset Name'       = $assetNameStr
            'OS'               = $osVal
            'IP Address'       = ''
            'Description'      = if ($obj.description) { [string]$obj.description } else { '' }
            'Application Name' = $nameVal
            'Product Name'     = $nameVal
            'First Seen'       = ''
            'Last Seen'        = $lastSeenVal
            'Owner'            = $ownerVal
            'Fix'              = if ($obj.fix) { [string]$obj.fix } else { '' }
        }
    }
    return $result
}

function ConvertFrom-NetworkVulnerabilitiesFormat {
    <#
    .SYNOPSIS
    Maps application_vulnerabilities_net rows to flat Excel format.
    Swagger: software_name, cvss_score, severity, affected_assets, ports[], title, affected_operating_systems, software_version.
    No per-asset host_name/ip - software-level aggregated.
    #>
    param([array]$Data)
    if (-not $Data -or $Data.Count -eq 0) { return @() }
    $result = @()
    foreach ($v in $Data) {
        $obj = if ($null -ne $v._source) { $v._source } else { $v }
        $portsStr = ''
        if ($obj.ports -and $obj.ports.Count -gt 0) {
            $portsStr = ($obj.ports | ForEach-Object { [string]$_ }) -join '; '
        }
        $result += [PSCustomObject]@{
            'Product Name'      = if ($obj.software_name) { [string]$obj.software_name } else { '' }
            'Severity'          = if ($obj.severity) { [string]$obj.severity } else { '' }
            'CVSS Score'        = if ($null -ne $obj.cvss_score) { $obj.cvss_score } else { '' }
            'Affected Assets'   = if ($null -ne $obj.affected_assets) { $obj.affected_assets } else { '' }
            'Ports'             = $portsStr
            'Title'             = if ($obj.title) { [string]$obj.title } else { '' }
            'Affected OS'       = if ($obj.affected_operating_systems) { [string]$obj.affected_operating_systems } else { '' }
            'Software Version'  = if ($obj.software_version) { [string]$obj.software_version } else { '' }
            'Software Type'     = if ($obj.software_type) { [string]$obj.software_type } else { '' }
        }
    }
    return $result
}

function ConvertFrom-NetworkTo13ColumnFormat {
    <#
    .SYNOPSIS
    Maps application_vulnerabilities_net rows to 13-column format for inclusion in All Vulnerabilities.
    Used when merging Registry and Network into Top N / ticket instructions. Asset Name = "(Network scan)".
    #>
    param([array]$Data)
    if (-not $Data -or $Data.Count -eq 0) { return @() }
    $result = @()
    foreach ($v in $Data) {
        $obj = if ($null -ne $v._source) { $v._source } else { $v }
        $prod = if ($obj.software_name) { [string]$obj.software_name } else { '' }
        $title = if ($obj.title) { [string]$obj.title } else { '' }
        $affStr = if ($null -ne $obj.affected_assets) { "Affects $($obj.affected_assets) assets" } else { '' }
        $desc = if ($title) { $title } else { $affStr }
        $result += [PSCustomObject]@{
            'Source'           = 'Network'
            'CVE ID'           = $title
            'Severity'         = if ($obj.severity) { [string]$obj.severity } else { '' }
            'CVSS Score'       = if ($null -ne $obj.cvss_score) { $obj.cvss_score } else { '' }
            'EPSS Score'       = ''
            'Asset Name'       = '(Network scan)'
            'OS'               = if ($obj.affected_operating_systems) { [string]$obj.affected_operating_systems } else { '' }
            'IP Address'       = ''
            'Description'      = $desc
            'Application Name' = $prod
            'Product Name'     = $prod
            'First Seen'       = ''
            'Last Seen'        = ''
            'Owner'            = ''
            'Fix'              = ''
        }
    }
    return $result
}

function Get-ConnectSecureAssetIdToNameMap {
    <#
    .SYNOPSIS
    Builds a hashtable mapping asset ID -> asset name (hostname, name, or id as fallback).
    Call Get-ConnectSecureAssets first and pass the result.
    #>
    param([array]$Assets)
    $map = @{}
    if (-not $Assets -or $Assets.Count -eq 0) { return $map }
    foreach ($a in $Assets) {
        $id = $null
        if ($null -ne $a.id) { $id = $a.id }
        elseif ($null -ne $a.asset_id) { $id = $a.asset_id }
        elseif ($null -ne $a._id) { $id = $a._id }
        if ($null -eq $id) { continue }
        $key = [string]$id
        if ($map.ContainsKey($key)) { continue }
        $name = $null
        if ($null -ne $a.hostname -and -not [string]::IsNullOrWhiteSpace([string]$a.hostname)) { $name = [string]$a.hostname }
        elseif ($null -ne $a.host_name -and -not [string]::IsNullOrWhiteSpace([string]$a.host_name)) { $name = [string]$a.host_name }
        elseif ($null -ne $a.name -and -not [string]::IsNullOrWhiteSpace([string]$a.name)) { $name = [string]$a.name }
        elseif ($null -ne $a.asset_name -and -not [string]::IsNullOrWhiteSpace([string]$a.asset_name)) { $name = [string]$a.asset_name }
        elseif ($null -ne $a.fqdn_name -and -not [string]::IsNullOrWhiteSpace([string]$a.fqdn_name)) { $name = [string]$a.fqdn_name }
        if ([string]::IsNullOrWhiteSpace($name)) { $name = $key }
        $map[$key] = $name
    }
    return $map
}

function Invoke-AssetNameResolution {
    <#
    .SYNOPSIS
    Fetches assets, builds id->name map, and enriches data with asset_names. If asset fetch fails, returns data unchanged.
    #>
    param(
        [array]$Data,
        [int]$CompanyId = 0
    )
    if (-not $Data -or $Data.Count -eq 0) { return $Data }
    try {
        $assets = Get-ConnectSecureAssets -CompanyId $CompanyId -Limit 5000 -FetchAll:$true
        $assetMap = Get-ConnectSecureAssetIdToNameMap -Assets $assets
        if ($assetMap.Count -gt 0) {
            return Resolve-AssetIdsToNames -Data $Data -AssetMap $assetMap
        }
    } catch {
        Write-CSApiLog ('Asset name resolution skipped: ' + $_.Exception.Message) -Level Warning
    }
    return $Data
}

function Resolve-AssetIdsToNames {
    <#
    .SYNOPSIS
    Enriches vulnerability data: adds asset_names column by resolving asset_ids using the asset lookup map.
    If asset_ids is present, creates asset_names (semicolon-separated names).
    #>
    param(
        [array]$Data,
        [hashtable]$AssetMap,
        [switch]$ReplaceAssetIds
    )
    if (-not $Data -or $Data.Count -eq 0 -or -not $AssetMap) { return $Data }
    $result = @()
    foreach ($row in $Data) {
        $obj = if ($null -ne $row -and $row.PSObject.Properties['_source']) { $row._source } else { $row }
        $aidVal = $null
        if ($null -ne $obj) {
            if ($obj.PSObject.Properties['asset_ids']) { $aidVal = $obj.asset_ids }
            elseif ($obj -is [hashtable] -and $obj.ContainsKey('asset_ids')) { $aidVal = $obj['asset_ids'] }
        }
        $nameStr = ''
        if ($null -ne $aidVal) {
            $ids = if ($aidVal -is [string]) { ($aidVal -split ';') | ForEach-Object { $_.Trim() } | Where-Object { $_ } } else { @($aidVal) | Where-Object { $null -ne $_ } }
            $names = @()
            foreach ($id in $ids) {
                $k = [string]$id
                if ($AssetMap.ContainsKey($k)) { $names += $AssetMap[$k] } else { $names += $k }
            }
            $nameStr = $names -join '; '
        }
        # Build new object with asset_names (and optionally replace asset_ids)
        $props = @{}
        if ($obj -is [PSCustomObject]) {
            $obj.PSObject.Properties | ForEach-Object { $props[$_.Name] = $_.Value }
        } elseif ($obj -is [hashtable]) {
            $obj.Keys | ForEach-Object { $props[$_] = $obj[$_] }
        } else {
            $result += $row
            continue
        }
        if ($nameStr) {
            $props['asset_names'] = $nameStr
            if ($ReplaceAssetIds) { $props['asset_ids'] = $nameStr }
        }
        $result += [PSCustomObject]$props
    }
    return $result
}

function Convert-ConnectSecureDataToExcelFormat {
    param(
        [array]$ConnectSecureData,
        [scriptblock]$OnProgress = $null
    )

    Write-CSApiLog 'Converting ConnectSecure data to Excel format...' -Level Info

    $excelData = @()
    $total = if ($ConnectSecureData) { $ConnectSecureData.Count } else { 0 }
    $progressInterval = if ($total -gt 0) { [Math]::Max(100, [Math]::Floor($total / 20)) } else { 99999 }

    # ConnectSecure native report uses: Host Name, FQDN Name, IP, Software Name, CVE, Problem Name, Solution, Severity
    $hostPaths = @('Host Name','host_name','hostName','hostname','fqdn_name','fqdn','host.hostname','host.host_name','host.name','asset_name','computer_name','device_name','asset.hostname','asset.host_name','machine.hostname','device.hostname','_source.Host Name','_source.host_name','_source.hostname','_source.fqdn_name','_source.host.hostname','_source.asset.hostname')
    $ipPaths = @('IP','ip','ip_address','ipAddress','host.ip','host.ip_address','asset.ip','asset.ip_address','_source.IP','_source.ip','_source.ip_address','_source.host.ip','_source.asset.ip')
    $productPaths = @('Software Name','software_name','softwareName','product','name','cpe','vulnerability.product','vulnerability.software_name','asset.product','_source.Software Name','_source.product','_source.software_name','_source.name')
    $sevPaths = @('Severity','severity','severity.keyword','vulnerability.severity','vulnerability.severity.keyword','_source.Severity','_source.severity','_source.severity.keyword')
    $idx = 0
    foreach ($item in $ConnectSecureData) {
        if ($OnProgress -and $total -gt 0 -and $progressInterval -gt 0 -and ($idx % $progressInterval -eq 0)) {
            $pct = [Math]::Min(99, [Math]::Floor(100 * $idx / $total))
            $msg = 'Converting to Excel format... {0} of {1} ({2} percent)' -f $idx, $total, $pct
            & $OnProgress $msg
        }
        $idx++
        $src = if ($item._source) { $item._source } else { $item }
        # Aggregated format: one row per vuln with AffectedHosts, HostCount
        if ($src._aggregated) {
            $hostVal = $src.AffectedHosts
            if ($src.HostCount) { $hostVal = ('(' + $src.HostCount + ' hosts) ') + $hostVal }
            $ipVal = ""
            $prodVal = $src.software_name
            $sevVal = $src.severity
            $cveVal = $src.problem_name
            $descVal = $src.description
            $epssVal = $src.epss_score
            $fixVal = ""; $evPathVal = ""; $evVerVal = ""
        } else {
        # Path-based lookup first, then deep recursive search as fallback
        $hostVal = Get-ConnectSecureVulnField -Item $src -Paths $hostPaths
        if ($null -eq $hostVal) { $hostVal = Get-ConnectSecureVulnField -Item $item -Paths $hostPaths }
        if ($null -eq $hostVal) { $hostVal = Get-ConnectSecureVulnFieldDeep -Item $item -FieldNames @('Host Name','host_name','hostName','hostname','fqdn_name','asset_name','computer_name','device_name') }
        $ipVal = Get-ConnectSecureVulnField -Item $src -Paths $ipPaths
        if ($null -eq $ipVal) { $ipVal = Get-ConnectSecureVulnField -Item $item -Paths $ipPaths }
        if ($null -eq $ipVal) { $ipVal = Get-ConnectSecureVulnFieldDeep -Item $item -FieldNames @('ip','ip_address','ipAddress') }
        $prodVal = Get-ConnectSecureVulnField -Item $src -Paths $productPaths
        if ($null -eq $prodVal) { $prodVal = Get-ConnectSecureVulnField -Item $item -Paths $productPaths }
        $sevVal = Get-ConnectSecureVulnField -Item $src -Paths $sevPaths
        if ($null -eq $sevVal) { $sevVal = Get-ConnectSecureVulnField -Item $item -Paths $sevPaths }
        if ($null -eq $sevVal) { $sevVal = Get-ConnectSecureVulnFieldDeep -Item $item -FieldNames @('Severity','severity') }
        $epssVal = Get-ConnectSecureVulnField -Item $src -Paths @('epss_score','epssScore','_source.epss_score')
        if ($null -eq $epssVal) { $epssVal = Get-ConnectSecureVulnField -Item $item -Paths @('epss_score','epssScore') }
        if ($null -eq $epssVal) { $epssVal = Get-ConnectSecureVulnFieldDeep -Item $item -FieldNames @('epss_score','epssScore') }
        $fixVal = Get-ConnectSecureVulnField -Item $src -Paths @('Solution','solution','fix','remediation','_source.Solution','_source.solution','_source.fix')
        if ($null -eq $fixVal) { $fixVal = Get-ConnectSecureVulnField -Item $item -Paths @('Solution','solution','fix') }
        if ($null -eq $fixVal) { $fixVal = Get-ConnectSecureVulnFieldDeep -Item $item -FieldNames @('Solution','solution','fix','remediation') }
        $evPathVal = Get-ConnectSecureVulnField -Item $src -Paths @('evidence_path','evidencePath','_source.evidence_path')
        if ($null -eq $evPathVal) { $evPathVal = Get-ConnectSecureVulnField -Item $item -Paths @('evidence_path') }
        if ($null -eq $evPathVal) { $evPathVal = Get-ConnectSecureVulnFieldDeep -Item $item -FieldNames @('evidence_path','evidencePath') }
        $evVerVal = Get-ConnectSecureVulnField -Item $src -Paths @('evidence_version','evidenceVersion','_source.evidence_version')
        if ($null -eq $evVerVal) { $evVerVal = Get-ConnectSecureVulnField -Item $item -Paths @('evidence_version') }
        if ($null -eq $evVerVal) { $evVerVal = Get-ConnectSecureVulnFieldDeep -Item $item -FieldNames @('evidence_version','evidenceVersion') }
        $cveVal = Get-ConnectSecureVulnField -Item $src -Paths @('CVE','cve','cve_id','_source.CVE','_source.cve')
        if ($null -eq $cveVal) { $cveVal = Get-ConnectSecureVulnField -Item $item -Paths @('CVE','cve','cve_id') }
        if ($null -eq $cveVal) { $cveVal = Get-ConnectSecureVulnFieldDeep -Item $item -FieldNames @('CVE','cve','cve_id') }
        $descVal = Get-ConnectSecureVulnField -Item $src -Paths @('Problem Name','problem_name','description','summary','_source.Problem Name','_source.problem_name','_source.description')
        if ($null -eq $descVal) { $descVal = Get-ConnectSecureVulnField -Item $item -Paths @('Problem Name','problem_name','description') }
        if ($null -eq $descVal) { $descVal = Get-ConnectSecureVulnFieldDeep -Item $item -FieldNames @('Problem Name','problem_name','description','summary') }
        }
        $emptyStr = [string]::Empty
        $excelRow = [PSCustomObject]@{
            'Host Name' = if ($hostVal) { $hostVal } else { $emptyStr }
            'IP' = if ($ipVal) { $ipVal } else { $emptyStr }
            'Product' = if ($prodVal) { $prodVal } else { $emptyStr }
            'Severity' = if ($sevVal) { $sevVal } else { $emptyStr }
            'EPSS Score' = if ($null -ne $epssVal) { [double]$epssVal } else { 0.0 }
            'Fix' = if ($fixVal) { $fixVal } else { $emptyStr }
            'Evidence Path' = if ($evPathVal) { $evPathVal } else { $emptyStr }
            'Evidence Version' = if ($evVerVal) { $evVerVal } else { $emptyStr }
            'CVE' = if ($cveVal) { $cveVal } else { $emptyStr }
            'Description' = if ($descVal) { $descVal } else { $emptyStr }
        }
        
        $excelData += $excelRow
    }

    Write-CSApiLog ('Converted ' + $excelData.Count + ' records to Excel format') -Level Success
    return $excelData
}

function Export-ConnectSecureDataToExcel {
    param(
        [array]$Data,
        [string]$OutputPath,
        [string]$SheetName = 'Vulnerabilities',
        [scriptblock]$OnProgress = $null
    )

    Write-CSApiLog ('Exporting data to Excel: ' + $OutputPath) -Level Info

    $excel = $null
    $workbook = $null

    try {
        if ($OnProgress) { & $OnProgress 'Opening Excel...' }
        $excel = New-Object -ComObject Excel.Application
        $excel.Visible = $false
        $excel.DisplayAlerts = $false
        $excel.ScreenUpdating = $false

        $workbook = $excel.Workbooks.Add()
        $worksheet = $workbook.Worksheets.Item(1)
        $worksheet.Name = $SheetName

        # Derive headers from API data (union of all rows - API output is source of truth)
        $headers = @()
        if ($Data -and $Data.Count -gt 0) {
            $allKeys = [System.Collections.Generic.HashSet[string]]::new()
            foreach ($item in $Data) {
                $obj = if ($null -ne $item -and $item.PSObject.Properties['_source']) { $item._source } else { $item }
                if ($null -ne $obj) {
                    if ($obj -is [hashtable]) {
                        foreach ($k in $obj.Keys) { $null = $allKeys.Add([string]$k) }
                    } else {
                        $obj.PSObject.Properties.Name | Where-Object { $_ -ne '_source' } | ForEach-Object { $null = $allKeys.Add($_) }
                    }
                }
            }
            $headers = [string[]]$allKeys
        }

        for ($col = 1; $col -le $headers.Count; $col++) {
            $r = $worksheet.Cells.Item(1, $col)
            $r.Value2 = [string]$headers[$col - 1]
            $r.Font.Bold = $true
        }

        # Write data row by row - API output as source of truth
        $row = 2
        $total = if ($Data) { $Data.Count } else { 0 }
        $progressInterval = if ($total -gt 0) { [Math]::Max(100, [Math]::Floor($total / 20)) } else { 99999 }
        $written = 0
        foreach ($item in $Data) {
            $obj = if ($null -ne $item -and $item.PSObject.Properties['_source']) { $item._source } else { $item }
            for ($col = 1; $col -le $headers.Count; $col++) {
                $key = $headers[$col - 1]
                $val = $null
                if ($null -ne $obj) {
                    if ($obj -is [hashtable] -and $obj.ContainsKey($key)) { $val = $obj[$key] }
                    elseif ($obj.PSObject.Properties[$key]) { $val = $obj.PSObject.Properties[$key].Value }
                }
                if ($null -eq $val) { $val = '' }
                elseif ($val -is [Array] -or ($val -is [System.Collections.IEnumerable] -and $val -isnot [string])) {
                    $arr = @($val) | Where-Object { $null -ne $_ }
                    $val = ($arr | ForEach-Object { $_.ToString() }) -join '; '
                }
                elseif ($val -is [PSCustomObject] -or $val -is [hashtable]) {
                    try { $val = ($val | ConvertTo-Json -Compress -Depth 2) } catch { $val = $val.ToString() }
                }
                elseif ($val -isnot [string]) { $val = $val.ToString() }
                $cellVal = [string]$val
                $targetCell = $worksheet.Cells.Item([int]$row, [int]$col)
                $targetCell.Value2 = $cellVal
            }
            $row++
            $written++
            if ($OnProgress -and $total -gt 0 -and $progressInterval -gt 0 -and ($written % $progressInterval -eq 0)) {
                $pct = [Math]::Min(99, [Math]::Floor(100 * $written / $total))
                $msg = [string]('Exporting to Excel... ' + $written + ' of ' + $total + ' rows (' + $pct + ' pct)')
                & $OnProgress $msg
            }
        }

        # Auto-fit columns
        if ($OnProgress) { & $OnProgress 'Formatting columns...' }
        $usedRange = $worksheet.UsedRange
        $usedRange.Columns.AutoFit() | Out-Null

        # Save workbook
        if ($OnProgress) { & $OnProgress 'Saving Excel file...' }
        if (Test-Path $OutputPath) {
            Remove-Item $OutputPath -Force
        }
        $workbook.SaveAs($OutputPath)
        $workbook.Close($false)

        Write-CSApiLog ('Excel file saved: ' + $OutputPath) -Level Success

    } catch {
        Write-CSApiLog ('Error exporting to Excel: ' + $_.Exception.Message) -Level Error
        throw
    } finally {
        if ($workbook) {
            try {
                $workbook.Close($false)
                Clear-ComObject $workbook
            } catch { }
        }
        if ($excel) {
            try {
                $excel.Quit()
                Clear-ComObject $excel
            } catch { }
        }
        [System.GC]::Collect()
        [System.GC]::WaitForPendingFinalizers()
    }
}

function Export-ConnectSecureMultiSheetToExcel {
    <#
    .SYNOPSIS
    Exports multiple sheets to one Excel workbook. Each sheet uses 13-column format.
    .PARAMETER Sheets
    Array of @{ SheetName = '...'; Data = @(...) } - Data is array of PSCustomObjects with 13-column properties.
    #>
    param(
        [array]$Sheets,
        [string]$OutputPath,
        [scriptblock]$OnProgress = $null
    )
    if (-not $Sheets -or $Sheets.Count -eq 0) { throw 'Sheets parameter must have at least one sheet.' }
    $excel = $null
    $workbook = $null
    try {
        if ($OnProgress) { & $OnProgress 'Opening Excel...' }
        $excel = New-Object -ComObject Excel.Application
        $excel.Visible = $false
        $excel.DisplayAlerts = $false
        $excel.ScreenUpdating = $false
        $workbook = $excel.Workbooks.Add()
        $sheetIndex = 0
        foreach ($sh in $Sheets) {
            $sheetName = $sh.SheetName
            $data = $sh.Data
            if (-not $data) { $data = @() }
            $ws = if ($sheetIndex -eq 0) { $workbook.Worksheets.Item(1) } else { $workbook.Worksheets.Add([System.Reflection.Missing]::Value, $workbook.Worksheets.Item($workbook.Worksheets.Count)) }
            $ws.Name = $sheetName
            $headers = @()
            if ($data -and $data.Count -gt 0) {
                $allKeys = [System.Collections.Generic.HashSet[string]]::new()
                foreach ($item in $data) {
                    $obj = if ($null -ne $item -and $item.PSObject.Properties['_source']) { $item._source } else { $item }
                    if ($null -ne $obj -and $obj.PSObject.Properties) {
                        $obj.PSObject.Properties.Name | Where-Object { $_ -ne '_source' } | ForEach-Object { $null = $allKeys.Add($_) }
                    }
                }
                $headers = [string[]]$allKeys
            }
            if ($headers.Count -eq 0 -and $data.Count -eq 0) {
                $emptyRow = [PSCustomObject]@{ 'CVE ID'=''; 'Severity'=''; 'CVSS Score'=''; 'EPSS Score'=''; 'Asset Name'=''; 'OS'=''; 'IP Address'=''; 'Description'=''; 'Application Name'=''; 'Product Name'=''; 'First Seen'=''; 'Last Seen'=''; 'Owner'=''; 'Fix'='' }
                $headers = $emptyRow.PSObject.Properties.Name
                $data = @($emptyRow)
            }
            for ($col = 1; $col -le $headers.Count; $col++) {
                $ws.Cells.Item(1, $col).Value2 = [string]$headers[$col - 1]
                $ws.Cells.Item(1, $col).Font.Bold = $true
            }
            $row = 2
            foreach ($item in $data) {
                $obj = if ($null -ne $item -and $item.PSObject.Properties['_source']) { $item._source } else { $item }
                for ($col = 1; $col -le $headers.Count; $col++) {
                    $key = $headers[$col - 1]
                    $val = if ($null -ne $obj -and $obj.PSObject.Properties[$key]) { $obj.PSObject.Properties[$key].Value } else { '' }
                    if ($null -eq $val) { $val = '' }
                    elseif ($val -is [Array] -or ($val -is [System.Collections.IEnumerable] -and $val -isnot [string])) {
                        $arr = @($val) | Where-Object { $null -ne $_ }
                        $val = ($arr | ForEach-Object { $_.ToString() }) -join '; '
                    }
                    elseif ($val -is [PSCustomObject] -or $val -is [hashtable]) {
                        try { $val = ($val | ConvertTo-Json -Compress -Depth 2) } catch { $val = $val.ToString() }
                    }
                    elseif ($val -isnot [string]) { $val = $val.ToString() }
                    $ws.Cells.Item($row, $col).Value2 = [string]$val
                }
                $row++
            }
            if ($data.Count -gt 0) {
                $usedRange = $ws.UsedRange
                $usedRange.Columns.AutoFit() | Out-Null
                Clear-ComObject $usedRange
            }
            Clear-ComObject $ws
            $sheetIndex++
        }
        if ($OnProgress) { & $OnProgress 'Saving Excel file...' }
        if (Test-Path $OutputPath) { Remove-Item $OutputPath -Force }
        $workbook.SaveAs($OutputPath)
        $workbook.Close($false)
        Write-CSApiLog ('Multi-sheet Excel saved: ' + $OutputPath) -Level Success
    } catch {
        Write-CSApiLog ('Error exporting multi-sheet Excel: ' + $_.Exception.Message) -Level Error
        throw
    } finally {
        if ($workbook) { try { $workbook.Close($false); Clear-ComObject $workbook } catch { } }
        if ($excel) { try { $excel.Quit(); Clear-ComObject $excel } catch { } }
        [System.GC]::Collect()
        [System.GC]::WaitForPendingFinalizers()
    }
}

# --- Report Generation from ConnectSecure Data (no API server needed) ---

function Get-ConnectSecureVulnField {
    param([object]$Item, [string[]]$Paths)

    # Helper: extract value from object (handle .keyword sub-property, arrays like [ios])
    $extractVal = {
        param($o)
        if ($null -eq $o -or $o -eq '') { return $null }
        if ($o -is [string] -or $o -is [int] -or $o -is [long] -or $o -is [double] -or $o -is [bool]) { return $o }
        if ($o -is [Array]) {
            $first = $o | Where-Object { $null -ne $_ -and $_ -ne '' } | Select-Object -First 1
            if ($first) {
                if ($first -is [string] -or $first -is [int] -or $first -is [long] -or $first -is [double]) { return $first }
                try { if ($first.PSObject.Properties['keyword']) { return $first.keyword } } catch { }
                return $first.ToString()
            }
            return $null
        }
        try { if ($o.PSObject.Properties['keyword']) { return $o.keyword } } catch { }
        return $o.ToString()
    }

    foreach ($p in $Paths) {
        try {
            $o = $Item
            # Try literal property name first (e.g. severity.keyword as single prop)
            if ($p.Contains('.')) {
                $litVal = $null
                try { $litVal = $o.PSObject.Properties[$p].Value } catch { }
                if ($null -ne $litVal -and $litVal -ne '') {
                    $v = & $extractVal $litVal
                    if ($null -ne $v) { return $v }
                }
            }
            # Traverse path
            $parts = $p -split '\.'
            foreach ($part in $parts) {
                if ($null -eq $o) { break }
                try { $o = $o.$part } catch { $o = $null; break }
            }
            if ($null -ne $o -and $o -ne '') {
                $v = & $extractVal $o
                if ($null -ne $v) { return $v }
            }
        } catch { }
    }
    return $null
}

function Get-ConnectSecureVulnFieldDeep {
    <# Recursively search object tree for first match of any known field name. Use when path-based lookup fails. #>
    param([object]$Item, [string[]]$FieldNames, [int]$MaxDepth = 5)
    if ($null -eq $Item -or $MaxDepth -le 0) { return $null }
    try {
        foreach ($prop in $Item.PSObject.Properties) {
            $name = $prop.Name
            $val = $prop.Value
            if ($name -in $FieldNames -and $null -ne $val -and $val -ne '') {
                if ($val -is [string] -or $val -is [int] -or $val -is [long] -or $val -is [double]) { return $val }
                if ($val -is [Array]) {
                    $first = $val | Where-Object { $null -ne $_ -and $_ -ne '' } | Select-Object -First 1
                    if ($first -is [string]) { return $first }
                    if ($null -ne $first) { return $first.ToString() }
                    return $null
                }
                try { if ($val.PSObject.Properties['keyword']) { return $val.keyword } } catch { }
                return $val.ToString()
            }
        }
        foreach ($prop in $Item.PSObject.Properties) {
            $val = $prop.Value
            if ($null -ne $val -and $val -is [System.Management.Automation.PSCustomObject]) {
                $found = Get-ConnectSecureVulnFieldDeep -Item $val -FieldNames $FieldNames -MaxDepth ($MaxDepth - 1)
                if ($null -ne $found) { return $found }
            }
        }
    } catch { }
    return $null
}

function Convert-ConnectSecureToVulnData {
    param([array]$ConnectSecureData)

    # Align with ConnectSecure native report: Host Name, FQDN Name, IP, Software Name, CVE, Problem Name, Solution, Severity
    $hostPaths = @('Host Name','host_name','hostName','hostname','fqdn_name','fqdn','host.hostname','host.host_name','asset_name','computer_name','device_name','asset.hostname','asset.host_name','machine.hostname','host.hostname','device.hostname','_source.Host Name','_source.host_name','_source.hostname','_source.fqdn_name','_source.asset.hostname')
    $ipPaths = @('IP','ip','ip_address','ipAddress','host.ip','asset.ip','asset.ip_address','machine.ip','device.ip','_source.IP','_source.ip','_source.ip_address','_source.asset.ip')
    $productPaths = @('Software Name','software_name','softwareName','product','name','cpe','vulnerability.product','vulnerability.software_name','asset.product','_source.Software Name','_source.product','_source.software_name','_source.name')
    $sevPaths = @('Severity','severity','severity.keyword','vulnerability.severity','vulnerability.severity.keyword','_source.Severity','_source.severity','_source.severity.keyword')
    $epssPaths = @('epss_score','epssScore','vulnerability.epss_score','_source.epss_score')

    $vulnData = @()
    foreach ($item in $ConnectSecureData) {
        $src = if ($item._source) { $item._source } else { $item }
        if ($src._aggregated) {
            $hostName = $src.AffectedHosts
            if ($src.HostCount) { $hostName = ('(' + $src.HostCount + ' hosts) ') + $hostName }
            $ip = ""
            $product = $src.software_name
            $severity = $src.severity
            $epssScore = if ($null -ne $src.epss_score) { [double]$src.epss_score } else { 0.0 }
            $critical = if ($severity -eq 'Critical') { 1 } else { 0 }
            $high = if ($severity -eq 'High') { 1 } else { 0 }
            $medium = if ($severity -eq 'Medium') { 1 } else { 0 }
            $low = if ($severity -eq 'Low') { 1 } else { 0 }
            $vulnData += [PSCustomObject]@{ 'Host Name' = $hostName; 'IP' = $ip; 'Username' = ''; 'Product' = $product; 'Critical' = $critical; 'High' = $high; 'Medium' = $medium; 'Low' = $low; 'Vulnerability Count' = [Math]::Max(1, $src.HostCount); 'EPSS Score' = $epssScore }
            continue
        }
        $hostName = (Get-ConnectSecureVulnField -Item $src -Paths $hostPaths)
        if ($null -eq $hostName) { $hostName = (Get-ConnectSecureVulnField -Item $item -Paths $hostPaths) }
        if ($null -eq $hostName) { $hostName = (Get-ConnectSecureVulnFieldDeep -Item $item -FieldNames @('host_name','hostName','hostname','asset_name','computer_name','device_name')) }
        if ($null -eq $hostName) { $hostName = "" }
        $ip = (Get-ConnectSecureVulnField -Item $src -Paths $ipPaths)
        if ($null -eq $ip) { $ip = (Get-ConnectSecureVulnField -Item $item -Paths $ipPaths) }
        if ($null -eq $ip) { $ip = (Get-ConnectSecureVulnFieldDeep -Item $item -FieldNames @('ip','ip_address','ipAddress')) }
        if ($null -eq $ip) { $ip = "" }
        $product = (Get-ConnectSecureVulnField -Item $src -Paths $productPaths)
        if ($null -eq $product) { $product = (Get-ConnectSecureVulnField -Item $item -Paths $productPaths) }
        if ($null -eq $product) { $product = (Get-ConnectSecureVulnFieldDeep -Item $item -FieldNames @('product','software_name','softwareName','name')) }
        if ($null -eq $product) { $product = "" }
        $sevRaw = (Get-ConnectSecureVulnField -Item $src -Paths $sevPaths)
        if ($null -eq $sevRaw) { $sevRaw = (Get-ConnectSecureVulnField -Item $item -Paths $sevPaths) }
        if ($null -eq $sevRaw) { $sevRaw = (Get-ConnectSecureVulnFieldDeep -Item $item -FieldNames @('severity')) }
        $severity = if ($sevRaw) { if ($sevRaw -is [string]) { $sevRaw } else { $sevRaw.keyword } } else { 'Medium' }
        $epssRaw = (Get-ConnectSecureVulnField -Item $src -Paths $epssPaths)
        if ($null -eq $epssRaw) { $epssRaw = (Get-ConnectSecureVulnField -Item $item -Paths $epssPaths) }
        $epssScore = if ($null -ne $epssRaw) { [double]$epssRaw } else { 0.0 }
        $critical = if ($severity -eq 'Critical') { 1 } else { 0 }
        $high = if ($severity -eq 'High') { 1 } else { 0 }
        $medium = if ($severity -eq 'Medium') { 1 } else { 0 }
        $low = if ($severity -eq 'Low') { 1 } else { 0 }
        $vulnData += [PSCustomObject]@{
            'Host Name' = $hostName
            'IP' = $ip
            'Username' = ''
            'Product' = $product
            'Critical' = $critical
            'High' = $high
            'Medium' = $medium
            'Low' = $low
            'Vulnerability Count' = 1
            'EPSS Score' = $epssScore
        }
    }
    return $vulnData
}

