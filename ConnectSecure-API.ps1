#Requires -Version 5.1
<#
.SYNOPSIS
ConnectSecure API Client Module

.DESCRIPTION
Provides functions to interact with the ConnectSecure API v4 for fetching vulnerability data.

Implementation is split into dot-sourced parts for maintainability:
- ConnectSecure-API.Part1.ps1 — configuration, Invoke-ConnectSecureRequest, auth, vulnerability queries, company review endpoint map
- ConnectSecure-API.Part2.ps1 — company review data, asset resolution, Excel conversion helpers
- ConnectSecure-API.Part3.ps1 — report generation, report builder jobs, batch download, local fallback

.NOTES
Version: 1.0.0
Author: River Run MSP
#>

Add-Type -AssemblyName System.Web

$csApiRoot = $PSScriptRoot
if ([string]::IsNullOrEmpty($csApiRoot)) {
    $csApiRoot = Split-Path -Parent $MyInvocation.MyCommand.Path
}

. (Join-Path $csApiRoot 'ConnectSecure-API.Part1.ps1')
. (Join-Path $csApiRoot 'ConnectSecure-API.Part2.ps1')
. (Join-Path $csApiRoot 'ConnectSecure-API.Part3.ps1')
