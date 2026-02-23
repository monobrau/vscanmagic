# Test loading first N lines
$path = "c:\Git\vscanmagicv3\ConnectSecure-API.ps1"
$lines = Get-Content $path
foreach ($max in 440, 441, 442, 443, 444, 445, 450, 455, 460) {
    $part = $lines[0..($max-1)] + "}"
    $partStr = $part -join "`n"
    $tokens = $null
    $errs = $null
    $null = [System.Management.Automation.Language.Parser]::ParseInput($partStr, [ref]$tokens, [ref]$errs)
    if ($errs.Count -eq 0) { Write-Host "Lines 1-$max : OK" } else { Write-Host "Lines 1-$max : FAIL - $($errs[0].Message)" }
}
