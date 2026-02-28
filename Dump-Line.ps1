param([int]$LineNum = 990)
$path = "c:\Git\vscanmagicv3\ConnectSecure-API.ps1"
$content = [System.IO.File]::ReadAllText($path, [System.Text.Encoding]::UTF8)
$lines = $content -split "`r?`n"
$line = $lines[$LineNum - 1]
Write-Host "Line $LineNum : $line"
Write-Host "Hex:"
[byte[]][char[]]$line | ForEach-Object { Write-Host -NoNewline ("{0:X2} " -f $_) }
Write-Host ""
# Check for non-ASCII
$line.ToCharArray() | ForEach-Object { $c = [int]$_; if ($c -gt 127) { Write-Host "Non-ASCII at pos: $c (U+$($c.ToString('X4')))" } }
