$path = Join-Path $PSScriptRoot 'ConnectSecure-API.ps1'
$content = [System.IO.File]::ReadAllText($path)
# Replace Unicode smart/curly double quotes with ASCII
$content = $content -replace [char]0x201C, [char]0x22  # "
$content = $content -replace [char]0x201D, [char]0x22  # "
$content = $content -replace [char]0x2018, [char]0x27  # '
$content = $content -replace [char]0x2019, [char]0x27  # '
[System.IO.File]::WriteAllText($path, $content)
Write-Host 'Normalized quotes in ConnectSecure-API.ps1'
