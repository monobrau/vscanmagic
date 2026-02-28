# Replace smart/curly quotes with ASCII quotes
$path = "c:\Git\vscanmagicv3\ConnectSecure-API.ps1"
$content = [System.IO.File]::ReadAllText($path, [System.Text.Encoding]::UTF8)
$before = $content
$content = $content.Replace([char]0x2018, [char]0x27)  # ' U+2018
$content = $content.Replace([char]0x2019, [char]0x27)  # ' U+2019
$content = $content.Replace([char]0x201C, [char]0x22)  # " U+201C
$content = $content.Replace([char]0x201D, [char]0x22)  # " U+201D
$content = $content.Replace([char]0x92, [char]0x27)    # Windows-1252 right single quote
$content = $content.Replace([char]0x91, [char]0x27)   # Windows-1252 left single quote
$content = $content.Replace([char]0x93, [char]0x22)    # Windows-1252 left double quote
$content = $content.Replace([char]0x94, [char]0x22)    # Windows-1252 right double quote
$utf8NoBom = [System.Text.UTF8Encoding]::new($false)
[System.IO.File]::WriteAllText($path, $content, $utf8NoBom)
$count = ([int[]][char[]]$before | Where-Object { $_ -in 0x91,0x92,0x93,0x94,0x2018,0x2019,0x201C,0x201D }).Count
Write-Host "Fixed smart quotes in ConnectSecure-API.ps1 ($count replacements)"
