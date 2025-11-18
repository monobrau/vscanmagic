# Build EXE using ps2exe module

$ErrorActionPreference = 'Stop'

# Create Modules directory if needed
$modulesDir = Join-Path $PSScriptRoot "Modules"
if (-not (Test-Path $modulesDir)) {
    New-Item -ItemType Directory -Path $modulesDir | Out-Null
}

# Download module if not present
$modulePath = Join-Path $modulesDir "ps2exe\1.0.17\ps2exe.psm1"
if (-not (Test-Path $modulePath)) {
    Write-Host "Downloading ps2exe module..."
    Save-Module -Name ps2exe -Path $modulesDir -Repository PSGallery -Force
}

# Import the module properly
$modulePath = Join-Path $modulesDir "ps2exe\1.0.17\ps2exe.psm1"
if (Test-Path $modulePath) {
    Write-Host "Loading module from: $modulePath"
    Import-Module $modulePath -Force
    
    # Try to find the function
    $func = Get-Command -Name "Invoke-PS2EXE" -ErrorAction SilentlyContinue
    if (-not $func) {
        $func = Get-Command -Name "Invoke-ps2exe" -ErrorAction SilentlyContinue
    }
    if (-not $func) {
        $func = Get-Command -Name "ps2exe" -ErrorAction SilentlyContinue
    }
    
    if ($func) {
        Write-Host "Found function: $($func.Name)"
        
        # Set paths
        $iconPath = Join-Path $PSScriptRoot "VScanMagic.ico"
        $inputScript = Join-Path $PSScriptRoot "VScanMagic-GUI.ps1"
        $outputExe = Join-Path $PSScriptRoot "VScanMagic.exe"
        
        # Build parameters
        $params = @{
            inputFile = $inputScript
            outputFile = $outputExe
            title = "VScanMagic v3"
            description = "Vulnerability Report Generator"
            company = "River Run MSP"
            product = "VScanMagic"
            version = "3.1.1"
            copyright = "Copyright (c) 2025 Chris Knospe"
        }
        
        if (Test-Path $iconPath) {
            $params.iconFile = $iconPath
            Write-Host "Using icon: $iconPath"
        } else {
            Write-Host "Icon file not found, creating one..."
            # Create icon if it doesn't exist
            Add-Type -AssemblyName System.Drawing
            $bitmap = New-Object System.Drawing.Bitmap(256, 256)
            $graphics = [System.Drawing.Graphics]::FromImage($bitmap)
            $graphics.Clear([System.Drawing.Color]::FromArgb(0, 51, 102))
            $font = New-Object System.Drawing.Font('Arial', 120, [System.Drawing.FontStyle]::Bold)
            $brush = New-Object System.Drawing.SolidBrush([System.Drawing.Color]::White)
            $graphics.DrawString('VS', $font, $brush, 30, 60)
            $graphics.Dispose()
            $bitmap.Save($iconPath, [System.Drawing.Imaging.ImageFormat]::Png)
            
            # Convert PNG to ICO
            $icoFile = [System.IO.File]::Create($iconPath.Replace('.ico', '_temp.ico'))
            $icoWriter = New-Object System.IO.BinaryWriter($icoFile)
            $icoWriter.Write([UInt16]0)
            $icoWriter.Write([UInt16]1)
            $icoWriter.Write([UInt16]1)
            $icoWriter.Write([byte]0)
            $icoWriter.Write([byte]0)
            $icoWriter.Write([byte]0)
            $icoWriter.Write([byte]0)
            $icoWriter.Write([UInt16]1)
            $icoWriter.Write([UInt16]32)
            $pngBytes = [System.IO.File]::ReadAllBytes($iconPath)
            $icoWriter.Write([UInt32]$pngBytes.Length)
            $icoWriter.Write([UInt32]22)
            $icoFile.Position = 22
            $icoWriter.Write($pngBytes)
            $icoWriter.Close()
            $icoFile.Close()
            Remove-Item $iconPath -ErrorAction SilentlyContinue
            Rename-Item $iconPath.Replace('.ico', '_temp.ico') $iconPath -ErrorAction SilentlyContinue
            $bitmap.Dispose()
            Write-Host "Icon created: $iconPath"
            $params.iconFile = $iconPath
        }
        
        Write-Host "Converting script to EXE..."
        & $func.Name @params
        
        if (Test-Path $outputExe) {
            Write-Host "SUCCESS! Created: $outputExe" -ForegroundColor Green
            $fileInfo = Get-Item $outputExe
            Write-Host "File size: $([math]::Round($fileInfo.Length / 1MB, 2)) MB"
            
            # Create ZIP
            Write-Host "Creating ZIP file..."
            $zipPath = Join-Path $PSScriptRoot "VScanMagic.zip"
            if (Test-Path $zipPath) {
                Remove-Item $zipPath -Force
            }
            Compress-Archive -Path $outputExe -DestinationPath $zipPath -Force
            Write-Host "Created ZIP: $zipPath" -ForegroundColor Green
        } else {
            throw "EXE file was not created"
        }
    } else {
        Write-Host "Available commands:" -ForegroundColor Yellow
        Get-Command | Where-Object { $_.Name -like "*ps2exe*" -or $_.Name -like "*PS2EXE*" } | Select-Object Name, CommandType
        throw "ps2exe function not found"
    }
} else {
    throw "Module file not found at: $modulePath"
}

