# Create Resources directory if it doesn't exist
$resourcesDir = Join-Path $PSScriptRoot "Resources"
if (-not (Test-Path $resourcesDir)) {
    New-Item -ItemType Directory -Path $resourcesDir | Out-Null
}

# Create a simple PowerPoint-themed icon using PowerShell
Add-Type -AssemblyName System.Drawing

$size = 32
$bitmap = New-Object System.Drawing.Bitmap($size, $size)
$g = [System.Drawing.Graphics]::FromImage($bitmap)

# Set high quality rendering
$g.SmoothingMode = [System.Drawing.Drawing2D.SmoothingMode]::AntiAlias
$g.InterpolationMode = [System.Drawing.Drawing2D.InterpolationMode]::HighQualityBicubic
$g.PixelOffsetMode = [System.Drawing.Drawing2D.PixelOffsetMode]::HighQuality

# Set background color (PowerPoint orange)
$backgroundColor = [System.Drawing.Color]::FromArgb(255, 209, 71, 0)
$g.Clear($backgroundColor)

# Draw "P" in white
$font = New-Object System.Drawing.Font("Arial", [float]24, [System.Drawing.FontStyle]::Bold)
$brush = [System.Drawing.Brushes]::White
$format = New-Object System.Drawing.StringFormat
$format.Alignment = [System.Drawing.StringAlignment]::Center
$format.LineAlignment = [System.Drawing.StringAlignment]::Center
$rect = New-Object System.Drawing.RectangleF(0, 0, $size, $size)
$g.DrawString("P", $font, $brush, $rect, $format)

# Convert to icon and save
$iconPath = Join-Path $resourcesDir "app.ico"
$icon = [System.Drawing.Icon]::FromHandle($bitmap.GetHicon())

# Save the icon
$fileStream = [System.IO.File]::Create($iconPath)
$icon.Save($fileStream)
$fileStream.Close()

# Clean up
$icon.Dispose()
$bitmap.Dispose()
$g.Dispose()
$font.Dispose()

Write-Host "Icon generated at: $iconPath"
