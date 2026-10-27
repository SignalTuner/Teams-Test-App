# Reuses the existing SignalTuner logo; no image generation or package creation.
$ErrorActionPreference = "Stop"
Add-Type -AssemblyName System.Drawing
$root = Split-Path $PSScriptRoot -Parent
$source = [System.Drawing.Bitmap]::FromFile((Join-Path $root "src/Tab/assets/signaltuner-logo-horizontal.png"))
try {
    # The symbol occupies the first square of the horizontal logo. Exclude text.
    $left = $source.Height; $top = $source.Height; $right = 0; $bottom = 0
    for ($y = 0; $y -lt $source.Height; $y++) {
        for ($x = 0; $x -lt $source.Height; $x++) {
            if ($source.GetPixel($x, $y).A -gt 0) {
                $left = [Math]::Min($left, $x); $right = [Math]::Max($right, $x)
                $top = [Math]::Min($top, $y); $bottom = [Math]::Max($bottom, $y)
            }
        }
    }
    $bounds = New-Object System.Drawing.Rectangle($left, $top, ($right - $left + 1), ($bottom - $top + 1))
    foreach ($variant in @(@{ File = "color.png"; Size = 192; Safe = 120 }, @{ File = "outline.png"; Size = 32; Safe = 28 })) {
        $bitmap = New-Object System.Drawing.Bitmap($variant.Size, $variant.Size)
        $graphics = [System.Drawing.Graphics]::FromImage($bitmap)
        try {
            if ($variant.File -eq "color.png") { $graphics.Clear([System.Drawing.Color]::White) }
            else { $graphics.Clear([System.Drawing.Color]::Transparent) }
            $graphics.InterpolationMode = [System.Drawing.Drawing2D.InterpolationMode]::HighQualityBicubic
            $graphics.PixelOffsetMode = [System.Drawing.Drawing2D.PixelOffsetMode]::HighQuality
            $scale = $variant.Safe / [double][Math]::Max($bounds.Width, $bounds.Height)
            $width = [int][Math]::Round($bounds.Width * $scale)
            $height = [int][Math]::Round($bounds.Height * $scale)
            $target = New-Object System.Drawing.Rectangle(([int](($variant.Size - $width) / 2)), ([int](($variant.Size - $height) / 2)), $width, $height)
            $graphics.DrawImage($source, $target, $bounds, [System.Drawing.GraphicsUnit]::Pixel)
            if ($variant.File -eq "outline.png") {
                for ($y = 0; $y -lt $bitmap.Height; $y++) {
                    for ($x = 0; $x -lt $bitmap.Width; $x++) {
                        $bitmap.SetPixel($x, $y, [System.Drawing.Color]::FromArgb($bitmap.GetPixel($x, $y).A, 255, 255, 255))
                    }
                }
            }
            $bitmap.Save((Join-Path $root "appPackage/$($variant.File)"), [System.Drawing.Imaging.ImageFormat]::Png)
        } finally { $graphics.Dispose(); $bitmap.Dispose() }
    }
} finally { $source.Dispose() }
Write-Host "Updated SignalTuner color (192px) and white/transparent outline (32px) icons from the existing logo."
