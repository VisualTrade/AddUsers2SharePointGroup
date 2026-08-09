<#
.SYNOPSIS
    Generates the 32x32 ribbon icons for the AddUsers Outlook add-in.

.DESCRIPTION
    Renders each icon at 4x resolution with GDI+ anti-aliasing, then downscales
    to a crisp 32x32 PNG with a transparent background. Idempotent: overwrites
    AddUsers\Resources\AddUsers32.png and AddUsers\Resources\Configure32.png.

    Windows PowerShell 5 compatible. Requires only System.Drawing.
#>

Set-StrictMode -Version 2.0
$ErrorActionPreference = 'Stop'

Add-Type -AssemblyName System.Drawing

$repoRoot    = Split-Path -Parent $PSScriptRoot
$resourceDir = Join-Path $repoRoot 'AddUsers\Resources'
if (-not (Test-Path -LiteralPath $resourceDir)) {
    New-Item -ItemType Directory -Path $resourceDir -Force | Out-Null
}

$IconSize    = 32
$SuperSample = 4

function New-Canvas {
    # Returns a supersampled bitmap; caller draws in 32-unit world coordinates.
    $px  = $script:IconSize * $script:SuperSample
    $bmp = New-Object System.Drawing.Bitmap($px, $px, [System.Drawing.Imaging.PixelFormat]::Format32bppArgb)
    return $bmp
}

function New-CanvasGraphics {
    param([System.Drawing.Bitmap]$Bitmap)
    $g = [System.Drawing.Graphics]::FromImage($Bitmap)
    $g.SmoothingMode      = [System.Drawing.Drawing2D.SmoothingMode]::AntiAlias
    $g.PixelOffsetMode    = [System.Drawing.Drawing2D.PixelOffsetMode]::HighQuality
    $g.CompositingQuality = [System.Drawing.Drawing2D.CompositingQuality]::HighQuality
    $g.Clear([System.Drawing.Color]::Transparent)
    $g.ScaleTransform($script:SuperSample, $script:SuperSample)
    return $g
}

function New-VerticalBrush {
    param([double]$X, [double]$Y, [double]$W, [double]$H,
          [System.Drawing.Color]$Top, [System.Drawing.Color]$Bottom)
    # Pad the gradient rect so anti-aliased edge pixels never sample the wrap seam.
    $rect  = New-Object System.Drawing.RectangleF(($X - 1), ($Y - 1), ($W + 2), ($H + 2))
    $brush = New-Object System.Drawing.Drawing2D.LinearGradientBrush(
        $rect, $Top, $Bottom, [System.Drawing.Drawing2D.LinearGradientMode]::Vertical)
    $brush.WrapMode = [System.Drawing.Drawing2D.WrapMode]::TileFlipXY
    return $brush
}

function Save-Png32 {
    param([System.Drawing.Bitmap]$Source, [string]$Path)
    $dst = New-Object System.Drawing.Bitmap($script:IconSize, $script:IconSize, [System.Drawing.Imaging.PixelFormat]::Format32bppArgb)
    $g   = [System.Drawing.Graphics]::FromImage($dst)
    try {
        $g.CompositingMode    = [System.Drawing.Drawing2D.CompositingMode]::SourceCopy
        $g.CompositingQuality = [System.Drawing.Drawing2D.CompositingQuality]::HighQuality
        $g.InterpolationMode  = [System.Drawing.Drawing2D.InterpolationMode]::HighQualityBicubic
        $g.SmoothingMode      = [System.Drawing.Drawing2D.SmoothingMode]::AntiAlias
        $g.PixelOffsetMode    = [System.Drawing.Drawing2D.PixelOffsetMode]::HighQuality

        $attr = New-Object System.Drawing.Imaging.ImageAttributes
        $attr.SetWrapMode([System.Drawing.Drawing2D.WrapMode]::TileFlipXY)
        $destRect = New-Object System.Drawing.Rectangle(0, 0, $script:IconSize, $script:IconSize)
        $g.DrawImage($Source, $destRect, 0, 0, $Source.Width, $Source.Height,
                     [System.Drawing.GraphicsUnit]::Pixel, $attr)
        $attr.Dispose()
    }
    finally {
        $g.Dispose()
    }

    if (Test-Path -LiteralPath $Path) {
        Remove-Item -LiteralPath $Path -Force
    }
    $dst.Save($Path, [System.Drawing.Imaging.ImageFormat]::Png)
    $dst.Dispose()
}

function New-AddUsersBitmap {
    # Two overlapping person silhouettes (slate blues) with a bold green plus
    # badge at bottom-right.
    $bmp = New-Canvas
    $g   = New-CanvasGraphics -Bitmap $bmp
    try {
        # --- Geometry (32-unit space) ------------------------------------
        # Back person (upper right, lighter)
        $bHeadCx = 20.5; $bHeadCy = 9.0;  $bHeadR = 4.4
        $bBodyCx = 20.5; $bBodyCy = 21.0; $bBodyRx = 7.0; $bBodyRy = 8.0; $bBase = 26.0
        # Front person (lower left, darker)
        $fHeadCx = 12.0; $fHeadCy = 10.5; $fHeadR = 5.0
        $fBodyCx = 12.0; $fBodyCy = 26.2; $fBodyRx = 8.0; $fBodyRy = 9.4; $fBase = 30.0
        $gap = 1.0   # separation halo carved out of the back silhouette

        # Front silhouette expanded by the gap width; excluded while painting
        # the back person so the front figure reads cleanly against it.
        $frontHalo = New-Object System.Drawing.Drawing2D.GraphicsPath
        $frontHalo.AddEllipse(($fHeadCx - $fHeadR - $gap), ($fHeadCy - $fHeadR - $gap),
                              (2 * ($fHeadR + $gap)), (2 * ($fHeadR + $gap)))
        $frontHalo.AddEllipse(($fBodyCx - $fBodyRx - $gap), ($fBodyCy - $fBodyRy - $gap),
                              (2 * ($fBodyRx + $gap)), (2 * ($fBodyRy + $gap)))

        # --- Back person -------------------------------------------------
        $backBrush = New-VerticalBrush -X ($bBodyCx - $bBodyRx) -Y ($bHeadCy - $bHeadR) `
            -W (2 * $bBodyRx) -H ($bBase - ($bHeadCy - $bHeadR)) `
            -Top ([System.Drawing.Color]::FromArgb(255, 0x93, 0xA9, 0xC2)) `
            -Bottom ([System.Drawing.Color]::FromArgb(255, 0x7B, 0x93, 0xAE))

        $backClip = New-Object System.Drawing.RectangleF(0, 0, 32, $bBase)
        $g.SetClip($backClip)
        $g.SetClip($frontHalo, [System.Drawing.Drawing2D.CombineMode]::Exclude)
        $g.FillEllipse($backBrush, ($bHeadCx - $bHeadR), ($bHeadCy - $bHeadR), (2 * $bHeadR), (2 * $bHeadR))
        $g.FillEllipse($backBrush, ($bBodyCx - $bBodyRx), ($bBodyCy - $bBodyRy), (2 * $bBodyRx), (2 * $bBodyRy))
        $g.ResetClip()
        $backBrush.Dispose()
        $frontHalo.Dispose()

        # --- Front person ------------------------------------------------
        $frontBrush = New-VerticalBrush -X ($fBodyCx - $fBodyRx) -Y ($fHeadCy - $fHeadR) `
            -W (2 * $fBodyRx) -H ($fBase - ($fHeadCy - $fHeadR)) `
            -Top ([System.Drawing.Color]::FromArgb(255, 0x4E, 0x6C, 0x90)) `
            -Bottom ([System.Drawing.Color]::FromArgb(255, 0x33, 0x50, 0x6E))

        $g.FillEllipse($frontBrush, ($fHeadCx - $fHeadR), ($fHeadCy - $fHeadR), (2 * $fHeadR), (2 * $fHeadR))
        $frontClip = New-Object System.Drawing.RectangleF(0, 0, 32, $fBase)
        $g.SetClip($frontClip)
        $g.FillEllipse($frontBrush, ($fBodyCx - $fBodyRx), ($fBodyCy - $fBodyRy), (2 * $fBodyRx), (2 * $fBodyRy))
        $g.ResetClip()
        $frontBrush.Dispose()

        # --- Green plus badge (bottom-right) -----------------------------
        $bcx = 24.3; $bcy = 24.3; $br = 6.9
        $badgeBrush = New-VerticalBrush -X ($bcx - $br) -Y ($bcy - $br) -W (2 * $br) -H (2 * $br) `
            -Top ([System.Drawing.Color]::FromArgb(255, 0x35, 0xB9, 0x5E)) `
            -Bottom ([System.Drawing.Color]::FromArgb(255, 0x1E, 0x91, 0x46))
        $g.FillEllipse($badgeBrush, ($bcx - $br), ($bcy - $br), (2 * $br), (2 * $br))
        $badgeBrush.Dispose()

        $ringPen = New-Object System.Drawing.Pen([System.Drawing.Color]::FromArgb(255, 0x15, 0x80, 0x3A), 1.0)
        $rr = $br - 0.5
        $g.DrawEllipse($ringPen, ($bcx - $rr), ($bcy - $rr), (2 * $rr), (2 * $rr))
        $ringPen.Dispose()

        $plusPen = New-Object System.Drawing.Pen([System.Drawing.Color]::White, 2.9)
        $plusPen.StartCap = [System.Drawing.Drawing2D.LineCap]::Round
        $plusPen.EndCap   = [System.Drawing.Drawing2D.LineCap]::Round
        $half = 3.2
        $g.DrawLine($plusPen, ($bcx - $half), $bcy, ($bcx + $half), $bcy)
        $g.DrawLine($plusPen, $bcx, ($bcy - $half), $bcx, ($bcy + $half))
        $plusPen.Dispose()
    }
    finally {
        $g.Dispose()
    }
    return $bmp
}

function New-GearPath {
    param([double]$Cx, [double]$Cy, [double]$OuterR, [double]$RootR, [double]$HoleR, [int]$Teeth)
    $path = New-Object System.Drawing.Drawing2D.GraphicsPath([System.Drawing.Drawing2D.FillMode]::Alternate)
    $pts  = New-Object 'System.Collections.Generic.List[System.Drawing.PointF]'
    $step      = 360.0 / $Teeth
    $toothHalf = 10.0   # half angular width of a tooth tip (degrees)
    $flank     = 4.5    # angular width of each tooth flank (degrees)

    for ($i = 0; $i -lt $Teeth; $i++) {
        $a = $i * $step
        $stops = @(
            @(($a - $toothHalf), $OuterR),
            @(($a + $toothHalf), $OuterR),
            @(($a + $toothHalf + $flank), $RootR),
            @(($a + $step - $toothHalf - $flank), $RootR)
        )
        foreach ($s in $stops) {
            $rad = $s[0] * [Math]::PI / 180.0
            $pts.Add((New-Object System.Drawing.PointF(
                ($Cx + $s[1] * [Math]::Cos($rad)),
                ($Cy + $s[1] * [Math]::Sin($rad)))))
        }
    }
    $path.AddPolygon($pts.ToArray())
    # Center hole; Alternate fill mode turns the inner ellipse into a hole.
    $path.AddEllipse(($Cx - $HoleR), ($Cy - $HoleR), (2 * $HoleR), (2 * $HoleR))
    return $path
}

function New-ConfigureBitmap {
    # Slate gear: 8 teeth, center hole, transparent background.
    $bmp = New-Canvas
    $g   = New-CanvasGraphics -Bitmap $bmp
    try {
        $gear = New-GearPath -Cx 16 -Cy 16 -OuterR 14.0 -RootR 10.7 -HoleR 4.7 -Teeth 8

        $rect  = New-Object System.Drawing.RectangleF(1, 1, 30, 30)
        $brush = New-Object System.Drawing.Drawing2D.LinearGradientBrush(
            $rect,
            [System.Drawing.Color]::FromArgb(255, 0x75, 0x87, 0x9C),
            [System.Drawing.Color]::FromArgb(255, 0x48, 0x59, 0x6C),
            [System.Drawing.Drawing2D.LinearGradientMode]::ForwardDiagonal)
        $brush.WrapMode = [System.Drawing.Drawing2D.WrapMode]::TileFlipXY

        $g.FillPath($brush, $gear)

        # Subtle inner rim around the hole for definition.
        $rimPen = New-Object System.Drawing.Pen([System.Drawing.Color]::FromArgb(90, 0x2E, 0x3B, 0x4A), 1.0)
        $g.DrawEllipse($rimPen, (16 - 5.2), (16 - 5.2), 10.4, 10.4)
        $rimPen.Dispose()

        $brush.Dispose()
        $gear.Dispose()
    }
    finally {
        $g.Dispose()
    }
    return $bmp
}

# --- Generate ------------------------------------------------------------
$icons = @(
    @{ Name = 'AddUsers32.png';  Builder = 'New-AddUsersBitmap'  },
    @{ Name = 'Configure32.png'; Builder = 'New-ConfigureBitmap' }
)

foreach ($icon in $icons) {
    $outPath = Join-Path $resourceDir $icon.Name
    $big = & $icon.Builder
    try {
        Save-Png32 -Source $big -Path $outPath
    }
    finally {
        $big.Dispose()
    }

    $check = [System.Drawing.Image]::FromFile($outPath)
    try {
        Write-Host ("Wrote {0} ({1}x{2})" -f $outPath, $check.Width, $check.Height)
    }
    finally {
        $check.Dispose()
    }
}
