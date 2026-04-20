param(
    [string]$OutputDir = "c:\Project\Powerpoint\HFP\UnifiedHPC-Docs\Slide_Images"
)

Set-StrictMode -Version Latest
$ErrorActionPreference = "Stop"

Add-Type -AssemblyName System.Drawing

function New-RoundedRectPath {
    param(
        [System.Drawing.RectangleF]$Rect,
        [float]$Radius
    )

    $path = New-Object System.Drawing.Drawing2D.GraphicsPath
    $d = $Radius * 2
    $path.AddArc($Rect.X, $Rect.Y, $d, $d, 180, 90)
    $path.AddArc($Rect.Right - $d, $Rect.Y, $d, $d, 270, 90)
    $path.AddArc($Rect.Right - $d, $Rect.Bottom - $d, $d, $d, 0, 90)
    $path.AddArc($Rect.X, $Rect.Bottom - $d, $d, $d, 90, 90)
    $path.CloseFigure()
    return $path
}

function Draw-RoundedRect {
    param(
        $Graphics,
        [System.Drawing.RectangleF]$Rect,
        [float]$Radius,
        [System.Drawing.Brush]$Fill,
        [System.Drawing.Pen]$Pen
    )

    $path = New-RoundedRectPath -Rect $Rect -Radius $Radius
    $Graphics.FillPath($Fill, $path)
    if ($Pen) {
        $Graphics.DrawPath($Pen, $path)
    }
    $path.Dispose()
}

function Draw-CenteredText {
    param(
        $Graphics,
        [string]$Text,
        [System.Drawing.Font]$Font,
        [System.Drawing.Brush]$Brush,
        [System.Drawing.RectangleF]$Rect,
        [System.Drawing.StringAlignment]$Alignment = [System.Drawing.StringAlignment]::Center,
        [System.Drawing.StringAlignment]$LineAlignment = [System.Drawing.StringAlignment]::Center
    )

    $fmt = New-Object System.Drawing.StringFormat
    $fmt.Alignment = $Alignment
    $fmt.LineAlignment = $LineAlignment
    $fmt.Trimming = [System.Drawing.StringTrimming]::EllipsisWord
    $Graphics.DrawString($Text, $Font, $Brush, $Rect, $fmt)
    $fmt.Dispose()
}

function New-Canvas {
    param(
        [string]$Path,
        [int]$Width,
        [int]$Height,
        [scriptblock]$Draw
    )

    $bmp = New-Object System.Drawing.Bitmap $Width, $Height
    $graphics = [System.Drawing.Graphics]::FromImage($bmp)
    $graphics.SmoothingMode = [System.Drawing.Drawing2D.SmoothingMode]::AntiAlias
    $graphics.TextRenderingHint = [System.Drawing.Text.TextRenderingHint]::ClearTypeGridFit
    & $Draw $graphics $Width $Height
    $bmp.Save($Path, [System.Drawing.Imaging.ImageFormat]::Png)
    $graphics.Dispose()
    $bmp.Dispose()
}

if (-not (Test-Path -LiteralPath $OutputDir)) {
    New-Item -ItemType Directory -Path $OutputDir -Force | Out-Null
}

$titleFont = [System.Drawing.Font]::new('Segoe UI Semibold', 34, [System.Drawing.FontStyle]::Bold, [System.Drawing.GraphicsUnit]::Point)
$subFont = [System.Drawing.Font]::new('Segoe UI', 18, [System.Drawing.FontStyle]::Regular, [System.Drawing.GraphicsUnit]::Point)
$boxFont = [System.Drawing.Font]::new('Segoe UI Semibold', 22, [System.Drawing.FontStyle]::Bold, [System.Drawing.GraphicsUnit]::Point)
$smallFont = [System.Drawing.Font]::new('Segoe UI', 15, [System.Drawing.FontStyle]::Regular, [System.Drawing.GraphicsUnit]::Point)

# Slide 1 image
New-Canvas -Path (Join-Path $OutputDir 'slide1_unified_hpc.png') -Width 1600 -Height 900 -Draw {
    param($g, $w, $h)
    $bg = New-Object System.Drawing.Drawing2D.LinearGradientBrush(
        (New-Object System.Drawing.Rectangle 0, 0, $w, $h),
        [System.Drawing.Color]::FromArgb(12, 23, 38),
        [System.Drawing.Color]::FromArgb(20, 104, 126),
        30
    )
    $g.FillRectangle($bg, 0, 0, $w, $h)
    $bg.Dispose()

    $glow = New-Object System.Drawing.SolidBrush ([System.Drawing.Color]::FromArgb(30, 255, 255, 255))
    $g.FillEllipse($glow, 1120, -120, 500, 500)
    $g.FillEllipse($glow, -180, 590, 420, 420)
    $glow.Dispose()

    $titleRect = New-Object System.Drawing.RectangleF 120, 42, 1360, 60
    $subRect = New-Object System.Drawing.RectangleF 180, 104, 1240, 42
    Draw-CenteredText -Graphics $g -Text 'FDA Unified HPC' -Font $titleFont -Brush ([System.Drawing.Brushes]::White) -Rect $titleRect -LineAlignment ([System.Drawing.StringAlignment]::Center)
    Draw-CenteredText -Graphics $g -Text 'Agency-wide shared service built on existing center infrastructure' -Font $subFont -Brush ([System.Drawing.Brushes]::Gainsboro) -Rect $subRect -LineAlignment ([System.Drawing.StringAlignment]::Center)

    $hubRect = New-Object System.Drawing.RectangleF 540, 330, 520, 180
    $hubFill = New-Object System.Drawing.SolidBrush ([System.Drawing.Color]::FromArgb(245, 249, 255))
    $hubPen = New-Object System.Drawing.Pen([System.Drawing.Color]::FromArgb(130, 202, 229, 255), 5)
    Draw-RoundedRect -Graphics $g -Rect $hubRect -Radius 24 -Fill $hubFill -Pen $hubPen
    $hubFill.Dispose()
    $hubPen.Dispose()

    $hubTitle = [System.Drawing.Font]::new('Segoe UI Semibold', 30, [System.Drawing.FontStyle]::Bold, [System.Drawing.GraphicsUnit]::Point)
    $hubBody = [System.Drawing.Font]::new('Segoe UI', 17, [System.Drawing.FontStyle]::Regular, [System.Drawing.GraphicsUnit]::Point)
    $hubTextBrush = New-Object System.Drawing.SolidBrush ([System.Drawing.Color]::FromArgb(24, 44, 64))
    $hubBodyBrush = New-Object System.Drawing.SolidBrush ([System.Drawing.Color]::FromArgb(63, 80, 100))
    Draw-CenteredText -Graphics $g -Text 'Unified HPC Service Layer' -Font $hubTitle -Brush $hubTextBrush -Rect (New-Object System.Drawing.RectangleF 575, 360, 450, 48)
    Draw-CenteredText -Graphics $g -Text 'Shared access, governance, and support across FDA' -Font $hubBody -Brush $hubBodyBrush -Rect (New-Object System.Drawing.RectangleF 575, 415, 450, 48)
    $hubTitle.Dispose()
    $hubBody.Dispose()
    $hubTextBrush.Dispose()
    $hubBodyBrush.Dispose()

    $nodeTextBrush = New-Object System.Drawing.SolidBrush ([System.Drawing.Color]::FromArgb(23, 43, 63))
    $nodeFill = New-Object System.Drawing.SolidBrush ([System.Drawing.Color]::FromArgb(243, 250, 255))
    $nodePen = New-Object System.Drawing.Pen([System.Drawing.Color]::FromArgb(160, 255, 255, 255), 4)
    $linePen = New-Object System.Drawing.Pen([System.Drawing.Color]::FromArgb(210, 158, 227, 255), 5)
    $nodes = @(
        @{X=120;Y=320;W=290;H=96;Label='CDRH / HIVE'},
        @{X=1190;Y=320;W=290;H=96;Label='CFSAN'},
        @{X=120;Y=560;W=290;H=96;Label='NCTR / Research'},
        @{X=1190;Y=560;W=290;H=96;Label='Cloud Burst'}
    )
    foreach ($n in $nodes) {
        $r = New-Object System.Drawing.RectangleF $n.X, $n.Y, $n.W, $n.H
        Draw-RoundedRect -Graphics $g -Rect $r -Radius 18 -Fill $nodeFill -Pen $nodePen
        Draw-CenteredText -Graphics $g -Text $n.Label -Font $boxFont -Brush $nodeTextBrush -Rect $r
    }
    $nodeFill.Dispose()
    $nodePen.Dispose()
    $nodeTextBrush.Dispose()

    $g.DrawLine($linePen, 410, 368, 540, 382)
    $g.DrawLine($linePen, 1190, 368, 1060, 382)
    $g.DrawLine($linePen, 410, 608, 540, 520)
    $g.DrawLine($linePen, 1190, 608, 1060, 520)
    $linePen.Dispose()
}

# Slide 4 image
New-Canvas -Path (Join-Path $OutputDir 'slide4_siloed_hpc.png') -Width 1600 -Height 900 -Draw {
    param($g, $w, $h)
    $bg = New-Object System.Drawing.Drawing2D.LinearGradientBrush(
        (New-Object System.Drawing.Rectangle 0, 0, $w, $h),
        [System.Drawing.Color]::FromArgb(18, 25, 39),
        [System.Drawing.Color]::FromArgb(82, 95, 118),
        35
    )
    $g.FillRectangle($bg, 0, 0, $w, $h)
    $bg.Dispose()

    Draw-CenteredText -Graphics $g -Text 'Why Siloed HPC Creates Agency-Wide Friction' -Font $titleFont -Brush ([System.Drawing.Brushes]::White) -Rect (New-Object System.Drawing.RectangleF 70, 30, 1460, 58)
    Draw-CenteredText -Graphics $g -Text 'Fragmentation makes access, sharing, and support harder than they should be' -Font $subFont -Brush ([System.Drawing.Brushes]::Gainsboro) -Rect (New-Object System.Drawing.RectangleF 160, 90, 1280, 36)

    $colFill = New-Object System.Drawing.SolidBrush ([System.Drawing.Color]::FromArgb(238, 243, 248))
    $colPen = New-Object System.Drawing.Pen([System.Drawing.Color]::FromArgb(160, 255, 255, 255), 4)
    $barrierPen = New-Object System.Drawing.Pen([System.Drawing.Color]::FromArgb(230, 210, 70, 70), 10)
    $arrowPen = New-Object System.Drawing.Pen([System.Drawing.Color]::FromArgb(240, 252, 196, 71), 6)
    $labelBrush = New-Object System.Drawing.SolidBrush ([System.Drawing.Color]::FromArgb(27, 48, 70))
    $detailBrush = New-Object System.Drawing.SolidBrush ([System.Drawing.Color]::FromArgb(74, 89, 109))
    $columns = @(
        @{X=120; Label='Center A'; Detail="Separate login`nSeparate storage"},
        @{X=650; Label='Center B'; Detail="Separate tools`nSeparate support"},
        @{X=1180; Label='Center C'; Detail="Idle capacity`nLocked away"}
    )
    foreach ($c in $columns) {
        $rect = New-Object System.Drawing.RectangleF $c.X, 260, 300, 380
        Draw-RoundedRect -Graphics $g -Rect $rect -Radius 20 -Fill $colFill -Pen $colPen
        Draw-CenteredText -Graphics $g -Text $c.Label -Font $boxFont -Brush $labelBrush -Rect (New-Object System.Drawing.RectangleF $c.X, 290, 300, 42)
        Draw-CenteredText -Graphics $g -Text $c.Detail -Font $smallFont -Brush $detailBrush -Rect (New-Object System.Drawing.RectangleF $c.X, 365, 300, 88)
    }
    $colFill.Dispose()
    $colPen.Dispose()
    $labelBrush.Dispose()
    $detailBrush.Dispose()

    $g.DrawLine($barrierPen, 530, 250, 530, 685)
    $g.DrawLine($barrierPen, 1060, 250, 1060, 685)
    $barrierPen.Dispose()

    $g.DrawLine($arrowPen, 460, 730, 1140, 730)
    $g.DrawLine($arrowPen, 1140, 730, 1118, 720)
    $g.DrawLine($arrowPen, 1140, 730, 1118, 740)
    $arrowPen.Dispose()

    $impactFont = [System.Drawing.Font]::new('Segoe UI Semibold', 22, [System.Drawing.FontStyle]::Bold, [System.Drawing.GraphicsUnit]::Point)
    Draw-CenteredText -Graphics $g -Text 'Hard to discover and access available resources' -Font $impactFont -Brush ([System.Drawing.Brushes]::White) -Rect (New-Object System.Drawing.RectangleF 150, 700, 1300, 45)
    $impactFont.Dispose()
}

# Slide 8 image
New-Canvas -Path (Join-Path $OutputDir 'slide8_service_value.png') -Width 1600 -Height 900 -Draw {
    param($g, $w, $h)
    $bg = New-Object System.Drawing.Drawing2D.LinearGradientBrush(
        (New-Object System.Drawing.Rectangle 0, 0, $w, $h),
        [System.Drawing.Color]::FromArgb(15, 52, 82),
        [System.Drawing.Color]::FromArgb(26, 150, 156),
        25
    )
    $g.FillRectangle($bg, 0, 0, $w, $h)
    $bg.Dispose()

    Draw-CenteredText -Graphics $g -Text 'What a Unified FDA HPC Service Should Provide' -Font $titleFont -Brush ([System.Drawing.Brushes]::White) -Rect (New-Object System.Drawing.RectangleF 80, 30, 1440, 58)
    Draw-CenteredText -Graphics $g -Text 'Single access, shared storage, workflows, accounting, and guided interfaces' -Font $subFont -Brush ([System.Drawing.Brushes]::Gainsboro) -Rect (New-Object System.Drawing.RectangleF 130, 92, 1340, 36)

    $boxFill = New-Object System.Drawing.SolidBrush ([System.Drawing.Color]::FromArgb(242, 249, 255))
    $boxPen = New-Object System.Drawing.Pen([System.Drawing.Color]::FromArgb(150, 255, 255, 255), 4)
    $arrowPen = New-Object System.Drawing.Pen([System.Drawing.Color]::FromArgb(250, 255, 223, 99), 7)

    $left = New-Object System.Drawing.RectangleF 120, 290, 360, 300
    $mid = New-Object System.Drawing.RectangleF 620, 290, 360, 300
    $right = New-Object System.Drawing.RectangleF 1120, 290, 360, 300
    foreach ($r in @($left, $mid, $right)) {
        Draw-RoundedRect -Graphics $g -Rect $r -Radius 22 -Fill $boxFill -Pen $boxPen
    }

    $serviceFont = [System.Drawing.Font]::new('Segoe UI Semibold', 24, [System.Drawing.FontStyle]::Bold, [System.Drawing.GraphicsUnit]::Point)
    $bodyFont = [System.Drawing.Font]::new('Segoe UI', 17, [System.Drawing.FontStyle]::Regular, [System.Drawing.GraphicsUnit]::Point)
    $serviceBrush = New-Object System.Drawing.SolidBrush ([System.Drawing.Color]::FromArgb(23, 43, 63))
    $bodyBrush = New-Object System.Drawing.SolidBrush ([System.Drawing.Color]::FromArgb(63, 82, 99))
    Draw-CenteredText -Graphics $g -Text 'Use what FDA already owns' -Font $serviceFont -Brush $serviceBrush -Rect (New-Object System.Drawing.RectangleF 155, 330, 290, 44)
    Draw-CenteredText -Graphics $g -Text 'Improve return on existing infrastructure first' -Font $bodyFont -Brush $bodyBrush -Rect (New-Object System.Drawing.RectangleF 150, 390, 300, 78)

    Draw-CenteredText -Graphics $g -Text 'Shared service layer' -Font $serviceFont -Brush $serviceBrush -Rect (New-Object System.Drawing.RectangleF 655, 330, 290, 44)
    Draw-CenteredText -Graphics $g -Text 'Common access, workflows, and accounting' -Font $bodyFont -Brush $bodyBrush -Rect (New-Object System.Drawing.RectangleF 650, 390, 300, 78)

    Draw-CenteredText -Graphics $g -Text 'Selective cloud bursting' -Font $serviceFont -Brush $serviceBrush -Rect (New-Object System.Drawing.RectangleF 1155, 330, 290, 44)
    Draw-CenteredText -Graphics $g -Text 'Use cloud only where it clearly adds value' -Font $bodyFont -Brush $bodyBrush -Rect (New-Object System.Drawing.RectangleF 1150, 390, 300, 78)

    $boxFill.Dispose()
    $boxPen.Dispose()

    $g.DrawLine($arrowPen, 480, 440, 620, 440)
    $g.DrawLine($arrowPen, 980, 440, 1120, 440)
    $g.DrawLine($arrowPen, 600, 445, 580, 435)
    $g.DrawLine($arrowPen, 600, 445, 580, 455)
    $g.DrawLine($arrowPen, 1100, 445, 1080, 435)
    $g.DrawLine($arrowPen, 1100, 445, 1080, 455)
    $arrowPen.Dispose()

    $badgeBrush = New-Object System.Drawing.SolidBrush ([System.Drawing.Color]::FromArgb(255, 247, 208, 74))
    Draw-RoundedRect -Graphics $g -Rect (New-Object System.Drawing.RectangleF 575, 620, 450, 88) -Radius 18 -Fill $badgeBrush -Pen ([System.Drawing.Pens]::Transparent)
    $badgeTextBrush = New-Object System.Drawing.SolidBrush ([System.Drawing.Color]::FromArgb(27, 48, 70))
    Draw-CenteredText -Graphics $g -Text 'Lower barriers. Improve utilization. Keep options open.' -Font ([System.Drawing.Font]::new('Segoe UI Semibold', 24, [System.Drawing.FontStyle]::Bold, [System.Drawing.GraphicsUnit]::Point)) -Brush $badgeTextBrush -Rect (New-Object System.Drawing.RectangleF 590, 632, 420, 56)
    $badgeBrush.Dispose()
    $serviceBrush.Dispose()
    $bodyBrush.Dispose()
    $badgeTextBrush.Dispose()

    $serviceFont.Dispose()
    $bodyFont.Dispose()
}

$titleFont.Dispose()
$subFont.Dispose()
$boxFont.Dispose()
$smallFont.Dispose()

Write-Output "Created slide images in: $OutputDir"
