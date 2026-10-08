Add-Type -AssemblyName System.Drawing
$ErrorActionPreference = 'Stop'
. "$PSScriptRoot\write-raw-ico.ps1"

function New-RoundedGradientBitmap {
    param([int]$Size, [System.Drawing.Color]$TopColor, [System.Drawing.Color]$BottomColor)
    $bmp = New-Object System.Drawing.Bitmap($Size, $Size, [System.Drawing.Imaging.PixelFormat]::Format32bppArgb)
    $g = [System.Drawing.Graphics]::FromImage($bmp)
    $g.SmoothingMode = [System.Drawing.Drawing2D.SmoothingMode]::AntiAlias
    $g.Clear([System.Drawing.Color]::Transparent)

    $radius = [single]($Size * 0.20)
    $rect = New-Object System.Drawing.RectangleF(0, 0, $Size, $Size)
    $path = New-Object System.Drawing.Drawing2D.GraphicsPath
    $d = $radius * 2
    $path.AddArc(0, 0, $d, $d, 180, 90)
    $path.AddArc(($Size - $d), 0, $d, $d, 270, 90)
    $path.AddArc(($Size - $d), ($Size - $d), $d, $d, 0, 90)
    $path.AddArc(0, ($Size - $d), $d, $d, 90, 90)
    $path.CloseFigure()
    $brush = New-Object System.Drawing.Drawing2D.LinearGradientBrush($rect, $TopColor, $BottomColor, 45)
    $g.FillPath($brush, $path)

    $barColor = [System.Drawing.Color]::FromArgb(255, 255, 255, 255)
    $barBrush = New-Object System.Drawing.SolidBrush($barColor)
    $nBars = 3
    $barWidth = [single]($Size * 0.14)
    $gap = [single]($Size * 0.09)
    $totalWidth = ($nBars * $barWidth) + (($nBars - 1) * $gap)
    $startX = ($Size - $totalWidth) / 2
    $baseY = [single]($Size * 0.74)
    $heights = @(0.28, 0.46, 0.62)
    for ($i = 0; $i -lt $nBars; $i++) {
        $bh = [single]($Size * $heights[$i])
        $bx = $startX + ($i * ($barWidth + $gap))
        $by = $baseY - $bh
        $bp = New-Object System.Drawing.Drawing2D.GraphicsPath
        $br = [single]($barWidth * 0.28)
        $bd = $br * 2
        $bp.AddArc($bx, $by, $bd, $bd, 180, 90)
        $bp.AddArc(($bx + $barWidth - $bd), $by, $bd, $bd, 270, 90)
        $bp.AddLine(($bx + $barWidth), ($by + $bh), $bx, ($by + $bh))
        $bp.CloseFigure()
        $g.FillPath($barBrush, $bp)
    }
    $g.Dispose()
    return $bmp
}

$top = [System.Drawing.Color]::FromArgb(254, 226, 226)
$bottom = [System.Drawing.Color]::FromArgb(239, 68, 68)

$bmp128 = New-RoundedGradientBitmap -Size 128 -TopColor $top -BottomColor $bottom
$outIco = 'C:\Users\haaklun\Documents\excelVBAkoding\Excel Macro Installer\Resources\icon-statistikkern.ico'
Write-RawIco -Bmp $bmp128 -OutPath $outIco
$bmp128.Dispose()
Write-Output "Skrev icon-statistikkern.ico (raw DIB-metode)"
