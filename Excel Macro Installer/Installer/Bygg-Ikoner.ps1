# ============================================================================
# Bygg-Ikoner.ps1 - lager de fem makro-ikonene (icon-*.ico) fra Håkons flate
# kilde-PNG-er i Resources\Kilde-PNG\ (avrundet gradient-firkant bak motivet).
# Bruk:  powershell -ExecutionPolicy Bypass -File Bygg-Ikoner.ps1
#        (skriver .ico + .png-forhandsvisning til %TEMP%\ikonbygg - kopier
#         deretter .ico-filene MANUELT til ..\Resources\ og rekompiler EMI).
# Viktig: se biHeight-doblingen (ICO-format-fallgruve) og PathGradientBrush-
# Point[]-fallgruven i kommentarene under. Verifiser ALLTID ferdig .ico ved a
# laste den tilbake med System.Drawing.Icon (Width/Height 128x128).
# Nytt ikon: legg PNG i Resources\Kilde-PNG\, legg en linje i $oppgaver nederst.
# MERK (2026-09-21): de opprinnelige kilde-PNG-ene (Kolonnevelger.png, Mail utsender.png,
# Kontaktsentralen.png, Makromeny.png, Nettskjema henter.png - flate stock-ikoner Håkon
# ga oss, transparent bakgrunn, ca 512x512) lå i Downloads og er SLETTET derfra. Kilde-PNG-mappa
# er derfor TOM - be Håkon levere PNG-ene pa nytt hvis ikonene skal bygges om.
# ============================================================================
param(
    [string]$OutDir = (Join-Path $env:TEMP "ikonbygg"),
    [string]$PngMappe = (Join-Path (Split-Path $PSScriptRoot -Parent) "Resources\Kilde-PNG")
)
if (-not (Test-Path $OutDir)) { New-Item -ItemType Directory -Path $OutDir -Force | Out-Null }
Add-Type -AssemblyName System.Drawing

# ---- Hjelpefunksjoner (samme BMP/DIB-teknikk og kjente fallgruver som
# icon.ico/icon-kolonnevelger.ico/icon-mailutsender.ico - se minnet
# excel-macro-installer-launcher for hvorfor: PNG-i-ICO krasjer
# System.Drawing.Icon, range-slicing gir feil byte[]-type, osv.) ----

function New-RoundedRectPath {
    param([float]$X, [float]$Y, [float]$W, [float]$H, [float]$Radius)
    $path = New-Object System.Drawing.Drawing2D.GraphicsPath
    $d = $Radius * 2
    $path.AddArc($X, $Y, $d, $d, 180, 90)
    $path.AddArc($X + $W - $d, $Y, $d, $d, 270, 90)
    $path.AddArc($X + $W - $d, $Y + $H - $d, $d, $d, 0, 90)
    $path.AddArc($X, $Y + $H - $d, $d, $d, 90, 90)
    $path.CloseFigure()
    return $path
}

function New-MacroIcon {
    param(
        [string]$PngSti,
        [System.Drawing.Color]$Lys,
        [System.Drawing.Color]$Mork,
        [string]$UtNavn,
        [int]$Storrelse = 128
    )
    $bmp = New-Object System.Drawing.Bitmap $Storrelse, $Storrelse, ([System.Drawing.Imaging.PixelFormat]::Format32bppArgb)
    $g = [System.Drawing.Graphics]::FromImage($bmp)
    $g.SmoothingMode = [System.Drawing.Drawing2D.SmoothingMode]::AntiAlias
    $g.InterpolationMode = [System.Drawing.Drawing2D.InterpolationMode]::HighQualityBicubic
    $g.PixelOffsetMode = [System.Drawing.Drawing2D.PixelOffsetMode]::HighQuality

    $radius = [float]($Storrelse * 0.19)
    $margin = [float]($Storrelse * 0.035)
    $rectPath = New-RoundedRectPath -X $margin -Y $margin -W ($Storrelse - $margin * 2) -H ($Storrelse - $margin * 2) -Radius $radius
    $g.SetClip($rectPath)

    # Diagonal gradient bakgrunn (samme prinsipp som icon.ico)
    $rectF = New-Object System.Drawing.RectangleF(0, 0, $Storrelse, $Storrelse)
    $grad = New-Object System.Drawing.Drawing2D.LinearGradientBrush($rectF, $Lys, $Mork, 45.0)
    $g.FillRectangle($grad, 0, 0, $Storrelse, $Storrelse)

    # Svak blank hoydeglans ovre halvdel (matcher icon.ico sin "glossy radial
    # highlight"), lav opasitet, kun for litt dybde/polish.
    $glansPunkter = [System.Drawing.Point[]]@(
        (New-Object System.Drawing.Point(0, 0)),
        (New-Object System.Drawing.Point($Storrelse, 0)),
        (New-Object System.Drawing.Point($Storrelse, [int]($Storrelse * 0.55))),
        (New-Object System.Drawing.Point(0, [int]($Storrelse * 0.55)))
    )
    $glansBrush = New-Object System.Drawing.Drawing2D.PathGradientBrush(,$glansPunkter)
    $glansBrush.CenterColor = [System.Drawing.Color]::FromArgb(60, 255, 255, 255)
    $glansBrush.SurroundColors = @([System.Drawing.Color]::FromArgb(0, 255, 255, 255))
    $glansBrush.CenterPoint = New-Object System.Drawing.PointF(($Storrelse * 0.5), ($Storrelse * 0.05))
    $g.FillRectangle($glansBrush, 0, 0, $Storrelse, [int]($Storrelse * 0.55))

    $g.ResetClip()

    # Selve PNG-illustrasjonen, sentrert, skalert til ca 72% av flaten (en
    # synlig kant av den avrundede firkanten rundt - matcher hvordan
    # icon-mailutsender.ico sin konvolutt sitter pa den gamle ikon-platen).
    $src = [System.Drawing.Image]::FromFile($PngSti)
    $maalStorrelse = [float]($Storrelse * 0.72)
    $skala = [Math]::Min($maalStorrelse / $src.Width, $maalStorrelse / $src.Height)
    $mBredde = $src.Width * $skala
    $mHoyde = $src.Height * $skala
    $mX = ($Storrelse - $mBredde) / 2
    $mY = ($Storrelse - $mHoyde) / 2
    $g.DrawImage($src, $mX, $mY, $mBredde, $mHoyde)
    $src.Dispose()

    $g.Dispose()

    # Lagre en PNG-forhandsvisning (for visuell kontroll - IKKE det som
    # brukes i selve appen).
    $bmp.Save((Join-Path $OutDir "$UtNavn.png"), [System.Drawing.Imaging.ImageFormat]::Png)

    # ---- Bygg ekte .ico (BMP/DIB-basert, ikke PNG-i-ICO) ----
    # 1) Lagre som ekte BMP, strip 14-byte BITMAPFILEHEADER, behold
    #    BITMAPINFOHEADER + pikseldata.
    $bmpSti = Join-Path $OutDir "$UtNavn.tmp.bmp"
    $bmp.Save($bmpSti, [System.Drawing.Imaging.ImageFormat]::Bmp)
    $alleBytes = [System.IO.File]::ReadAllBytes($bmpSti)
    Remove-Item $bmpSti -Force

    $dibLengde = $alleBytes.Length - 14
    $dib = New-Object byte[] $dibLengde
    [Array]::Copy($alleBytes, 14, $dib, 0, $dibLengde)

    # ICO-formatet krever at BITMAPINFOHEADER.biHeight (byte-offset 8 i DIB-
    # en) er DOBBEL den faktiske pikselhoyden - selv om kun ETT sett pikseldata
    # (ikke to) faktisk folger, pluss en egen AND-maske. En vanlig BMP (som
    # $bmp.Save(...,Bmp) skrev) har korrekt, IKKE doblet biHeight - lot vi den
    # sta som den var, leser ikonlastere (bekreftet: System.Drawing.Icon)
    # tilbake halve hoyden (128x64 i stedet for 128x128). Fiks: doble verdien
    # i selve byte-arrayet for hand.
    $biHeight = [BitConverter]::ToInt32($dib, 8)
    $dobbelHeightBytes = [BitConverter]::GetBytes($biHeight * 2)
    [Array]::Copy($dobbelHeightBytes, 0, $dib, 8, 4)

    # 2) AND-maske fra faktisk per-piksel alpha (via LockBits) - alpha<128 = maskert.
    $bredde = $bmp.Width
    $hoyde = $bmp.Height
    $rd = New-Object System.Drawing.Rectangle(0, 0, $bredde, $hoyde)
    $bd = $bmp.LockBits($rd, [System.Drawing.Imaging.ImageLockMode]::ReadOnly, [System.Drawing.Imaging.PixelFormat]::Format32bppArgb)
    $pikselByteAntall = [Math]::Abs($bd.Stride) * $hoyde
    $pikselBytes = New-Object byte[] $pikselByteAntall
    [System.Runtime.InteropServices.Marshal]::Copy($bd.Scan0, $pikselBytes, 0, $pikselByteAntall)
    $bmp.UnlockBits($bd)

    $maskeRadBytes = [Math]::Ceiling($bredde / 8.0)
    if (($maskeRadBytes % 4) -ne 0) { $maskeRadBytes = $maskeRadBytes + (4 - ($maskeRadBytes % 4)) } # 4-byte-justert
    $maske = New-Object byte[] ($maskeRadBytes * $hoyde)
    for ($y = 0; $y -lt $hoyde; $y++) {
        for ($x = 0; $x -lt $bredde; $x++) {
            $pikselOffset = ($y * $bd.Stride) + ($x * 4)
            $alpha = $pikselBytes[$pikselOffset + 3]
            if ($alpha -lt 128) {
                # Bitmap er lagret bunn-opp i DIB - maske-raden skal matche.
                $maskeY = $hoyde - 1 - $y
                $byteIndeks = ($maskeY * $maskeRadBytes) + ($x -shr 3)
                $bitIndeks = 7 - ($x -band 7)
                $maske[$byteIndeks] = $maske[$byteIndeks] -bor (1 -shl $bitIndeks)
            }
        }
    }

    # 3) ICONDIR + ICONDIRENTRY + (BITMAPINFOHEADER+pikseldata+maske)
    $icoData = New-Object byte[] ($dib.Length + $maske.Length)
    [Array]::Copy($dib, 0, $icoData, 0, $dib.Length)
    [Array]::Copy($maske, 0, $icoData, $dib.Length, $maske.Length)

    $ms = New-Object System.IO.MemoryStream
    $bw = New-Object System.IO.BinaryWriter($ms)
    # ICONDIR
    $bw.Write([UInt16]0)       # reserved
    $bw.Write([UInt16]1)       # type = icon
    $bw.Write([UInt16]1)       # antall bilder
    # ICONDIRENTRY
    $bBredde = if ($bredde -ge 256) { 0 } else { [byte]$bredde }
    $bHoyde = if ($hoyde -ge 256) { 0 } else { [byte]$hoyde }
    $bw.Write([byte]$bBredde)
    $bw.Write([byte]$bHoyde)
    $bw.Write([byte]0)        # fargepalett
    $bw.Write([byte]0)        # reserved
    $bw.Write([UInt16]1)      # fargeplan
    $bw.Write([UInt16]32)     # bits per piksel
    $bw.Write([UInt32]$icoData.Length)  # storrelse pa bildedata
    $bw.Write([UInt32]22)     # offset til bildedata (6 + 16)
    $bw.Write($icoData)
    $bw.Flush()

    [System.IO.File]::WriteAllBytes((Join-Path $OutDir "$UtNavn.ico"), $ms.ToArray())
    $ms.Dispose()
    $bmp.Dispose()

    Write-Output "Bygget: $UtNavn (.ico + .png forhandsvisning)"
}

$oppgaver = @(
    @{ Navn = "icon-kolonnevelger";    Png = (Join-Path $PngMappe "Kolonnevelger.png");      Lys = [System.Drawing.Color]::FromArgb(196,181,253); Mork = [System.Drawing.Color]::FromArgb(109,40,217) },
    @{ Navn = "icon-mailutsender";     Png = (Join-Path $PngMappe "Mail utsender.png");      Lys = [System.Drawing.Color]::FromArgb(147,197,253); Mork = [System.Drawing.Color]::FromArgb(29,78,216) },
    @{ Navn = "icon-kontaktsentralen"; Png = (Join-Path $PngMappe "Kontaktsentralen.png");   Lys = [System.Drawing.Color]::FromArgb(110,231,183); Mork = [System.Drawing.Color]::FromArgb(4,120,87) },
    @{ Navn = "icon-makromeny";        Png = (Join-Path $PngMappe "Makromeny.png");          Lys = [System.Drawing.Color]::FromArgb(253,186,116); Mork = [System.Drawing.Color]::FromArgb(194,65,12) },
    @{ Navn = "icon-nettskjemahenter"; Png = (Join-Path $PngMappe "Nettskjema henter.png");  Lys = [System.Drawing.Color]::FromArgb(103,232,249); Mork = [System.Drawing.Color]::FromArgb(21,94,117) }
)

foreach ($o in $oppgaver) {
    New-MacroIcon -PngSti $o.Png -Lys $o.Lys -Mork $o.Mork -UtNavn $o.Navn
}
