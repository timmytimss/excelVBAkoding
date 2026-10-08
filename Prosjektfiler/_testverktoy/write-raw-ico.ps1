Add-Type -AssemblyName System.Drawing
$ErrorActionPreference = 'Stop'

function Write-RawIco {
    param([System.Drawing.Bitmap]$Bmp, [string]$OutPath)
    $w = $Bmp.Width
    $h = $Bmp.Height

    $rect = New-Object System.Drawing.Rectangle(0, 0, $w, $h)
    $bd = $Bmp.LockBits($rect, [System.Drawing.Imaging.ImageLockMode]::ReadOnly, [System.Drawing.Imaging.PixelFormat]::Format32bppArgb)
    $stride = $bd.Stride
    $byteCount = $stride * $h
    $pixelBytes = New-Object byte[] $byteCount
    [System.Runtime.InteropServices.Marshal]::Copy($bd.Scan0, $pixelBytes, 0, $byteCount)
    $Bmp.UnlockBits($bd)

    # DIB-rader er bunn-opp i XOR-dataen (samme rekkefolge som BMP-filformatet)
    # - Bitmap.LockBits gir allerede topp-til-bunn i minnet, saa vi reverserer
    # radrekkefolgen her.
    $xor = New-Object byte[] $byteCount
    for ($row = 0; $row -lt $h; $row++) {
        [Array]::Copy($pixelBytes, $row * $stride, $xor, ($h - 1 - $row) * $stride, $stride)
    }

    # AND-maske: 1bpp, radbredde paddet til 4-byte multiplum, bunn-opp. Satt
    # til alt 0 (uansett maskert) - alfakanalen i XOR-dataen styrer faktisk
    # gjennomsiktighet for moderne lesere (inkl. .NET sin Icon-klasse).
    $andStride = [Math]::Ceiling($w / 32.0) * 4
    $andBytes = New-Object byte[] ($andStride * $h)

    $dibHeaderSize = 40
    $dibSize = $dibHeaderSize + $xor.Length + $andBytes.Length

    $ms = New-Object System.IO.MemoryStream
    $bw = New-Object System.IO.BinaryWriter($ms)

    # ICONDIR
    $bw.Write([UInt16]0)   # reserved
    $bw.Write([UInt16]1)   # type = icon
    $bw.Write([UInt16]1)   # count

    # ICONDIRENTRY
    $bw.Write([byte]($(if ($w -ge 256) { 0 } else { $w })))
    $bw.Write([byte]($(if ($h -ge 256) { 0 } else { $h })))
    $bw.Write([byte]0)     # color count
    $bw.Write([byte]0)     # reserved
    $bw.Write([UInt16]1)   # planes
    $bw.Write([UInt16]32)  # bit count
    $bw.Write([UInt32]$dibSize)
    $bw.Write([UInt32]22)  # offset (6 + 16)

    # BITMAPINFOHEADER
    $bw.Write([UInt32]$dibHeaderSize)
    $bw.Write([Int32]$w)
    $bw.Write([Int32]($h * 2))   # dobbel hoyde (XOR + AND) per ICO-konvensjon
    $bw.Write([UInt16]1)         # planes
    $bw.Write([UInt16]32)        # bit count
    $bw.Write([UInt32]0)         # compression = BI_RGB
    $bw.Write([UInt32]($xor.Length + $andBytes.Length))
    $bw.Write([Int32]0)          # x pixels per meter
    $bw.Write([Int32]0)          # y pixels per meter
    $bw.Write([UInt32]0)         # colors used
    $bw.Write([UInt32]0)         # colors important

    $bw.Write($xor)
    $bw.Write($andBytes)

    $bw.Flush()
    [System.IO.File]::WriteAllBytes($OutPath, $ms.ToArray())
    $bw.Close()
    $ms.Close()
}
