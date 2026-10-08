param([long]$Hwnd, [string]$OutPath)
Add-Type -AssemblyName System.Drawing
Add-Type @"
using System;
using System.Runtime.InteropServices;
public class WinShot {
    [DllImport("user32.dll")] public static extern bool SetProcessDpiAwarenessContext(IntPtr value);
    [DllImport("user32.dll")] public static extern bool GetWindowRect(IntPtr hWnd, out RECT lpRect);
    [DllImport("user32.dll")] public static extern bool PrintWindow(IntPtr hWnd, IntPtr hdcBlt, uint nFlags);
    [DllImport("user32.dll")] public static extern bool SetForegroundWindow(IntPtr hWnd);
    public struct RECT { public int Left; public int Top; public int Right; public int Bottom; }
}
"@
[WinShot]::SetProcessDpiAwarenessContext([IntPtr](-4)) | Out-Null

$h = [IntPtr]$Hwnd
$rect = New-Object WinShot+RECT
[WinShot]::GetWindowRect($h, [ref]$rect) | Out-Null
$w = $rect.Right - $rect.Left
$ht = $rect.Bottom - $rect.Top

$bmp = New-Object System.Drawing.Bitmap $w, $ht
$g = [System.Drawing.Graphics]::FromImage($bmp)
$hdc = $g.GetHdc()
[WinShot]::PrintWindow($h, $hdc, 2) | Out-Null
$g.ReleaseHdc($hdc)
$bmp.Save($OutPath, [System.Drawing.Imaging.ImageFormat]::Png)
$g.Dispose()
$bmp.Dispose()
Write-Output "Lagret: $OutPath ($w x $ht)"
