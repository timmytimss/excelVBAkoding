param([long]$Hwnd,[string]$Key)
Add-Type @"
using System; using System.Runtime.InteropServices;
public class Kb {
    [DllImport("user32.dll")] public static extern bool SetForegroundWindow(IntPtr h);
    [DllImport("user32.dll")] public static extern void keybd_event(byte bVk, byte bScan, uint dwFlags, IntPtr dwExtraInfo);
}
"@
[Kb]::SetForegroundWindow([IntPtr]$Hwnd) | Out-Null
Start-Sleep -Milliseconds 200
$vk = switch ($Key) {
    "Down" { 0x28 }
    "Enter" { 0x0D }
    "Up" { 0x26 }
    "Escape" { 0x1B }
    default { 0 }
}
if ($vk -eq 0) { Write-Output "Ukjent tast"; exit }
[Kb]::keybd_event($vk, 0, 0, [IntPtr]::Zero)
Start-Sleep -Milliseconds 60
[Kb]::keybd_event($vk, 0, 2, [IntPtr]::Zero)
Write-Output "Tast sendt: $Key"
