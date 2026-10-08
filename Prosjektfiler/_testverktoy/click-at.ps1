param([long]$Hwnd,[int]$X,[int]$Y)
Add-Type @"
using System; using System.Runtime.InteropServices;
public class Clk { [DllImport("user32.dll")] public static extern bool SetProcessDpiAwarenessContext(IntPtr v);
 [DllImport("user32.dll")] public static extern bool GetWindowRect(IntPtr h, out RC r);
 [DllImport("user32.dll")] public static extern bool SetCursorPos(int x,int y);
 [DllImport("user32.dll")] public static extern bool GetCursorPos(out PT p);
 [DllImport("user32.dll")] public static extern void mouse_event(uint f,uint dx,uint dy,uint d,IntPtr e);
 [DllImport("user32.dll")] public static extern bool SetForegroundWindow(IntPtr h);
 public struct RC{public int L,T,R,B;} public struct PT{public int X,Y;} }
"@
[Clk]::SetProcessDpiAwarenessContext([IntPtr](-4)) | Out-Null
$r = New-Object Clk+RC; [Clk]::GetWindowRect([IntPtr]$Hwnd,[ref]$r) | Out-Null
$o = New-Object Clk+PT; [Clk]::GetCursorPos([ref]$o) | Out-Null
[Clk]::SetForegroundWindow([IntPtr]$Hwnd) | Out-Null
Start-Sleep -Milliseconds 300
[Clk]::SetCursorPos($r.L+$X, $r.T+$Y) | Out-Null
Start-Sleep -Milliseconds 150
[Clk]::mouse_event(2,0,0,0,[IntPtr]::Zero); Start-Sleep -Milliseconds 60; [Clk]::mouse_event(4,0,0,0,[IntPtr]::Zero)
Start-Sleep -Milliseconds 200
[Clk]::SetCursorPos($o.X,$o.Y) | Out-Null
"klikket ($($r.L+$X),$($r.T+$Y)), markor gjenopprettet"
