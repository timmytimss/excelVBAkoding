param([long]$Hwnd)
Add-Type -MemberDefinition '[DllImport("user32.dll")] public static extern bool PostMessage(IntPtr h, uint m, IntPtr w, IntPtr l);' -Name CW -Namespace CWN
[CWN.CW]::PostMessage([IntPtr]$Hwnd, 0x10, [IntPtr]::Zero, [IntPtr]::Zero) | Out-Null
