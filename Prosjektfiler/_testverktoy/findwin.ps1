param([string]$Title)
$sig = '[DllImport("user32.dll")] public static extern bool EnumWindows(EnumWindowsProc f, IntPtr l); public delegate bool EnumWindowsProc(IntPtr h, IntPtr l); [DllImport("user32.dll")] public static extern int GetWindowText(IntPtr h, System.Text.StringBuilder t, int c);'
Add-Type -MemberDefinition $sig -Name FW -Namespace FWN
$script:t = [IntPtr]::Zero
[FWN.FW]::EnumWindows({ param($h,$l) $sb = New-Object System.Text.StringBuilder 256; [FWN.FW]::GetWindowText($h,$sb,256)|Out-Null; if ($sb.ToString() -eq $Title) { $script:t = $h; return $false }; return $true }, [IntPtr]::Zero) | Out-Null
Write-Output $script:t.ToInt64()
