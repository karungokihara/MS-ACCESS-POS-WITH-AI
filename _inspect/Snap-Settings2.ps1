$ErrorActionPreference = "Stop"
Add-Type -AssemblyName System.Windows.Forms, System.Drawing
Add-Type @"
using System;
using System.Runtime.InteropServices;
public class Win2 {
  [DllImport("user32.dll")] public static extern bool SetForegroundWindow(IntPtr h);
  [DllImport("user32.dll")] public static extern bool ShowWindow(IntPtr h, int c);
  [DllImport("user32.dll")] public static extern bool GetWindowRect(IntPtr h, out RECT r);
  public struct RECT { public int Left, Top, Right, Bottom; }
}
"@
function Snap($h, $path) {
    $r = New-Object Win2+RECT
    [Win2]::GetWindowRect($h, [ref]$r) | Out-Null
    $bmp = New-Object System.Drawing.Bitmap ($r.Right - $r.Left), ($r.Bottom - $r.Top)
    $g = [System.Drawing.Graphics]::FromImage($bmp)
    $g.CopyFromScreen($r.Left, $r.Top, 0, 0, $bmp.Size)
    $bmp.Save($path, [System.Drawing.Imaging.ImageFormat]::Png)
    $g.Dispose(); $bmp.Dispose()
}
Get-Process MSACCESS -ErrorAction SilentlyContinue | Stop-Process -Force -ErrorAction SilentlyContinue
Start-Sleep -Seconds 1
$acc = New-Object -ComObject Access.Application
$acc.OpenCurrentDatabase("C:\Users\user\Downloads\LLC\LLC.accdb")
$acc.Visible = $true
try { $acc.DoCmd.Close(2, "F_Login", 2) } catch {}
$acc.TempVars.Add("CurrentUserID", $acc.DLookup("UserID", "T_Users", "Username='admin'"))
$acc.TempVars.Add("CurrentRoleID", $acc.DLookup("RoleID", "T_Users", "Username='admin'"))
$h = [IntPtr]$acc.hWndAccessApp()
[Win2]::ShowWindow($h, 3) | Out-Null
$acc.DoCmd.OpenForm("F_SystemSettings", 0)
$acc.DoCmd.Maximize()
[Win2]::SetForegroundWindow($h) | Out-Null
Start-Sleep -Seconds 3
Snap $h "C:\Users\user\Downloads\LLC\_inspect\snap-settings-alone.png"
$acc.DoCmd.Close(2, "F_SystemSettings", 2)
$acc.DoCmd.OpenForm("Dashboard", 0)
$arg = "F_SystemSettings"
$acc.Run("UX_ShowModule", [ref]$arg) | Out-Null
[Win2]::SetForegroundWindow($h) | Out-Null
Start-Sleep -Seconds 3
Snap $h "C:\Users\user\Downloads\LLC\_inspect\snap-settings-panel.png"
$acc.Quit(2)
Write-Host "SNAPPED"
