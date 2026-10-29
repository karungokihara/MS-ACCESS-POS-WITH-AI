$ErrorActionPreference = "Stop"
Add-Type -AssemblyName System.Windows.Forms, System.Drawing
Add-Type @"
using System;
using System.Runtime.InteropServices;
public class Win {
  [DllImport("user32.dll")] public static extern bool SetForegroundWindow(IntPtr h);
  [DllImport("user32.dll")] public static extern bool ShowWindow(IntPtr h, int c);
  [DllImport("user32.dll")] public static extern bool GetWindowRect(IntPtr h, out RECT r);
  public struct RECT { public int Left, Top, Right, Bottom; }
}
"@
Get-Process MSACCESS -ErrorAction SilentlyContinue | Stop-Process -Force -ErrorAction SilentlyContinue
Start-Sleep -Seconds 1
$acc = New-Object -ComObject Access.Application
$acc.OpenCurrentDatabase("C:\Users\user\Downloads\LLC\LLC.accdb")
$acc.Visible = $true
$acc.UserControl = $false
try { $acc.DoCmd.Close(2, "F_Login", 2) } catch {}
$adminID = $acc.DLookup("UserID", "T_Users", "Username='admin'")
$roleID = $acc.DLookup("RoleID", "T_Users", "Username='admin'")
$acc.TempVars.Add("CurrentUserID", $adminID)
$acc.TempVars.Add("CurrentRoleID", $roleID)
$acc.DoCmd.OpenForm("Dashboard", 1)
$d = $acc.Forms.Item("Dashboard")
Write-Host ("Picture=[" + $d.Picture + "] Type=" + $d.PictureType + " Size=" + $d.PictureSizeMode)
$acc.DoCmd.Close(2, "Dashboard", 2)
$acc.DoCmd.OpenForm("Dashboard", 0)
$acc.DoCmd.Maximize()
Start-Sleep -Seconds 3
$h = [IntPtr]$acc.hWndAccessApp()
[Win]::ShowWindow($h, 3) | Out-Null
[Win]::SetForegroundWindow($h) | Out-Null
Start-Sleep -Seconds 2
function Snap($path) {
    $r = New-Object Win+RECT
    [Win]::GetWindowRect($h, [ref]$r) | Out-Null
    $w = $r.Right - $r.Left; $hh = $r.Bottom - $r.Top
    $bmp = New-Object System.Drawing.Bitmap $w, $hh
    $g = [System.Drawing.Graphics]::FromImage($bmp)
    $g.CopyFromScreen($r.Left, $r.Top, 0, 0, $bmp.Size)
    $bmp.Save($path, [System.Drawing.Imaging.ImageFormat]::Png)
    $g.Dispose(); $bmp.Dispose()
}
Snap "C:\Users\user\Downloads\LLC\_inspect\snap-home.png"
$arg = "F_SystemSettings"
$acc.Run("UX_ShowModule", [ref]$arg) | Out-Null
Start-Sleep -Seconds 3
Snap "C:\Users\user\Downloads\LLC\_inspect\snap-settings.png"
$arg2 = "F_BusinessReadiness"
$acc.Run("UX_ShowModule", [ref]$arg2) | Out-Null
Start-Sleep -Seconds 3
Snap "C:\Users\user\Downloads\LLC\_inspect\snap-sales.png"
$acc.Quit(2)
Write-Host "SNAPPED"
