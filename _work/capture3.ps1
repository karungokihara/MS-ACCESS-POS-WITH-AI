$ErrorActionPreference = "Continue"
$src = @"
using System;
using System.Drawing;
using System.Drawing.Imaging;
using System.Text;
using System.Runtime.InteropServices;
public class WinCap3 {
    public delegate bool EnumWindowsProc(IntPtr hWnd, IntPtr lParam);
    [DllImport("user32.dll")] public static extern bool EnumWindows(EnumWindowsProc lpEnumFunc, IntPtr lParam);
    [DllImport("user32.dll")] public static extern int GetWindowText(IntPtr hWnd, StringBuilder lpString, int nMaxCount);
    [DllImport("user32.dll")] public static extern uint GetWindowThreadProcessId(IntPtr hWnd, out uint lpdwProcessId);
    [DllImport("user32.dll")] public static extern bool PrintWindow(IntPtr hWnd, IntPtr hdcBlt, uint nFlags);
    [DllImport("user32.dll")] public static extern bool GetWindowRect(IntPtr hWnd, out RECT lpRect);
    [StructLayout(LayoutKind.Sequential)] public struct RECT { public int Left, Top, Right, Bottom; }
    public static IntPtr Find(uint pid, string title) {
        IntPtr found = IntPtr.Zero;
        EnumWindows(delegate(IntPtr hWnd, IntPtr lParam) {
            uint wpid;
            GetWindowThreadProcessId(hWnd, out wpid);
            if (wpid == pid) {
                StringBuilder t = new StringBuilder(512);
                GetWindowText(hWnd, t, 512);
                if (t.ToString() == title) { found = hWnd; return false; }
            }
            return true;
        }, IntPtr.Zero);
        return found;
    }
    public static void Capture(IntPtr hWnd, string path) {
        RECT r;
        GetWindowRect(hWnd, out r);
        Bitmap bmp = new Bitmap(r.Right - r.Left, r.Bottom - r.Top);
        Graphics g = Graphics.FromImage(bmp);
        IntPtr hdc = g.GetHdc();
        PrintWindow(hWnd, hdc, 2);
        g.ReleaseHdc(hdc);
        g.Dispose();
        bmp.Save(path, ImageFormat.Png);
        bmp.Dispose();
    }
}
"@
Add-Type -TypeDefinition $src -ReferencedAssemblies System.Drawing

$dbPath = "C:\Users\user\Downloads\LLC\LLC_v3_2026-10-29.accdb"
Get-Process MSACCESS -ErrorAction SilentlyContinue | Stop-Process -Force -ErrorAction SilentlyContinue
Start-Sleep -Seconds 2
$acc = New-Object -ComObject Access.Application
$acc.Visible = $true
$acc.AutomationSecurity = 1
$acc.OpenCurrentDatabase($dbPath, $false)
Start-Sleep -Seconds 1
$p = Get-Process MSACCESS | Select-Object -First 1

$acc.DoCmd.OpenForm("F_UnitList")
Start-Sleep -Seconds 2
$h = [WinCap3]::Find([uint32]$p.Id, "INVENTORY SYSTEM")
if ($h -ne [IntPtr]::Zero) { [WinCap3]::Capture($h, "C:\Users\user\Downloads\LLC\_work\screen_unitlist2.png"); Write-Host "captured units" }

try { $acc.DoCmd.OpenForm("F_AssemblyTemplate", 0, [System.Reflection.Missing]::Value, [System.Reflection.Missing]::Value, 1) } catch { Write-Host ("open error: " + $_.Exception.Message) }
Start-Sleep -Seconds 2
Write-Host ("template loaded = " + $acc.CurrentProject.AllForms("F_AssemblyTemplate").IsLoaded)
if ($acc.CurrentProject.AllForms("F_AssemblyTemplate").IsLoaded) {
    $at = $acc.Forms.Item("F_AssemblyTemplate")
    Write-Host ("controls = " + $at.Controls.Count)
    foreach ($n in @("txtManufacturerPartNumber","txtVendorPartNumber","txtInternalPartNumber","chkAutoGeneratePartNumber","lblAutoGeneratePartNumber")) {
        try { $c = $at.Controls.Item($n); Write-Host ("{0} L={1} T={2} W={3} Visible={4}" -f $n,$c.Left,$c.Top,$c.Width,$c.Visible) } catch { Write-Host ($n + " MISSING") }
    }
    $h = [WinCap3]::Find([uint32]$p.Id, "INVENTORY SYSTEM")
    if ($h -ne [IntPtr]::Zero) { [WinCap3]::Capture($h, "C:\Users\user\Downloads\LLC\_work\screen_template2.png"); Write-Host "captured template" }
}
$acc.CloseCurrentDatabase()
$acc.Quit()
Write-Host "DONE"
