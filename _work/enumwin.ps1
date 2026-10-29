$ErrorActionPreference = "Continue"
$src = @"
using System;
using System.Text;
using System.Runtime.InteropServices;
public class WinEnum {
    public delegate bool EnumWindowsProc(IntPtr hWnd, IntPtr lParam);
    [DllImport("user32.dll")] public static extern bool EnumWindows(EnumWindowsProc lpEnumFunc, IntPtr lParam);
    [DllImport("user32.dll")] public static extern int GetWindowText(IntPtr hWnd, StringBuilder lpString, int nMaxCount);
    [DllImport("user32.dll")] public static extern uint GetWindowThreadProcessId(IntPtr hWnd, out uint lpdwProcessId);
    [DllImport("user32.dll")] public static extern bool IsWindowVisible(IntPtr hWnd);
    public static string List(uint pid) {
        StringBuilder sb = new StringBuilder();
        EnumWindows(delegate(IntPtr hWnd, IntPtr lParam) {
            uint wpid;
            GetWindowThreadProcessId(hWnd, out wpid);
            if (wpid == pid) {
                StringBuilder t = new StringBuilder(512);
                GetWindowText(hWnd, t, 512);
                sb.AppendLine("hwnd=" + hWnd.ToInt64() + " visible=" + IsWindowVisible(hWnd) + " title=[" + t.ToString() + "]");
            }
            return true;
        }, IntPtr.Zero);
        return sb.ToString();
    }
}
"@
Add-Type -TypeDefinition $src

$dbPath = "C:\Users\user\Downloads\LLC\LLC_v3_2026-10-29.accdb"
Get-Process MSACCESS -ErrorAction SilentlyContinue | Stop-Process -Force -ErrorAction SilentlyContinue
Start-Sleep -Seconds 2
$acc = New-Object -ComObject Access.Application
$acc.Visible = $true
$acc.AutomationSecurity = 1
$acc.OpenCurrentDatabase($dbPath, $false)
Start-Sleep -Seconds 1
$acc.DoCmd.OpenForm("F_CustomerList")
Start-Sleep -Seconds 2
$p = Get-Process MSACCESS | Select-Object -First 1
Write-Host ("pid=" + $p.Id)
Write-Host ([WinEnum]::List([uint32]$p.Id))
$acc.CloseCurrentDatabase()
$acc.Quit()
Write-Host "DONE"
