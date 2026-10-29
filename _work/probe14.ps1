$ErrorActionPreference="Continue"
$src = @"
using System;
using System.Text;
using System.Collections.Generic;
using System.Runtime.InteropServices;
public class WinEnum {
    public delegate bool EnumWindowsProc(IntPtr hWnd, IntPtr lParam);
    [DllImport("user32.dll")] public static extern bool EnumWindows(EnumWindowsProc lpEnumFunc, IntPtr lParam);
    [DllImport("user32.dll")] public static extern int GetWindowText(IntPtr hWnd, StringBuilder lpString, int nMaxCount);
    [DllImport("user32.dll")] public static extern uint GetWindowThreadProcessId(IntPtr hWnd, out uint lpdwProcessId);
    [DllImport("user32.dll")] public static extern bool IsWindowVisible(IntPtr hWnd);
    public static List<string> Titles(uint pid) {
        List<string> list = new List<string>();
        EnumWindows(delegate(IntPtr hWnd, IntPtr lParam) {
            uint wpid; GetWindowThreadProcessId(hWnd, out wpid);
            if (wpid == pid && IsWindowVisible(hWnd)) {
                StringBuilder t = new StringBuilder(512); GetWindowText(hWnd, t, 512);
                if (t.Length > 0) list.Add(hWnd.ToInt64() + " | " + t.ToString());
            }
            return true;
        }, IntPtr.Zero);
        return list;
    }
}
"@
Add-Type -TypeDefinition $src
$dbPath="C:\Users\user\Downloads\LLC\LLC_v3_2026-10-29.accdb"
Get-Process MSACCESS -ErrorAction SilentlyContinue | Stop-Process -Force -ErrorAction SilentlyContinue
Start-Sleep -Seconds 1
$acc=New-Object -ComObject Access.Application
$acc.Visible=$true; $acc.AutomationSecurity=1
$acc.OpenCurrentDatabase($dbPath,$false)
Start-Sleep -Seconds 1
$acc.DoCmd.OpenForm("F_AssemblyTemplate")
Start-Sleep -Seconds 2
$p=Get-Process MSACCESS | Select-Object -First 1
[WinEnum]::Titles([uint32]$p.Id) | ForEach-Object { Write-Host $_ }
try { $acc.DoCmd.SelectObject(2,"F_AssemblyTemplate",$true) } catch { Write-Host ("sel err " + $_.Exception.Message) }
Start-Sleep -Milliseconds 800
Write-Host ("active=" + $acc.Screen.ActiveForm.Name)
$f=$acc.Forms.Item("F_AssemblyTemplate")
Write-Host ("count=" + $f.Controls.Count + " name=" + $f.Name + " caption=" + $f.Caption + " popup=" + $f.Popup)
$acc.CloseCurrentDatabase(); $acc.Quit()
Write-Host "DONE"
