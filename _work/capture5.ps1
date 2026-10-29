$ErrorActionPreference = "Continue"
$src = @"
using System;
using System.Drawing;
using System.Drawing.Imaging;
using System.Text;
using System.Runtime.InteropServices;
public class WinCap5 {
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

$srcDb = "C:\Users\user\Downloads\LLC\LLC_v3_2026-10-29.accdb"
$testDb = "C:\Users\user\Downloads\LLC\_work\test_customer.accdb"
Copy-Item $srcDb $testDb -Force
Get-Process MSACCESS -ErrorAction SilentlyContinue | Stop-Process -Force -ErrorAction SilentlyContinue
Start-Sleep -Seconds 2

$acc = New-Object -ComObject Access.Application
$acc.Visible = $true
$acc.AutomationSecurity = 1
$acc.OpenCurrentDatabase($testDb, $false)
Start-Sleep -Seconds 1

$acc.CurrentDb().Execute("DELETE FROM T_Customers WHERE CustomerName='TEST CUSTOMER'") | Out-Null
$acc.CurrentDb().Execute("INSERT INTO T_Customers (CustomerName, CustomerType, Status) VALUES ('TEST CUSTOMER','Retail','In Review')") | Out-Null
Write-Host ("incomplete customers = " + $acc.Eval('UX_AreaIncomplete("customer")'))
Write-Host ("total customers = " + $acc.Eval('UX_AreaTotal("customer")'))
Write-Host ("filter = " + $acc.Eval('UX_AreaDrillFilter("customer")'))

$p = Get-Process MSACCESS | Select-Object -First 1

$acc.DoCmd.OpenForm("F_SystemSettings")
Start-Sleep -Seconds 2
$h = [WinCap5]::Find([uint32]$p.Id, "INVENTORY SYSTEM")
if ($h -ne [IntPtr]::Zero) { [WinCap5]::Capture($h, "C:\Users\user\Downloads\LLC\_work\screen_settings.png"); Write-Host "captured settings" }

# invoke the drill-down exactly as the label click does
$acc.Eval('UX_OpenAreaList("customer",True)') | Out-Null
Start-Sleep -Seconds 2
$cl = $acc.Forms.Item("F_CustomerList")
Write-Host ("list filter = " + $cl.Filter + " on=" + $cl.FilterOn)
Write-Host ("list rows = " + $cl.RecordsetClone.RecordCount)
$h = [WinCap5]::Find([uint32]$p.Id, "INVENTORY SYSTEM")
if ($h -ne [IntPtr]::Zero) { [WinCap5]::Capture($h, "C:\Users\user\Downloads\LLC\_work\screen_customerlist.png"); Write-Host "captured list" }

# open the setup form for the incomplete customer
$acc.DoCmd.OpenForm("F_CustomerSetup", 0, [System.Reflection.Missing]::Value, "CustomerName='TEST CUSTOMER'")
Start-Sleep -Seconds 2
$h = [WinCap5]::Find([uint32]$p.Id, "F_CustomerSetup")
if ($h -eq [IntPtr]::Zero) { $h = [WinCap5]::Find([uint32]$p.Id, "INVENTORY SYSTEM") }
if ($h -ne [IntPtr]::Zero) { [WinCap5]::Capture($h, "C:\Users\user\Downloads\LLC\_work\screen_customersetup.png"); Write-Host "captured setup" }

$cs = $acc.Forms.Item("F_CustomerSetup")
Write-Host ("phone back = " + $cs.Controls.Item("txtPhone").BackColor + " address back = " + $cs.Controls.Item("txtBillingAddress").BackColor)
Write-Host ("name back = " + $cs.Controls.Item("txtCustomerName").BackColor)

$acc.CurrentDb().Execute("DELETE FROM T_Customers WHERE CustomerName='TEST CUSTOMER'") | Out-Null
$acc.CloseCurrentDatabase()
$acc.Quit()
Write-Host "DONE"
