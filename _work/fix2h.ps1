$ErrorActionPreference = "Continue"
$dbPath = "C:\Users\user\Downloads\LLC\LLC_v3_2026-10-29.accdb"
Get-Process MSACCESS -ErrorAction SilentlyContinue | Stop-Process -Force -ErrorAction SilentlyContinue
Start-Sleep -Seconds 1

$acc = New-Object -ComObject Access.Application
$acc.Visible = $false
$acc.AutomationSecurity = 1
$acc.OpenCurrentDatabase($dbPath, $false)
$acc.DoCmd.OpenForm("F_UnitList", 1)
Start-Sleep -Milliseconds 500
$f = $acc.Forms.Item("F_UnitList")

$f.ScrollBars = 2
$f.AutoResize = $true
$f.AutoCenter = $true
$f.Section(1).Height = 700
$f.Section(0).Height = 420
$f.Width = 15300

foreach ($c in $f.Controls) {
    if ($c.Name -like "hdr*") {
        $c.Top = 380
        $c.Height = 260
        $c.FontName = "Segoe UI Semibold"
        $c.FontSize = 9.5
        $c.FontBold = $true
    }
    if ($c.Name -like "txt*") { $c.Top = 60; $c.Height = 330 }
    if ($c.Name -like "cmd*" -and $c.ControlType -eq 104) { $c.Top = 60; $c.Height = 340 }
}

$acc.DoCmd.Save(2, "F_UnitList")
$acc.DoCmd.Close(2, "F_UnitList", 1)

$f2 = $acc.Forms.Item("F_UnitList")
Write-Host ("saved ScrollBars=" + $f2.ScrollBars + " HeaderH=" + $f2.Section(1).Height + " Width=" + $f2.Width)
$acc.DoCmd.Close(2, "F_UnitList", 2)

$acc.CloseCurrentDatabase()
$acc.Quit()
Write-Host "DONE"
