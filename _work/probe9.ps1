$ErrorActionPreference = "Continue"
$dbPath = "C:\Users\user\Downloads\LLC\LLC_v3_2026-10-29.accdb"
Get-Process MSACCESS -ErrorAction SilentlyContinue | Stop-Process -Force -ErrorAction SilentlyContinue
Start-Sleep -Seconds 1

$acc = New-Object -ComObject Access.Application
$acc.Visible = $false
$acc.AutomationSecurity = 1
$acc.OpenCurrentDatabase($dbPath, $false)
$acc.DoCmd.OpenForm("F_UnitList", 1)
Start-Sleep -Milliseconds 400
$f = $acc.Forms.Item("F_UnitList")

Write-Host ("DefaultView=" + $f.DefaultView + " ScrollBars=" + $f.ScrollBars)
Write-Host ("HeaderH=" + $f.Section(1).Height + " DetailH=" + $f.Section(0).Height + " InsideWidth=" + $f.InsideWidth)

foreach ($c in $f.Controls) {
    $ct = $c.ControlType
    $line = ("{0,-24} type={1,-4} L={2,-6} T={3,-6} W={4,-6} H={5,-6}" -f $c.Name, $ct, $c.Left, $c.Top, $c.Width, $c.Height)
    if ($ct -eq 100) { $line = $line + (" align=" + $c.TextAlign) }
    Write-Host $line
}

$acc.DoCmd.Close(2, "F_UnitList", 2)
$acc.CloseCurrentDatabase()
$acc.Quit()
Write-Host "DONE"
