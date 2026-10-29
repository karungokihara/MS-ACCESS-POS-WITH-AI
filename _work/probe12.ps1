$ErrorActionPreference="Continue"
$dbPath="C:\Users\user\Downloads\LLC\LLC_v3_2026-10-29.accdb"
Get-Process MSACCESS -ErrorAction SilentlyContinue | Stop-Process -Force -ErrorAction SilentlyContinue
Start-Sleep -Seconds 1
$acc=New-Object -ComObject Access.Application
$acc.Visible=$false; $acc.AutomationSecurity=1
$acc.OpenCurrentDatabase($dbPath,$false)
$acc.DoCmd.OpenForm("F_UnitList",1)
Start-Sleep -Milliseconds 400
$f=$acc.Forms.Item("F_UnitList")
Write-Host ("ScrollBars=" + $f.ScrollBars + " HeaderH=" + $f.Section(1).Height + " DetailH=" + $f.Section(0).Height + " Width=" + $f.Width + " AutoResize=" + $f.AutoResize + " AutoCenter=" + $f.AutoCenter)
foreach($c in $f.Controls){ if($c.Name -like "hdr*"){ Write-Host ("{0} L={1} T={2} W={3} H={4}" -f $c.Name,$c.Left,$c.Top,$c.Width,$c.Height) } }
$acc.DoCmd.Close(2,"F_UnitList",2)
$acc.CloseCurrentDatabase(); $acc.Quit()
