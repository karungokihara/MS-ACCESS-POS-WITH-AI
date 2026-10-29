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
$f.Controls.Item("Label12").Height=520
$f.Section(1).Height=700
$acc.DoCmd.Save(2,"F_UnitList")
$acc.DoCmd.Close(2,"F_UnitList",1)
$acc.DoCmd.OpenForm("F_UnitList",1)
Start-Sleep -Milliseconds 400
$f=$acc.Forms.Item("F_UnitList")
Write-Host ("HeaderH=" + $f.Section(1).Height + " TitleH=" + $f.Controls.Item("Label12").Height + " ScrollBars=" + $f.ScrollBars)
$acc.DoCmd.Close(2,"F_UnitList",2)
$acc.CloseCurrentDatabase(); $acc.Quit()
Write-Host "DONE"
