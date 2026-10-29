$ErrorActionPreference="Continue"
$dbPath="C:\Users\user\Downloads\LLC\LLC_v1_2026-10-29.accdb"
Get-Process MSACCESS -ErrorAction SilentlyContinue | Stop-Process -Force -ErrorAction SilentlyContinue
Start-Sleep -Seconds 1
$acc=New-Object -ComObject Access.Application
$acc.Visible=$true; $acc.AutomationSecurity=1
$acc.OpenCurrentDatabase($dbPath,$false)
$acc.Eval('UX_OpenAreaList("unit",False)') | Out-Null
Start-Sleep -Seconds 1
$f=$acc.Forms.Item("F_UnitList")
Write-Host ("unitlist loaded=" + $acc.CurrentProject.AllForms("F_UnitList").IsLoaded + " rows=" + $f.RecordsetClone.RecordCount + " status=" + [string]$f.Controls.Item("txtRecordStatus").Value)
$acc.DoCmd.Close(2,"F_UnitList",2)
$acc.Eval('UX_OpenAreaList("location",False)') | Out-Null
Start-Sleep -Seconds 1
Write-Host ("locationSetup loaded=" + $acc.CurrentProject.AllForms("F_LocationSetup").IsLoaded)
$acc.DoCmd.Close(2,"F_LocationSetup",2)
$acc.CloseCurrentDatabase(); $acc.Quit()
Write-Host "DONE"
