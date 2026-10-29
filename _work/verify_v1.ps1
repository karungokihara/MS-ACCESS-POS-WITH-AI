$ErrorActionPreference="Continue"
$dbPath="C:\Users\user\Downloads\LLC\LLC_v1_2026-10-29.accdb"
Get-Process MSACCESS -ErrorAction SilentlyContinue | Stop-Process -Force -ErrorAction SilentlyContinue
Start-Sleep -Seconds 1
$acc=New-Object -ComObject Access.Application
$acc.Visible=$false; $acc.AutomationSecurity=1
$acc.OpenCurrentDatabase($dbPath,$false)
$acc.RunCommand(126)
Write-Host ("compiled=" + $acc.IsCompiled)
Write-Host ("forms=" + $acc.CurrentProject.AllForms.Count + " modules=" + $acc.VBE.VBProjects.Item(1).VBComponents.Count + " queries=" + $acc.CurrentDb().QueryDefs.Count)
Write-Host ("units=" + $acc.DCount("*","T_Units") + " unitReady=" + $acc.Eval('UX_AreaReady("unit")') + " qRows=" + $acc.DCount("*","Q_UnitList"))
Write-Host ("statuses=" + $acc.Eval('UX_StatusList()'))
Write-Host ("backupBtns=" + ($acc.CurrentProject.AllForms("F_SystemSettings").Name))
$acc.CloseCurrentDatabase(); $acc.Quit()
Write-Host "DONE"
