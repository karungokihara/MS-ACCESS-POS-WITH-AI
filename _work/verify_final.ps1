$ErrorActionPreference="Continue"
$dbPath="C:\Users\user\Downloads\LLC\LLC_v3_2026-10-29.accdb"
Get-Process MSACCESS -ErrorAction SilentlyContinue | Stop-Process -Force -ErrorAction SilentlyContinue
Start-Sleep -Seconds 1
$acc=New-Object -ComObject Access.Application
$acc.Visible=$false; $acc.AutomationSecurity=1
$acc.OpenCurrentDatabase($dbPath,$false)
$acc.RunCommand(126)
Write-Host ("compiled=" + $acc.IsCompiled)
$names=@()
foreach($c in $acc.VBE.VBProjects.Item(1).VBComponents){ $names += $c.Name }
Write-Host ("modules=" + $names.Count + " hasUXT=" + ($names -contains "UXT_Test"))
$loc=$acc.VBE.VBProjects.Item(1).VBComponents.Item("Form_F_LocationSetup").CodeModule
$t=$loc.Lines(1,$loc.CountOfLines)
Write-Host ("loc has TempVars=" + $t.Contains("TempVars"))
$it=$acc.VBE.VBProjects.Item(1).VBComponents.Item("Form_F_ItemSetup").CodeModule
$itText=$it.Lines(1,$it.CountOfLines)
Write-Host ("item has MsgBox confirm=" + $itText.Contains("Would you like to create it now"))
Write-Host ("forms=" + $acc.CurrentProject.AllForms.Count + " reports=" + $acc.CurrentProject.AllReports.Count + " queries=" + $acc.CurrentDb().QueryDefs.Count)
Write-Host ("units=" + $acc.DCount("*","T_Units") + " customers=" + $acc.DCount("*","T_Customers") + " locations=" + $acc.DCount("*","T_Locations"))
$acc.CloseCurrentDatabase(); $acc.Quit()
Write-Host "DONE"
