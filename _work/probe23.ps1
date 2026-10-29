$ErrorActionPreference="Continue"
$dbPath="C:\Users\user\Downloads\LLC\_work\test_itemloc.accdb"
Get-Process MSACCESS -ErrorAction SilentlyContinue | Stop-Process -Force -ErrorAction SilentlyContinue
Start-Sleep -Seconds 1
$acc=New-Object -ComObject Access.Application
$acc.Visible=$true; $acc.AutomationSecurity=1
$acc.OpenCurrentDatabase($dbPath,$false)
$cm=$acc.VBE.VBProjects.Item(1).VBComponents.Item("Form_F_LocationSetup").CodeModule
Write-Host "--- load lines 40..85"
Write-Host $cm.Lines(40,46)
$acc.Eval('TempVars.Add "UXT_AutoSave", True') | Out-Null
$acc.DoCmd.OpenForm("F_LocationSetup", 0, [System.Reflection.Missing]::Value, [System.Reflection.Missing]::Value, 1, 0, "QUICKADD|TEST-ITEM-LOC")
Start-Sleep -Seconds 2
Write-Host ("loaded=" + $acc.CurrentProject.AllForms("F_LocationSetup").IsLoaded)
$acc.Eval('Empty_Noop') 2>$null
$vals=$acc.Eval('UXT_ProbeLocation()')
Write-Host ("PROBE: " + $vals)
$acc.CloseCurrentDatabase(); $acc.Quit()
