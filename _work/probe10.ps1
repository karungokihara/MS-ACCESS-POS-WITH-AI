$ErrorActionPreference="Continue"
$dbPath="C:\Users\user\Downloads\LLC\LLC_v3_2026-10-29.accdb"
Get-Process MSACCESS -ErrorAction SilentlyContinue | Stop-Process -Force -ErrorAction SilentlyContinue
Start-Sleep -Seconds 1
$acc=New-Object -ComObject Access.Application
$acc.Visible=$false; $acc.AutomationSecurity=1
$acc.OpenCurrentDatabase($dbPath,$false)
$comp=$acc.VBE.VBProjects.Item(1).VBComponents.Item("Form_F_UnitList")
Write-Host "LINES:" $comp.CodeModule.CountOfLines
Write-Host $comp.CodeModule.Lines(1,$comp.CodeModule.CountOfLines)
$acc.CloseCurrentDatabase(); $acc.Quit()
Write-Host "DONE"
