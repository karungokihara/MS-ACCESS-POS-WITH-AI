$ErrorActionPreference="Continue"
$dbPath="C:\Users\user\Downloads\LLC\LLC_v3_2026-10-29.accdb"
Get-Process MSACCESS -ErrorAction SilentlyContinue | Stop-Process -Force -ErrorAction SilentlyContinue
Start-Sleep -Seconds 1
$acc=New-Object -ComObject Access.Application
$acc.Visible=$false; $acc.AutomationSecurity=1
$acc.OpenCurrentDatabase($dbPath,$false)
$cm=$acc.VBE.VBProjects.Item(1).VBComponents.Item("Form_F_ItemSetup").CodeModule
Write-Host $cm.Lines(1020, 135)
$acc.CloseCurrentDatabase(); $acc.Quit()
