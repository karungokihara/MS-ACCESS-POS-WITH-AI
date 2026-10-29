$ErrorActionPreference="Continue"
$dbPath="C:\Users\user\Downloads\LLC\LLC_v1_2026-10-29.accdb"
Get-Process MSACCESS -ErrorAction SilentlyContinue | Stop-Process -Force -ErrorAction SilentlyContinue
Start-Sleep -Seconds 1
$acc=New-Object -ComObject Access.Application
$acc.Visible=$false; $acc.AutomationSecurity=1
$acc.OpenCurrentDatabase($dbPath,$false)
Write-Host "FORMS:"
foreach($f in $acc.CurrentProject.AllForms){ Write-Host ("  " + $f.Name) }
Write-Host "REPORTS:"
foreach($r in $acc.CurrentProject.AllReports){ Write-Host ("  " + $r.Name) }
Write-Host "QUERIES:"
foreach($q in $acc.CurrentDb().QueryDefs){ Write-Host ("  " + $q.Name) }
$acc.CloseCurrentDatabase(); $acc.Quit()
