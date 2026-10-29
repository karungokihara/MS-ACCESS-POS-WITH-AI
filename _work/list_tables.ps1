$ErrorActionPreference="Continue"
$dbPath="C:\Users\user\Downloads\LLC\LLC_v1_2026-10-29.accdb"
Get-Process MSACCESS -ErrorAction SilentlyContinue | Stop-Process -Force -ErrorAction SilentlyContinue
Start-Sleep -Seconds 1
$acc=New-Object -ComObject Access.Application
$acc.Visible=$false; $acc.AutomationSecurity=1
$acc.OpenCurrentDatabase($dbPath,$false)
$names=@()
foreach($t in $acc.CurrentDb().TableDefs){ if($t.Name -notlike "MSys*" -and $t.Name -notlike "~*"){ $names += $t.Name } }
$names | Sort-Object | ForEach-Object { Write-Host $_ }
Write-Host ("TOTAL=" + $names.Count)
$acc.CloseCurrentDatabase(); $acc.Quit()
