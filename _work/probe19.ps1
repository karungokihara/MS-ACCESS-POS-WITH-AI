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
foreach($n in @("cmdNewUnit","cmdEditUnit","cmdCloseUnit")){
  try { $c=$f.Controls.Item($n); Write-Host ("{0} cap='{1}' onclick='{2}'" -f $n,$c.Caption,$c.OnClick) } catch { Write-Host ($n + " ERR") }
}
$acc.DoCmd.Close(2,"F_UnitList",2)
$q=$acc.CurrentDb().QueryDefs.Item("Q_UnitList")
Write-Host ("Q_UnitList SQL: " + $q.SQL)
$acc.CloseCurrentDatabase(); $acc.Quit()
