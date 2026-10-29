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
$f.Section(1).Height=900
foreach($c in $f.Controls){
  if($c.Name -like "hdr*"){ $c.Top=560 }
}
$acc.DoCmd.Save(2,"F_UnitList")
$acc.DoCmd.Close(2,"F_UnitList",1)
$acc.DoCmd.OpenForm("F_UnitList",1)
Start-Sleep -Milliseconds 300
$f=$acc.Forms.Item("F_UnitList")
Write-Host ("HeaderH=" + $f.Section(1).Height)
foreach($c in $f.Controls){ if($c.Name -like "hdr*"){ Write-Host ("{0} T={1}" -f $c.Name,$c.Top) } }
$acc.DoCmd.Close(2,"F_UnitList",2)
$acc.CloseCurrentDatabase(); $acc.Quit()
Write-Host "DONE"
