$ErrorActionPreference="Continue"
$dbPath="C:\Users\user\Downloads\LLC\LLC_v3_2026-10-29.accdb"
Get-Process MSACCESS -ErrorAction SilentlyContinue | Stop-Process -Force -ErrorAction SilentlyContinue
Start-Sleep -Seconds 1
$acc=New-Object -ComObject Access.Application
$acc.Visible=$false; $acc.AutomationSecurity=1
$acc.OpenCurrentDatabase($dbPath,$false)
$acc.DoCmd.OpenForm("F_LocationSetup",1)
Start-Sleep -Milliseconds 400
$f=$acc.Forms.Item("F_LocationSetup")
$map=@{
 "WarehouseID_Label"="Warehouse (optional)"
 "BinNumber_Label"="Bin (optional)"
 "ShelfNumber_Label"="Shelf (optional)"
 "LocationName_Label"="Location Name"
}
foreach($k in $map.Keys){
  try { $c=$f.Controls.Item($k); $c.Caption=$map[$k]; Write-Host ("set " + $k + " -> " + $map[$k]) } catch { Write-Host ("fail " + $k) }
}
$acc.DoCmd.Save(2,"F_LocationSetup")
$acc.DoCmd.Close(2,"F_LocationSetup",1)

Write-Host "== system settings buttons"
$acc.DoCmd.OpenForm("F_SystemSettings",1)
Start-Sleep -Milliseconds 400
$f=$acc.Forms.Item("F_SystemSettings")
foreach($c in $f.Controls){ if($c.ControlType -eq 104){ Write-Host ("{0} cap='{1}' onclick='{2}'" -f $c.Name,$c.Caption,$c.OnClick) } }
$acc.DoCmd.Close(2,"F_SystemSettings",2)

$acc.CloseCurrentDatabase(); $acc.Quit()
Write-Host "DONE"
