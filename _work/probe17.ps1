$ErrorActionPreference="Continue"
$dbPath="C:\Users\user\Downloads\LLC\LLC_v3_2026-10-29.accdb"
Get-Process MSACCESS -ErrorAction SilentlyContinue | Stop-Process -Force -ErrorAction SilentlyContinue
Start-Sleep -Seconds 1
$acc=New-Object -ComObject Access.Application
$acc.Visible=$false; $acc.AutomationSecurity=1
$acc.OpenCurrentDatabase($dbPath,$false)
$subs=@("SF_CustomerList","SF_SupplierList","SF_ItemSetupList","SF_WarehouseList","SF_ProjectList","SF_StockMovementList","SF_InventoryList","SF_CustomerPOList","SF_SalesOrderList","SF_InvoiceList","SF_TemplateList")
foreach($s in $subs){
  try {
    $acc.DoCmd.OpenForm($s,1)
    Start-Sleep -Milliseconds 200
    $f=$acc.Forms.Item($s)
    Write-Host ("{0,-24} ScrollBars={1} DefaultView={2} InsideWidth={3}" -f $s,$f.ScrollBars,$f.DefaultView,$f.InsideWidth)
    $acc.DoCmd.Close(2,$s,2)
  } catch { Write-Host ($s + " ERR " + $_.Exception.Message) }
}
$lists=@("F_CustomerList","F_SupplierList","F_ItemSetupList","F_WarehouseList","F_ProjectList","F_StockMovementList","F_InventoryList")
foreach($l in $lists){
  try {
    $acc.DoCmd.OpenForm($l,1)
    Start-Sleep -Milliseconds 200
    $f=$acc.Forms.Item($l)
    $sf=""
    foreach($c in $f.Controls){ if($c.ControlType -eq 112){ $sf = $sf + $c.Name + "->" + $c.SourceObject + " W=" + $c.Width + "; " } }
    Write-Host ("{0,-24} ScrollBars={1} subforms: {2}" -f $l,$f.ScrollBars,$sf)
    $acc.DoCmd.Close(2,$l,2)
  } catch { Write-Host ($l + " ERR " + $_.Exception.Message) }
}
$acc.CloseCurrentDatabase(); $acc.Quit()
Write-Host "DONE"
