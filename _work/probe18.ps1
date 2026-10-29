$ErrorActionPreference="Continue"
$dbPath="C:\Users\user\Downloads\LLC\_work\test_customer.accdb"
Get-Process MSACCESS -ErrorAction SilentlyContinue | Stop-Process -Force -ErrorAction SilentlyContinue
Start-Sleep -Seconds 1
$acc=New-Object -ComObject Access.Application
$acc.Visible=$true; $acc.AutomationSecurity=1
$acc.OpenCurrentDatabase($dbPath,$false)
$acc.CurrentDb().Execute("DELETE FROM T_Customers WHERE CustomerName='TEST CUSTOMER'") | Out-Null
$acc.CurrentDb().Execute("INSERT INTO T_Customers (CustomerName, CustomerType, Status) VALUES ('TEST CUSTOMER','Retail','In Review')") | Out-Null
$acc.DoCmd.OpenForm("F_SystemSettings")
Start-Sleep -Seconds 2
$f=$acc.Forms.Item("F_SystemSettings")
Write-Host ("controls=" + $f.Controls.Count)
foreach($n in @("lblCustomerStatus","lblSupplierStatus","lblCategoryStatus","lblUnitStatus","lblItemStatus","lblTemplateStatus","lblWarehouseStatus","lblLocationStatus")){
  try { $c=$f.Controls.Item($n); Write-Host ("{0,-20} caption='{1}' onclick='{2}'" -f $n,$c.Caption,$c.OnClick) } catch { Write-Host ($n + " ERR") }
}
$acc.Eval('UX_OpenAreaList("customer",True)') | Out-Null
Start-Sleep -Seconds 1
$cl=$acc.Forms.Item("F_CustomerList")
$sub=$cl.Controls.Item("subCustomerList").Form
Write-Host ("sub filter='" + $sub.Filter + "' on=" + $sub.FilterOn)
Write-Host ("sub records=" + $sub.RecordsetClone.RecordCount)
$acc.CurrentDb().Execute("DELETE FROM T_Customers WHERE CustomerName='TEST CUSTOMER'") | Out-Null
$acc.CloseCurrentDatabase(); $acc.Quit()
Write-Host "DONE"
