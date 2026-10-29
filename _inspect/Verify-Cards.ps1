Get-Process MSACCESS -ErrorAction SilentlyContinue | Stop-Process -Force -ErrorAction SilentlyContinue
Start-Sleep -Seconds 1
$acc = New-Object -ComObject Access.Application
$acc.Visible = $false
$acc.OpenCurrentDatabase("C:\Users\user\Downloads\LLC\LLC.accdb")
try { $acc.DoCmd.Close(2, "F_Login", 2) } catch {}
$acc.TempVars.Add("CurrentUserID", $acc.DLookup("UserID", "T_Users", "Username='admin'"))
$acc.TempVars.Add("CurrentRoleID", $acc.DLookup("RoleID", "T_Users", "Username='admin'"))
$acc.DoCmd.OpenForm("Dashboard", 0)
foreach ($n in @("txtActiveProjets","txtInventoryQty","txtTotalItems","txtLowStockItems","txtTotalTools","txtActiveSaleOrders","txtReadyToShipProjects","txtSentInvoices","txtPaidInvoices")) {
    $v = $acc.Eval("Forms!Dashboard!" + $n)
    $oc = $acc.Eval("Forms!Dashboard!" + $n + ".OnClick")
    Write-Host ($n + " value=" + $v + " click=" + $oc)
}
foreach ($f in @("F_ProjectList","F_InventoryList","F_ItemSetupList","F_SalesOrderList","F_InvoiceList")) {
    $null = $acc.Eval("UX_ShowModule(""" + $f + """)")
    Write-Host ("Panel shows " + $acc.Eval("Forms!Dashboard!subNavModule.SourceObject") + " visible=" + $acc.Eval("Forms!Dashboard!subNavModule.Visible"))
}
$null = $acc.Eval("UX_ShowHome()")
Write-Host ("Home again: card visible=" + $acc.Eval("Forms!Dashboard!txtActiveProjets.Visible"))
$acc.Quit(2)
