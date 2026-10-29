$ErrorActionPreference = "Stop"
$root = "C:\Users\user\Downloads\LLC"
$out = "$root\_inspect\verify"
New-Item -ItemType Directory -Force -Path $out | Out-Null
Get-Process MSACCESS -ErrorAction SilentlyContinue | Stop-Process -Force -ErrorAction SilentlyContinue
Start-Sleep -Seconds 1
$acc = New-Object -ComObject Access.Application
$acc.Visible = $false
$acc.OpenCurrentDatabase("$root\LLC.accdb")
$acc.DoCmd.SetWarnings($false)
$db = $acc.CurrentDb()
$report = New-Object System.Collections.Generic.List[string]
function Note($m) { $report.Add($m); Write-Host $m }

$tables = @(
  "T_Users","T_Roles","T_Currencies","T_ItemTypes","T_Units","T_Customers",
  "T_Categories","T_Warehouses","T_Locations","T_Items","T_StockBalances",
  "T_Invoices","T_InvoiceLines","T_SalesOrders","T_SalesOrderLines",
  "T_CustomerPOs","T_Projects","T_Suppliers","T_Manufacturers","T_DocumentNumbers"
)
foreach ($t in $tables) {
  try {
    $rs = $db.OpenRecordset("SELECT Count(*) AS N FROM [$t]")
    Note ("COUNT $t = " + $rs.Fields.Item("N").Value)
    $rs.Close()
  } catch { Note "COUNT $t FAIL $($_.Exception.Message)" }
}
$rs = $db.OpenRecordset("SELECT Username FROM T_Users")
$names = @()
while (-not $rs.EOF) { $names += [string]$rs.Fields.Item("Username").Value; $rs.MoveNext() }
$rs.Close()
Note ("USERS " + ($names -join ", "))
$rs = $db.OpenRecordset("SELECT CustomerCode, CustomerName, CustomerType FROM T_Customers")
while (-not $rs.EOF) {
  Note ("CUSTOMER " + $rs.Fields.Item("CustomerCode").Value + " / " + $rs.Fields.Item("CustomerName").Value + " / " + $rs.Fields.Item("CustomerType").Value)
  $rs.MoveNext()
}
$rs.Close()
$rs = $db.OpenRecordset("SELECT UnitCode, UnitName FROM T_Units")
while (-not $rs.EOF) {
  Note ("UNIT " + $rs.Fields.Item("UnitCode").Value + " / " + $rs.Fields.Item("UnitName").Value)
  $rs.MoveNext()
}
$rs.Close()
$rs = $db.OpenRecordset("SELECT DocumentType, Prefix, NextNumber FROM T_DocumentNumbers")
while (-not $rs.EOF) {
  Note ("DOC " + $rs.Fields.Item("DocumentType").Value + " / " + $rs.Fields.Item("Prefix").Value + " = " + $rs.Fields.Item("NextNumber").Value)
  $rs.MoveNext()
}
$rs.Close()

$forms = @(
  "F_SalesOrders","F_Invoices","F_CustomerPO","F_SendPO","F_StockIn","F_StockOut",
  "F_ItemSetup","F_Warehouses","F_LocationSetup","F_Projects","F_ProjectIssue",
  "F_InvoiceList","SF_InvoiceList","SF_SalesOrderLines","SF_InvoiceLines","Dashboard",
  "F_SalesOrderList","SF_ItemSetupList"
)
foreach ($f in $forms) {
  $acc.SaveAsText(2, $f, "$out\$f.txt")
  Note "EXPORTED $f"
}
$mods = @("M_UX")
foreach ($m in $mods) { $acc.SaveAsText(5, $m, "$out\$m.txt"); Note "EXPORTED $m" }
$queries = @(
  "Q_InventoryListAll","Q_ItemSetupList","Q_SalesOrderReportHeader","Q_DeliveryNoteReportHeader",
  "Q_CustomerPOReportHeader","Q_SupplierPOReportHeader"
)
foreach ($q in $queries) { $acc.SaveAsText(1, $q, "$out\$q.txt"); Note "EXPORTED $q" }
$reports = @("R_Invoice","R_Quotation","R_SalesOrder","R_PurchaseOrder","R_SupplierPO","R_DeliveryNote")
foreach ($rp in $reports) {
  try { $acc.SaveAsText(3, $rp, "$out\$rp.txt"); Note "EXPORTED $rp" }
  catch { Note "REPORT MISSING $rp $($_.Exception.Message)" }
}

$acc.CloseCurrentDatabase()
$acc.Quit()
$report | Set-Content "$out\summary.txt" -Encoding UTF8
