$ErrorActionPreference = "Stop"
function Log($m) { Write-Host $m }
function Try-Set($obj, $name, $value) { try { $obj.$name = $value } catch {} }
$dbPath = "C:\Users\user\Downloads\LLC\LLC.accdb"
$backup = "C:\Users\user\Downloads\LLC\LLC_before_clear_20261012.accdb"
$tmp = "C:\Users\user\Downloads\LLC\LLC_compact_tmp.accdb"

Get-Process MSACCESS -ErrorAction SilentlyContinue | Stop-Process -Force -ErrorAction SilentlyContinue
Start-Sleep -Seconds 2
if (-not (Test-Path $backup)) { Copy-Item $dbPath $backup }
Log "Backup ready"

$acc = New-Object -ComObject Access.Application
$acc.Visible = $false
$acc.OpenCurrentDatabase($dbPath)
$acc.DoCmd.SetWarnings($false)
try {
    $cards = @(
        @("txtActiveProjets",       "Label0",  "F_ProjectList",    "Open Projects"),
        @("txtInventoryQty",        "Label1",  "F_InventoryList",  "Open Inventory"),
        @("txtTotalItems",          "Label4",  "F_ItemSetupList",  "Open Items"),
        @("txtLowStockItems",       "Label2",  "F_InventoryList",  "Open Inventory"),
        @("txtTotalTools",          "Label3",  "F_ItemSetupList",  "Open Items"),
        @("txtActiveSaleOrders",    "Label6",  "F_SalesOrderList", "Open Sales Orders"),
        @("txtReadyToShipProjects", "Label8",  "F_ProjectList",    "Open Projects"),
        @("txtSentInvoices",        "Label10", "F_InvoiceList",    "Open Invoices"),
        @("txtPaidInvoices",        "Label12", "F_InvoiceList",    "Open Invoices")
    )
    $acc.DoCmd.OpenForm("Dashboard", 1)
    $d = $acc.Forms.Item("Dashboard")
    foreach ($c in $cards) {
        $action = "=UX_ShowModule(""" + $c[2] + """)"
        $t = $d.Controls.Item($c[0])
        $t.OnClick = $action
        Try-Set $t "Locked" $true
        Try-Set $t "TabStop" $false
        Try-Set $t "FontUnderline" $true
        Try-Set $t "ControlTipText" ($c[3] + " (click)")
        $l = $d.Controls.Item($c[1])
        $l.OnClick = $action
        Try-Set $l "ControlTipText" ($c[3] + " (click)")
        Log ("Card " + $c[0] + " -> " + $c[2])
    }
    $acc.DoCmd.Close(2, "Dashboard", 1)

    $db = $acc.CurrentDb()
    $clear = @(
        "T_AssemblyJobMaterials","T_AssemblyJobTools","T_AssemblyTemplateMaterials","T_AssemblyTemplateSteps","T_AssemblyTemplateTools",
        "T_InvoiceLines","T_GoodsReceiptItems","T_TransferItems","T_PurchaseOrderItems","T_SupplierPOLines","T_CustomerPOReleases",
        "T_StockMovements","T_StockBalances","T_ItemInstances","T_ItemPartNumbers","T_ItemSuppliers","T_SalesOrderLines","T_CustomerPOLines",
        "T_Invoices","T_GoodsReceipts","T_Transfers","T_PurchaseOrders","T_SupplierPOs","T_SalesOrders","T_CustomerPOs",
        "T_AssemblyJobs","T_AssemblyTemplates","T_WorkInstructions","T_ProjectIssueStaging","T_DocumentAttachments",
        "T_Projects","T_Items","T_Categories","T_Locations","T_Warehouses","T_Customers","T_Suppliers","T_Manufacturers",
        "T_Expenses","T_Attendance","T_Leave","T_Staff"
    )
    for ($pass = 1; $pass -le 5; $pass++) {
        $left = 0
        foreach ($t in $clear) {
            try { $db.Execute("DELETE FROM [" + $t + "]", 128) } catch {}
            $rs = $db.OpenRecordset("SELECT COUNT(*) FROM [" + $t + "]")
            $n = $rs.Fields.Item(0).Value
            $rs.Close()
            $left += $n
        }
        Log ("Pass " + $pass + " rows left: " + $left)
        if ($left -eq 0) { break }
    }
    $db.Execute("DELETE FROM T_ItemTypes WHERE ItemTypeName IN ('testing new again','tool1')", 128)
    $db.Execute("UPDATE T_DocumentNumbers SET NextNumber=1, LastUpdatedOn=Now()", 128)
    foreach ($t in $clear) {
        $rs = $db.OpenRecordset("SELECT COUNT(*) FROM [" + $t + "]")
        $n = $rs.Fields.Item(0).Value
        $rs.Close()
        if ($n -ne 0) { Log ("NOT EMPTY " + $t + " = " + $n) }
    }

    $acc.DoCmd.RunCommand(126)
    $acc.CloseCurrentDatabase()
    if (Test-Path $tmp) { Remove-Item $tmp -Force }
    $ok = $acc.CompactRepair($dbPath, $tmp, $false)
    $acc.Quit()
    Start-Sleep -Seconds 2
    if ($ok -and (Test-Path $tmp)) {
        Remove-Item $dbPath -Force
        Move-Item $tmp $dbPath
        Log "Compacted"
    } else {
        Log "Compact skipped"
    }
    Log "DONE"
} catch {
    Log ("FATAL " + $_.Exception.Message)
    try { $acc.Quit() } catch {}
    throw
}
