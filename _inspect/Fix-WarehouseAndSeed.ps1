$ErrorActionPreference = "Stop"
$root = "C:\Users\user\Downloads\LLC"
function Log($msg) {
    $line = "$(Get-Date -Format HH:mm:ss) $msg"
    Add-Content "$root\_inspect\fix-warehouse-log.txt" $line
    Write-Host $line
}
"" | Set-Content "$root\_inspect\fix-warehouse-log.txt" -Encoding UTF8

Get-Process MSACCESS -ErrorAction SilentlyContinue | Stop-Process -Force -ErrorAction SilentlyContinue
Start-Sleep -Seconds 1

$acc = New-Object -ComObject Access.Application
$acc.Visible = $false
$acc.OpenCurrentDatabase("$root\LLC.accdb")
$acc.DoCmd.SetWarnings($false)

try {
    # ----- Warehouse form labels, matching the working layout -----
    Log "Fixing warehouse form"
    while ($acc.CurrentProject.AllForms.Item("F_Warehouses").IsLoaded) {
        $acc.DoCmd.Close(2, "F_Warehouses", 2)
    }
    $acc.DoCmd.OpenForm("F_Warehouses", 1)
    $frm = $acc.Forms.Item("F_Warehouses")
    $captions = @{
        "WarehouseCode_Label" = "WarehouseCode"
        "WarehouseName_Label" = "WarehouseName"
        "Address_Label" = "Address"
        "ContactPerson_Label" = "ContactPerson"
        "SF_WarehouseLocations Label" = "Warehouse Locations"
        "Label10" = "Warehouses"
        "Command14" = "Save"
        "Command15" = "Close"
        "Command16" = "Add New"
    }
    foreach ($name in $captions.Keys) {
        $frm.Controls.Item($name).Caption = $captions[$name]
        Log "Caption $name = $($captions[$name])"
    }
    $acc.DoCmd.Close(2, "F_Warehouses", 1)

    # ----- Location rows can be typed on the warehouse form -----
    Log "Fixing warehouse locations subform"
    while ($acc.CurrentProject.AllForms.Item("SF_WarehouseLocations").IsLoaded) {
        $acc.DoCmd.Close(2, "SF_WarehouseLocations", 2)
    }
    $acc.DoCmd.OpenForm("SF_WarehouseLocations", 1)
    $sub = $acc.Forms.Item("SF_WarehouseLocations")
    $sub.RecordSource = "SELECT * FROM T_Locations"
    foreach ($p in $sub.Properties) {
        if ($p.Name -eq "HasModule") { $p.Value = $true }
        if ($p.Name -eq "AllowAdditions") { $p.Value = $true }
        if ($p.Name -eq "AllowEdits") { $p.Value = $true }
    }
    $mod = $sub.Module
    $has = $false
    for ($i = 1; $i -le $mod.CountOfLines; $i++) {
        if (([string]$mod.Lines($i, 1)).Contains("UX-LOCATION-ROW")) { $has = $true }
    }
    if (-not $has) {
        $mod.InsertLines($mod.CountOfLines + 1, @'
' UX-LOCATION-ROW
Private Sub Form_BeforeInsert(Cancel As Integer)
    On Error Resume Next
    If Nz(Me!WarehouseID, 0) = 0 Then
        Me!WarehouseID = Nz(Me.Parent!WarehouseID, 0)
    End If
End Sub

Private Sub Form_BeforeUpdate(Cancel As Integer)
    On Error GoTo Err_Handler
    If Len(Trim$(Nz(Me!LocationName, ""))) = 0 Then
        MsgBox "Type the location name in this row.", vbExclamation, "Location"
        Cancel = True
        Exit Sub
    End If
    If Len(Trim$(Nz(Me!LocationCode, ""))) = 0 Then
        Me!LocationCode = UCase$(Left$(Replace(Trim$(Me!LocationName), " ", ""), 8))
    End If
    If Nz(Me!WarehouseID, 0) = 0 Then
        Me!WarehouseID = Nz(Me.Parent!WarehouseID, 0)
    End If
    Exit Sub
Err_Handler:
    Cancel = True
    MsgBox Err.Description, vbExclamation, "Location"
End Sub
'@)
    }
    $acc.DoCmd.Close(2, "SF_WarehouseLocations", 1)
    Log "Location subform saved"

    # ----- Item review list query (simpler, so the list can open) -----
    Log "Fixing item setup list query"
    $sql = @"
SELECT I.ItemID, I.InternalSKU, I.ItemName, I.Description, I.CategoryID, C.CategoryName, IIf(Nz(I.RequiresSerial,0)=0,'Bulk','Serialized') AS StockType, I.BaseUnitID, IIf(Len(Trim(Nz(U.UnitCode,'')))=0,Nz(U.UnitName,''),U.UnitCode) AS UnitCode, Nz(B.QtyOnHand,0) AS QuantityOnHand, I.ReorderLevel, I.DefaultLocationID, L.LocationCode AS DefaultLocationCode, L.LocationName AS DefaultLocationName, W.WarehouseName AS DefaultWarehouseName, I.CreatedOn, I.UpdatedOn
FROM ((((T_Items AS I LEFT JOIN T_Categories AS C ON I.CategoryID = C.CategoryID) LEFT JOIN T_Units AS U ON I.BaseUnitID = U.UnitID) LEFT JOIN T_Locations AS L ON I.DefaultLocationID = L.LocationID) LEFT JOIN T_Warehouses AS W ON L.WarehouseID = W.WarehouseID) LEFT JOIN (SELECT ItemID, Sum(Nz(QuantityOnHand,0)) AS QtyOnHand FROM T_StockBalances GROUP BY ItemID) AS B ON I.ItemID = B.ItemID
ORDER BY I.ItemName, I.InternalSKU;
"@
    $db = $acc.CurrentDb()
    $found = $false
    foreach ($q in $db.QueryDefs) {
        if ($q.Name -eq "Q_ItemSetupList") { $q.SQL = $sql; $found = $true }
    }
    if (-not $found) { [void]$db.CreateQueryDef("Q_ItemSetupList", $sql) }
    $rs = $db.OpenRecordset("SELECT TOP 1 ItemID, UnitCode, QuantityOnHand FROM Q_ItemSetupList")
    Log "Item query opens. EOF=$($rs.EOF)"
    $rs.Close()

    # ----- Clear business data -----
    Log "Clearing business data"
    $db = $acc.CurrentDb()
    $tables = @(
        "T_AssemblyJobMaterials","T_AssemblyJobTools","T_AssemblyJobs",
        "T_AssemblyTemplateSteps","T_AssemblyTemplateMaterials","T_AssemblyTemplateTools","T_AssemblyTemplates",
        "T_ProjectIssueStaging","T_GoodsReceiptItems","T_GoodsReceipts",
        "T_TransferItems","T_Transfers","T_InvoiceLines","T_Invoices",
        "T_SalesOrderLines","T_SalesOrders","T_CustomerPOReleases","T_CustomerPOLines","T_CustomerPOs",
        "T_SupplierPOLines","T_SupplierPOs","T_PurchaseOrderItems","T_PurchaseOrders",
        "T_StockMovements","T_StockBalances","T_ItemInstances","T_ItemPartNumbers","T_ItemSuppliers","T_Items",
        "T_Projects","T_DocumentAttachments","T_WorkInstructions","T_Locations","T_Warehouses",
        "T_Categories","T_Manufacturers","T_Suppliers","T_Customers","T_Units"
    )
    foreach ($pass in 1..4) {
        foreach ($t in $tables) {
            try { $db.Execute("DELETE FROM [$t]") } catch { Log "Delete $t pass $pass : $($_.Exception.Message)" }
        }
    }
    try { $db.Execute("UPDATE T_DocumentNumbers SET NextNumber=1") } catch { Log "Doc numbers: $($_.Exception.Message)" }

    Log "Seeding one sample of each main record"
    $acc.Run("UX_EnsureStandardUnits")
    $acc.Run("UX_EnsureInternalCustomer")
    $acc.Run("UX_BackfillUnitCodes")
    $db = $acc.CurrentDb()

    $db.Execute("UPDATE T_Customers SET Phone='000', Email='internal@local', BillingAddress='Internal', ShippingAddress='Internal', BillingName='Internal Customer', ContactPerson='Internal Test', Status='Active' WHERE CustomerCode='INTERNAL'")

    $db.Execute("INSERT INTO T_Categories (ItemTypeID, CategoryName, Description, RequiresSerial, RequiresExpiry) SELECT TOP 1 ItemTypeID, 'General', 'General stock items', 0, 0 FROM T_ItemTypes")
    $db.Execute("INSERT INTO T_Warehouses (WarehouseCode, WarehouseName, Address, ContactPerson) VALUES ('MAIN-001', 'Main Warehouse', 'Main site', 'Store')")
    $db.Execute("INSERT INTO T_Locations (WarehouseID, LocationCode, LocationName, BinNumber, ShelfNumber) SELECT WarehouseID, 'MAIN', 'Main Store', 'A1', '1' FROM T_Warehouses WHERE WarehouseCode='MAIN-001'")
    $db.Execute("INSERT INTO T_Manufacturers (ManufacturerName, Website, Notes) VALUES ('Sample Manufacturer', '', 'Sample manufacturer')")
    $db.Execute("INSERT INTO T_Suppliers (SupplierCode, CompanyName, ContactPerson, Phone, Email, Address, PaymentTerms) VALUES ('SUP-001', 'Sample Supplier', 'Sales Desk', '000', 'supplier@local', 'Supplier address', '30 days')")
    $db.Execute("INSERT INTO T_Customers (CustomerCode, CustomerName, CustomerType, ContactPerson, Phone, Email, BillingAddress, ShippingAddress, BillingName, Status, Notes, CreatedOn, UpdatedOn) VALUES ('SAMPLE', 'Sample Customer', 'Customer', 'Buyer', '000', 'buyer@local', 'Billing address', 'Shipping address', 'Sample Customer', 'Active', 'Sample customer for normal sales.', Now(), Now())")

    $db.Execute(@"
INSERT INTO T_Items (InternalSKU, ItemName, Description, CategoryID, BaseUnitID, ReorderLevel, MinStockLevel, MaxStockLevel, DefaultLocationID, RequiresSerial, ItemClass, ItemType, IsSellable, DefaultSalesPrice, DefaultTaxable, CreatedOn, UpdatedOn)
SELECT 'ITEM-0001', 'Sample Product', 'Sample product for testing', C.CategoryID, U.UnitID, 5, 5, 100, L.LocationID, 0, 'Stock', 'Product', -1, 10, 0, Now(), Now()
FROM T_Categories AS C, T_Units AS U, T_Locations AS L
WHERE C.CategoryName='General' AND U.UnitCode='PCS' AND L.LocationCode='MAIN'
"@)

    $db.Execute("INSERT INTO T_StockBalances (ItemID, CurrentLocationID, QuantityOnHand, QuantityReserved, LastUpdated) SELECT I.ItemID, L.LocationID, 20, 0, Now() FROM T_Items AS I, T_Locations AS L WHERE I.InternalSKU='ITEM-0001' AND L.LocationCode='MAIN'")
    $db.Execute("INSERT INTO T_ItemSuppliers (ItemID, SupplierID, PreferredSupplier, LastUnitPrice, IsActive) SELECT I.ItemID, S.SupplierID, -1, 10, -1 FROM T_Items AS I, T_Suppliers AS S WHERE I.InternalSKU='ITEM-0001' AND S.SupplierCode='SUP-001'")
    $db.Execute("INSERT INTO T_Projects (ProjectCode, ProjectName, Status, IsActive, CustomerID, FinishedItemID, BuildQuantity, OutputLocationID, WorkflowStatus, StartDate) SELECT 'PRJ-001', 'Sample Project', 'Open', -1, CU.CustomerID, I.ItemID, 1, L.LocationID, 'Open', Date() FROM T_Customers AS CU, T_Items AS I, T_Locations AS L WHERE CU.CustomerCode='SAMPLE' AND I.InternalSKU='ITEM-0001' AND L.LocationCode='MAIN'")

    foreach ($t in @("T_Categories","T_Warehouses","T_Locations","T_Items","T_StockBalances","T_Customers","T_Suppliers","T_Manufacturers","T_Projects","T_Units")) {
        $c = $db.OpenRecordset("SELECT Count(*) AS N FROM [$t]")
        Log "COUNT $t = $($c.Fields.Item('N').Value)"
        $c.Close()
    }

    Log "Opening item setup list"
    $acc.DoCmd.OpenForm("F_ItemSetupList", 0)
    Start-Sleep -Seconds 1
    $acc.DoCmd.Close(2, "F_ItemSetupList", 2)
    Log "Item setup list opened"
    Log "DONE"
    $acc.CloseCurrentDatabase()
    $acc.Quit()
} catch {
    Log "FATAL $($_.Exception.Message)"
    try { $acc.Quit() } catch {}
    throw
}
