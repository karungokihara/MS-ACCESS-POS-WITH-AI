param(
    [string]$DatabasePath = "C:\Users\user\Downloads\LLC\LLC_sample_reset_working.accdb"
)

$ErrorActionPreference = "Stop"
$root = "C:\Users\user\Downloads\LLC"
$expectedName = "LLC_sample_reset_working.accdb"
$fullPath = [System.IO.Path]::GetFullPath($DatabasePath)

if ([System.IO.Path]::GetFileName($fullPath) -cne $expectedName -or
    [System.IO.Path]::GetDirectoryName($fullPath) -ine $root) {
    throw "Refusing to reset a database other than the isolated working copy: $expectedName"
}

$backupPath = Join-Path $root "LLC_sample_reset_preseed_verified_20261028.accdb"
if (-not (Test-Path -LiteralPath $fullPath)) { throw "Database not found: $fullPath" }
if (-not (Test-Path -LiteralPath $backupPath)) { throw "Required pre-seed backup not found: $backupPath" }
if ((Get-FileHash -LiteralPath $fullPath -Algorithm SHA256).Hash -ne
    (Get-FileHash -LiteralPath $backupPath -Algorithm SHA256).Hash) {
    throw "Pre-seed backup does not match the isolated working copy."
}

$connectionString = "Provider=Microsoft.ACE.OLEDB.12.0;Data Source=$fullPath;Persist Security Info=False;"
$connection = New-Object System.Data.OleDb.OleDbConnection($connectionString)
$transaction = $null

function Invoke-Sql([string]$Sql) {
    $command = $connection.CreateCommand()
    $command.Transaction = $transaction
    $command.CommandText = $Sql
    [void]$command.ExecuteNonQuery()
    $command.Dispose()
}

function Get-Scalar([string]$Sql) {
    $command = $connection.CreateCommand()
    $command.Transaction = $transaction
    $command.CommandText = $Sql
    $value = $command.ExecuteScalar()
    $command.Dispose()
    return $value
}

try {
    $connection.Open()

    foreach ($check in @(
        @{ Sql = "SELECT Count(*) FROM T_ItemTypes"; Name = "item types" },
        @{ Sql = "SELECT Count(*) FROM T_Units WHERE UnitCode='PCS'"; Name = "PCS unit" },
        @{ Sql = "SELECT Count(*) FROM T_StockMovementTypes WHERE MovementTypeCode='STOCK_IN'"; Name = "STOCK_IN movement type" },
        @{ Sql = "SELECT Count(*) FROM T_PartNumberTypes"; Name = "part-number types" },
        @{ Sql = "SELECT Count(*) FROM T_Users"; Name = "users" }
    )) {
        if ([int](Get-Scalar $check.Sql) -lt 1) { throw "Required preserved setup is missing: $($check.Name)" }
    }

    $transaction = $connection.BeginTransaction()

    $clearTables = @(
        "T_ProjectIssueStaging", "T_StockMovements",
        "T_AssemblyJobMaterials", "T_AssemblyJobTools",
        "T_AssemblyJobs", "T_Projects",
        "T_AssemblyTemplateMaterials", "T_AssemblyTemplateSteps", "T_AssemblyTemplateTools",
        "T_InvoiceLines", "T_GoodsReceiptItems", "T_TransferItems", "T_PurchaseOrderItems",
        "T_SupplierPOLines", "T_CustomerPOReleases", "T_StockBalances",
        "T_ItemInstances", "T_ItemPartNumbers", "T_ItemSuppliers",
        "T_SalesOrderLines", "T_CustomerPOLines", "T_Invoices", "T_GoodsReceipts", "T_Transfers",
        "T_PurchaseOrders", "T_SupplierPOs", "T_SalesOrders", "T_CustomerPOs",
        "T_AssemblyTemplates", "T_WorkInstructions",
        "T_DocumentAttachments", "T_Items", "T_Locations", "T_Warehouses",
        "T_Categories", "T_Manufacturers", "T_Suppliers", "T_Customers",
        "T_Attendance", "T_Leave", "T_Expenses", "T_Staff"
    )

    foreach ($table in $clearTables) {
        Invoke-Sql "DELETE FROM [$table]"
    }
    Invoke-Sql "UPDATE T_DocumentNumbers SET NextNumber=1, LastUpdatedOn=Now()"

    Invoke-Sql @"
INSERT INTO T_Customers
    (CustomerCode, CustomerName, CustomerType, ContactPerson, Phone, Email, BillingAddress, ShippingAddress, Status, Notes, CreatedOn, UpdatedOn, BillingName)
VALUES
    ('INTERNAL', 'Internal Customer', 'Internal', 'Internal Test', '000', 'internal@example.invalid', 'Internal', 'Internal', 'Active', 'Reserved internal customer record.', Now(), Now(), 'Internal Customer')
"@

    foreach ($record in @(
        @{ Code = "SAMPLE-CUST-001"; Name = "Sample Customer A"; Contact = "Buyer A"; City = "Test City A" },
        @{ Code = "SAMPLE-CUST-002"; Name = "Sample Customer B"; Contact = "Buyer B"; City = "Test City B" },
        @{ Code = "SAMPLE-CUST-003"; Name = "Sample Customer C"; Contact = "Buyer C"; City = "Test City C" }
    )) {
        $code = $record.Code
        $name = $record.Name
        $contact = $record.Contact
        $city = $record.City
        Invoke-Sql @"
INSERT INTO T_Customers
    (CustomerCode, CustomerName, CustomerType, ContactPerson, Phone, Email, BillingAddress, ShippingAddress, Status, Notes, CreatedOn, UpdatedOn, BillingName, BillingCity)
VALUES
    ('$code', '$name', 'Customer', '$contact', '000-000-0000', '$code@example.invalid', '$city Test Address', '$city Ship Address', 'Active', 'Test record created by the sample-data reset.', Now(), Now(), '$name', '$city')
"@
    }

    foreach ($record in @(
        @{ Code = "SAMPLE-SUP-001"; Name = "Sample Supplier A"; Contact = "Sales A" },
        @{ Code = "SAMPLE-SUP-002"; Name = "Sample Supplier B"; Contact = "Sales B" },
        @{ Code = "SAMPLE-SUP-003"; Name = "Sample Supplier C"; Contact = "Sales C" }
    )) {
        Invoke-Sql @"
INSERT INTO T_Suppliers (SupplierCode, CompanyName, ContactPerson, Phone, Email, Address, PaymentTerms)
VALUES ('$($record.Code)', '$($record.Name)', '$($record.Contact)', '000-000-0000', '$($record.Code)@example.invalid', 'Sample supplier address', '30 days')
"@
    }

    foreach ($record in @(
        @{ Name = "Sample Components"; Description = "Sample component category." },
        @{ Name = "Sample Tools"; Description = "Sample tool category." },
        @{ Name = "Sample Finished Goods"; Description = "Sample finished-goods category." }
    )) {
        Invoke-Sql @"
INSERT INTO T_Categories (ItemTypeID, CategoryName, Description, RequiresSerial, RequiresExpiry)
SELECT TOP 1 ItemTypeID, '$($record.Name)', '$($record.Description)', 0, 0 FROM T_ItemTypes
"@
    }

    foreach ($record in @(
        @{ Code = "SAMPLE-WH-001"; Name = "Sample Warehouse A"; Address = "Sample Site A"; Contact = "Warehouse Contact A" },
        @{ Code = "SAMPLE-WH-002"; Name = "Sample Warehouse B"; Address = "Sample Site B"; Contact = "Warehouse Contact B" },
        @{ Code = "SAMPLE-WH-003"; Name = "Sample Warehouse C"; Address = "Sample Site C"; Contact = "Warehouse Contact C" }
    )) {
        Invoke-Sql @"
INSERT INTO T_Warehouses (WarehouseCode, WarehouseName, Address, ContactPerson)
VALUES ('$($record.Code)', '$($record.Name)', '$($record.Address)', '$($record.Contact)')
"@
    }

    foreach ($record in @(
        @{ Code = "SAMPLE-LOC-001"; Name = "Sample Location A"; Warehouse = "SAMPLE-WH-001"; Bin = "A1" },
        @{ Code = "SAMPLE-LOC-002"; Name = "Sample Location B"; Warehouse = "SAMPLE-WH-002"; Bin = "B1" },
        @{ Code = "SAMPLE-LOC-003"; Name = "Sample Location C"; Warehouse = "SAMPLE-WH-003"; Bin = "C1" }
    )) {
        Invoke-Sql @"
INSERT INTO T_Locations (WarehouseID, LocationCode, LocationName, BinNumber, ShelfNumber)
SELECT WarehouseID, '$($record.Code)', '$($record.Name)', '$($record.Bin)', '1'
FROM T_Warehouses WHERE WarehouseCode='$($record.Warehouse)'
"@
    }

    foreach ($record in @(
        @{ Name = "Sample Manufacturer A"; Code = "SAMPLE-MFR-001" },
        @{ Name = "Sample Manufacturer B"; Code = "SAMPLE-MFR-002" },
        @{ Name = "Sample Manufacturer C"; Code = "SAMPLE-MFR-003" }
    )) {
        Invoke-Sql @"
INSERT INTO T_Manufacturers (ManufacturerName, Website, Notes)
VALUES ('$($record.Name)', 'https://example.invalid', '$($record.Code) test record')
"@
    }

    foreach ($record in @(
        @{ Sku = "SAMPLE-ITEM-001"; Name = "Sample Controller"; Category = "Sample Components"; Location = "SAMPLE-LOC-001"; Reorder = 5; Quantity = 20; Price = 12.5; Supplier = "SAMPLE-SUP-001"; Manufacturer = "Sample Manufacturer A"; Part = "SAMPLE-MPN-001"; VendorPart = "SAMPLE-SPN-001" },
        @{ Sku = "SAMPLE-ITEM-002"; Name = "Sample Connector Set"; Category = "Sample Tools"; Location = "SAMPLE-LOC-002"; Reorder = 10; Quantity = 30; Price = 4.25; Supplier = "SAMPLE-SUP-002"; Manufacturer = "Sample Manufacturer B"; Part = "SAMPLE-MPN-002"; VendorPart = "SAMPLE-SPN-002" },
        @{ Sku = "SAMPLE-ITEM-003"; Name = "Sample Sensor Unit"; Category = "Sample Finished Goods"; Location = "SAMPLE-LOC-003"; Reorder = 3; Quantity = 12; Price = 18.75; Supplier = "SAMPLE-SUP-003"; Manufacturer = "Sample Manufacturer C"; Part = "SAMPLE-MPN-003"; VendorPart = "SAMPLE-SPN-003" }
    )) {
        Invoke-Sql @"
INSERT INTO T_Items
    (InternalSKU, ItemName, Description, CategoryID, BaseUnitID, ReorderLevel, MinStockLevel, MaxStockLevel, DefaultLocationID, RequiresSerial, CreatedOn, UpdatedOn, ItemClass, ItemType, IsSellable, DefaultSalesPrice, DefaultTaxable)
SELECT '$($record.Sku)', '$($record.Name)', 'Sample item for workflow testing.', C.CategoryID, U.UnitID, $($record.Reorder), $($record.Reorder), 100,
       L.LocationID, 0, Now(), Now(), 'Stock', 'Product', -1, $($record.Price), 0
FROM T_Categories AS C, T_Units AS U, T_Locations AS L
WHERE C.CategoryName='$($record.Category)' AND U.UnitCode='PCS' AND L.LocationCode='$($record.Location)'
"@

        Invoke-Sql @"
INSERT INTO T_ItemSuppliers (ItemID, SupplierID, PreferredSupplier, LastUnitPrice, LeadTimeDays, IsActive)
SELECT I.ItemID, S.SupplierID, -1, $($record.Price), 14, -1
FROM T_Items AS I, T_Suppliers AS S
WHERE I.InternalSKU='$($record.Sku)' AND S.SupplierCode='$($record.Supplier)'
"@

        Invoke-Sql @"
INSERT INTO T_ItemPartNumbers
    (ItemID, PartNumberTypeID, PartNumber, SupplierID, ManufacturerID, IsPrimary, Notes, IsVendorPartNumber, IsActive, UpdatedOn)
SELECT I.ItemID, (SELECT TOP 1 PartNumberTypeID FROM T_PartNumberTypes ORDER BY PartNumberTypeID),
       '$($record.Part)', S.SupplierID, M.ManufacturerID, -1, 'Sample manufacturer part number.', 0, -1, Now()
FROM T_Items AS I, T_Suppliers AS S, T_Manufacturers AS M
WHERE I.InternalSKU='$($record.Sku)' AND S.SupplierCode='$($record.Supplier)' AND M.ManufacturerName='$($record.Manufacturer)'
"@

        Invoke-Sql @"
INSERT INTO T_ItemPartNumbers
    (ItemID, PartNumberTypeID, PartNumber, SupplierID, ManufacturerID, IsPrimary, Notes, IsVendorPartNumber, IsActive, UpdatedOn)
SELECT I.ItemID, (SELECT TOP 1 PartNumberTypeID FROM T_PartNumberTypes ORDER BY PartNumberTypeID),
       '$($record.VendorPart)', S.SupplierID, M.ManufacturerID, 0, 'Sample supplier part number.', -1, -1, Now()
FROM T_Items AS I, T_Suppliers AS S, T_Manufacturers AS M
WHERE I.InternalSKU='$($record.Sku)' AND S.SupplierCode='$($record.Supplier)' AND M.ManufacturerName='$($record.Manufacturer)'
"@

        Invoke-Sql @"
INSERT INTO T_StockBalances (ItemID, CurrentLocationID, QuantityOnHand, QuantityReserved, LastUpdated)
SELECT I.ItemID, L.LocationID, $($record.Quantity), 0, Now()
FROM T_Items AS I, T_Locations AS L
WHERE I.InternalSKU='$($record.Sku)' AND L.LocationCode='$($record.Location)'
"@

        Invoke-Sql @"
INSERT INTO T_StockMovements
    (MovementDate, ItemID, MovementTypeID, Quantity, ToLocationID, ReferenceDocument, PerformedByUserID, Reason, UnitCost, SourceDocumentType, MovementNotes)
SELECT Now(), I.ItemID, MT.MovementTypeID, $($record.Quantity), L.LocationID, '$($record.Sku)',
       (SELECT TOP 1 UserID FROM T_Users ORDER BY UserID), 'Sample opening stock', $($record.Price), 'SampleData', 'Matching opening movement for seeded stock balance.'
FROM T_Items AS I, T_StockMovementTypes AS MT, T_Locations AS L
WHERE I.InternalSKU='$($record.Sku)' AND MT.MovementTypeCode='STOCK_IN' AND L.LocationCode='$($record.Location)'
"@
    }

    foreach ($record in @(
        @{ Code = "SAMPLE-PRJ-001"; Name = "Sample Engineering Project A"; Customer = "SAMPLE-CUST-001"; Sku = "SAMPLE-ITEM-003"; Location = "SAMPLE-LOC-001" },
        @{ Code = "SAMPLE-PRJ-002"; Name = "Sample Engineering Project B"; Customer = "SAMPLE-CUST-002"; Sku = "SAMPLE-ITEM-001"; Location = "SAMPLE-LOC-002" },
        @{ Code = "SAMPLE-PRJ-003"; Name = "Sample Engineering Project C"; Customer = "SAMPLE-CUST-003"; Sku = "SAMPLE-ITEM-002"; Location = "SAMPLE-LOC-003" }
    )) {
        Invoke-Sql @"
INSERT INTO T_Projects
    (ProjectCode, ProjectName, StartDate, Status, Budget, Notes, IsActive, CustomerID, FinishedItemID, BuildQuantity, OutputLocationID, WorkflowStatus)
SELECT '$($record.Code)', '$($record.Name)', Date(), 'Open', 1000, 'Sample engineering project.', -1, C.CustomerID, I.ItemID, 1, L.LocationID, 'Open'
FROM T_Customers AS C, T_Items AS I, T_Locations AS L
WHERE C.CustomerCode='$($record.Customer)' AND I.InternalSKU='$($record.Sku)' AND L.LocationCode='$($record.Location)'
"@
    }

    foreach ($record in @(
        @{ Code = "SAMPLE-WI-001"; Name = "Sample Work Instruction A"; Template = "SAMPLE-BOM-001"; FinishedSku = "SAMPLE-ITEM-003"; Customer = "SAMPLE-CUST-001"; ComponentSku = "SAMPLE-ITEM-001" },
        @{ Code = "SAMPLE-WI-002"; Name = "Sample Work Instruction B"; Template = "SAMPLE-BOM-002"; FinishedSku = "SAMPLE-ITEM-001"; Customer = "SAMPLE-CUST-002"; ComponentSku = "SAMPLE-ITEM-002" },
        @{ Code = "SAMPLE-WI-003"; Name = "Sample Work Instruction C"; Template = "SAMPLE-BOM-003"; FinishedSku = "SAMPLE-ITEM-002"; Customer = "SAMPLE-CUST-003"; ComponentSku = "SAMPLE-ITEM-003" }
    )) {
        Invoke-Sql @"
INSERT INTO T_WorkInstructions (InstructionName, Description, FilePath, CreatedOn, Active)
VALUES ('$($record.Code)', '$($record.Name)', '', Now(), -1)
"@
        Invoke-Sql @"
INSERT INTO T_AssemblyTemplates
    (TemplateCode, TemplateName, FinishedItemID, CustomerID, CustomerPartNumber, OutputQuantity, RevisionNo, Status, Instructions, Notes, CreatedOn, UpdatedOn, WorkInstructionID)
SELECT '$($record.Template)', '$($record.Name)', I.ItemID, C.CustomerID, 'SAMPLE-CPN-$($record.Template)', 1, 'A', 'Draft',
       'Sample engineering BOM template.', 'Sample template for workflow testing.', Now(), Now(), W.WorkInstructionID
FROM T_Items AS I, T_Customers AS C, T_WorkInstructions AS W
WHERE I.InternalSKU='$($record.FinishedSku)' AND C.CustomerCode='$($record.Customer)' AND W.InstructionName='$($record.Code)'
"@
        Invoke-Sql @"
INSERT INTO T_AssemblyTemplateMaterials (AssemblyTemplateID, ComponentItemID, QuantityPerOutput, EstimatedUnitCost, SortOrder, Notes)
SELECT T.AssemblyTemplateID, I.ItemID, 1, 1, 1, 'Sample BOM component.'
FROM T_AssemblyTemplates AS T, T_Items AS I
WHERE T.TemplateCode='$($record.Template)' AND I.InternalSKU='$($record.ComponentSku)'
"@
    }

    foreach ($record in @(
        @{ Name = "Sample Staff A"; Title = "Engineering"; Phone = "000-000-0001"; Category = "Supplies A"; Amount = 25.00 },
        @{ Name = "Sample Staff B"; Title = "Warehouse"; Phone = "000-000-0002"; Category = "Supplies B"; Amount = 40.00 },
        @{ Name = "Sample Staff C"; Title = "Operations"; Phone = "000-000-0003"; Category = "Supplies C"; Amount = 55.00 }
    )) {
        Invoke-Sql @"
INSERT INTO T_Staff (StaffName, JobTitle, Phone, Status)
VALUES ('$($record.Name)', '$($record.Title)', '$($record.Phone)', 'Active')
"@
        Invoke-Sql @"
INSERT INTO T_Attendance (StaffID, WorkDate, TimeIn, Notes)
SELECT StaffID, Date(), Now(), 'Sample attendance record.' FROM T_Staff WHERE StaffName='$($record.Name)'
"@
        Invoke-Sql @"
INSERT INTO T_Leave (StaffID, LeaveType, StartDate, EndDate, Days, Status, Notes)
SELECT StaffID, 'Sample Leave', Date()+30, Date()+31, 2, 'Pending', 'Sample leave record.'
FROM T_Staff WHERE StaffName='$($record.Name)'
"@
        Invoke-Sql @"
INSERT INTO T_Expenses (ExpenseDate, Category, Description, Amount, PaidTo, Notes)
VALUES (Date(), '$($record.Category)', 'Sample expense for testing.', $($record.Amount), 'Sample Payee', 'Test-only sample expense.')
"@
    }

    $expectedCounts = @{
        "T_Customers" = 4
        "T_Suppliers" = 3
        "T_Categories" = 3
        "T_Warehouses" = 3
        "T_Locations" = 3
        "T_Manufacturers" = 3
        "T_Items" = 3
        "T_ItemSuppliers" = 3
        "T_ItemPartNumbers" = 6
        "T_Projects" = 3
        "T_StockBalances" = 3
        "T_StockMovements" = 3
        "T_WorkInstructions" = 3
        "T_AssemblyTemplates" = 3
        "T_AssemblyTemplateMaterials" = 3
        "T_Staff" = 3
        "T_Attendance" = 3
        "T_Leave" = 3
        "T_Expenses" = 3
    }

    foreach ($table in $expectedCounts.Keys) {
        $actualCount = [int](Get-Scalar "SELECT Count(*) FROM [$table]")
        if ($actualCount -ne $expectedCounts[$table]) {
            throw "Expected $($expectedCounts[$table]) records in $table; found $actualCount."
        }
    }

    $transaction.Commit()
    $transaction = $null
    Write-Output "Sample data reset and seed committed to $fullPath"

    foreach ($table in @(
        "T_Customers", "T_Suppliers", "T_Categories", "T_Warehouses", "T_Locations",
        "T_Manufacturers", "T_Items", "T_ItemSuppliers", "T_ItemPartNumbers", "T_Projects",
        "T_StockBalances", "T_StockMovements", "T_WorkInstructions", "T_AssemblyTemplates",
        "T_AssemblyTemplateMaterials", "T_Staff", "T_Attendance", "T_Leave", "T_Expenses", "T_Units"
    )) {
        $count = Get-Scalar "SELECT Count(*) FROM [$table]"
        Write-Output ("COUNT {0}={1}" -f $table, $count)
    }
}
catch {
    if ($transaction) {
        try { $transaction.Rollback() } catch {}
    }
    throw
}
finally {
    if ($connection.State -eq [System.Data.ConnectionState]::Open) { $connection.Close() }
    $connection.Dispose()
}
