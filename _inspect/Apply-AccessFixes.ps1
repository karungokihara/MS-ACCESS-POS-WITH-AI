$ErrorActionPreference = "Stop"
$root = "C:\Users\user\Downloads\LLC"
$logPath = "$root\_inspect\apply-log.txt"
$modulePath = "$root\_inspect\M_UX.txt"
"" | Set-Content $logPath -Encoding UTF8

function Log($msg) {
    $line = "$(Get-Date -Format HH:mm:ss) $msg"
    Add-Content $logPath $line
    Write-Host $line
}

function Module-HasText($mod, $text) {
    $n = $mod.CountOfLines
    for ($i = 1; $i -le $n; $i++) {
        $line = [string]$mod.Lines($i, 1)
        if ($line.Contains($text)) { return $true }
    }
    return $false
}

function Find-LineStarts($mod, $prefix) {
    $n = $mod.CountOfLines
    for ($i = 1; $i -le $n; $i++) {
        $t = ([string]$mod.Lines($i, 1)).Trim()
        if ($t.StartsWith($prefix)) { return $i }
    }
    return 0
}

function Find-ExactLine($mod, $exact) {
    $n = $mod.CountOfLines
    for ($i = 1; $i -le $n; $i++) {
        $t = ([string]$mod.Lines($i, 1)).Trim()
        if ($t -eq $exact) { return $i }
    }
    return 0
}

function Get-ProcEnd($mod, $start) {
    if ($start -lt 1) { return 0 }
    $depth = 0
    $n = $mod.CountOfLines
    for ($i = $start; $i -le $n; $i++) {
        $t = ([string]$mod.Lines($i, 1)).Trim()
        if ($t -match '^(Private |Public )?(Sub |Function |Property )') { $depth++ }
        if ($t -eq "End Sub" -or $t -eq "End Function" -or $t -eq "End Property") {
            $depth--
            if ($depth -le 0) { return $i }
        }
    }
    return 0
}

function Add-ModuleCode($mod, $marker, $code) {
    if (Module-HasText $mod $marker) { return }
    $mod.InsertLines($mod.CountOfLines + 1, $code)
}

function Insert-AfterProcStart($mod, $procPrefix, $marker, $codeLine) {
    if (Module-HasText $mod $marker) { return }
    $start = Find-LineStarts $mod $procPrefix
    if ($start -eq 0) { Log "MISSING PROC $procPrefix"; return }
    $insertAt = $start
    $first = ([string]$mod.Lines($start, 1)).Trim()
    if (-not $first.Contains(")")) {
        $n = $mod.CountOfLines
        for ($i = $start; $i -le [Math]::Min($start + 40, $n); $i++) {
            $t = ([string]$mod.Lines($i, 1)).Trim()
            if ($t -eq ")" -or $t.EndsWith(")")) { $insertAt = $i; break }
        }
    }
    $mod.InsertLines($insertAt + 1, $codeLine)
}

function Replace-Procedure($mod, $procPrefix, $marker, $newCode) {
    if (Module-HasText $mod $marker) { return }
    $start = Find-LineStarts $mod $procPrefix
    if ($start -eq 0) { Log "MISSING PROC $procPrefix"; return }
    $end = Get-ProcEnd $mod $start
    if ($end -eq 0) { Log "MISSING END $procPrefix"; return }
    $mod.DeleteLines($start, ($end - $start + 1))
    $mod.InsertLines($start, $newCode)
}

function Replace-ContainingLine($mod, $contains, $newLine) {
    $n = $mod.CountOfLines
    for ($i = 1; $i -le $n; $i++) {
        $line = [string]$mod.Lines($i, 1)
        if ($line.Contains($contains)) {
            $mod.ReplaceLine($i, $newLine)
            return $true
        }
    }
    return $false
}

function Replace-Snippet($mod, $old, $new) {
    $n = $mod.CountOfLines
    $count = 0
    for ($i = 1; $i -le $n; $i++) {
        $line = [string]$mod.Lines($i, 1)
        if ($line.Contains($old)) {
            $mod.ReplaceLine($i, $line.Replace($old, $new))
            $count++
        }
    }
    return $count
}

function Open-DesignForm($acc, $name) {
    while ($acc.CurrentProject.AllForms.Item($name).IsLoaded) {
        $acc.DoCmd.Close(2, $name, 2)
    }
    $acc.DoCmd.OpenForm($name, 1)
    return ,$acc.Forms.Item($name)
}

function Save-DesignForm($acc, $name) {
    $acc.DoCmd.Close(2, $name, 1)
}

function Test-Control($frm, $name) {
    try {
        $null = $frm.Controls.Item($name)
        return $true
    } catch { return $false }
}

function Test-Section($frm, $section) {
    try {
        $null = $frm.Section($section).Height
        return $true
    } catch { return $false }
}

function Add-Button($acc, $frm, $formName, $buttonName, $caption, $onClick, $left, $top, $width, $height, $section) {
    if (Test-Control $frm $buttonName) { return }
    $use = $section
    if (-not (Test-Section $frm $use)) {
        if ($section -eq 1) {
            try { $acc.DoCmd.RunCommand(36) } catch {}
        }
        if (-not (Test-Section $frm 1)) { $use = 0 }
        else { $use = 1 }
    }
    try {
        $btn = $acc.CreateControl($formName, 104, $use, "", "", $left, $top, $width, $height)
        $btn.Name = $buttonName
        $btn.Caption = $caption
        $btn.OnClick = $onClick
        try {
            $sec = $frm.Section($use)
            $need = $top + $height + 80
            if ($sec.Height -lt $need) { $sec.Height = $need }
        } catch {}
    } catch {
        Log "Button $buttonName on $formName skipped: $($_.Exception.Message)"
    }
}

Get-Process MSACCESS -ErrorAction SilentlyContinue | Stop-Process -Force -ErrorAction SilentlyContinue
Start-Sleep -Seconds 1

Log "Opening Access"
$acc = New-Object -ComObject Access.Application
$acc.Visible = $false
$acc.OpenCurrentDatabase("$root\LLC.accdb")
$acc.DoCmd.SetWarnings($false)

try {
    while ($acc.Forms.Count -gt 0) {
        $acc.DoCmd.Close(2, $acc.Forms.Item(0).Name, 2)
    }
} catch {}

try {
Log "Loading M_UX"
$loaded = $false
try {
    $acc.LoadFromText(5, "M_UX", $modulePath)
    $loaded = $true
    Log "LoadFromText M_UX ok"
} catch {
    Log "LoadFromText failed: $($_.Exception.Message)"
}

if (-not $loaded) {
    $proj = $acc.VBE.ActiveVBProject
    $exists = $false
    foreach ($c in $proj.VBComponents) {
        if ($c.Name -eq "M_UX") { $exists = $true; $comp = $c }
    }
    if (-not $exists) {
        $comp = $proj.VBComponents.Add(1)
        $comp.Name = "M_UX"
    }
    $code = Get-Content $modulePath -Raw
    $code = $code -replace '(?m)^Attribute VB_Name = "M_UX"\r?\n', ''
    if ($comp.CodeModule.CountOfLines -gt 0) {
        $comp.CodeModule.DeleteLines(1, $comp.CodeModule.CountOfLines)
    }
    $comp.CodeModule.AddFromString($code)
    Log "VBE M_UX ok"
}

# ---------- Queries ----------
Log "Updating queries"
$db = $acc.CurrentDb()

function Set-Query($db, $name, $sql) {
    $found = $false
    foreach ($q in $db.QueryDefs) {
        if ($q.Name -eq $name) {
            $q.SQL = $sql
            $found = $true
            break
        }
    }
    if (-not $found) { [void]$db.CreateQueryDef($name, $sql) }
}

$company = @"
"Value Transformation LLC" AS CompanyName, "1211 Heath Church Road" AS CompanyAddress1, "Lexington, NC, 27292, phone: 336-963-0119" AS CompanyAddress2, "email: Jon.Quigley@ValueTransform.com" AS CompanyEmail, "Product Development Experts - Training and Consulting" AS CompanyFooter
"@.Trim()

$inventorySql = @"
SELECT 'Bulk' AS StockType, Nz(B.StockBalanceID, I.ItemID) AS SourceRecordID, I.ItemID AS ItemID, Null AS ItemInstanceID, I.CategoryID AS CategoryID, L.WarehouseID AS WarehouseID, B.CurrentLocationID AS LocationID, I.InternalSKU AS InternalSKU, I.ItemName AS ItemName, I.Description AS Description, C.CategoryName AS CategoryName, Null AS SerialNumber, Nz(B.QuantityOnHand,0) AS QuantityOnHand, Nz(B.QuantityReserved,0) AS QuantityReserved, Nz(B.QuantityOnHand,0)-Nz(B.QuantityReserved,0) AS AvailableQuantity, IIf(Nz(B.QuantityOnHand,0)-Nz(B.QuantityReserved,0)<=0,'Out of Stock',IIf(Nz(I.ReorderLevel,0)>0 AND Nz(B.QuantityOnHand,0)-Nz(B.QuantityReserved,0)<=Nz(I.ReorderLevel,0),'Low Stock','In Stock')) AS StockStatus, IIf(Nz(I.ReorderLevel,0)>0 AND Nz(B.QuantityOnHand,0)-Nz(B.QuantityReserved,0)<=0,'Restock Required',IIf(Nz(I.ReorderLevel,0)>0 AND Nz(B.QuantityOnHand,0)-Nz(B.QuantityReserved,0)<=Nz(I.ReorderLevel,0),'Low Stock','OK')) AS StockAlert, W.WarehouseName AS WarehouseName, L.LocationCode AS LocationCode, L.LocationName AS LocationName, L.BinNumber AS BinNumber, L.ShelfNumber AS ShelfNumber, IIf(Len(Trim(Nz(U.UnitCode,'')))=0,Nz(U.UnitName,''),U.UnitCode) AS UnitCode, Nz(I.ReorderLevel,0) AS ReorderLevel, B.LastUpdated AS LastUpdated
FROM ((((T_Items AS I LEFT JOIN T_StockBalances AS B ON I.ItemID = B.ItemID) LEFT JOIN T_Categories AS C ON I.CategoryID = C.CategoryID) LEFT JOIN T_Units AS U ON I.BaseUnitID = U.UnitID) LEFT JOIN T_Locations AS L ON B.CurrentLocationID = L.LocationID) LEFT JOIN T_Warehouses AS W ON L.WarehouseID = W.WarehouseID
WHERE Nz(I.RequiresSerial,False)=False
UNION ALL SELECT 'Serialized' AS StockType, S.ItemInstanceID AS SourceRecordID, I.ItemID AS ItemID, S.ItemInstanceID AS ItemInstanceID, I.CategoryID AS CategoryID, L.WarehouseID AS WarehouseID, S.CurrentLocationID AS LocationID, I.InternalSKU AS InternalSKU, I.ItemName AS ItemName, I.Description AS Description, C.CategoryName AS CategoryName, S.SerialNumber AS SerialNumber, 1 AS QuantityOnHand, 0 AS QuantityReserved, IIf(Nz(S.CurrentStatus,'')='In Stock',1,0) AS AvailableQuantity, Nz(S.CurrentStatus,'Unknown') AS StockStatus, IIf(Nz(S.CurrentStatus,'')='In Stock','OK',Nz(S.CurrentStatus,'Unknown')) AS StockAlert, W.WarehouseName AS WarehouseName, L.LocationCode AS LocationCode, L.LocationName AS LocationName, L.BinNumber AS BinNumber, L.ShelfNumber AS ShelfNumber, IIf(Len(Trim(Nz(U.UnitCode,'')))=0,Nz(U.UnitName,''),U.UnitCode) AS UnitCode, Nz(I.ReorderLevel,0) AS ReorderLevel, S.AcquisitionDate AS LastUpdated
FROM ((((T_ItemInstances AS S INNER JOIN T_Items AS I ON S.ItemID = I.ItemID) LEFT JOIN T_Categories AS C ON I.CategoryID = C.CategoryID) LEFT JOIN T_Units AS U ON I.BaseUnitID = U.UnitID) LEFT JOIN T_Locations AS L ON S.CurrentLocationID = L.LocationID) LEFT JOIN T_Warehouses AS W ON L.WarehouseID = W.WarehouseID
WHERE Nz(I.RequiresSerial,False)=True
UNION ALL SELECT 'Serialized' AS StockType, I.ItemID AS SourceRecordID, I.ItemID AS ItemID, Null AS ItemInstanceID, I.CategoryID AS CategoryID, Null AS WarehouseID, Null AS LocationID, I.InternalSKU AS InternalSKU, I.ItemName AS ItemName, I.Description AS Description, C.CategoryName AS CategoryName, Null AS SerialNumber, 0 AS QuantityOnHand, 0 AS QuantityReserved, 0 AS AvailableQuantity, 'No Stock' AS StockStatus, 'No Stock' AS StockAlert, Null AS WarehouseName, Null AS LocationCode, Null AS LocationName, Null AS BinNumber, Null AS ShelfNumber, IIf(Len(Trim(Nz(U.UnitCode,'')))=0,Nz(U.UnitName,''),U.UnitCode) AS UnitCode, Nz(I.ReorderLevel,0) AS ReorderLevel, I.UpdatedOn AS LastUpdated
FROM (T_Items AS I LEFT JOIN T_Categories AS C ON I.CategoryID = C.CategoryID) LEFT JOIN T_Units AS U ON I.BaseUnitID = U.UnitID
WHERE Nz(I.RequiresSerial,False)=True AND Not Exists (SELECT 1 FROM T_ItemInstances AS S WHERE S.ItemID = I.ItemID)
ORDER BY ItemName, StockType, SerialNumber;
"@

$itemSetupSql = @"
SELECT I.ItemID, I.InternalSKU, I.ItemName, I.Description, I.CategoryID, C.CategoryName, IIf(Nz(I.RequiresSerial,False)=True,'Serialized','Bulk') AS StockType, I.BaseUnitID, IIf(Len(Trim(Nz(U.UnitCode,'')))=0,Nz(U.UnitName,''),U.UnitCode) AS UnitCode, IIf(Nz(I.RequiresSerial,False), Nz((SELECT Count(*) FROM T_ItemInstances AS S WHERE S.ItemID=I.ItemID AND Nz(S.CurrentStatus,'')='In Stock'),0), Nz(B.QuantityOnHand,0)) AS QuantityOnHand, I.ReorderLevel, I.DefaultLocationID, L.LocationCode AS DefaultLocationCode, L.LocationName AS DefaultLocationName, W.WarehouseName AS DefaultWarehouseName, I.CreatedOn, I.UpdatedOn
FROM ((((T_Items AS I LEFT JOIN T_Categories AS C ON I.CategoryID = C.CategoryID) LEFT JOIN T_Units AS U ON I.BaseUnitID = U.UnitID) LEFT JOIN T_Locations AS L ON I.DefaultLocationID = L.LocationID) LEFT JOIN T_Warehouses AS W ON L.WarehouseID = W.WarehouseID) LEFT JOIN (SELECT ItemID, Sum(Nz(QuantityOnHand,0)) AS QuantityOnHand FROM T_StockBalances GROUP BY ItemID) AS B ON I.ItemID = B.ItemID
ORDER BY I.ItemName, I.InternalSKU;
"@

$soHeader = @"
SELECT SO.SalesOrderID AS InvoiceID, SO.SalesOrderNo AS InvoiceNo, SO.CustomerID, C.CustomerName, SO.CustomerPOID, PO.PONumber, SO.SalesOrderID, SO.SalesOrderNo, SO.OrderDate AS InvoiceDate, Nz(SO.SalesOrderStatus,'Sales Order') AS InvoiceType, SO.PaymentTerms, IIf(Len(Trim(Nz(C.BillingName,'')))=0,Nz(C.CustomerName,''),C.BillingName) AS BillToName, IIf(Len(Trim(Nz(C.BillingAddress,'')))=0,Nz(SO.ShipToAddress,''),C.BillingAddress) AS BillToAddress, Nz(L.SubtotalAmount,0) AS SubtotalAmount, 0 AS TaxAmount, Nz(L.SubtotalAmount,0) AS GrandTotalAmount, SO.SalesOrderStatus AS InvoiceStatus, SO.Notes, $company, Nz(SO.Notes,'') AS InvoiceDescription
FROM ((T_SalesOrders AS SO LEFT JOIN T_Customers AS C ON SO.CustomerID = C.CustomerID) LEFT JOIN T_CustomerPOs AS PO ON SO.CustomerPOID = PO.CustomerPOID) LEFT JOIN (SELECT SalesOrderID, Sum(Nz(LineSubtotal,0)) AS SubtotalAmount FROM T_SalesOrderLines GROUP BY SalesOrderID) AS L ON SO.SalesOrderID = L.SalesOrderID;
"@

$soLines = @"
SELECT L.SalesOrderLineID AS InvoiceLineID, L.SalesOrderID AS InvoiceID, L.ItemID, Nz(I.InternalSKU,'') AS ItemCode, Nz(L.LineDescription,'') AS LineDescription, Nz(L.QuantityOrdered,0) AS Quantity, L.UnitID, IIf(Len(Trim(Nz(U.UnitCode,'')))=0,Nz(U.UnitName,''),U.UnitCode) AS UnitCode, Nz(L.UnitPrice,0) AS UnitPrice, L.Taxable, Nz(L.LineSubtotal,0) AS LineSubtotal, Trim(Nz(I.InternalSKU,'') & ' ' & Nz(L.LineDescription,'')) AS ReportItem, Nz(L.LineDescription,'') AS ReportDescription, IIf(Len(Trim(Nz(U.UnitCode,'')))=0,Nz(U.UnitName,''),U.UnitCode) AS ReportUnit
FROM (T_SalesOrderLines AS L LEFT JOIN T_Items AS I ON L.ItemID = I.ItemID) LEFT JOIN T_Units AS U ON L.UnitID = U.UnitID;
"@

$poHeader = @"
SELECT P.CustomerPOID AS InvoiceID, P.PONumber AS InvoiceNo, P.CustomerID, C.CustomerName, P.CustomerPOID, P.PONumber, P.SalesOrderID, SO.SalesOrderNo, P.PODate AS InvoiceDate, Nz(P.POType,'Purchase Order') AS InvoiceType, P.PaymentTerms, IIf(Len(Trim(Nz(C.BillingName,'')))=0,Nz(C.CustomerName,''),C.BillingName) AS BillToName, IIf(Len(Trim(Nz(P.ShipToAddress,'')))=0,Nz(C.BillingAddress,''),P.ShipToAddress) AS BillToAddress, Nz(P.SubtotalAmount,0) AS SubtotalAmount, Nz(P.TaxAmount,0) AS TaxAmount, Nz(P.TotalAmount,0) AS GrandTotalAmount, P.POStatus AS InvoiceStatus, P.Notes, $company, Nz(P.Notes,'') AS InvoiceDescription
FROM (T_CustomerPOs AS P LEFT JOIN T_Customers AS C ON P.CustomerID = C.CustomerID) LEFT JOIN T_SalesOrders AS SO ON P.SalesOrderID = SO.SalesOrderID;
"@

$poLines = @"
SELECT L.CustomerPOLineID AS InvoiceLineID, L.CustomerPOID AS InvoiceID, L.ItemID, Nz(I.InternalSKU,Nz(L.CustomerPartNumber,'')) AS ItemCode, Nz(L.LineDescription,'') AS LineDescription, Nz(L.QuantityOrdered,0) AS Quantity, L.UnitID, IIf(Len(Trim(Nz(U.UnitCode,'')))=0,Nz(U.UnitName,''),U.UnitCode) AS UnitCode, Nz(L.UnitPrice,0) AS UnitPrice, L.Taxable, Nz(L.LineSubtotal,0) AS LineSubtotal, Trim(Nz(I.InternalSKU,Nz(L.CustomerPartNumber,'')) & ' ' & Nz(L.LineDescription,'')) AS ReportItem, Nz(L.LineDescription,'') AS ReportDescription, IIf(Len(Trim(Nz(U.UnitCode,'')))=0,Nz(U.UnitName,''),U.UnitCode) AS ReportUnit
FROM (T_CustomerPOLines AS L LEFT JOIN T_Items AS I ON L.ItemID = I.ItemID) LEFT JOIN T_Units AS U ON L.UnitID = U.UnitID;
"@

$spoHeader = @"
SELECT P.SupplierPOID AS InvoiceID, P.PONumber AS InvoiceNo, P.SupplierID AS CustomerID, S.CompanyName AS CustomerName, Null AS CustomerPOID, P.PONumber, Null AS SalesOrderID, Null AS SalesOrderNo, P.PODate AS InvoiceDate, Nz(P.POType,'Purchase Order') AS InvoiceType, P.PaymentTerms, Nz(P.ShipToName,S.CompanyName) AS BillToName, Nz(P.ShipToAddress,'') AS BillToAddress, Nz(P.SubtotalAmount,0) AS SubtotalAmount, Nz(P.TaxAmount,0) AS TaxAmount, Nz(P.TotalAmount,0) AS GrandTotalAmount, P.POStatus AS InvoiceStatus, P.Notes, $company, Nz(P.Notes,'') AS InvoiceDescription
FROM T_SupplierPOs AS P LEFT JOIN T_Suppliers AS S ON P.SupplierID = S.SupplierID;
"@

$spoLines = @"
SELECT L.SupplierPOLineID AS InvoiceLineID, L.SupplierPOID AS InvoiceID, L.ItemID, Nz(I.InternalSKU,Nz(L.SupplierPartNumber,'')) AS ItemCode, Nz(L.ItemDescription,'') AS LineDescription, Nz(L.OrderedQuantity,0) AS Quantity, L.UnitID, IIf(Len(Trim(Nz(U.UnitCode,'')))=0,Nz(U.UnitName,''),U.UnitCode) AS UnitCode, Nz(L.UnitPrice,0) AS UnitPrice, False AS Taxable, Nz(L.OrderedQuantity,0)*Nz(L.UnitPrice,0) AS LineSubtotal, Trim(Nz(I.InternalSKU,'') & ' ' & Nz(L.ItemDescription,'')) AS ReportItem, Nz(L.ItemDescription,'') AS ReportDescription, IIf(Len(Trim(Nz(U.UnitCode,'')))=0,Nz(U.UnitName,''),U.UnitCode) AS ReportUnit
FROM (T_SupplierPOLines AS L LEFT JOIN T_Items AS I ON L.ItemID = I.ItemID) LEFT JOIN T_Units AS U ON L.UnitID = U.UnitID;
"@

$dnHeader = @"
SELECT I.InvoiceID, I.InvoiceNo, I.CustomerID, C.CustomerName, I.CustomerPOID, PO.PONumber, I.SalesOrderID, SO.SalesOrderNo, I.InvoiceDate, 'Delivery Note' AS InvoiceType, I.PaymentTerms, Nz(SO.ShipToName,Nz(C.CustomerName,I.BillToName)) AS BillToName, Nz(SO.ShipToAddress,Nz(C.ShippingAddress,I.BillToAddress)) AS BillToAddress, I.SubtotalAmount, I.TaxAmount, I.GrandTotalAmount, I.InvoiceStatus, I.Notes, $company, Nz(SO.ShipVia,'') & ' ' & Nz(I.Notes,'') AS InvoiceDescription
FROM ((T_Invoices AS I LEFT JOIN T_Customers AS C ON I.CustomerID = C.CustomerID) LEFT JOIN T_CustomerPOs AS PO ON I.CustomerPOID = PO.CustomerPOID) LEFT JOIN T_SalesOrders AS SO ON I.SalesOrderID = SO.SalesOrderID;
"@

Set-Query $db "Q_InventoryListAll" $inventorySql
Set-Query $db "Q_ItemSetupList" $itemSetupSql
Set-Query $db "Q_SalesOrderReportHeader" $soHeader
Set-Query $db "Q_SalesOrderReportLines" $soLines
Set-Query $db "Q_CustomerPOReportHeader" $poHeader
Set-Query $db "Q_CustomerPOReportLines" $poLines
Set-Query $db "Q_SupplierPOReportHeader" $spoHeader
Set-Query $db "Q_SupplierPOReportLines" $spoLines
Set-Query $db "Q_DeliveryNoteReportHeader" $dnHeader
Log "Queries saved"

# ---------- Reports ----------
Log "Building printable documents"
function Ensure-ReportCopy($acc, $newName, $sourceName) {
    $exists = $false
    foreach ($r in $acc.CurrentProject.AllReports) { if ($r.Name -eq $newName) { $exists = $true } }
    if (-not $exists) { $acc.DoCmd.CopyObject("", $newName, 3, $sourceName) }
}

function Tune-Report($acc, $reportName, $recordSource, $titleCaption, $subName, $subObject) {
    if ($acc.CurrentProject.AllReports.Item($reportName).IsLoaded) {
        $acc.DoCmd.Close(3, $reportName, 2)
    }
    $acc.DoCmd.OpenReport($reportName, 1)
    $rpt = $acc.Reports.Item($reportName)
    $rpt.RecordSource = $recordSource
    $rpt.Caption = $titleCaption
    foreach ($ctl in $rpt.Controls) {
        try {
            if ($ctl.Name -eq "Label10") { $ctl.Caption = $titleCaption }
            if ($subName -and $ctl.Name -eq "R_InvoiceLines") {
                $ctl.SourceObject = $subObject
            }
        } catch {}
    }
    $acc.DoCmd.Close(3, $reportName, 1)
}

Ensure-ReportCopy $acc "R_SalesDocLines" "R_InvoiceLines"
Tune-Report $acc "R_SalesDocLines" "Q_SalesOrderReportLines" "Lines" $null $null
Ensure-ReportCopy $acc "R_Quotation" "R_Invoice"
Tune-Report $acc "R_Quotation" "Q_SalesOrderReportHeader" "QUOTATION #" "R_InvoiceLines" "Report.R_SalesDocLines"
Ensure-ReportCopy $acc "R_SalesOrder" "R_Invoice"
Tune-Report $acc "R_SalesOrder" "Q_SalesOrderReportHeader" "SALES ORDER #" "R_InvoiceLines" "Report.R_SalesDocLines"
Ensure-ReportCopy $acc "R_POLines" "R_InvoiceLines"
Tune-Report $acc "R_POLines" "Q_CustomerPOReportLines" "Lines" $null $null
Ensure-ReportCopy $acc "R_PurchaseOrder" "R_Invoice"
Tune-Report $acc "R_PurchaseOrder" "Q_CustomerPOReportHeader" "PURCHASE ORDER #" "R_InvoiceLines" "Report.R_POLines"
Ensure-ReportCopy $acc "R_SupplierPOLinesRpt" "R_InvoiceLines"
Tune-Report $acc "R_SupplierPOLinesRpt" "Q_SupplierPOReportLines" "Lines" $null $null
Ensure-ReportCopy $acc "R_SupplierPO" "R_Invoice"
Tune-Report $acc "R_SupplierPO" "Q_SupplierPOReportHeader" "PURCHASE ORDER #" "R_InvoiceLines" "Report.R_SupplierPOLinesRpt"
Ensure-ReportCopy $acc "R_DeliveryNote" "R_Invoice"
Tune-Report $acc "R_DeliveryNote" "Q_DeliveryNoteReportHeader" "DELIVERY NOTE #" $null $null
Log "Reports saved"

# ---------- Boolean defaults ----------
Log "Setting yes/no defaults"
foreach ($td in $db.TableDefs) {
    if ($td.Name.StartsWith("MSys") -or $td.Name.StartsWith("~")) { continue }
    foreach ($f in $td.Fields) {
        if ($f.Type -eq 1) {
            try { $f.DefaultValue = "0" } catch {}
            try {
                $db.Execute("UPDATE [" + $td.Name + "] SET [" + $f.Name + "]=0 WHERE [" + $f.Name + "] Is Null")
            } catch {}
        }
    }
}

# ---------- Form patches ----------
Log "Patching forms"

# Warehouses
$frm = Open-DesignForm $acc "F_Warehouses"
foreach ($p in $frm.Properties) { if ($p.Name -eq "HasModule") { $p.Value = $true } }
foreach ($p in $frm.Properties) {
    if ($p.Name -eq "Filter") { $p.Value = "" }
    if ($p.Name -eq "FilterOn") { $p.Value = $false }
    if ($p.Name -eq "Caption") { $p.Value = "Warehouses" }
    if ($p.Name -eq "AllowAdditions") { $p.Value = $true }
    if ($p.Name -eq "AllowEdits") { $p.Value = $true }
}
if (Test-Control $frm "Command14") { $frm.Controls("Command14").OnClick = "[Event Procedure]"; $frm.Controls("Command14").Caption = "Save" }
if (Test-Control $frm "Command15") { $frm.Controls("Command15").OnClick = "[Event Procedure]"; $frm.Controls("Command15").Caption = "Back" }
if (Test-Control $frm "Command16") { $frm.Controls("Command16").OnClick = "[Event Procedure]"; $frm.Controls("Command16").Caption = "Add New" }
$wmod = $frm.Module
Add-ModuleCode $wmod "UX-WAREHOUSE" @"
' UX-WAREHOUSE
Private Sub Form_Load()
    On Error Resume Next
    Me.Filter = ""
    Me.FilterOn = False
    Me.Caption = "Warehouses"
    If Me.NewRecord Then ApplyWarehouseCode
    If Left`$(Nz(Me.OpenArgs, ""), 9) = "QUICKADD|" Then
        Me!WarehouseName = Mid`$(Me.OpenArgs, 10)
        ApplyWarehouseCode
    End If
    UX_MarkRequired Me, "WarehouseName", "WarehouseCode"
End Sub

Private Sub Form_BeforeUpdate(Cancel As Integer)
    On Error GoTo Err_Handler
    If Len(Trim`$(Nz(Me!WarehouseName, ""))) = 0 Then
        MsgBox "Enter the warehouse name.", vbExclamation, "Warehouse"
        Cancel = True
        Exit Sub
    End If
    ApplyWarehouseCode
    Exit Sub
Err_Handler:
    Cancel = True
    MsgBox Err.Description, vbExclamation, "Warehouse"
End Sub

Private Sub WarehouseName_AfterUpdate()
    ApplyWarehouseCode
End Sub

Private Sub ApplyWarehouseCode()
    On Error Resume Next
    Dim prefix As String
    Dim n As Long
    Dim candidate As String
    If Len(Trim`$(Nz(Me!WarehouseCode, ""))) > 0 Then Exit Sub
    If Len(Trim`$(Nz(Me!WarehouseName, ""))) = 0 Then Exit Sub
    prefix = UCase`$(Left`$(Trim`$(Me!WarehouseName), 4))
    If Len(prefix) < 4 Then prefix = prefix & String`$(4 - Len(prefix), "X")
    n = 1
    Do
        candidate = prefix & "-" & Format`$(n, "000")
        If IsNull(DLookup("WarehouseID", "T_Warehouses", "WarehouseCode=" & UX_SqlText(candidate))) Then Exit Do
        n = n + 1
    Loop
    Me!WarehouseCode = candidate
End Sub

Private Sub Command14_Click()
    On Error GoTo Err_Handler
    If Me.Dirty Then Me.Dirty = False
    MsgBox "Warehouse saved. You can type locations in the list below.", vbInformation, "Warehouse"
    Exit Sub
Err_Handler:
    MsgBox Err.Description, vbExclamation, "Warehouse"
End Sub

Private Sub Command15_Click()
    UX_GoBack
End Sub

Private Sub Command16_Click()
    On Error Resume Next
    If Me.Dirty Then Me.Dirty = False
    DoCmd.GoToRecord , , acNewRec
    Me!WarehouseCode = Null
    Me!WarehouseName = Null
End Sub

Private Sub Form_Unload(Cancel As Integer)
    On Error Resume Next
    If Me.NewRecord And Me.Dirty Then Me.Undo
End Sub
"@
Save-DesignForm $acc "F_Warehouses"
Log "Warehouses form saved"

# Location setup points warehouse quick-add at the real warehouse form
$frm = Open-DesignForm $acc "F_LocationSetup"
$lmod = $frm.Module
Replace-Snippet $lmod "F_WarehouseList" "F_Warehouses" | Out-Null
Insert-AfterProcStart $lmod "Private Sub Form_Load()" "UX_MarkRequired Me, ""LocationName""" "    UX_MarkRequired Me, ""LocationName"", ""cboWarehouseID"""
if (Test-Control $frm "WarehouseID_Label") { $frm.Controls("WarehouseID_Label").Caption = "Warehouse (type a new name to add it)" }
if (Test-Control $frm "LocationName_Label") { $frm.Controls("LocationName_Label").Caption = "Location Name" }
Add-Button $acc $frm "F_LocationSetup" "cmdUXBack" "Back" "=UX_GoBack()" 200 80 1400 360 1
Save-DesignForm $acc "F_LocationSetup"
Log "Location setup saved"

# Stock in
$frm = Open-DesignForm $acc "F_StockIn"
$smod = $frm.Module
Insert-AfterProcStart $smod "Private Sub Form_Load()" "Negative Adjustment" @"
    gSkipInventory = False
    Me.cboReason.RowSource = "Opening Balance;Purchase Receipt;Return to Stock;Positive Adjustment;Negative Adjustment"
    Me.cboToLocationID.LimitToList = True
    Me.cboToLocationID.OnNotInList = "[Event Procedure]"
    UX_MarkRequired Me, "cboItemID", "cboToLocationID", "txtQuantity", "cboReason"
"@
Insert-AfterProcStart $smod "Private Sub cmdPostStockIn_Click()" "gSkipInventory = False" "    gSkipInventory = False"
if (-not (Module-HasText $smod "UX_AdjustStockQuantity")) {
    $start = Find-ExactLine $smod "If Quantity <= 0 Then"
    if ($start -gt 0) {
        # delete until the matching End If of this small block: next 10 lines include End If
        $end = $start
        $n = $smod.CountOfLines
        for ($i = $start; $i -le [Math]::Min($start + 20, $n); $i++) {
            if (([string]$smod.Lines($i, 1)).Trim() -eq "End If") { $end = $i; break }
        }
        $smod.DeleteLines($start, ($end - $start + 1))
        $smod.InsertLines($start, @"
        If Not UX_AdjustStockQuantity(Quantity, Nz(Me.cboReason, vbNullString)) Then
            Me.txtQuantity.SetFocus
            Exit Function
        End If
"@)
    } else { Log "Quantity check block not found" }
}
if (-not (Module-HasText $smod "negative adjustment is larger")) {
    $idx = 0
    $n = $smod.CountOfLines
    for ($i = 1; $i -le $n; $i++) {
        if (([string]$smod.Lines($i, 1)).Contains("rs!QuantityOnHand = currentQuantity + Quantity")) { $idx = $i; break }
    }
    if ($idx -gt 0) {
        $smod.InsertLines($idx + 1, @"
        If Nz(rs!QuantityOnHand, 0) < -0.0000001 Then
            Err.Raise vbObjectError + 1006, "UpdateBulkStockBalance", "That negative adjustment is larger than the quantity on hand (" & currentQuantity & ")."
        End If
"@)
    }
}
Add-ModuleCode $smod "cboToLocationID_NotInList" @"
Private Sub cboToLocationID_NotInList(NewData As String, Response As Integer)
    On Error GoTo Err_Handler
    Response = acDataErrContinue
    If HandleComboNotInList(Me.cboToLocationID, NewData, "F_LocationSetup", "T_Locations", "LocationID", "LocationName", "LocationCode") Then
        Response = acDataErrAdded
    Else
        Me.cboToLocationID.Undo
    End If
    Exit Sub
Err_Handler:
    Response = acDataErrContinue
    MsgBox Err.Description, vbExclamation, "Add Location"
End Sub

Private Sub cmdAddLocation_Click()
    On Error GoTo Err_Handler
    Dim oldMax As Long
    Dim newMax As Long
    oldMax = Nz(DMax("LocationID", "T_Locations"), 0)
    DoCmd.OpenForm "F_LocationSetup", acNormal, , , acFormAdd, acDialog
    newMax = Nz(DMax("LocationID", "T_Locations"), 0)
    Me.cboToLocationID.Requery
    If newMax > oldMax Then Me.cboToLocationID = newMax
    Exit Sub
Err_Handler:
    If Err.Number <> 2501 Then MsgBox Err.Description, vbExclamation, "Add Location"
End Sub
"@
if (Test-Control $frm "cboToLocationID") {
    $loc = $frm.Controls("cboToLocationID")
    $loc.LimitToList = $true
    $loc.OnNotInList = "[Event Procedure]"
    $loc.ControlTipText = "Select a location, or type a new location name to create it."
    Add-Button $acc $frm "F_StockIn" "cmdAddLocation" "Add Location" "[Event Procedure]" ($loc.Left + $loc.Width + 120) $loc.Top 1800 $loc.Height 0
}
Add-Button $acc $frm "F_StockIn" "cmdUXBack" "Back" "=UX_GoBack()" 200 80 1400 360 1
Save-DesignForm $acc "F_StockIn"
Log "Stock in saved"

# Stock out skip flag and back button
$frm = Open-DesignForm $acc "F_StockOut"
$omod = $frm.Module
Insert-AfterProcStart $omod "Private Sub cmdPostStockOut_Click()" "UX_ProjectSkipsInventory(Nz(mProjectID, 0))" "    gSkipInventory = UX_ProjectSkipsInventory(Nz(mProjectID, 0))"
Insert-AfterProcStart $omod "Private Sub ReduceBulkStockBalance(" "If gSkipInventory Then Exit Sub" "    If gSkipInventory Then Exit Sub"
Insert-AfterProcStart $omod "Private Sub IssueSerializedInstance(" "If gSkipInventory Then Exit Sub" "    If gSkipInventory Then Exit Sub"
Insert-AfterProcStart $omod "Private Sub InsertStockMovement(" "If gSkipInventory Then Exit Sub" "    If gSkipInventory Then Exit Sub"
Add-Button $acc $frm "F_StockOut" "cmdUXBack" "Back" "=UX_GoBack()" 200 80 1400 360 1
Save-DesignForm $acc "F_StockOut"
Log "Stock out saved"

# Item setup duplicate name and required marks
$frm = Open-DesignForm $acc "F_ItemSetup"
$imod = $frm.Module
if (-not (Module-HasText $imod "UX_ReviewDuplicateItemName")) {
    $idx = Find-ExactLine $imod "currentItemID = Nz(Me.txtItemID, 0)"
    if ($idx -gt 0) {
        $imod.InsertLines($idx, @"
    If Not UX_ReviewDuplicateItemName(Nz(Me.txtItemName, ""), Nz(Me.txtItemID, 0), Nz(Me.txtDescription, "")) Then
        Cancel = True
        Exit Sub
    End If
"@)
    } else { Log "Item duplicate insert point missing" }
}
Insert-AfterProcStart $imod "Private Sub Form_Load()" "UX_MarkRequired Me, ""txtItemName""" "    UX_MarkRequired Me, ""txtItemName"", ""cboCategoryID"", ""cboBaseUnitID"", ""cboItemClass"", ""cboDefaultLocationID"""
Replace-Snippet $imod "SELECT UnitID, UnitCode, UnitName " "SELECT UnitID, IIf(Len(Trim(Nz([UnitCode],'')))=0,[UnitName],[UnitCode]), UnitName " | Out-Null
Add-Button $acc $frm "F_ItemSetup" "cmdUXBack" "Back" "=UX_GoBack()" 200 80 1400 360 1
Save-DesignForm $acc "F_ItemSetup"
Log "Item setup saved"

# Item list shows unit code and quantity
$frm = Open-DesignForm $acc "SF_ItemSetupList"
if (Test-Control $frm "BaseUnitID") {
    $frm.Controls("BaseUnitID").ControlSource = "UnitCode"
}
if (Test-Control $frm "BaseUnitID_Label") { $frm.Controls("BaseUnitID_Label").Caption = "Unit Code" }
if (-not (Test-Control $frm "QuantityOnHand")) {
    $txt = $acc.CreateControl("SF_ItemSetupList", 109, 0, "", "QuantityOnHand", 300, 5400, 1800, 300)
    $txt.Name = "QuantityOnHand"
}
Save-DesignForm $acc "SF_ItemSetupList"
Log "Item review list saved"

# Sales order lines: item code, unit, quick add
$frm = Open-DesignForm $acc "SF_SalesOrderLines"
$mod = $frm.Module
Replace-Snippet $mod "SELECT ItemID, ItemName " "SELECT ItemID, [InternalSKU] & ' - ' & [ItemName] " | Out-Null
Replace-Snippet $mod "SELECT UnitID, UnitCode " "SELECT UnitID, IIf(Len(Trim(Nz([UnitCode],'')))=0,[UnitName],[UnitCode]) " | Out-Null
Replace-Procedure $mod "Private Sub ItemID_NotInList(" "UX-ITEM-NOTINLIST-SO" @"
Private Sub ItemID_NotInList(NewData As String, Response As Integer)
    On Error GoTo Err_Handler
    Response = acDataErrContinue
    If HandleComboNotInList(Me!ItemID, NewData, "F_ItemSetup", "T_Items", "ItemID", "ItemName", "InternalSKU") Then
        Response = acDataErrAdded
    Else
        Me!ItemID.Undo
    End If
    Exit Sub
Err_Handler:
    Response = acDataErrContinue
    MsgBox Err.Description, vbExclamation, "Add Item"
End Sub
"@
if (-not (Module-HasText $mod "UX_FillLineFromItem")) {
    $start = Find-LineStarts $mod "Private Sub ItemID_AfterUpdate("
    $end = Get-ProcEnd $mod $start
    if ($end -gt 0) { $mod.InsertLines($end, "    UX_FillLineFromItem Me") }
}
Save-DesignForm $acc "SF_SalesOrderLines"
Log "Sales order lines saved"

# Invoice lines
$frm = Open-DesignForm $acc "SF_InvoiceLines"
$mod = $frm.Module
Replace-Snippet $mod "SELECT ItemID, ItemName " "SELECT ItemID, [InternalSKU] & ' - ' & [ItemName] " | Out-Null
Replace-Snippet $mod "SELECT UnitID, UnitCode " "SELECT UnitID, IIf(Len(Trim(Nz([UnitCode],'')))=0,[UnitName],[UnitCode]) " | Out-Null
if (Find-LineStarts $mod "Private Sub ItemID_NotInList(") {
    Replace-Procedure $mod "Private Sub ItemID_NotInList(" "UX-ITEM-NOTINLIST-INV" @"
Private Sub ItemID_NotInList(NewData As String, Response As Integer)
    On Error GoTo Err_Handler
    Response = acDataErrContinue
    If HandleComboNotInList(Me!ItemID, NewData, "F_ItemSetup", "T_Items", "ItemID", "ItemName", "InternalSKU") Then
        Response = acDataErrAdded
    Else
        Me!ItemID.Undo
    End If
    Exit Sub
Err_Handler:
    Response = acDataErrContinue
    MsgBox Err.Description, vbExclamation, "Add Item"
End Sub
"@
}
if (-not (Module-HasText $mod "UX_FillLineFromItem")) {
    $start = Find-LineStarts $mod "Private Sub ItemID_AfterUpdate("
    $end = Get-ProcEnd $mod $start
    if ($end -gt 0) { $mod.InsertLines($end, "    UX_FillLineFromItem Me") }
}
Save-DesignForm $acc "SF_InvoiceLines"
Log "Invoice lines saved"

# Customer PO lines
$frm = Open-DesignForm $acc "SF_CustomerPOLines"
$mod = $frm.Module
Replace-Snippet $mod "SELECT ItemID, ItemName " "SELECT ItemID, [InternalSKU] & ' - ' & [ItemName] " | Out-Null
Replace-Snippet $mod "SELECT UnitID, UnitCode " "SELECT UnitID, IIf(Len(Trim(Nz([UnitCode],'')))=0,[UnitName],[UnitCode]) " | Out-Null
if (Find-LineStarts $mod "Private Sub cboItemID_NotInList(") {
    Replace-Procedure $mod "Private Sub cboItemID_NotInList(" "UX-ITEM-NOTINLIST-PO" @"
Private Sub cboItemID_NotInList(NewData As String, Response As Integer)
    On Error GoTo Err_Handler
    Response = acDataErrContinue
    If HandleComboNotInList(Me.cboItemID, NewData, "F_ItemSetup", "T_Items", "ItemID", "ItemName", "InternalSKU") Then
        Response = acDataErrAdded
    Else
        Me.cboItemID.Undo
    End If
    Exit Sub
Err_Handler:
    Response = acDataErrContinue
    MsgBox Err.Description, vbExclamation, "Add Item"
End Sub
"@
}
$after = Find-LineStarts $mod "Private Sub cboItemID_AfterUpdate("
if ($after -eq 0) { $after = Find-LineStarts $mod "Private Sub ItemID_AfterUpdate(" }
if ($after -gt 0 -and -not (Module-HasText $mod "UX_FillLineFromItem")) {
    $end = Get-ProcEnd $mod $after
    if ($end -gt 0) { $mod.InsertLines($end, "    UX_FillLineFromItem Me") }
}
Save-DesignForm $acc "SF_CustomerPOLines"
Log "Customer PO lines saved"

# Supplier PO lines unit/item display
$frm = Open-DesignForm $acc "SF_SupplierPOLines"
$mod = $frm.Module
Replace-Snippet $mod "SELECT UnitID, UnitCode, UnitName " "SELECT UnitID, IIf(Len(Trim(Nz([UnitCode],'')))=0,[UnitName],[UnitCode]), UnitName " | Out-Null
if (-not (Module-HasText $mod "UX_FillLineFromItem")) {
    $after = Find-LineStarts $mod "Private Sub cboItemID_AfterUpdate("
    if ($after -gt 0) {
        $end = Get-ProcEnd $mod $after
        if ($end -gt 0) { $mod.InsertLines($end, "    UX_FillLineFromItem Me") }
    }
}
Save-DesignForm $acc "SF_SupplierPOLines"
Log "Supplier PO lines saved"

# Sales orders
$frm = Open-DesignForm $acc "F_SalesOrders"
$linkName = "SF_SalesOrderLines"
if (Test-Control $frm "subSalesOrderLines") { $linkName = "subSalesOrderLines" }
$sub = $frm.Controls($linkName)
$sub.LinkChildFields = "SalesOrderID"
$sub.LinkMasterFields = "SalesOrderID"
if ($linkName -ne "subSalesOrderLines") {
    foreach ($p in $sub.Properties) { if ($p.Name -eq "Name") { $p.Value = "subSalesOrderLines" } }
}
$mod = $frm.Module
Replace-ContainingLine $mod 'SafeSetRSField rsInvoice, "CustomerID", Nz(Me!CustomerID, Null)' '    SafeSetRSField rsInvoice, "CustomerID", UX_SalesOrderCustomerID(salesOrderIDValue)' | Out-Null
Replace-ContainingLine $mod 'SafeSetRSField rsInvoice, "BillToName", Nz(Me!ShipToName, vbNullString)' '    SafeSetRSField rsInvoice, "BillToName", UX_BillToName(UX_SalesOrderCustomerID(salesOrderIDValue))' | Out-Null
Replace-ContainingLine $mod 'SafeSetRSField rsInvoice, "BillToAddress", Nz(Me!ShipToAddress, vbNullString)' '    SafeSetRSField rsInvoice, "BillToAddress", UX_BillToAddress(UX_SalesOrderCustomerID(salesOrderIDValue))' | Out-Null
if (-not (Module-HasText $mod "existingInvoiceID = Nz(DLookup(")) {
    # guard lives in cmdCreateInvoice already; add a second guard inside the creator
}
if (-not (Module-HasText $mod 'newInvoiceID = Nz(DLookup("InvoiceID"')) {
    $idx = 0
    $n = $mod.CountOfLines
    for ($i = 1; $i -le $n; $i++) {
        if (([string]$mod.Lines($i, 1)).Contains('GetNextDocumentNumber("Invoice"')) { $idx = $i; break }
    }
    if ($idx -gt 0) {
        $mod.InsertLines($idx, @"
    newInvoiceID = Nz(DLookup("InvoiceID", "T_Invoices", "SalesOrderID=" & salesOrderIDValue), 0)
    If newInvoiceID > 0 Then
        CreateInvoiceFromCurrentSalesOrder = newInvoiceID
        Exit Function
    End If
"@)
    }
}
if (-not (Module-HasText $mod 'DoCmd.Close acForm, "F_Invoices", acSaveNo')) {
    $start = Find-LineStarts $mod "Private Sub OpenInvoiceByID("
    if ($start -gt 0) {
        $mod.InsertLines($start + 1, @"
    On Error Resume Next
    DoCmd.Close acForm, "F_Invoices", acSaveNo
    On Error GoTo Err_Handler
"@)
    }
}
Insert-AfterProcStart $mod "Private Function ShipSalesOrderStock(" "gSkipInventory = False" "    gSkipInventory = False"
# set skip once customer is known
if (-not (Module-HasText $mod "gSkipInventory = UX_IsInternalCustomer")) {
    $idx = 0
    $n = $mod.CountOfLines
    for ($i = 1; $i -le $n; $i++) {
        if (([string]$mod.Lines($i, 1)).Contains("customerIDValue = Nz(Me!CustomerID, Null)")) { $idx = $i; break }
    }
    if ($idx -gt 0) {
        $mod.InsertLines($idx + 1, "    gSkipInventory = UX_IsInternalCustomer(Nz(customerIDValue, 0))")
    }
}
Insert-AfterProcStart $mod "Private Sub IssueStockForSalesOrderLine(" "If gSkipInventory Then Exit Sub" "    If gSkipInventory Then Exit Sub"
Insert-AfterProcStart $mod "Private Sub Form_Load()" "UX_MarkRequired Me, ""CustomerID""" "    UX_MarkRequired Me, ""CustomerID"", ""OrderDate"""
Insert-AfterProcStart $mod "Private Sub Form_BeforeUpdate(" "UX_CoerceBooleans Me" "    UX_CoerceBooleans Me"
Add-ModuleCode $mod "Private Sub CustomerID_NotInList(" @"
Private Sub CustomerID_NotInList(NewData As String, Response As Integer)
    On Error GoTo Err_Handler
    Response = acDataErrContinue
    If HandleComboNotInList(Me!CustomerID, NewData, "F_CustomerSetup", "T_Customers", "CustomerID", "CustomerName", "CustomerCode") Then
        Response = acDataErrAdded
    Else
        Me!CustomerID.Undo
    End If
    Exit Sub
Err_Handler:
    Response = acDataErrContinue
    MsgBox Err.Description, vbExclamation, "Add Customer"
End Sub

Private Sub Form_Error(DataErr As Integer, Response As Integer)
    Dim uxHandled As Boolean
    UX_SuppressFormError DataErr, Response, uxHandled
    If uxHandled Then Exit Sub
    Response = acDataErrDisplay
End Sub

Private Sub Form_Unload(Cancel As Integer)
    On Error Resume Next
    gSkipInventory = False
    If Me.NewRecord And Me.Dirty Then Me.Undo
End Sub
"@
if (Test-Control $frm "CustomerID") { $frm.Controls("CustomerID").OnNotInList = "[Event Procedure]" }
$top = 80
Add-Button $acc $frm "F_SalesOrders" "cmdUXBack" "Back" "=UX_GoBack()" 200 $top 1400 340 1
Add-Button $acc $frm "F_SalesOrders" "cmdPrintQuote" "Print Quotation" "=UX_PrintActive(""R_Quotation"",""SalesOrderID"",""Quotation"")" 1700 $top 2200 340 1
Add-Button $acc $frm "F_SalesOrders" "cmdPrintSO" "Print Sales Order" "=UX_PrintActive(""R_SalesOrder"",""SalesOrderID"",""Sales Order"")" 4000 $top 2400 340 1
Save-DesignForm $acc "F_SalesOrders"
Log "Sales orders saved"

# Invoices
$frm = Open-DesignForm $acc "F_Invoices"
$mod = $frm.Module
$idx = Find-ExactLine $mod "If m_IsExplicitSave = False Then"
if ($idx -gt 0) {
    $mod.DeleteLines($idx, 4)
    Log "Removed invoice explicit-save block"
}
Insert-AfterProcStart $mod "Private Sub Form_Error(" "UX_SuppressFormError DataErr" @"
    Dim uxHandled As Boolean
    UX_SuppressFormError DataErr, Response, uxHandled
    If uxHandled Then Exit Sub
"@
Insert-AfterProcStart $mod "Private Sub Form_BeforeUpdate(" "UX_CoerceBooleans Me" "    UX_CoerceBooleans Me"
Insert-AfterProcStart $mod "Private Sub Form_Load()" "UX_MarkRequired Me, ""CustomerID""" "    UX_MarkRequired Me, ""CustomerID"", ""InvoiceDate"""
Add-ModuleCode $mod "Private Sub CustomerID_NotInList(" @"
Private Sub CustomerID_NotInList(NewData As String, Response As Integer)
    On Error GoTo Err_Handler
    Response = acDataErrContinue
    If HandleComboNotInList(Me!CustomerID, NewData, "F_CustomerSetup", "T_Customers", "CustomerID", "CustomerName", "CustomerCode") Then
        Response = acDataErrAdded
    Else
        Me!CustomerID.Undo
    End If
    Exit Sub
Err_Handler:
    Response = acDataErrContinue
    MsgBox Err.Description, vbExclamation, "Add Customer"
End Sub

Private Sub Form_Unload(Cancel As Integer)
    On Error Resume Next
    If Me.NewRecord And Me.Dirty Then Me.Undo
End Sub
"@
if (Test-Control $frm "CustomerID") { $frm.Controls("CustomerID").OnNotInList = "[Event Procedure]" }
Add-Button $acc $frm "F_Invoices" "cmdUXBack" "Back" "=UX_GoBack()" 200 80 1400 340 1
if (Test-Control $frm "cmdMakePDF") {
    $src = $frm.Controls("cmdMakePDF")
    Add-Button $acc $frm "F_Invoices" "cmdPrintDN" "Print Delivery Note" "=UX_PrintActive(""R_DeliveryNote"",""InvoiceID"",""Delivery Note"")" $src.Left ($src.Top + $src.Height + 80) $src.Width $src.Height $src.Section
}
Save-DesignForm $acc "F_Invoices"
Log "Invoices saved"

# Customer PO
$frm = Open-DesignForm $acc "F_CustomerPO"
$mod = $frm.Module
Replace-Procedure $mod "Private Sub txtPONumber_BeforeUpdate(" "PO number is validated when the record is saved." @"
Private Sub txtPONumber_BeforeUpdate(Cancel As Integer)
    'PO number is validated when the record is saved.
    'Do not trap the user in this field while they are still entering the PO.
End Sub
"@
if (-not (Module-HasText $mod "<> subtotalAmountValue Then Me.txtSubtotalAmount")) {
    Replace-ContainingLine $mod "Me.txtSubtotalAmount = subtotalAmountValue" "    If Nz(Me.txtSubtotalAmount, 0) <> subtotalAmountValue Then Me.txtSubtotalAmount = subtotalAmountValue" | Out-Null
    Replace-ContainingLine $mod "Me.txtTotalAmount = totalAmountValue" "    If Nz(Me.txtTotalAmount, 0) <> totalAmountValue Then Me.txtTotalAmount = totalAmountValue" | Out-Null
}
Insert-AfterProcStart $mod "Private Sub Form_BeforeUpdate(" "UX_CoerceBooleans Me" "    UX_CoerceBooleans Me"
Insert-AfterProcStart $mod "Private Sub Form_Load()" "UX_MarkRequired Me, ""cboCustomerID""" "    UX_MarkRequired Me, ""cboCustomerID"", ""txtPODate"""
Add-ModuleCode $mod "Private Sub cboCustomerID_NotInList(" @"
Private Sub cboCustomerID_NotInList(NewData As String, Response As Integer)
    On Error GoTo Err_Handler
    Response = acDataErrContinue
    If HandleComboNotInList(Me.cboCustomerID, NewData, "F_CustomerSetup", "T_Customers", "CustomerID", "CustomerName", "CustomerCode") Then
        Response = acDataErrAdded
    Else
        Me.cboCustomerID.Undo
    End If
    Exit Sub
Err_Handler:
    Response = acDataErrContinue
    MsgBox Err.Description, vbExclamation, "Add Customer"
End Sub

Private Sub Form_Error(DataErr As Integer, Response As Integer)
    Dim uxHandled As Boolean
    UX_SuppressFormError DataErr, Response, uxHandled
    If uxHandled Then Exit Sub
    Response = acDataErrDisplay
End Sub

Private Sub Form_Unload(Cancel As Integer)
    On Error Resume Next
    If Me.NewRecord And Me.Dirty Then Me.Undo
End Sub
"@
if (Test-Control $frm "cboCustomerID") { $frm.Controls("cboCustomerID").OnNotInList = "[Event Procedure]" }
Add-Button $acc $frm "F_CustomerPO" "cmdUXBack" "Back" "=UX_GoBack()" 200 80 1400 340 1
Add-Button $acc $frm "F_CustomerPO" "cmdPrintPO" "Print P.O." "=UX_PrintActive(""R_PurchaseOrder"",""txtCustomerPOID"",""Purchase Order"")" 1700 80 1800 340 1
Add-Button $acc $frm "F_CustomerPO" "cmdUploadPO" "Upload Client PO" "=UX_AttachToCurrentForm(""txtCustomerPOID"",""CustomerPO"")" 3600 80 2400 340 1
Save-DesignForm $acc "F_CustomerPO"
Log "Customer PO saved"

# Send PO / supplier PO
$frm = Open-DesignForm $acc "F_SendPO"
$mod = $frm.Module
Insert-AfterProcStart $mod "Private Sub Form_BeforeUpdate(" "UX_CoerceBooleans Me" "    UX_CoerceBooleans Me"
if (-not (Module-HasText $mod "Private Sub Form_Error(")) {
    Add-ModuleCode $mod "UX-SENDPO-ERR" @"
Private Sub Form_Error(DataErr As Integer, Response As Integer)
    Dim uxHandled As Boolean
    UX_SuppressFormError DataErr, Response, uxHandled
    If uxHandled Then Exit Sub
    Response = acDataErrDisplay
End Sub
"@
} else {
    Insert-AfterProcStart $mod "Private Sub Form_Error(" "UX_SuppressFormError DataErr" @"
    Dim uxHandled As Boolean
    UX_SuppressFormError DataErr, Response, uxHandled
    If uxHandled Then Exit Sub
"@
}
Add-Button $acc $frm "F_SendPO" "cmdUXBack" "Back" "=UX_GoBack()" 200 80 1400 340 1
Add-Button $acc $frm "F_SendPO" "cmdPrintSupplierPO" "Print P.O." "=UX_PrintActive(""R_SupplierPO"",""txtSupplierPOID"",""Purchase Order"")" 1700 80 1800 340 1
Add-Button $acc $frm "F_SendPO" "cmdUploadSupplierPO" "Upload PO" "=UX_AttachToCurrentForm(""txtSupplierPOID"",""SupplierPO"")" 3600 80 1800 340 1
Save-DesignForm $acc "F_SendPO"
Log "Supplier PO saved"

# Projects
$frm = Open-DesignForm $acc "F_Projects"
$mod = $frm.Module
Insert-AfterProcStart $mod "Private Sub Form_BeforeUpdate(" "UX_CoerceBooleans Me" "    UX_CoerceBooleans Me"
if (-not (Module-HasText $mod "Private Sub Form_Error(")) {
    Add-ModuleCode $mod "UX-PROJ-ERR" @"
Private Sub Form_Error(DataErr As Integer, Response As Integer)
    Dim uxHandled As Boolean
    UX_SuppressFormError DataErr, Response, uxHandled
    If uxHandled Then Exit Sub
    Response = acDataErrDisplay
End Sub

Private Sub Form_Unload(Cancel As Integer)
    On Error Resume Next
    gSkipInventory = False
    If Me.NewRecord And Me.Dirty Then Me.Undo
End Sub
"@
}
Add-Button $acc $frm "F_Projects" "cmdUXBack" "Back" "=UX_GoBack()" 200 80 1400 340 1
Save-DesignForm $acc "F_Projects"
Log "Projects saved"

# Project issue inventory skip
$frm = Open-DesignForm $acc "F_ProjectIssue"
$mod = $frm.Module
Insert-AfterProcStart $mod "Private Sub cmdIssueSelected_Click()" "UX-PROJ-SKIP" "    gSkipInventory = UX_ProjectSkipsInventory(mProjectID) ' UX-PROJ-SKIP"
Insert-AfterProcStart $mod "Private Sub IssueBulkProjectItem(" "UX-SKIP-BULK" "    If gSkipInventory Then Exit Sub ' UX-SKIP-BULK"
Insert-AfterProcStart $mod "Private Sub ReturnBulkProjectItem(" "UX-SKIP-RETBULK" "    If gSkipInventory Then Exit Sub ' UX-SKIP-RETBULK"
Insert-AfterProcStart $mod "Private Sub IssueSerializedProjectItem(" "UX-SKIP-SER" "    If gSkipInventory Then Exit Sub ' UX-SKIP-SER"
Insert-AfterProcStart $mod "Private Sub ReturnSerializedProjectItem(" "UX-SKIP-RETSER" "    If gSkipInventory Then Exit Sub ' UX-SKIP-RETSER"
Add-Button $acc $frm "F_ProjectIssue" "cmdUXBack" "Back" "=UX_GoBack()" 200 80 1400 340 1
Save-DesignForm $acc "F_ProjectIssue"
Log "Project issue saved"

# Invoice list can always open a row
$frm = Open-DesignForm $acc "SF_InvoiceList"
$frm.OnDblClick = "=OpenSelectedRecord()"
foreach ($ctl in @($frm.Controls)) {
    try { if ($ctl.ControlType -eq 109 -or $ctl.ControlType -eq 111) { $ctl.OnDblClick = "=OpenSelectedRecord()" } } catch {}
}
Save-DesignForm $acc "SF_InvoiceList"
$frm = Open-DesignForm $acc "F_InvoiceList"
Add-Button $acc $frm "F_InvoiceList" "cmdOpenSelected" "Open Invoice" "=UX_OpenSubformRecord(""subInvoiceList"")" 200 80 2000 360 1
Add-Button $acc $frm "F_InvoiceList" "cmdUXBack" "Back" "=UX_GoBack()" 2300 80 1400 360 1
Save-DesignForm $acc "F_InvoiceList"
Log "Invoice list saved"

# Sales order list open
$frm = Open-DesignForm $acc "SF_SalesOrderList"
$frm.OnDblClick = "=OpenSelectedRecord()"
foreach ($ctl in @($frm.Controls)) {
    try { if ($ctl.ControlType -eq 109 -or $ctl.ControlType -eq 111) { $ctl.OnDblClick = "=OpenSelectedRecord()" } } catch {}
}
Save-DesignForm $acc "SF_SalesOrderList"
$frm = Open-DesignForm $acc "F_SalesOrderList"
Add-Button $acc $frm "F_SalesOrderList" "cmdUXBack" "Back" "=UX_GoBack()" 200 80 1400 360 1
Save-DesignForm $acc "F_SalesOrderList"

# Dashboard startup
$frm = Open-DesignForm $acc "Dashboard"
$mod = $frm.Module
Insert-AfterProcStart $mod "Private Sub Form_Load()" "UX_Startup" "    UX_Startup"
Save-DesignForm $acc "Dashboard"
Log "Dashboard saved"

# Back buttons on the remaining main forms
$skip = @{
    "F_Login" = $true
    "F_Warehouses" = $true
    "F_LocationSetup" = $true
    "F_StockIn" = $true
    "F_StockOut" = $true
    "F_ItemSetup" = $true
    "F_SalesOrders" = $true
    "F_Invoices" = $true
    "F_CustomerPO" = $true
    "F_SendPO" = $true
    "F_Projects" = $true
    "F_ProjectIssue" = $true
    "F_InvoiceList" = $true
    "F_SalesOrderList" = $true
    "Dashboard" = $true
}
foreach ($f in $acc.CurrentProject.AllForms) {
    $name = $f.Name
    if ($name.StartsWith("SF_")) { continue }
    if ($skip.ContainsKey($name)) { continue }
    try {
        $frm = Open-DesignForm $acc $name
        Add-Button $acc $frm $name "cmdUXBack" "Back" "=UX_GoBack()" 200 60 1400 340 1
        Save-DesignForm $acc $name
        Log "Back button on $name"
    } catch {
        Log "Back button skipped $name : $($_.Exception.Message)"
        try { $acc.DoCmd.Close(2, $name, 2) } catch {}
    }
}

# ---------- Clear business data and seed ----------
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

Log "Seeding units and internal customer"
$acc.Run("UX_EnsureStandardUnits")
$acc.Run("UX_EnsureInternalCustomer")
$acc.Run("UX_BackfillUnitCodes")

$cust = $db.OpenRecordset("SELECT CustomerID, CustomerCode, CustomerName, CustomerType FROM T_Customers")
if (-not $cust.EOF) { Log "Customer seed: $($cust.Fields.Item('CustomerCode').Value) / $($cust.Fields.Item('CustomerName').Value)" }
$cust.Close()
$units = $db.OpenRecordset("SELECT Count(*) AS N FROM T_Units")
Log "Units remaining: $($units.Fields.Item('N').Value)"
$units.Close()
$items = $db.OpenRecordset("SELECT Count(*) AS N FROM T_Items")
Log "Items remaining: $($items.Fields.Item('N').Value)"
$items.Close()

Log "DONE"
$acc.DoCmd.SetWarnings($true)
$acc.CloseCurrentDatabase()
$acc.Quit()
} catch {
    Log "FATAL $($_.Exception.Message)"
    try { $acc.Quit() } catch {}
    throw
}
