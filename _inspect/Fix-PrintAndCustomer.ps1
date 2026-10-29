$ErrorActionPreference = "Stop"
$root = "C:\Users\user\Downloads\LLC"
function Log($m) { Write-Host $m }

Get-Process MSACCESS -ErrorAction SilentlyContinue | Stop-Process -Force -ErrorAction SilentlyContinue
Start-Sleep -Seconds 1
$acc = New-Object -ComObject Access.Application
$acc.Visible = $false
$acc.OpenCurrentDatabase("$root\LLC.accdb")
$acc.DoCmd.SetWarnings($false)

function Set-Query($db, $name, $sql) {
    $found = $false
    foreach ($q in $db.QueryDefs) {
        if ($q.Name -eq $name) { $q.SQL = $sql; $found = $true }
    }
    if (-not $found) { [void]$db.CreateQueryDef($name, $sql) }
}

function Test-Control($obj, $name) {
    try { $null = $obj.Controls.Item($name); return $true } catch { return $false }
}

try {
    Log "Loading M_UX"
    $acc.LoadFromText(5, "M_UX", "$root\_inspect\M_UX.txt")

    $itemText = "IIf(Len(Trim(Nz(I.ItemName,'')))=0,Nz(L.LineDescription,''),IIf(Len(Trim(Nz(I.InternalSKU,'')))=0,I.ItemName,I.InternalSKU & ' - ' & I.ItemName))"
    $unitText = "IIf(Len(Trim(Nz(U.UnitCode,'')))=0,Nz(U.UnitName,''),U.UnitCode)"
    $db = $acc.CurrentDb()

    Set-Query $db "Q_SalesOrderReportLines" @"
SELECT L.SalesOrderLineID AS InvoiceLineID, L.SalesOrderID AS InvoiceID, L.ItemID, Nz(I.InternalSKU,'') AS ItemCode, Nz(L.LineDescription,'') AS LineDescription, Nz(L.QuantityOrdered,0) AS Quantity, L.UnitID, $unitText AS UnitCode, Nz(L.UnitPrice,0) AS UnitPrice, L.Taxable, Nz(L.LineSubtotal,0) AS LineSubtotal, Trim($itemText) AS ReportItem, Nz(I.ItemName,'') AS ReportDescription, $unitText AS ReportUnit
FROM (T_SalesOrderLines AS L LEFT JOIN T_Items AS I ON L.ItemID = I.ItemID) LEFT JOIN T_Units AS U ON L.UnitID = U.UnitID;
"@
    Set-Query $db "Q_CustomerPOReportLines" @"
SELECT L.CustomerPOLineID AS InvoiceLineID, L.CustomerPOID AS InvoiceID, L.ItemID, Nz(I.InternalSKU,Nz(L.CustomerPartNumber,'')) AS ItemCode, Nz(L.LineDescription,'') AS LineDescription, Nz(L.QuantityOrdered,0) AS Quantity, L.UnitID, $unitText AS UnitCode, Nz(L.UnitPrice,0) AS UnitPrice, L.Taxable, Nz(L.LineSubtotal,0) AS LineSubtotal, Trim($itemText) AS ReportItem, Nz(I.ItemName,'') AS ReportDescription, $unitText AS ReportUnit
FROM (T_CustomerPOLines AS L LEFT JOIN T_Items AS I ON L.ItemID = I.ItemID) LEFT JOIN T_Units AS U ON L.UnitID = U.UnitID;
"@
    Set-Query $db "Q_SupplierPOReportLines" @"
SELECT L.SupplierPOLineID AS InvoiceLineID, L.SupplierPOID AS InvoiceID, L.ItemID, Nz(I.InternalSKU,Nz(L.SupplierPartNumber,'')) AS ItemCode, Nz(L.ItemDescription,'') AS LineDescription, Nz(L.OrderedQuantity,0) AS Quantity, L.UnitID, $unitText AS UnitCode, Nz(L.UnitPrice,0) AS UnitPrice, False AS Taxable, Nz(L.OrderedQuantity,0)*Nz(L.UnitPrice,0) AS LineSubtotal, Trim(IIf(Len(Trim(Nz(I.ItemName,'')))=0,Nz(L.ItemDescription,''),IIf(Len(Trim(Nz(I.InternalSKU,'')))=0,I.ItemName,I.InternalSKU & ' - ' & I.ItemName))) AS ReportItem, Nz(I.ItemName,'') AS ReportDescription, $unitText AS ReportUnit
FROM (T_SupplierPOLines AS L LEFT JOIN T_Items AS I ON L.ItemID = I.ItemID) LEFT JOIN T_Units AS U ON L.UnitID = U.UnitID;
"@
    Set-Query $db "Q_InvoiceReportLines" @"
SELECT L.InvoiceLineID, L.InvoiceID, L.ItemID, L.ItemCode, L.LineDescription, L.Quantity, L.UnitID, IIf(Len(Trim(Nz(U.UnitCode,'')))=0,Nz(U.UnitName,''),U.UnitCode) AS UnitCode, L.UnitPrice, L.Taxable, L.LineSubtotal, IIf(Len(Trim(Nz(I.ItemName,'')))=0,Nz(L.LineDescription,L.ItemCode),IIf(Len(Trim(Nz(I.InternalSKU,L.ItemCode)))=0,I.ItemName,Nz(I.InternalSKU,L.ItemCode) & ' - ' & I.ItemName)) AS ReportItem, Nz(I.ItemName,L.LineDescription) AS ReportDescription, IIf(Len(Trim(Nz(U.UnitCode,'')))=0,Nz(U.UnitName,''),U.UnitCode) AS ReportUnit
FROM (T_InvoiceLines AS L LEFT JOIN T_Items AS I ON L.ItemID = I.ItemID) LEFT JOIN T_Units AS U ON L.UnitID = U.UnitID
ORDER BY L.InvoiceLineID;
"@
    Log "Line queries saved"

    function Tune-LineReport($name) {
        while ($acc.CurrentProject.AllReports.Item($name).IsLoaded) { $acc.DoCmd.Close(3, $name, 2) }
        $acc.DoCmd.OpenReport($name, 1)
        $rpt = $acc.Reports.Item($name)
        if (-not (Test-Control $rpt "ReportUnit")) {
            $txt = $acc.CreateReportControl($name, 109, 0, "", "ReportUnit", 4800, 0, 1100, 324)
            $txt.Name = "ReportUnit"
        }
        $u = $rpt.Controls.Item("ReportUnit")
        $u.ControlSource = "ReportUnit"
        $u.Left = 4800; $u.Width = 1100; $u.Height = 324; $u.Top = 0
        $u.TextAlign = 2
        $item = $rpt.Controls.Item("ReportItem")
        $item.Left = 120; $item.Width = 4600; $item.Height = 324; $item.CanGrow = $true
        $price = $rpt.Controls.Item("UnitPrice")
        $price.Left = 5980; $price.Width = 1400; $price.TextAlign = 3; $price.Format = "#,##0.00"
        $qty = $rpt.Controls.Item("Quantity")
        $qty.Left = 7460; $qty.Width = 1100; $qty.TextAlign = 3; $qty.Format = "#,##0.00"
        if (Test-Control $rpt "Check13") { $chk = $rpt.Controls.Item("Check13"); $chk.Left = 8700; $chk.Width = 400 }
        if (Test-Control $rpt "LineSubtotal") {
            $tot = $rpt.Controls.Item("LineSubtotal")
            $tot.Left = 9200; $tot.Width = 1800; $tot.TextAlign = 3; $tot.Format = "#,##0.00"
        }
        try { $rpt.Section(0).CanGrow = $true } catch {}
        $acc.DoCmd.Close(3, $name, 1)
        Log "Line report $name"
    }

    foreach ($n in @("R_InvoiceLines","R_SalesDocLines","R_POLines","R_SupplierPOLinesRpt")) {
        Tune-LineReport $n
    }

    function Tune-Header($name) {
        while ($acc.CurrentProject.AllReports.Item($name).IsLoaded) { $acc.DoCmd.Close(3, $name, 2) }
        $acc.DoCmd.OpenReport($name, 1)
        $rpt = $acc.Reports.Item($name)
        if (Test-Control $rpt "Label10") {
            $lbl = $rpt.Controls.Item("Label10")
            $cap = [string]$lbl.Caption
            $cap = $cap.Replace(" #", "").Trim()
            if ($cap.Length -gt 0) { $lbl.Caption = $cap }
            $lbl.Width = 3200
            $lbl.Height = 420
        }
        if (Test-Control $rpt "Text11" -and (Test-Control $rpt "Label10")) {
            $num = $rpt.Controls.Item("Text11")
            $lbl = $rpt.Controls.Item("Label10")
            $num.Left = $lbl.Left + $lbl.Width + 80
            $num.Width = 2800
            $num.Height = 420
            $num.Top = $lbl.Top
            $num.FontSize = 14
        }
        $headers = @{
            "ReportItem_Label" = @{ Caption = "Item"; Left = 180; Width = 4500 }
            "UnitPrice_Label" = @{ Caption = "Price"; Left = 6040; Width = 1400 }
            "Quantity_Label" = @{ Caption = "Qty"; Left = 7520; Width = 1100 }
            "Taxable_Label" = @{ Caption = "Tax"; Left = 8640; Width = 700 }
            "LineSubtotal_Label" = @{ Caption = "Amount"; Left = 9400; Width = 1800 }
        }
        foreach ($key in $headers.Keys) {
            if (Test-Control $rpt $key) {
                $c = $rpt.Controls.Item($key)
                $c.Caption = $headers[$key].Caption
                $c.Left = $headers[$key].Left
                $c.Width = $headers[$key].Width
            }
        }
        if (-not (Test-Control $rpt "Unit_Label")) {
            $lab = $acc.CreateReportControl($name, 100, 0, "", "", 4860, 5400, 1100, 324)
            $lab.Name = "Unit_Label"
            $lab.Caption = "Unit"
            $lab.FontWeight = 700
            $lab.TextAlign = 2
            $lab.ForeColor = 0
            if (Test-Control $rpt "Quantity_Label") { $lab.Top = $rpt.Controls.Item("Quantity_Label").Top }
        } else {
            $lab = $rpt.Controls.Item("Unit_Label")
            $lab.Caption = "Unit"
            $lab.Left = 4860
            $lab.Width = 1100
        }
        foreach ($ctl in @($rpt.Controls)) {
            try {
                if ($ctl.ControlType -eq 103 -or $ctl.ControlType -eq 110) { $ctl.SizeMode = 3 }
            } catch {}
        }
        $acc.DoCmd.Close(3, $name, 1)
        Log "Header $name"
    }

    foreach ($n in @("R_Invoice","R_Quotation","R_SalesOrder","R_PurchaseOrder","R_SupplierPO","R_DeliveryNote")) {
        Tune-Header $n
    }

    function Add-CustomerButton($formName, $comboName) {
        while ($acc.CurrentProject.AllForms.Item($formName).IsLoaded) { $acc.DoCmd.Close(2, $formName, 2) }
        $acc.DoCmd.OpenForm($formName, 1)
        $frm = $acc.Forms.Item($formName)
        if (-not (Test-Control $frm $comboName)) { $acc.DoCmd.Close(2, $formName, 2); Log "No combo $comboName on $formName"; return }
        $cbo = $frm.Controls.Item($comboName)
        $cbo.LimitToList = $true
        $cbo.OnNotInList = "[Event Procedure]"
        $cbo.ControlTipText = "Type a new customer name to add it, or click Add Customer."
        foreach ($ctl in @($frm.Controls)) {
            try {
                if ($ctl.Name -eq ($comboName + "_Label") -or $ctl.Name -eq "CustomerID_Label") {
                    if ($ctl.Caption -like "*Customer*") { $ctl.Caption = "Customer" }
                }
            } catch {}
        }
        if (-not (Test-Control $frm "cmdAddCustomer")) {
            $btn = $acc.CreateControl($formName, 104, $cbo.Section, "", "", ($cbo.Left + $cbo.Width + 120), $cbo.Top, 2100, $cbo.Height)
            $btn.Name = "cmdAddCustomer"
            $btn.Caption = "Add Customer"
            $btn.OnClick = "=UX_AddCustomer(""$comboName"")"
        } else {
            $frm.Controls.Item("cmdAddCustomer").OnClick = "=UX_AddCustomer(""$comboName"")"
            $frm.Controls.Item("cmdAddCustomer").Caption = "Add Customer"
        }
        $acc.DoCmd.Close(2, $formName, 1)
        Log "Add Customer on $formName"
    }

    Add-CustomerButton "F_SalesOrders" "CustomerID"
    Add-CustomerButton "F_Invoices" "CustomerID"
    Add-CustomerButton "F_CustomerPO" "cboCustomerID"
    Add-CustomerButton "F_Projects" "cboCustomerID"

    $acc.DoCmd.RunCommand(126)
    Log "Compiled"
    $acc.CloseCurrentDatabase()
    $acc.Quit()
    Log "DONE"
} catch {
    Log "FATAL $($_.Exception.Message)"
    try { $acc.Quit() } catch {}
    throw
}
