$ErrorActionPreference = "Stop"
$root = "C:\Users\user\Downloads\LLC"
function Log($m) { Write-Host $m }
$bg = 16316664
$white = 16777215
$ink = 3947580
$teal = 7241006
$soft = 16119283
$muted = 7895160

Get-Process MSACCESS -ErrorAction SilentlyContinue | Stop-Process -Force -ErrorAction SilentlyContinue
Start-Sleep -Seconds 1
$acc = New-Object -ComObject Access.Application
$acc.Visible = $false
$acc.OpenCurrentDatabase("$root\LLC.accdb")
$acc.DoCmd.SetWarnings($false)

function Test-Control($obj, $name) {
    try { $null = $obj.Controls.Item($name); return $true } catch { return $false }
}
function Form-Exists($acc, $name) {
    foreach ($f in $acc.CurrentProject.AllForms) { if ($f.Name -eq $name) { return $true } }
    return $false
}
function Table-Exists($db, $name) {
    foreach ($t in $db.TableDefs) { if ($t.Name -eq $name) { return $true } }
    return $false
}

try {
    Log "Loading M_UX"
    $acc.LoadFromText(5, "M_UX", "$root\_inspect\M_UX.txt")
    $db = $acc.CurrentDb()

    function Add-Field($td, $name, $type, $size, $auto) {
        $f = $td.CreateField($name, $type)
        if ($size -gt 0) { $f.Size = $size }
        if ($auto) { $f.Attributes = 17 }
        $td.Fields.Append($f)
    }
    function New-Table($db, $name, $script) {
        if (Table-Exists $db $name) { Log "Table $name exists"; return }
        $td = $db.CreateTableDef($name)
        & $script $td
        $db.TableDefs.Append($td)
        $td = $db.TableDefs.Item($name)
        $idx = $td.CreateIndex("PrimaryKey")
        $idx.Primary = $true
        $idx.Unique = $true
        $idf = $idx.CreateField(($td.Fields.Item(0).Name))
        $idx.Fields.Append($idf)
        $td.Indexes.Append($idx)
        Log "Created $name"
    }

    New-Table $db "T_Expenses" { param($td)
        Add-Field $td "ExpenseID" 4 0 $true
        Add-Field $td "ExpenseDate" 8 0 $false
        Add-Field $td "Category" 10 50 $false
        Add-Field $td "Description" 10 255 $false
        Add-Field $td "Amount" 5 0 $false
        Add-Field $td "PaidTo" 10 120 $false
        Add-Field $td "Notes" 12 0 $false
    }
    New-Table $db "T_Staff" { param($td)
        Add-Field $td "StaffID" 4 0 $true
        Add-Field $td "StaffName" 10 120 $false
        Add-Field $td "JobTitle" 10 80 $false
        Add-Field $td "Phone" 10 40 $false
        Add-Field $td "Status" 10 20 $false
    }
    New-Table $db "T_Attendance" { param($td)
        Add-Field $td "AttendanceID" 4 0 $true
        Add-Field $td "StaffID" 4 0 $false
        Add-Field $td "WorkDate" 8 0 $false
        Add-Field $td "TimeIn" 8 0 $false
        Add-Field $td "TimeOut" 8 0 $false
        Add-Field $td "Notes" 10 255 $false
    }
    New-Table $db "T_Leave" { param($td)
        Add-Field $td "LeaveID" 4 0 $true
        Add-Field $td "StaffID" 4 0 $false
        Add-Field $td "LeaveType" 10 30 $false
        Add-Field $td "StartDate" 8 0 $false
        Add-Field $td "EndDate" 8 0 $false
        Add-Field $td "Days" 7 0 $false
        Add-Field $td "Status" 10 20 $false
        Add-Field $td "Notes" 12 0 $false
    }

    $n = $db.OpenRecordset("SELECT Count(*) AS N FROM T_Staff")
    if ([int]$n.Fields.Item("N").Value -eq 0) {
        $db.Execute("INSERT INTO T_Staff (StaffName, JobTitle, Phone, Status) VALUES ('Store Keeper', 'Stores', '000', 'Active')")
        $db.Execute("INSERT INTO T_Expenses (ExpenseDate, Category, Description, Amount, PaidTo, Notes) VALUES (Date(), 'Supplies', 'Sample packing supplies', 25, 'Local shop', 'Sample expense')")
        $db.Execute("INSERT INTO T_Attendance (StaffID, WorkDate, TimeIn, Notes) SELECT StaffID, Date(), Time(), 'Signed in' FROM T_Staff WHERE StaffName='Store Keeper'")
        $db.Execute("INSERT INTO T_Leave (StaffID, LeaveType, StartDate, EndDate, Days, Status, Notes) SELECT StaffID, 'Off day', Date()+7, Date()+7, 1, 'Approved', 'Sample off day' FROM T_Staff WHERE StaffName='Store Keeper'")
        Log "Sample accounting rows added"
    }
    $n.Close()

    function Set-Query($db, $name, $sql) {
        $found = $false
        foreach ($q in $db.QueryDefs) { if ($q.Name -eq $name) { $q.SQL = $sql; $found = $true } }
        if (-not $found) { [void]$db.CreateQueryDef($name, $sql) }
    }
    Set-Query $db "Q_ExpenseByCategory" "SELECT Category, Count(*) AS Items, Sum(Nz(Amount,0)) AS TotalAmount FROM T_Expenses GROUP BY Category ORDER BY Category;"
    Set-Query $db "Q_AttendanceList" "SELECT A.AttendanceID, A.WorkDate, S.StaffName, A.TimeIn, A.TimeOut, A.Notes FROM T_Attendance AS A LEFT JOIN T_Staff AS S ON A.StaffID = S.StaffID ORDER BY A.WorkDate DESC;"
    Set-Query $db "Q_LeaveList" "SELECT L.LeaveID, S.StaffName, L.LeaveType, L.StartDate, L.EndDate, L.Days, L.Status, L.Notes FROM T_Leave AS L LEFT JOIN T_Staff AS S ON L.StaffID = S.StaffID ORDER BY L.StartDate DESC;"

    function New-SavedForm($acc, $finalName) {
        if (Form-Exists $acc $finalName) { return }
        $acc.DoCmd.CopyObject("", $finalName, 2, "F_Accounting")
        Log "Form shell $finalName"
    }
    foreach ($fn in @("F_Accounting","F_Expenses","F_HR","F_Statistics","SF_ExpenseList","SF_StaffList","SF_Attendance","SF_Leave","SF_ExpenseStats")) {
        New-SavedForm $acc $fn
    }

    function Open-Design($acc, $name) {
        while ($acc.CurrentProject.AllForms.Item($name).IsLoaded) { $acc.DoCmd.Close(2, $name, 2) }
        $acc.DoCmd.OpenForm($name, 1)
        return ,$acc.Forms.Item($name)
    }
    function Add-Label($acc, $form, $name, $caption, $left, $top, $width, $height, $size) {
        $frm = $acc.Forms.Item($form)
        if (Test-Control $frm $name) { $frm.Controls.Item($name).Caption = $caption; return }
        $c = $acc.CreateControl([string]$form, 100, 0, "", "", [int]$left, [int]$top, [int]$width, [int]$height)
        $c.Name = [string]$name
        $c.Caption = $caption
        $c.FontSize = $size
        $c.FontWeight = 700
        $c.ForeColor = $ink
    }
    function Add-Text($acc, $form, $name, $left, $top, $width, $height) {
        $frm = $acc.Forms.Item($form)
        if (Test-Control $frm $name) { return }
        $c = $acc.CreateControl($form, 109, 0, "", "", $left, $top, $width, $height)
        $c.Name = $name
    }
    function Add-Combo($acc, $form, $name, $left, $top, $width, $height, $rowSource, $valueList) {
        $frm = $acc.Forms.Item($form)
        if (Test-Control $frm $name) { return }
        $c = $acc.CreateControl($form, 111, 0, "", "", $left, $top, $width, $height)
        $c.Name = $name
        if ($valueList) { $c.RowSourceType = "Value List" } else { $c.RowSourceType = "Table/Query" }
        $c.RowSource = $rowSource
        $c.ColumnCount = 2
        $c.ColumnWidths = "0;6cm"
        $c.BoundColumn = 1
    }
    function Add-Button($acc, $form, $name, $caption, $click, $left, $top, $width, $height) {
        $frm = $acc.Forms.Item($form)
        if (Test-Control $frm $name) {
            $b = $frm.Controls.Item($name)
            $b.Caption = $caption
            $b.OnClick = $click
            $b.Width = $width
            $b.Height = $height
            return
        }
        $c = $acc.CreateControl($form, 104, 0, "", "", $left, $top, $width, $height)
        $c.Name = $name
        $c.Caption = $caption
        $c.OnClick = $click
    }
    function Add-Sub($acc, $form, $name, $source, $left, $top, $width, $height) {
        $frm = $acc.Forms.Item($form)
        if (Test-Control $frm $name) { $frm.Controls.Item($name).SourceObject = $source; return }
        $c = $acc.CreateControl($form, 112, 0, "", "", $left, $top, $width, $height)
        $c.Name = $name
        $c.SourceObject = $source
    }
    function Style-Button($btn, $quiet) {
        try { $btn.UseTheme = 0 } catch {}
        $btn.FontName = "Calibri"
        $btn.FontSize = 10
        if ($quiet) {
            $btn.BackColor = 15790320
            $btn.ForeColor = $ink
        } else {
            $btn.BackColor = $teal
            $btn.ForeColor = $white
            try { $btn.HoverColor = 5390651 } catch {}
            try { $btn.HoverForeColor = $white } catch {}
        }
    }

    # Datasheet lists
    $lists = @{
        "SF_ExpenseList" = "T_Expenses"
        "SF_StaffList" = "T_Staff"
        "SF_Attendance" = "Q_AttendanceList"
        "SF_Leave" = "Q_LeaveList"
        "SF_ExpenseStats" = "Q_ExpenseByCategory"
    }
    foreach ($key in $lists.Keys) {
        $frm = Open-Design $acc $key
        $frm.RecordSource = $lists[$key]
        $frm.DefaultView = 2
        $frm.AllowAdditions = $false
        try { $frm.Section(0).BackColor = $white } catch {}
        $acc.DoCmd.Close(2, $key, 1)
    }

    # Accounting hub
    $frm = Open-Design $acc "F_Accounting"
    $frm.Caption = "Accounting"
    try { $frm.Section(0).BackColor = $bg; $frm.Section(0).Height = 4200 } catch {}
    Add-Label $acc "F_Accounting" "lblTitle" "Accounting" 300 200 5000 420 22
    Add-Label $acc "F_Accounting" "lblSub" "Expenses, staff attendance, leave, and statistics." 300 680 9000 300 11
    $acc.Forms.Item("F_Accounting").Controls.Item("lblSub").FontWeight = 400
    Add-Button $acc "F_Accounting" "cmdExpenses" "Expenses" "=UX_OpenForm(""F_Expenses"")" 300 1300 2800 700
    Add-Button $acc "F_Accounting" "cmdHR" "Staff and Leave" "=UX_OpenForm(""F_HR"")" 3300 1300 2800 700
    Add-Button $acc "F_Accounting" "cmdStats" "Statistics" "=UX_OpenForm(""F_Statistics"")" 6300 1300 2800 700
    Add-Button $acc "F_Accounting" "cmdUXBack" "Back" "=UX_GoBack()" 300 2400 1400 360
    foreach ($bn in @("cmdExpenses","cmdHR","cmdStats")) { Style-Button $acc.Forms.Item("F_Accounting").Controls.Item($bn) $false }
    Style-Button $acc.Forms.Item("F_Accounting").Controls.Item("cmdUXBack") $true
    $acc.DoCmd.Close(2, "F_Accounting", 1)

    # Expenses
    $frm = Open-Design $acc "F_Expenses"
    $frm.Caption = "Expenses"
    try { $frm.Section(0).BackColor = $bg; $frm.Section(0).Height = 7200 } catch {}
    Add-Label $acc "F_Expenses" "lblTitle" "Expenses" 300 160 4000 400 20
    Add-Label $acc "F_Expenses" "lblDate" "Date" 300 700 1200 280 10
    Add-Text $acc "F_Expenses" "txtExpenseDate" 1600 680 1800 300
    Add-Label $acc "F_Expenses" "lblCat" "Category" 3600 700 1400 280 10
    Add-Combo $acc "F_Expenses" "cboCategory" 5100 680 2800 300 "Rent;Utilities;Salaries;Supplies;Transport;Other" $true
    $acc.Forms.Item("F_Expenses").Controls.Item("cboCategory").ColumnCount = 1
    $acc.Forms.Item("F_Expenses").Controls.Item("cboCategory").ColumnWidths = "6cm"
    Add-Label $acc "F_Expenses" "lblDesc" "Description" 300 1120 1400 280 10
    Add-Text $acc "F_Expenses" "txtDescription" 1800 1100 6100 300
    Add-Label $acc "F_Expenses" "lblAmt" "Amount" 300 1540 1400 280 10
    Add-Text $acc "F_Expenses" "txtAmount" 1800 1520 1800 300
    Add-Label $acc "F_Expenses" "lblPaid" "Paid to" 3800 1540 1200 280 10
    Add-Text $acc "F_Expenses" "txtPaidTo" 5000 1520 2900 300
    Add-Label $acc "F_Expenses" "lblNotes" "Notes" 300 1960 1400 280 10
    Add-Text $acc "F_Expenses" "txtNotes" 1800 1940 6100 300
    Add-Button $acc "F_Expenses" "cmdSaveExpense" "Save Expense" "=UX_SaveExpense()" 1800 2380 2000 380
    Add-Button $acc "F_Expenses" "cmdUXBack" "Back" "=UX_GoBack()" 4000 2380 1400 380
    Add-Sub $acc "F_Expenses" "SF_ExpenseList" "SF_ExpenseList" 300 3000 11000 3800
    $acc.Forms.Item("F_Expenses").Controls.Item("txtExpenseDate").DefaultValue = "=Date()"
    $acc.Forms.Item("F_Expenses").Controls.Item("cboCategory").DefaultValue = '"Supplies"'
    Style-Button $acc.Forms.Item("F_Expenses").Controls.Item("cmdSaveExpense") $false
    Style-Button $acc.Forms.Item("F_Expenses").Controls.Item("cmdUXBack") $true
    $acc.DoCmd.Close(2, "F_Expenses", 1)

    # HR
    $frm = Open-Design $acc "F_HR"
    $frm.Caption = "Staff and Leave"
    try { $frm.Section(0).BackColor = $bg; $frm.Section(0).Height = 9800 } catch {}
    Add-Label $acc "F_HR" "lblTitle" "Staff, arrivals, and leave" 300 140 7000 400 20
    Add-Label $acc "F_HR" "lblPerson" "Person" 300 680 1200 280 10
    Add-Combo $acc "F_HR" "cboStaff" 1600 660 3600 300 "SELECT StaffID, StaffName FROM T_Staff ORDER BY StaffName;" $false
    Add-Label $acc "F_HR" "lblNew" "New person" 5400 680 1400 280 10
    Add-Text $acc "F_HR" "txtStaffName" 6900 660 2800 300
    Add-Text $acc "F_HR" "txtJobTitle" 6900 1020 1800 300
    Add-Text $acc "F_HR" "txtPhone" 8800 1020 1600 300
    Add-Button $acc "F_HR" "cmdAddStaff" "Add Person" "=UX_AddStaff()" 10500 660 1800 320
    Add-Label $acc "F_HR" "lblArr" "Arrival" 300 1500 2000 300 14
    Add-Label $acc "F_HR" "lblWork" "Date" 300 1900 800 260 10
    Add-Text $acc "F_HR" "txtWorkDate" 1200 1880 1600 300
    Add-Label $acc "F_HR" "lblIn" "Time in" 3000 1900 900 260 10
    Add-Text $acc "F_HR" "txtTimeIn" 4000 1880 1400 300
    Add-Label $acc "F_HR" "lblOut" "Time out" 5600 1900 1000 260 10
    Add-Text $acc "F_HR" "txtTimeOut" 6700 1880 1400 300
    Add-Text $acc "F_HR" "txtAttendanceNotes" 8300 1880 2200 300
    Add-Button $acc "F_HR" "cmdSign" "Save Arrival" "=UX_SaveAttendance()" 300 2300 2000 360
    Add-Sub $acc "F_HR" "SF_Attendance" "SF_Attendance" 300 2780 12000 1800
    Add-Label $acc "F_HR" "lblLeave" "Off days and leave" 300 4700 4000 300 14
    Add-Combo $acc "F_HR" "cboLeaveType" 300 5100 1800 300 "Annual;Sick;Off day;Unpaid" $true
    $acc.Forms.Item("F_HR").Controls.Item("cboLeaveType").ColumnCount = 1
    Add-Text $acc "F_HR" "txtLeaveStart" 2300 5100 1600 300
    Add-Text $acc "F_HR" "txtLeaveEnd" 4100 5100 1600 300
    Add-Text $acc "F_HR" "txtLeaveDays" 5900 5100 900 300
    Add-Combo $acc "F_HR" "cboLeaveStatus" 7000 5100 1800 300 "Requested;Approved;Taken" $true
    $acc.Forms.Item("F_HR").Controls.Item("cboLeaveStatus").ColumnCount = 1
    Add-Text $acc "F_HR" "txtLeaveNotes" 9000 5100 2400 300
    Add-Button $acc "F_HR" "cmdLeave" "Save Leave" "=UX_SaveLeave()" 300 5540 2000 360
    Add-Sub $acc "F_HR" "SF_Leave" "SF_Leave" 300 6040 12000 2200
    Add-Sub $acc "F_HR" "SF_StaffList" "SF_StaffList" 300 8400 12000 1200
    Add-Button $acc "F_HR" "cmdUXBack" "Back" "=UX_GoBack()" 2500 2300 1400 360
    $acc.Forms.Item("F_HR").Controls.Item("txtWorkDate").DefaultValue = "=Date()"
    $acc.Forms.Item("F_HR").Controls.Item("cboLeaveType").DefaultValue = '"Off day"'
    $acc.Forms.Item("F_HR").Controls.Item("cboLeaveStatus").DefaultValue = '"Approved"'
    Style-Button $acc.Forms.Item("F_HR").Controls.Item("cmdAddStaff") $false
    Style-Button $acc.Forms.Item("F_HR").Controls.Item("cmdSign") $false
    Style-Button $acc.Forms.Item("F_HR").Controls.Item("cmdLeave") $false
    Style-Button $acc.Forms.Item("F_HR").Controls.Item("cmdUXBack") $true
    $acc.DoCmd.Close(2, "F_HR", 1)

    # Statistics
    $frm = Open-Design $acc "F_Statistics"
    $frm.Caption = "Statistics"
    try { $frm.Section(0).BackColor = $bg; $frm.Section(0).Height = 7800 } catch {}
    Add-Label $acc "F_Statistics" "lblTitle" "Statistics" 300 160 4000 400 20
    function Add-Stat($acc, $labelName, $textName, $caption, $source, $left) {
        Add-Label $acc "F_Statistics" $labelName $caption $left 800 2000 260 10
        Add-Text $acc "F_Statistics" $textName $left 1100 2000 420
        $box = $acc.Forms.Item("F_Statistics").Controls.Item($textName)
        $box.ControlSource = $source
        $box.FontSize = 16
        $box.FontWeight = 700
        $box.BackColor = $white
        $box.ForeColor = $teal
        try { $box.Locked = $true; $box.BorderStyle = 1 } catch {}
    }
    Add-Stat $acc "lblSales" "txtSales" "Sales total" '=Format(Nz(DSum("GrandTotalAmount","T_Invoices"),0),"#,##0.00")' 300
    Add-Stat $acc "lblInv" "txtInv" "Invoices" '=Nz(DCount("*","T_Invoices"),0)' 2500
    Add-Stat $acc "lblExp" "txtExp" "Expenses" '=Format(Nz(DSum("Amount","T_Expenses"),0),"#,##0.00")' 4700
    Add-Stat $acc "lblStock" "txtStock" "Stock on hand" '=Nz(DSum("QuantityOnHand","T_StockBalances"),0)' 6900
    Add-Stat $acc "lblProj" "txtProj" "Projects" '=Nz(DCount("*","T_Projects"),0)' 9100
    Add-Stat $acc "lblLeaveN" "txtLeaveN" "Approved leave" "=Nz(DCount(`"*`",`"T_Leave`",`"Status='Approved'`"),0)" 11300
    Add-Label $acc "F_Statistics" "lblByCat" "Expenses by category" 300 1800 4000 300 14
    Add-Sub $acc "F_Statistics" "SF_ExpenseStats" "SF_ExpenseStats" 300 2200 7000 2800
    Add-Button $acc "F_Statistics" "cmdUXBack" "Back" "=UX_GoBack()" 300 5200 1400 360
    Style-Button $acc.Forms.Item("F_Statistics").Controls.Item("cmdUXBack") $true
    $acc.DoCmd.Close(2, "F_Statistics", 1)
    Log "Accounting forms built"

    # Dashboard button
    $frm = Open-Design $acc "Dashboard"
    if (-not (Test-Control $frm "cmdAccounting")) {
        $c = $acc.CreateControl("Dashboard", 104, 0, "", "", 60, 3000, 2760, 420)
        $c.Name = "cmdAccounting"
    }
    $b = $frm.Controls.Item("cmdAccounting")
    $b.Caption = "ACCOUNTING"
    $b.OnClick = "=UX_OpenForm(""F_Accounting"")"
    $b.FontWeight = 700
    Style-Button $b $false
    $acc.DoCmd.Close(2, "Dashboard", 1)
    Log "Dashboard accounting button"

    # Shrink Add Customer on the order screens
    foreach ($pair in @(@("F_SalesOrders","CustomerID"), @("F_Invoices","CustomerID"), @("F_CustomerPO","cboCustomerID"), @("F_Projects","cboCustomerID"))) {
        $formName = $pair[0]
        $frm = Open-Design $acc $formName
        if (Test-Control $frm "cmdAddCustomer") {
            $btn = $frm.Controls.Item("cmdAddCustomer")
            $combo = $frm.Controls.Item($pair[1])
            $btn.Width = 1480
            $btn.Height = 300
            $btn.Top = $combo.Top
            $btn.Left = $combo.Left + $combo.Width + 80
            $btn.FontSize = 9
            $btn.Caption = "Add Customer"
            Style-Button $btn $false
        }
        $acc.DoCmd.Close(2, $formName, 1)
        Log "Sized Add Customer on $formName"
    }

    # Theme main screens
    Log "Applying theme"
    $skipBtn = @{}
    foreach ($f in @($acc.CurrentProject.AllForms)) {
        $name = $f.Name
        try {
            $frm = Open-Design $acc $name
            try { $frm.Section(0).BackColor = $bg } catch {}
            try { if ($frm.Section(1).Height -gt 0) { $frm.Section(1).BackColor = $white } } catch {}
            foreach ($ctl in @($frm.Controls)) {
                try {
                    if ($ctl.ControlType -eq 104) {
                        $cap = [string]$ctl.Caption
                        $quiet = ($cap -eq "Back" -or $cap -eq "Close" -or $cap -eq "EXIT")
                        Style-Button $ctl $quiet
                    } elseif ($ctl.ControlType -eq 109 -or $ctl.ControlType -eq 111) {
                        $ctl.BackColor = $white
                        $ctl.ForeColor = $ink
                        try { $ctl.FontName = "Calibri" } catch {}
                    }
                } catch {}
            }
            $acc.DoCmd.Close(2, $name, 1)
        } catch {
            Log "Theme skip $name $($_.Exception.Message)"
            try { $acc.DoCmd.Close(2, $name, 2) } catch {}
        }
    }

    # Cleaner printable documents
    Log "Styling documents"
    foreach ($rep in @("R_Invoice","R_Quotation","R_SalesOrder","R_PurchaseOrder","R_SupplierPO","R_DeliveryNote")) {
        try {
            while ($acc.CurrentProject.AllReports.Item($rep).IsLoaded) { $acc.DoCmd.Close(3, $rep, 2) }
            $acc.DoCmd.OpenReport($rep, 1)
            $rpt = $acc.Reports.Item($rep)
            try { $rpt.Section(0).BackColor = $white } catch {}
            if (Test-Control $rpt "Label10") {
                $t = $rpt.Controls.Item("Label10")
                $t.FontName = "Calibri"
                $t.FontSize = 22
                $t.FontWeight = 700
                $t.ForeColor = $ink
                $t.Width = 4200
                $t.Height = 480
            }
            if (Test-Control $rpt "Text11" -and (Test-Control $rpt "Label10")) {
                $n = $rpt.Controls.Item("Text11")
                $t = $rpt.Controls.Item("Label10")
                $n.Left = $t.Left + 4300
                $n.Top = $t.Top
                $n.Width = 2600
                $n.Height = 480
                $n.FontName = "Calibri"
                $n.FontSize = 16
                $n.ForeColor = $teal
            }
            if (Test-Control $rpt "Text9") {
                $c = $rpt.Controls.Item("Text9")
                $c.FontName = "Calibri"
                $c.FontSize = 16
                $c.FontWeight = 700
                $c.ForeColor = $ink
            }
            foreach ($key in @("ReportItem_Label","Unit_Label","UnitPrice_Label","Quantity_Label","Taxable_Label","LineSubtotal_Label")) {
                if (Test-Control $rpt $key) {
                    $h = $rpt.Controls.Item($key)
                    $h.FontName = "Calibri"
                    $h.FontSize = 10
                    $h.FontWeight = 700
                    $h.ForeColor = $ink
                    $h.BackStyle = 1
                    $h.BackColor = $soft
                    $h.Height = 360
                }
            }
            $acc.DoCmd.Close(3, $rep, 1)
            Log "Styled $rep"
        } catch {
            Log "Report skip $rep $($_.Exception.Message)"
            try { $acc.DoCmd.Close(3, $rep, 2) } catch {}
        }
    }
    foreach ($rep in @("R_InvoiceLines","R_SalesDocLines","R_POLines","R_SupplierPOLinesRpt")) {
        try {
            while ($acc.CurrentProject.AllReports.Item($rep).IsLoaded) { $acc.DoCmd.Close(3, $rep, 2) }
            $acc.DoCmd.OpenReport($rep, 1)
            $rpt = $acc.Reports.Item($rep)
            try { $rpt.Section(0).BackColor = $white } catch {}
            foreach ($ctl in @($rpt.Controls)) {
                try {
                    $ctl.FontName = "Calibri"
                    $ctl.ForeColor = $ink
                } catch {}
            }
            $acc.DoCmd.Close(3, $rep, 1)
        } catch {
            Log "Line report skip $rep $($_.Exception.Message)"
            try { $acc.DoCmd.Close(3, $rep, 2) } catch {}
        }
    }

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
