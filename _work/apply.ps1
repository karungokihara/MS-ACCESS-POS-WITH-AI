$ErrorActionPreference = "Continue"

$root = "C:\Users\user\Downloads\LLC"
$dbPath = Join-Path $root "LLC_v3_2026-10-29.accdb"
$work = Join-Path $root "_work"
$preFixCopy = Join-Path $root "LLC_v3_before_fixes.accdb"

Get-Process MSACCESS -ErrorAction SilentlyContinue | Stop-Process -Force -ErrorAction SilentlyContinue
Start-Sleep -Seconds 1

if (-not (Test-Path $preFixCopy)) {
    Copy-Item -LiteralPath $dbPath -Destination $preFixCopy -Force
    Write-Host "safety copy created"
}

function Normalize([string]$s) {
    if ($null -eq $s) { return "" }
    return ($s -replace "`r`n", "`n")
}

function Get-Code($compName) {
    $cm = $acc.VBE.VBProjects.Item(1).VBComponents.Item($compName).CodeModule
    return $cm.Lines(1, $cm.CountOfLines)
}

function Set-Code($compName, [string]$text) {
    $cm = $acc.VBE.VBProjects.Item(1).VBComponents.Item($compName).CodeModule
    $n = $cm.CountOfLines
    if ($n -gt 0) { $cm.DeleteLines(1, $n) }
    $text = Normalize $text
    $text = $text -replace "`n", "`r`n"
    $cm.AddFromString($text)
}

function Patch-Code($compName, $pairs) {
    $text = Normalize (Get-Code $compName)
    $changed = $false
    foreach ($p in $pairs) {
        $old = Normalize $p.Old
        $new = Normalize $p.New
        if ($p.ContainsKey("Marker") -and $text.Contains($p.Marker)) {
            Write-Host ("  skip (already patched): " + $compName + " :: " + $p.Label)
            continue
        }
        if ($text.Contains($old)) {
            $text = $text.Replace($old, $new)
            $changed = $true
            Write-Host ("  patched: " + $compName + " :: " + $p.Label)
        } else {
            Write-Host ("  MISS   : " + $compName + " :: " + $p.Label)
        }
    }
    if ($changed) { Set-Code $compName $text }
}

$acc = New-Object -ComObject Access.Application
$acc.Visible = $false
$acc.AutomationSecurity = 1
$acc.OpenCurrentDatabase($dbPath, $false)
$db = $acc.CurrentDb()

# ---------------------------------------------------------------
# 1. SCHEMA: template part numbers
# ---------------------------------------------------------------
Write-Host "== schema"
$tbl = $db.TableDefs.Item("T_AssemblyTemplates")
$wanted = @(
    @{ Name = "ManufacturerPartNumber"; Type = 10; Size = 120 },
    @{ Name = "VendorPartNumber"; Type = 10; Size = 120 },
    @{ Name = "InternalPartNumber"; Type = 10; Size = 120 },
    @{ Name = "AutoGeneratePartNumber"; Type = 1; Size = 0 }
)
foreach ($spec in $wanted) {
    $exists = $false
    foreach ($f in $tbl.Fields) { if ($f.Name -eq $spec.Name) { $exists = $true } }
    if ($exists) {
        Write-Host ("  exists: " + $spec.Name)
    } else {
        $f = $tbl.CreateField($spec.Name, [int]$spec.Type)
        if ($spec.Size -gt 0) { $f.Size = [int]$spec.Size }
        if ($spec.Type -eq 1) { $f.Required = $false }
        $tbl.Fields.Append($f)
        Write-Host ("  added : " + $spec.Name)
    }
}
$tbl.Fields.Refresh()

# ---------------------------------------------------------------
# 2. QUERY: Q_UnitList
# ---------------------------------------------------------------
Write-Host "== query Q_UnitList"
$unitSql = @'
SELECT U.UnitID, U.UnitCode, U.UnitName, U.Description, IIf(Len(Trim(Nz(U.UnitCode,'')))=0 Or Len(Trim(Nz(U.UnitName,'')))=0,'Incomplete','Complete') AS RecordStatus, (SELECT Count(*) FROM T_Items AS I WHERE I.BaseUnitID = U.UnitID) AS ItemCount FROM T_Units AS U ORDER BY U.UnitName, U.UnitCode;
'@
$unitSql = $unitSql.Trim()
$qdef = $null
foreach ($q in $db.QueryDefs) { if ($q.Name -eq "Q_UnitList") { $qdef = $q } }
if ($null -eq $qdef) {
    $qdef = $db.CreateQueryDef("Q_UnitList", $unitSql)
    Write-Host "  created"
} else {
    $qdef.SQL = $unitSql
    Write-Host "  updated"
}

# ---------------------------------------------------------------
# 3. MODULES: import M_Fixes and M_Backup
# ---------------------------------------------------------------
Write-Host "== modules"
$proj = $acc.VBE.VBProjects.Item(1)
function Import-Module-Bas([string]$name, [string]$file) {
    $existing = $null
    foreach ($c in $proj.VBComponents) { if ($c.Name -eq $name) { $existing = $c } }
    if ($null -ne $existing) {
        $proj.VBComponents.Remove($existing)
        Write-Host ("  removed old " + $name)
    }
    $proj.VBComponents.Import($file) | Out-Null
    Write-Host ("  imported " + $name)
}
Import-Module-Bas "M_Fixes"  (Join-Path $work "M_Fixes.bas")
Import-Module-Bas "M_Backup" (Join-Path $work "M_Backup.bas")

# ---------------------------------------------------------------
# 4. FORM CODE: F_LocationSetup full replacement
# ---------------------------------------------------------------
Write-Host "== form code F_LocationSetup"
$locLines = Get-Content -LiteralPath (Join-Path $work "F_LocationSetup.vba")
# drop the leading Attribute lines (module keeps its own name)
$start = 0
while ($start -lt $locLines.Count -and $locLines[$start].StartsWith("Attribute ")) { $start++ }
$locText = ($locLines[$start..($locLines.Count - 1)] -join "`r`n")
Set-Code "Form_F_LocationSetup" $locText
Write-Host ("  lines now: " + (Get-Code "Form_F_LocationSetup").Split("`n").Count)

# ---------------------------------------------------------------
# 5. PATCHES
# ---------------------------------------------------------------
Write-Host "== patch F_Warehouses"
Patch-Code "Form_F_Warehouses" @(
    @{ Label = "keep incoming filter"; Marker = "keepFilter"; Old = @'
Private Sub Form_Load()
    On Error Resume Next
    Me.Filter = ""
    Me.FilterOn = False
    Me.Caption = "Warehouses"
'@; New = @'
Private Sub Form_Load()
    Dim keepFilter As String
    Dim keepFilterOn As Boolean
    On Error Resume Next
    keepFilter = Nz(Me.Filter, "")
    keepFilterOn = Me.FilterOn
    Me.Filter = ""
    Me.FilterOn = False
    Me.Caption = "Warehouses"
'@ }
    @{ Label = "restore incoming filter"; Marker = "keepFilterOn"; Old = @'
    UX_MarkRequired Me, "WarehouseName", "WarehouseCode"
End Sub
'@; New = @'
    UX_MarkRequired Me, "WarehouseName", "WarehouseCode"
    If Len(keepFilter) > 0 Then
        Me.Filter = keepFilter
        Me.FilterOn = keepFilterOn
    End If
End Sub
'@ }
)

Write-Host "== patch F_ItemSetup"
Patch-Code "Form_F_ItemSetup" @(
    @{ Label = "declare max ids"; Marker = "newMaximumID"; Old = @'
    Dim enteredText As String
    Dim savedLocationID As Variant

    CreateTypedLocation = False
'@; New = @'
    Dim enteredText As String
    Dim savedLocationID As Variant
    Dim oldMaximumID As Long
    Dim newMaximumID As Long

    CreateTypedLocation = False
'@ }
    @{ Label = "remember max id"; Marker = "oldMaximumID ="; Old = @'
    If Not CloseLocationSetupIfLoaded Then
        Exit Function
    End If

    DoCmd.OpenForm _
        formName:="F_LocationSetup", _
        View:=acNormal, _
        DataMode:=acFormAdd, _
        WindowMode:=acDialog, _
        OpenArgs:="QUICKADD|" & enteredText
'@; New = @'
    If Not CloseLocationSetupIfLoaded Then
        Exit Function
    End If

    oldMaximumID = Nz(DMax("LocationID", "T_Locations"), 0)

    DoCmd.OpenForm _
        formName:="F_LocationSetup", _
        View:=acNormal, _
        DataMode:=acFormAdd, _
        WindowMode:=acDialog, _
        OpenArgs:="QUICKADD|" & enteredText
'@ }
    @{ Label = "resolve saved id"; Marker = "newMaximumID > oldMaximumID"; Old = @'
    savedLocationID = DLookup( _
        "LocationID", _
        "T_Locations", _
        "Nz([LocationName],'')=" & SqlText(enteredText))

    If IsNull(savedLocationID) Then

        MsgBox _
            "The location was not found after Location Setup " & _
            "closed." & vbCrLf & vbCrLf & _
            "Save the location using the same name that you typed.", _
            vbInformation, _
            "Location Not Saved"

        Exit Function

    End If

    CreateTypedLocation = True
'@; New = @'
    Me.cboDefaultLocationID.Requery

    savedLocationID = DLookup( _
        "LocationID", _
        "T_Locations", _
        "StrComp(Trim(Nz([LocationName],''))," & _
            SqlText(enteredText) & ",1)=0")

    If IsNull(savedLocationID) Then

        newMaximumID = Nz(DMax("LocationID", "T_Locations"), 0)

        If newMaximumID > oldMaximumID Then
            savedLocationID = newMaximumID
        End If

    End If

    If IsNull(savedLocationID) Then

        MsgBox _
            "The location was not saved." & vbCrLf & vbCrLf & _
            "Enter the location name in Location Setup and press Save.", _
            vbInformation, _
            "Location Not Saved"

        Exit Function

    End If

    Me.cboDefaultLocationID.Value = CLng(savedLocationID)
    CreateTypedLocation = True
'@ }
)

Write-Host "== patch F_CustomerSetup"
Patch-Code "Form_F_CustomerSetup" @(
    @{ Label = "highlight missing"; Marker = "UX_HighlightMissingCustomer"; Old = @'
    Else
        Me.cboFindCustomer = Me.txtCustomerID
    End If

    SetSubformAvailability
'@; New = @'
    Else
        Me.cboFindCustomer = Me.txtCustomerID
    End If

    SetSubformAvailability
    UX_HighlightMissingCustomer Me
'@ }
)

Write-Host "== patch F_AssemblyTemplate"
Patch-Code "Form_F_AssemblyTemplate" @(
    @{ Label = "status list"; Marker = "UX_StatusList()"; Old = '        .RowSource = "Draft;Active;On Hold;Retired"'; New = '        .RowSource = UX_StatusList()' }
    @{ Label = "normalize part numbers"; Marker = "ApplyTemplatePartNumbers"; Old = @'
    Me.cboStatus = Trim$(Nz(Me.cboStatus, vbNullString))
    Me.txtNOtes = Trim$(Nz(Me.txtNOtes, vbNullString))
'@; New = @'
    Me.cboStatus = Trim$(Nz(Me.cboStatus, vbNullString))
    Me.txtNOtes = Trim$(Nz(Me.txtNOtes, vbNullString))
    ApplyTemplatePartNumbers
'@ }
    @{ Label = "lock internal part number"; Marker = "ApplyPartNumberLock"; Old = '    Me.txtTemplateCode.Locked = True'; New = @'
    Me.txtTemplateCode.Locked = True
    ApplyPartNumberLock
'@ }
    @{ Label = "part number helpers"; Marker = "Private Sub ApplyPartNumberLock"; Old = @'
'=============================================================
' TEMPLATE CODE GENERATION
'=============================================================
'@; New = @'
'=============================================================
' PART NUMBERS
'=============================================================

Private Sub ApplyPartNumberLock()

    On Error Resume Next

    If Not ControlExists("txtInternalPartNumber") Then Exit Sub

    Me.txtInternalPartNumber.Locked = _
        Nz(Me.chkAutoGeneratePartNumber, False)

    If Me.txtInternalPartNumber.Locked Then
        Me.txtInternalPartNumber.BackColor = RGB(236, 236, 236)
    Else
        Me.txtInternalPartNumber.BackColor = RGB(255, 255, 255)
    End If

    On Error GoTo 0

End Sub


Private Sub ApplyTemplatePartNumbers()

    Dim generatedNumber As String

    On Error Resume Next

    If ControlExists("txtManufacturerPartNumber") Then
        Me.txtManufacturerPartNumber = _
            Trim$(Nz(Me.txtManufacturerPartNumber, vbNullString))
    End If

    If ControlExists("txtVendorPartNumber") Then
        Me.txtVendorPartNumber = _
            Trim$(Nz(Me.txtVendorPartNumber, vbNullString))
    End If

    If Not ControlExists("txtInternalPartNumber") Then
        On Error GoTo 0
        Exit Sub
    End If

    Me.txtInternalPartNumber = _
        Trim$(Nz(Me.txtInternalPartNumber, vbNullString))

    If Nz(Me.chkAutoGeneratePartNumber, False) Then

        generatedNumber = UX_GeneratePartNumber( _
            Me.txtTemplateCode, Me.txtRevisionNo)

        If Len(generatedNumber) > 0 Then
            Me.txtInternalPartNumber = generatedNumber
        End If

    End If

    On Error GoTo 0

End Sub


Private Sub chkAutoGeneratePartNumber_Click()

    On Error Resume Next

    If Nz(Me.chkAutoGeneratePartNumber, False) Then
        Me.txtInternalPartNumber = _
            UX_GeneratePartNumber(Me.txtTemplateCode, Me.txtRevisionNo)
    End If

    ApplyPartNumberLock

    On Error GoTo 0

End Sub


'=============================================================
' TEMPLATE CODE GENERATION
'=============================================================
'@ }
)

Write-Host "== patch M_UX"
Patch-Code "M_UX" @(
    @{ Label = "list form guard"; Marker = "UX_IsListForm(frm)"; Old = @'
Public Sub UX_FitForm(ByVal frm As Form, ByVal availW As Long, ByVal availH As Long)

    On Error Resume Next

    Dim ctl As Control
'@; New = @'
Public Sub UX_FitForm(ByVal frm As Form, ByVal availW As Long, ByVal availH As Long)

    On Error Resume Next

    Dim ctl As Control

    'List holder forms keep their own layout so the datasheet keeps
    'its horizontal scrollbar.
    If UX_IsListForm(frm) Then
        UX_FitListForm frm
        Exit Sub
    End If
'@ }
)

Write-Host "== patch F_SystemSettings"
Patch-Code "Form_F_SystemSettings" @(
    @{ Label = "next action order"; Marker = 'A warehouse is optional.'; Old = @'
    'Storage must exist before item default locations can be used.
    If mWarehouseTotal = 0 Then
        SetNextAction _
            "Start by creating a warehouse.", _
            "Create Warehouse"
        Exit Sub
    End If

    If mWarehouseIncomplete > 0 Then
        SetNextAction _
            MissingDetailsMessage( _
                mWarehouseIncomplete, _
                "warehouse record", _
                "warehouse records"), _
            "Review Warehouses"
        Exit Sub
    End If

    If mLocationTotal = 0 Then
        SetNextAction _
            "Create at least one warehouse location.", _
            "Create Location"
        Exit Sub
    End If

    If mLocationIncomplete > 0 Then
        SetNextAction _
            MissingDetailsMessage( _
                mLocationIncomplete, _
                "location record", _
                "location records"), _
            "Review Locations"
        Exit Sub
    End If
'@; New = @'
    'Locations come first. A location can be created before any
    'warehouse exists.
    If mLocationTotal = 0 Then
        SetNextAction _
            "Create your first stock location. A warehouse is optional.", _
            "Create Location"
        Exit Sub
    End If

    If mLocationIncomplete > 0 Then
        SetNextAction _
            MissingDetailsMessage( _
                mLocationIncomplete, _
                "location record", _
                "location records"), _
            "Review Locations"
        Exit Sub
    End If

    If mWarehouseTotal = 0 Then
        SetNextAction _
            "Create at least one warehouse.", _
            "Create Warehouse"
        Exit Sub
    End If

    If mWarehouseIncomplete > 0 Then
        SetNextAction _
            MissingDetailsMessage( _
                mWarehouseIncomplete, _
                "warehouse record", _
                "warehouse records"), _
            "Review Warehouses"
        Exit Sub
    End If
'@ }
    @{ Label = "customer criteria"; Marker = 'UX_AreaCriteria("customer")'; Old = @'
    CustomerIncompleteCriteria = AnyMissing( _
        BlankTextCriteria("CustomerName"))
'@; New = '    CustomerIncompleteCriteria = UX_AreaCriteria("customer")' }
    @{ Label = "supplier criteria"; Marker = 'UX_AreaCriteria("supplier")'; Old = @'
    SupplierIncompleteCriteria = AnyMissing( _
        BlankTextCriteria("CompanyName"))
'@; New = '    SupplierIncompleteCriteria = UX_AreaCriteria("supplier")' }
    @{ Label = "category criteria"; Marker = 'UX_AreaCriteria("category")'; Old = @'
    CategoryIncompleteCriteria = AnyMissing( _
        BlankTextCriteria("CategoryName"), _
        MissingKeyCriteria("ItemTypeID"), _
        "[RequiresSerial] Is Null", _
        "[RequiresExpiry] Is Null")
'@; New = '    CategoryIncompleteCriteria = UX_AreaCriteria("category")' }
    @{ Label = "unit criteria"; Marker = 'UX_AreaCriteria("unit")'; Old = @'
    UnitIncompleteCriteria = AnyMissing( _
        BlankTextCriteria("UnitCode"), _
        BlankTextCriteria("UnitName"))
'@; New = '    UnitIncompleteCriteria = UX_AreaCriteria("unit")' }
    @{ Label = "item criteria"; Marker = 'UX_AreaCriteria("item")'; Old = @'
    ItemIncompleteCriteria = AnyMissing( _
        BlankTextCriteria("InternalSKU"), _
        BlankTextCriteria("ItemName"), _
        BlankTextCriteria("ItemClass"), _
        MissingKeyCriteria("CategoryID"), _
        MissingKeyCriteria("BaseUnitID"), _
        MissingKeyCriteria("DefaultLocationID"), _
        "[ReorderLevel] Is Null", _
        "[MinStockLevel] Is Null", _
        "[MaxStockLevel] Is Null", _
        "[RequiresSerial] Is Null", _
        "(Nz([RequiresSerial],0)<>0 AND " & _
            BlankTextCriteria("SerialNumber") & ")")
'@; New = '    ItemIncompleteCriteria = UX_AreaCriteria("item")' }
    @{ Label = "template criteria"; Marker = 'UX_AreaCriteria("template")'; Old = @'
    TemplateIncompleteCriteria = AnyMissing( _
        BlankTextCriteria("TemplateCode"), _
        BlankTextCriteria("TemplateName"), _
        MissingKeyCriteria("CustomerID"), _
        "([OutputQuantity] Is Null OR [OutputQuantity]<=0)", _
        BlankTextCriteria("RevisionNo"), _
        BlankTextCriteria("Status"), _
        "[EstimatedLaborHours] Is Null", _
        "[EstimatedMaterialCost] Is Null", _
        "[EstimatedTotalCost] Is Null")
'@; New = '    TemplateIncompleteCriteria = UX_AreaCriteria("template")' }
    @{ Label = "warehouse criteria"; Marker = 'UX_AreaCriteria("warehouse")'; Old = @'
    WarehouseIncompleteCriteria = AnyMissing( _
        BlankTextCriteria("WarehouseName"))
'@; New = '    WarehouseIncompleteCriteria = UX_AreaCriteria("warehouse")' }
    @{ Label = "location criteria"; Marker = 'UX_AreaCriteria("location")'; Old = @'
    LocationIncompleteCriteria = AnyMissing( _
        MissingKeyCriteria("WarehouseID"), _
        BlankTextCriteria("LocationCode"), _
        BlankTextCriteria("LocationName"))
'@; New = '    LocationIncompleteCriteria = UX_AreaCriteria("location")' }
    @{ Label = "form load wiring"; Marker = 'WireAreaDrilldowns'; Old = @'
Private Sub Form_Load()

    RefreshSetup

End Sub
'@; New = @'
Private Sub Form_Load()

    WireAreaDrilldowns
    RefreshSetup

End Sub
'@ }
    @{ Label = "refresh wiring"; Marker = 'UpdateNextAction
    WireAreaDrilldowns'; Old = @'
    UpdateCardColors
    UpdateCompletion
    UpdateIncompleteSummary
    UpdateNextAction
'@; New = @'
    UpdateCardColors
    UpdateCompletion
    UpdateIncompleteSummary
    UpdateNextAction
    WireAreaDrilldowns
'@ }
    @{ Label = "drilldown wiring sub"; Marker = 'Private Sub WireAreaDrilldowns'; Old = @'
'=============================================================
' FORM EVENTS
'=============================================================
'@; New = @'
'=============================================================
' AREA DRILL-DOWN
'=============================================================

Private Sub WireAreaDrilldowns()

    SetAreaDrilldown Me.lblCustomerStatus, "customer", mCustomerIncomplete
    SetAreaDrilldown Me.lblSupplierStatus, "supplier", mSupplierIncomplete
    SetAreaDrilldown Me.lblCategoryStatus, "category", mCategoryIncomplete
    SetAreaDrilldown Me.lblUnitStatus, "unit", mUnitIncomplete
    SetAreaDrilldown Me.lblItemStatus, "item", mItemIncomplete
    SetAreaDrilldown Me.lblTemplateStatus, "template", mTemplateIncomplete
    SetAreaDrilldown Me.lblWarehouseStatus, "warehouse", mWarehouseIncomplete
    SetAreaDrilldown Me.lblLocationStatus, "location", mLocationIncomplete

End Sub


Private Sub SetAreaDrilldown( _
    ByVal statusLabel As Control, _
    ByVal areaName As String, _
    ByVal incompleteCount As Long _
)

    Dim openArguments As String

    On Error Resume Next

    If incompleteCount > 0 Then
        openArguments = "=UX_OpenAreaList(""" & areaName & """,True)"
    Else
        openArguments = "=UX_OpenAreaList(""" & areaName & """,False)"
    End If

    statusLabel.OnClick = openArguments
    statusLabel.HyperlinkAddress = vbNullString

    On Error GoTo 0

End Sub


'=============================================================
' FORM EVENTS
'=============================================================
'@ }
    @{ Label = "customer button"; Marker = 'UX_OpenAreaList "customer", (mCustomerIncomplete > 0)'; Old = @'
Private Sub cmdCustomerSetup_Click()

    DoCmd.OpenForm "F_CustomerList"

End Sub
'@; New = @'
Private Sub cmdCustomerSetup_Click()

    UX_OpenAreaList "customer", (mCustomerIncomplete > 0)

End Sub
'@ }
    @{ Label = "supplier button"; Marker = 'UX_OpenAreaList "supplier"'; Old = @'
Private Sub cmdSupplierSetup_Click()

    DoCmd.OpenForm "F_SupplierList"

End Sub
'@; New = @'
Private Sub cmdSupplierSetup_Click()

    UX_OpenAreaList "supplier", (mSupplierIncomplete > 0)

End Sub
'@ }
    @{ Label = "category button"; Marker = 'UX_OpenAreaList "category"'; Old = @'
Private Sub cmdCategorySetup_Click()

    DoCmd.OpenForm "F_CategoriesList"

End Sub
'@; New = @'
Private Sub cmdCategorySetup_Click()

    UX_OpenAreaList "category", (mCategoryIncomplete > 0)

End Sub
'@ }
    @{ Label = "unit button"; Marker = 'UX_OpenAreaList "unit"'; Old = @'
Private Sub cmdUnitSetup_Click()

    DoCmd.OpenForm "F_UnitSetup"

End Sub
'@; New = @'
Private Sub cmdUnitSetup_Click()

    UX_OpenAreaList "unit", (mUnitIncomplete > 0)

End Sub
'@ }
    @{ Label = "item button"; Marker = 'UX_OpenAreaList "item"'; Old = @'
Private Sub cmdItemSetup_Click()

    DoCmd.OpenForm "F_ItemSetupList"

End Sub
'@; New = @'
Private Sub cmdItemSetup_Click()

    UX_OpenAreaList "item", (mItemIncomplete > 0)

End Sub
'@ }
    @{ Label = "template button"; Marker = 'UX_OpenAreaList "template"'; Old = @'
Private Sub cmdTemplateSetup_Click()

    DoCmd.OpenForm "F_TemplateList"

End Sub
'@; New = @'
Private Sub cmdTemplateSetup_Click()

    UX_OpenAreaList "template", (mTemplateIncomplete > 0)

End Sub
'@ }
    @{ Label = "warehouse button"; Marker = 'UX_OpenAreaList "warehouse"'; Old = @'
Private Sub cmdWarehouseSetup_Click()

    DoCmd.OpenForm "F_WarehouseList"

End Sub
'@; New = @'
Private Sub cmdWarehouseSetup_Click()

    UX_OpenAreaList "warehouse", (mWarehouseIncomplete > 0)

End Sub
'@ }
    @{ Label = "location button"; Marker = 'UX_OpenAreaList "location"'; Old = @'
Private Sub cmdLocationSetup_Click()

    DoCmd.OpenForm "F_LocationSetup"

End Sub
'@; New = @'
Private Sub cmdLocationSetup_Click()

    UX_OpenAreaList "location", (mLocationIncomplete > 0)

End Sub
'@ }
    @{ Label = "next action routing"; Marker = 'actionCaption'; Old = @'
    Select Case Nz(Me.cmdNextAction.Caption, vbNullString)
        Case "Create Warehouse", "Review Warehouses", "Manage Warehouses"
            DoCmd.OpenForm "F_WarehouseList"
        Case "Create Location", "Review Locations"
            DoCmd.OpenForm "F_LocationSetup"
        Case "Create Categories", "Review Categories"
            DoCmd.OpenForm "F_CategoriesList"
        Case "Create Units", "Review Units"
            DoCmd.OpenForm "F_UnitSetup"
        Case "Create Customer", "Review Customers", "Manage Customers"
            DoCmd.OpenForm "F_CustomerList"
        Case "Create Supplier", "Review Suppliers", "Manage Suppliers"
            DoCmd.OpenForm "F_SupplierList"
        Case "Add Items", "Review Items"
            DoCmd.OpenForm "F_ItemSetupList"
        Case "Create Template", "Review Templates"
            DoCmd.OpenForm "F_TemplateList"
        Case Else
            RefreshSetup
    End Select
'@; New = @'
    Dim actionCaption As String

    actionCaption = Nz(Me.cmdNextAction.Caption, vbNullString)

    Select Case actionCaption
        Case "Create Warehouse", "Manage Warehouses"
            UX_OpenAreaList "warehouse", False
        Case "Review Warehouses"
            UX_OpenAreaList "warehouse", True
        Case "Create Location", "Manage Locations"
            UX_OpenAreaList "location", False
        Case "Review Locations"
            UX_OpenAreaList "location", True
        Case "Create Categories", "Add Category"
            UX_OpenAreaList "category", False
        Case "Review Categories"
            UX_OpenAreaList "category", True
        Case "Create Units", "Add Units"
            UX_OpenAreaList "unit", False
        Case "Review Units"
            UX_OpenAreaList "unit", True
        Case "Create Customer", "Manage Customers"
            UX_OpenAreaList "customer", False
        Case "Review Customers"
            UX_OpenAreaList "customer", True
        Case "Create Supplier", "Manage Suppliers"
            UX_OpenAreaList "supplier", False
        Case "Review Suppliers"
            UX_OpenAreaList "supplier", True
        Case "Add Items", "Manage Items"
            UX_OpenAreaList "item", False
        Case "Review Items"
            UX_OpenAreaList "item", True
        Case "Create Template", "Manage Templates"
            UX_OpenAreaList "template", False
        Case "Review Templates"
            UX_OpenAreaList "template", True
        Case Else
            RefreshSetup
    End Select
'@ }
)

# backup / restore handlers
$ssCode = Get-Code "Form_F_SystemSettings"
if (-not $ssCode.Contains("cmdBackupData_Click")) {
    $cm = $acc.VBE.VBProjects.Item(1).VBComponents.Item("Form_F_SystemSettings").CodeModule
    $extra = @'


'=============================================================
' BACKUP AND RESTORE
'=============================================================

Private Sub cmdBackupData_Click()

    UX_BackupAllData

End Sub


Private Sub cmdRestoreData_Click()

    UX_RestoreAllData

End Sub
'@
    $cm.AddFromString(($extra -replace "`n", "`r`n"))
    Write-Host "  added backup/restore handlers"
}

# ---------------------------------------------------------------
# 6. DESIGN: System Settings buttons, template part-number controls
# ---------------------------------------------------------------
Write-Host "== design F_SystemSettings"
$acc.DoCmd.OpenForm("F_SystemSettings", 1)
Start-Sleep -Milliseconds 400
$f = $acc.Forms.Item("F_SystemSettings")
function Ensure-Button($form, $name, $left, $top, $width, $height, $caption) {
    $existing = $null
    foreach ($c in $form.Controls) { if ($c.Name -eq $name) { $existing = $c } }
    if ($null -eq $existing) {
        $existing = $acc.CreateControl($form.Name, 104, 0, "", $name, $left, $top, $width, $height)
        Write-Host ("  created " + $name)
    } else {
        Write-Host ("  exists  " + $name)
    }
    $existing.Caption = $caption
    $existing.OnClick = "[Event Procedure]"
    $existing.FontName = "Segoe UI Semibold"
    $existing.FontSize = 11
    $existing.ForeColor = 16777215
    $existing.BackColor = 0
    $existing.FontWeight = 700
}
Ensure-Button $f "cmdBackupData"  1200 7500 3000 540 "Backup All Data"
Ensure-Button $f "cmdRestoreData" 4400 7500 3000 540 "Restore All Data"
$acc.DoCmd.Close(2, "F_SystemSettings", 1)

Write-Host "== design F_AssemblyTemplate"
$acc.DoCmd.OpenForm("F_AssemblyTemplate", 1)
Start-Sleep -Milliseconds 400
$f = $acc.Forms.Item("F_AssemblyTemplate")
function Ensure-Text($form, $name, $controlSource, $left, $top, $width, $height) {
    $existing = $null
    foreach ($c in $form.Controls) { if ($c.Name -eq $name) { $existing = $c } }
    if ($null -eq $existing) {
        $existing = $acc.CreateControl($form.Name, 109, 0, "", $name, $left, $top, $width, $height)
        Write-Host ("  created " + $name)
    } else {
        Write-Host ("  exists  " + $name)
    }
    $existing.ControlSource = $controlSource
    $existing.FontName = "Segoe UI"
    $existing.FontSize = 11
}
function Ensure-Label($form, $name, $caption, $left, $top, $width, $height) {
    $existing = $null
    foreach ($c in $form.Controls) { if ($c.Name -eq $name) { $existing = $c } }
    if ($null -eq $existing) {
        $existing = $acc.CreateControl($form.Name, 100, 0, "", $name, $left, $top, $width, $height)
        Write-Host ("  created " + $name)
    } else {
        Write-Host ("  exists  " + $name)
    }
    $existing.Caption = $caption
    $existing.FontName = "Segoe UI"
    $existing.FontSize = 10
}
$f.Section(0).Height = 14000
Ensure-Label $f "lblPartNumbersTitle" "Part Numbers" 480 12300 3000 300
Ensure-Label $f "lblManufacturerPartNumber" "Manufacturer Part No." 480 12660 2900 300
Ensure-Text  $f "txtManufacturerPartNumber" "ManufacturerPartNumber" 3500 12660 3200 300
Ensure-Label $f "lblVendorPartNumber" "Vendor Part No." 7000 12660 2300 300
Ensure-Text  $f "txtVendorPartNumber" "VendorPartNumber" 9400 12660 3200 300
Ensure-Label $f "lblInternalPartNumber" "Internal Part No." 480 13040 2900 300
Ensure-Text  $f "txtInternalPartNumber" "InternalPartNumber" 3500 13040 3200 300

$existing = $null
foreach ($c in $f.Controls) { if ($c.Name -eq "chkAutoGeneratePartNumber") { $existing = $c } }
if ($null -eq $existing) {
    $existing = $acc.CreateControl($f.Name, 106, 0, "", "chkAutoGeneratePartNumber", 7000, 13040, 4200, 300)
    Write-Host "  created chkAutoGeneratePartNumber"
}
$existing.ControlSource = "AutoGeneratePartNumber"
$existing.Caption = "Auto generate internal part number"
$existing.FontName = "Segoe UI"
$existing.FontSize = 10

$acc.DoCmd.Close(2, "F_AssemblyTemplate", 1)

# ---------------------------------------------------------------
# 7. DESIGN: datasheet subforms keep both scrollbars
# ---------------------------------------------------------------
Write-Host "== scrollbars"
$sfNames = @(
    "SF_CustomerList", "SF_SupplierList", "SF_ItemSetupList", "SF_TemplateList",
    "SF_WarehouseList", "SF_ProjectList", "SF_InventoryList", "SF_WarehouseLocations"
)
foreach ($n in $sfNames) {
    try {
        $acc.DoCmd.OpenForm($n, 1)
        Start-Sleep -Milliseconds 250
        $ff = $acc.Forms.Item($n)
        if ($ff.DefaultView -eq 2) { $ff.ScrollBars = 3 }
        $acc.DoCmd.Close(2, $n, 1)
        Write-Host ("  " + $n + " ok")
    } catch {
        Write-Host ("  " + $n + " FAILED " + $_.Exception.Message)
        try { $acc.DoCmd.Close(2, $n, 2) } catch {}
    }
}

# ---------------------------------------------------------------
# 8. COMPILE
# ---------------------------------------------------------------
Write-Host "== compile"
$acc.DoCmd.RunCommand(126) | Out-Null
Write-Host ("IsCompiled = " + $acc.IsCompiled)

$acc.CloseCurrentDatabase()
$acc.Quit()
Write-Host "DONE"
