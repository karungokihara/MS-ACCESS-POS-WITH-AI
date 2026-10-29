$ErrorActionPreference = 'Stop'

Add-Type @'
using System;
using System.Runtime.InteropServices;
public static class Version2AccessKeys
{
    [DllImport("user32.dll")]
    public static extern void keybd_event(byte key, byte scan, uint flags, UIntPtr extra);
}
'@

$databasePath = (Resolve-Path (Join-Path $PSScriptRoot '..\LLC_Version2.accdb')).Path
$access = $null
$database = $null

function Test-AccessFormExists {
    param([string]$Name)

    foreach ($form in $access.CurrentProject.AllForms) {
        if ($form.Name -eq $Name) {
            return $true
        }
    }

    return $false
}

function Copy-AccessFormIfMissing {
    param(
        [string]$Source,
        [string]$Destination
    )

    if (-not (Test-AccessFormExists $Destination)) {
        $access.DoCmd.CopyObject($null, $Destination, 2, $Source)
    }
}

function Open-FormDesign {
    param([string]$Name)

    $access.DoCmd.OpenForm($Name, 1)
    return $access.Forms.Item($Name)
}

function Get-AccessControl {
    param(
        $Form,
        [string[]]$Names
    )

    foreach ($name in $Names) {
        foreach ($control in $Form.Controls) {
            if ($control.Name -eq $name) {
                return $control
            }
        }
    }

    throw "Could not find any of these controls on $($Form.Name): $($Names -join ', ')"
}

function Set-FormModule {
    param(
        [string]$Name,
        [string]$Code
    )

    $module = $access.Forms.Item($Name).Module
    if ($module.CountOfLines -gt 0) {
        $module.DeleteLines(1, $module.CountOfLines)
    }
    $module.InsertLines(1, $Code.Trim())
}

function Set-Procedure {
    param(
        [string]$Module,
        [string]$Name,
        [string]$Code
    )

    $procedurePattern = '(?ms)^(Private|Public) Sub ' + [regex]::Escape($Name) +
        '\([^\r\n]*\)\r?\n.*?^End Sub[ \t]*\r?\n?'
    $expression = [regex]::new($procedurePattern)
    $existing = $expression.Match($Module)
    if ($existing.Success) {
        return $expression.Replace($Module, $Code.Trim() + "`r`n", 1)
    }

    return $Module.TrimEnd() + "`r`n`r`n" + $Code.Trim() + "`r`n"
}

function Write-AccessModule {
    param(
        [string]$Name,
        [string]$Code
    )

    Set-FormModule -Name $Name -Code $Code
}

try {
    $access = New-Object -ComObject Access.Application
    $access.Visible = $true
    Write-Output "Opening Version 2 database: $databasePath"
    [Version2AccessKeys]::keybd_event(0x10, 0, 0, [UIntPtr]::Zero)
    try {
        $access.OpenCurrentDatabase($databasePath)
    }
    finally {
        [Version2AccessKeys]::keybd_event(0x10, 0, 2, [UIntPtr]::Zero)
    }
    Write-Output 'Access database opened.'

    $database = $access.CurrentDb()
    $database.QueryDefs.Item('Q_V2_LocationList').SQL =
        'SELECT L.LocationID, L.LocationCode, L.LocationName, L.WarehouseID, ' +
        'W.WarehouseName, L.BinNumber, L.ShelfNumber ' +
        'FROM T_Locations AS L LEFT JOIN T_Warehouses AS W ' +
        'ON L.WarehouseID = W.WarehouseID ORDER BY L.LocationCode;'
    $database.QueryDefs.Item('Q_V2_UnitList').SQL =
        'SELECT UnitID, UnitCode, UnitName, Description ' +
        'FROM T_Units ORDER BY UnitCode;'

    Copy-AccessFormIfMissing -Source 'SF_WarehouseList' -Destination 'SF_LocationList'
    Copy-AccessFormIfMissing -Source 'F_WarehouseList' -Destination 'F_LocationList'
    Copy-AccessFormIfMissing -Source 'SF_WarehouseList' -Destination 'SF_UnitList'
    Copy-AccessFormIfMissing -Source 'F_WarehouseList' -Destination 'F_UnitList'

    Write-Output 'Building Location list form.'
    $locationSubform = Open-FormDesign 'SF_LocationList'
    $locationSubform.RecordSource = 'Q_V2_LocationList'
    $locationSubform.DefaultView = 2
    $locationSubform.ViewsAllowed = 2
    $locationSubform.AllowDatasheetView = $true
    $locationSubform.AllowEdits = $false
    $locationSubform.AllowAdditions = $false
    $locationSubform.AllowDeletions = $false
    $locationSubform.ScrollBars = 3
    $locationSubform.FitToScreen = $false

    $locationColumns = @(
        @{ Old = 'WarehouseID'; New = 'LocationID'; Caption = 'LocationID'; Width = 900; Hidden = $true },
        @{ Old = 'WarehouseCode'; New = 'LocationCode'; Caption = 'Location Code'; Width = 1900; Hidden = $false },
        @{ Old = 'WarehouseName'; New = 'LocationName'; Caption = 'Location Name'; Width = 2800; Hidden = $false },
        @{ Old = 'Address'; New = 'WarehouseName'; Caption = 'Warehouse'; Width = 2500; Hidden = $false },
        @{ Old = 'ContactPerson'; New = 'BinNumber'; Caption = 'Bin'; Width = 1500; Hidden = $false },
        @{ Old = 'LocationCount'; New = 'ShelfNumber'; Caption = 'Shelf'; Width = 1500; Hidden = $false }
    )
    foreach ($column in $locationColumns) {
        $control = Get-AccessControl -Form $locationSubform -Names @($column.Old, $column.New)
        $label = Get-AccessControl -Form $locationSubform -Names @(
            $column.Old + '_Label',
            $column.New + '_Label'
        )
        $control.Name = $column.New
        $control.ControlSource = $column.New
        $control.Width = $column.Width
        $control.Visible = -not $column.Hidden
        $label.Caption = $column.Caption
        $label.Visible = -not $column.Hidden
    }
    foreach ($columnName in @('LocationID', 'LocationCode', 'LocationName', 'WarehouseName', 'BinNumber', 'ShelfNumber')) {
        $locationSubform.Controls.Item($columnName).OnDblClick = '=OpenSelectedLocation()'
    }
    Write-AccessModule -Name 'SF_LocationList' -Code @'
Option Compare Database
Option Explicit

Public Function OpenSelectedLocation() As Boolean
    On Error GoTo Err_Handler

    Dim selectedID As Long
    Dim detailForm As String

    If IsNull(Me!LocationID) Then Exit Function
    selectedID = CLng(Me!LocationID)
    detailForm = "F_LocationSetup"

    If CurrentProject.AllForms(detailForm).IsLoaded Then
        If Forms(detailForm).Dirty Then Forms(detailForm).Dirty = False
        DoCmd.Close acForm, detailForm, acSaveNo
    End If

    DoCmd.OpenForm detailForm, acNormal, , "[LocationID]=" & CStr(selectedID)
    If Forms(detailForm)!LocationID <> selectedID Then
        Err.Raise vbObjectError + 2101, "OpenSelectedLocation", _
            "The location form did not open the selected location."
    End If

    OpenSelectedLocation = True
    Exit Function

Err_Handler:
    MsgBox "The selected location could not be opened." & vbCrLf & _
        "Error Number: " & Err.Number & vbCrLf & _
        "Details: " & Err.Description, vbExclamation, "Locations"
    OpenSelectedLocation = False
End Function
'@
    $access.DoCmd.Close(2, 'SF_LocationList', 1)

    Write-Output 'Building Unit list form.'
    $unitSubform = Open-FormDesign 'SF_UnitList'
    $unitSubform.RecordSource = 'Q_V2_UnitList'
    $unitSubform.DefaultView = 2
    $unitSubform.ViewsAllowed = 2
    $unitSubform.AllowDatasheetView = $true
    $unitSubform.AllowEdits = $true
    $unitSubform.AllowAdditions = $true
    $unitSubform.AllowDeletions = $false
    $unitSubform.ScrollBars = 3
    $unitSubform.FitToScreen = $false

    $unitColumns = @(
        @{ Old = 'WarehouseID'; New = 'UnitID'; Caption = 'UnitID'; Width = 900; Hidden = $true },
        @{ Old = 'WarehouseCode'; New = 'UnitCode'; Caption = 'Unit Code'; Width = 2200; Hidden = $false },
        @{ Old = 'WarehouseName'; New = 'UnitName'; Caption = 'Unit Name'; Width = 3000; Hidden = $false },
        @{ Old = 'Address'; New = 'Description'; Caption = 'Description'; Width = 4200; Hidden = $false },
        @{ Old = 'ContactPerson'; New = 'UnitUnused1'; Caption = ''; Width = 1000; Hidden = $true },
        @{ Old = 'LocationCount'; New = 'UnitUnused2'; Caption = ''; Width = 1000; Hidden = $true }
    )
    foreach ($column in $unitColumns) {
        $control = Get-AccessControl -Form $unitSubform -Names @($column.Old, $column.New)
        $label = Get-AccessControl -Form $unitSubform -Names @(
            $column.Old + '_Label',
            $column.New + '_Label'
        )
        $control.Name = $column.New
        if ($column.New -in @('UnitID', 'UnitCode', 'UnitName', 'Description')) {
            $control.ControlSource = $column.New
        } else {
            $control.ControlSource = ''
        }
        $control.Width = $column.Width
        $control.Visible = -not $column.Hidden
        $label.Caption = $column.Caption
        $label.Visible = -not $column.Hidden
    }
    foreach ($columnName in @('UnitID', 'UnitCode', 'UnitName', 'Description')) {
        $unitSubform.Controls.Item($columnName).OnDblClick = '=OpenSelectedUnit()'
    }
    Write-AccessModule -Name 'SF_UnitList' -Code @'
Option Compare Database
Option Explicit

Public Function OpenSelectedUnit() As Boolean
    On Error GoTo Err_Handler

    Dim selectedID As Long
    Dim detailForm As String

    If IsNull(Me!UnitID) Then Exit Function
    selectedID = CLng(Me!UnitID)
    detailForm = "F_UnitSetup"

    If CurrentProject.AllForms(detailForm).IsLoaded Then
        If Forms(detailForm).Dirty Then Forms(detailForm).Dirty = False
        DoCmd.Close acForm, detailForm, acSaveNo
    End If

    DoCmd.OpenForm detailForm, acNormal, , "[UnitID]=" & CStr(selectedID)
    If Forms(detailForm)!UnitID <> selectedID Then
        Err.Raise vbObjectError + 2102, "OpenSelectedUnit", _
            "The unit form did not open the selected unit."
    End If

    OpenSelectedUnit = True
    Exit Function

Err_Handler:
    MsgBox "The selected unit could not be opened." & vbCrLf & _
        "Error Number: " & Err.Number & vbCrLf & _
        "Details: " & Err.Description, vbExclamation, "Units"
    OpenSelectedUnit = False
End Function
'@
    $access.DoCmd.Close(2, 'SF_UnitList', 1)

    foreach ($list in @(
        @{ Form = 'F_LocationList'; Subform = 'subLocationList'; Source = 'SF_LocationList'; Caption = 'Location List'; Button = 'ADD LOCATION'; Label = 'Locations' },
        @{ Form = 'F_UnitList'; Subform = 'subUnitList'; Source = 'SF_UnitList'; Caption = 'Unit List'; Button = 'ADD UNIT'; Label = 'Units' }
    )) {
        Write-Output "Configuring $($list.Form)."
        $hostForm = Open-FormDesign $list.Form
        $hostForm.Caption = $list.Caption
        $hostForm.ScrollBars = 3
        $hostForm.FitToScreen = $false
        $hostForm.Controls.Item('btnNew').Caption = $list.Button
        $subformControl = Get-AccessControl -Form $hostForm -Names @('subWarehouseList', $list.Subform)
        $subformControl.Name = $list.Subform
        $hostForm.Controls.Item($list.Subform).SourceObject = 'Form.' + $list.Source
        $hostForm.Controls.Item($list.Subform).ScrollBars = 3
        foreach ($control in $hostForm.Controls) {
            if ($control.ControlType -eq 100 -and $control.Caption -eq 'subWarehouseList') {
                $control.Caption = $list.Label
            }
        }

        if ($list.Form -eq 'F_LocationList') {
            Write-AccessModule -Name $list.Form -Code @'
Option Compare Database
Option Explicit

Private Sub Form_Load()
    Me.subLocationList.Requery
End Sub

Private Sub Form_Activate()
    Me.subLocationList.Requery
End Sub

Private Sub txtSearch_AfterUpdate()
    ApplyListFilter
End Sub

Private Sub txtSearch_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyReturn Then
        KeyCode = 0
        ApplyListFilter
    End If
End Sub

Private Sub btnClearFilter_Click()
    Me.txtSearch = Null
    ApplyListFilter
End Sub

Private Sub btnNew_Click()
    DoCmd.OpenForm "F_LocationSetup", acNormal, , , acFormAdd
End Sub

Private Sub EXIT_Click()
    If Me.Dirty Then Me.Dirty = False
    DoCmd.Close acForm, Me.Name
End Sub

Private Sub ApplyListFilter()
    Dim searchText As String
    Dim criteria As String

    searchText = Replace(Trim$(Nz(Me.txtSearch.Value, vbNullString)), "'", "''")
    With Me.subLocationList.Form
        If Len(searchText) = 0 Then
            .FilterOn = False
        Else
            criteria = "[LocationCode] Like '*" & searchText & "*' OR " & _
                "[LocationName] Like '*" & searchText & "*' OR " & _
                "Nz([WarehouseName],'') Like '*" & searchText & "*'"
            .Filter = criteria
            .FilterOn = True
        End If
    End With
End Sub
'@
        } else {
            Write-AccessModule -Name $list.Form -Code @'
Option Compare Database
Option Explicit

Private Sub Form_Load()
    Me.subUnitList.Requery
End Sub

Private Sub Form_Activate()
    Me.subUnitList.Requery
End Sub

Private Sub txtSearch_AfterUpdate()
    ApplyListFilter
End Sub

Private Sub txtSearch_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyReturn Then
        KeyCode = 0
        ApplyListFilter
    End If
End Sub

Private Sub btnClearFilter_Click()
    Me.txtSearch = Null
    ApplyListFilter
End Sub

Private Sub btnNew_Click()
    DoCmd.OpenForm "F_UnitSetup", acNormal, , , acFormAdd
End Sub

Private Sub EXIT_Click()
    If Me.Dirty Then Me.Dirty = False
    DoCmd.Close acForm, Me.Name
End Sub

Private Sub ApplyListFilter()
    Dim searchText As String
    Dim criteria As String

    searchText = Replace(Trim$(Nz(Me.txtSearch.Value, vbNullString)), "'", "''")
    With Me.subUnitList.Form
        If Len(searchText) = 0 Then
            .FilterOn = False
        Else
            criteria = "[UnitCode] Like '*" & searchText & "*' OR " & _
                "[UnitName] Like '*" & searchText & "*' OR " & _
                "Nz([Description],'') Like '*" & searchText & "*'"
            .Filter = criteria
            .FilterOn = True
        End If
    End With
End Sub
'@
        }

        $access.DoCmd.Close(2, $list.Form, 1)
    }

    Write-Output 'Routing System Settings to the new lists.'
    $settings = Open-FormDesign 'F_SystemSettings'
    $settings.Controls.Item('lblUnitStatus').OnClick = '[Event Procedure]'
    $settings.Controls.Item('lblLocationStatus').OnClick = '[Event Procedure]'
    $settings.Controls.Item('lblUnitStatus').IsHyperlink = $true
    $settings.Controls.Item('lblLocationStatus').IsHyperlink = $true

    $module = $settings.Module
    $settingsCode = $module.Lines(1, $module.CountOfLines)
    $settingsCode = Set-Procedure -Module $settingsCode -Name 'cmdUnitSetup_Click' -Code @'
Private Sub cmdUnitSetup_Click()
    DoCmd.OpenForm "F_UnitList"
End Sub
'@
    $settingsCode = Set-Procedure -Module $settingsCode -Name 'cmdLocationSetup_Click' -Code @'
Private Sub cmdLocationSetup_Click()
    DoCmd.OpenForm "F_LocationList"
End Sub
'@
    $settingsCode = Set-Procedure -Module $settingsCode -Name 'cmdNextAction_Click' -Code @'
Private Sub cmdNextAction_Click()
    On Error GoTo Err_Handler

    Select Case Nz(Me.cmdNextAction.Caption, vbNullString)
        Case "Create Warehouse", "Review Warehouses", "Manage Warehouses"
            DoCmd.OpenForm "F_WarehouseList"
        Case "Create Location", "Review Locations"
            DoCmd.OpenForm "F_LocationList"
        Case "Create Categories", "Review Categories"
            DoCmd.OpenForm "F_CategoriesList"
        Case "Create Units", "Review Units"
            DoCmd.OpenForm "F_UnitList"
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

    Exit Sub

Err_Handler:
    MsgBox "The selected setup action could not be opened." & vbCrLf & _
        "Action: " & Nz(Me.cmdNextAction.Caption, vbNullString) & vbCrLf & _
        "Error Number: " & Err.Number & vbCrLf & _
        "Details: " & Err.Description, vbExclamation, "System Setup"
End Sub
'@
    $settingsCode = Set-Procedure -Module $settingsCode -Name 'lblUnitStatus_Click' -Code @'
Private Sub lblUnitStatus_Click()
    DoCmd.OpenForm "F_UnitList"
End Sub
'@
    $settingsCode = Set-Procedure -Module $settingsCode -Name 'lblLocationStatus_Click' -Code @'
Private Sub lblLocationStatus_Click()
    DoCmd.OpenForm "F_LocationList"
End Sub
'@
    Set-FormModule -Name 'F_SystemSettings' -Code $settingsCode
    $access.DoCmd.Close(2, 'F_SystemSettings', 1)

    Write-Output 'Correcting warehouse record selection.'
    $warehouseList = Open-FormDesign 'SF_WarehouseList'
    foreach ($columnName in @('WarehouseID', 'WarehouseCode', 'WarehouseName', 'Address', 'ContactPerson', 'LocationCount')) {
        $warehouseList.Controls.Item($columnName).OnDblClick = '=OpenSelectedWarehouse()'
    }
    Write-AccessModule -Name 'SF_WarehouseList' -Code @'
Option Compare Database
Option Explicit

Public Function OpenSelectedWarehouse() As Boolean
    On Error GoTo Err_Handler

    Dim selectedID As Long
    Dim detailForm As String

    If IsNull(Me!WarehouseID) Then Exit Function
    selectedID = CLng(Me!WarehouseID)
    detailForm = "F_Warehouses"

    If CurrentProject.AllForms(detailForm).IsLoaded Then
        If Forms(detailForm).Dirty Then Forms(detailForm).Dirty = False
        DoCmd.Close acForm, detailForm, acSaveNo
    End If

    DoCmd.OpenForm detailForm, acNormal, , "[WarehouseID]=" & CStr(selectedID)
    If Forms(detailForm)!WarehouseID <> selectedID Then
        Err.Raise vbObjectError + 2103, "OpenSelectedWarehouse", _
            "The warehouse form did not open the selected warehouse."
    End If

    OpenSelectedWarehouse = True
    Exit Function

Err_Handler:
    MsgBox "The selected warehouse could not be opened." & vbCrLf & _
        "Error Number: " & Err.Number & vbCrLf & _
        "Details: " & Err.Description, vbExclamation, "Warehouses"
    OpenSelectedWarehouse = False
End Function
'@
    $access.DoCmd.Close(2, 'SF_WarehouseList', 1)

    $wideDatasheetForms = @(
        'SF_WarehouseList',
        'SF_TemplateList',
        'SF_SupplierList',
        'SF_StockMovementList',
        'SF_SalesOrderList',
        'SF_ProjectList',
        'SF_ItemSetupList',
        'SF_InvoiceList',
        'SF_InventoryList',
        'SF_CustomerPOList',
        'SF_CustomerList'
    )
    foreach ($formName in $wideDatasheetForms) {
        Write-Output "Configuring horizontal scrolling: $formName."
        $form = Open-FormDesign $formName
        $form.ScrollBars = 3
        $form.FitToScreen = $false
        $form.ViewsAllowed = 2
        $form.AllowDatasheetView = $true
        $access.DoCmd.Close(2, $formName, 1)
    }

    $wideDatasheetHosts = @(
        'F_WarehouseList',
        'F_TemplateList',
        'F_SupplierList',
        'F_StockMovementList',
        'F_SalesOrderList',
        'F_ProjectList',
        'F_ItemSetupList',
        'F_InvoiceList',
        'F_InventoryList',
        'F_CustomerPOList',
        'F_CustomerList'
    )
    foreach ($formName in $wideDatasheetHosts) {
        Write-Output "Configuring host scrolling: $formName."
        $form = Open-FormDesign $formName
        $form.ScrollBars = 3
        $form.FitToScreen = $false
        foreach ($control in $form.Controls) {
            if ($control.ControlType -eq 112) {
                $control.ScrollBars = 3
            }
        }
        $access.DoCmd.Close(2, $formName, 1)
    }

    $access.CloseCurrentDatabase()
    $access.Quit()
    Write-Output "Version 2 list screens, routing, warehouse selection, and datasheet settings were saved."
}
finally {
    if ($database -ne $null) {
        [void][Runtime.InteropServices.Marshal]::ReleaseComObject($database)
    }
    if ($access -ne $null) {
        try {
            $access.CloseCurrentDatabase()
        }
        catch {
        }
        try {
            $access.Quit()
        }
        catch {
        }
        [void][Runtime.InteropServices.Marshal]::ReleaseComObject($access)
    }
    [GC]::Collect()
    [GC]::WaitForPendingFinalizers()
}
