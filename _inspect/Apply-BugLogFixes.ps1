$ErrorActionPreference = 'Stop'
$root = 'C:\Users\user\Downloads\LLC'
$databasePath = Join-Path $root 'LLC.accdb'
$logPath = Join-Path $root '_inspect\buglog-apply.txt'
'' | Set-Content $logPath -Encoding UTF8

function Log([string]$Message) {
    $line = '{0} {1}' -f (Get-Date -Format 'HH:mm:ss'), $Message
    Add-Content $logPath $line
    Write-Host $line
}

Add-Type @'
using System;
using System.Runtime.InteropServices;
public static class BugLogAccessKeys {
    [DllImport("user32.dll")]
    public static extern void keybd_event(byte key, byte scan, uint flags, UIntPtr extra);
}
'@

function FormExists([string]$Name) {
    foreach ($form in $access.CurrentProject.AllForms) {
        if ($form.Name -eq $Name) { return $true }
    }
    return $false
}

function Copy-FormIfMissing([string]$Source, [string]$Destination) {
    if (-not (FormExists $Destination)) {
        $access.DoCmd.CopyObject($null, $Destination, 2, $Source)
        Log "Copied $Source -> $Destination"
    }
}

function Open-FormDesign([string]$Name) {
    $access.DoCmd.OpenForm($Name, 1)
    return $access.Forms.Item($Name)
}

function Get-ModuleText([string]$Name) {
    $module = $access.Forms.Item($Name).Module
    return $module.Lines(1, $module.CountOfLines)
}

function Set-FormModule([string]$Name, [string]$Code) {
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
    $pattern = '(?ms)^(Private|Public) (Sub|Function) ' + [regex]::Escape($Name) +
        '\([^\r\n]*\)\r?\n.*?^End (Sub|Function)[ \t]*\r?\n?'
    $expression = [regex]::new($pattern)
    if ($expression.IsMatch($Module)) {
        return $expression.Replace($Module, $Code.Trim() + "`r`n", 1)
    }
    return $Module.TrimEnd() + "`r`n`r`n" + $Code.Trim() + "`r`n"
}

function Ensure-Query([string]$Name, [string]$Sql) {
    $db = $access.CurrentDb()
    $found = $false
    foreach ($query in $db.QueryDefs) {
        if ($query.Name -eq $Name) { $found = $true; break }
    }
    if ($found) {
        $db.QueryDefs.Item($Name).SQL = $Sql
    } else {
        $db.CreateQueryDef($Name, $Sql) | Out-Null
    }
}

function Add-FieldIfMissing([string]$TableName, [string]$FieldName, [string]$SqlType) {
    $db = $access.CurrentDb()
    $table = $db.TableDefs.Item($TableName)
    foreach ($field in $table.Fields) {
        if ($field.Name -eq $FieldName) { return }
    }
    $db.Execute("ALTER TABLE [$TableName] ADD COLUMN [$FieldName] $SqlType", 128)
    $db.TableDefs.Refresh()
    Log "Added $TableName.$FieldName"
}

function TableExists([string]$Name) {
    $db = $access.CurrentDb()
    foreach ($table in $db.TableDefs) {
        if ($table.Name -eq $Name) { return $true }
    }
    return $false
}

function Ensure-Control($form, [string]$formName, [string]$name, [int]$type, [string]$source, [int]$left, [int]$top, [int]$width, [int]$height) {
    foreach ($existing in $form.Controls) {
        if ($existing.Name -eq $name) { return $existing }
    }
    $created = $access.CreateControl($formName, $type, 0, '', $source, $left, $top, $width, $height)
    $created.Name = $name
    return $created
}

function Style-BlackButton($button, $sample) {
    try {
        $button.FontName = $sample.FontName
        $button.FontSize = $sample.FontSize
        $button.FontWeight = 700
        $button.ForeColor = 16777215
        $button.BackColor = 0
        $button.HoverColor = 11485214
        $button.PressedColor = 0
        $button.BorderColor = 0
        $button.HoverForeColor = 16777215
        $button.PressedForeColor = 16777215
        $button.Gradient = 0
    } catch {}
}

Get-Process MSACCESS -ErrorAction SilentlyContinue | Stop-Process -Force -ErrorAction SilentlyContinue
Start-Sleep -Seconds 2

$access = $null
try {
    $access = New-Object -ComObject Access.Application
    $access.Visible = $true
    Log "Opening $databasePath"
    [BugLogAccessKeys]::keybd_event(0x10, 0, 0, [UIntPtr]::Zero)
    try {
        $access.OpenCurrentDatabase($databasePath)
    } finally {
        [BugLogAccessKeys]::keybd_event(0x10, 0, 2, [UIntPtr]::Zero)
    }
    Log 'Database opened'
    $db = $access.CurrentDb()

    # ------------------------------------------------------------------
    # Schema: optional location fields, status fields, part numbers
    # ------------------------------------------------------------------
    foreach ($fieldName in @('WarehouseID', 'BinNumber', 'ShelfNumber')) {
        try {
            $db.TableDefs.Item('T_Locations').Fields.Item($fieldName).Required = $false
            Log "T_Locations.$fieldName Required=False"
        } catch {
            Log "WARN could not relax T_Locations.$fieldName : $($_.Exception.Message)"
        }
    }
    foreach ($fieldName in @('BinNumber', 'ShelfNumber')) {
        try {
            $db.TableDefs.Item('T_Locations').Fields.Item($fieldName).AllowZeroLength = $true
        } catch {}
    }

    Add-FieldIfMissing 'T_Items' 'ItemStatus' 'TEXT(30)'
    try {
        $db.TableDefs.Item('T_Items').Fields.Item('ItemStatus').DefaultValue = '"Draft"'
        $db.Execute("UPDATE T_Items SET ItemStatus='Complete' WHERE ItemStatus Is Null OR Trim(ItemStatus)=''", 128)
    } catch {
        Log "WARN ItemStatus default: $($_.Exception.Message)"
    }

    if (TableExists 'T_Assemblies') {
        Add-FieldIfMissing 'T_Assemblies' 'AssemblyStatus' 'TEXT(30)'
        try {
            $db.Execute("UPDATE T_Assemblies SET AssemblyStatus='Complete' WHERE AssemblyStatus Is Null OR Trim(AssemblyStatus)=''", 128)
        } catch {}
    }

    Add-FieldIfMissing 'T_AssemblyTemplates' 'ManufacturerPartNumber' 'TEXT(100)'
    Add-FieldIfMissing 'T_AssemblyTemplates' 'VendorPartNumber' 'TEXT(100)'
    Add-FieldIfMissing 'T_AssemblyTemplates' 'InternalPartNumber' 'TEXT(100)'
    Add-FieldIfMissing 'T_AssemblyTemplates' 'AutoGeneratePartNumber' 'YESNO'
    try {
        $db.TableDefs.Item('T_AssemblyTemplates').Fields.Item('AutoGeneratePartNumber').DefaultValue = 'False'
    } catch {}

    Ensure-Query 'Q_V2_LocationList' @'
SELECT L.LocationID, L.LocationCode, L.LocationName, L.WarehouseID,
W.WarehouseName, L.BinNumber, L.ShelfNumber
FROM T_Locations AS L LEFT JOIN T_Warehouses AS W
ON L.WarehouseID = W.WarehouseID
ORDER BY L.LocationCode;
'@
    Ensure-Query 'Q_V2_UnitList' @'
SELECT UnitID, UnitCode, UnitName, Description
FROM T_Units
ORDER BY UnitCode;
'@

    try {
        $customerList = $db.QueryDefs.Item('Q_CustomerList')
        if ($customerList.SQL -notmatch 'MissingFields') {
            $missingFields = @"
IIf(Len(Trim(Nz(C.CustomerName,'')))=0,'Customer Name; ','') &
IIf(Len(Trim(Nz(C.ContactPerson,'')))=0,'Contact Name; ','') &
IIf(Len(Trim(Nz(C.Email,'')))=0,'Email; ','') &
IIf(Len(Trim(Nz(C.BillingAddress,'')))=0,'Billing Address; ','') AS MissingFields,
"@
            $customerList.SQL = $customerList.SQL.Replace(
                'C.CustomerID, C.CustomerName,',
                "C.CustomerID, C.CustomerName, $missingFields"
            )
            Log 'Updated Q_CustomerList MissingFields'
        }
    } catch {
        Log "WARN Q_CustomerList: $($_.Exception.Message)"
    }

    # ------------------------------------------------------------------
    # Backup module
    # ------------------------------------------------------------------
    $backupText = Get-Content -Raw (Join-Path $root '_inspect\modules\M_Backup.txt')
    try { $access.DoCmd.DeleteObject(5, 'M_Backup') } catch {}
    $tempBackup = Join-Path $root '_inspect\modules\M_Backup_load.txt'
    Set-Content -Path $tempBackup -Value $backupText -Encoding Unicode
    $access.LoadFromText(5, 'M_Backup', $tempBackup)
    Log 'Loaded M_Backup'

    # ------------------------------------------------------------------
    # Location setup: optional warehouse/bin/shelf and DAO save (fixes 3021)
    # ------------------------------------------------------------------
    $locationForm = Open-FormDesign 'F_LocationSetup'
    $locationForm.NavigationButtons = $true
    $locationForm.RecordSelectors = $true
    $locationForm.DataEntry = $false
    $locationCode = Get-ModuleText 'F_LocationSetup'
    $locationCode = $locationCode.Replace(
        'UX_MarkRequired Me, "LocationName", "cboWarehouseID"',
        'UX_MarkRequired Me, "LocationName"'
    )
    $locationCode = Set-Procedure -Module $locationCode -Name 'Form_BeforeInsert' -Code @'
Private Sub Form_BeforeInsert(Cancel As Integer)
    On Error Resume Next
    If Not IsNull(Me.CreatedOn) Then Exit Sub
    Me.CreatedOn = Now()
    On Error GoTo 0
End Sub
'@
    $locationCode = Set-Procedure -Module $locationCode -Name 'Form_Error' -Code @'
Private Sub Form_Error(DataErr As Integer, Response As Integer)
    If DataErr = 3022 Then
        MsgBox "The location code already exists.", vbExclamation, "Duplicate Location"
        Response = acDataErrContinue
    ElseIf DataErr = 3021 Then
        Response = acDataErrContinue
    Else
        Response = acDataErrDisplay
    End If
End Sub
'@
    $locationCode = Set-Procedure -Module $locationCode -Name 'SaveLocationRecord' -Code @'
Private Function SaveLocationRecord() As Boolean
    On Error GoTo Err_Handler

    Dim db As DAO.Database
    Dim rs As DAO.Recordset
    Dim locationIDValue As Long
    Dim locationNameValue As String
    Dim locationCodeValue As String
    Dim warehouseIDValue As Variant
    Dim binValue As Variant
    Dim shelfValue As Variant
    Dim notesValue As Variant

    SaveLocationRecord = False
    NormalizeLocationValues
    EnsureLocationCode

    If Not ValidateLocationRecord Then
        Exit Function
    End If

    locationNameValue = Trim$(Nz(ReadFormValue("LocationName", Array("txtLocationName", "LocationName")), vbNullString))
    locationCodeValue = UCase$(Trim$(Nz(ReadFormValue("LocationCode", Array("txtLocationCode", "LocationCode")), vbNullString)))
    If Nz(Me.cboWarehouseID, 0) > 0 Then
        warehouseIDValue = CLng(Me.cboWarehouseID)
    Else
        warehouseIDValue = Null
    End If
    binValue = Trim$(Nz(ReadFormValue("BinNumber", Array("txtBinNumber", "BinNumber")), vbNullString))
    If Len(binValue) = 0 Then binValue = Null
    shelfValue = Trim$(Nz(ReadFormValue("ShelfNumber", Array("txtShelfNumber", "ShelfNumber")), vbNullString))
    If Len(shelfValue) = 0 Then shelfValue = Null
    notesValue = Trim$(Nz(ReadFormValue("Notes", Array("txtNotes", "txtNOtes", "Notes")), vbNullString))
    If Len(notesValue) = 0 Then notesValue = Null

    locationIDValue = CurrentLocationID()
    Set db = CurrentDb()

    If locationIDValue > 0 Then
        Set rs = db.OpenRecordset("SELECT * FROM T_Locations WHERE LocationID=" & locationIDValue, dbOpenDynaset)
        If rs.EOF Then
            Err.Raise vbObjectError + 7501, "SaveLocationRecord", "The location record was not found."
        End If
        rs.Edit
    Else
        Set rs = db.OpenRecordset("T_Locations", dbOpenDynaset)
        rs.AddNew
        On Error Resume Next
        rs.Fields("CreatedOn").Value = Now()
        On Error GoTo Err_Handler
    End If

    rs.Fields("LocationName").Value = locationNameValue
    rs.Fields("LocationCode").Value = locationCodeValue
    On Error Resume Next
    rs.Fields("WarehouseID").Value = warehouseIDValue
    rs.Fields("BinNumber").Value = binValue
    rs.Fields("ShelfNumber").Value = shelfValue
    rs.Fields("Notes").Value = notesValue
    rs.Fields("UpdatedOn").Value = Now()
    On Error GoTo Err_Handler
    rs.Update

    If locationIDValue = 0 Then
        rs.Bookmark = rs.LastModified
        locationIDValue = Nz(rs.Fields("LocationID").Value, 0)
    End If
    rs.Close

    If Me.Dirty Then Me.Undo
    SaveLocationRecord = (locationIDValue > 0)

Exit_Handler:
    On Error Resume Next
    rs.Close
    Set rs = Nothing
    Set db = Nothing
    Exit Function

Err_Handler:
    MsgBox "The location record could not be saved." & vbCrLf & vbCrLf & _
        "Error Number: " & Err.Number & vbCrLf & _
        "Details: " & Err.Description, vbExclamation, "Location Setup"
    Resume Exit_Handler
End Function
'@
    $locationCode = Set-Procedure -Module $locationCode -Name 'ValidateLocationRecord' -Code @'
Private Function ValidateLocationRecord() As Boolean
    Dim locationIDValue As Long
    Dim locationNameValue As String
    Dim locationCodeValue As String

    ValidateLocationRecord = False

    locationNameValue = Trim$(Nz(ReadFormValue("LocationName", Array("txtLocationName", "LocationName")), vbNullString))
    If Len(locationNameValue) = 0 Then
        MsgBox "Enter the location name.", vbExclamation, "Location Name Required"
        FocusControlIfExists "txtLocationName", "LocationName"
        Exit Function
    End If

    locationCodeValue = UCase$(Trim$(Nz(ReadFormValue("LocationCode", Array("txtLocationCode", "LocationCode")), vbNullString)))
    If Len(locationCodeValue) = 0 Then
        MsgBox "The location code could not be generated.", vbExclamation, "Location Code Required"
        Exit Function
    End If

    locationIDValue = CurrentLocationID()
    If DCount("*", "T_Locations", "[LocationCode]=" & SqlText(locationCodeValue) & " AND [LocationID]<>" & CStr(locationIDValue)) > 0 Then
        MsgBox "The location code already exists.", vbExclamation, "Duplicate Location Code"
        Exit Function
    End If

    ValidateLocationRecord = True
End Function
'@
    $locationCode = Set-Procedure -Module $locationCode -Name 'ApplyNewRecordDefaults' -Code @'
Private Sub ApplyNewRecordDefaults()
    If Not Me.NewRecord Then Exit Sub
    On Error Resume Next
    If IsNull(Me.CreatedOn) Then Me.CreatedOn = Now()
    On Error GoTo 0
End Sub
'@
    $locationCode = Set-Procedure -Module $locationCode -Name 'CurrentLocationID' -Code @'
Private Function CurrentLocationID() As Long
    On Error Resume Next
    If Me.NewRecord Then
        CurrentLocationID = 0
        Exit Function
    End If
    CurrentLocationID = Nz(Me!LocationID, 0)
    On Error GoTo 0
End Function
'@
    $locationCode = Set-Procedure -Module $locationCode -Name 'WriteFormValue' -Code @'
Private Sub WriteFormValue(ByVal fieldName As String, ByVal candidateControls As Variant, ByVal newValue As Variant, ByVal optionalField As Boolean)
    Dim controlName As Variant
    For Each controlName In candidateControls
        If ControlExists(CStr(controlName)) Then
            Me.Controls(CStr(controlName)).Value = newValue
            Exit Sub
        End If
    Next controlName
    If optionalField Then Exit Sub
    On Error Resume Next
    If Not Me.NewRecord Then Me.Recordset.Fields(fieldName).Value = newValue
    If Err.Number <> 0 And Not optionalField Then
        Err.Raise vbObjectError + 7404, "WriteFormValue", "The field or control '" & fieldName & "' could not be found."
    End If
    On Error GoTo 0
End Sub
'@
    $locationCode = Set-Procedure -Module $locationCode -Name 'SetInitialFocus' -Code @'
Private Sub SetInitialFocus()
    FocusControlIfExists "txtLocationName", "LocationName"
    If Nz(Me.cboWarehouseID, 0) <= 0 And Len(Nz(ReadFormValue("LocationName", Array("txtLocationName", "LocationName")), vbNullString)) = 0 Then
        On Error Resume Next
        Me.cboWarehouseID.SetFocus
    End If
End Sub
'@
    Set-FormModule 'F_LocationSetup' $locationCode
    $access.DoCmd.Close(2, 'F_LocationSetup', 1)
    Log 'Updated F_LocationSetup'

    # ------------------------------------------------------------------
    # Warehouse opens the selected record
    # ------------------------------------------------------------------
    $warehouseForm = Open-FormDesign 'F_Warehouses'
    $warehouseForm.Filter = ''
    $warehouseForm.FilterOn = $false
    $warehouseForm.FilterOnLoad = $false
    $warehouseCode = Get-ModuleText 'F_Warehouses'
    $warehouseCode = $warehouseCode.Replace(
        "    Me.Filter = `"`"`r`n    Me.FilterOn = False`r`n",
        "    If Len(Trim$(Nz(Me.Filter, `"`"))) > 0 Then`r`n        Me.FilterOn = True`r`n    ElseIf IsNumeric(Nz(Me.OpenArgs, `"`")) Then`r`n        Me.Filter = `"[WarehouseID]=`" & CLng(Me.OpenArgs)`r`n        Me.FilterOn = True`r`n    End If`r`n"
    )
    if ($warehouseCode -notmatch 'Private Sub Form_Open') {
        $warehouseCode = Set-Procedure -Module $warehouseCode -Name 'Form_Open' -Code @'
Private Sub Form_Open(Cancel As Integer)
    On Error Resume Next
    If Len(Trim$(Nz(Me.Filter, ""))) > 0 Then
        Me.FilterOn = True
    End If
End Sub
'@
    }
    Set-FormModule 'F_Warehouses' $warehouseCode
    $access.DoCmd.Close(2, 'F_Warehouses', 1)

    $sfWarehouse = Open-FormDesign 'SF_WarehouseList'
    foreach ($columnName in @('WarehouseID', 'WarehouseCode', 'WarehouseName', 'Address', 'ContactPerson', 'LocationCount')) {
        try { $sfWarehouse.Controls.Item($columnName).OnDblClick = '=OpenSelectedWarehouse()' } catch {}
    }
    $sfCode = Get-ModuleText 'SF_WarehouseList'
    $sfCode = Set-Procedure -Module $sfCode -Name 'OpenSelectedWarehouse' -Code @'
Public Function OpenSelectedWarehouse() As Boolean
    On Error GoTo Err_Handler
    Dim selectedID As Long
    Dim detailForm As String
    If IsNull(Me!WarehouseID) Then Exit Function
    selectedID = CLng(Me!WarehouseID)
    detailForm = "F_Warehouses"
    If CurrentProject.AllForms(detailForm).IsLoaded Then
        On Error Resume Next
        If Forms(detailForm).Dirty Then Forms(detailForm).Undo
        DoCmd.Close acForm, detailForm, acSaveNo
        On Error GoTo Err_Handler
    End If
    DoCmd.OpenForm detailForm, acNormal, , "[WarehouseID]=" & CStr(selectedID), , , CStr(selectedID)
    Forms(detailForm).Filter = "[WarehouseID]=" & CStr(selectedID)
    Forms(detailForm).FilterOn = True
    If Nz(Forms(detailForm)!WarehouseID, 0) <> selectedID Then
        Err.Raise vbObjectError + 2103, "OpenSelectedWarehouse", "The warehouse form did not open the selected warehouse."
    End If
    OpenSelectedWarehouse = True
    Exit Function
Err_Handler:
    MsgBox "The selected warehouse could not be opened." & vbCrLf & _
        "Error Number: " & Err.Number & vbCrLf & "Details: " & Err.Description, _
        vbExclamation, "Warehouses"
    OpenSelectedWarehouse = False
End Function
'@
    Set-FormModule 'SF_WarehouseList' $sfCode
    $access.DoCmd.Close(2, 'SF_WarehouseList', 1)
    Log 'Updated warehouse record selection'

    # ------------------------------------------------------------------
    # Unit and location list screens
    # ------------------------------------------------------------------
    Copy-FormIfMissing 'SF_WarehouseList' 'SF_LocationList'
    Copy-FormIfMissing 'F_WarehouseList' 'F_LocationList'
    Copy-FormIfMissing 'SF_WarehouseList' 'SF_UnitList'
    Copy-FormIfMissing 'F_WarehouseList' 'F_UnitList'

    $locationSubform = Open-FormDesign 'SF_LocationList'
    $locationSubform.RecordSource = 'Q_V2_LocationList'
    $locationSubform.DefaultView = 2
    $locationSubform.ViewsAllowed = 2
    $locationSubform.AllowDatasheetView = $true
    $locationSubform.ScrollBars = 3
    $locationSubform.FitToScreen = $false
    $locationColumns = @(
        @{ Old = @('WarehouseID','LocationID'); New = 'LocationID'; Caption = 'LocationID'; Width = 900; Hidden = $true },
        @{ Old = @('WarehouseCode','LocationCode'); New = 'LocationCode'; Caption = 'Location Code'; Width = 1900; Hidden = $false },
        @{ Old = @('WarehouseName','LocationName'); New = 'LocationName'; Caption = 'Location Name'; Width = 2800; Hidden = $false },
        @{ Old = @('Address','WarehouseName'); New = 'WarehouseName'; Caption = 'Warehouse'; Width = 2500; Hidden = $false },
        @{ Old = @('ContactPerson','BinNumber'); New = 'BinNumber'; Caption = 'Bin'; Width = 1500; Hidden = $false },
        @{ Old = @('LocationCount','ShelfNumber'); New = 'ShelfNumber'; Caption = 'Shelf'; Width = 1500; Hidden = $false }
    )
    foreach ($column in $locationColumns) {
        $control = $null
        foreach ($name in $column.Old) {
            foreach ($candidate in $locationSubform.Controls) {
                if ($candidate.Name -eq $name) { $control = $candidate; break }
            }
            if ($control) { break }
        }
        if (-not $control) { continue }
        $control.Name = $column.New
        $control.ControlSource = $column.New
        $control.Width = $column.Width
        $control.Visible = -not $column.Hidden
        $control.OnDblClick = '=OpenSelectedLocation()'
        try {
            $label = $locationSubform.Controls.Item($column.New + '_Label')
            $label.Caption = $column.Caption
            $label.Visible = -not $column.Hidden
        } catch {}
    }
    Set-FormModule 'SF_LocationList' @'
Option Compare Database
Option Explicit
Public Function OpenSelectedLocation() As Boolean
    On Error GoTo Err_Handler
    Dim selectedID As Long
    If IsNull(Me!LocationID) Then Exit Function
    selectedID = CLng(Me!LocationID)
    If CurrentProject.AllForms("F_LocationSetup").IsLoaded Then
        If Forms("F_LocationSetup").Dirty Then Forms("F_LocationSetup").Undo
        DoCmd.Close acForm, "F_LocationSetup", acSaveNo
    End If
    DoCmd.OpenForm "F_LocationSetup", acNormal, , "[LocationID]=" & CStr(selectedID)
    OpenSelectedLocation = True
    Exit Function
Err_Handler:
    MsgBox "The selected location could not be opened." & vbCrLf & Err.Description, vbExclamation, "Locations"
    OpenSelectedLocation = False
End Function
'@
    $access.DoCmd.Close(2, 'SF_LocationList', 1)

    $unitSubform = Open-FormDesign 'SF_UnitList'
    $unitSubform.RecordSource = 'Q_V2_UnitList'
    $unitSubform.DefaultView = 2
    $unitSubform.ViewsAllowed = 2
    $unitSubform.AllowDatasheetView = $true
    $unitSubform.AllowEdits = $true
    $unitSubform.AllowAdditions = $true
    $unitSubform.ScrollBars = 3
    $unitSubform.FitToScreen = $false
    $unitColumns = @(
        @{ Old = @('WarehouseID','LocationID','UnitID'); New = 'UnitID'; Caption = 'UnitID'; Width = 900; Hidden = $true },
        @{ Old = @('WarehouseCode','LocationCode','UnitCode'); New = 'UnitCode'; Caption = 'Unit Code'; Width = 2200; Hidden = $false },
        @{ Old = @('WarehouseName','LocationName','UnitName'); New = 'UnitName'; Caption = 'Unit Name'; Width = 3000; Hidden = $false },
        @{ Old = @('Address','WarehouseName','Description'); New = 'Description'; Caption = 'Description'; Width = 4200; Hidden = $false },
        @{ Old = @('ContactPerson','BinNumber','UnitUnused1'); New = 'UnitUnused1'; Caption = ''; Width = 1000; Hidden = $true },
        @{ Old = @('LocationCount','ShelfNumber','UnitUnused2'); New = 'UnitUnused2'; Caption = ''; Width = 1000; Hidden = $true }
    )
    foreach ($column in $unitColumns) {
        $control = $null
        foreach ($name in $column.Old) {
            foreach ($candidate in $unitSubform.Controls) {
                if ($candidate.Name -eq $name) { $control = $candidate; break }
            }
            if ($control) { break }
        }
        if (-not $control) { continue }
        $control.Name = $column.New
        if ($column.New -in @('UnitID','UnitCode','UnitName','Description')) {
            $control.ControlSource = $column.New
            $control.OnDblClick = '=OpenSelectedUnit()'
        } else {
            $control.ControlSource = ''
        }
        $control.Width = $column.Width
        $control.Visible = -not $column.Hidden
    }
    Set-FormModule 'SF_UnitList' @'
Option Compare Database
Option Explicit
Public Function OpenSelectedUnit() As Boolean
    On Error GoTo Err_Handler
    Dim selectedID As Long
    If IsNull(Me!UnitID) Then Exit Function
    selectedID = CLng(Me!UnitID)
    If CurrentProject.AllForms("F_UnitSetup").IsLoaded Then
        If Forms("F_UnitSetup").Dirty Then Forms("F_UnitSetup").Undo
        DoCmd.Close acForm, "F_UnitSetup", acSaveNo
    End If
    DoCmd.OpenForm "F_UnitSetup", acNormal, , "[UnitID]=" & CStr(selectedID)
    OpenSelectedUnit = True
    Exit Function
Err_Handler:
    MsgBox "The selected unit could not be opened." & vbCrLf & Err.Description, vbExclamation, "Units"
    OpenSelectedUnit = False
End Function
'@
    $access.DoCmd.Close(2, 'SF_UnitList', 1)

    foreach ($list in @(
        @{ Form = 'F_LocationList'; SubformCandidates = @('subLocationList','subWarehouseList','subUnitList'); Source = 'SF_LocationList'; Caption = 'Location List'; Button = 'ADD LOCATION' },
        @{ Form = 'F_UnitList'; SubformCandidates = @('subUnitList','subWarehouseList','subLocationList'); Source = 'SF_UnitList'; Caption = 'Unit List'; Button = 'ADD UNIT' }
    )) {
        $hostForm = Open-FormDesign $list.Form
        $hostForm.Caption = $list.Caption
        $hostForm.ScrollBars = 3
        $hostForm.FitToScreen = $false
        try { $hostForm.Controls.Item('btnNew').Caption = $list.Button } catch {}
        $subformControl = $null
        foreach ($candidate in $list.SubformCandidates) {
            foreach ($ctl in $hostForm.Controls) {
                if ($ctl.Name -eq $candidate) { $subformControl = $ctl; break }
            }
            if ($subformControl) { break }
        }
        if ($subformControl) {
            if ($list.Form -eq 'F_LocationList') { $subformControl.Name = 'subLocationList' } else { $subformControl.Name = 'subUnitList' }
            $subformControl.SourceObject = 'Form.' + $list.Source
            $subformControl.ScrollBars = 3
        }
        if ($list.Form -eq 'F_LocationList') {
            Set-FormModule $list.Form @'
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
    If Me.Dirty Then Me.Undo
    DoCmd.Close acForm, Me.Name
End Sub
Private Sub ApplyListFilter()
    Dim searchText As String
    searchText = Replace(Trim$(Nz(Me.txtSearch.Value, vbNullString)), "'", "''")
    With Me.subLocationList.Form
        If Len(searchText) = 0 Then
            .FilterOn = False
        Else
            .Filter = "[LocationCode] Like '*" & searchText & "*' OR [LocationName] Like '*" & searchText & "*' OR Nz([WarehouseName],'') Like '*" & searchText & "*'"
            .FilterOn = True
        End If
    End With
End Sub
'@
        } else {
            Set-FormModule $list.Form @'
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
    If Me.Dirty Then Me.Undo
    DoCmd.Close acForm, Me.Name
End Sub
Private Sub ApplyListFilter()
    Dim searchText As String
    searchText = Replace(Trim$(Nz(Me.txtSearch.Value, vbNullString)), "'", "''")
    With Me.subUnitList.Form
        If Len(searchText) = 0 Then
            .FilterOn = False
        Else
            .Filter = "[UnitCode] Like '*" & searchText & "*' OR [UnitName] Like '*" & searchText & "*' OR Nz([Description],'') Like '*" & searchText & "*'"
            .FilterOn = True
        End If
    End With
End Sub
'@
        }
        $access.DoCmd.Close(2, $list.Form, 1)
        Log "Configured $($list.Form)"
    }

    $unitSetup = Open-FormDesign 'F_UnitSetup'
    $unitSetup.NavigationButtons = $true
    $unitSetup.RecordSelectors = $true
    $unitSetup.AllowAdditions = $true
    $unitSetup.DataEntry = $false
    $access.DoCmd.Close(2, 'F_UnitSetup', 1)

    # ------------------------------------------------------------------
    # Customer list: missing-fields column and incomplete OpenArgs
    # ------------------------------------------------------------------
    $sfCustomer = Open-FormDesign 'SF_CustomerList'
    $missingControl = $null
    foreach ($ctl in $sfCustomer.Controls) {
        if ($ctl.Name -eq 'txtMissingFields') { $missingControl = $ctl }
    }
    if (-not $missingControl) {
        $missingControl = $access.CreateControl('SF_CustomerList', 109, 0, '', 'MissingFields', 15000, 0, 3000, 300)
        $missingControl.Name = 'txtMissingFields'
    }
    $missingControl.ControlSource = 'MissingFields'
    try { $missingControl.ColumnWidth = 3000 } catch {}
    $access.DoCmd.Close(2, 'SF_CustomerList', 1)

    $customerListForm = Open-FormDesign 'F_CustomerList'
    $customerCode = Get-ModuleText 'F_CustomerList'
    $customerCode = Set-Procedure -Module $customerCode -Name 'Form_Load' -Code @'
Private Sub Form_Load()
    If StrComp(Nz(Me.OpenArgs, vbNullString), "INCOMPLETE", vbTextCompare) = 0 Then
        On Error Resume Next
        With Me.Controls("subCustomerList").Form
            .Filter = "Len(Trim(Nz([MissingFields],'')))>0"
            .FilterOn = True
        End With
        On Error GoTo 0
    Else
        ApplyListFilter
    End If
End Sub
'@
    Set-FormModule 'F_CustomerList' $customerCode
    $access.DoCmd.Close(2, 'F_CustomerList', 1)

    # ------------------------------------------------------------------
    # Item status + assembly template part numbers
    # ------------------------------------------------------------------
    $itemForm = Open-FormDesign 'F_ItemSetup'
    $itemStatus = $null
    foreach ($ctl in $itemForm.Controls) { if ($ctl.Name -eq 'cboItemStatus') { $itemStatus = $ctl } }
    if (-not $itemStatus) {
        $itemStatus = $access.CreateControl('F_ItemSetup', 111, 0, '', 'ItemStatus', 9738, 2280, 3648, 312)
        $itemStatus.Name = 'cboItemStatus'
        $statusLabel = $access.CreateControl('F_ItemSetup', 100, 0, '', '', 7200, 2280, 2448, 312)
        $statusLabel.Name = 'ItemStatus_Label'
        $statusLabel.Caption = 'Status'
    }
    $itemStatus.RowSourceType = 'Value List'
    $itemStatus.RowSource = 'Draft;In Review;Pending Approval;Complete'
    $itemStatus.LimitToList = $true
    $itemStatus.DefaultValue = '"Draft"'
    $access.DoCmd.Close(2, 'F_ItemSetup', 1)

    $templateForm = Open-FormDesign 'F_AssemblyTemplate'
    $referenceText = $null
    foreach ($ctl in $templateForm.Controls) { if ($ctl.Name -eq 'txtRevisionNo') { $referenceText = $ctl } }
    $partControls = @(
        @{ Name='txtManufacturerPartNumber'; Source='ManufacturerPartNumber'; Caption='Manufacturer Part Number'; Top=2400 },
        @{ Name='txtVendorPartNumber'; Source='VendorPartNumber'; Caption='Vendor Part Number'; Top=2820 },
        @{ Name='txtInternalPartNumber'; Source='InternalPartNumber'; Caption='Internal Part Number'; Top=3240 }
    )
    foreach ($part in $partControls) {
        $exists = $false
        foreach ($ctl in $templateForm.Controls) { if ($ctl.Name -eq $part.Name) { $exists = $true } }
        if (-not $exists) {
            $control = $access.CreateControl('F_AssemblyTemplate', 109, 0, '', $part.Source, 10662, $part.Top, 2052, 312)
            $control.Name = $part.Name
            $label = $access.CreateControl('F_AssemblyTemplate', 100, 0, '', '', 8124, $part.Top, 2448, 312)
            $label.Name = $part.Name + '_Label'
            $label.Caption = $part.Caption
        }
    }
    $autoPart = $null
    foreach ($ctl in $templateForm.Controls) { if ($ctl.Name -eq 'chkAutoGeneratePartNumber') { $autoPart = $ctl } }
    if (-not $autoPart) {
        $autoPart = $access.CreateControl('F_AssemblyTemplate', 106, 0, '', 'AutoGeneratePartNumber', 8124, 3660, 300, 312)
        $autoPart.Name = 'chkAutoGeneratePartNumber'
        $autoLabel = $access.CreateControl('F_AssemblyTemplate', 100, 0, '', '', 8460, 3660, 3750, 312)
        $autoLabel.Name = 'chkAutoGeneratePartNumber_Label'
        $autoLabel.Caption = 'Auto Generate Part Number'
    }
    $autoPart.AfterUpdate = '[Event Procedure]'
    $access.DoCmd.Close(2, 'F_AssemblyTemplate', 1)

    $templateForm = Open-FormDesign 'F_AssemblyTemplate'
    $templateCode = Get-ModuleText 'F_AssemblyTemplate'
    $templateCode = $templateCode.Replace(
        '.RowSource = "Draft;Active;On Hold;Retired"',
        '.RowSource = "Draft;In Review;Pending Approval;Complete"'
    )
    $templateCode = Set-Procedure -Module $templateCode -Name 'chkAutoGeneratePartNumber_AfterUpdate' -Code @'
Private Sub chkAutoGeneratePartNumber_AfterUpdate()
    UpdatePartNumberState
    If Nz(Me.chkAutoGeneratePartNumber, False) Then
        EnsureGeneratedPartNumbers
    End If
End Sub
'@
    $templateCode = Set-Procedure -Module $templateCode -Name 'UpdatePartNumberState' -Code @'
Private Sub UpdatePartNumberState()
    On Error Resume Next
    Dim autoGenerate As Boolean
    autoGenerate = Nz(Me.chkAutoGeneratePartNumber, False)
    Me.txtManufacturerPartNumber.Locked = autoGenerate
    Me.txtVendorPartNumber.Locked = autoGenerate
    Me.txtInternalPartNumber.Locked = autoGenerate
    On Error GoTo 0
End Sub
'@
    $templateCode = Set-Procedure -Module $templateCode -Name 'EnsureGeneratedPartNumbers' -Code @'
Private Sub EnsureGeneratedPartNumbers()
    On Error Resume Next
    Dim codeValue As String
    If Not Nz(Me.chkAutoGeneratePartNumber, False) Then Exit Sub
    codeValue = UCase$(Trim$(Nz(Me.txtTemplateCode, vbNullString)))
    If Len(codeValue) = 0 Then codeValue = "TEMPLATE"
    If Len(Trim$(Nz(Me.txtInternalPartNumber, vbNullString))) = 0 Then
        Me.txtInternalPartNumber = "ASM-" & codeValue
    End If
    If Len(Trim$(Nz(Me.txtManufacturerPartNumber, vbNullString))) = 0 Then
        Me.txtManufacturerPartNumber = "MFR-" & codeValue
    End If
    If Len(Trim$(Nz(Me.txtVendorPartNumber, vbNullString))) = 0 Then
        Me.txtVendorPartNumber = "VND-" & codeValue
    End If
    On Error GoTo 0
End Sub
'@
    if ($templateCode -notmatch 'UpdatePartNumberState') {
        $templateCode = $templateCode.Replace('SetSubformAvailability', "SetSubformAvailability`r`n    UpdatePartNumberState")
    }
    $oldNormalize = 'Private Sub NormalizeTemplateValues()'
    if ($templateCode -match [regex]::Escape($oldNormalize)) {
        $templateCode = Set-Procedure -Module $templateCode -Name 'NormalizeTemplateValues' -Code @'
Private Sub NormalizeTemplateValues()
    Me.txtTemplateName = Trim$(Nz(Me.txtTemplateName, vbNullString))
    If Len(Trim$(Nz(Me.txtTemplateCode, vbNullString))) = 0 Then
        Me.txtTemplateCode = GenerateTemplateCodeFromName(Me.txtTemplateName)
    End If
    Me.txtTemplateCode = UCase$(Trim$(Nz(Me.txtTemplateCode, vbNullString)))
    Me.txtRevisionNo = Trim$(Nz(Me.txtRevisionNo, vbNullString))
    Me.cboStatus = Trim$(Nz(Me.cboStatus, vbNullString))
    On Error Resume Next
    Me.txtManufacturerPartNumber = Trim$(Nz(Me.txtManufacturerPartNumber, vbNullString))
    Me.txtVendorPartNumber = Trim$(Nz(Me.txtVendorPartNumber, vbNullString))
    Me.txtInternalPartNumber = Trim$(Nz(Me.txtInternalPartNumber, vbNullString))
    Me.txtNOtes = Trim$(Nz(Me.txtNOtes, vbNullString))
    On Error GoTo 0
    If IsNull(Me.txtOutputQuantity) Then Me.txtOutputQuantity = 1
    If IsNull(Me.txtEstimatedLaborHours) Then Me.txtEstimatedLaborHours = 0
    EnsureGeneratedPartNumbers
End Sub
'@
    }
    Set-FormModule 'F_AssemblyTemplate' $templateCode
    $access.DoCmd.Close(2, 'F_AssemblyTemplate', 1)
    Log 'Updated assembly template part numbers'

    # ------------------------------------------------------------------
    # System Settings: clickable counters, backup, draft rules
    # ------------------------------------------------------------------
    $settings = Open-FormDesign 'F_SystemSettings'
    foreach ($labelName in @(
        'lblCustomerStatus','lblSupplierStatus','lblCategoryStatus','lblUnitStatus',
        'lblItemStatus','lblTemplateStatus','lblWarehouseStatus','lblLocationStatus'
    )) {
        try {
            $label = $settings.Controls.Item($labelName)
            $label.OnClick = '[Event Procedure]'
            $label.ControlTipText = 'Click to review these records.'
            try { $label.IsHyperlink = $true } catch {}
        } catch {
            Log "WARN missing $labelName"
        }
    }

    $sampleButton = $settings.Controls.Item('cmdUserManagement')
    $cmdBackup = $null
    $cmdRestore = $null
    foreach ($ctl in $settings.Controls) {
        if ($ctl.Name -eq 'cmdBackupData') { $cmdBackup = $ctl }
        if ($ctl.Name -eq 'cmdRestoreData') { $cmdRestore = $ctl }
    }
    if (-not $cmdBackup) {
        $cmdBackup = $access.CreateControl('F_SystemSettings', 104, 1, '', '', 11340, 360, 2100, 540)
        $cmdBackup.Name = 'cmdBackupData'
    }
    $cmdBackup.Caption = 'Backup Data'
    $cmdBackup.OnClick = '[Event Procedure]'
    Style-BlackButton $cmdBackup $sampleButton
    if (-not $cmdRestore) {
        $cmdRestore = $access.CreateControl('F_SystemSettings', 104, 1, '', '', 13500, 360, 2100, 540)
        $cmdRestore.Name = 'cmdRestoreData'
    }
    $cmdRestore.Caption = 'Restore Data'
    $cmdRestore.OnClick = '[Event Procedure]'
    Style-BlackButton $cmdRestore $sampleButton

    $settingsCode = Get-ModuleText 'F_SystemSettings'
    $settingsCode = Set-Procedure -Module $settingsCode -Name 'CustomerIncompleteCriteria' -Code @'
Private Function CustomerIncompleteCriteria() As String
    CustomerIncompleteCriteria = AnyMissing( _
        BlankTextCriteria("CustomerName"), _
        BlankTextCriteria("ContactPerson"), _
        BlankTextCriteria("Email"), _
        BlankTextCriteria("BillingAddress"))
End Function
'@
    $settingsCode = Set-Procedure -Module $settingsCode -Name 'ItemIncompleteCriteria' -Code @'
Private Function ItemIncompleteCriteria() As String
    ItemIncompleteCriteria = _
        "(Nz([ItemStatus],'Draft')<>'Draft') AND (" & _
        AnyMissing( _
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
            "(Nz([RequiresSerial],0)<>0 AND " & BlankTextCriteria("SerialNumber") & ")") & ")"
End Function
'@
    $settingsCode = Set-Procedure -Module $settingsCode -Name 'TemplateIncompleteCriteria' -Code @'
Private Function TemplateIncompleteCriteria() As String
    TemplateIncompleteCriteria = _
        "(Nz([Status],'Draft')<>'Draft') AND (" & _
        AnyMissing( _
            BlankTextCriteria("TemplateCode"), _
            BlankTextCriteria("TemplateName"), _
            MissingKeyCriteria("CustomerID"), _
            "([OutputQuantity] Is Null OR [OutputQuantity]<=0)", _
            BlankTextCriteria("RevisionNo"), _
            "[EstimatedLaborHours] Is Null", _
            "[EstimatedMaterialCost] Is Null", _
            "[EstimatedTotalCost] Is Null") & ")"
End Function
'@
    $settingsCode = Set-Procedure -Module $settingsCode -Name 'LocationIncompleteCriteria' -Code @'
Private Function LocationIncompleteCriteria() As String
    LocationIncompleteCriteria = AnyMissing( _
        BlankTextCriteria("LocationCode"), _
        BlankTextCriteria("LocationName"))
End Function
'@
    $settingsCode = Set-Procedure -Module $settingsCode -Name 'OpenReadinessForm' -Code @'
Private Sub OpenReadinessForm(ByVal formName As String, Optional ByVal subformName As String = "", Optional ByVal filterText As String = "", Optional ByVal openArgs As String = "")
    On Error GoTo Err_Handler
    If CurrentProject.AllForms(formName).IsLoaded Then
        DoCmd.Close acForm, formName, acSaveNo
    End If
    If Len(openArgs) > 0 Then
        DoCmd.OpenForm formName, acNormal, , , , , openArgs
    Else
        DoCmd.OpenForm formName
    End If
    If Len(subformName) > 0 Then
        On Error Resume Next
        With Forms(formName).Controls(subformName).Form
            .Filter = filterText
            .FilterOn = (Len(filterText) > 0)
        End With
        On Error GoTo Err_Handler
    End If
    Exit Sub
Err_Handler:
    MsgBox "Could not open " & formName & "." & vbCrLf & Err.Description, vbExclamation, "System Setup"
End Sub
'@
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
    MsgBox "The selected setup action could not be opened." & vbCrLf & Err.Description, vbExclamation, "System Setup"
End Sub
'@
    $settingsCode = Set-Procedure -Module $settingsCode -Name 'lblUnitStatus_Click' -Code @'
Private Sub lblUnitStatus_Click()
    DoCmd.OpenForm "F_UnitList"
End Sub
'@
    $settingsCode = Set-Procedure -Module $settingsCode -Name 'lblLocationStatus_Click' -Code @'
Private Sub lblLocationStatus_Click()
    Dim filterText As String
    If mLocationIncomplete > 0 Then filterText = LocationIncompleteCriteria()
    OpenReadinessForm "F_LocationList", "subLocationList", filterText
End Sub
'@
    $settingsCode = Set-Procedure -Module $settingsCode -Name 'lblCustomerStatus_Click' -Code @'
Private Sub lblCustomerStatus_Click()
    If mCustomerIncomplete > 0 Then
        OpenReadinessForm "F_CustomerList", "subCustomerList", "Len(Trim(Nz([MissingFields],'')))>0", "INCOMPLETE"
    Else
        DoCmd.OpenForm "F_CustomerList"
    End If
End Sub
'@
    $settingsCode = Set-Procedure -Module $settingsCode -Name 'lblSupplierStatus_Click' -Code @'
Private Sub lblSupplierStatus_Click()
    Dim filterText As String
    If mSupplierIncomplete > 0 Then filterText = SupplierIncompleteCriteria()
    OpenReadinessForm "F_SupplierList", "subSupplierList", filterText
End Sub
'@
    $settingsCode = Set-Procedure -Module $settingsCode -Name 'lblCategoryStatus_Click' -Code @'
Private Sub lblCategoryStatus_Click()
    DoCmd.OpenForm "F_CategoriesList"
End Sub
'@
    $settingsCode = Set-Procedure -Module $settingsCode -Name 'lblItemStatus_Click' -Code @'
Private Sub lblItemStatus_Click()
    Dim filterText As String
    If mItemIncomplete > 0 Then filterText = ItemIncompleteCriteria()
    OpenReadinessForm "F_ItemSetupList", "subItemSetupList", filterText
End Sub
'@
    $settingsCode = Set-Procedure -Module $settingsCode -Name 'lblTemplateStatus_Click' -Code @'
Private Sub lblTemplateStatus_Click()
    Dim filterText As String
    If mTemplateIncomplete > 0 Then filterText = TemplateIncompleteCriteria()
    OpenReadinessForm "F_TemplateList", "subTemplateList", filterText
End Sub
'@
    $settingsCode = Set-Procedure -Module $settingsCode -Name 'lblWarehouseStatus_Click' -Code @'
Private Sub lblWarehouseStatus_Click()
    Dim filterText As String
    If mWarehouseIncomplete > 0 Then filterText = WarehouseIncompleteCriteria()
    OpenReadinessForm "F_WarehouseList", "subWarehouseList", filterText
End Sub
'@
    $settingsCode = Set-Procedure -Module $settingsCode -Name 'cmdBackupData_Click' -Code @'
Private Sub cmdBackupData_Click()
    UX_BackupAllData
End Sub
'@
    $settingsCode = Set-Procedure -Module $settingsCode -Name 'cmdRestoreData_Click' -Code @'
Private Sub cmdRestoreData_Click()
    UX_RestoreAllData
End Sub
'@
    Set-FormModule 'F_SystemSettings' $settingsCode
    $access.DoCmd.Close(2, 'F_SystemSettings', 1)
    Log 'Updated F_SystemSettings counters and backup'

    # ------------------------------------------------------------------
    # Horizontal scrolling on datasheets
    # ------------------------------------------------------------------
    $wideDatasheetForms = @(
        'SF_WarehouseList','SF_LocationList','SF_UnitList','SF_TemplateList','SF_SupplierList',
        'SF_StockMovementList','SF_SalesOrderList','SF_ProjectList','SF_ItemSetupList',
        'SF_InvoiceList','SF_InventoryList','SF_CustomerPOList','SF_CustomerList'
    )
    foreach ($formName in $wideDatasheetForms) {
        if (-not (FormExists $formName)) { continue }
        $form = Open-FormDesign $formName
        $form.ScrollBars = 3
        $form.FitToScreen = $false
        $form.ViewsAllowed = 2
        try { $form.AllowDatasheetView = $true } catch {}
        $access.DoCmd.Close(2, $formName, 1)
        Log "Scroll $formName"
    }
    $wideHosts = @(
        'F_WarehouseList','F_LocationList','F_UnitList','F_TemplateList','F_SupplierList',
        'F_StockMovementList','F_SalesOrderList','F_ProjectList','F_ItemSetupList',
        'F_InvoiceList','F_InventoryList','F_CustomerPOList','F_CustomerList'
    )
    foreach ($formName in $wideHosts) {
        if (-not (FormExists $formName)) { continue }
        $form = Open-FormDesign $formName
        $form.ScrollBars = 3
        $form.FitToScreen = $false
        foreach ($control in $form.Controls) {
            if ($control.ControlType -eq 112) {
                try { $control.ScrollBars = 3 } catch {}
            }
        }
        $access.DoCmd.Close(2, $formName, 1)
        Log "Host scroll $formName"
    }

    $uxPath = Join-Path $root '_inspect\M_UX_live.txt'
    $access.SaveAsText(5, 'M_UX', $uxPath)
    $uxText = Get-Content -Raw $uxPath
    if ($uxText -notmatch 'Public Sub UX_EnableHorizontalScroll') {
        $uxHelper = @'

Public Sub UX_EnableHorizontalScroll(ByVal frm As Form)
    On Error Resume Next
    Dim ctl As Control
    If frm.DefaultView = 2 Then
        frm.ScrollBars = 3
        frm.FitToScreen = False
    End If
    For Each ctl In frm.Controls
        If ctl.ControlType = acSubform Then
            ctl.Form.ScrollBars = 3
            ctl.Form.FitToScreen = False
        End If
    Next ctl
    On Error GoTo 0
End Sub
'@
        $uxText = $uxText.TrimEnd() + "`r`n" + $uxHelper
        $uxText = $uxText.Replace(
            "Public Sub UX_FitForm(ByVal frm As Form, ByVal availW As Long, ByVal availH As Long)`r`n",
            "Public Sub UX_FitForm(ByVal frm As Form, ByVal availW As Long, ByVal availH As Long)`r`n`r`n    UX_EnableHorizontalScroll frm`r`n    If frm.DefaultView = 2 Then Exit Sub`r`n"
        )
        Set-Content -Path $uxPath -Value $uxText -Encoding Unicode
        $access.LoadFromText(5, 'M_UX', $uxPath)
        Log 'Updated M_UX horizontal scroll helper'
    }

    try {
        $access.DoCmd.RunCommand(126)
        Log ('Compile ran, IsCompiled=' + $access.IsCompiled)
    } catch {
        Log ('Compile warning: ' + $_.Exception.Message)
    }

    $access.CloseCurrentDatabase()
    $access.Quit()
    Log 'DONE'
}
catch {
    Log ('FATAL ' + $_.Exception.Message)
    try { $access.Quit() } catch {}
    throw
}
finally {
    if ($access -ne $null) {
        [void][Runtime.InteropServices.Marshal]::ReleaseComObject($access)
    }
    [GC]::Collect()
    [GC]::WaitForPendingFinalizers()
}
