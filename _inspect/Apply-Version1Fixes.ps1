$ErrorActionPreference = 'Stop'

$databasePath = [System.IO.Path]::GetFullPath(
    (Join-Path $PSScriptRoot '..\LLC_Version1.accdb')
)

if ((Split-Path $databasePath -Leaf) -ne 'LLC_Version1.accdb') {
    throw "Refusing to edit unexpected database: $databasePath"
}
if (-not (Test-Path -LiteralPath $databasePath)) {
    throw "Version 1 database not found: $databasePath"
}

$access = New-Object -ComObject Access.Application
$access.Visible = $true
$access.OpenCurrentDatabase($databasePath)
$db = $access.CurrentDb()

function Get-ModuleText($formName) {
    $module = $access.Forms.Item($formName).Module
    return $module.Lines(1, $module.CountOfLines)
}

function Set-ModuleText($formName, [string]$text) {
    $module = $access.Forms.Item($formName).Module
    if ($module.CountOfLines -gt 0) {
        $module.DeleteLines(1, $module.CountOfLines)
    }
    if (-not [string]::IsNullOrWhiteSpace($text)) {
        $module.InsertLines(1, $text)
    }
}

function Replace-Procedure([string]$text, [string]$signature, [string]$replacement) {
    $escaped = [regex]::Escape($signature)
    $pattern = "(?ms)^$escaped\s.*?^End (Sub|Function)\s*$"
    $matches = [regex]::Matches($text, $pattern)
    if ($matches.Count -ne 1) {
        throw "Expected one procedure matching '$signature'; found $($matches.Count)."
    }
    return $text.Substring(0, $matches[0].Index) +
        $replacement.TrimEnd() +
        $text.Substring($matches[0].Index + $matches[0].Length)
}

function Add-FieldIfMissing([string]$tableName, [string]$fieldName, [string]$sqlType) {
    $table = $db.TableDefs.Item($tableName)
    $found = $false
    foreach ($field in $table.Fields) {
        if ($field.Name -eq $fieldName) { $found = $true; break }
    }
    if (-not $found) {
        $db.Execute(
            "ALTER TABLE [$tableName] ADD COLUMN [$fieldName] $sqlType",
            128
        )
        $db.TableDefs.Refresh()
    }
}

function Ensure-Control($form, [string]$name, [int]$type, [string]$source, [int]$left, [int]$top, [int]$width, [int]$height) {
    foreach ($existing in $form.Controls) {
        if ($existing.Name -eq $name) { return $existing }
    }
    $created = $access.CreateControl(
        $script:currentFormName,
        $type,
        0,
        '',
        $source,
        $left,
        $top,
        $width,
        $height
    )
    $created.Name = $name
    return $created
}

function Ensure-Label($form, [string]$name, [string]$caption, [int]$left, [int]$top, [int]$width, [int]$height) {
    foreach ($existing in $form.Controls) {
        if ($existing.Name -eq $name) { return $existing }
    }
    $created = $access.CreateControl(
        $script:currentFormName,
        100,
        0,
        '',
        '',
        $left,
        $top,
        $width,
        $height
    )
    $created.Name = $name
    $created.Caption = $caption
    return $created
}

function Copy-ControlStyle($source, $target, [string[]]$properties) {
    foreach ($property in $properties) {
        $target.$property = $source.$property
    }
}

function Open-FormDesign([string]$formName) {
    $script:currentFormName = $formName
    $access.DoCmd.OpenForm($formName, 1)
    return ,$access.Forms.Item($formName)
}

try {
    Add-FieldIfMissing 'T_Items' 'ItemStatus' 'TEXT(30)'
    Add-FieldIfMissing 'T_AssemblyTemplates' 'ManufacturerPartNumber' 'TEXT(100)'
    Add-FieldIfMissing 'T_AssemblyTemplates' 'VendorPartNumber' 'TEXT(100)'
    Add-FieldIfMissing 'T_AssemblyTemplates' 'InternalPartNumber' 'TEXT(100)'
    Add-FieldIfMissing 'T_AssemblyTemplates' 'AutoGenerateInternalPartNumber' 'YESNO'

    $db = $access.CurrentDb()
    $itemStatusField = $db.TableDefs.Item('T_Items').Fields.Item('ItemStatus')
    $itemStatusField.DefaultValue = '"Draft"'
    $db.Execute(
        "UPDATE T_Items SET ItemStatus='Complete' " +
        "WHERE ItemStatus Is Null OR Trim(ItemStatus)=''",
        128
    )
    $db.TableDefs.Item('T_AssemblyTemplates').Fields.Item('AutoGenerateInternalPartNumber').DefaultValue = 'False'

    foreach ($fieldName in @('WarehouseID', 'BinNumber', 'ShelfNumber')) {
        $db.TableDefs.Item('T_Locations').Fields.Item($fieldName).Required = $false
    }

    $itemList = $db.QueryDefs.Item('Q_ItemSetupList')
    if ($itemList.SQL -notmatch '\bItemStatus\b') {
        $itemList.SQL = $itemList.SQL.Replace(
            'I.ItemName, I.Description,',
            'I.ItemName, I.Description, I.ItemStatus, I.ItemClass, I.MinStockLevel, I.MaxStockLevel, I.RequiresSerial, I.SerialNumber,'
        )
    }

    $templateList = $db.QueryDefs.Item('Q_TemplateList')
    if ($templateList.SQL -notmatch '\bManufacturerPartNumber\b') {
        $templateList.SQL = $templateList.SQL.Replace(
            'T.CustomerPartNumber, T.OutputQuantity,',
            'T.CustomerPartNumber, T.ManufacturerPartNumber, T.VendorPartNumber, T.InternalPartNumber, T.AutoGenerateInternalPartNumber, T.OutputQuantity,'
        )
    }

    $customerList = $db.QueryDefs.Item('Q_CustomerList')
    if ($customerList.SQL -notmatch '\bMissingFields\b') {
        $missingFields = @"
IIf(Len(Trim(Nz(C.CustomerName,'')))=0,'Customer Name; ','') &
IIf(Len(Trim(Nz(C.ContactPerson,'')))=0,'Contact Name; ','') &
IIf(Len(Trim(Nz(C.Email,'')))=0,'Email; ','') &
IIf(Len(Trim(Nz(C.BillingAddress,'')))=0,'Billing Address; ','') AS MissingFields,
"@
        $customerList.SQL = $customerList.SQL.Replace(
            'C.CustomerID, C.CustomerName,',
            "C.CustomerID, C.CustomerName, C.CustomerType, $missingFields"
        )
    }

    $form = Open-FormDesign 'F_ItemSetup'
    $itemStatus = Ensure-Control $form 'cboItemStatus' 111 'ItemStatus' 9738 2280 3648 312
    $itemStatus.RowSourceType = 'Value List'
    $itemStatus.RowSource = 'Draft;In Review;Pending Approval;Complete'
    $itemStatus.BoundColumn = 1
    $itemStatus.ColumnCount = 1
    $itemStatus.LimitToList = $true
    $itemStatus.DefaultValue = '"Draft"'
    $itemStatus.ControlTipText = 'Draft items do not count as incomplete.'
    $statusLabel = Ensure-Label $form 'ItemStatus_Label' 'Status' 7200 2280 2448 312
    Copy-ControlStyle $form.Controls.Item('ReorderLevel_Label') $statusLabel @('FontName','FontSize','FontWeight','ForeColor','BackColor','BackStyle')
    Copy-ControlStyle $form.Controls.Item('txtReorderLevel') $itemStatus @('FontName','FontSize','ForeColor','BackColor','BackStyle','BorderStyle','SpecialEffect')
    $access.DoCmd.Close(2, 'F_ItemSetup', 1)

    $form = Open-FormDesign 'F_AssemblyTemplate'
    $referenceText = $form.Controls.Item('txtRevisionNo')
    $referenceLabel = $form.Controls.Item('RevisionNo_Label')
    $partControls = @(
        @{ Name='txtManufacturerPartNumber'; Source='ManufacturerPartNumber'; Caption='Manufacturer Part Number'; Top=2400 },
        @{ Name='txtVendorPartNumber'; Source='VendorPartNumber'; Caption='Vendor Part Number'; Top=2820 },
        @{ Name='txtInternalPartNumber'; Source='InternalPartNumber'; Caption='Internal Part Number'; Top=3240 }
    )
    foreach ($part in $partControls) {
        $control = Ensure-Control $form $part.Name 109 $part.Source 10662 $part.Top 2052 312
        $label = Ensure-Label $form ($part.Name + '_Label') $part.Caption 8124 $part.Top 2448 312
        Copy-ControlStyle $referenceText $control @('FontName','FontSize','ForeColor','BackColor','BackStyle','BorderStyle','SpecialEffect')
        Copy-ControlStyle $referenceLabel $label @('FontName','FontSize','FontWeight','ForeColor','BackColor','BackStyle')
    }
    $autoPart = Ensure-Control $form 'chkAutoGenerateInternalPartNumber' 106 'AutoGenerateInternalPartNumber' 8124 3660 300 312
    $autoPart.ControlTipText = 'When selected, only the internal part number is generated.'
    $autoPart.AfterUpdate = '[Event Procedure]'
    $autoPartLabel = Ensure-Label $form 'chkAutoGenerateInternalPartNumber_Label' 'Auto Generate Internal Part Number' 8460 3660 3750 312
    Copy-ControlStyle $referenceLabel $autoPartLabel @('FontName','FontSize','FontWeight','ForeColor','BackColor','BackStyle')
    $form.Controls.Item('txtNotes').Width = 6900
    $form.Controls.Item('Notes_Label').Left = 2400
    $form.Controls.Item('Notes_Label').Width = 3000
    $access.DoCmd.Close(2, 'F_AssemblyTemplate', 1)

    $form = Open-FormDesign 'SF_CustomerList'
    $missingControl = Ensure-Control $form 'txtMissingFields' 109 'MissingFields' 15000 0 3000 300
    $missingControl.ColumnWidth = 3000
    $access.DoCmd.Close(2, 'SF_CustomerList', 1)

    $unitForm = Open-FormDesign 'F_UnitSetup'
    $unitForm.NavigationButtons = $true
    $unitForm.RecordSelectors = $true
    $unitForm.AllowAdditions = $true
    $unitForm.DataEntry = $false
    $access.DoCmd.Close(2, 'F_UnitSetup', 1)

    $settings = Open-FormDesign 'F_SystemSettings'
    foreach ($labelName in @(
        'lblCustomerStatus','lblSupplierStatus','lblCategoryStatus','lblUnitStatus',
        'lblItemStatus','lblTemplateStatus','lblWarehouseStatus','lblLocationStatus'
    )) {
        $label = $settings.Controls.Item($labelName)
        $label.OnClick = '[Event Procedure]'
        $label.ControlTipText = 'Click to review these records.'
    }
    $access.DoCmd.Close(2, 'F_SystemSettings', 1)

    $locationForm = Open-FormDesign 'F_LocationSetup'
    $locationModule = Get-ModuleText 'F_LocationSetup'
    $locationModule = Replace-Procedure $locationModule 'Private Function ValidateLocationRecord() As Boolean' @'
Private Function ValidateLocationRecord() As Boolean

    Dim locationIDValue As Long
    Dim locationNameValue As String
    Dim locationCodeValue As String

    ValidateLocationRecord = False

    locationNameValue = Trim$(Nz(ReadFormValue( _
        "LocationName", _
        Array("txtLocationName", "LocationName")), _
        vbNullString))

    If Len(locationNameValue) = 0 Then
        MsgBox "Enter the location name.", vbExclamation, "Location Name Required"
        FocusControlIfExists "txtLocationName", "LocationName"
        Exit Function
    End If

    locationCodeValue = UCase$(Trim$(Nz(ReadFormValue( _
        "LocationCode", _
        Array("txtLocationCode", "LocationCode")), _
        vbNullString)))

    If Len(locationCodeValue) = 0 Then
        MsgBox "The location code could not be generated.", vbExclamation, "Location Code Required"
        Exit Function
    End If

    locationIDValue = CurrentLocationID()
    If DCount( _
        "*", _
        "T_Locations", _
        "[LocationCode]=" & SqlText(locationCodeValue) & _
        " AND [LocationID]<>" & CStr(locationIDValue)) > 0 Then
        MsgBox "The location code already exists.", vbExclamation, "Duplicate Location Code"
        Exit Function
    End If

    ValidateLocationRecord = True

End Function
'@
    $locationModule = Replace-Procedure $locationModule 'Private Function CurrentLocationID() As Long' @'
Private Function CurrentLocationID() As Long

    If Me.NewRecord Then Exit Function

    If RecordsetHasField("LocationID") Then
        CurrentLocationID = Nz(Me.Recordset.Fields("LocationID").Value, 0)
    End If

End Function
'@
    Set-ModuleText 'F_LocationSetup' $locationModule

    $locationForm.NavigationButtons = $true
    $locationForm.RecordSelectors = $true
    $locationForm.DataEntry = $false
    $access.DoCmd.Close(2, 'F_LocationSetup', 1)

    $warehouseForm = Open-FormDesign 'SF_WarehouseList'
    $warehouseModule = Get-ModuleText 'SF_WarehouseList'
    $oldOpen = @'
    DoCmd.OpenForm _
        detailFormName, _
        acNormal, _
        , _
        "WarehouseID=" & warehouseIDValue
'@
    $newOpen = @'
    If CurrentProject.AllForms(detailFormName).IsLoaded Then
        If Forms(detailFormName).Dirty Then
            Forms(detailFormName).Dirty = False
        End If

        With Forms(detailFormName)
            .Filter = "WarehouseID=" & CStr(warehouseIDValue)
            .FilterOn = True
            .Requery
        End With
    Else
        DoCmd.OpenForm _
            detailFormName, _
            acNormal, _
            , _
            "WarehouseID=" & warehouseIDValue
    End If
'@
    if ($warehouseModule.Contains($oldOpen)) {
        $warehouseModule = $warehouseModule.Replace($oldOpen, $newOpen)
    } elseif ($warehouseModule -notmatch '\.FilterOn\s*=\s*True') {
        throw 'Warehouse detail-open statement did not match the live module.'
    }
    Set-ModuleText 'SF_WarehouseList' $warehouseModule
    $access.DoCmd.Close(2, 'SF_WarehouseList', 1)

    $settings = Open-FormDesign 'F_SystemSettings'
    $settingsModule = Get-ModuleText 'F_SystemSettings'
    $settingsModule = Replace-Procedure $settingsModule 'Private Function CustomerIncompleteCriteria() As String' @'
Private Function CustomerIncompleteCriteria() As String

    CustomerIncompleteCriteria = AnyMissing( _
        BlankTextCriteria("CustomerName"), _
        BlankTextCriteria("ContactPerson"), _
        BlankTextCriteria("Email"), _
        BlankTextCriteria("BillingAddress"))

End Function
'@
    $settingsModule = Replace-Procedure $settingsModule 'Private Function ItemIncompleteCriteria() As String' @'
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
            "(Nz([RequiresSerial],0)<>0 AND " & _
                BlankTextCriteria("SerialNumber") & ")") & ")"

End Function
'@
    $settingsModule = Replace-Procedure $settingsModule 'Private Function TemplateIncompleteCriteria() As String' @'
Private Function TemplateIncompleteCriteria() As String

    TemplateIncompleteCriteria = _
        "(Nz([Status],'Draft')<>'Draft') AND (" & _
        AnyMissing( _
            BlankTextCriteria("TemplateCode"), _
            BlankTextCriteria("TemplateName"), _
            MissingKeyCriteria("CustomerID"), _
            "([OutputQuantity] Is Null OR [OutputQuantity]<=0)", _
            BlankTextCriteria("RevisionNo"), _
            BlankTextCriteria("Status"), _
            "[EstimatedLaborHours] Is Null", _
            "[EstimatedMaterialCost] Is Null", _
            "[EstimatedTotalCost] Is Null") & ")"

End Function
'@
    $settingsModule = Replace-Procedure $settingsModule 'Private Function LocationIncompleteCriteria() As String' @'
Private Function LocationIncompleteCriteria() As String

    LocationIncompleteCriteria = AnyMissing( _
        BlankTextCriteria("LocationCode"), _
        BlankTextCriteria("LocationName"))

End Function
'@
    if ($settingsModule -notmatch 'Private Sub OpenReadinessForm\(') {
        $settingsModule = Replace-Procedure $settingsModule 'Private Sub cmdCustomerSetup_Click()' @'
Private Sub OpenReadinessForm( _
    ByVal formName As String, _
    ByVal subformName As String, _
    ByVal filterText As String _
)

    DoCmd.OpenForm formName

    If Len(subformName) > 0 Then
        With Forms(formName).Controls(subformName).Form
            .Filter = filterText
            .FilterOn = (Len(filterText) > 0)
        End With
    End If

End Sub


Private Sub lblCustomerStatus_Click()
    Dim filterText As String
    If mCustomerIncomplete > 0 Then
        filterText = _
            "UCase(Trim(Nz([CustomerType],'')))<>'INTERNAL' AND " & _
            "Len(Trim(Nz([MissingFields],'')))>0"
    End If
    OpenReadinessForm "F_CustomerList", "subCustomerList", filterText
End Sub


Private Sub lblSupplierStatus_Click()
    Dim filterText As String
    If mSupplierIncomplete > 0 Then filterText = SupplierIncompleteCriteria()
    OpenReadinessForm "F_SupplierList", "subSupplierList", filterText
End Sub


Private Sub lblCategoryStatus_Click()
    DoCmd.OpenForm "F_CategoriesList"
End Sub


Private Sub lblUnitStatus_Click()
    DoCmd.OpenForm "F_UnitSetup"
End Sub


Private Sub lblItemStatus_Click()
    Dim filterText As String
    If mItemIncomplete > 0 Then filterText = ItemIncompleteCriteria()
    OpenReadinessForm "F_ItemSetupList", "subItemSetupList", filterText
End Sub


Private Sub lblTemplateStatus_Click()
    Dim filterText As String
    If mTemplateIncomplete > 0 Then filterText = TemplateIncompleteCriteria()
    OpenReadinessForm "F_TemplateList", "subTemplateList", filterText
End Sub


Private Sub lblWarehouseStatus_Click()
    Dim filterText As String
    If mWarehouseIncomplete > 0 Then filterText = WarehouseIncompleteCriteria()
    OpenReadinessForm "F_WarehouseList", "subWarehouseList", filterText
End Sub


Private Sub lblLocationStatus_Click()
    DoCmd.OpenForm "F_LocationSetup"
End Sub


Private Sub cmdCustomerSetup_Click()
    DoCmd.OpenForm "F_CustomerList"
End Sub
'@
    }
    $settingsModule = Replace-Procedure $settingsModule 'Private Sub lblCustomerStatus_Click()' @'
Private Sub lblCustomerStatus_Click()
    Dim filterText As String
    If mCustomerIncomplete > 0 Then
        filterText = _
            "UCase(Trim(Nz([CustomerType],'')))<>'INTERNAL' AND " & _
            "Len(Trim(Nz([MissingFields],'')))>0"
    End If
    OpenReadinessForm "F_CustomerList", "subCustomerList", filterText
End Sub
'@
    $settingsModule = [regex]::Replace(
        $settingsModule,
        '(?m)^Private Sub cmdCustomerSetup_Click\(\)\s*Private Sub cmdSupplierSetup_Click\(\)',
        "Private Sub cmdCustomerSetup_Click()`r`n    DoCmd.OpenForm `"F_CustomerList`"`r`nEnd Sub`r`n`r`nPrivate Sub cmdSupplierSetup_Click()",
        1
    )
    Set-ModuleText 'F_SystemSettings' $settingsModule
    $access.DoCmd.Close(2, 'F_SystemSettings', 1)

    $templateForm = Open-FormDesign 'F_AssemblyTemplate'
    $templateModule = Get-ModuleText 'F_AssemblyTemplate'
    $templateModule = $templateModule.Replace(
        '.RowSource = "Draft;Active;On Hold;Retired"',
        '.RowSource = "Draft;In Review;Pending Approval;Complete"'
    )
    $templateModule = Replace-Procedure $templateModule 'Private Sub NormalizeTemplateValues()' @'
Private Sub NormalizeTemplateValues()

    Me.txtTemplateName = Trim$(Nz(Me.txtTemplateName, vbNullString))

    If Len(Trim$(Nz(Me.txtTemplateCode, vbNullString))) = 0 Then
        Me.txtTemplateCode = GenerateTemplateCodeFromName(Me.txtTemplateName)
    End If

    Me.txtTemplateCode = UCase$(Trim$(Nz(Me.txtTemplateCode, vbNullString)))
    Me.txtRevisionNo = Trim$(Nz(Me.txtRevisionNo, vbNullString))
    Me.cboStatus = Trim$(Nz(Me.cboStatus, vbNullString))
    Me.txtManufacturerPartNumber = Trim$(Nz(Me.txtManufacturerPartNumber, vbNullString))
    Me.txtVendorPartNumber = Trim$(Nz(Me.txtVendorPartNumber, vbNullString))
    Me.txtInternalPartNumber = Trim$(Nz(Me.txtInternalPartNumber, vbNullString))
    Me.txtNOtes = Trim$(Nz(Me.txtNOtes, vbNullString))

    If IsNull(Me.txtOutputQuantity) Then Me.txtOutputQuantity = 1
    If IsNull(Me.txtEstimatedLaborHours) Then Me.txtEstimatedLaborHours = 0

End Sub
'@
    $templateModule = $templateModule.Replace(
        '    SetSubformAvailability' + "`r`n" + "`r`n" + 'Exit_Handler:',
        '    SetSubformAvailability' + "`r`n" + '    UpdateInternalPartNumberState' + "`r`n" + "`r`n" + 'Exit_Handler:'
    )
    $templateModule = $templateModule.Replace(
        '    SetSubformAvailability' + "`r`n" + "`r`n" + '    On Error GoTo 0',
        '    SetSubformAvailability' + "`r`n" + '    UpdateInternalPartNumberState' + "`r`n" + "`r`n" + '    On Error GoTo 0'
    )
    if ($templateModule -notmatch 'Private Sub chkAutoGenerateInternalPartNumber_AfterUpdate') {
        $templateModule += @'


Private Sub chkAutoGenerateInternalPartNumber_AfterUpdate()
    UpdateInternalPartNumberState
    If Nz(Me.chkAutoGenerateInternalPartNumber, False) Then
        EnsureGeneratedInternalPartNumber
    End If
End Sub


Private Sub UpdateInternalPartNumberState()
    Me.txtInternalPartNumber.Locked = Nz(Me.chkAutoGenerateInternalPartNumber, False)
End Sub


Private Sub GenerateInternalPartNumber()
    Dim candidate As String
    Dim suffix As Long
    Dim templateID As Long

    candidate = "ASM-" & UCase$(Trim$(Nz(Me.txtTemplateCode, vbNullString)))
    If candidate = "ASM-" Then
        candidate = "ASM-TEMPLATE-" & Format$(Now(), "yyyymmddhhnnss")
    End If

    templateID = currentTemplateID()
    suffix = 1
    Do While DCount( _
        "*", _
        "T_AssemblyTemplates", _
        "InternalPartNumber=" & SqlText(candidate) & _
        " AND AssemblyTemplateID<>" & templateID) > 0
        suffix = suffix + 1
        candidate = "ASM-" & UCase$(Trim$(Nz(Me.txtTemplateCode, vbNullString))) & _
            "-" & CStr(suffix)
    Loop

    Me.txtInternalPartNumber = candidate
End Sub


Private Sub EnsureGeneratedInternalPartNumber()
    If Nz(Me.chkAutoGenerateInternalPartNumber, False) And _
       Len(Trim$(Nz(Me.txtInternalPartNumber, vbNullString))) = 0 Then
        GenerateInternalPartNumber
    End If
End Sub
'@
    }
    if ($templateModule -notmatch 'EnsureGeneratedInternalPartNumber') {
        throw 'Template internal part-number helper was not added.'
    }
    $templateModule = [regex]::Replace(
        $templateModule,
        'NormalizeTemplateValues\(\)\s+If Not ValidateTemplateRecord Then',
        "NormalizeTemplateValues()`r`n    EnsureGeneratedInternalPartNumber`r`n`r`n    If Not ValidateTemplateRecord Then",
        2
    )
    Set-ModuleText 'F_AssemblyTemplate' $templateModule
    $access.DoCmd.Close(2, 'F_AssemblyTemplate', 1)

    $datasheetForms = @()
    $allForms = $access.CurrentProject.AllForms
    for ($formIndex = 0; $formIndex -lt $allForms.Count; $formIndex++) {
        $entry = $allForms.Item($formIndex)
        if ($entry.Name -like 'SF_*') {
            $formName = $entry.Name
            $access.DoCmd.OpenForm($formName, 1)
            $candidateForm = $access.Forms.Item($formName)
            if ($candidateForm.DefaultView -eq 2) {
                $candidateForm.ScrollBars = 3
                $datasheetForms += $formName
            }
            $access.DoCmd.Close(2, $formName, 1)
        }
    }

    $access.CloseCurrentDatabase()
    Write-Output "Updated Version 1 database: $databasePath"
    Write-Output ("Datasheet forms with horizontal scrolling: " + ($datasheetForms -join ', '))
}
finally {
    if ($access.CurrentProject.Name) {
        try { $access.CloseCurrentDatabase() } catch {}
    }
    $access.Quit()
    [System.Runtime.InteropServices.Marshal]::ReleaseComObject($db) | Out-Null
    [System.Runtime.InteropServices.Marshal]::ReleaseComObject($access) | Out-Null
}
