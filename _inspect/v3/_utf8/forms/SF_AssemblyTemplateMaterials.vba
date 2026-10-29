Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Option Compare Database
Option Explicit

'=============================================================
' FORM: SF_AssemblyTemplateMaterials
' PURPOSE:
'   Maintains component/material requirements and supports direct
'   Item Setup creation from the component combo or Add button.
'=============================================================

Private Const ITEM_SETUP_FORM As String = "F_ItemSetup"


'=============================================================
' FORM EVENTS
'=============================================================

Private Sub Form_Load()

    On Error GoTo Err_Handler

    ConfigureMaterialControls

Exit_Handler:
    Exit Sub

Err_Handler:
    MsgBox _
        "The assembly-material subform could not be configured." & _
        vbCrLf & vbCrLf & _
        "Error Number: " & Err.Number & vbCrLf & _
        "Details: " & Err.Description, _
        vbExclamation, _
        "Assembly Materials"
    Resume Exit_Handler

End Sub


Private Sub Form_BeforeInsert(Cancel As Integer)

    On Error GoTo Err_Handler

    'The parent header is saved by the main form's subform-control
    'Enter event. Link Master/Child Fields supplies the foreign key.
    If ParentTemplateID() <= 0 Then
        Cancel = True
        Exit Sub
    End If

    If IsNull(Me!SortOrder) Then
        Me!SortOrder = GetNextSortOrder()
    End If

Exit_Handler:
    Exit Sub

Err_Handler:
    Cancel = True
    MsgBox _
        "The material line could not be started." & _
        vbCrLf & vbCrLf & _
        "Error Number: " & Err.Number & vbCrLf & _
        "Details: " & Err.Description, _
        vbExclamation, _
        "Assembly Material"
    Resume Exit_Handler

End Sub


Private Sub Form_BeforeUpdate(Cancel As Integer)

    On Error GoTo Err_Handler

    'Do not save the parent or rewrite AssemblyTemplateID here.
    'Doing that while Access validates this row causes error 2115.
    If Nz(Me!AssemblyTemplateID, 0) <= 0 Then
        MsgBox _
            "The material row is not linked to a saved template." & _
            vbCrLf & vbCrLf & _
            "Set Link Master Fields to txtAssemblyTemplateID and " & _
            "Link Child Fields to AssemblyTemplateID.", _
            vbExclamation, _
            "Template Link Required"
        Cancel = True
        Exit Sub
    End If

    If IsNull(Me.cboComponentItemID) Then
        MsgBox _
            "Select a component, material, consumable, or subassembly.", _
            vbExclamation, _
            "Component Required"
        Me.cboComponentItemID.SetFocus
        Cancel = True
        Exit Sub
    End If

    If Not IsPositiveNumber(Me.txtQuantityPerOutput) Then
        MsgBox _
            "Quantity per output must be a number greater than zero.", _
            vbExclamation, _
            "Invalid Quantity"
        Me.txtQuantityPerOutput.SetFocus
        Cancel = True
        Exit Sub
    End If

    If Not IsNull(Me.txtEstimatedUnitCost) Then
        If Not IsNumeric(Me.txtEstimatedUnitCost) Then
            MsgBox _
                "Estimated unit cost must be zero or greater.", _
                vbExclamation, _
                "Invalid Estimated Cost"
            Me.txtEstimatedUnitCost.SetFocus
            Cancel = True
            Exit Sub
        End If

        If CCur(Me.txtEstimatedUnitCost) < 0 Then
            MsgBox _
                "Estimated unit cost must be zero or greater.", _
                vbExclamation, _
                "Invalid Estimated Cost"
            Me.txtEstimatedUnitCost.SetFocus
            Cancel = True
            Exit Sub
        End If
    End If

Exit_Handler:
    Exit Sub

Err_Handler:
    Cancel = True
    MsgBox _
        "The material line could not be validated." & _
        vbCrLf & vbCrLf & _
        "Error Number: " & Err.Number & vbCrLf & _
        "Details: " & Err.Description, _
        vbExclamation, _
        "Assembly Material"
    Resume Exit_Handler

End Sub


Private Sub Form_Error(DataErr As Integer, Response As Integer)

    If DataErr = 2115 Then
        MsgBox _
            "Access rejected a material field while validating the row." & _
            vbCrLf & vbCrLf & _
            "Confirm the subform link fields, then clear any old " & _
            "Validation Rule or control Before Update macro attached " & _
            "to the material controls.", _
            vbExclamation, _
            "Assembly Material"
        Response = acDataErrContinue
    Else
        Response = acDataErrDisplay
    End If

End Sub


Private Sub Form_AfterUpdate()

    RecalculateParentTemplate

End Sub


Private Sub Form_AfterDelConfirm(Status As Integer)

    RecalculateParentTemplate

End Sub


'=============================================================
' COMPONENT COMBO EVENTS
'=============================================================

Private Sub cboComponentItemID_AfterUpdate()

    On Error GoTo Err_Handler

    If IsNull(Me.cboComponentItemID) Then
        Me.txtEstimatedUnitCost = 0
    Else
        Me.txtEstimatedUnitCost = _
            LatestUnitCost(CLng(Me.cboComponentItemID))
    End If

Exit_Handler:
    Exit Sub

Err_Handler:
    Me.txtEstimatedUnitCost = 0
    MsgBox _
        "The latest unit cost could not be loaded." & _
        vbCrLf & vbCrLf & Err.Description, _
        vbExclamation, _
        "Material Cost"
    Resume Exit_Handler

End Sub


Private Sub cboComponentItemID_NotInList( _
    NewData As String, _
    Response As Integer _
)

    Response = acDataErrContinue

    If CreateTypedTemplateItem( _
        NewData, _
        "material/component") Then

        'Access safely requeries and selects the saved item after this
        'NotInList event has finished.
        Response = acDataErrAdded
    Else
        On Error Resume Next
        Me.cboComponentItemID.Undo
        On Error GoTo 0
    End If

End Sub


Private Sub cboComponentItemID_DblClick(Cancel As Integer)

    Cancel = True

    If Nz(Me.cboComponentItemID, 0) > 0 Then
        EditSelectedTemplateItem Me.cboComponentItemID
    Else
        AddNewTemplateItem Me.cboComponentItemID
    End If

End Sub


'=============================================================
' ADD BUTTON
'=============================================================

Private Sub cmdAddComponentItem_Click()

    AddNewComponentFromTemplate

End Sub


Public Sub AddNewComponentFromTemplate()

    AddNewTemplateItem Me.cboComponentItemID

End Sub


'=============================================================
' CONFIGURATION
'=============================================================

Private Sub ConfigureMaterialControls()

    With Me.cboComponentItemID
        .RowSourceType = "Table/Query"
        .RowSource = _
            "SELECT ItemID, ItemName, InternalSKU, ItemClass " & _
            "FROM T_Items " & _
            "WHERE Nz([ItemClass],'') Not In ('Tool','Equipment') " & _
            "ORDER BY ItemName, InternalSKU;"
        .BoundColumn = 1
        .ColumnCount = 4
        .ColumnWidths = "0cm;7cm;3cm;3cm"
        .LimitToList = True
        .AutoExpand = True
    End With

    Me.txtEstimatedUnitCost.Locked = True
    Me.txtEstimatedUnitCost.TabStop = False

    If ControlExists("cmdAddComponentItem") Then
        Me.Controls("cmdAddComponentItem").Caption = _
            "+ New Material / Component"
        Me.Controls("cmdAddComponentItem").Enabled = True
    End If

    If ControlExists("txtAssemblyTemplateID") Then
        Me.Controls("txtAssemblyTemplateID").Visible = False
        Me.Controls("txtAssemblyTemplateID").Enabled = False
        Me.Controls("txtAssemblyTemplateID").Locked = True
    End If

End Sub


'=============================================================
' ITEM SETUP QUICK ADD / EDIT
'=============================================================

Private Function CreateTypedTemplateItem( _
    ByVal NewData As String, _
    ByVal ItemDescription As String _
) As Boolean

    On Error GoTo Err_Handler

    Dim enteredName As String
    Dim savedItemID As Long
    Dim oldMaximumID As Long
    Dim newMaximumID As Long
    Dim savedItemName As String

    CreateTypedTemplateItem = False
    enteredName = Trim$(Nz(NewData, vbNullString))

    If Len(enteredName) = 0 Then
        Exit Function
    End If

    If MsgBox( _
        "'" & enteredName & "' is not in the item list." & _
        vbCrLf & vbCrLf & _
        "Would you like to create it as a new " & _
        ItemDescription & "?" & _
        vbCrLf & vbCrLf & _
        "In Item Setup, select a material-side Item Class " & _
        "(not Tool or Equipment).", _
        vbQuestion + vbYesNo, _
        "Create New Item") <> vbYes Then

        Exit Function
    End If

    If Not PrepareItemSetupForm Then
        Exit Function
    End If

    oldMaximumID = Nz(DMax("ItemID", "T_Items"), 0)

    DoCmd.OpenForm _
        formName:=ITEM_SETUP_FORM, _
        View:=acNormal, _
        DataMode:=acFormAdd, _
        WindowMode:=acDialog, _
        OpenArgs:=BuildItemOpenArgs(enteredName)

    newMaximumID = Nz(DMax("ItemID", "T_Items"), 0)

    If newMaximumID <= oldMaximumID Then
        MsgBox _
            "The new item was not found after Item Setup closed." & _
            vbCrLf & _
            "Save it using the same item name that you typed.", _
            vbInformation, _
            "Item Not Added"
        Exit Function
    End If

    savedItemID = newMaximumID
    savedItemName = Trim$(Nz(DLookup( _
        "ItemName", _
        "T_Items", _
        "ItemID=" & CStr(savedItemID)), vbNullString))

    If StrComp(savedItemName, enteredName, vbTextCompare) <> 0 Then
        MsgBox _
            "The item was saved with a different name." & _
            vbCrLf & _
            "To fill the typed combo value automatically, save it " & _
            "using the same name: " & enteredName, _
            vbInformation, _
            "Item Name Changed"
        Exit Function
    End If

    If Not ItemIsVisibleForThisSubform(savedItemID) Then
        MsgBox _
            "The item was saved, but its Item Class is Tool or " & _
            "Equipment, so it cannot appear in Materials." & _
            vbCrLf & vbCrLf & _
            "Open Item Setup and change its Item Class to Component, " & _
            "Raw Material, Consumable, Other, or Finished Assembly.", _
            vbExclamation, _
            "Material Item Class"
        Exit Function
    End If

    CreateTypedTemplateItem = True

Exit_Handler:
    Exit Function

Err_Handler:
    If Err.Number <> 2501 Then
        MsgBox _
            "The new material/component could not be created." & _
            vbCrLf & vbCrLf & _
            "Error Number: " & Err.Number & vbCrLf & _
            "Details: " & Err.Description, _
            vbExclamation, _
            "Create Material / Component"
    End If

    Resume Exit_Handler

End Function


Private Sub AddNewTemplateItem( _
    ByVal targetCombo As Access.ComboBox _
)

    On Error GoTo Err_Handler

    Dim oldMaximumID As Long
    Dim newMaximumID As Long

    If Not PrepareItemSetupForm Then
        Exit Sub
    End If

    oldMaximumID = Nz(DMax("ItemID", "T_Items"), 0)

    DoCmd.OpenForm _
        formName:=ITEM_SETUP_FORM, _
        View:=acNormal, _
        DataMode:=acFormAdd, _
        WindowMode:=acDialog, _
        OpenArgs:=BuildItemOpenArgs(vbNullString)

    newMaximumID = Nz(DMax("ItemID", "T_Items"), 0)

    targetCombo.Requery

    If newMaximumID > oldMaximumID Then
        If ItemIsVisibleForThisSubform(newMaximumID) Then
            targetCombo.value = newMaximumID
            Me.txtEstimatedUnitCost = LatestUnitCost(newMaximumID)
        Else
            MsgBox _
                "The item was saved, but its Item Class does not " & _
                "belong in the Materials list.", _
                vbExclamation, _
                "Material Item Class"
        End If
    End If

Exit_Handler:
    Exit Sub

Err_Handler:
    If Err.Number <> 2501 Then
        MsgBox _
            "Item Setup could not create the new material/component." & _
            vbCrLf & vbCrLf & _
            "Error Number: " & Err.Number & vbCrLf & _
            "Details: " & Err.Description, _
            vbExclamation, _
            "New Material / Component"
    End If

    Resume Exit_Handler

End Sub


Private Sub EditSelectedTemplateItem( _
    ByVal targetCombo As Access.ComboBox _
)

    On Error GoTo Err_Handler

    Dim selectedItemID As Long

    selectedItemID = Nz(targetCombo.value, 0)

    If selectedItemID <= 0 Then
        AddNewTemplateItem targetCombo
        Exit Sub
    End If

    If Not PrepareItemSetupForm Then
        Exit Sub
    End If

    DoCmd.OpenForm _
        formName:=ITEM_SETUP_FORM, _
        View:=acNormal, _
        WhereCondition:="[ItemID]=" & CStr(selectedItemID), _
        DataMode:=acFormEdit, _
        WindowMode:=acDialog

    targetCombo.Requery

    If DCount( _
        "*", _
        "T_Items", _
        "ItemID=" & CStr(selectedItemID)) = 0 Then

        targetCombo.value = Null
    ElseIf ItemIsVisibleForThisSubform(selectedItemID) Then
        targetCombo.value = selectedItemID
        Me.txtEstimatedUnitCost = LatestUnitCost(selectedItemID)
    Else
        targetCombo.value = Null
        MsgBox _
            "The edited item's class no longer belongs in Materials, " & _
            "so it was cleared from this line.", _
            vbInformation, _
            "Material Item Class Changed"
    End If

Exit_Handler:
    Exit Sub

Err_Handler:
    If Err.Number <> 2501 Then
        MsgBox _
            "The selected item could not be opened." & _
            vbCrLf & vbCrLf & _
            "Error Number: " & Err.Number & vbCrLf & _
            "Details: " & Err.Description, _
            vbExclamation, _
            "Edit Material / Component"
    End If

    Resume Exit_Handler

End Sub


Private Function PrepareItemSetupForm() As Boolean

    On Error GoTo Err_Handler

    PrepareItemSetupForm = False

    If Not FormExists(ITEM_SETUP_FORM) Then
        MsgBox _
            "The form '" & ITEM_SETUP_FORM & "' does not exist.", _
            vbExclamation, _
            "Item Setup Missing"
        Exit Function
    End If

    If IsFormLoaded(ITEM_SETUP_FORM) Then
        If Forms(ITEM_SETUP_FORM).Dirty Then
            MsgBox _
                "Finish or cancel the current record in Item Setup, " & _
                "then try again.", _
                vbInformation, _
                "Item Setup Already Open"
            Exit Function
        End If

        DoCmd.close acForm, ITEM_SETUP_FORM, acSaveNo
    End If

    PrepareItemSetupForm = True

Exit_Handler:
    Exit Function

Err_Handler:
    MsgBox _
        "Item Setup could not be prepared." & _
        vbCrLf & vbCrLf & Err.Description, _
        vbExclamation, _
        "Item Setup"
    Resume Exit_Handler

End Function


Private Function BuildItemOpenArgs( _
    ByVal requestedItemName As String _
) As String

    'Use the existing Quick Add contract already supported by
    'F_ItemSetup. No change to F_ItemSetup is required.
    BuildItemOpenArgs = _
        "QUICKADD|" & requestedItemName

End Function


Private Function ItemIsVisibleForThisSubform( _
    ByVal ItemID As Long _
) As Boolean

    Dim itemClass As String

    If DCount( _
        "*", _
        "T_Items", _
        "ItemID=" & CStr(ItemID)) = 0 Then

        Exit Function
    End If

    itemClass = Trim$(Nz(DLookup( _
        "ItemClass", _
        "T_Items", _
        "ItemID=" & CStr(ItemID)), vbNullString))

    ItemIsVisibleForThisSubform = _
        (StrComp(itemClass, "Tool", vbTextCompare) <> 0) And _
        (StrComp(itemClass, "Equipment", vbTextCompare) <> 0)

End Function


'=============================================================
' PARENT AND SORT HELPERS
'=============================================================

Private Function ParentTemplateID() As Long

    On Error GoTo Err_Handler

    If Me.Parent Is Nothing Then
        Err.Raise _
            vbObjectError + 5301, _
            "ParentTemplateID", _
            "This subform must be opened inside F_AssemblyTemplate."
    End If

    ParentTemplateID = Nz( _
        Me.Parent.Recordset.Fields("AssemblyTemplateID").value, _
        0)

    If ParentTemplateID <= 0 Then
        Err.Raise _
            vbObjectError + 5302, _
            "ParentTemplateID", _
            "Save the template header before adding materials."
    End If

Exit_Handler:
    Exit Function

Err_Handler:
    ParentTemplateID = 0
    Err.Raise Err.Number, Err.Source, Err.Description

End Function


Private Function GetNextSortOrder() As Long

    Dim templateID As Long
    Dim currentMaximum As Variant

    templateID = ParentTemplateID()

    currentMaximum = DMax( _
        "SortOrder", _
        "T_AssemblyTemplateMaterials", _
        "AssemblyTemplateID=" & CStr(templateID))

    If IsNull(currentMaximum) Then
        GetNextSortOrder = 10
    Else
        GetNextSortOrder = CLng(currentMaximum) + 10
    End If

End Function


Private Sub RecalculateParentTemplate()

    On Error Resume Next

    Me.Parent.RecalculateTemplateCosts

    On Error GoTo 0

End Sub


'=============================================================
' COST HELPER
'=============================================================

Private Function LatestUnitCost(ByVal ItemID As Long) As Currency

    On Error GoTo Err_Handler

    Dim rows As DAO.Recordset

    LatestUnitCost = 0

    Set rows = CurrentDb.OpenRecordset( _
        "SELECT TOP 1 UnitCost " & _
        "FROM T_StockMovements " & _
        "WHERE ItemID=" & CStr(ItemID) & " " & _
        "AND UnitCost Is Not Null " & _
        "ORDER BY MovementDate DESC, StockMovementID DESC;", _
        dbOpenSnapshot)

    If Not rows.EOF Then
        LatestUnitCost = Nz(rows!UnitCost, 0)
    End If

Exit_Handler:
    On Error Resume Next
    If Not rows Is Nothing Then rows.close
    Set rows = Nothing
    On Error GoTo 0
    Exit Function

Err_Handler:
    LatestUnitCost = 0
    Resume Exit_Handler

End Function


'=============================================================
' SAFE HELPERS
'=============================================================

Private Function ControlExists( _
    ByVal controlName As String _
) As Boolean

    On Error GoTo Not_Found

    Dim controlObject As Control

    Set controlObject = Me.Controls(controlName)
    ControlExists = True

Exit_Function:
    Set controlObject = Nothing
    Exit Function

Not_Found:
    ControlExists = False
    Resume Exit_Function

End Function


Private Function FormExists(ByVal formName As String) As Boolean

    On Error GoTo Not_Found

    Dim formObject As AccessObject

    Set formObject = CurrentProject.AllForms(formName)
    FormExists = True

Exit_Function:
    Set formObject = Nothing
    Exit Function

Not_Found:
    FormExists = False
    Resume Exit_Function

End Function


Private Function IsFormLoaded(ByVal formName As String) As Boolean

    On Error GoTo Not_Loaded

    IsFormLoaded = CurrentProject.AllForms(formName).IsLoaded
    Exit Function

Not_Loaded:
    IsFormLoaded = False

End Function


Private Function IsPositiveNumber( _
    ByVal fieldValue As Variant _
) As Boolean

    If IsNull(fieldValue) Then
        Exit Function
    End If

    If Len(Trim$(CStr(fieldValue))) = 0 Then
        Exit Function
    End If

    If Not IsNumeric(fieldValue) Then
        Exit Function
    End If

    IsPositiveNumber = (CDbl(fieldValue) > 0)

End Function
