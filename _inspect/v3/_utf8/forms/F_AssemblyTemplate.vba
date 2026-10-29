Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Compare Database
Option Explicit

'=============================================================
' FORM: F_AssemblyTemplate
' PURPOSE:
'   Maintains an assembly template header and its material/tool
'   requirements. A template defines requirements only; it does
'   not create stock movements.
'=============================================================

Private Const PROJECT_FORM As String = "F_Projects"
Private Const NOTES_REPORT As String = "R_TemplateNotes"
Private Const PDF_MAIN_FOLDER As String = "\Documents\Inventory System PDFs"
Private Const PDF_TEMPLATE_FOLDER As String = "\Documents\Inventory System PDFs\Template Notes"

Private mProjectContext As Boolean
Private mContextProjectID As Long
Private mLinkInProgress As Boolean


Private Sub cmdManageWorkInstructions_Click()

    DoCmd.OpenForm "F_WorkInstructions", _
        View:=acNormal, _
        DataMode:=acFormEdit, _
        WindowMode:=acDialog


End Sub

'=============================================================
' FORM EVENTS
'=============================================================

Private Sub Form_Load()

    On Error GoTo Err_Handler

    ConfigureFormControls
    ConfigureComboBoxes
    ApplyTemplateDefaults
    ApplyQuickAddOpenArgs
    ApplyProjectTemplateOpenArgs
    ConfigureContextControls
    SetSubformAvailability

Exit_Handler:
    Exit Sub

Err_Handler:
    MsgBox _
        "The Template / Estimate form could not be loaded." & vbCrLf & vbCrLf & _
        "Error Number: " & Err.Number & vbCrLf & _
        "Details: " & Err.Description, _
        vbExclamation, _
        "Template / Estimate"
    Resume Exit_Handler

End Sub


Private Sub Form_Current()

    On Error Resume Next

    ApplyTemplateDefaults
    ConfigureContextControls
    SetSubformAvailability

    On Error GoTo 0

End Sub


Private Sub Form_BeforeInsert(Cancel As Integer)

    ApplyTemplateDefaults

End Sub


Private Sub Form_BeforeUpdate(Cancel As Integer)

    On Error GoTo Err_Handler

    NormalizeTemplateValues

    If Not ValidateTemplateRecord Then
        Cancel = True
        Exit Sub
    End If

    If Me.NewRecord Then
        SetRecordField "CreatedOn", Now()
    End If

    SetRecordField "UpdatedOn", Now()

Exit_Handler:
    Exit Sub

Err_Handler:
    Cancel = True
    MsgBox _
        "The template could not be validated." & vbCrLf & vbCrLf & _
        "Error Number: " & Err.Number & vbCrLf & _
        "Details: " & Err.Description, _
        vbExclamation, _
        "Template Validation"
    Resume Exit_Handler

End Sub


Private Sub Form_AfterInsert()

    On Error Resume Next

    SetSubformAvailability
    LinkTemplateBackToProjectIfNeeded

    On Error GoTo 0

End Sub


Private Sub Form_AfterUpdate()

    On Error Resume Next

    SetSubformAvailability
    LinkTemplateBackToProjectIfNeeded

    On Error GoTo 0

End Sub

'=============================================================
' FIELD EVENTS
'=============================================================

Private Sub txtTemplateName_AfterUpdate()

    If Len(Trim$(Nz(Me.txtTemplateCode, vbNullString))) = 0 Then
        Me.txtTemplateCode = GenerateTemplateCodeFromName(Me.txtTemplateName)
    End If

    SetSubformAvailability

End Sub


Private Sub txtOutputQuantity_BeforeUpdate(Cancel As Integer)

    If IsNull(Me.txtOutputQuantity) Or _
       Not IsNumeric(Me.txtOutputQuantity) Or _
       CDbl(Me.txtOutputQuantity) <= 0 Then

        MsgBox _
            "Output quantity must be a number greater than zero.", _
            vbExclamation, _
            "Invalid Output Quantity"
        Cancel = True
    End If

End Sub


Private Sub cboFinishedItemID_AfterUpdate()

    On Error Resume Next

    Dim finishedItemCombo As ComboBox

    If ControlExists("cboFinishedItemID") Then
        Set finishedItemCombo = Me.Controls("cboFinishedItemID")

        If Len(Trim$(Nz(Me.txtTemplateName, vbNullString))) = 0 Then
            Me.txtTemplateName = Nz(finishedItemCombo.Column(1), vbNullString)
        End If
    End If

    Set finishedItemCombo = Nothing
    On Error GoTo 0

End Sub

'=============================================================
' PRIMARY BUTTONS
'=============================================================

Private Sub cmdNewTemplate_Click()

    On Error GoTo Err_Handler

    Dim reuseCurrentNewRecord As Boolean

    If mProjectContext Then
        MsgBox _
            "This template was opened from a project." & vbCrLf & vbCrLf & _
            "Finish and close it before creating another template.", _
            vbInformation, _
            "Project Template"
        Exit Sub
    End If

    If Me.NewRecord And IsBlankNewTemplateRecord Then
        reuseCurrentNewRecord = True

        If Me.Dirty Then
            Me.Undo
        End If

    ElseIf Me.Dirty Or _
       SubformIsDirty("SF_AssemblyTemplateMaterials") Or _
       SubformIsDirty("SF_AssemblyTemplateTools") Then

        If Not SaveTemplateRecord(False) Then
            Exit Sub
        End If
    End If

    If Not reuseCurrentNewRecord Then
        DoCmd.GoToRecord acDataForm, Me.Name, acNewRec
    End If

    ApplyTemplateDefaults
    ConfigureContextControls
    SetSubformAvailability
Exit_Handler:
    Exit Sub

Err_Handler:
    MsgBox _
        "A new template could not be opened." & vbCrLf & vbCrLf & _
        "Error Number: " & Err.Number & vbCrLf & _
        "Details: " & Err.Description, _
        vbExclamation, _
        "New Template"
    Resume Exit_Handler

End Sub


Private Sub cmdSaveTemplate_Click()

    If SaveTemplateRecord(True) Then
        MsgBox _
            "Template / Estimate saved successfully.", _
            vbInformation, _
            "Template Saved"
    End If

End Sub


Private Sub cmdRecalculateCosts_Click()

    On Error GoTo Err_Handler

    If Not SaveTemplateRecord(False) Then
        Exit Sub
    End If

    MsgBox _
        "Template material cost recalculated.", _
        vbInformation, _
        "Cost Recalculation"

Exit_Handler:
    Exit Sub

Err_Handler:
    MsgBox _
        "Template costs could not be recalculated." & vbCrLf & vbCrLf & _
        Err.Description, _
        vbExclamation, _
        "Cost Recalculation"
    Resume Exit_Handler

End Sub


Private Sub cmdPrintNotesPDF_Click()

    CreateTemplateNotesPDF

End Sub


Private Sub cmdClose_Click()

    On Error GoTo Err_Handler

    Dim Response As VbMsgBoxResult
    Dim hasChanges As Boolean

    hasChanges = _
        Me.Dirty Or _
        SubformIsDirty("SF_AssemblyTemplateMaterials") Or _
        SubformIsDirty("SF_AssemblyTemplateTools")

    If hasChanges Then

        Response = MsgBox( _
            "Save the template changes before closing?", _
            vbQuestion + vbYesNoCancel, _
            "Close Template / Estimate")

        Select Case Response

            Case vbYes
                If Not SaveTemplateRecord(True) Then
                    Exit Sub
                End If

            Case vbNo
                UndoSubformIfDirty "SF_AssemblyTemplateMaterials"
                UndoSubformIfDirty "SF_AssemblyTemplateTools"

                If Me.Dirty Then
                    Me.Undo
                End If

            Case vbCancel
                Exit Sub

        End Select

    End If

    DoCmd.close acForm, Me.Name

Exit_Handler:
    Exit Sub

Err_Handler:
    MsgBox _
        "The Template / Estimate form could not be closed." & vbCrLf & vbCrLf & _
        "Error Number: " & Err.Number & vbCrLf & _
        "Details: " & Err.Description, _
        vbExclamation, _
        "Close Template"
    Resume Exit_Handler

End Sub

'=============================================================
' CUSTOMER / ITEM BUTTONS AND NOT-IN-LIST EVENTS
'=============================================================

Private Sub cmdNewCustomer_Click()

    On Error GoTo Err_Handler

    AddComboEntry _
        Me.cboCustomerID, _
        "F_CustomerSetup", _
        "T_Customers", _
        "CustomerID"

Exit_Handler:
    Exit Sub

Err_Handler:
    MsgBox "Customer Setup could not be opened." & vbCrLf & Err.Description, vbExclamation
    Resume Exit_Handler

End Sub


Private Sub cmdNewItem_Click()

    On Error GoTo Err_Handler

    Dim finishedItemCombo As Access.ComboBox

    If Not ControlExists("cboFinishedItemID") Then
        MsgBox _
            "Add the cboFinishedItemID combo to this form before creating " & _
            "standalone templates.", _
            vbExclamation, _
            "Finished Item Control Missing"
        Exit Sub
    End If

    Set finishedItemCombo = Me.Controls("cboFinishedItemID")

    AddComboEntry _
        finishedItemCombo, _
        "F_ItemSetup", _
        "T_Items", _
        "ItemID"

Exit_Handler:
    Set finishedItemCombo = Nothing
    Exit Sub

Err_Handler:
    MsgBox "Item Setup could not be opened." & vbCrLf & Err.Description, vbExclamation
    Resume Exit_Handler

End Sub


Private Sub cboCustomerID_NotInList(NewData As String, Response As Integer)

    Response = acDataErrContinue

    If HandleComboNotInList( _
        Me.cboCustomerID, _
        NewData, _
        "F_CustomerSetup", _
        "T_Customers", _
        "CustomerID", _
        "CustomerName") Then

        Response = acDataErrAdded
    Else
        Me.cboCustomerID.Undo
    End If

End Sub


Private Sub cboFinishedItemID_NotInList(NewData As String, Response As Integer)

    Response = acDataErrContinue

    Dim finishedItemCombo As Access.ComboBox

    If Not ControlExists("cboFinishedItemID") Then
        Exit Sub
    End If

    Set finishedItemCombo = Me.Controls("cboFinishedItemID")

    If HandleComboNotInList( _
        finishedItemCombo, _
        NewData, _
        "F_ItemSetup", _
        "T_Items", _
        "ItemID", _
        "ItemName") Then

        Response = acDataErrAdded
    Else
        finishedItemCombo.Undo
    End If

    Set finishedItemCombo = Nothing

End Sub

'=============================================================
' CONFIGURATION
'=============================================================

Private Sub ConfigureFormControls()

    Me.txtAssemblyTemplateID.Visible = False
    Me.txtAssemblyTemplateID.Enabled = False
    Me.txtAssemblyTemplateID.Locked = True

    Me.txtTemplateCode.Locked = True
    Me.txtEstimatedMaterialCost.Locked = True
    Me.txtEstimatedTotalCost.Locked = True

    Me.txtNOtes.EnterKeyBehavior = True
    Me.txtNOtes.ScrollBars = 2

End Sub


Private Sub ConfigureComboBoxes()

    Dim finishedItemCombo As ComboBox

    With Me.cboCustomerID
        .RowSourceType = "Table/Query"
        .RowSource = _
            "SELECT CustomerID, CustomerName " & _
            "FROM T_Customers " & _
            "ORDER BY CustomerName;"
        .BoundColumn = 1
        .ColumnCount = 2
        .ColumnWidths = "0cm;8cm"
        .LimitToList = True
        .AutoExpand = True
    End With

    If ControlExists("cboFinishedItemID") Then
        Set finishedItemCombo = Me.Controls("cboFinishedItemID")

        With finishedItemCombo
            .RowSourceType = "Table/Query"
            .RowSource = _
                "SELECT ItemID, ItemName, InternalSKU, ItemClass " & _
                "FROM T_Items " & _
                "ORDER BY ItemName, InternalSKU;"
            .BoundColumn = 1
            .ColumnCount = 4
            .ColumnWidths = "0cm;7cm;3cm;3cm"
            .LimitToList = True
            .AutoExpand = True
        End With
    End If

    Set finishedItemCombo = Nothing

    With Me.cboStatus
        .RowSourceType = "Value List"
        .RowSource = "Draft;Active;On Hold;Retired"
        .BoundColumn = 1
        .ColumnCount = 1
        .LimitToList = True
        .AutoExpand = True
    End With

End Sub


Private Sub ConfigureContextControls()

    On Error Resume Next

    Dim finishedItemCombo As ComboBox

    If ControlExists("cboFinishedItemID") Then
        Set finishedItemCombo = Me.Controls("cboFinishedItemID")
        finishedItemCombo.Locked = mProjectContext
        finishedItemCombo.Enabled = True
    End If

    Me.cmdNewTemplate.Enabled = Not mProjectContext

    Set finishedItemCombo = Nothing
    On Error GoTo 0

End Sub


Private Sub ApplyTemplateDefaults()

    If Not Me.NewRecord Then
        Exit Sub
    End If

    If IsNull(Me.txtOutputQuantity) Then
        Me.txtOutputQuantity = 1
    End If

    If Len(Trim$(Nz(Me.cboStatus, vbNullString))) = 0 Then
        Me.cboStatus = "Draft"
    End If

    If IsNull(Me.txtEstimatedLaborHours) Then
        Me.txtEstimatedLaborHours = 0
    End If

    If IsNull(Me.txtEstimatedMaterialCost) Then
        Me.txtEstimatedMaterialCost = 0
    End If

    If IsNull(Me.txtEstimatedTotalCost) Then
        Me.txtEstimatedTotalCost = 0
    End If

End Sub

'=============================================================
' PROJECT OPENARGS AND LINK-BACK
'=============================================================

Private Sub ApplyQuickAddOpenArgs()

    On Error GoTo Err_Handler

    If Not Me.NewRecord Then
        Exit Sub
    End If

    If Left$(Nz(Me.OpenArgs, vbNullString), 9) <> "QUICKADD|" Then
        Exit Sub
    End If

    SeedQuickAddField Me, "txtTemplateName"

    If Len(Trim$(Nz(Me.txtTemplateName, vbNullString))) > 0 Then
        Me.txtTemplateCode = _
            GenerateTemplateCodeFromName(Me.txtTemplateName)
    End If

    Me.txtTemplateName.SetFocus

Exit_Handler:
    Exit Sub

Err_Handler:
    MsgBox _
        "The typed template name could not be prefilled." & vbCrLf & vbCrLf & _
        "Error Number: " & Err.Number & vbCrLf & _
        "Details: " & Err.Description, _
        vbExclamation, _
        "Template Quick Add"
    Resume Exit_Handler

End Sub


Private Sub ApplyProjectTemplateOpenArgs()

    On Error GoTo Err_Handler

    Dim contextText As String
    Dim modeText As String
    Dim requestedTemplateName As String
    Dim ProjectCode As String
    Dim ProjectName As String
    Dim WorkInstructions As String
    Dim CustomerID As Long
    Dim FinishedItemID As Long
    Dim rs As DAO.Recordset

    If Len(Trim$(Nz(Me.OpenArgs, vbNullString))) = 0 Then
        Exit Sub
    End If

    contextText = GetOpenArgValue("Context")
    modeText = GetOpenArgValue("Mode")

    If StrComp(contextText, "Project", vbTextCompare) <> 0 Then
        Exit Sub
    End If

    If StrComp(modeText, "NewTemplate", vbTextCompare) <> 0 Then
        Exit Sub
    End If

    mContextProjectID = Val(GetOpenArgValue("ProjectID"))

    If mContextProjectID <= 0 Then
        Err.Raise _
            vbObjectError + 5201, _
            "ApplyProjectTemplateOpenArgs", _
            "The project ID was not supplied."
    End If

    Set rs = CurrentDb.OpenRecordset( _
        "SELECT ProjectCode, ProjectName, CustomerID, FinishedItemID, " & _
        "WorkInstructions " & _
        "FROM T_Projects " & _
        "WHERE ProjectID=" & mContextProjectID, _
        dbOpenSnapshot)

    If rs.EOF Then
        Err.Raise _
            vbObjectError + 5202, _
            "ApplyProjectTemplateOpenArgs", _
            "The selected project no longer exists."
    End If

    ProjectCode = Nz(rs!ProjectCode, vbNullString)
    ProjectName = Nz(rs!ProjectName, vbNullString)
    CustomerID = Nz(rs!CustomerID, 0)
    FinishedItemID = Nz(rs!FinishedItemID, 0)
    WorkInstructions = Nz(rs!WorkInstructions, vbNullString)

    rs.close
    Set rs = Nothing

    If FinishedItemID <= 0 Then
        Err.Raise _
            vbObjectError + 5203, _
            "ApplyProjectTemplateOpenArgs", _
            "Select and save the project's finished item first."
    End If

    mProjectContext = True

    If Not Me.NewRecord Then
        DoCmd.GoToRecord acDataForm, Me.Name, acNewRec
    End If

    requestedTemplateName = GetOpenArgValue("RequestedTemplateName")

    If Len(Trim$(requestedTemplateName)) > 0 Then
        Me.txtTemplateName = requestedTemplateName
    ElseIf Len(Trim$(ProjectName)) > 0 Then
        Me.txtTemplateName = ProjectName & " Template"
    Else
        Me.txtTemplateName = "New Project Template"
    End If

    Me.txtTemplateCode = MakeTemplateCodeFromProject(ProjectCode)

    Me!FinishedItemID = FinishedItemID

    If ControlExists("cboFinishedItemID") Then
        Me.Controls("cboFinishedItemID").value = FinishedItemID
    End If

    If CustomerID > 0 Then
        Me.cboCustomerID = CustomerID
    End If

    'A template is stored per one output unit. Project.BuildQuantity
    'controls how many units are required for a particular project.
    Me.txtOutputQuantity = 1
    Me.cboStatus = "Draft"

    If Len(Trim$(WorkInstructions)) > 0 Then
        Me.txtNOtes = WorkInstructions
    End If

Exit_Handler:
    On Error Resume Next
    If Not rs Is Nothing Then
        rs.close
    End If
    Set rs = Nothing
    On Error GoTo 0
    Exit Sub

Err_Handler:
    MsgBox _
        "The new template could not be prefilled from the project." & vbCrLf & vbCrLf & _
        "Error Number: " & Err.Number & vbCrLf & _
        "Details: " & Err.Description, _
        vbExclamation, _
        "Template Prefill"
    Resume Exit_Handler

End Sub


Private Function LinkTemplateBackToProjectIfNeeded() As Boolean

    On Error GoTo Err_Handler

    Dim templateID As Long
    Dim projectFinishedItemID As Variant
    Dim templateFinishedItemID As Long
    Dim projectFormObject As Form

    LinkTemplateBackToProjectIfNeeded = True

    If Not mProjectContext Then
        Exit Function
    End If

    If mLinkInProgress Then
        Exit Function
    End If

    templateID = currentTemplateID()

    If mContextProjectID <= 0 Or templateID <= 0 Then
        LinkTemplateBackToProjectIfNeeded = False
        Exit Function
    End If

    projectFinishedItemID = DLookup( _
        "FinishedItemID", _
        "T_Projects", _
        "ProjectID=" & mContextProjectID)

    templateFinishedItemID = Nz(Me!FinishedItemID, 0)

    If IsNull(projectFinishedItemID) Or _
       CLng(projectFinishedItemID) <> templateFinishedItemID Then

        Err.Raise _
            vbObjectError + 5210, _
            "LinkTemplateBackToProjectIfNeeded", _
            "The template finished item does not match the project finished item."
    End If

    mLinkInProgress = True

    CurrentDb.Execute _
        "UPDATE T_Projects " & _
        "SET AssemblyTemplateID=" & templateID & " " & _
        "WHERE ProjectID=" & mContextProjectID, _
        dbFailOnError

    If CurrentProject.AllForms(PROJECT_FORM).IsLoaded Then

        Set projectFormObject = Forms(PROJECT_FORM)

        If Nz( _
            projectFormObject.Recordset.Fields("ProjectID").value, _
            0) = mContextProjectID Then
            projectFormObject!cboAssemblyTemplateID.Requery
            projectFormObject!cboAssemblyTemplateID = templateID

            If projectFormObject.Dirty Then
                projectFormObject.Dirty = False
            End If
        End If
    End If

    mLinkInProgress = False
    LinkTemplateBackToProjectIfNeeded = True

Exit_Handler:
    Set projectFormObject = Nothing
    Exit Function

Err_Handler:
    mLinkInProgress = False
    LinkTemplateBackToProjectIfNeeded = False

    MsgBox _
        "The template was saved, but it could not be linked back to the project." & _
        vbCrLf & vbCrLf & _
        "Error Number: " & Err.Number & vbCrLf & _
        "Details: " & Err.Description, _
        vbExclamation, _
        "Project Link"
    Resume Exit_Handler

End Function


Private Function GetOpenArgValue(ByVal keyName As String) As String

    Dim rawText As String
    Dim parts() As String
    Dim onePart As Variant
    Dim equalPosition As Long
    Dim thisKey As String
    Dim thisValue As String

    rawText = Nz(Me.OpenArgs, vbNullString)

    If Len(Trim$(rawText)) = 0 Then
        Exit Function
    End If

    parts = Split(rawText, ";")

    For Each onePart In parts

        equalPosition = InStr(1, CStr(onePart), "=", vbBinaryCompare)

        If equalPosition > 0 Then

            thisKey = Left$(CStr(onePart), equalPosition - 1)
            thisValue = Mid$(CStr(onePart), equalPosition + 1)

            If StrComp(thisKey, keyName, vbTextCompare) = 0 Then
                GetOpenArgValue = DecodeArg(thisValue)
                Exit Function
            End If
        End If
    Next onePart

End Function


Private Function DecodeArg(ByVal value As String) As String

    value = Replace(value, "%0D", vbCr, 1, -1, vbTextCompare)
    value = Replace(value, "%0A", vbLf, 1, -1, vbTextCompare)
    value = Replace(value, "%3B", ";", 1, -1, vbTextCompare)
    value = Replace(value, "%3D", "=", 1, -1, vbTextCompare)
    value = Replace(value, "%25", "%", 1, -1, vbTextCompare)

    DecodeArg = value

End Function

'=============================================================
' SAVING, VALIDATION AND NORMALIZATION
'=============================================================

Private Function SaveTemplateRecord( _
    ByVal linkToProject As Boolean _
) As Boolean

    On Error GoTo Err_Handler

    SaveTemplateRecord = False

    NormalizeTemplateValues

    If Not ValidateTemplateRecord Then
        Exit Function
    End If

    If Me.Dirty Then
        Me.Dirty = False
    End If

    If currentTemplateID() <= 0 Then
        MsgBox _
            "The template header could not be saved.", _
            vbExclamation, _
            "Save Template"
        Exit Function
    End If

    SetSubformAvailability

    If Not SaveSubformIfDirty( _
        "SF_AssemblyTemplateMaterials", _
        "material lines") Then
        Exit Function
    End If

    If Not SaveSubformIfDirty( _
        "SF_AssemblyTemplateTools", _
        "tool lines") Then
        Exit Function
    End If

    RecalculateTemplateCosts

    If Me.Dirty Then
        Me.Dirty = False
    End If

    If linkToProject Then
        If Not LinkTemplateBackToProjectIfNeeded Then
            Exit Function
        End If
    End If

    SaveTemplateRecord = True

Exit_Handler:
    Exit Function

Err_Handler:
    MsgBox _
        "The template could not be saved." & vbCrLf & vbCrLf & _
        "Error Number: " & Err.Number & vbCrLf & _
        "Details: " & Err.Description, _
        vbExclamation, _
        "Save Template"
    Resume Exit_Handler

End Function


Private Sub NormalizeTemplateValues()

    Me.txtTemplateName = Trim$(Nz(Me.txtTemplateName, vbNullString))

    If Len(Trim$(Nz(Me.txtTemplateCode, vbNullString))) = 0 Then
        Me.txtTemplateCode = GenerateTemplateCodeFromName(Me.txtTemplateName)
    End If

    Me.txtTemplateCode = UCase$(Trim$(Nz(Me.txtTemplateCode, vbNullString)))
    Me.txtRevisionNo = Trim$(Nz(Me.txtRevisionNo, vbNullString))
    Me.cboStatus = Trim$(Nz(Me.cboStatus, vbNullString))
    Me.txtNOtes = Trim$(Nz(Me.txtNOtes, vbNullString))

    If IsNull(Me.txtOutputQuantity) Then
        Me.txtOutputQuantity = 1
    End If

    If IsNull(Me.txtEstimatedLaborHours) Then
        Me.txtEstimatedLaborHours = 0
    End If

End Sub


Private Function ValidateTemplateRecord() As Boolean

    Dim templateID As Long
    Dim duplicateCriteria As String

    ValidateTemplateRecord = False
    templateID = currentTemplateID()

    If Len(Trim$(Nz(Me.txtTemplateName, vbNullString))) = 0 Then
        MsgBox "Enter the template name.", vbExclamation, "Template Name Required"
        Me.txtTemplateName.SetFocus
        Exit Function
    End If

    If Len(Trim$(Nz(Me.txtTemplateCode, vbNullString))) = 0 Then
        MsgBox "The template code is required.", vbExclamation, "Template Code Required"
        Exit Function
    End If

    duplicateCriteria = _
        "TemplateCode=" & SqlText(Me.txtTemplateCode) & _
        " AND AssemblyTemplateID<>" & templateID

    If DCount("*", "T_AssemblyTemplates", duplicateCriteria) > 0 Then
        MsgBox _
            "The template code already exists.", _
            vbExclamation, _
            "Duplicate Template Code"
        Exit Function
    End If

    If IsNull(Me.txtOutputQuantity) Or _
       Not IsNumeric(Me.txtOutputQuantity) Or _
       CDbl(Me.txtOutputQuantity) <= 0 Then

        MsgBox _
            "Output quantity must be a number greater than zero.", _
            vbExclamation, _
            "Invalid Output Quantity"
        Me.txtOutputQuantity.SetFocus
        Exit Function
    End If

    If Len(Trim$(Nz(Me.cboStatus, vbNullString))) = 0 Then
        MsgBox "Select the template status.", vbExclamation, "Status Required"
        Me.cboStatus.SetFocus
        Exit Function
    End If

    If Not IsNumeric(Me.txtEstimatedLaborHours) Or _
       CDbl(Me.txtEstimatedLaborHours) < 0 Then

        MsgBox _
            "Estimated labor hours must be zero or greater.", _
            vbExclamation, _
            "Invalid Labor Hours"
        Me.txtEstimatedLaborHours.SetFocus
        Exit Function
    End If

    ValidateTemplateRecord = True

End Function


Private Function IsBlankNewTemplateRecord() As Boolean

    IsBlankNewTemplateRecord = _
        (Len(Trim$(Nz(Me.txtTemplateName, vbNullString))) = 0) And _
        (Nz(Me!FinishedItemID, 0) = 0) And _
        (Nz(Me.cboCustomerID, 0) = 0) And _
        (Len(Trim$(Nz(Me.txtRevisionNo, vbNullString))) = 0) And _
        (Len(Trim$(Nz(Me.txtNOtes, vbNullString))) = 0)

End Function

'=============================================================
' TEMPLATE CODE GENERATION
'=============================================================

Private Function GenerateTemplateCodeFromName( _
    ByVal nameText As String _
) As String

    On Error GoTo Err_Handler

    Dim prefix As String
    Dim firstWord As String
    Dim spacePosition As Long
    Dim characterPosition As Long
    Dim oneCharacter As String
    Dim nextNumber As Long
    Dim existingMaximum As Variant

    nameText = Trim$(Nz(nameText, vbNullString))

    If Len(nameText) = 0 Then
        GenerateTemplateCodeFromName = _
            "TEMP-" & Format$(Now(), "yyyymmddhhnnss")
        Exit Function
    End If

    spacePosition = InStr(1, nameText, " ")

    If spacePosition > 0 Then
        firstWord = Left$(nameText, spacePosition - 1)
    Else
        firstWord = nameText
    End If

    For characterPosition = 1 To Len(firstWord)

        oneCharacter = Mid$(firstWord, characterPosition, 1)

        If oneCharacter Like "[A-Za-z]" Then
            prefix = prefix & UCase$(oneCharacter)
        End If
    Next characterPosition

    If Len(prefix) = 0 Then
        prefix = "TEMP"
    End If

    existingMaximum = DMax( _
        "Val(Mid([TemplateCode]," & CStr(Len(prefix) + 2) & "))", _
        "T_AssemblyTemplates", _
        "TemplateCode Like " & SqlText(prefix & "-*") & _
        " AND IsNumeric(Mid([TemplateCode]," & CStr(Len(prefix) + 2) & "))")

    If IsNull(existingMaximum) Then
        nextNumber = 10001
    Else
        nextNumber = CLng(existingMaximum) + 1
    End If

    GenerateTemplateCodeFromName = _
        prefix & "-" & Format$(nextNumber, "00000")

Exit_Handler:
    Exit Function

Err_Handler:
    GenerateTemplateCodeFromName = _
        "TEMP-" & Format$(Now(), "yyyymmddhhnnss")
    Resume Exit_Handler

End Function


Private Function MakeTemplateCodeFromProject( _
    ByVal ProjectCode As String _
) As String

    Dim candidateCode As String

    candidateCode = UCase$(Trim$(ProjectCode))
    candidateCode = Replace(candidateCode, " ", "-")
    candidateCode = Replace(candidateCode, "/", "-")
    candidateCode = Replace(candidateCode, "\", "-")
    candidateCode = Replace(candidateCode, ".", "-")
    candidateCode = Replace(candidateCode, ",", "-")

    If Len(candidateCode) = 0 Then
        candidateCode = "TEMP-" & Format$(Now(), "yyyymmddhhnnss")
    ElseIf Left$(candidateCode, 5) <> "TEMP-" Then
        candidateCode = "TEMP-" & candidateCode
    End If

    If DCount( _
        "*", _
        "T_AssemblyTemplates", _
        "TemplateCode=" & SqlText(candidateCode)) > 0 Then

        candidateCode = _
            candidateCode & "-" & Format$(Now(), "hhnnss")
    End If

    MakeTemplateCodeFromProject = candidateCode

End Function

'=============================================================
' SUBFORM HANDLING
'=============================================================

Private Sub SetSubformAvailability()

    On Error Resume Next

    Dim canEditLines As Boolean

    'Permit focus to enter a line subform after the user has entered
    'a template name. The subform-control Enter event below saves the
    'header before the user starts the first material or tool row.
    canEditLines = _
        (currentTemplateID() > 0) Or _
        (Len(Trim$(Nz(Me.txtTemplateName, vbNullString))) > 0)

    Me.Controls("SF_AssemblyTemplateMaterials").Enabled = canEditLines
    Me.Controls("SF_AssemblyTemplateTools").Enabled = canEditLines

    On Error GoTo 0

End Sub



Private Function SaveSubformIfDirty( _
    ByVal subformControlName As String, _
    ByVal friendlyName As String _
) As Boolean

    On Error GoTo Err_Handler

    SaveSubformIfDirty = True

    If Not ControlExists(subformControlName) Then
        Exit Function
    End If

    With Me.Controls(subformControlName).Form
        If .Dirty Then
            .Dirty = False
        End If
    End With

Exit_Handler:
    Exit Function

Err_Handler:
    SaveSubformIfDirty = False
    MsgBox _
        "The " & friendlyName & " could not be saved." & vbCrLf & vbCrLf & _
        "Error Number: " & Err.Number & vbCrLf & _
        "Details: " & Err.Description, _
        vbExclamation, _
        "Save Template Lines"
    Resume Exit_Handler

End Function


Private Function SubformIsDirty( _
    ByVal subformControlName As String _
) As Boolean

    On Error GoTo Err_Handler

    If ControlExists(subformControlName) Then
        SubformIsDirty = Me.Controls(subformControlName).Form.Dirty
    End If

Exit_Handler:
    Exit Function

Err_Handler:
    SubformIsDirty = False
    Resume Exit_Handler

End Function


Private Sub UndoSubformIfDirty( _
    ByVal subformControlName As String _
)

    On Error Resume Next

    If ControlExists(subformControlName) Then
        With Me.Controls(subformControlName).Form
            If .Dirty Then
                .Undo
            End If
        End With
    End If

    On Error GoTo 0

End Sub

'=============================================================
' COST CALCULATION
'=============================================================

Public Sub RecalculateTemplateCosts()

    On Error GoTo Err_Handler

    Dim templateID As Long
    Dim materialCost As Currency

    templateID = currentTemplateID()

    If templateID <= 0 Then
        Exit Sub
    End If

    materialCost = Nz(DSum( _
        "Nz([QuantityPerOutput],0)*Nz([EstimatedUnitCost],0)", _
        "T_AssemblyTemplateMaterials", _
        "AssemblyTemplateID=" & templateID), 0)

    Me.txtEstimatedMaterialCost = materialCost

    'No labor-rate field currently exists, so total cost contains
    'material cost only. Do not invent a monetary labor cost.
    Me.txtEstimatedTotalCost = materialCost

Exit_Handler:
    Exit Sub

Err_Handler:
    MsgBox _
        "Template costs could not be recalculated." & vbCrLf & vbCrLf & _
        "Error Number: " & Err.Number & vbCrLf & _
        "Details: " & Err.Description, _
        vbExclamation, _
        "Cost Recalculation"
    Resume Exit_Handler

End Sub

'=============================================================
' TEMPLATE NOTES PDF
'=============================================================

Private Sub CreateTemplateNotesPDF()

    On Error GoTo Err_Handler

    Dim templateID As Long
    Dim TemplateName As String
    Dim mainFolder As String
    Dim outputFolder As String
    Dim outputFile As String

    If Not SaveTemplateRecord(True) Then
        Exit Sub
    End If

    templateID = currentTemplateID()
    TemplateName = CleanFileName(Nz(Me.txtTemplateName, vbNullString))

    If Len(TemplateName) = 0 Then
        TemplateName = "Template_" & templateID
    End If

    If Not ReportExists(NOTES_REPORT) Then
        MsgBox _
            "Report " & NOTES_REPORT & " was not found.", _
            vbExclamation, _
            "Template Notes Report"
        Exit Sub
    End If

    mainFolder = Environ$("USERPROFILE") & PDF_MAIN_FOLDER
    outputFolder = Environ$("USERPROFILE") & PDF_TEMPLATE_FOLDER

    EnsureFolderExists mainFolder
    EnsureFolderExists outputFolder

    outputFile = _
        outputFolder & "\Template_Notes_" & TemplateName & ".pdf"

    If CurrentProject.AllReports(NOTES_REPORT).IsLoaded Then
        DoCmd.close acReport, NOTES_REPORT, acSaveNo
    End If

    DoCmd.OpenReport _
        NOTES_REPORT, _
        acViewPreview, _
        , _
        "AssemblyTemplateID=" & templateID, _
        acHidden

    DoCmd.OutputTo _
        acOutputReport, _
        NOTES_REPORT, _
        acFormatPDF, _
        outputFile, _
        True

    DoCmd.close acReport, NOTES_REPORT, acSaveNo

    MsgBox _
        "Template notes PDF created:" & vbCrLf & vbCrLf & _
        outputFile, _
        vbInformation, _
        "Template Notes PDF"

Exit_Handler:
    Exit Sub

Err_Handler:
    On Error Resume Next
    DoCmd.close acReport, NOTES_REPORT, acSaveNo
    On Error GoTo 0

    MsgBox _
        "The template notes PDF could not be created." & vbCrLf & vbCrLf & _
        "Error Number: " & Err.Number & vbCrLf & _
        "Details: " & Err.Description, _
        vbExclamation, _
        "Template Notes PDF"
    Resume Exit_Handler

End Sub


Private Sub EnsureFolderExists(ByVal folderPath As String)

    If Len(Dir$(folderPath, vbDirectory)) = 0 Then
        MkDir folderPath
    End If

End Sub


Private Function ReportExists(ByVal reportName As String) As Boolean

    On Error GoTo Err_Handler

    Dim temporaryName As String

    temporaryName = CurrentProject.AllReports(reportName).Name
    ReportExists = True

Exit_Handler:
    Exit Function

Err_Handler:
    ReportExists = False
    Resume Exit_Handler

End Function


Private Function CleanFileName(ByVal value As String) As String

    Dim badCharacters As Variant
    Dim characterItem As Variant

    badCharacters = Array("\", "/", ":", "*", "?", """", "<", ">", "|")
    value = Trim$(value)

    For Each characterItem In badCharacters
        value = Replace(value, CStr(characterItem), "_")
    Next characterItem

    CleanFileName = value

End Function

'=============================================================
' SAFE HELPERS
'=============================================================

Private Function currentTemplateID() As Long

    On Error GoTo Err_Handler

    If ControlExists("txtAssemblyTemplateID") Then
        currentTemplateID = Nz(Me.Controls("txtAssemblyTemplateID").value, 0)
    Else
        currentTemplateID = Nz(Me.Recordset.Fields("AssemblyTemplateID").value, 0)
    End If

Exit_Handler:
    Exit Function

Err_Handler:
    currentTemplateID = 0
    Resume Exit_Handler

End Function


Private Sub SetRecordField( _
    ByVal fieldName As String, _
    ByVal fieldValue As Variant _
)

    On Error Resume Next

    Me(fieldName) = fieldValue

    On Error GoTo 0

End Sub


Private Function ControlExists(ByVal controlName As String) As Boolean

    On Error GoTo Err_Handler

    Dim controlObject As Control

    Set controlObject = Me.Controls(controlName)
    ControlExists = True

Exit_Handler:
    Set controlObject = Nothing
    Exit Function

Err_Handler:
    ControlExists = False
    Resume Exit_Handler

End Function


Private Function SqlText(ByVal value As String) As String

    SqlText = "'" & Replace(value, "'", "''") & "'"

End Function
Public Function EnsureTemplateHeaderSaved() As Boolean

    On Error GoTo Err_Handler

    EnsureTemplateHeaderSaved = False

    If currentTemplateID() > 0 Then
        EnsureTemplateHeaderSaved = True
        Exit Function
    End If

    If Len(Trim$(Nz(Me.txtTemplateName, vbNullString))) = 0 Then
        MsgBox _
            "Enter the template name before adding materials or tools.", _
            vbExclamation, _
            "Template Name Required"
        Me.txtTemplateName.SetFocus
        Exit Function
    End If

    'False means save the template without running the optional
    'project link-back step.
    If Not SaveTemplateRecord(False) Then
        Exit Function
    End If

    SetSubformAvailability
    EnsureTemplateHeaderSaved = (currentTemplateID() > 0)

Exit_Handler:
    Exit Function

Err_Handler:
    MsgBox _
        "The template header could not be saved before opening its lines." & _
        vbCrLf & vbCrLf & _
        "Error Number: " & Err.Number & vbCrLf & _
        "Details: " & Err.Description, _
        vbExclamation, _
        "Template Header"
    Resume Exit_Handler

End Function



Private Sub SF_AssemblyTemplateMaterials_Enter()

    If Not EnsureTemplateHeaderSaved() Then
        Me.txtTemplateName.SetFocus
    End If

End Sub


Private Sub SF_AssemblyTemplateTools_Enter()

    If Not EnsureTemplateHeaderSaved() Then
        Me.txtTemplateName.SetFocus
    End If

End Sub
Private Sub cmdAddMaterialItem_Click()

    On Error GoTo Err_Handler

    If Not EnsureTemplateHeaderForItemCreation Then
        Exit Sub
    End If

    Me.Controls("SF_AssemblyTemplateMaterials").Form.AddNewComponentFromTemplate

Exit_Handler:
    Exit Sub

Err_Handler:
    MsgBox _
        "The material/component setup screen could not be opened." & _
        vbCrLf & vbCrLf & _
        "Error Number: " & Err.Number & vbCrLf & _
        "Details: " & Err.Description, _
        vbExclamation, _
        "New Material / Component"
    Resume Exit_Handler

End Sub


Private Sub cmdAddToolItem_Click()

    On Error GoTo Err_Handler

    If Not EnsureTemplateHeaderForItemCreation Then
        Exit Sub
    End If

    Me.Controls("SF_AssemblyTemplateTools").Form.AddNewToolFromTemplate

Exit_Handler:
    Exit Sub

Err_Handler:
    MsgBox _
        "The tool/equipment setup screen could not be opened." & _
        vbCrLf & vbCrLf & _
        "Error Number: " & Err.Number & vbCrLf & _
        "Details: " & Err.Description, _
        vbExclamation, _
        "New Tool / Equipment"
    Resume Exit_Handler

End Sub


Private Function EnsureTemplateHeaderForItemCreation() As Boolean

    On Error GoTo Err_Handler

    Dim templateID As Long

    EnsureTemplateHeaderForItemCreation = False

    If Me.NewRecord And _
       Len(Trim$(Nz(Me.txtTemplateName, vbNullString))) = 0 Then

        MsgBox _
            "Enter the template name first.", _
            vbInformation, _
            "Template Name Required"
        Me.txtTemplateName.SetFocus
        Exit Function
    End If

    'This invokes the existing F_AssemblyTemplate validation and
    'saves the header automatically. The user does not press Save.
    If Me.Dirty Then
        Me.Dirty = False
    End If

    On Error Resume Next
    templateID = Nz(Me.Controls("txtAssemblyTemplateID").value, 0)

    If templateID <= 0 Then
        templateID = Nz( _
            Me.Recordset.Fields("AssemblyTemplateID").value, _
            0)
    End If
    On Error GoTo Err_Handler

    If templateID <= 0 Then
        MsgBox _
            "The template header could not be saved. Complete the " & _
            "required template fields and try again.", _
            vbExclamation, _
            "Template Required"
        Exit Function
    End If

    EnsureTemplateHeaderForItemCreation = True

Exit_Handler:
    Exit Function

Err_Handler:
    MsgBox _
        "The template header could not be prepared." & _
        vbCrLf & vbCrLf & _
        "Error Number: " & Err.Number & vbCrLf & _
        "Details: " & Err.Description, _
        vbExclamation, _
        "Template Header"
    Resume Exit_Handler

End Function
