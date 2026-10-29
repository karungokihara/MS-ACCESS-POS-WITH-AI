Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Compare Database
Option Explicit

'=============================================================
' FORM: F_Projects
' CLEAN PROJECT HEADER / NAVIGATION / WORKFLOW MODULE
'
' Main behavior:
'   - Customer PO reference is optional.
'   - Setup buttons work before the project is saved.
'   - Workflow buttons save the project automatically when a
'     permanent ProjectID is required.
'   - Double-clicking Customer, Template, or Output Location opens
'     the selected record; an empty combo opens a new record.
'   - cmdItems opens F_ItemSetup.
'   - No item or template finished-item field is used.
'   - Production Receipt movements establish the project output
'     item; Project Shipment movements consume that output.
'=============================================================

Private Const CUSTOMER_FORM As String = "F_CustomerSetup"
Private Const ITEM_FORM As String = "F_ItemSetup"
Private Const LOCATION_FORM As String = "F_LocationSetup"
Private Const TEMPLATE_FORM As String = "F_AssemblyTemplate"
Private Const WAREHOUSE_LIST_FORM As String = "F_WarehouseList"
Private Const PROJECT_ISSUE_FORM As String = "F_ProjectIssue"
Private Const STOCK_IN_FORM As String = "F_StockIn"
Private Const STOCK_OUT_FORM As String = "F_StockOut"

Private Const REASON_PRODUCTION_RECEIPT As String = "Production Receipt"
Private Const REASON_PROJECT_SHIPMENT As String = "Project Shipment"
Private Const QUANTITY_TOLERANCE As Double = 0.000001

Private mSavingProject As Boolean


'=============================================================
' FORM EVENTS
'=============================================================

Private Sub Form_Load()

    On Error GoTo Err_Handler

    ConfigureFormControls
    ConfigureComboBoxes
    ApplyNewRecordDefaults
    RefreshWorkflowState

Exit_Handler:
    Exit Sub

Err_Handler:
    MsgBox _
        "The project form could not be loaded." & _
        vbCrLf & vbCrLf & _
        "Error Number: " & Err.Number & vbCrLf & _
        "Details: " & Err.Description, _
        vbExclamation, _
        "Project Setup"
    Resume Exit_Handler

End Sub


Private Sub Form_Current()

    On Error Resume Next

    ApplyNewRecordDefaults
    RefreshWorkflowState

    On Error GoTo 0

End Sub


Private Sub Form_BeforeInsert(Cancel As Integer)

    ApplyNewRecordDefaults

End Sub


Private Sub Form_BeforeUpdate(Cancel As Integer)
    UX_CoerceBooleans Me

    On Error GoTo Err_Handler

    NormalizeProjectValues

    If Not ValidateProjectRecord Then
        Cancel = True
        Exit Sub
    End If

Exit_Handler:
    Exit Sub

Err_Handler:
    Cancel = True
    MsgBox _
        "The project could not be validated." & _
        vbCrLf & vbCrLf & _
        "Error Number: " & Err.Number & vbCrLf & _
        "Details: " & Err.Description, _
        vbExclamation, _
        "Project Validation"
    Resume Exit_Handler

End Sub


Private Sub Form_AfterUpdate()

    On Error Resume Next

    RefreshWorkflowState

    On Error GoTo 0

End Sub


Private Sub Form_Activate()

    On Error Resume Next

    'Combo setup helpers refresh only the combo they changed.
    'Avoid a general requery here because Activate can occur while
    'a NotInList event is still completing.
    RefreshWorkflowState

    On Error GoTo 0

End Sub


'=============================================================
' FIELD / COMBO EVENTS
'=============================================================

Private Sub txtProjectName_AfterUpdate()

    RefreshWorkflowState

End Sub


Private Sub txtBuildQuantity_AfterUpdate()

    RefreshWorkflowState

End Sub


Private Sub cboCustomerID_AfterUpdate()

    RefreshWorkflowState

End Sub


Private Sub cboAssemblyTemplateID_AfterUpdate()

    RefreshWorkflowState

End Sub


Private Sub cboFinishedItemID_AfterUpdate()

    RefreshWorkflowState

End Sub


Private Sub cboOutputLocationID_AfterUpdate()

    RefreshWorkflowState

End Sub


Private Sub cboStatus_AfterUpdate()

    RefreshWorkflowState

End Sub


'=============================================================
' COMBO DOUBLE-CLICK NAVIGATION
'=============================================================

Private Sub cboCustomerID_DblClick(Cancel As Integer)

    Cancel = True
    OpenCustomerSetup

End Sub


Private Sub cboAssemblyTemplateID_DblClick(Cancel As Integer)

    Cancel = True
    OpenTemplateSetup

End Sub


Private Sub cboOutputLocationID_DblClick(Cancel As Integer)

    Cancel = True
    OpenLocationSetup

End Sub


'=============================================================
' PRIMARY BUTTONS
'=============================================================

Private Sub cmdNewProject_Click()

    On Error GoTo Err_Handler

    Dim reuseCurrentNewRecord As Boolean

    If Me.NewRecord And IsBlankNewProjectRecord Then
        reuseCurrentNewRecord = True

        If Me.Dirty Then
            Me.Undo
        End If
    ElseIf Me.Dirty Then
        If Not SaveCurrentProject Then
            Exit Sub
        End If
    End If

    If Not reuseCurrentNewRecord Then
        DoCmd.GoToRecord acDataForm, Me.Name, acNewRec
    End If

    ApplyNewRecordDefaults
    RefreshWorkflowState
    Me.txtProjectName.SetFocus

Exit_Handler:
    Exit Sub

Err_Handler:
    MsgBox _
        "A new project could not be opened." & _
        vbCrLf & vbCrLf & _
        "Error Number: " & Err.Number & vbCrLf & _
        "Details: " & Err.Description, _
        vbExclamation, _
        "New Project"
    Resume Exit_Handler

End Sub


Private Sub cmdSaveProject_Click()

    If SaveCurrentProject Then
        MsgBox _
            "Project saved successfully.", _
            vbInformation, _
            "Project Saved"
    End If

End Sub


Private Sub cmdOpenTemplate_Click()

    OpenTemplateSetup

End Sub


Private Sub cmdClose_Click()

    On Error GoTo Err_Handler

    Dim responseValue As VbMsgBoxResult

    If Me.Dirty Then
        responseValue = MsgBox( _
            "Save the project changes before closing?", _
            vbQuestion + vbYesNoCancel, _
            "Close Project")

        Select Case responseValue
            Case vbYes
                If Not SaveCurrentProject Then
                    Exit Sub
                End If

            Case vbNo
                Me.Undo

            Case vbCancel
                Exit Sub
        End Select
    End If

    DoCmd.close acForm, Me.Name

Exit_Handler:
    Exit Sub

Err_Handler:
    MsgBox _
        "The project form could not be closed." & _
        vbCrLf & vbCrLf & _
        "Error Number: " & Err.Number & vbCrLf & _
        "Details: " & Err.Description, _
        vbExclamation, _
        "Close Project"
    Resume Exit_Handler

End Sub


'=============================================================
' RELATED SETUP BUTTONS
'=============================================================

Private Sub cmdOpenCustomer_Click()

    OpenCustomerSetup

End Sub


Private Sub cmdItems_Click()

    On Error GoTo Err_Handler

    OpenStandaloneForm ITEM_FORM

Exit_Handler:
    Exit Sub

Err_Handler:
    MsgBox _
        "Item Setup could not be opened." & _
        vbCrLf & vbCrLf & _
        "Error Number: " & Err.Number & vbCrLf & _
        "Details: " & Err.Description, _
        vbExclamation, _
        "Item Setup"
    Resume Exit_Handler

End Sub


Private Sub cmdLocation_Click()

    OpenLocationSetup

End Sub


Private Sub OpenCustomerSetup()

    On Error GoTo Err_Handler

    OpenSelectedOrNewComboRecord _
        Me.cboCustomerID, _
        CUSTOMER_FORM, _
        "T_Customers", _
        "CustomerID", _
        "CustomerName", _
        vbNullString

    RefreshWorkflowState

Exit_Handler:
    Exit Sub

Err_Handler:
    MsgBox _
        "Customer Setup could not be opened." & _
        vbCrLf & vbCrLf & _
        "Error Number: " & Err.Number & vbCrLf & _
        "Details: " & Err.Description, _
        vbExclamation, _
        "Customer Setup"
    Resume Exit_Handler

End Sub


Private Sub OpenTemplateSetup()

    On Error GoTo Err_Handler

    Dim suggestedName As String

    If Nz(Me.cboAssemblyTemplateID, 0) <= 0 Then
        suggestedName = Trim$(Nz(Me.txtProjectName, vbNullString))

        If Len(suggestedName) > 0 Then
            suggestedName = suggestedName & " Template"
        End If
    End If

    OpenSelectedOrNewComboRecord _
        Me.cboAssemblyTemplateID, _
        TEMPLATE_FORM, _
        "T_AssemblyTemplates", _
        "AssemblyTemplateID", _
        "TemplateName", _
        suggestedName

    RefreshWorkflowState

Exit_Handler:
    Exit Sub

Err_Handler:
    MsgBox _
        "The Template / Estimate form could not be opened." & _
        vbCrLf & vbCrLf & _
        "Error Number: " & Err.Number & vbCrLf & _
        "Details: " & Err.Description, _
        vbExclamation, _
        "Template / Estimate"
    Resume Exit_Handler

End Sub


Private Sub OpenLocationSetup()

    On Error GoTo Err_Handler

    If Not WarehouseAvailableForLocation Then
        Exit Sub
    End If

    OpenSelectedOrNewComboRecord _
        Me.cboOutputLocationID, _
        LOCATION_FORM, _
        "T_Locations", _
        "LocationID", _
        "LocationName", _
        vbNullString

    RefreshWorkflowState

Exit_Handler:
    Exit Sub

Err_Handler:
    MsgBox _
        "Location Setup could not be opened." & _
        vbCrLf & vbCrLf & _
        "Error Number: " & Err.Number & vbCrLf & _
        "Details: " & Err.Description, _
        vbExclamation, _
        "Location Setup"
    Resume Exit_Handler

End Sub


'=============================================================
' NOT-IN-LIST EVENTS
'=============================================================

Private Sub cboCustomerID_NotInList( _
    NewData As String, _
    Response As Integer _
)

    Response = acDataErrContinue

    If CreateTypedComboEntry( _
        NewData, _
        CUSTOMER_FORM, _
        "T_Customers", _
        "CustomerID", _
        "CustomerName") Then

        Response = acDataErrAdded
    Else
        Me.cboCustomerID.Undo
    End If

End Sub


Private Sub cboAssemblyTemplateID_NotInList( _
    NewData As String, _
    Response As Integer _
)

    Response = acDataErrContinue

    If CreateTypedComboEntry( _
        NewData, _
        TEMPLATE_FORM, _
        "T_AssemblyTemplates", _
        "AssemblyTemplateID", _
        "TemplateName") Then

        Response = acDataErrAdded
    Else
        Me.cboAssemblyTemplateID.Undo
    End If

End Sub


Private Sub cboOutputLocationID_NotInList( _
    NewData As String, _
    Response As Integer _
)

    Response = acDataErrContinue

    If Not WarehouseAvailableForLocation Then
        Me.cboOutputLocationID.Undo
        Exit Sub
    End If

    If CreateTypedComboEntry( _
        NewData, _
        LOCATION_FORM, _
        "T_Locations", _
        "LocationID", _
        "LocationName") Then

        Response = acDataErrAdded
    Else
        Me.cboOutputLocationID.Undo
    End If

End Sub


'=============================================================
' PROJECT WORKFLOW BUTTONS
'=============================================================

Private Sub cmdIssueMaterials_Click()

    OpenProjectIssue "Materials"

End Sub


Private Sub cmdIssueTools_Click()

    OpenProjectIssue "Tools"

End Sub


Private Sub OpenProjectIssue(ByVal issueType As String)

    On Error GoTo Err_Handler

    If Not ProjectReadyForIssue Then
        Exit Sub
    End If

    CloseProjectIssueIfOpen

    DoCmd.OpenForm _
        PROJECT_ISSUE_FORM, _
        acNormal, _
        , _
        , _
        , _
        , _
        "ProjectID=" & CStr(CurrentProjectID()) & _
        ";IssueType=" & issueType

Exit_Handler:
    Exit Sub

Err_Handler:
    MsgBox _
        issueType & " issue could not be opened." & _
        vbCrLf & vbCrLf & _
        "Error Number: " & Err.Number & vbCrLf & _
        "Details: " & Err.Description, _
        vbExclamation, _
        "Project Issue"
    Resume Exit_Handler

End Sub


Private Sub cmdReturnTools_Click()

    On Error GoTo Err_Handler

    If Not SaveCurrentProject Then
        Exit Sub
    End If

    CloseProjectIssueIfOpen

    DoCmd.OpenForm _
        PROJECT_ISSUE_FORM, _
        acNormal, _
        , _
        , _
        , _
        , _
        "ProjectID=" & CStr(CurrentProjectID()) & _
        ";IssueType=ToolReturn"

Exit_Handler:
    Exit Sub

Err_Handler:
    MsgBox _
        "Tool Return could not be opened." & _
        vbCrLf & vbCrLf & _
        "Error Number: " & Err.Number & vbCrLf & _
        "Details: " & Err.Description, _
        vbExclamation, _
        "Tool Return"
    Resume Exit_Handler

End Sub


Private Sub CloseProjectIssueIfOpen()

    If IsFormLoaded(PROJECT_ISSUE_FORM) Then
        DoCmd.close acForm, PROJECT_ISSUE_FORM, acSaveNo
    End If

End Sub


Private Sub cmdReceiveFinishedItem_Click()

    On Error GoTo Err_Handler

    Dim projectIDValue As Long
    Dim finishedItemIDValue As Long

    If Not ProjectReadyForReceipt Then
        Exit Sub
    End If

    projectIDValue = CurrentProjectID()
    finishedItemIDValue = Nz(Me.cboFinishedItemID, 0)

    If IsFormLoaded(STOCK_IN_FORM) Then

        Call Forms(STOCK_IN_FORM).LoadProjectReceipt( _
            projectIDValue, _
            finishedItemIDValue)

        Forms(STOCK_IN_FORM).SetFocus

    Else

        DoCmd.OpenForm _
            STOCK_IN_FORM, _
            acNormal, _
            , _
            , _
            acFormAdd, _
            , _
            "Context=ProjectReceipt;" & _
            "ProjectID=" & CStr(projectIDValue) & ";" & _
            "ItemID=" & CStr(finishedItemIDValue)

    End If

Exit_Handler:
    Exit Sub

Err_Handler:
    MsgBox _
        "Project output receipt could not be opened." & _
        vbCrLf & vbCrLf & _
        "Error Number: " & Err.Number & vbCrLf & _
        "Details: " & Err.Description, _
        vbExclamation, _
        "Project Receipt"
    Resume Exit_Handler

End Sub


Private Sub cmdMarkReadyToShip_Click()

    On Error GoTo Err_Handler

    If Not ProjectHasShippableOutput Then
        Exit Sub
    End If

    Me.cboStatus = "Ready to Ship"

    If Me.Dirty Then
        Me.Dirty = False
    End If

    RefreshWorkflowState

    MsgBox _
        "Project marked Ready to Ship.", _
        vbInformation, _
        "Project Status"

Exit_Handler:
    Exit Sub

Err_Handler:
    MsgBox _
        "The project status could not be updated." & _
        vbCrLf & vbCrLf & _
        "Error Number: " & Err.Number & vbCrLf & _
        "Details: " & Err.Description, _
        vbExclamation, _
        "Project Status"
    Resume Exit_Handler

End Sub


Private Sub cmdPostShipment_Click()

    On Error GoTo Err_Handler

    If Not ProjectReadyForShipment Then
        Exit Sub
    End If

    If IsFormLoaded(STOCK_OUT_FORM) Then
        Call Forms(STOCK_OUT_FORM).LoadProjectShipment( _
            CurrentProjectID())
        Forms(STOCK_OUT_FORM).SetFocus
    Else
        DoCmd.OpenForm _
            STOCK_OUT_FORM, _
            acNormal, _
            , _
            , _
            acFormAdd, _
            , _
            "Context=ProjectShipment;ProjectID=" & _
                CStr(CurrentProjectID())
    End If

Exit_Handler:
    Exit Sub

Err_Handler:
    MsgBox _
        "Project shipment could not be opened." & _
        vbCrLf & vbCrLf & _
        "Error Number: " & Err.Number & vbCrLf & _
        "Details: " & Err.Description, _
        vbExclamation, _
        "Project Shipment"
    Resume Exit_Handler

End Sub


'=============================================================
' FORM / COMBO CONFIGURATION
'=============================================================

Private Sub ConfigureFormControls()

    If ControlExists("txtProjectID") Then
        Me.Controls("txtProjectID").Visible = False
        Me.Controls("txtProjectID").Enabled = False
        Me.Controls("txtProjectID").Locked = True
    End If

    Me.txtProjectCode.Locked = True
    Me.txtProjectCode.TabStop = False
    Me.txtShippedDate.Locked = True

    Me.txtWorkInstructions.EnterKeyBehavior = True
    Me.txtWorkInstructions.ScrollBars = 2

    Me.txtNOtes.EnterKeyBehavior = True
    Me.txtNOtes.ScrollBars = 2

    Me.txtBuildQuantity.Format = "General Number"
    Me.txtQuantityShipped.Format = "General Number"
    Me.txtBudget.Format = "Currency"

    SetControlEnabled "cmdOpenCustomer", True
    SetControlEnabled "cmdOpenTemplate", True
    SetControlEnabled "cmdItems", True
    SetControlEnabled "cmdLocation", True

End Sub


Private Sub ConfigureComboBoxes()

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
        .Requery
    End With

    With Me.cboAssemblyTemplateID
        .RowSourceType = "Table/Query"
        .RowSource = _
            "SELECT [AssemblyTemplateID], [TemplateName], " & _
                "[TemplateCode], Nz([Status],'') AS TemplateStatus " & _
            "FROM [T_AssemblyTemplates] " & _
            "ORDER BY [TemplateName], [TemplateCode];"
        .BoundColumn = 1
        .ColumnCount = 4
        .ColumnWidths = "0cm;7cm;3.5cm;3cm"
        .LimitToList = True
        .AutoExpand = True
        .Requery
    End With

    With Me.cboFinishedItemID
        .RowSourceType = "Table/Query"
        .RowSource = _
            "SELECT ItemID, ItemName, InternalSKU " & _
            "FROM T_Items " & _
            "ORDER BY ItemName, InternalSKU;"
        .BoundColumn = 1
        .ColumnCount = 3
        .ColumnWidths = "0cm;7cm;3.5cm"
        .LimitToList = True
        .AutoExpand = True
        .Requery
    End With

    With Me.cboOutputLocationID
        .RowSourceType = "Table/Query"
        .RowSource = _
            "SELECT L.LocationID, L.LocationName, " & _
                "L.LocationCode, W.WarehouseName " & _
            "FROM T_Locations AS L " & _
            "LEFT JOIN T_Warehouses AS W " & _
                "ON L.WarehouseID=W.WarehouseID " & _
            "ORDER BY W.WarehouseName, L.LocationName;"
        .BoundColumn = 1
        .ColumnCount = 4
        .ColumnWidths = "0cm;6cm;3cm;5cm"
        .LimitToList = True
        .AutoExpand = True
        .Requery
    End With

    With Me.cboStatus
        .RowSourceType = "Value List"
        .RowSource = _
            "Quoted;Approved;In Progress;Materials Issued;" & _
            "Tools Issued;Finished Received;Ready to Ship;" & _
            "Shipped;Cancelled"
        .BoundColumn = 1
        .ColumnCount = 1
        .LimitToList = True
    End With

End Sub


'=============================================================
' DEFAULTS, NORMALIZATION, VALIDATION, AND SAVING
'=============================================================

Private Sub ApplyNewRecordDefaults()

    If Not Me.NewRecord Then
        Exit Sub
    End If

    If Len(Trim$(Nz(Me.txtProjectCode, vbNullString))) = 0 Then
        Me.txtProjectCode = GenerateProjectCode()
    End If

    If IsNull(Me.txtBuildQuantity) Then
        Me.txtBuildQuantity = 1
    End If

    If IsNull(Me.txtQuantityShipped) Then
        Me.txtQuantityShipped = 0
    End If

    If Len(Trim$(Nz(Me.cboStatus, vbNullString))) = 0 Then
        Me.cboStatus = "Quoted"
    End If

    If IsNull(Me.txtStartDate) Then
        Me.txtStartDate = Date
    End If

End Sub


Private Sub NormalizeProjectValues()

    Me.txtProjectCode = _
        UCase$(Trim$(Nz(Me.txtProjectCode, vbNullString)))
    Me.txtProjectName = _
        Trim$(Nz(Me.txtProjectName, vbNullString))

    NormalizeOptionalTextControl "txtCustomerPOReference"
    NormalizeOptionalTextControl "txtWorkInstructions"
    NormalizeOptionalTextControl "txtNOtes"

    If Len(Me.txtProjectCode) = 0 Then
        Me.txtProjectCode = GenerateProjectCode()
    End If

    If IsNull(Me.txtQuantityShipped) Then
        Me.txtQuantityShipped = 0
    End If

End Sub


Private Sub NormalizeOptionalTextControl( _
    ByVal controlName As String _
)

    Dim normalizedText As String

    normalizedText = _
        Trim$(Nz(Me.Controls(controlName).value, vbNullString))

    If Len(normalizedText) = 0 Then
        Me.Controls(controlName).value = Null
    Else
        Me.Controls(controlName).value = normalizedText
    End If

End Sub


Private Function ValidateProjectRecord() As Boolean

    ValidateProjectRecord = False

    If Len(Trim$(Nz(Me.txtProjectCode, vbNullString))) = 0 Then
        MsgBox _
            "The project code is required.", _
            vbExclamation, _
            "Project Code"
        Me.txtProjectCode.SetFocus
        Exit Function
    End If

    If DCount( _
        "*", _
        "T_Projects", _
        "ProjectCode=" & SqlText(Me.txtProjectCode) & _
        " AND ProjectID<>" & CStr(CurrentProjectID())) > 0 Then

        MsgBox _
            "The project code already exists.", _
            vbExclamation, _
            "Duplicate Project Code"
        Exit Function
    End If

    If Len(Trim$(Nz(Me.txtProjectName, vbNullString))) = 0 Then
        MsgBox _
            "Enter the project name.", _
            vbExclamation, _
            "Project Name Required"
        Me.txtProjectName.SetFocus
        Exit Function
    End If

    If IsNull(Me.txtBuildQuantity) Or _
       Not IsNumeric(Me.txtBuildQuantity) Or _
       CDbl(Me.txtBuildQuantity) <= 0 Then

        MsgBox _
            "Build quantity must be a number greater than zero.", _
            vbExclamation, _
            "Invalid Build Quantity"
        Me.txtBuildQuantity.SetFocus
        Exit Function
    End If

    If Len(Trim$(Nz(Me.cboStatus, vbNullString))) = 0 Then
        MsgBox _
            "Select the project status.", _
            vbExclamation, _
            "Status Required"
        Me.cboStatus.SetFocus
        Exit Function
    End If

    'Customer PO reference is intentionally not required.
    ValidateProjectRecord = True

End Function


Private Function SaveCurrentProject() As Boolean

    On Error GoTo Err_Handler

    SaveCurrentProject = False

    If mSavingProject Then
        SaveCurrentProject = (CurrentProjectID() > 0)
        Exit Function
    End If

    NormalizeProjectValues

    If Not ValidateProjectRecord Then
        Exit Function
    End If

    mSavingProject = True

    If Me.Dirty Then
        Me.Dirty = False
    End If

    mSavingProject = False
    RefreshWorkflowState

    SaveCurrentProject = (CurrentProjectID() > 0)

Exit_Handler:
    mSavingProject = False
    Exit Function

Err_Handler:
    MsgBox _
        "The project could not be saved." & _
        vbCrLf & vbCrLf & _
        "Error Number: " & Err.Number & vbCrLf & _
        "Details: " & Err.Description, _
        vbExclamation, _
        "Save Project"
    Resume Exit_Handler

End Function


Private Function IsBlankNewProjectRecord() As Boolean

    IsBlankNewProjectRecord = _
        (Len(Trim$(Nz(Me.txtProjectName, vbNullString))) = 0) And _
        (Nz(Me.cboCustomerID, 0) = 0) And _
        (Nz(Me.cboAssemblyTemplateID, 0) = 0) And _
        (Nz(Me.cboFinishedItemID, 0) = 0) And _
        (Nz(Me.cboOutputLocationID, 0) = 0) And _
        (Len(Trim$(Nz( _
            Me.txtCustomerPOReference, vbNullString))) = 0) And _
        (Len(Trim$(Nz( _
            Me.txtWorkInstructions, vbNullString))) = 0) And _
        (Len(Trim$(Nz(Me.txtNOtes, vbNullString))) = 0)

End Function


Private Function GenerateProjectCode() As String

    Dim nextNumber As Long
    Dim candidate As String

    nextNumber = Nz(DMax("ProjectID", "T_Projects"), 0) + 1

    Do
        candidate = "PRJ-" & Format$(nextNumber, "00000")

        If DCount( _
            "*", _
            "T_Projects", _
            "ProjectCode=" & SqlText(candidate)) = 0 Then

            Exit Do
        End If

        nextNumber = nextNumber + 1
    Loop

    GenerateProjectCode = candidate

End Function


'=============================================================
' WORKFLOW READINESS / BUTTON STATE
'=============================================================

Private Function ProjectReadyForIssue() As Boolean

    ProjectReadyForIssue = False

    If Not SaveCurrentProject Then
        Exit Function
    End If

    If Nz(Me.cboAssemblyTemplateID, 0) <= 0 Then
        MsgBox _
            "Select or create an assembly template first.", _
            vbExclamation, _
            "Template Required"
        Me.cboAssemblyTemplateID.SetFocus
        Exit Function
    End If

    If Nz(Me.txtBuildQuantity, 0) <= 0 Then
        MsgBox _
            "Enter a Build Quantity greater than zero.", _
            vbExclamation, _
            "Build Quantity Required"
        Me.txtBuildQuantity.SetFocus
        Exit Function
    End If

    ProjectReadyForIssue = True

End Function


Private Function ProjectReadyForReceipt() As Boolean

    Dim outputItemID As Long
    Dim finishedItemIDValue As Long
    Dim receivedQuantity As Double
    Dim shippedQuantity As Double

    ProjectReadyForReceipt = False

    If Not ProjectReadyForIssue Then
        Exit Function
    End If

    finishedItemIDValue = Nz(Me.cboFinishedItemID, 0)

    If finishedItemIDValue <= 0 Then
        MsgBox _
            "Select the Finished Item first.", _
            vbExclamation, _
            "Finished Item Required"

        Me.cboFinishedItemID.SetFocus
        Exit Function
    End If

    If DCount( _
        "*", _
        "T_Items", _
        "ItemID=" & CStr(finishedItemIDValue)) = 0 Then

        MsgBox _
            "The selected Finished Item could not be found in T_Items.", _
            vbExclamation, _
            "Finished Item Not Found"

        Me.cboFinishedItemID.Requery
        Me.cboFinishedItemID.SetFocus
        Exit Function
    End If

    If Nz(Me.cboOutputLocationID, 0) <= 0 Then
        MsgBox _
            "Select the project output location first.", _
            vbExclamation, _
            "Output Location Required"
        Me.cboOutputLocationID.SetFocus
        Exit Function
    End If

    GetProjectOutputSummary _
        CurrentProjectID(), _
        outputItemID, _
        receivedQuantity, _
        shippedQuantity

    If outputItemID > 0 And _
       outputItemID <> finishedItemIDValue Then

        MsgBox _
            "Earlier production receipts for this project used a " & _
            "different finished item." & vbCrLf & vbCrLf & _
            "Select the same finished item used by the earlier " & _
            "receipts before continuing.", _
            vbExclamation, _
            "Finished Item Does Not Match"

        Me.cboFinishedItemID.SetFocus
        Exit Function
    End If

    If receivedQuantity + QUANTITY_TOLERANCE >= _
       CDbl(Nz(Me.txtBuildQuantity, 0)) Then

        MsgBox _
            "The full project Build Quantity has already been received.", _
            vbInformation, _
            "Project Receipt Complete"
        Exit Function
    End If

    ProjectReadyForReceipt = True

End Function


Private Function ProjectHasShippableOutput() As Boolean

    On Error GoTo Err_Handler

    Dim outputItemID As Long
    Dim receivedQuantity As Double
    Dim shippedQuantity As Double

    ProjectHasShippableOutput = False

    If Not SaveCurrentProject Then
        Exit Function
    End If

    GetProjectOutputSummary _
        CurrentProjectID(), _
        outputItemID, _
        receivedQuantity, _
        shippedQuantity

    If outputItemID <= 0 Or _
       receivedQuantity <= QUANTITY_TOLERANCE Then

        MsgBox _
            "Receive at least one project output item before " & _
            "marking the project Ready to Ship.", _
            vbExclamation, _
            "Project Output Required"
        Exit Function
    End If

    If receivedQuantity - shippedQuantity <= _
       QUANTITY_TOLERANCE Then

        MsgBox _
            "There is no received project output waiting to ship.", _
            vbInformation, _
            "Nothing to Ship"
        Exit Function
    End If

    ProjectHasShippableOutput = True

Exit_Handler:
    Exit Function

Err_Handler:
    MsgBox _
        "Project shipment readiness could not be checked." & _
        vbCrLf & vbCrLf & _
        "Error Number: " & Err.Number & vbCrLf & _
        "Details: " & Err.Description, _
        vbExclamation, _
        "Project Shipment"
    Resume Exit_Handler

End Function


Private Function ProjectReadyForShipment() As Boolean

    ProjectReadyForShipment = False

    If Not ProjectHasShippableOutput Then
        Exit Function
    End If

    If StrComp( _
        Trim$(Nz(Me.cboStatus, vbNullString)), _
        "Ready to Ship", _
        vbTextCompare) <> 0 Then

        MsgBox _
            "Mark the project Ready to Ship before posting " & _
            "the shipment.", _
            vbExclamation, _
            "Project Not Ready"
        Exit Function
    End If

    ProjectReadyForShipment = True

End Function


Private Sub RefreshWorkflowState()

    Dim savedProject As Boolean
    Dim hasProjectName As Boolean
    Dim hasTemplate As Boolean
    Dim hasFinishedItem As Boolean
    Dim hasOutputLocation As Boolean
    Dim hasBuildQuantity As Boolean
    Dim outputItemID As Long
    Dim receivedQuantity As Double
    Dim shippedQuantity As Double
    Dim unshippedQuantity As Double
    Dim projectStatus As String

    savedProject = (CurrentProjectID() > 0)
    hasProjectName = _
        (Len(Trim$(Nz(Me.txtProjectName, vbNullString))) > 0)
    hasTemplate = (Nz(Me.cboAssemblyTemplateID, 0) > 0)
    hasFinishedItem = (Nz(Me.cboFinishedItemID, 0) > 0)
    hasOutputLocation = (Nz(Me.cboOutputLocationID, 0) > 0)
    hasBuildQuantity = (Nz(Me.txtBuildQuantity, 0) > 0)
    projectStatus = Trim$(Nz(Me.cboStatus, vbNullString))

    If savedProject Then
        On Error Resume Next

        GetProjectOutputSummary _
            CurrentProjectID(), _
            outputItemID, _
            receivedQuantity, _
            shippedQuantity

        On Error GoTo 0
    End If

    unshippedQuantity = receivedQuantity - shippedQuantity

    If unshippedQuantity < 0 Then
        unshippedQuantity = 0
    End If

    'Setup/navigation never requires a saved project.
    SetControlEnabled "cmdOpenCustomer", True
    SetControlEnabled "cmdOpenTemplate", True
    SetControlEnabled "cmdItems", True
    SetControlEnabled "cmdLocation", True

    'Workflow buttons use the values currently on screen. Their click
    'procedures save automatically when a ProjectID is required.
    SetControlEnabled _
        "cmdIssueMaterials", _
        hasProjectName And hasTemplate And hasBuildQuantity

    SetControlEnabled _
        "cmdIssueTools", _
        hasProjectName And hasTemplate And hasBuildQuantity

    SetControlEnabled _
        "cmdReturnTools", _
        savedProject Or hasProjectName

    SetControlEnabled _
        "cmdReceiveFinishedItem", _
        hasProjectName And hasTemplate And _
        hasFinishedItem And hasBuildQuantity And _
        hasOutputLocation And _
        (receivedQuantity + QUANTITY_TOLERANCE < _
            CDbl(Nz(Me.txtBuildQuantity, 0)))

    SetControlEnabled _
        "cmdMarkReadyToShip", _
        savedProject And _
        (outputItemID > 0) And _
        (unshippedQuantity > QUANTITY_TOLERANCE)

    SetControlEnabled _
        "cmdPostShipment", _
        savedProject And _
        (unshippedQuantity > QUANTITY_TOLERANCE) And _
        (StrComp( _
            projectStatus, _
            "Ready to Ship", _
            vbTextCompare) = 0)

End Sub


Private Sub GetProjectOutputSummary( _
    ByVal ProjectID As Long, _
    ByRef outputItemID As Long, _
    ByRef receivedQuantity As Double, _
    ByRef shippedQuantity As Double _
)

    On Error GoTo Err_Handler

    Dim db As DAO.Database
    Dim rs As DAO.Recordset
    Dim itemCount As Long
    Dim movementItemID As Long
    Dim errorNumber As Long
    Dim errorDescription As String
    Dim errorSource As String

    outputItemID = 0
    receivedQuantity = 0
    shippedQuantity = 0

    If ProjectID <= 0 Then
        Exit Sub
    End If

    Set db = CurrentDb
    Set rs = db.OpenRecordset( _
        "SELECT M.ItemID, Sum(Nz(M.Quantity,0)) " & _
            "AS MovementQuantity " & _
        "FROM T_StockMovements AS M " & _
        "INNER JOIN T_StockMovementTypes AS MT " & _
            "ON M.MovementTypeID=MT.MovementTypeID " & _
        "WHERE M.ProjectID=" & ProjectID & " " & _
            "AND MT.MovementTypeCode='STOCK_IN' " & _
            "AND M.Reason=" & _
                SqlText(REASON_PRODUCTION_RECEIPT) & " " & _
        "GROUP BY M.ItemID;", _
        dbOpenSnapshot)

    Do While Not rs.EOF
        itemCount = itemCount + 1

        If itemCount = 1 Then
            outputItemID = Nz(rs!ItemID, 0)
        End If

        receivedQuantity = _
            receivedQuantity + Nz(rs!MovementQuantity, 0)

        rs.MoveNext
    Loop

    rs.close
    Set rs = Nothing

    If itemCount > 1 Then
        Err.Raise _
            vbObjectError + 7501, _
            "GetProjectOutputSummary", _
            "This project has production receipts for more than " & _
            "one output item. Correct the movement history first."
    End If

    Set rs = db.OpenRecordset( _
        "SELECT M.ItemID, Sum(Nz(M.Quantity,0)) " & _
            "AS MovementQuantity " & _
        "FROM T_StockMovements AS M " & _
        "INNER JOIN T_StockMovementTypes AS MT " & _
            "ON M.MovementTypeID=MT.MovementTypeID " & _
        "WHERE M.ProjectID=" & ProjectID & " " & _
            "AND MT.MovementTypeCode='STOCK_OUT' " & _
            "AND M.Reason=" & _
                SqlText(REASON_PROJECT_SHIPMENT) & " " & _
        "GROUP BY M.ItemID;", _
        dbOpenSnapshot)

    Do While Not rs.EOF
        movementItemID = Nz(rs!ItemID, 0)

        If outputItemID <= 0 Or movementItemID <> outputItemID Then
            Err.Raise _
                vbObjectError + 7502, _
                "GetProjectOutputSummary", _
                "A project shipment is not linked to the " & _
                "project's received output item. Correct the " & _
                "movement history first."
        End If

        shippedQuantity = _
            shippedQuantity + Nz(rs!MovementQuantity, 0)

        rs.MoveNext
    Loop

    If shippedQuantity > receivedQuantity + QUANTITY_TOLERANCE Then
        Err.Raise _
            vbObjectError + 7503, _
            "GetProjectOutputSummary", _
            "Project shipment quantity exceeds its production " & _
            "receipt quantity. Correct the movement history first."
    End If

Exit_Handler:
    On Error Resume Next
    If Not (rs Is Nothing) Then rs.close
    Set rs = Nothing
    Set db = Nothing
    On Error GoTo 0
    Exit Sub

Err_Handler:
    errorNumber = Err.Number
    errorDescription = Err.Description
    errorSource = Err.Source

    On Error Resume Next
    If Not (rs Is Nothing) Then rs.close
    Set rs = Nothing
    Set db = Nothing
    On Error GoTo 0

    Err.Raise errorNumber, errorSource, errorDescription

End Sub


'=============================================================
' COMBO SETUP NAVIGATION
'=============================================================

Private Sub OpenSelectedOrNewComboRecord( _
    ByVal comboControl As Access.ComboBox, _
    ByVal setupFormName As String, _
    ByVal tableName As String, _
    ByVal idFieldName As String, _
    ByVal matchingFieldName As String, _
    ByVal suggestedText As String _
)

    If Nz(comboControl.value, 0) > 0 Then
        OpenExistingComboRecord _
            comboControl, _
            setupFormName, _
            tableName, _
            idFieldName
    Else
        OpenNewComboRecord _
            comboControl, _
            setupFormName, _
            tableName, _
            idFieldName, _
            matchingFieldName, _
            suggestedText
    End If

End Sub


Private Sub OpenExistingComboRecord( _
    ByVal comboControl As Access.ComboBox, _
    ByVal setupFormName As String, _
    ByVal tableName As String, _
    ByVal idFieldName As String _
)

    On Error GoTo Err_Handler

    Dim selectedID As Long

    selectedID = Nz(comboControl.value, 0)

    If selectedID <= 0 Then
        Exit Sub
    End If

    If Not FormExists(setupFormName) Then
        Err.Raise _
            vbObjectError + 7301, _
            "OpenExistingComboRecord", _
            "The setup form '" & setupFormName & "' does not exist."
    End If

    If Not CloseLoadedSetupForm(setupFormName) Then
        Exit Sub
    End If

    DoCmd.OpenForm _
        formName:=setupFormName, _
        View:=acNormal, _
        WhereCondition:= _
            QuoteIdentifier(idFieldName) & "=" & CStr(selectedID), _
        DataMode:=acFormEdit, _
        WindowMode:=acDialog

    comboControl.Requery

    If DCount( _
        "*", _
        QuoteIdentifier(tableName), _
        QuoteIdentifier(idFieldName) & "=" & CStr(selectedID)) > 0 Then

        comboControl.value = selectedID
    Else
        comboControl.value = Null
    End If

Exit_Handler:
    Exit Sub

Err_Handler:
    If Err.Number = 2501 Then
        Resume Exit_Handler
    End If

    Err.Raise Err.Number, Err.Source, Err.Description

End Sub


Private Sub OpenNewComboRecord( _
    ByVal comboControl As Access.ComboBox, _
    ByVal setupFormName As String, _
    ByVal tableName As String, _
    ByVal idFieldName As String, _
    ByVal matchingFieldName As String, _
    ByVal suggestedText As String _
)

    On Error GoTo Err_Handler

    Dim oldMaximumID As Long
    Dim newMaximumID As Long
    Dim savedID As Variant
    Dim openArguments As String

    If Not FormExists(setupFormName) Then
        Err.Raise _
            vbObjectError + 7302, _
            "OpenNewComboRecord", _
            "The setup form '" & setupFormName & "' does not exist."
    End If

    If Not CloseLoadedSetupForm(setupFormName) Then
        Exit Sub
    End If

    oldMaximumID = Nz( _
        DMax(QuoteIdentifier(idFieldName), _
            QuoteIdentifier(tableName)), _
        0)

    suggestedText = Trim$(suggestedText)

    If Len(suggestedText) > 0 Then
        openArguments = "QUICKADD|" & suggestedText

        DoCmd.OpenForm _
            formName:=setupFormName, _
            View:=acNormal, _
            DataMode:=acFormAdd, _
            WindowMode:=acDialog, _
            OpenArgs:=openArguments
    Else
        DoCmd.OpenForm _
            formName:=setupFormName, _
            View:=acNormal, _
            DataMode:=acFormAdd, _
            WindowMode:=acDialog
    End If

    newMaximumID = Nz( _
        DMax(QuoteIdentifier(idFieldName), _
            QuoteIdentifier(tableName)), _
        0)

    If newMaximumID > oldMaximumID Then
        savedID = newMaximumID
    ElseIf Len(suggestedText) > 0 Then
        savedID = DLookup( _
            QuoteIdentifier(idFieldName), _
            QuoteIdentifier(tableName), _
            "Nz(" & QuoteIdentifier(matchingFieldName) & ",'')=" & _
                SqlText(suggestedText))
    End If

    comboControl.Requery

    If Not IsNull(savedID) Then
        If Nz(savedID, 0) > 0 Then
            comboControl.value = CLng(savedID)
        End If
    End If

Exit_Handler:
    Exit Sub

Err_Handler:
    If Err.Number = 2501 Then
        Resume Exit_Handler
    End If

    Err.Raise Err.Number, Err.Source, Err.Description

End Sub


Private Function CreateTypedComboEntry( _
    ByVal NewData As String, _
    ByVal setupFormName As String, _
    ByVal tableName As String, _
    ByVal idFieldName As String, _
    ByVal matchingFieldName As String _
) As Boolean

    On Error GoTo Err_Handler

    Dim enteredText As String
    Dim savedID As Variant

    CreateTypedComboEntry = False
    enteredText = Trim$(Nz(NewData, vbNullString))

    If Len(enteredText) = 0 Then
        Exit Function
    End If

    If MsgBox( _
        "'" & enteredText & "' is not currently in the list." & _
        vbCrLf & vbCrLf & _
        "Would you like to create it now?", _
        vbQuestion + vbYesNo, _
        "Add New Entry") <> vbYes Then

        Exit Function
    End If

    If Not FormExists(setupFormName) Then
        Err.Raise _
            vbObjectError + 7303, _
            "CreateTypedComboEntry", _
            "The setup form '" & setupFormName & "' does not exist."
    End If

    If Not CloseLoadedSetupForm(setupFormName) Then
        Exit Function
    End If

    DoCmd.OpenForm _
        formName:=setupFormName, _
        View:=acNormal, _
        DataMode:=acFormAdd, _
        WindowMode:=acDialog, _
        OpenArgs:="QUICKADD|" & enteredText

    savedID = DLookup( _
        QuoteIdentifier(idFieldName), _
        QuoteIdentifier(tableName), _
        "Nz(" & QuoteIdentifier(matchingFieldName) & ",'')=" & _
            SqlText(enteredText))

    If IsNull(savedID) Then
        MsgBox _
            "The new entry was not found after the setup form closed." & _
            vbCrLf & _
            "Save it using the same name that you typed.", _
            vbInformation, _
            "Add New Entry"
        Exit Function
    End If

    CreateTypedComboEntry = True

Exit_Handler:
    Exit Function

Err_Handler:
    If Err.Number <> 2501 Then
        MsgBox _
            "The new entry could not be created." & _
            vbCrLf & vbCrLf & _
            "Error Number: " & Err.Number & vbCrLf & _
            "Details: " & Err.Description, _
            vbExclamation, _
            "Add New Entry"
    End If

    Resume Exit_Handler

End Function


Private Function CloseLoadedSetupForm( _
    ByVal formName As String _
) As Boolean

    On Error GoTo Err_Handler

    CloseLoadedSetupForm = False

    If Not IsFormLoaded(formName) Then
        CloseLoadedSetupForm = True
        Exit Function
    End If

    If Forms(formName).Dirty Then
        Forms(formName).Dirty = False
    End If

    DoCmd.close acForm, formName, acSaveNo

    CloseLoadedSetupForm = Not IsFormLoaded(formName)

Exit_Handler:
    Exit Function

Err_Handler:
    MsgBox _
        "Finish or cancel the current record in " & formName & _
        ", then try again." & vbCrLf & vbCrLf & _
        "Details: " & Err.Description, _
        vbExclamation, _
        "Setup Form Already Open"
    Resume Exit_Handler

End Function


Private Sub OpenStandaloneForm(ByVal formName As String)

    If Not FormExists(formName) Then
        Err.Raise _
            vbObjectError + 7304, _
            "OpenStandaloneForm", _
            "The form '" & formName & "' does not exist."
    End If

    If IsFormLoaded(formName) Then
        DoCmd.SelectObject acForm, formName, False
    Else
        DoCmd.OpenForm formName, acNormal
    End If

End Sub


Private Function WarehouseAvailableForLocation() As Boolean

    WarehouseAvailableForLocation = _
        (DCount("*", "T_Warehouses") > 0)

    If WarehouseAvailableForLocation Then
        Exit Function
    End If

    MsgBox _
        "Create a warehouse before adding a stock location." & _
        vbCrLf & vbCrLf & _
        "The warehouse list will open now. After saving a warehouse, " & _
        "return here and open Location Setup again.", _
        vbInformation, _
        "Warehouse Required"

    If FormExists(WAREHOUSE_LIST_FORM) Then
        OpenStandaloneForm WAREHOUSE_LIST_FORM
    End If

End Function


'=============================================================
' GENERAL HELPERS
'=============================================================

Private Function CurrentProjectID() As Long

    On Error GoTo Err_Handler

    CurrentProjectID = Nz( _
        Me.Recordset.Fields("ProjectID").value, _
        0)

Exit_Handler:
    Exit Function

Err_Handler:
    CurrentProjectID = 0
    Resume Exit_Handler

End Function


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


Private Sub SetControlEnabled( _
    ByVal controlName As String, _
    ByVal enabledValue As Boolean _
)

    If ControlExists(controlName) Then
        Me.Controls(controlName).Enabled = enabledValue
    End If

End Sub


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


Private Function QuoteIdentifier( _
    ByVal identifierName As String _
) As String

    QuoteIdentifier = _
        "[" & Replace(identifierName, "]", "]]") & "]"

End Function


Private Function SqlText(ByVal value As String) As String

    SqlText = "'" & Replace(value, "'", "''") & "'"

End Function


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
