Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Compare Database
Option Explicit

'=============================================================
' FORM: F_StockOut
'
' PURPOSE:
'   Posts negative inventory movements:
'       - Issue to project
'       - Tool issue
'       - Sale
'       - Supplier return
'       - Damage write-off
'       - Negative adjustment
'       - Partial or complete project shipment
'
' TABLES UPDATED:
'   T_StockBalances   - bulk items only
'   T_ItemInstances   - serialized items only
'   T_StockMovements  - permanent audit trail
'
' SERIAL-NUMBER RULE:
'   RequiresSerial is stored directly on T_Items.
'
' TOOL RETURN NOTIFICATION RULE:
'   ExpectedReturnDate is optional.
'   It is only saved for returnable stock-out reasons such as:
'       - Issue to Project
'       - Tool Issue
'       - Project Tool Issue
'       - Issued to Project
'
' IMPORTANT:
'   No IsActive fields are used.
'   No legacy finished-item field is used. Shipment derives its item
'   and available quantity from Production Receipt movements.
'=============================================================

Private Const RECENT_SUBFORM_CONTROL As String = "SF_RecentStockOut"
Private Const PROJECT_FORM As String = "F_Projects"

Private Const CONTEXT_PROJECT_SHIPMENT As String = "ProjectShipment"
Private Const REASON_PRODUCTION_RECEIPT As String = "Production Receipt"
Private Const REASON_PROJECT_SHIPMENT As String = "Project Shipment"
Private Const SOURCE_PROJECT As String = "Project"
Private Const QUANTITY_TOLERANCE As Double = 0.000001

Private mProjectShipmentMode As Boolean
Private mProjectID As Long
Private mProjectOutputItemID As Long
Private mProjectOutputLocationID As Long
Private mProjectCustomerID As Long
Private mProjectAssemblyTemplateID As Long
Private mProjectSalesOrderID As Long
Private mProjectBuildQuantity As Double
Private mProjectReceivedQuantity As Double
Private mProjectShippedQuantity As Double
Private mProjectCode As String


'=============================================================
' INITIALIZE FORM
'=============================================================
Private Sub Form_Load()

    ResetForm
    ApplyStockOutOpenArgs

End Sub


'=============================================================
' PROJECT SHIPMENT HANDOFF
'=============================================================

Private Sub ApplyStockOutOpenArgs()

    Dim contextText As String
    Dim ProjectID As Long

    contextText = GetStockOutOpenArgValue("Context")

    If StrComp( _
        contextText, _
        CONTEXT_PROJECT_SHIPMENT, _
        vbTextCompare) <> 0 Then

        Exit Sub
    End If

    ProjectID = CLng(Val(GetStockOutOpenArgValue("ProjectID")))

    If ProjectID <= 0 Then
        MsgBox _
            "F_StockOut was opened for a project shipment " & _
            "without a valid ProjectID.", _
            vbExclamation, _
            "Project Shipment"
        Exit Sub
    End If

    LoadProjectShipment ProjectID

End Sub


Public Sub LoadProjectShipment(ByVal ProjectID As Long)

    On Error GoTo Err_Handler

    Dim db As DAO.Database
    Dim rs As DAO.Recordset
    Dim projectStatus As String
    Dim unshippedQuantity As Double
    Dim CustomerName As String

    If ProjectID <= 0 Then
        Err.Raise _
            vbObjectError + 8201, _
            "LoadProjectShipment", _
            "A valid ProjectID is required."
    End If

    ResetForm

    Set db = CurrentDb
    Set rs = db.OpenRecordset( _
        "SELECT * FROM T_Projects " & _
        "WHERE ProjectID=" & ProjectID & ";", _
        dbOpenSnapshot)

    If rs.EOF Then
        Err.Raise _
            vbObjectError + 8202, _
            "LoadProjectShipment", _
            "Project ID " & ProjectID & " was not found."
    End If

    mProjectID = ProjectID
    mProjectCode = Trim$(Nz(rs!ProjectCode, vbNullString))
    mProjectCustomerID = Nz(rs!CustomerID, 0)
    mProjectAssemblyTemplateID = Nz(rs!AssemblyTemplateID, 0)
    mProjectOutputLocationID = Nz(rs!OutputLocationID, 0)
    mProjectBuildQuantity = Nz(rs!BuildQuantity, 0)
    projectStatus = UCase$(Trim$(Nz(rs!Status, vbNullString)))

    mProjectSalesOrderID = _
        ResolveLinkedSalesOrderID(db, rs, mProjectID)

    rs.close
    Set rs = Nothing

    If projectStatus = "CANCELLED" Then
        Err.Raise _
            vbObjectError + 8203, _
            "LoadProjectShipment", _
            "A cancelled project cannot be shipped."
    End If

    If projectStatus = "SHIPPED" Then
        Err.Raise _
            vbObjectError + 8204, _
            "LoadProjectShipment", _
            "This project is already fully shipped."
    End If

    If projectStatus <> "READY TO SHIP" Then
        Err.Raise _
            vbObjectError + 8205, _
            "LoadProjectShipment", _
            "Mark the project Ready to Ship first."
    End If

    GetProjectOutputSummary _
        db, _
        mProjectID, _
        mProjectOutputItemID, _
        mProjectReceivedQuantity, _
        mProjectShippedQuantity

    If mProjectOutputItemID <= 0 Then
        Err.Raise _
            vbObjectError + 8206, _
            "LoadProjectShipment", _
            "No production receipt exists for this project."
    End If

    unshippedQuantity = ProjectUnshippedQuantity()

    If unshippedQuantity <= QUANTITY_TOLERANCE Then
        Err.Raise _
            vbObjectError + 8207, _
            "LoadProjectShipment", _
            "There is no received project output waiting to ship."
    End If

    mProjectShipmentMode = True

    EnsureValueListItem Me.cboReason, REASON_PROJECT_SHIPMENT

    Me.cboReason = REASON_PROJECT_SHIPMENT
    Me.cboReason.Enabled = True
    Me.cboReason.Locked = True

    Me.cboProjectID.Requery
    Me.cboProjectID = mProjectID
    Me.cboProjectID.Enabled = True
    Me.cboProjectID.Locked = True

    Me.cboItemID.Requery
    Me.cboItemID = mProjectOutputItemID
    Me.cboItemID.Locked = True

    Me.txtReferenceDocument = mProjectCode
    Me.txtExpectedReturnDate = Null
    Me.txtExpectedReturnDate.Enabled = False
    Me.txtExpectedReturnDate.Locked = True

    CustomerName = vbNullString

    If mProjectCustomerID > 0 Then
        CustomerName = Nz(DLookup( _
            "CustomerName", _
            "T_Customers", _
            "CustomerID=" & mProjectCustomerID), vbNullString)
    End If

    Me.txtIssuedTo = CustomerName
    Me.Caption = "Project Shipment - " & mProjectCode

    LoadSelectedItem
    ApplyProjectShipmentControlValues

    On Error Resume Next

    If ItemRequiresSerial(mProjectOutputItemID) Then
        Me.cboItemInstanceID.SetFocus
    Else
        Me.txtQuantity.SetFocus
    End If

    On Error GoTo 0

Exit_Handler:
    On Error Resume Next
    If Not (rs Is Nothing) Then rs.close
    Set rs = Nothing
    Set db = Nothing
    On Error GoTo 0
    Exit Sub

Err_Handler:
    MsgBox _
        "The project shipment could not be prepared." & _
        vbCrLf & vbCrLf & _
        "Error Number: " & Err.Number & vbCrLf & _
        "Details: " & Err.Description, _
        vbExclamation, _
        "Project Shipment"
    Resume Exit_Handler

End Sub


Private Function GetStockOutOpenArgValue( _
    ByVal keyName As String _
) As String

    Dim rawArguments As String
    Dim argumentParts() As String
    Dim onePart As Variant
    Dim equalsPosition As Long
    Dim thisKey As String

    rawArguments = Nz(Me.OpenArgs, vbNullString)

    If Len(Trim$(rawArguments)) = 0 Then
        Exit Function
    End If

    argumentParts = Split(rawArguments, ";")

    For Each onePart In argumentParts
        equalsPosition = InStr(1, CStr(onePart), "=", vbBinaryCompare)

        If equalsPosition > 0 Then
            thisKey = Left$(CStr(onePart), equalsPosition - 1)

            If StrComp(thisKey, keyName, vbTextCompare) = 0 Then
                GetStockOutOpenArgValue = _
                    Mid$(CStr(onePart), equalsPosition + 1)
                Exit Function
            End If
        End If
    Next onePart

End Function


'=============================================================
' ITEM SELECTED
'=============================================================
Private Sub cboItemID_AfterUpdate()

    LoadSelectedItem

End Sub


'=============================================================
' SERIAL NUMBER SELECTED
'=============================================================
Private Sub cboItemInstanceID_AfterUpdate()

    LoadSelectedSerializedInstance

End Sub


'=============================================================
' REASON SELECTED
'=============================================================
Private Sub cboReason_AfterUpdate()

    ConfigureProjectControl
    ConfigureExpectedReturnControl

End Sub


'=============================================================
' POST STOCK-OUT TRANSACTION
'=============================================================
Private Sub cmdPostStockOut_Click()
    gSkipInventory = UX_ProjectSkipsInventory(Nz(mProjectID, 0))

    On Error GoTo Err_Handler

    Dim ws As DAO.Workspace
    Dim db As DAO.Database

    Dim ItemID As Long
    Dim ItemInstanceID As Long
    Dim FromLocationID As Long
    Dim performedByUserID As Long
    Dim MovementTypeID As Long
    Dim ProjectID As Long

    Dim Quantity As Double
    Dim MovementDate As Date
    Dim ExpectedReturnDate As Variant
    Dim RequiresSerial As Boolean

    Dim Reason As String
    Dim ReferenceDocument As String
    Dim MovementNotes As String

    Dim UnitCost As Variant
    Dim transactionStarted As Boolean

    Dim originalErrorNumber As Long
    Dim originalErrorDescription As String
    Dim originalErrorSource As String

    '---------------------------------------------------------
    ' Validate all values before changing any table.
    '---------------------------------------------------------
    If Not ValidateEntry( _
        ItemID, _
        ItemInstanceID, _
        FromLocationID, _
        performedByUserID, _
        ProjectID, _
        Quantity, _
        MovementDate, _
        ExpectedReturnDate, _
        RequiresSerial, _
        Reason, _
        ReferenceDocument, _
        MovementNotes, _
        UnitCost _
    ) Then
        Exit Sub
    End If

    Set ws = DBEngine.Workspaces(0)
    Set db = CurrentDb

    MovementTypeID = GetMovementTypeID(db, "STOCK_OUT")

    '---------------------------------------------------------
    ' Begin transaction.
    ' Every table update succeeds or every update rolls back.
    '---------------------------------------------------------
    ws.BeginTrans
    transactionStarted = True

    If RequiresSerial Then

        IssueSerializedInstance _
            db, _
            ItemID, _
            ItemInstanceID, _
            FromLocationID, _
            Reason, _
            MovementDate, _
            MovementNotes

    Else

        ReduceBulkStockBalance _
            db, _
            ItemID, _
            FromLocationID, _
            Quantity

    End If

    '---------------------------------------------------------
    ' Create permanent audit record.
    '---------------------------------------------------------
    InsertStockMovement _
        db, _
        MovementDate, _
        ExpectedReturnDate, _
        ItemID, _
        ItemInstanceID, _
        MovementTypeID, _
        Quantity, _
        FromLocationID, _
        performedByUserID, _
        ProjectID, _
        Reason, _
        ReferenceDocument, _
        MovementNotes, _
        UnitCost

    If mProjectShipmentMode Then
        If mProjectSalesOrderID > 0 Then
            ApplyShipmentToLinkedSalesOrder _
                db, _
                mProjectSalesOrderID, _
                ItemID, _
                Quantity, _
                MovementDate
        End If

        UpdateProjectShipmentStatus db, Quantity, MovementDate
    End If

    ws.CommitTrans
    transactionStarted = False

    '---------------------------------------------------------
    ' Refresh display after committed transaction.
    '---------------------------------------------------------
    On Error Resume Next

    RefreshSelectedItemStatus
    RequeryRecentSubform

    If mProjectShipmentMode Then
        mProjectShippedQuantity = _
            mProjectShippedQuantity + Quantity
        PrepareNextProjectShipment RequiresSerial
        RefreshProjectFormIfOpen
    Else
        PrepareNextEntry RequiresSerial
    End If

    On Error GoTo 0

    If mProjectShipmentMode Then
        MsgBox _
            "Project shipment posted successfully." & _
            vbCrLf & vbCrLf & _
            "Shipped for project: " & _
                Format$(mProjectShippedQuantity, "0.######") & _
            vbCrLf & _
            "Build quantity: " & _
                Format$(mProjectBuildQuantity, "0.######"), _
            vbInformation, _
            "Project Shipment Completed"

        If ProjectUnshippedQuantity() <= QUANTITY_TOLERANCE Then
            RefreshProjectFormIfOpen
            On Error Resume Next
            DoCmd.close acForm, Me.Name, acSaveNo
            On Error GoTo 0
        End If
    Else
        MsgBox _
            "Stock-out transaction posted successfully.", _
            vbInformation, _
            "Stock Out Completed"
    End If

Exit_Handler:

    Set db = Nothing
    Set ws = Nothing

    Exit Sub

Err_Handler:

    originalErrorNumber = Err.Number
    originalErrorDescription = Err.Description
    originalErrorSource = Err.Source

    On Error Resume Next

    If transactionStarted Then

        If Not (ws Is Nothing) Then
            ws.Rollback
        End If

    End If

    On Error GoTo 0

    MsgBox _
        "The stock-out transaction was not posted." & vbCrLf & vbCrLf & _
        "Error Number: " & originalErrorNumber & vbCrLf & _
        "Source: " & originalErrorSource & vbCrLf & _
        "Details: " & originalErrorDescription, _
        vbExclamation, _
        "Stock Out Failed"

    Resume Exit_Handler

End Sub


'=============================================================
' CLEAR FORM
'=============================================================
Private Sub cmdClear_Click()

    Dim ProjectID As Long

    If mProjectShipmentMode Then
        ProjectID = mProjectID
        LoadProjectShipment ProjectID
    Else
        ResetForm
    End If

End Sub


'=============================================================
' CLOSE FORM
'=============================================================
Private Sub cmdClose_Click()

    If mProjectShipmentMode Then
        RefreshProjectFormIfOpen
    End If

    DoCmd.close acForm, Me.Name

End Sub


'=============================================================
' VALIDATE ENTRY
'=============================================================
Private Function ValidateEntry( _
    ByRef ItemID As Long, _
    ByRef ItemInstanceID As Long, _
    ByRef FromLocationID As Long, _
    ByRef performedByUserID As Long, _
    ByRef ProjectID As Long, _
    ByRef Quantity As Double, _
    ByRef MovementDate As Date, _
    ByRef ExpectedReturnDate As Variant, _
    ByRef RequiresSerial As Boolean, _
    ByRef Reason As String, _
    ByRef ReferenceDocument As String, _
    ByRef MovementNotes As String, _
    ByRef UnitCost As Variant _
) As Boolean

    Dim issuedTo As String
    Dim Notes As String

    ValidateEntry = False
    ExpectedReturnDate = Null

    '---------------------------------------------------------
    ' Item
    '---------------------------------------------------------
    If IsNull(Me.cboItemID) Then

        MsgBox _
            "Select the item being issued.", _
            vbExclamation, _
            "Item Required"

        Me.cboItemID.SetFocus
        Exit Function

    End If

    ItemID = CLng(Me.cboItemID)

    RequiresSerial = ItemRequiresSerial(ItemID)

    Me.chkRequiresSerial = RequiresSerial

    '---------------------------------------------------------
    ' Movement date
    '---------------------------------------------------------
    If IsNull(Me.txtMovementDate) Or _
       Not IsDate(Me.txtMovementDate) Then

        MsgBox _
            "Enter a valid movement date.", _
            vbExclamation, _
            "Movement Date Required"

        Me.txtMovementDate.SetFocus
        Exit Function

    End If

    MovementDate = CDate(Me.txtMovementDate)

    '---------------------------------------------------------
    ' Reason
    '---------------------------------------------------------
    Reason = Trim$(Nz(Me.cboReason, vbNullString))

    If Len(Reason) = 0 Then

        MsgBox _
            "Select the reason for the stock-out transaction.", _
            vbExclamation, _
            "Reason Required"

        Me.cboReason.SetFocus
        Exit Function

    End If

    If mProjectShipmentMode And _
       StrComp( _
            Reason, _
            REASON_PROJECT_SHIPMENT, _
            vbTextCompare) <> 0 Then

        MsgBox _
            "Project output must be shipped with Reason = " & _
            REASON_PROJECT_SHIPMENT & ".", _
            vbExclamation, _
            "Project Shipment Reason"

        Me.cboReason = REASON_PROJECT_SHIPMENT
        Exit Function
    End If

    If Not mProjectShipmentMode And _
       StrComp( _
            Reason, _
            REASON_PROJECT_SHIPMENT, _
            vbTextCompare) = 0 Then

        MsgBox _
            "Open Stock Out from the project's Post Shipment " & _
            "button when posting a Project Shipment.", _
            vbExclamation, _
            "Project Context Required"
        Exit Function
    End If

    '---------------------------------------------------------
    ' Expected return date.
    '
    ' Optional, but only relevant for returnable issue reasons.
    ' If entered, it must be a valid date and cannot be before
    ' the movement date.
    '---------------------------------------------------------
    If IsReturnableIssueReason(Reason) Then

        If Len(Trim$(Nz(Me.txtExpectedReturnDate, vbNullString))) > 0 Then

            If Not IsDate(Me.txtExpectedReturnDate) Then

                MsgBox _
                    "Enter a valid expected return date, or leave it blank.", _
                    vbExclamation, _
                    "Invalid Expected Return Date"

                Me.txtExpectedReturnDate.SetFocus
                Exit Function

            End If

            ExpectedReturnDate = CDate(Me.txtExpectedReturnDate)

            If DateValue(CDate(ExpectedReturnDate)) < DateValue(MovementDate) Then

                MsgBox _
                    "Expected return date cannot be earlier than the movement date.", _
                    vbExclamation, _
                    "Invalid Expected Return Date"

                Me.txtExpectedReturnDate.SetFocus
                Exit Function

            End If

        End If

    Else

        Me.txtExpectedReturnDate = Null
        ExpectedReturnDate = Null

    End If

    '---------------------------------------------------------
    ' Project:
    ' Mandatory only for project issue reasons.
    '---------------------------------------------------------
    ProjectID = 0

    If IsProjectIssueReason(Reason) Then

        If IsNull(Me.cboProjectID) Then

            MsgBox _
                "Select the project receiving the stock.", _
                vbExclamation, _
                "Project Required"

            Me.cboProjectID.SetFocus
            Exit Function

        End If

        ProjectID = CLng(Me.cboProjectID)

    End If

    '---------------------------------------------------------
    ' Issued by
    '---------------------------------------------------------
    If IsNull(Me.cboPerformedByUserID) Then

        MsgBox _
            "Select the user issuing or recording the stock.", _
            vbExclamation, _
            "Issued By Required"

        Me.cboPerformedByUserID.SetFocus
        Exit Function

    End If

    performedByUserID = CLng(Me.cboPerformedByUserID)

    '---------------------------------------------------------
    ' Optional text fields
    '---------------------------------------------------------
    ReferenceDocument = Trim$(Nz( _
        Me.txtReferenceDocument, _
        vbNullString _
    ))

    issuedTo = Trim$(Nz(Me.txtIssuedTo, vbNullString))
    Notes = Trim$(Nz(Me.txtNOtes, vbNullString))

    '---------------------------------------------------------
    ' Recipient remains mandatory for direct tool issues and sales.
    '
    ' For project issues, ProjectID is the required reference.
    ' A recipient can still be entered optionally.
    '---------------------------------------------------------
    If UCase$(Reason) = "TOOL ISSUE" Or _
       UCase$(Reason) = "SALE" Then

        If Len(issuedTo) = 0 Then

            MsgBox _
                "Enter the recipient, department, or customer.", _
                vbExclamation, _
                "Recipient Required"

            Me.txtIssuedTo.SetFocus
            Exit Function

        End If

    End If

    MovementNotes = BuildMovementNotes(issuedTo, Notes)

    '---------------------------------------------------------
    ' Serialized item
    '---------------------------------------------------------
    If RequiresSerial Then

        Quantity = 1
        Me.txtQuantity = 1

        If IsNull(Me.cboItemInstanceID) Then

            MsgBox _
                "Select the serial number being issued.", _
                vbExclamation, _
                "Serial Number Required"

            Me.cboItemInstanceID.SetFocus
            Exit Function

        End If

        ItemInstanceID = CLng(Me.cboItemInstanceID)

        If Not ValidateSerializedInstance( _
            ItemID, _
            ItemInstanceID, _
            FromLocationID, _
            UnitCost _
        ) Then
            Exit Function
        End If

    Else

        '-----------------------------------------------------
        ' Bulk item
        '-----------------------------------------------------
        ItemInstanceID = 0
        UnitCost = Null

        If IsNull(Me.txtQuantity) Or _
           Len(Trim$(Nz(Me.txtQuantity, vbNullString))) = 0 Then

            MsgBox _
                "Enter the quantity being issued.", _
                vbExclamation, _
                "Quantity Required"

            Me.txtQuantity.SetFocus
            Exit Function

        End If

        If Not IsNumeric(Me.txtQuantity) Then

            MsgBox _
                "Quantity must be numeric.", _
                vbExclamation, _
                "Invalid Quantity"

            Me.txtQuantity.SetFocus
            Exit Function

        End If

        Quantity = CDbl(Me.txtQuantity)

        If Quantity <= 0 Then

            MsgBox _
                "Quantity must be greater than zero.", _
                vbExclamation, _
                "Invalid Quantity"

            Me.txtQuantity.SetFocus
            Exit Function

        End If

        If Not ValidateBulkQuantity( _
            ItemID, _
            Quantity, _
            FromLocationID _
        ) Then
            Exit Function
        End If

    End If

    If mProjectShipmentMode Then
        If Not ValidateProjectShipmentEntry( _
            ItemID, _
            ItemInstanceID, _
            ProjectID, _
            Quantity) Then

            Exit Function
        End If
    End If

    Me.txtFromLocationID = FromLocationID

    ValidateEntry = True

End Function


'=============================================================
' VALIDATE BULK QUANTITY
'=============================================================
Private Function ValidateBulkQuantity( _
    ByVal ItemID As Long, _
    ByVal requestedQuantity As Double, _
    ByRef FromLocationID As Long _
) As Boolean

    Dim db As DAO.Database
    Dim rs As DAO.Recordset

    Dim QuantityOnHand As Double
    Dim QuantityReserved As Double
    Dim quantityAvailable As Double

    ValidateBulkQuantity = False

    Set db = CurrentDb

    Set rs = db.OpenRecordset( _
        "SELECT " & _
            "CurrentLocationID, " & _
            "QuantityOnHand, " & _
            "QuantityReserved " & _
        "FROM T_StockBalances " & _
        "WHERE ItemID = " & ItemID & ";", _
        dbOpenSnapshot _
    )

    If rs.EOF Then

        rs.close
        Set rs = Nothing
        Set db = Nothing

        MsgBox _
            "This bulk item does not have a stock balance.", _
            vbExclamation, _
            "No Stock Available"

        Exit Function

    End If

    rs.MoveLast

    If rs.recordCount > 1 Then

        rs.close
        Set rs = Nothing
        Set db = Nothing

        MsgBox _
            "More than one stock-balance row exists for this item.", _
            vbExclamation, _
            "Stock-Balance Data Issue"

        Exit Function

    End If

    rs.MoveFirst

    FromLocationID = Nz(rs!CurrentLocationID, 0)
    QuantityOnHand = Nz(rs!QuantityOnHand, 0)
    QuantityReserved = Nz(rs!QuantityReserved, 0)

    quantityAvailable = QuantityOnHand - QuantityReserved

    rs.close
    Set rs = Nothing
    Set db = Nothing

    If FromLocationID = 0 Then

        MsgBox _
            "This item has stock but no source location.", _
            vbExclamation, _
            "Location Required"

        Exit Function

    End If

    If quantityAvailable <= 0 Then

        MsgBox _
            "No unreserved stock is currently available.", _
            vbExclamation, _
            "No Available Stock"

        Exit Function

    End If

    If requestedQuantity > quantityAvailable Then

        MsgBox _
            "The requested quantity exceeds available stock." & _
            vbCrLf & vbCrLf & _
            "Quantity on hand: " & QuantityOnHand & vbCrLf & _
            "Quantity reserved: " & QuantityReserved & vbCrLf & _
            "Available quantity: " & quantityAvailable & vbCrLf & _
            "Requested quantity: " & requestedQuantity, _
            vbExclamation, _
            "Insufficient Stock"

        Exit Function

    End If

    ValidateBulkQuantity = True

End Function


'=============================================================
' VALIDATE SERIALIZED ITEM
'=============================================================
Private Function ValidateSerializedInstance( _
    ByVal ItemID As Long, _
    ByVal ItemInstanceID As Long, _
    ByRef FromLocationID As Long, _
    ByRef UnitCost As Variant _
) As Boolean

    Dim db As DAO.Database
    Dim rs As DAO.Recordset

    ValidateSerializedInstance = False

    Set db = CurrentDb

    Set rs = db.OpenRecordset( _
        "SELECT " & _
            "ItemID, " & _
            "CurrentLocationID, " & _
            "CurrentStatus, " & _
            "UnitCost " & _
        "FROM T_ItemInstances " & _
        "WHERE ItemInstanceID = " & ItemInstanceID & ";", _
        dbOpenSnapshot _
    )

    If rs.EOF Then

        rs.close
        Set rs = Nothing
        Set db = Nothing

        MsgBox _
            "The selected serial number could not be found.", _
            vbExclamation, _
            "Serial Number Not Found"

        Exit Function

    End If

    If Nz(rs!ItemID, 0) <> ItemID Then

        rs.close
        Set rs = Nothing
        Set db = Nothing

        MsgBox _
            "The selected serial number does not belong to this item.", _
            vbExclamation, _
            "Serial Number Mismatch"

        Exit Function

    End If

    FromLocationID = Nz(rs!CurrentLocationID, 0)

    If FromLocationID = 0 Then

        rs.close
        Set rs = Nothing
        Set db = Nothing

        MsgBox _
            "The selected serial number is not stored in a warehouse.", _
            vbExclamation, _
            "Serialized Item Unavailable"

        Exit Function

    End If

    If Nz(rs!currentStatus, vbNullString) <> "In Stock" Then

        rs.close
        Set rs = Nothing
        Set db = Nothing

        MsgBox _
            "The selected serial number is not marked as In Stock.", _
            vbExclamation, _
            "Serialized Item Unavailable"

        Exit Function

    End If

    If IsNull(rs!UnitCost) Then
        UnitCost = Null
    Else
        UnitCost = rs!UnitCost
    End If

    rs.close
    Set rs = Nothing
    Set db = Nothing

    ValidateSerializedInstance = True

End Function


'=============================================================
' REDUCE BULK STOCK BALANCE
'=============================================================
Private Sub ReduceBulkStockBalance( _
    ByVal db As DAO.Database, _
    ByVal ItemID As Long, _
    ByVal FromLocationID As Long, _
    ByVal Quantity As Double _
)
    If gSkipInventory Then Exit Sub

    Dim rs As DAO.Recordset

    Dim existingLocationID As Long
    Dim QuantityOnHand As Double
    Dim QuantityReserved As Double
    Dim quantityAvailable As Double

    Set rs = db.OpenRecordset( _
        "SELECT * " & _
        "FROM T_StockBalances " & _
        "WHERE ItemID = " & ItemID & ";", _
        dbOpenDynaset _
    )

    If rs.EOF Then

        rs.close
        Set rs = Nothing

        Err.Raise _
            vbObjectError + 2001, _
            "ReduceBulkStockBalance", _
            "This bulk item does not have a stock-balance row."

    End If

    rs.MoveLast

    If rs.recordCount > 1 Then

        rs.close
        Set rs = Nothing

        Err.Raise _
            vbObjectError + 2002, _
            "ReduceBulkStockBalance", _
            "More than one stock-balance row exists for this item."

    End If

    rs.MoveFirst

    existingLocationID = Nz(rs!CurrentLocationID, 0)
    QuantityOnHand = Nz(rs!QuantityOnHand, 0)
    QuantityReserved = Nz(rs!QuantityReserved, 0)

    quantityAvailable = QuantityOnHand - QuantityReserved

    If existingLocationID <> FromLocationID Then

        rs.close
        Set rs = Nothing

        Err.Raise _
            vbObjectError + 2003, _
            "ReduceBulkStockBalance", _
            "The source location changed before posting."

    End If

    If Quantity > quantityAvailable Then

        rs.close
        Set rs = Nothing

        Err.Raise _
            vbObjectError + 2004, _
            "ReduceBulkStockBalance", _
            "The requested quantity now exceeds available stock."

    End If

    rs.Edit

    rs!QuantityOnHand = QuantityOnHand - Quantity
    rs!LastUpdated = Now()

    rs.Update

    rs.close
    Set rs = Nothing

End Sub


'=============================================================
' ISSUE SERIALIZED ITEM
'=============================================================
Private Sub IssueSerializedInstance( _
    ByVal db As DAO.Database, _
    ByVal ItemID As Long, _
    ByVal ItemInstanceID As Long, _
    ByVal FromLocationID As Long, _
    ByVal Reason As String, _
    ByVal MovementDate As Date, _
    ByVal MovementNotes As String _
)
    If gSkipInventory Then Exit Sub

    Dim rs As DAO.Recordset
    Dim newStatus As String
    Dim instanceNote As String

    newStatus = GetIssuedInstanceStatus(Reason)

    Set rs = db.OpenRecordset( _
        "SELECT * " & _
        "FROM T_ItemInstances " & _
        "WHERE ItemInstanceID = " & ItemInstanceID & ";", _
        dbOpenDynaset _
    )

    If rs.EOF Then

        rs.close
        Set rs = Nothing

        Err.Raise _
            vbObjectError + 2101, _
            "IssueSerializedInstance", _
            "The selected serialized item could not be found."

    End If

    If Nz(rs!ItemID, 0) <> ItemID Then

        rs.close
        Set rs = Nothing

        Err.Raise _
            vbObjectError + 2102, _
            "IssueSerializedInstance", _
            "The selected serial number does not belong to this item."

    End If

    If Nz(rs!CurrentLocationID, 0) <> FromLocationID Then

        rs.close
        Set rs = Nothing

        Err.Raise _
            vbObjectError + 2103, _
            "IssueSerializedInstance", _
            "The serialized-item location changed before posting."

    End If

    If Nz(rs!currentStatus, vbNullString) <> "In Stock" Then

        rs.close
        Set rs = Nothing

        Err.Raise _
            vbObjectError + 2104, _
            "IssueSerializedInstance", _
            "The serialized item is no longer marked as In Stock."

    End If

    instanceNote = _
        "Stock out posted on " & _
        Format$(MovementDate, "yyyy-mm-dd hh:nn") & _
        ". Status changed to: " & newStatus & "."

    If Len(MovementNotes) > 0 Then
        instanceNote = instanceNote & vbCrLf & MovementNotes
    End If

    rs.Edit

    rs!CurrentLocationID = Null
    rs!currentStatus = newStatus
    rs!Notes = AppendNote( _
        Nz(rs!Notes, vbNullString), _
        instanceNote _
    )

    rs.Update

    rs.close
    Set rs = Nothing

End Sub


'=============================================================
' INSERT STOCK MOVEMENT
'=============================================================
Private Sub InsertStockMovement( _
    ByVal db As DAO.Database, _
    ByVal MovementDate As Date, _
    ByVal ExpectedReturnDate As Variant, _
    ByVal ItemID As Long, _
    ByVal ItemInstanceID As Long, _
    ByVal MovementTypeID As Long, _
    ByVal Quantity As Double, _
    ByVal FromLocationID As Long, _
    ByVal performedByUserID As Long, _
    ByVal ProjectID As Long, _
    ByVal Reason As String, _
    ByVal ReferenceDocument As String, _
    ByVal Notes As String, _
    ByVal UnitCost As Variant _
)
    If gSkipInventory Then Exit Sub

    Dim rs As DAO.Recordset

    Set rs = db.OpenRecordset( _
        "T_StockMovements", _
        dbOpenDynaset _
    )

    rs.addnew

    rs!MovementDate = MovementDate
    rs!ItemID = ItemID

    If Not IsNull(ExpectedReturnDate) Then
        rs!ExpectedReturnDate = CDate(ExpectedReturnDate)
    End If

    If ItemInstanceID > 0 Then
        rs!ItemInstanceID = ItemInstanceID
    End If

    If ProjectID > 0 Then
        rs!ProjectID = ProjectID
    End If

    rs!MovementTypeID = MovementTypeID
    rs!Quantity = Quantity
    rs!FromLocationID = FromLocationID
    rs!performedByUserID = performedByUserID
    rs!Reason = Reason

    If mProjectShipmentMode Then
        SetRecordsetFieldIfExists _
            rs, _
            "CustomerID", _
            NullIfZero(mProjectCustomerID)

        SetRecordsetFieldIfExists _
            rs, _
            "AssemblyTemplateID", _
            NullIfZero(mProjectAssemblyTemplateID)

        SetRecordsetFieldIfExists _
            rs, _
            "SalesOrderID", _
            NullIfZero(mProjectSalesOrderID)

        SetRecordsetFieldIfExists _
            rs, _
            "SourceDocumentType", _
            SOURCE_PROJECT

        SetRecordsetFieldIfExists _
            rs, _
            "MovementNotes", _
            "Finished output shipped for project " & mProjectCode & "."
    End If

    If Len(ReferenceDocument) > 0 Then
        rs!ReferenceDocument = ReferenceDocument
    End If

    If Len(Notes) > 0 Then
        rs!Notes = Notes
    End If

    If Not IsNull(UnitCost) Then
        rs!UnitCost = UnitCost
    End If

    rs.Update

    rs.close
    Set rs = Nothing

End Sub


'=============================================================
' RETRIEVE MOVEMENT TYPE
'=============================================================
Private Function GetMovementTypeID( _
    ByVal db As DAO.Database, _
    ByVal MovementTypeCode As String _
) As Long

    Dim rs As DAO.Recordset

    Set rs = db.OpenRecordset( _
        "SELECT MovementTypeID " & _
        "FROM T_StockMovementTypes " & _
        "WHERE MovementTypeCode = " & _
            SqlText(MovementTypeCode) & ";", _
        dbOpenSnapshot _
    )

    If rs.EOF Then

        rs.close
        Set rs = Nothing

        Err.Raise _
            vbObjectError + 2201, _
            "GetMovementTypeID", _
            "Movement type '" & MovementTypeCode & "' was not found."

    End If

    GetMovementTypeID = CLng(rs!MovementTypeID)

    rs.close
    Set rs = Nothing

End Function


'=============================================================
' LOAD SELECTED ITEM
'=============================================================
Private Sub LoadSelectedItem()

    Dim ItemID As Long
    Dim RequiresSerial As Boolean

    If IsNull(Me.cboItemID) Then

        ResetItemStatusDisplay

        If mProjectShipmentMode Then
            ApplyProjectShipmentControlValues
        End If

        Exit Sub

    End If

    ItemID = CLng(Me.cboItemID)

    RequiresSerial = ItemRequiresSerial(ItemID)

    Me.chkRequiresSerial = RequiresSerial

    ConfigureSerialControls RequiresSerial

    RefreshSelectedItemStatus

    If mProjectShipmentMode Then
        ApplyProjectShipmentControlValues
    End If

End Sub


'=============================================================
' SHOW CURRENT STOCK STATUS
'=============================================================
Private Sub RefreshSelectedItemStatus()

    On Error GoTo Err_Handler

    Dim db As DAO.Database
    Dim rs As DAO.Recordset

    Dim ItemID As Long
    Dim RequiresSerial As Boolean

    Dim FromLocationID As Long
    Dim QuantityOnHand As Double
    Dim QuantityReserved As Double
    Dim quantityAvailable As Double

    If IsNull(Me.cboItemID) Then

        ResetItemStatusDisplay
        Exit Sub

    End If

    ItemID = CLng(Me.cboItemID)

    RequiresSerial = ItemRequiresSerial(ItemID)

    Me.chkRequiresSerial = RequiresSerial

    If RequiresSerial Then

        Me.txtAvailableQuantity = DCount( _
            "*", _
            "T_ItemInstances", _
            "ItemID = " & ItemID & _
            " AND CurrentLocationID Is Not Null" & _
            " AND CurrentStatus = 'In Stock'" _
        )

        Me.txtCurrentLocation = _
            "Select an available serial number"

        Me.txtShelfLocation = _
            "Select an available serial number"

        Me.txtFromLocationID = Null

        PopulateSerialNumberCombo ItemID

        If mProjectShipmentMode Then
            Me.txtAvailableQuantity = _
                Me.cboItemInstanceID.ListCount
        End If

    Else

        Me.cboItemInstanceID = Null
        Me.cboItemInstanceID.RowSource = vbNullString

        Set db = CurrentDb

        Set rs = db.OpenRecordset( _
            "SELECT " & _
                "CurrentLocationID, " & _
                "QuantityOnHand, " & _
                "QuantityReserved " & _
            "FROM T_StockBalances " & _
            "WHERE ItemID = " & ItemID & ";", _
            dbOpenSnapshot _
        )

        If rs.EOF Then

            Me.txtAvailableQuantity = 0
            Me.txtCurrentLocation = "No stock recorded"
            Me.txtShelfLocation = "No shelf location recorded"
            Me.txtFromLocationID = Null

        Else

            rs.MoveLast

            If rs.recordCount > 1 Then

                Me.txtAvailableQuantity = Null
                Me.txtCurrentLocation = _
                    "Data issue: multiple stock-balance rows found"
                Me.txtShelfLocation = _
                    "Data issue: multiple stock-balance rows found"
                Me.txtFromLocationID = Null

            Else

                rs.MoveFirst

                FromLocationID = Nz(rs!CurrentLocationID, 0)
                QuantityOnHand = Nz(rs!QuantityOnHand, 0)
                QuantityReserved = Nz(rs!QuantityReserved, 0)

                quantityAvailable = _
                    QuantityOnHand - QuantityReserved

                Me.txtAvailableQuantity = quantityAvailable

                If FromLocationID > 0 Then

                    Me.txtFromLocationID = FromLocationID

                    Me.txtCurrentLocation = _
                        GetWarehouseName(FromLocationID)

                    Me.txtShelfLocation = _
                        GetShelfLocationDescription(FromLocationID)

                Else

                    Me.txtFromLocationID = Null
                    Me.txtCurrentLocation = "No warehouse recorded"
                    Me.txtShelfLocation = "No shelf location recorded"

                End If

            End If

        End If

        rs.close
        Set rs = Nothing
        Set db = Nothing

    End If

Exit_Handler:

    Exit Sub

Err_Handler:

    Me.txtAvailableQuantity = Null
    Me.txtCurrentLocation = "Unable to retrieve warehouse"
    Me.txtShelfLocation = "Unable to retrieve shelf location"
    Me.txtFromLocationID = Null

    Resume Exit_Handler

End Sub


'=============================================================
' LOAD SELECTED SERIAL NUMBER LOCATION
'=============================================================
Private Sub LoadSelectedSerializedInstance()

    On Error GoTo Err_Handler

    Dim db As DAO.Database
    Dim rs As DAO.Recordset

    Dim ItemInstanceID As Long
    Dim FromLocationID As Long

    If IsNull(Me.cboItemInstanceID) Then

        Me.txtFromLocationID = Null
        Me.txtCurrentLocation = _
            "Select an available serial number"
        Me.txtShelfLocation = _
            "Select an available serial number"

        Exit Sub

    End If

    ItemInstanceID = CLng(Me.cboItemInstanceID)

    Set db = CurrentDb

    Set rs = db.OpenRecordset( _
        "SELECT CurrentLocationID " & _
        "FROM T_ItemInstances " & _
        "WHERE ItemInstanceID = " & ItemInstanceID & ";", _
        dbOpenSnapshot _
    )

    If rs.EOF Then

        Me.txtFromLocationID = Null
        Me.txtCurrentLocation = _
            "Selected serial number could not be found"
        Me.txtShelfLocation = _
            "Selected serial number could not be found"

    Else

        FromLocationID = Nz(rs!CurrentLocationID, 0)

        If FromLocationID > 0 Then

            Me.txtFromLocationID = FromLocationID

            Me.txtCurrentLocation = _
                GetWarehouseName(FromLocationID)

            Me.txtShelfLocation = _
                GetShelfLocationDescription(FromLocationID)

        Else

            Me.txtFromLocationID = Null
            Me.txtCurrentLocation = _
                "Selected serial number has no warehouse"
            Me.txtShelfLocation = _
                "Selected serial number has no shelf location"

        End If

    End If

    rs.close
    Set rs = Nothing
    Set db = Nothing

Exit_Handler:

    Exit Sub

Err_Handler:

    Me.txtFromLocationID = Null
    Me.txtCurrentLocation = "Unable to retrieve warehouse"
    Me.txtShelfLocation = "Unable to retrieve shelf location"

    Resume Exit_Handler

End Sub


'=============================================================
' CONFIGURE SERIAL CONTROLS
'=============================================================
Private Sub ConfigureSerialControls( _
    ByVal RequiresSerial As Boolean _
)

    If RequiresSerial Then

        Me.txtQuantity = 1
        Me.txtQuantity.Locked = True

        Me.cboItemInstanceID.Enabled = True
        Me.cboItemInstanceID.Locked = False

    Else

        Me.txtQuantity = Null
        Me.txtQuantity.Locked = False

        Me.cboItemInstanceID = Null
        Me.cboItemInstanceID.RowSource = vbNullString
        Me.cboItemInstanceID.Enabled = False
        Me.cboItemInstanceID.Locked = True

    End If

End Sub


'=============================================================
' CONFIGURE PROJECT CONTROL
'=============================================================
Private Sub ConfigureProjectControl()

    Dim Reason As String

    If mProjectShipmentMode Then
        Me.cboProjectID = mProjectID
        Me.cboProjectID.Enabled = True
        Me.cboProjectID.Locked = True
        Exit Sub
    End If

    Reason = Trim$(Nz(Me.cboReason, vbNullString))

    If IsProjectIssueReason(Reason) Then

        Me.cboProjectID.Enabled = True
        Me.cboProjectID.Locked = False
        Me.cboProjectID.Requery

    Else

        Me.cboProjectID = Null
        Me.cboProjectID.Enabled = False
        Me.cboProjectID.Locked = True

    End If

End Sub


'=============================================================
' CONFIGURE EXPECTED RETURN DATE CONTROL
'=============================================================
Private Sub ConfigureExpectedReturnControl()

    Dim Reason As String

    If mProjectShipmentMode Then
        Me.txtExpectedReturnDate = Null
        Me.txtExpectedReturnDate.Enabled = False
        Me.txtExpectedReturnDate.Locked = True
        Exit Sub
    End If

    Reason = Trim$(Nz(Me.cboReason, vbNullString))

    If IsReturnableIssueReason(Reason) Then

        Me.txtExpectedReturnDate.Enabled = True
        Me.txtExpectedReturnDate.Locked = False

    Else

        Me.txtExpectedReturnDate = Null
        Me.txtExpectedReturnDate.Enabled = False
        Me.txtExpectedReturnDate.Locked = True

    End If

End Sub


'=============================================================
' RETURNABLE ISSUE REASON CHECK
'=============================================================
Private Function IsReturnableIssueReason( _
    ByVal Reason As String _
) As Boolean

    Dim r As String

    r = UCase$(Trim$(Nz(Reason, vbNullString)))

    IsReturnableIssueReason = _
        r = "ISSUE TO PROJECT" _
        Or r = "PROJECT ISSUE" _
        Or r = "PROJECT TOOL ISSUE" _
        Or r = "ISSUED TO PROJECT" _
        Or r = "TOOL ISSUE"

End Function


'=============================================================
' PROJECT ISSUE REASON CHECK
'=============================================================
Private Function IsProjectIssueReason( _
    ByVal Reason As String _
) As Boolean

    Dim r As String

    r = UCase$(Trim$(Nz(Reason, vbNullString)))

    IsProjectIssueReason = _
        r = "ISSUE TO PROJECT" _
        Or r = "PROJECT ISSUE" _
        Or r = "PROJECT TOOL ISSUE" _
        Or r = "ISSUED TO PROJECT" _
        Or r = "PROJECT SHIPMENT"

End Function


'=============================================================
' POPULATE SERIAL-NUMBER COMBO
'=============================================================
Private Sub PopulateSerialNumberCombo( _
    ByVal ItemID As Long _
)

    Me.cboItemInstanceID = Null

    Me.cboItemInstanceID.RowSource = _
        "SELECT II.ItemInstanceID, II.SerialNumber, " & _
            "L.LocationCode, L.LocationName " & _
        "FROM T_ItemInstances AS II " & _
        "LEFT JOIN T_Locations AS L " & _
            "ON II.CurrentLocationID=L.LocationID " & _
        "WHERE II.ItemID=" & ItemID & " " & _
            "AND II.CurrentLocationID Is Not Null " & _
            "AND II.CurrentStatus='In Stock' "

    If mProjectShipmentMode Then
        Me.cboItemInstanceID.RowSource = _
            Me.cboItemInstanceID.RowSource & _
            "AND II.ItemInstanceID In (" & _
                "SELECT M.ItemInstanceID " & _
                "FROM T_StockMovements AS M " & _
                "INNER JOIN T_StockMovementTypes AS MT " & _
                    "ON M.MovementTypeID=MT.MovementTypeID " & _
                "WHERE M.ProjectID=" & mProjectID & " " & _
                    "AND MT.MovementTypeCode='STOCK_IN' " & _
                    "AND M.Reason=" & _
                        SqlText(REASON_PRODUCTION_RECEIPT) & _
                    " AND M.ItemInstanceID Is Not Null) "
    End If

    Me.cboItemInstanceID.RowSource = _
        Me.cboItemInstanceID.RowSource & _
        "ORDER BY II.SerialNumber;"

    Me.cboItemInstanceID.Requery

End Sub


'=============================================================
' PROJECT SHIPMENT HELPERS
'=============================================================

Private Sub GetProjectOutputSummary( _
    ByVal db As DAO.Database, _
    ByVal ProjectID As Long, _
    ByRef outputItemID As Long, _
    ByRef receivedQuantity As Double, _
    ByRef shippedQuantity As Double _
)

    On Error GoTo Err_Handler

    Dim rs As DAO.Recordset
    Dim itemCount As Long
    Dim movementItemID As Long
    Dim errorNumber As Long
    Dim errorDescription As String
    Dim errorSource As String

    outputItemID = 0
    receivedQuantity = 0
    shippedQuantity = 0

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
            vbObjectError + 8210, _
            "GetProjectOutputSummary", _
            "The project has production receipts for more than " & _
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
                vbObjectError + 8211, _
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
            vbObjectError + 8223, _
            "GetProjectOutputSummary", _
            "Project shipment quantity exceeds its production " & _
            "receipt quantity. Correct the movement history first."
    End If

Exit_Handler:
    On Error Resume Next
    If Not (rs Is Nothing) Then rs.close
    Set rs = Nothing
    On Error GoTo 0
    Exit Sub

Err_Handler:
    errorNumber = Err.Number
    errorDescription = Err.Description
    errorSource = Err.Source

    On Error Resume Next
    If Not (rs Is Nothing) Then rs.close
    Set rs = Nothing
    On Error GoTo 0

    Err.Raise errorNumber, errorSource, errorDescription

End Sub


Private Function ProjectUnshippedQuantity() As Double

    ProjectUnshippedQuantity = _
        mProjectReceivedQuantity - mProjectShippedQuantity

    If ProjectUnshippedQuantity < 0 Then
        ProjectUnshippedQuantity = 0
    End If

End Function


Private Function ValidateProjectShipmentEntry( _
    ByVal ItemID As Long, _
    ByVal ItemInstanceID As Long, _
    ByVal ProjectID As Long, _
    ByVal Quantity As Double _
) As Boolean

    On Error GoTo Err_Handler

    Dim db As DAO.Database
    Dim rs As DAO.Recordset
    Dim outputItemID As Long
    Dim receivedQuantity As Double
    Dim shippedQuantity As Double
    Dim unshippedQuantity As Double
    Dim projectStatus As String

    ValidateProjectShipmentEntry = False

    If Not mProjectShipmentMode Or mProjectID <= 0 Then
        Err.Raise _
            vbObjectError + 8212, _
            "ValidateProjectShipmentEntry", _
            "The project shipment context is not active."
    End If

    If ProjectID <> mProjectID Then
        MsgBox _
            "Use the project supplied by the Project form.", _
            vbExclamation, _
            "Project Shipment"

        Me.cboProjectID = mProjectID
        Exit Function
    End If

    If ItemID <> mProjectOutputItemID Then
        MsgBox _
            "Only the output item established by the project's " & _
            "production receipts can be shipped.", _
            vbExclamation, _
            "Project Output Item"

        Me.cboItemID = mProjectOutputItemID
        LoadSelectedItem
        Exit Function
    End If

    Set db = CurrentDb
    Set rs = db.OpenRecordset( _
        "SELECT * FROM T_Projects " & _
        "WHERE ProjectID=" & mProjectID & ";", _
        dbOpenSnapshot)

    If rs.EOF Then
        Err.Raise _
            vbObjectError + 8213, _
            "ValidateProjectShipmentEntry", _
            "The project no longer exists."
    End If

    projectStatus = UCase$(Trim$(Nz(rs!Status, vbNullString)))
    mProjectBuildQuantity = Nz(rs!BuildQuantity, 0)
    mProjectSalesOrderID = _
        ResolveLinkedSalesOrderID(db, rs, mProjectID)

    rs.close
    Set rs = Nothing

    If projectStatus <> "READY TO SHIP" Then
        MsgBox _
            "The project status must be Ready to Ship.", _
            vbExclamation, _
            "Project Not Ready"
        Exit Function
    End If

    GetProjectOutputSummary _
        db, _
        mProjectID, _
        outputItemID, _
        receivedQuantity, _
        shippedQuantity

    If outputItemID <> ItemID Then
        Err.Raise _
            vbObjectError + 8214, _
            "ValidateProjectShipmentEntry", _
            "The project output item changed before posting."
    End If

    unshippedQuantity = receivedQuantity - shippedQuantity

    If unshippedQuantity <= QUANTITY_TOLERANCE Then
        MsgBox _
            "There is no received project output waiting to ship.", _
            vbInformation, _
            "Nothing to Ship"
        Exit Function
    End If

    If Quantity > unshippedQuantity + QUANTITY_TOLERANCE Then
        MsgBox _
            "The shipment quantity exceeds the received quantity " & _
            "still waiting to ship." & vbCrLf & vbCrLf & _
            "Available for this project: " & _
                Format$(unshippedQuantity, "0.######") & vbCrLf & _
            "Entered quantity: " & _
                Format$(Quantity, "0.######"), _
            vbExclamation, _
            "Shipment Quantity Too High"

        Me.txtQuantity = unshippedQuantity
        Exit Function
    End If

    If ItemInstanceID > 0 Then
        If Not SerializedInstanceBelongsToProject( _
            db, _
            ItemInstanceID, _
            mProjectID) Then

            MsgBox _
                "The selected serial number was not received for " & _
                "this project.", _
                vbExclamation, _
                "Project Serial Number"
            Exit Function
        End If
    End If

    If mProjectSalesOrderID > 0 Then
        ValidateLinkedSalesOrderShipment _
            db, _
            mProjectSalesOrderID, _
            ItemID, _
            Quantity
    End If

    mProjectReceivedQuantity = receivedQuantity
    mProjectShippedQuantity = shippedQuantity
    ValidateProjectShipmentEntry = True

Exit_Handler:
    On Error Resume Next
    If Not (rs Is Nothing) Then rs.close
    Set rs = Nothing
    Set db = Nothing
    On Error GoTo 0
    Exit Function

Err_Handler:
    ValidateProjectShipmentEntry = False
    MsgBox _
        "The project shipment could not be validated." & _
        vbCrLf & vbCrLf & _
        "Error Number: " & Err.Number & vbCrLf & _
        "Details: " & Err.Description, _
        vbExclamation, _
        "Project Shipment"
    Resume Exit_Handler

End Function


Private Function SerializedInstanceBelongsToProject( _
    ByVal db As DAO.Database, _
    ByVal ItemInstanceID As Long, _
    ByVal ProjectID As Long _
) As Boolean

    Dim rs As DAO.Recordset

    Set rs = db.OpenRecordset( _
        "SELECT M.StockMovementID " & _
        "FROM T_StockMovements AS M " & _
        "INNER JOIN T_StockMovementTypes AS MT " & _
            "ON M.MovementTypeID=MT.MovementTypeID " & _
        "WHERE M.ProjectID=" & ProjectID & " " & _
            "AND M.ItemInstanceID=" & ItemInstanceID & " " & _
            "AND MT.MovementTypeCode='STOCK_IN' " & _
            "AND M.Reason=" & _
                SqlText(REASON_PRODUCTION_RECEIPT) & ";", _
        dbOpenSnapshot)

    SerializedInstanceBelongsToProject = Not rs.EOF

    rs.close
    Set rs = Nothing

End Function


Private Function ResolveLinkedSalesOrderID( _
    ByVal db As DAO.Database, _
    ByVal projectRecordset As DAO.Recordset, _
    ByVal ProjectID As Long _
) As Long

    On Error GoTo Err_Handler

    Dim rs As DAO.Recordset
    Dim foundID As Long
    Dim CustomerID As Long
    Dim CustomerPOReference As String
    Dim errorNumber As Long
    Dim errorDescription As String
    Dim errorSource As String

    ResolveLinkedSalesOrderID = 0

    If RecordsetHasField(projectRecordset, "SalesOrderID") Then
        foundID = Nz(projectRecordset.Fields("SalesOrderID").value, 0)

        If foundID > 0 Then
            If DCount( _
                "*", _
                "T_SalesOrders", _
                "SalesOrderID=" & foundID) = 0 Then

                Err.Raise _
                    vbObjectError + 8215, _
                    "ResolveLinkedSalesOrderID", _
                    "The SalesOrderID saved on this project no " & _
                    "longer exists."
            End If

            ResolveLinkedSalesOrderID = foundID
            GoTo Exit_Handler
        End If
    End If

    If TableHasField(db, "T_SalesOrders", "ProjectID") Then
        Set rs = db.OpenRecordset( _
            "SELECT SalesOrderID FROM T_SalesOrders " & _
            "WHERE ProjectID=" & ProjectID & " " & _
            "ORDER BY SalesOrderID;", _
            dbOpenSnapshot)

        If Not rs.EOF Then
            foundID = Nz(rs!SalesOrderID, 0)
            rs.MoveNext

            If Not rs.EOF Then
                Err.Raise _
                    vbObjectError + 8216, _
                    "ResolveLinkedSalesOrderID", _
                    "More than one Sales Order is linked to this project."
            End If

            ResolveLinkedSalesOrderID = foundID
            GoTo Exit_Handler
        End If

        rs.close
        Set rs = Nothing
    End If

    'Existing databases may link a project through its Customer PO
    'reference instead of a direct SalesOrderID or ProjectID field.
    If Not RecordsetHasField( _
        projectRecordset, _
        "CustomerPOReference") Then

        GoTo Exit_Handler
    End If

    CustomerPOReference = Trim$(Nz( _
        projectRecordset.Fields("CustomerPOReference").value, _
        vbNullString))
    CustomerID = Nz(RecordsetValue( _
        projectRecordset, _
        "CustomerID", _
        0), 0)

    If Len(CustomerPOReference) = 0 Or CustomerID <= 0 Then
        GoTo Exit_Handler
    End If

    If Not TableHasField(db, "T_CustomerPOs", "PONumber") Or _
       Not TableHasField(db, "T_CustomerPOs", "SalesOrderID") Then

        GoTo Exit_Handler
    End If

    Set rs = db.OpenRecordset( _
        "SELECT SalesOrderID FROM T_CustomerPOs " & _
        "WHERE CustomerID=" & CustomerID & " " & _
            "AND PONumber=" & SqlText(CustomerPOReference) & " " & _
            "AND SalesOrderID Is Not Null " & _
        "ORDER BY CustomerPOID;", _
        dbOpenSnapshot)

    If rs.EOF Then
        GoTo Exit_Handler
    End If

    foundID = Nz(rs!SalesOrderID, 0)
    rs.MoveNext

    If Not rs.EOF Then
        Err.Raise _
            vbObjectError + 8222, _
            "ResolveLinkedSalesOrderID", _
            "More than one Customer PO with this reference is " & _
            "linked to a Sales Order."
    End If

    ResolveLinkedSalesOrderID = foundID

Exit_Handler:
    On Error Resume Next
    If Not (rs Is Nothing) Then rs.close
    Set rs = Nothing
    On Error GoTo 0
    Exit Function

Err_Handler:
    errorNumber = Err.Number
    errorDescription = Err.Description
    errorSource = Err.Source

    On Error Resume Next
    If Not (rs Is Nothing) Then rs.close
    Set rs = Nothing
    On Error GoTo 0

    Err.Raise errorNumber, errorSource, errorDescription

End Function


Private Sub ValidateLinkedSalesOrderShipment( _
    ByVal db As DAO.Database, _
    ByVal SalesOrderID As Long, _
    ByVal ItemID As Long, _
    ByVal Quantity As Double _
)

    Dim rs As DAO.Recordset
    Dim lineRemaining As Double
    Dim totalRemaining As Double

    Set rs = db.OpenRecordset( _
        "SELECT * FROM T_SalesOrderLines " & _
        "WHERE SalesOrderID=" & SalesOrderID & " " & _
            "AND ItemID=" & ItemID & ";", _
        dbOpenSnapshot)

    Do While Not rs.EOF
        lineRemaining = _
            CDbl(Nz(RecordsetValue( _
                rs, _
                "QuantityOrdered", _
                0), 0)) - _
            CDbl(Nz(RecordsetValue( _
                rs, _
                "QuantityShipped", _
                0), 0))

        If lineRemaining > 0 Then
            totalRemaining = totalRemaining + lineRemaining
        End If

        rs.MoveNext
    Loop

    rs.close
    Set rs = Nothing

    If totalRemaining + QUANTITY_TOLERANCE < Quantity Then
        Err.Raise _
            vbObjectError + 8217, _
            "ValidateLinkedSalesOrderShipment", _
            "Linked Sales Order " & SalesOrderID & _
            " does not have enough remaining quantity for this " & _
            "project output item." & vbCrLf & vbCrLf & _
            "Sales Order remaining: " & _
                Format$(totalRemaining, "0.######") & vbCrLf & _
            "Shipment quantity: " & _
                Format$(Quantity, "0.######")
    End If

End Sub


Private Sub ApplyShipmentToLinkedSalesOrder( _
    ByVal db As DAO.Database, _
    ByVal SalesOrderID As Long, _
    ByVal ItemID As Long, _
    ByVal Quantity As Double, _
    ByVal MovementDate As Date _
)

    Dim rs As DAO.Recordset
    Dim rsHeader As DAO.Recordset
    Dim quantityToAllocate As Double
    Dim lineOrdered As Double
    Dim lineShipped As Double
    Dim lineRemaining As Double
    Dim allocatedQuantity As Double
    Dim allLinesShipped As Boolean

    quantityToAllocate = Quantity

    Set rs = db.OpenRecordset( _
        "SELECT * FROM T_SalesOrderLines " & _
        "WHERE SalesOrderID=" & SalesOrderID & " " & _
            "AND ItemID=" & ItemID & " " & _
        "ORDER BY SalesOrderLineID;", _
        dbOpenDynaset)

    Do While Not rs.EOF And _
             quantityToAllocate > QUANTITY_TOLERANCE

        lineOrdered = CDbl(Nz(RecordsetValue( _
            rs, _
            "QuantityOrdered", _
            0), 0))
        lineShipped = CDbl(Nz(RecordsetValue( _
            rs, _
            "QuantityShipped", _
            0), 0))
        lineRemaining = lineOrdered - lineShipped

        If lineRemaining > QUANTITY_TOLERANCE Then
            If lineRemaining >= quantityToAllocate Then
                allocatedQuantity = quantityToAllocate
            Else
                allocatedQuantity = lineRemaining
            End If

            rs.Edit
            SetRecordsetFieldIfExists _
                rs, _
                "QuantityShipped", _
                lineShipped + allocatedQuantity

            If lineShipped + allocatedQuantity + _
               QUANTITY_TOLERANCE >= lineOrdered Then

                SetRecordsetFieldIfExists rs, "LineStatus", "Shipped"
            Else
                SetRecordsetFieldIfExists _
                    rs, _
                    "LineStatus", _
                    "Partially Shipped"
            End If

            rs.Update
            quantityToAllocate = _
                quantityToAllocate - allocatedQuantity
        End If

        rs.MoveNext
    Loop

    rs.close
    Set rs = Nothing

    If quantityToAllocate > QUANTITY_TOLERANCE Then
        Err.Raise _
            vbObjectError + 8218, _
            "ApplyShipmentToLinkedSalesOrder", _
            "The linked Sales Order quantity changed before posting."
    End If

    allLinesShipped = SalesOrderAllLinesShipped(db, SalesOrderID)

    Set rsHeader = db.OpenRecordset( _
        "SELECT * FROM T_SalesOrders " & _
        "WHERE SalesOrderID=" & SalesOrderID & ";", _
        dbOpenDynaset)

    If rsHeader.EOF Then
        Err.Raise _
            vbObjectError + 8219, _
            "ApplyShipmentToLinkedSalesOrder", _
            "The linked Sales Order no longer exists."
    End If

    rsHeader.Edit

    If allLinesShipped Then
        SetRecordsetFieldIfExists _
            rsHeader, _
            "SalesOrderStatus", _
            "Shipped"
        SetRecordsetFieldIfExists _
            rsHeader, _
            "ShipDate", _
            DateValue(MovementDate)
    Else
        SetRecordsetFieldIfExists _
            rsHeader, _
            "SalesOrderStatus", _
            "Partially Shipped"
    End If

    SetRecordsetFieldIfExists rsHeader, "UpdatedOn", Now()
    rsHeader.Update

    rsHeader.close
    Set rsHeader = Nothing

End Sub


Private Function SalesOrderAllLinesShipped( _
    ByVal db As DAO.Database, _
    ByVal SalesOrderID As Long _
) As Boolean

    Dim rs As DAO.Recordset
    Dim LineCount As Long
    Dim OrderedQuantity As Double
    Dim shippedQuantity As Double

    SalesOrderAllLinesShipped = False

    Set rs = db.OpenRecordset( _
        "SELECT * FROM T_SalesOrderLines " & _
        "WHERE SalesOrderID=" & SalesOrderID & ";", _
        dbOpenSnapshot)

    Do While Not rs.EOF
        LineCount = LineCount + 1
        OrderedQuantity = CDbl(Nz(RecordsetValue( _
            rs, _
            "QuantityOrdered", _
            0), 0))
        shippedQuantity = CDbl(Nz(RecordsetValue( _
            rs, _
            "QuantityShipped", _
            0), 0))

        If shippedQuantity + QUANTITY_TOLERANCE < _
           OrderedQuantity Then

            GoTo Exit_Handler
        End If

        rs.MoveNext
    Loop

    SalesOrderAllLinesShipped = (LineCount > 0)

Exit_Handler:
    On Error Resume Next
    If Not (rs Is Nothing) Then rs.close
    Set rs = Nothing
    On Error GoTo 0

End Function


Private Sub UpdateProjectShipmentStatus( _
    ByVal db As DAO.Database, _
    ByVal quantityJustShipped As Double, _
    ByVal MovementDate As Date _
)

    Dim rs As DAO.Recordset
    Dim shippedAfterPosting As Double
    Dim currentStatus As String

    shippedAfterPosting = _
        mProjectShippedQuantity + quantityJustShipped

    Set rs = db.OpenRecordset( _
        "SELECT * FROM T_Projects " & _
        "WHERE ProjectID=" & mProjectID & ";", _
        dbOpenDynaset)

    If rs.EOF Then
        Err.Raise _
            vbObjectError + 8220, _
            "UpdateProjectShipmentStatus", _
            "The project could not be found while updating shipment status."
    End If

    currentStatus = UCase$(Trim$(Nz(rs!Status, vbNullString)))

    If currentStatus <> "READY TO SHIP" Then
        Err.Raise _
            vbObjectError + 8221, _
            "UpdateProjectShipmentStatus", _
            "The project status changed before shipment posting."
    End If

    rs.Edit
    SetRecordsetFieldIfExists _
        rs, _
        "QuantityShipped", _
        shippedAfterPosting

    If shippedAfterPosting + QUANTITY_TOLERANCE >= _
       mProjectBuildQuantity Then

        rs!Status = "Shipped"
        SetRecordsetFieldIfExists rs, "ShippedDate", MovementDate
    Else
        rs!Status = "Ready to Ship"
        SetRecordsetFieldIfExists rs, "ShippedDate", Null
    End If

    SetRecordsetFieldIfExists rs, "UpdatedOn", Now()
    rs.Update

    rs.close
    Set rs = Nothing

End Sub


Private Sub ApplyProjectShipmentControlValues()

    If Not mProjectShipmentMode Then
        Exit Sub
    End If

    Me.cboItemID = mProjectOutputItemID
    Me.cboItemID.Enabled = True
    Me.cboItemID.Locked = True

    Me.cboReason = REASON_PROJECT_SHIPMENT
    Me.cboReason.Enabled = True
    Me.cboReason.Locked = True

    Me.cboProjectID = mProjectID
    Me.cboProjectID.Enabled = True
    Me.cboProjectID.Locked = True

    Me.txtReferenceDocument = mProjectCode
    Me.txtExpectedReturnDate = Null
    Me.txtExpectedReturnDate.Enabled = False
    Me.txtExpectedReturnDate.Locked = True

    If ItemRequiresSerial(mProjectOutputItemID) Then
        Me.txtQuantity = 1
    Else
        Me.txtQuantity = ProjectUnshippedQuantity()
    End If

End Sub


Private Sub PrepareNextProjectShipment( _
    ByVal RequiresSerial As Boolean _
)

    Dim remainingQuantity As Double

    remainingQuantity = ProjectUnshippedQuantity()
    Me.txtNOtes = Null

    If remainingQuantity <= QUANTITY_TOLERANCE Then
        Me.cmdPostStockOut.Enabled = False
        Exit Sub
    End If

    If RequiresSerial Then
        Me.cboItemInstanceID = Null
        Me.txtQuantity = 1
        PopulateSerialNumberCombo mProjectOutputItemID
        Me.cboItemInstanceID.SetFocus
    Else
        Me.txtQuantity = remainingQuantity
        Me.txtQuantity.SetFocus
    End If

End Sub


Private Sub RefreshProjectFormIfOpen()

    On Error Resume Next

    If (SysCmd( _
        acSysCmdGetObjectState, _
        acForm, _
        PROJECT_FORM) And acObjStateOpen) <> 0 Then

        Forms(PROJECT_FORM).Refresh
        Forms(PROJECT_FORM).Recalc
    End If

    On Error GoTo 0

End Sub


Private Sub EnsureValueListItem( _
    ByVal comboControl As Access.ComboBox, _
    ByVal listValue As String _
)

    Dim entries() As String
    Dim entry As Variant
    Dim currentRowSource As String

    If StrComp( _
        comboControl.RowSourceType, _
        "Value List", _
        vbTextCompare) <> 0 Then

        comboControl.LimitToList = False
        Exit Sub
    End If

    currentRowSource = Nz(comboControl.RowSource, vbNullString)
    entries = Split(currentRowSource, ";")

    For Each entry In entries
        If StrComp( _
            Trim$(CStr(entry)), _
            listValue, _
            vbTextCompare) = 0 Then

            Exit Sub
        End If
    Next entry

    If Len(Trim$(currentRowSource)) > 0 Then
        currentRowSource = currentRowSource & ";"
    End If

    comboControl.RowSource = currentRowSource & listValue

End Sub


Private Function TableHasField( _
    ByVal db As DAO.Database, _
    ByVal tableName As String, _
    ByVal fieldName As String _
) As Boolean

    On Error GoTo Err_Handler

    Dim fieldObject As DAO.Field

    Set fieldObject = _
        db.TableDefs(tableName).Fields(fieldName)
    TableHasField = True

Exit_Handler:
    Set fieldObject = Nothing
    Exit Function

Err_Handler:
    TableHasField = False
    Resume Exit_Handler

End Function


Private Function RecordsetHasField( _
    ByVal rs As DAO.Recordset, _
    ByVal fieldName As String _
) As Boolean

    On Error GoTo Err_Handler

    Dim testName As String

    testName = rs.Fields(fieldName).Name
    RecordsetHasField = True

Exit_Handler:
    Exit Function

Err_Handler:
    RecordsetHasField = False
    Resume Exit_Handler

End Function


Private Function RecordsetValue( _
    ByVal rs As DAO.Recordset, _
    ByVal fieldName As String, _
    ByVal defaultValue As Variant _
) As Variant

    If RecordsetHasField(rs, fieldName) Then
        If IsNull(rs.Fields(fieldName).value) Then
            RecordsetValue = defaultValue
        Else
            RecordsetValue = rs.Fields(fieldName).value
        End If
    Else
        RecordsetValue = defaultValue
    End If

End Function


Private Sub SetRecordsetFieldIfExists( _
    ByVal rs As DAO.Recordset, _
    ByVal fieldName As String, _
    ByVal fieldValue As Variant _
)

    If RecordsetHasField(rs, fieldName) Then
        rs.Fields(fieldName).value = fieldValue
    End If

End Sub


Private Function NullIfZero(ByVal numberValue As Long) As Variant

    If numberValue > 0 Then
        NullIfZero = numberValue
    Else
        NullIfZero = Null
    End If

End Function


'=============================================================
' PREPARE NEXT ENTRY
'=============================================================
Private Sub PrepareNextEntry( _
    ByVal RequiresSerial As Boolean _
)

    Me.txtReferenceDocument = Null
    Me.txtIssuedTo = Null
    Me.txtNOtes = Null
    Me.txtExpectedReturnDate = Null

    ConfigureExpectedReturnControl

    If RequiresSerial Then

        Me.cboItemInstanceID = Null
        Me.txtQuantity = 1

        PopulateSerialNumberCombo CLng(Me.cboItemID)

        Me.cboItemInstanceID.SetFocus

    Else

        Me.txtQuantity = Null
        Me.txtQuantity.SetFocus

    End If

End Sub


'=============================================================
' RESET FORM
'=============================================================
Private Sub ResetForm()

    mProjectShipmentMode = False
    mProjectID = 0
    mProjectOutputItemID = 0
    mProjectOutputLocationID = 0
    mProjectCustomerID = 0
    mProjectAssemblyTemplateID = 0
    mProjectSalesOrderID = 0
    mProjectBuildQuantity = 0
    mProjectReceivedQuantity = 0
    mProjectShippedQuantity = 0
    mProjectCode = vbNullString

    Me.Caption = "Stock Out"
    Me.cboItemID.Enabled = True
    Me.cboItemID.Locked = False
    Me.cboReason.Enabled = True
    Me.cboReason.Locked = False
    Me.cboProjectID.Locked = False

    On Error Resume Next
    Me.cmdPostStockOut.Enabled = True
    On Error GoTo 0

    Me.cboItemID = Null
    Me.chkRequiresSerial = False

    Me.txtAvailableQuantity = Null
    Me.txtCurrentLocation = Null
    Me.txtShelfLocation = Null
    Me.txtFromLocationID = Null

    Me.cboItemInstanceID = Null
    Me.cboItemInstanceID.RowSource = vbNullString
    Me.cboItemInstanceID.Enabled = False
    Me.cboItemInstanceID.Locked = True

    Me.txtQuantity = Null
    Me.txtQuantity.Locked = False

    Me.txtMovementDate = Now()
    Me.cboReason = "Issue to Project"

    Me.cboProjectID = Null

    Me.txtReferenceDocument = Null
    Me.cboPerformedByUserID = Null
    Me.txtIssuedTo = Null
    Me.txtNOtes = Null
    Me.txtExpectedReturnDate = Null

    ConfigureProjectControl
    ConfigureExpectedReturnControl

    On Error Resume Next
    Me.cboItemID.SetFocus
    On Error GoTo 0

End Sub


'=============================================================
' RESET ITEM DISPLAY
'=============================================================
Private Sub ResetItemStatusDisplay()

    Me.chkRequiresSerial = False

    Me.txtAvailableQuantity = Null
    Me.txtCurrentLocation = Null
    Me.txtShelfLocation = Null
    Me.txtFromLocationID = Null

    Me.cboItemInstanceID = Null
    Me.cboItemInstanceID.RowSource = vbNullString
    Me.cboItemInstanceID.Enabled = False
    Me.cboItemInstanceID.Locked = True

    Me.txtQuantity = Null
    Me.txtQuantity.Locked = False

End Sub


'=============================================================
' REQUERY RECENT TRANSACTIONS
'=============================================================
Private Sub RequeryRecentSubform()

    On Error Resume Next

    Me.Controls(RECENT_SUBFORM_CONTROL).Form.Requery

    On Error GoTo 0

End Sub


'=============================================================
' ITEM SERIAL RULE
'=============================================================
Private Function ItemRequiresSerial( _
    ByVal ItemID As Long _
) As Boolean

    Dim db As DAO.Database
    Dim rs As DAO.Recordset

    Set db = CurrentDb

    Set rs = db.OpenRecordset( _
        "SELECT RequiresSerial " & _
        "FROM T_Items " & _
        "WHERE ItemID=" & ItemID & ";", _
        dbOpenSnapshot)

    If rs.EOF Then

        rs.close
        Set rs = Nothing
        Set db = Nothing

        Err.Raise _
            vbObjectError + 2301, _
            "ItemRequiresSerial", _
            "The selected item could not be found."

    End If

    ItemRequiresSerial = Nz(rs!RequiresSerial, False)

    rs.close
    Set rs = Nothing
    Set db = Nothing

End Function


'=============================================================
' STATUS AFTER SERIALIZED ISSUE
'=============================================================
Private Function GetIssuedInstanceStatus( _
    ByVal Reason As String _
) As String

    Select Case UCase$(Trim$(Reason))

        Case "ISSUE TO PROJECT", "PROJECT ISSUE", "PROJECT TOOL ISSUE", "ISSUED TO PROJECT", "TOOL ISSUE"
            GetIssuedInstanceStatus = "Issued"

        Case "SALE", "PROJECT SHIPMENT"
            GetIssuedInstanceStatus = "Sold"

        Case "SUPPLIER RETURN"
            GetIssuedInstanceStatus = "Returned to Supplier"

        Case "DAMAGE WRITE-OFF"
            GetIssuedInstanceStatus = "Written Off"

        Case "NEGATIVE ADJUSTMENT"
            GetIssuedInstanceStatus = "Removed from Stock"

        Case Else
            GetIssuedInstanceStatus = "Issued"

    End Select

End Function


'=============================================================
' BUILD MOVEMENT NOTES
'=============================================================
Private Function BuildMovementNotes( _
    ByVal issuedTo As String, _
    ByVal userNotes As String _
) As String

    Dim result As String

    result = vbNullString

    If Len(issuedTo) > 0 Then
        result = "Issued To / Recipient: " & issuedTo
    End If

    If Len(userNotes) > 0 Then

        If Len(result) > 0 Then
            result = result & vbCrLf
        End If

        result = result & userNotes

    End If

    BuildMovementNotes = result

End Function


'=============================================================
' APPEND NOTE
'=============================================================
Private Function AppendNote( _
    ByVal existingText As String, _
    ByVal newText As String _
) As String

    If Len(Trim$(existingText)) = 0 Then

        AppendNote = newText

    ElseIf Len(Trim$(newText)) = 0 Then

        AppendNote = existingText

    Else

        AppendNote = _
            existingText & _
            vbCrLf & _
            String$(40, "-") & _
            vbCrLf & _
            newText

    End If

End Function


'=============================================================
' RETURN WAREHOUSE NAME
'=============================================================
Private Function GetWarehouseName( _
    ByVal LocationID As Long _
) As String

    GetWarehouseName = Nz( _
        DLookup( _
            "WarehouseName", _
            "T_Warehouses", _
            "WarehouseID = " & _
                Nz( _
                    DLookup( _
                        "WarehouseID", _
                        "T_Locations", _
                        "LocationID = " & LocationID _
                    ), _
                    0 _
                ) _
        ), _
        "Warehouse not found" _
    )

End Function


'=============================================================
' RETURN SHELF / BIN DESCRIPTION
'=============================================================
Private Function GetShelfLocationDescription( _
    ByVal LocationID As Long _
) As String

    Dim db As DAO.Database
    Dim rs As DAO.Recordset
    Dim result As String

    Set db = CurrentDb

    Set rs = db.OpenRecordset( _
        "SELECT " & _
            "LocationName, " & _
            "BinNumber, " & _
            "ShelfNumber " & _
        "FROM T_Locations " & _
        "WHERE LocationID = " & LocationID & ";", _
        dbOpenSnapshot _
    )

    If rs.EOF Then

        result = "Shelf location not found"

    Else

        result = Nz(rs!LocationName, vbNullString)

        If Len(Trim$(Nz(rs!BinNumber, vbNullString))) > 0 Then

            If Len(result) > 0 Then
                result = result & " | "
            End If

            result = result & "Bin: " & rs!BinNumber

        End If

        If Len(Trim$(Nz(rs!ShelfNumber, vbNullString))) > 0 Then

            If Len(result) > 0 Then
                result = result & " | "
            End If

            result = result & "Shelf: " & rs!ShelfNumber

        End If

        If Len(Trim$(result)) = 0 Then
            result = "Shelf location not recorded"
        End If

    End If

    rs.close
    Set rs = Nothing
    Set db = Nothing

    GetShelfLocationDescription = result

End Function


'=============================================================
' ESCAPE TEXT FOR SQL
'=============================================================
Private Function SqlText( _
    ByVal value As String _
) As String

    SqlText = "'" & Replace(value, "'", "''") & "'"

End Function
