Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Compare Database
Option Explicit

'=============================================================
' FORM: SF_CustomerPartNumbers
'
' PURPOSE:
'   Stores customer-specific part numbers for internal items.
'
' EXAMPLE:
'   Customer Part Number: ABC-SHELF-A100
'   Internal Item:        SHELF-001 - Wooden Shelf
'
' TABLE:
'   T_ItemPartNumbers
'
' REQUIRED PARENT FORM:
'   F_CustomerSetup
'
' IMPORTANT:
'   The parent form/subform link should be:
'       Link Master Fields: CustomerID
'       Link Child Fields:  CustomerID
'=============================================================


Private Sub Form_Load()

    On Error Resume Next

    ConfigureItemCombo

    On Error GoTo 0

End Sub


Private Sub Form_BeforeInsert(Cancel As Integer)

    On Error GoTo Err_Handler

    EnsureParentCustomerSaved

    Me!CustomerID = CLng(Me.Parent!txtCustomerID)
    Me!PartNumberTypeID = GetCustomerPartNumberTypeID()

Exit_Handler:
    Exit Sub

Err_Handler:

    Cancel = True

    MsgBox _
        "The customer part-number row could not be started." & _
        vbCrLf & vbCrLf & _
        "Error Number: " & Err.Number & vbCrLf & _
        "Details: " & Err.Description, _
        vbExclamation, _
        "Customer Part Number"

    Resume Exit_Handler

End Sub


Private Sub Form_BeforeUpdate(Cancel As Integer)

    On Error GoTo Err_Handler

    EnsureParentCustomerSaved

    Me!CustomerID = CLng(Me.Parent!txtCustomerID)
    Me!PartNumberTypeID = GetCustomerPartNumberTypeID()

    If IsNull(Me!ItemID) Then

        MsgBox _
            "Select the internal item for this customer part number.", _
            vbExclamation, _
            "Internal Item Required"

        Cancel = True
        Exit Sub

    End If

    If Len(Trim$(Nz(Me!PartNumber, vbNullString))) = 0 Then

        MsgBox _
            "Enter the customer part number.", _
            vbExclamation, _
            "Customer Part Number Required"

        Cancel = True
        Exit Sub

    End If

    Me!PartNumber = UCase$(Trim$(Nz(Me!PartNumber, vbNullString)))

    If Me.NewRecord Then

        If CustomerPartNumberAlreadyExists( _
            CLng(Me!CustomerID), _
            CLng(Me!PartNumberTypeID), _
            CStr(Me!PartNumber) _
        ) Then

            MsgBox _
                "This customer part number already exists for this customer." & _
                vbCrLf & vbCrLf & _
                "Customer Part Number: " & Me!PartNumber, _
                vbExclamation, _
                "Duplicate Customer Part Number"

            Cancel = True
            Exit Sub

        End If

    End If

Exit_Handler:
    Exit Sub

Err_Handler:

    Cancel = True

    MsgBox _
        "The customer part-number row could not be validated." & _
        vbCrLf & vbCrLf & _
        "Error Number: " & Err.Number & vbCrLf & _
        "Details: " & Err.Description, _
        vbExclamation, _
        "Customer Part Number"

    Resume Exit_Handler

End Sub


Private Sub ConfigureItemCombo()

    On Error Resume Next

    Me.cboItemID.RowSourceType = "Table/Query"

    Me.cboItemID.RowSource = _
        "SELECT " & _
            "ItemID, " & _
            "InternalSKU, " & _
            "ItemName, " & _
            "Nz([ItemClass], '') AS ItemClassName " & _
        "FROM T_Items " & _
        "ORDER BY InternalSKU, ItemName;"

    Me.cboItemID.BoundColumn = 1
    Me.cboItemID.ColumnCount = 4
    Me.cboItemID.ColumnWidths = "0cm;3cm;7cm;3cm"
    Me.cboItemID.LimitToList = True
    Me.cboItemID.AutoExpand = True

    On Error GoTo 0

End Sub


Private Sub EnsureParentCustomerSaved()

    If Me.Parent Is Nothing Then

        Err.Raise _
            vbObjectError + 4001, _
            "EnsureParentCustomerSaved", _
            "This subform must be used inside F_CustomerSetup."

    End If

    If Me.Parent.NewRecord Then

        If Me.Parent.Dirty Then
            Me.Parent.Dirty = False
        End If

    ElseIf Me.Parent.Dirty Then

        Me.Parent.Dirty = False

    End If

    If IsNull(Me.Parent!txtCustomerID) Then

        Err.Raise _
            vbObjectError + 4002, _
            "EnsureParentCustomerSaved", _
            "Save the customer record before adding customer part numbers."

    End If

End Sub


Private Function GetCustomerPartNumberTypeID() As Long

    Dim valueFound As Variant

    valueFound = DLookup( _
        "PartNumberTypeID", _
        "T_PartNumberTypes", _
        "PartNumberTypeName = 'Customer'" _
    )

    If IsNull(valueFound) Then

        Err.Raise _
            vbObjectError + 4003, _
            "GetCustomerPartNumberTypeID", _
            "The part-number type 'Customer' was not found in T_PartNumberTypes."

    End If

    GetCustomerPartNumberTypeID = CLng(valueFound)

End Function


Private Function CustomerPartNumberAlreadyExists( _
    ByVal CustomerID As Long, _
    ByVal PartNumberTypeID As Long, _
    ByVal PartNumber As String _
) As Boolean

    Dim criteria As String

    criteria = _
        "CustomerID = " & CustomerID & _
        " AND PartNumberTypeID = " & PartNumberTypeID & _
        " AND PartNumber = " & SqlText(PartNumber)

    CustomerPartNumberAlreadyExists = _
        DCount("*", "T_ItemPartNumbers", criteria) > 0

End Function


Private Function SqlText( _
    ByVal value As String _
) As String

    SqlText = "'" & Replace(value, "'", "''") & "'"

End Function
