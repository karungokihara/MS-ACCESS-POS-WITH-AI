Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Compare Database
Option Explicit

'=============================================================
' FORM: SF_InvoiceLines
'
' PURPOSE:
'   Invoice line entry.
'
' IMPORTANT:
'   Invoice lines do NOT move stock.
'   Stock movement happens later during shipment/issue.
'
' REQUIRED CONTROLS:
'   InvoiceLineID
'   InvoiceID
'   SalesOrderLineID
'   CustomerPOLineID
'   ItemID
'   ItemCode
'   LineDescription
'   Quantity
'   UnitID
'   UnitPrice
'   Taxable
'   LineSubtotal
'   CreatedOn
'   UpdatedOn
'=============================================================


'=============================================================
' FORM EVENTS
'=============================================================
Private Sub Form_Load()

    On Error GoTo Err_Handler

    ConfigureInvoiceLineForm
    ConfigureInvoiceLineCombos

Exit_Handler:
    Exit Sub

Err_Handler:
    MsgBox _
        "The invoice line form could not be loaded." & _
        vbCrLf & vbCrLf & _
        "Error " & Err.Number & ": " & Err.Description, _
        vbExclamation, _
        "Invoice Lines"

    Resume Exit_Handler

End Sub


Private Sub Form_Current()

    On Error Resume Next

    If Not IsBlankNewInvoiceLine Then
        CalculateLineSubtotal
    End If

    On Error GoTo 0

End Sub


Private Sub Form_BeforeInsert(Cancel As Integer)

    On Error GoTo Err_Handler

    If ParentInvoiceID <= 0 Then
        Cancel = True
        Exit Sub
    End If

    Me!InvoiceID = ParentInvoiceID

    If IsNull(Me!Taxable) Then
        Me!Taxable = False
    End If

Exit_Handler:
    Exit Sub

Err_Handler:
    Cancel = True

    MsgBox _
        "The new invoice line could not be initialized." & _
        vbCrLf & vbCrLf & _
        "Error " & Err.Number & ": " & Err.Description, _
        vbExclamation, _
        "New Invoice Line"

    Resume Exit_Handler

End Sub


Private Sub Form_BeforeUpdate(Cancel As Integer)

    On Error GoTo Err_Handler

    'If this is just a blank new row, discard it silently.
    If IsBlankNewInvoiceLine Then
        Me.Undo
        Exit Sub
    End If

    If ParentInvoiceID <= 0 Then

        MsgBox _
            "Save the invoice header first before saving invoice lines.", _
            vbExclamation, _
            "Invoice Header Not Saved"

        Cancel = True
        Exit Sub

    End If

    Me!InvoiceID = ParentInvoiceID

    If Not ValidateLineRecord Then
        Cancel = True
        Exit Sub
    End If

    CalculateLineSubtotal

    If IsNull(Me!Taxable) Then
        Me!Taxable = False
    End If

    If Nz(Me!InvoiceLineID, 0) = 0 Then
        SafeSetField "CreatedOn", Now()
    End If

    SafeSetField "UpdatedOn", Now()

Exit_Handler:
    Exit Sub

Err_Handler:
    Cancel = True

    MsgBox _
        "The invoice line could not be validated." & _
        vbCrLf & vbCrLf & _
        "Error " & Err.Number & ": " & Err.Description, _
        vbExclamation, _
        "Invoice Line Validation"

    Resume Exit_Handler

End Sub


Private Sub Form_AfterUpdate()

    On Error Resume Next

    If Not Me.Parent Is Nothing Then
        Me.Parent.RecalculateInvoiceTotals
    End If

    On Error GoTo 0

End Sub


'=============================================================
' CONTROL EVENTS
'=============================================================
Private Sub ItemID_AfterUpdate()

    On Error Resume Next

    Dim salesPrice As Variant
    Dim BaseUnitID As Variant
    Dim skuValue As Variant

    If IsNull(Me!ItemID) Then
        Exit Sub
    End If

    If Len(Trim$(Nz(Me!LineDescription, vbNullString))) = 0 Then
        Me!LineDescription = Nz(Me!ItemID.Column(1), vbNullString)
    End If

    skuValue = DLookup( _
        "InternalSKU", _
        "T_Items", _
        "ItemID = " & CLng(Me!ItemID) _
    )

    If Not IsNull(skuValue) Then

        If Len(Trim$(Nz(Me!ItemCode, vbNullString))) = 0 Then
            Me!ItemCode = CStr(skuValue)
        End If

    End If

    salesPrice = DLookup( _
        "DefaultSalesPrice", _
        "T_Items", _
        "ItemID = " & CLng(Me!ItemID) _
    )

    If Not IsNull(salesPrice) Then

        If Nz(Me!UnitPrice, 0) = 0 Then
            Me!UnitPrice = salesPrice
        End If

    End If

    BaseUnitID = DLookup( _
        "BaseUnitID", _
        "T_Items", _
        "ItemID = " & CLng(Me!ItemID) _
    )

    If Not IsNull(BaseUnitID) Then

        If IsNull(Me!UnitID) Then
            Me!UnitID = CLng(BaseUnitID)
        End If

    End If

    CalculateLineSubtotal

    On Error GoTo 0

    UX_FillLineFromItem Me
End Sub


Private Sub Quantity_AfterUpdate()

    If Not IsBlankNewInvoiceLine Then
        CalculateLineSubtotal
    End If

End Sub


Private Sub UnitPrice_AfterUpdate()

    If Not IsBlankNewInvoiceLine Then
        CalculateLineSubtotal
    End If

End Sub


Private Sub ItemID_NotInList(NewData As String, Response As Integer)
    On Error GoTo Err_Handler
    Response = acDataErrContinue
    If HandleComboNotInList(Me!ItemID, NewData, "F_ItemSetup", "T_Items", "ItemID", "ItemName", "InternalSKU") Then
        Response = acDataErrAdded
    Else
        Me!ItemID.Undo
    End If
    Exit Sub
Err_Handler:
    Response = acDataErrContinue
    MsgBox Err.Description, vbExclamation, "Add Item"
End Sub


'=============================================================
' CONFIGURATION
'=============================================================
Private Sub ConfigureInvoiceLineForm()

    On Error Resume Next

    Me!InvoiceLineID.Visible = False
    Me!InvoiceID.Visible = False
    Me!SalesOrderLineID.Visible = False
    Me!CustomerPOLineID.Visible = False
    Me!CreatedOn.Visible = False
    Me!UpdatedOn.Visible = False

    Me!LineSubtotal.Locked = True
    Me!LineSubtotal.Enabled = True

    Me!LineDescription.EnterKeyBehavior = True
    Me!LineDescription.ScrollBars = 2

    On Error GoTo 0

End Sub


Private Sub ConfigureInvoiceLineCombos()

    On Error Resume Next

    Me!ItemID.RowSourceType = "Table/Query"
    Me!ItemID.RowSource = _
        "SELECT ItemID, [InternalSKU] & ' - ' & [ItemName] " & _
        "FROM T_Items " & _
        "ORDER BY ItemName;"

    Me!ItemID.BoundColumn = 1
    Me!ItemID.ColumnCount = 2
    Me!ItemID.ColumnWidths = "0cm;7cm"
    Me!ItemID.LimitToList = True
    Me!ItemID.AutoExpand = True
    Me!ItemID.Requery


    Me!UnitID.RowSourceType = "Table/Query"
    Me!UnitID.RowSource = _
        "SELECT UnitID, IIf(Len(Trim(Nz([UnitCode],'')))=0,[UnitName],[UnitCode]) " & _
        "FROM T_Units " & _
        "ORDER BY UnitCode;"

    Me!UnitID.BoundColumn = 1
    Me!UnitID.ColumnCount = 2
    Me!UnitID.ColumnWidths = "0cm;3cm"
    Me!UnitID.LimitToList = True
    Me!UnitID.AutoExpand = True
    Me!UnitID.Requery

    On Error GoTo 0

End Sub


'=============================================================
' VALIDATION
'=============================================================
Private Function ValidateLineRecord() As Boolean

    On Error GoTo Err_Handler

    ValidateLineRecord = False

    If IsNull(Me!ItemID) And _
       Len(Trim$(Nz(Me!LineDescription, vbNullString))) = 0 Then

        MsgBox _
            "Select an item or enter a line description.", _
            vbExclamation, _
            "Line Description Required"

        Me!LineDescription.SetFocus
        Exit Function

    End If

    If IsNull(Me!Quantity) Then

        MsgBox _
            "Enter the invoice quantity.", _
            vbExclamation, _
            "Quantity Required"

        Me!Quantity.SetFocus
        Exit Function

    End If

    If Not IsNumeric(Me!Quantity) Then

        MsgBox _
            "Quantity must be numeric.", _
            vbExclamation, _
            "Invalid Quantity"

        Me!Quantity.SetFocus
        Exit Function

    End If

    If CDbl(Nz(Me!Quantity, 0)) <= 0 Then

        MsgBox _
            "Quantity must be greater than zero.", _
            vbExclamation, _
            "Invalid Quantity"

        Me!Quantity.SetFocus
        Exit Function

    End If

    If IsNull(Me!UnitID) Then

        MsgBox _
            "Select the unit.", _
            vbExclamation, _
            "Unit Required"

        Me!UnitID.SetFocus
        Exit Function

    End If

    If IsNull(Me!UnitPrice) Then

        MsgBox _
            "Enter the unit price.", _
            vbExclamation, _
            "Unit Price Required"

        Me!UnitPrice.SetFocus
        Exit Function

    End If

    If Not IsNumeric(Me!UnitPrice) Then

        MsgBox _
            "Unit price must be numeric.", _
            vbExclamation, _
            "Invalid Unit Price"

        Me!UnitPrice.SetFocus
        Exit Function

    End If

    If CCur(Nz(Me!UnitPrice, 0)) < 0 Then

        MsgBox _
            "Unit price cannot be negative.", _
            vbExclamation, _
            "Invalid Unit Price"

        Me!UnitPrice.SetFocus
        Exit Function

    End If

    ValidateLineRecord = True

Exit_Handler:
    Exit Function

Err_Handler:
    MsgBox _
        "Invoice line validation failed." & _
        vbCrLf & vbCrLf & _
        "Error " & Err.Number & ": " & Err.Description, _
        vbExclamation, _
        "Invoice Line Validation"

    ValidateLineRecord = False
    Resume Exit_Handler

End Function


Private Function IsBlankNewInvoiceLine() As Boolean

    On Error GoTo Err_Handler

    If Not Me.NewRecord Then
        IsBlankNewInvoiceLine = False
        Exit Function
    End If

    IsBlankNewInvoiceLine = _
        IsNull(Me!ItemID) And _
        Len(Trim$(Nz(Me!ItemCode, vbNullString))) = 0 And _
        Len(Trim$(Nz(Me!LineDescription, vbNullString))) = 0 And _
        Nz(Me!Quantity, 0) = 0 And _
        IsNull(Me!UnitID) And _
        Nz(Me!UnitPrice, 0) = 0

    Exit Function

Err_Handler:
    IsBlankNewInvoiceLine = False

End Function


'=============================================================
' CALCULATION
'=============================================================
Private Sub CalculateLineSubtotal()

    On Error Resume Next

    If IsBlankNewInvoiceLine Then
        Exit Sub
    End If

    Me!LineSubtotal = Nz(Me!Quantity, 0) * Nz(Me!UnitPrice, 0)

    On Error GoTo 0

End Sub


'=============================================================
' HELPERS
'=============================================================
Private Function ParentInvoiceID() As Long

    On Error GoTo Err_Handler

    ParentInvoiceID = Nz(Me.Parent!InvoiceID, 0)
    Exit Function

Err_Handler:
    ParentInvoiceID = 0

End Function


Private Sub SafeSetField(ByVal fieldName As String, ByVal fieldValue As Variant)

    On Error Resume Next
    Me(fieldName) = fieldValue
    On Error GoTo 0

End Sub
