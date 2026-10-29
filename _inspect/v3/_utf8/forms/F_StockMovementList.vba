Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Compare Database
Option Explicit

Private Sub Form_Load()
    ApplyListFilter
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

Private Sub ApplyListFilter()

    On Error GoTo Err_Handler

    Dim s As String
    Dim filterText As String

    s = Trim$(Nz(Me.txtSearch, vbNullString))
    s = Replace(s, "'", "''")

    If Len(s) = 0 Then

        Me.Controls("subStockMovementList").Form.FilterOn = False
        Me.Controls("subStockMovementList").Form.Filter = vbNullString

    Else

        filterText = ""
        filterText = filterText & "Nz([MovementTypeCode],'') Like '*" & s & "*'"
        filterText = filterText & " OR Nz([MovementTypeName],'') Like '*" & s & "*'"
        filterText = filterText & " OR Nz([InternalSKU],'') Like '*" & s & "*'"
        filterText = filterText & " OR Nz([ItemName],'') Like '*" & s & "*'"
        filterText = filterText & " OR Nz([CustomerName],'') Like '*" & s & "*'"
        filterText = filterText & " OR Nz([SalesOrderNo],'') Like '*" & s & "*'"
        filterText = filterText & " OR Nz([PONumber],'') Like '*" & s & "*'"
        filterText = filterText & " OR Nz([InvoiceNo],'') Like '*" & s & "*'"
        filterText = filterText & " OR Nz([ReferenceDocument],'') Like '*" & s & "*'"
        filterText = filterText & " OR Nz([SourceDocumentType],'') Like '*" & s & "*'"
        filterText = filterText & " OR Nz([Reason],'') Like '*" & s & "*'"
        filterText = filterText & " OR Nz([MovementNotes],'') Like '*" & s & "*'"

        Me.Controls("subStockMovementList").Form.Filter = filterText
        Me.Controls("subStockMovementList").Form.FilterOn = True

    End If

Exit_Handler:
    Exit Sub

Err_Handler:
    MsgBox "Stock movement list filter failed." & vbCrLf & vbCrLf & _
           "Error Number: " & Err.Number & vbCrLf & _
           "Details: " & Err.Description, vbExclamation, "Stock Movement List"
    Resume Exit_Handler

End Sub
