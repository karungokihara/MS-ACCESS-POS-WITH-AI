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

Private Sub btnNew_Click()

    DoCmd.OpenForm "F_Invoices", acNormal, , , acFormAdd

End Sub

Private Sub ApplyListFilter()

    On Error GoTo Err_Handler

    Dim s As String
    Dim filterText As String

    s = Trim$(Nz(Me.txtSearch, vbNullString))
    s = Replace(s, "'", "''")

    If Len(s) = 0 Then

        Me.Controls("subInvoiceList").Form.FilterOn = False
        Me.Controls("subInvoiceList").Form.Filter = vbNullString

    Else

        filterText = ""
        filterText = filterText & "Nz([InvoiceNo],'') Like '*" & s & "*'"
        filterText = filterText & " OR Nz([InvoiceType],'') Like '*" & s & "*'"
        filterText = filterText & " OR Nz([InvoiceStatus],'') Like '*" & s & "*'"
        filterText = filterText & " OR Nz([CustomerName],'') Like '*" & s & "*'"
        filterText = filterText & " OR Nz([PONumber],'') Like '*" & s & "*'"
        filterText = filterText & " OR Nz([SalesOrderNo],'') Like '*" & s & "*'"
        filterText = filterText & " OR Nz([BillToName],'') Like '*" & s & "*'"

        Me.Controls("subInvoiceList").Form.Filter = filterText
        Me.Controls("subInvoiceList").Form.FilterOn = True

    End If

Exit_Handler:
    Exit Sub

Err_Handler:
    MsgBox "Invoice list filter failed." & vbCrLf & vbCrLf & _
           "Error Number: " & Err.Number & vbCrLf & _
           "Details: " & Err.Description, vbExclamation, "Invoice List"
    Resume Exit_Handler

End Sub
