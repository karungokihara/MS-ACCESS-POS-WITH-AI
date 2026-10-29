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

    DoCmd.OpenForm "F_ItemSetup", acNormal, , , acFormAdd

End Sub

Private Sub ApplyListFilter()

    On Error GoTo Err_Handler

    Dim s As String
    Dim filterText As String

    s = Trim$(Nz(Me.txtSearch, vbNullString))
    s = Replace(s, "'", "''")

    If Len(s) = 0 Then

        Me.Controls("subItemSetupList").Form.FilterOn = False
        Me.Controls("subItemSetupList").Form.Filter = vbNullString

    Else

        filterText = ""
        filterText = filterText & "Nz([InternalSKU],'') Like '*" & s & "*'"
        filterText = filterText & " OR Nz([ItemName],'') Like '*" & s & "*'"
        filterText = filterText & " OR Nz([Description],'') Like '*" & s & "*'"
        filterText = filterText & " OR Nz([CategoryName],'') Like '*" & s & "*'"
        filterText = filterText & " OR Nz([StockType],'') Like '*" & s & "*'"
        filterText = filterText & " OR Nz([UnitCode],'') Like '*" & s & "*'"

        Me.Controls("subItemSetupList").Form.Filter = filterText
        Me.Controls("subItemSetupList").Form.FilterOn = True

    End If

Exit_Handler:
    Exit Sub

Err_Handler:
    MsgBox "Item setup list filter failed." & vbCrLf & vbCrLf & _
           "Error Number: " & Err.Number & vbCrLf & _
           "Details: " & Err.Description, vbExclamation, "Item Setup List"
    Resume Exit_Handler

End Sub
