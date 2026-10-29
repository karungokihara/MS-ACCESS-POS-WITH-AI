Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Compare Database

' UX-WAREHOUSE
Private Sub Form_Load()
    On Error Resume Next
    Me.Filter = ""
    Me.FilterOn = False
    Me.Caption = "Warehouses"
    If Me.NewRecord Then ApplyWarehouseCode
    If Left$(Nz(Me.OpenArgs, ""), 9) = "QUICKADD|" Then
        Me!WarehouseName = Mid$(Me.OpenArgs, 10)
        ApplyWarehouseCode
    End If
    UX_MarkRequired Me, "WarehouseName", "WarehouseCode"
End Sub

Private Sub Form_BeforeUpdate(Cancel As Integer)
    On Error GoTo Err_Handler
    If Len(Trim$(Nz(Me!WarehouseName, ""))) = 0 Then
        MsgBox "Enter the warehouse name.", vbExclamation, "Warehouse"
        Cancel = True
        Exit Sub
    End If
    ApplyWarehouseCode
    Exit Sub
Err_Handler:
    Cancel = True
    MsgBox Err.Description, vbExclamation, "Warehouse"
End Sub

Private Sub WarehouseName_AfterUpdate()
    ApplyWarehouseCode
End Sub

Private Sub ApplyWarehouseCode()
    On Error Resume Next
    Dim prefix As String
    Dim n As Long
    Dim candidate As String
    If Len(Trim$(Nz(Me!WarehouseCode, ""))) > 0 Then Exit Sub
    If Len(Trim$(Nz(Me!WarehouseName, ""))) = 0 Then Exit Sub
    prefix = UCase$(Left$(Trim$(Me!WarehouseName), 4))
    If Len(prefix) < 4 Then prefix = prefix & String$(4 - Len(prefix), "X")
    n = 1
    Do
        candidate = prefix & "-" & Format$(n, "000")
        If IsNull(DLookup("WarehouseID", "T_Warehouses", "WarehouseCode=" & UX_SqlText(candidate))) Then Exit Do
        n = n + 1
    Loop
    Me!WarehouseCode = candidate
End Sub

Private Sub Command14_Click()
    On Error GoTo Err_Handler
    If Me.Dirty Then Me.Dirty = False
    MsgBox "Warehouse saved. You can type locations in the list below.", vbInformation, "Warehouse"
    Exit Sub
Err_Handler:
    MsgBox Err.Description, vbExclamation, "Warehouse"
End Sub

Private Sub Command15_Click()
    UX_GoBack
End Sub

Private Sub Command16_Click()
    On Error Resume Next
    If Me.Dirty Then Me.Dirty = False
    DoCmd.GoToRecord , , acNewRec
    Me!WarehouseCode = Null
    Me!WarehouseName = Null
End Sub

Private Sub Form_Unload(Cancel As Integer)
    On Error Resume Next
    If Me.NewRecord And Me.Dirty Then Me.Undo
End Sub
