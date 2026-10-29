Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Option Compare Database

Private Sub cmdLogin_Click()

    Dim UserID As Long
    Dim RoleID As Long
    Dim Username As String

    Username = Nz(Me.txtUsername, "")

    If Username = "" Or Nz(Me.txtPassword, "") = "" Then
        MsgBox "Enter username and password.", vbExclamation
        Exit Sub
    End If

    UserID = Nz(DLookup("UserID", "T_Users", _
        "Username='" & Replace(Username, "'", "''") & _
        "' AND PasswordHash='" & Replace(Me.txtPassword, "'", "''") & "'"), 0)

    If UserID = 0 Then
        MsgBox "Invalid username or password.", vbCritical
        Exit Sub
    End If

    RoleID = Nz(DLookup("RoleID", "T_Users", _
        "UserID=" & UserID), 0)

    TempVars("CurrentUserID") = UserID
    TempVars("CurrentRoleID") = RoleID

    DoCmd.OpenForm "Dashboard"
    DoCmd.close acForm, Me.Name


End Sub
