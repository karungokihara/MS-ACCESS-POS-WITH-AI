Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Compare Database
Private Sub Form_Load()

    SeedQuickAddField Me, "ManufacturerName"

End Sub
Private Sub cmdSaveClose_Click()

    On Error GoTo ErrorHandler

    If Me.Dirty Then
        Me.Dirty = False
    End If

    DoCmd.close acForm, Me.Name
    Exit Sub

ErrorHandler:

    MsgBox _
        "The entry could not be saved." & _
        vbCrLf & vbCrLf & Err.Description, _
        vbExclamation, _
        "Save Entry"

End Sub
