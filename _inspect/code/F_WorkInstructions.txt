Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Compare Database

Private Sub cmdBrowsePDF_Click()

    Dim fd As Object

    Set fd = Application.FileDialog(3)

    With fd
        .Title = "Select Work Instruction PDF"
        .AllowMultiSelect = False
        .Filters.Clear
        .Filters.Add "PDF Files", "*.pdf"

        If .Show = -1 Then
            Me.FilePath = .SelectedItems(1)
        End If
    End With

    Set fd = Nothing

End Sub

Private Sub cmdOpenPDF_Click()

If Nz(Me.FilePath, "") = "" Then
        MsgBox "Please select a PDF first.", vbInformation
        Exit Sub
    End If

    Application.FollowHyperlink Me.FilePath


End Sub
