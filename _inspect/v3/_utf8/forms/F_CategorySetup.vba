Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Compare Database

Private Sub Form_Load()

    Dim args() As String

    SeedQuickAddField Me, "CategoryName"

    If Len(Nz(Me.OpenArgs, "")) > 0 Then

        args = Split(Me.OpenArgs, "|")

        If UBound(args) >= 2 Then

            If args(0) = "QUICKADD" Then
                Me.cboItemTypeID = DLookup("ItemTypeID", "T_ItemTypes", _
                    "ItemTypeName='" & Replace(args(2), "'", "''") & "'")
            End If

        End If

    End If

End Sub
'=============================================================
' ITEM TYPE COMBO
'
' Type a missing name:
'   Adds it directly to T_ItemTypes.
'
' Double-click an existing name:
'   Allows the name to be edited.
'=============================================================

Private Sub cboItemTypeID_NotInList( _
    NewData As String, _
    Response As Integer _
)

    Dim db As DAO.Database
    Dim enteredName As String
    Dim criteria As String

    On Error GoTo ErrorHandler

    Response = acDataErrContinue
    enteredName = Trim$(NewData)

    If Len(enteredName) = 0 Then
        Me!cboItemTypeID.Undo
        Exit Sub
    End If

    If MsgBox( _
        """" & enteredName & """ is not in the Item Type list." & _
        vbCrLf & vbCrLf & _
        "Would you like to add it?", _
        vbQuestion + vbYesNo, _
        "Add Item Type" _
    ) <> vbYes Then

        Me!cboItemTypeID.Undo
        Exit Sub

    End If

    criteria = _
        "ItemTypeName = '" & _
        Replace(enteredName, "'", "''") & "'"

    'Add it only when it does not already exist.
    If DCount("*", "T_ItemTypes", criteria) = 0 Then

        Set db = CurrentDb

        db.Execute _
            "INSERT INTO T_ItemTypes " & _
            "(ItemTypeName) " & _
            "VALUES ('" & _
            Replace(enteredName, "'", "''") & "');", _
            dbFailOnError

    End If

    'Access will requery the combo and select the new value.
    Response = acDataErrAdded

CleanExit:

    Set db = Nothing
    Exit Sub

ErrorHandler:

    Response = acDataErrContinue

    MsgBox _
        "The Item Type could not be added." & _
        vbCrLf & vbCrLf & _
        Err.Number & " - " & Err.Description, _
        vbExclamation, _
        "Item Type"

    Resume CleanExit

End Sub

'=============================================================
' EDIT ITEM TYPE WITHOUT OPENING ANOTHER FORM
'=============================================================

Private Sub EditSelectedItemType()

    Dim db As DAO.Database

    Dim ItemTypeID As Long
    Dim currentName As String
    Dim updatedName As String
    Dim criteria As String

    On Error GoTo ErrorHandler

    If IsNull(Me!cboItemTypeID) Then

        MsgBox _
            "Select an Item Type to edit.", _
            vbInformation, _
            "Edit Item Type"

        Me!cboItemTypeID.SetFocus
        Exit Sub

    End If

    ItemTypeID = CLng(Me!cboItemTypeID)

    currentName = Nz( _
                      DLookup( _
                          "ItemTypeName", _
                          "T_ItemTypes", _
                          "ItemTypeID = " & ItemTypeID _
                      ), _
                      vbNullString _
                  )

    updatedName = Trim$( _
                        InputBox( _
                            "Edit the Item Type name." & _
                            vbCrLf & vbCrLf & _
                            "This change will appear in every category " & _
                            "using this Item Type.", _
                            "Edit Item Type", _
                            currentName _
                        ) _
                    )

    'Blank means the user cancelled or removed the text.
    If Len(updatedName) = 0 Then
        Exit Sub
    End If

    If StrComp( _
        updatedName, _
        currentName, _
        vbTextCompare _
    ) = 0 Then

        Exit Sub

    End If

    criteria = _
        "ItemTypeName = '" & _
        Replace(updatedName, "'", "''") & "'" & _
        " AND ItemTypeID <> " & ItemTypeID

    If DCount("*", "T_ItemTypes", criteria) > 0 Then

        MsgBox _
            "An Item Type named """ & updatedName & _
            """ already exists.", _
            vbExclamation, _
            "Duplicate Item Type"

        Exit Sub

    End If

    If MsgBox( _
        "Change """ & currentName & """ to """ & updatedName & """?", _
        vbQuestion + vbYesNo, _
        "Update Item Type" _
    ) <> vbYes Then

        Exit Sub

    End If

    Set db = CurrentDb

    db.Execute _
        "UPDATE T_ItemTypes " & _
        "SET ItemTypeName = '" & _
        Replace(updatedName, "'", "''") & "' " & _
        "WHERE ItemTypeID = " & ItemTypeID & ";", _
        dbFailOnError

    Me!cboItemTypeID.Requery
    Me!cboItemTypeID = ItemTypeID

    MsgBox _
        "The Item Type was updated.", _
        vbInformation, _
        "Item Type"

CleanExit:

    Set db = Nothing
    Exit Sub

ErrorHandler:

    MsgBox _
        "The Item Type could not be updated." & _
        vbCrLf & vbCrLf & Err.Description, _
        vbExclamation, _
        "Item Type"

    Resume CleanExit

End Sub

Private Sub Save_Click()


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
