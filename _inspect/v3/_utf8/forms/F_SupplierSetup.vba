Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Compare Database
Option Explicit

'=============================================================
' FORM: F_SupplierSetup
'
' EDITED FROM THE SUPPLIED DATABASE MODULE.
' Existing quick-add seeding and save/close behaviour are kept.
'
' SUPPLIER CODE FORMAT:
'   First four letters from CompanyName plus a five-digit
'   sequence for that prefix, for example LAWI-00001.
'=============================================================


'=============================================================
' FORM EVENTS
'=============================================================

Private Sub Form_Load()

    On Error GoTo Err_Handler

    'When opened by a combo-box NotInList event, this fills in
    'the company name that the user typed.
    SeedQuickAddField Me, "CompanyName"

    On Error Resume Next
    Me.SupplierCode.Locked = True
    Me.SupplierCode.TabStop = False
    On Error GoTo Err_Handler

    If Me.NewRecord Then
        EnsureSupplierCode
    End If

Exit_Handler:
    Exit Sub

Err_Handler:
    MsgBox _
        "Supplier Setup could not be initialized." & _
        vbCrLf & vbCrLf & _
        "Error Number: " & Err.Number & vbCrLf & _
        "Details: " & Err.Description, _
        vbExclamation, _
        "Supplier Setup"
    Resume Exit_Handler

End Sub


Private Sub Form_BeforeUpdate(Cancel As Integer)

    On Error GoTo Err_Handler

    NormalizeSupplierValues
    EnsureSupplierCode

    If Not ValidateSupplierRecord Then
        Cancel = True
    End If

Exit_Handler:
    Exit Sub

Err_Handler:
    Cancel = True
    MsgBox _
        "The supplier could not be validated." & _
        vbCrLf & vbCrLf & _
        "Error Number: " & Err.Number & vbCrLf & _
        "Details: " & Err.Description, _
        vbExclamation, _
        "Supplier Validation"
    Resume Exit_Handler

End Sub


Private Sub CompanyName_AfterUpdate()

    On Error GoTo Err_Handler

    If Me.NewRecord Then
        EnsureSupplierCode
    End If

Exit_Handler:
    Exit Sub

Err_Handler:
    MsgBox _
        "The supplier code could not be generated." & _
        vbCrLf & vbCrLf & _
        Err.Description, _
        vbExclamation, _
        "Supplier Code"
    Resume Exit_Handler

End Sub


'=============================================================
' EXISTING SAVE BUTTON - VALIDATE, SAVE AND CLOSE
'=============================================================

Private Sub Save_Click()

    On Error GoTo ErrorHandler

    NormalizeSupplierValues
    EnsureSupplierCode

    If Not ValidateSupplierRecord Then
        Exit Sub
    End If

    If Me.Dirty Then
        Me.Dirty = False
    End If

    DoCmd.close acForm, Me.Name
    Exit Sub

ErrorHandler:

    MsgBox _
        "The entry could not be saved." & _
        vbCrLf & vbCrLf & _
        "Error Number: " & Err.Number & vbCrLf & _
        "Details: " & Err.Description, _
        vbExclamation, _
        "Save Entry"

End Sub


'=============================================================
' GENERATED SUPPLIER CODE
'=============================================================

Private Sub EnsureSupplierCode()

    If Not Me.NewRecord Then
        Exit Sub
    End If

    If Len(Trim$(Nz(Me.CompanyName, vbNullString))) = 0 Then
        Me.SupplierCode = Null
    Else
        Me.SupplierCode = GenerateSupplierCode(Me.CompanyName)
    End If

End Sub


Private Function GenerateSupplierCode( _
    ByVal sourceName As String _
) As String

    Dim prefix As String
    Dim nextNumber As Long
    Dim existingMaximum As Variant
    Dim candidate As String

    prefix = FourCharacterPrefix(sourceName, "SUPP")

    existingMaximum = DMax( _
        "Val(Mid([SupplierCode],6))", _
        "T_Suppliers", _
        "SupplierCode Like " & SqlText(prefix & "-*") & _
        " AND IsNumeric(Mid([SupplierCode],6))")

    If IsNull(existingMaximum) Then
        nextNumber = 1
    Else
        nextNumber = CLng(existingMaximum) + 1
    End If

    Do
        candidate = prefix & "-" & Format$(nextNumber, "00000")

        If DCount( _
            "*", _
            "T_Suppliers", _
            "SupplierCode=" & SqlText(candidate)) = 0 Then
            Exit Do
        End If

        nextNumber = nextNumber + 1
    Loop

    GenerateSupplierCode = candidate

End Function


Private Function FourCharacterPrefix( _
    ByVal sourceText As String, _
    ByVal fallbackPrefix As String _
) As String

    Dim resultText As String
    Dim oneCharacter As String
    Dim position As Long

    sourceText = UCase$(Trim$(sourceText))

    For position = 1 To Len(sourceText)
        oneCharacter = Mid$(sourceText, position, 1)

        If oneCharacter Like "[A-Z]" Then
            resultText = resultText & oneCharacter

            If Len(resultText) = 4 Then
                Exit For
            End If
        End If
    Next position

    If Len(resultText) = 0 Then
        resultText = Left$(fallbackPrefix & "XXXX", 4)
    ElseIf Len(resultText) < 4 Then
        resultText = resultText & _
            String$(4 - Len(resultText), "X")
    End If

    FourCharacterPrefix = resultText

End Function


'=============================================================
' NORMALIZE AND VALIDATE
'=============================================================

Private Sub NormalizeSupplierValues()

    Me.CompanyName = Trim$(Nz(Me.CompanyName, vbNullString))

    Me.SupplierCode = _
        UCase$(Trim$(Nz(Me.SupplierCode, vbNullString)))

End Sub


Private Function ValidateSupplierRecord() As Boolean

    Dim currentSupplierID As Long
    Dim criteria As String

    ValidateSupplierRecord = False
    currentSupplierID = Nz(Me.SupplierID, 0)

    If Len(Trim$(Nz(Me.CompanyName, vbNullString))) = 0 Then
        MsgBox _
            "Enter the supplier company name.", _
            vbExclamation, _
            "Company Name Required"

        Me.CompanyName.SetFocus
        Exit Function
    End If

    If Len(Trim$(Nz(Me.SupplierCode, vbNullString))) = 0 Then
        MsgBox _
            "The supplier code could not be generated.", _
            vbExclamation, _
            "Supplier Code"
        Exit Function
    End If

    criteria = _
        "SupplierCode=" & SqlText(Me.SupplierCode) & _
        " AND SupplierID<>" & currentSupplierID

    If DCount("*", "T_Suppliers", criteria) > 0 Then
        MsgBox _
            "The supplier code already exists.", _
            vbExclamation, _
            "Duplicate Supplier Code"
        Exit Function
    End If

    criteria = _
        "CompanyName=" & SqlText(Me.CompanyName) & _
        " AND SupplierID<>" & currentSupplierID

    If DCount("*", "T_Suppliers", criteria) > 0 Then
        MsgBox _
            "A supplier with this company name already exists.", _
            vbExclamation, _
            "Duplicate Supplier"

        Me.CompanyName.SetFocus
        Exit Function
    End If

    ValidateSupplierRecord = True

End Function


Private Function SqlText(ByVal value As String) As String

    SqlText = "'" & Replace(value, "'", "''") & "'"

End Function
