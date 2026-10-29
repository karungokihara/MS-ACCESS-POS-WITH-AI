Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Compare Database
Option Explicit

'=============================================================
' FORM: F_LocationSetup
'
' Creates and edits stock locations.
'
' RULES
'   Location Name       required
'   Warehouse           OPTIONAL - a location may exist before
'                       any warehouse is created
'   Bin / Shelf         always optional
'   Location Code       generated automatically
'
' QUICK ADD
'   F_ItemSetup opens this form with OpenArgs "QUICKADD|<name>".
'   The typed name is filled in, saved, and the form closes so
'   the calling combo can select the new location.
'=============================================================

Private Const WAREHOUSE_TABLE As String = "T_Warehouses"
Private Const WAREHOUSE_FORM As String = "F_Warehouses"

Private mInitializing As Boolean
Private mQuickAddMode As Boolean
Private mCloseAfterLoad As Boolean
Private mLastSavedID As Long


'=============================================================
' FORM EVENTS
'=============================================================

Private Sub Form_Load()

    On Error GoTo Err_Handler

    mInitializing = True
    mQuickAddMode = IsQuickAddRequest()

    ConfigureFormControls
    ConfigureWarehouseCombo
    MarkRequiredFields

    If Me.NewRecord Then
        ApplyNewRecordDefaults
        ApplyQuickAddLocationName
        EnsureLocationCode
    End If

    mInitializing = False

    SetInitialFocus

Exit_Handler:
    mInitializing = False
    Exit Sub

Err_Handler:
    mInitializing = False

    MsgBox _
        "Location Setup could not be initialized." & _
        vbCrLf & vbCrLf & _
        "Error Number: " & Err.Number & vbCrLf & _
        "Details: " & Err.Description, _
        vbExclamation, _
        "Location Setup"

    Resume Exit_Handler

End Sub


Private Sub Form_Current()

    On Error Resume Next

    If Me.NewRecord Then
        ApplyNewRecordDefaults
        EnsureLocationCode
    End If

    On Error GoTo 0

End Sub


Private Sub Form_BeforeUpdate(Cancel As Integer)

    On Error GoTo Err_Handler

    NormalizeLocationValues
    EnsureLocationCode

    If Not ValidateLocationRecord() Then
        Cancel = True
        Exit Sub
    End If

Exit_Handler:
    Exit Sub

Err_Handler:

    Cancel = True

    MsgBox _
        "The location could not be validated." & _
        vbCrLf & vbCrLf & _
        "Error Number: " & Err.Number & vbCrLf & _
        "Details: " & Err.Description, _
        vbExclamation, _
        "Location Setup"

    Resume Exit_Handler

End Sub


Private Sub Form_Timer()

    On Error Resume Next

    Me.TimerInterval = 0

    If mCloseAfterLoad Then
        mCloseAfterLoad = False
        DoCmd.Close acForm, Me.Name, acSaveNo
    End If

    On Error GoTo 0

End Sub


Private Sub Form_Error( _
    DataErr As Integer, _
    Response As Integer _
)

    If DataErr = 3022 Then

        MsgBox _
            "That location code already exists. " & _
            "Clear the Location Code field and it will be " & _
            "generated again.", _
            vbExclamation, _
            "Duplicate Location"

        Response = acDataErrContinue

    Else

        Response = acDataErrDisplay

    End If

End Sub


'=============================================================
' WAREHOUSE COMBO
'=============================================================

Private Sub cboWarehouseID_NotInList( _
    NewData As String, _
    Response As Integer _
)

    On Error GoTo Err_Handler

    Response = acDataErrContinue

    If HandleComboNotInList( _
        Me.cboWarehouseID, _
        NewData, _
        WAREHOUSE_FORM, _
        WAREHOUSE_TABLE, _
        "WarehouseID", _
        "WarehouseName", _
        "WarehouseCode") Then

        Response = acDataErrAdded

    Else

        Me.cboWarehouseID.Undo

    End If

Exit_Handler:
    Exit Sub

Err_Handler:

    Response = acDataErrContinue

    MsgBox _
        "The warehouse could not be created." & _
        vbCrLf & vbCrLf & _
        "Error Number: " & Err.Number & vbCrLf & _
        "Details: " & Err.Description, _
        vbExclamation, _
        "Warehouse Quick Add"

    On Error Resume Next
    Me.cboWarehouseID.Undo
    On Error GoTo 0

    Resume Exit_Handler

End Sub


Private Sub cboWarehouseID_DblClick(Cancel As Integer)

    On Error GoTo Err_Handler

    Cancel = True

    If Nz(Me.cboWarehouseID, 0) > 0 Then

        EditComboEntry _
            Me.cboWarehouseID, _
            WAREHOUSE_FORM, _
            "WarehouseID"

    Else

        AddComboEntry _
            Me.cboWarehouseID, _
            WAREHOUSE_FORM, _
            WAREHOUSE_TABLE, _
            "WarehouseID"

    End If

    Me.cboWarehouseID.Requery

Exit_Handler:
    Exit Sub

Err_Handler:

    MsgBox _
        "Warehouse Setup could not be opened." & _
        vbCrLf & vbCrLf & _
        "Error Number: " & Err.Number & vbCrLf & _
        "Details: " & Err.Description, _
        vbExclamation, _
        "Warehouse Setup"

    Resume Exit_Handler

End Sub


'=============================================================
' LOCATION NAME
'=============================================================

Private Sub LocationName_AfterUpdate()

    If Not mInitializing Then

        If Me.NewRecord Then
            EnsureLocationCode
        End If

    End If

End Sub


'=============================================================
' BUTTONS
'=============================================================

Private Sub cmdSaveClose_Click()

    SaveAndCloseLocation

End Sub


Private Sub cmdSave_Click()

    SaveAndCloseLocation

End Sub


Private Sub Save_Click()

    SaveAndCloseLocation

End Sub


'=============================================================
' SAVE
'=============================================================

Private Sub SaveAndCloseLocation()

    On Error GoTo Err_Handler

    If Not SaveLocationRecord() Then
        Exit Sub
    End If

    DoCmd.Close acForm, Me.Name, acSaveNo

Exit_Handler:
    Exit Sub

Err_Handler:

    MsgBox _
        "The location could not be saved and closed." & _
        vbCrLf & vbCrLf & _
        "Error Number: " & Err.Number & vbCrLf & _
        "Details: " & Err.Description, _
        vbExclamation, _
        "Location Setup"

    Resume Exit_Handler

End Sub


Private Function SaveLocationRecord() As Boolean

    Dim messageText As String
    Dim errorNumber As Long
    Dim errorText As String

    On Error GoTo Err_Handler

    SaveLocationRecord = False

    NormalizeLocationValues
    EnsureLocationCode

    If Not ValidateLocationRecord() Then
        Exit Function
    End If

    If Me.Dirty Then
        Me.Dirty = False
    End If

    If Me.NewRecord Then
        Exit Function
    End If

    mLastSavedID = CurrentLocationID()

    If mLastSavedID <= 0 Then
        Exit Function
    End If

    SaveLocationRecord = True

Exit_Handler:
    Exit Function

Err_Handler:

    errorNumber = Err.Number
    errorText = Err.Description

    On Error Resume Next

    If errorNumber = 3021 Then
        Me.Requery
        messageText = _
            "The location form lost its current record while saving." & _
            vbCrLf & vbCrLf & _
            "The form has been refreshed. Enter the location name " & _
            "and save again."
    Else
        messageText = _
            "The location record could not be saved." & _
            vbCrLf & vbCrLf & _
            "Error Number: " & errorNumber & vbCrLf & _
            "Details: " & errorText
    End If

    MsgBox messageText, vbExclamation, "Location Setup"

    Resume Exit_Handler

End Function


'=============================================================
' DEFAULTS
'=============================================================

Private Sub ApplyNewRecordDefaults()

    If Not Me.NewRecord Then
        Exit Sub
    End If

    'Only select a warehouse automatically when exactly one exists.
    If IsNull(Me.cboWarehouseID) Then

        If DCount("*", WAREHOUSE_TABLE) = 1 Then
            Me.cboWarehouseID = DLookup("WarehouseID", WAREHOUSE_TABLE)
        End If

    End If

End Sub


Private Sub MarkRequiredFields()

    UX_MarkRequired Me, "LocationName"

    On Error Resume Next

    Me.Controls("cboWarehouseID").Controls(0).Caption = _
        "Warehouse (optional)"

    On Error GoTo 0

End Sub


'=============================================================
' QUICK ADD
'=============================================================

Private Function IsQuickAddRequest() As Boolean

    IsQuickAddRequest = _
        (StrComp( _
            Left$(Trim$(Nz(Me.OpenArgs, vbNullString)), 9), _
            "QUICKADD|", _
            vbTextCompare) = 0)

End Function


Private Function QuickAddTypedText() As String

    Dim openArguments As String

    openArguments = Trim$(Nz(Me.OpenArgs, vbNullString))

    If StrComp(Left$(openArguments, 9), "QUICKADD|", vbTextCompare) = 0 Then
        QuickAddTypedText = Trim$(Mid$(openArguments, 10))
    End If

End Function


Private Sub ApplyQuickAddLocationName()

    Dim typedLocationName As String
    Dim existingID As Variant
    Dim recordsetObject As DAO.Recordset

    typedLocationName = QuickAddTypedText()

    If Len(typedLocationName) = 0 Then Exit Sub

    'Reuse an existing location with the same name instead of
    'creating a duplicate row.
    existingID = DLookup( _
        "LocationID", _
        "T_Locations", _
        "StrComp(Trim(Nz([LocationName],''))," & _
            SqlText(typedLocationName) & ",1)=0")

    If Not IsNull(existingID) Then

        mLastSavedID = CLng(existingID)
        Me.Undo

        Set recordsetObject = Me.RecordsetClone
        recordsetObject.FindFirst "LocationID=" & mLastSavedID

        If Not recordsetObject.NoMatch Then
            Me.Bookmark = recordsetObject.Bookmark
        End If

        recordsetObject.Close
        Set recordsetObject = Nothing

        'The form cannot close while it is still loading.
        mCloseAfterLoad = True
        Me.TimerInterval = 50
        Exit Sub

    End If

    If ControlExists("LocationName") Then

        Me.Controls("LocationName").Value = typedLocationName
        EnsureLocationCode

    End If

End Sub


'=============================================================
' FORM SETUP
'=============================================================

Private Sub ConfigureFormControls()

    On Error Resume Next

    'Hide only the primary key.
    If ControlExists("LocationID") Then

        If ControlSourceIs(Me.Controls("LocationID"), "LocationID") Then
            Me.Controls("LocationID").Visible = False
            Me.Controls("LocationID").Enabled = False
            Me.Controls("LocationID").Locked = True
            Me.Controls("LocationID").TabStop = False
        End If

    End If

    'The code is always generated and never typed.
    If ControlExists("LocationCode") Then

        With Me.Controls("LocationCode")
            .Visible = True
            .Enabled = True
            .Locked = True
            .TabStop = False
        End With

    End If

    'Bin and shelf are optional and must stay editable.
    If ControlExists("BinNumber") Then
        Me.Controls("BinNumber").Enabled = True
        Me.Controls("BinNumber").Locked = False
    End If

    If ControlExists("ShelfNumber") Then
        Me.Controls("ShelfNumber").Enabled = True
        Me.Controls("ShelfNumber").Locked = False
    End If

    On Error GoTo 0

End Sub


Private Sub ConfigureWarehouseCombo()

    On Error Resume Next

    If Not ControlExists("cboWarehouseID") Then Exit Sub

    With Me.cboWarehouseID

        .ControlSource = "WarehouseID"
        .RowSourceType = "Table/Query"
        .RowSource = _
            "SELECT WarehouseID, WarehouseName, WarehouseCode " & _
            "FROM T_Warehouses " & _
            "ORDER BY WarehouseName, WarehouseCode;"
        .BoundColumn = 1
        .ColumnCount = 3
        .ColumnWidths = "0cm;6cm;3cm"
        .LimitToList = False
        .AutoExpand = True
        .Visible = True
        .Enabled = True
        .Locked = False
        .TabStop = True
        .Requery

    End With

    On Error GoTo 0

End Sub


'=============================================================
' NORMALIZATION AND VALIDATION
'=============================================================

Private Sub NormalizeLocationValues()

    On Error Resume Next

    If ControlExists("LocationName") Then
        Me.Controls("LocationName").Value = _
            Trim$(Nz(Me.Controls("LocationName").Value, vbNullString))
    End If

    If ControlExists("LocationCode") Then
        Me.Controls("LocationCode").Value = _
            UCase$(Trim$(Nz(Me.Controls("LocationCode").Value, vbNullString)))
    End If

    If ControlExists("BinNumber") Then
        Me.Controls("BinNumber").Value = _
            Trim$(Nz(Me.Controls("BinNumber").Value, vbNullString))
    End If

    If ControlExists("ShelfNumber") Then
        Me.Controls("ShelfNumber").Value = _
            Trim$(Nz(Me.Controls("ShelfNumber").Value, vbNullString))
    End If

    On Error GoTo 0

End Sub


Private Function ValidateLocationRecord() As Boolean

    Dim locationNameValue As String
    Dim locationCodeValue As String
    Dim locationIDValue As Long

    ValidateLocationRecord = False

    locationNameValue = vbNullString

    If ControlExists("LocationName") Then
        locationNameValue = _
            Trim$(Nz(Me.Controls("LocationName").Value, vbNullString))
    End If

    If Len(locationNameValue) = 0 Then

        MsgBox _
            "Enter the location name.", _
            vbExclamation, _
            "Location Name Required"

        If ControlExists("LocationName") Then
            Me.Controls("LocationName").SetFocus
        End If

        Exit Function

    End If

    locationCodeValue = vbNullString

    If ControlExists("LocationCode") Then
        locationCodeValue = _
            UCase$(Trim$(Nz(Me.Controls("LocationCode").Value, vbNullString)))
    End If

    If Len(locationCodeValue) = 0 Then

        MsgBox _
            "The location code could not be generated.", _
            vbExclamation, _
            "Location Code Required"

        Exit Function

    End If

    locationIDValue = CurrentLocationID()

    If DCount( _
        "*", _
        "T_Locations", _
        "[LocationCode]=" & SqlText(locationCodeValue) & _
        " AND [LocationID]<>" & CStr(locationIDValue)) > 0 Then

        MsgBox _
            "The location code already exists." & vbCrLf & vbCrLf & _
            "Change the location name or clear the code.", _
            vbExclamation, _
            "Duplicate Location Code"

        Exit Function

    End If

    ValidateLocationRecord = True

End Function


'=============================================================
' LOCATION CODE
'=============================================================

Private Sub EnsureLocationCode()

    Dim locationNameValue As String
    Dim locationCodeValue As String

    On Error Resume Next

    If Not Me.NewRecord Then Exit Sub

    If Not ControlExists("LocationName") Then Exit Sub
    If Not ControlExists("LocationCode") Then Exit Sub

    locationNameValue = _
        Trim$(Nz(Me.Controls("LocationName").Value, vbNullString))

    If Len(locationNameValue) = 0 Then Exit Sub

    locationCodeValue = _
        Trim$(Nz(Me.Controls("LocationCode").Value, vbNullString))

    If Len(locationCodeValue) > 0 Then Exit Sub

    Me.Controls("LocationCode").Value = _
        GenerateLocationCode(locationNameValue)

    On Error GoTo 0

End Sub


Private Function GenerateLocationCode( _
    ByVal locationName As String _
) As String

    Dim prefix As String
    Dim nextNumber As Long
    Dim existingMaximum As Variant
    Dim candidate As String

    prefix = FourCharacterPrefix(locationName, "LOCA")

    existingMaximum = DMax( _
        "Val(Mid([LocationCode],6))", _
        "T_Locations", _
        "[LocationCode] Like " & SqlText(prefix & "-*"))

    If IsNull(existingMaximum) Then
        nextNumber = 1
    Else
        nextNumber = CLng(existingMaximum) + 1
    End If

    Do

        candidate = prefix & "-" & Format$(nextNumber, "00000")

        If DCount( _
            "*", _
            "T_Locations", _
            "[LocationCode]=" & SqlText(candidate)) = 0 Then

            Exit Do

        End If

        nextNumber = nextNumber + 1

    Loop

    GenerateLocationCode = candidate

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

            If Len(resultText) = 4 Then Exit For

        End If

    Next position

    If Len(resultText) = 0 Then

        resultText = Left$(fallbackPrefix & "XXXX", 4)

    ElseIf Len(resultText) < 4 Then

        resultText = resultText & String$(4 - Len(resultText), "X")

    End If

    FourCharacterPrefix = resultText

End Function


'=============================================================
' SAFE HELPERS
'=============================================================

Private Function ControlExists(ByVal controlName As String) As Boolean

    Dim controlObject As Access.Control

    On Error GoTo Not_Found

    Set controlObject = Me.Controls(controlName)
    ControlExists = True

Exit_Handler:
    Set controlObject = Nothing
    Exit Function

Not_Found:
    ControlExists = False
    Resume Exit_Handler

End Function


Private Function ControlSourceIs( _
    ByVal controlObject As Access.Control, _
    ByVal expectedSource As String _
) As Boolean

    On Error GoTo Not_Match

    ControlSourceIs = _
        (StrComp( _
            Replace(Replace( _
                Trim$(Nz(controlObject.ControlSource, vbNullString)), _
                "[", vbNullString), "]", vbNullString), _
            expectedSource, _
            vbTextCompare) = 0)

    Exit Function

Not_Match:
    ControlSourceIs = False

End Function


Private Function HasCurrentRecord() As Boolean

    On Error Resume Next

    If Me.NewRecord Then Exit Function

    If Me.Recordset.RecordCount < 1 Then Exit Function
    If IsNull(Me.Recordset.Fields("LocationID").Value) Then Exit Function

    HasCurrentRecord = True

    On Error GoTo 0

End Function


Private Function CurrentLocationID() As Long

    On Error Resume Next

    CurrentLocationID = 0

    If Not HasCurrentRecord() Then Exit Function

    CurrentLocationID = _
        Nz(Me.Recordset.Fields("LocationID").Value, 0)

    On Error GoTo 0

End Function


Private Sub SetInitialFocus()

    On Error Resume Next

    If ControlExists("LocationName") Then

        If Len(Trim$(Nz(Me.Controls("LocationName").Value, vbNullString))) = 0 Then
            Me.Controls("LocationName").SetFocus
            Exit Sub
        End If

    End If

    If ControlExists("cboWarehouseID") Then
        Me.cboWarehouseID.SetFocus
    End If

    On Error GoTo 0

End Sub


Private Function SqlText(ByVal value As String) As String

    SqlText = "'" & Replace(value, "'", "''") & "'"

End Function
