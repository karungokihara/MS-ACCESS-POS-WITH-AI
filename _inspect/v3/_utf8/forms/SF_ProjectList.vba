Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Compare Database
Option Explicit

'=============================================================
' FORM: SF_ProjectList
'
' PURPOSE:
'   Displays the project list.
'
'   Double-clicking any project row opens the exact selected
'   project in F_Projects.
'
' RECORD SOURCE:
'   Q_ProjectList
'
' REQUIRED FIELD:
'   ProjectID
'
' DOUBLE-CLICK PROPERTY:
'   Set the On Dbl Click property of each visible control to:
'
'       =OpenSelectedProject()
'=============================================================

Private Const PROJECT_DETAIL_FORM As String = "F_Projects"
Private Const PROJECT_TABLE As String = "T_Projects"
Private Const PROJECT_ID_FIELD As String = "ProjectID"


'=============================================================
' OPEN SELECTED PROJECT
'=============================================================

Public Function OpenSelectedProject() As Boolean

    On Error GoTo Err_Handler

    Dim projectIDValue As Long
    Dim openedProjectID As Long

    OpenSelectedProject = False

    'Do not continue when the list is on its blank new row.
    If Me.NewRecord Then
        MsgBox _
            "Select an existing project first.", _
            vbInformation, _
            "Project Required"

        Exit Function
    End If

    'Read ProjectID directly from the selected datasheet row.
    projectIDValue = Nz( _
        Me.Recordset.Fields(PROJECT_ID_FIELD).value, _
        0)

    If projectIDValue <= 0 Then
        MsgBox _
            "The selected row does not contain a valid ProjectID.", _
            vbExclamation, _
            "Invalid Project"

        Exit Function
    End If

    'Make sure the project still exists in T_Projects.
    If DCount( _
        "*", _
        PROJECT_TABLE, _
        "[" & PROJECT_ID_FIELD & "]=" & _
            CStr(projectIDValue)) = 0 Then

        MsgBox _
            "Project ID " & CStr(projectIDValue) & _
            " could not be found in T_Projects." & _
            vbCrLf & vbCrLf & _
            "Refresh the project list and try again.", _
            vbExclamation, _
            "Project Not Found"

        Exit Function
    End If

    'The detail form must be the actual project form.
    If Not FormExists(PROJECT_DETAIL_FORM) Then

        MsgBox _
            "The project detail form '" & _
            PROJECT_DETAIL_FORM & "' could not be found.", _
            vbExclamation, _
            "Project Form Missing"

        Exit Function
    End If

    'If F_Projects is already open, close it first.
    '
    'This is important because calling OpenForm again does not
    'reliably apply a different WhereCondition to an open form.
    If IsFormLoaded(PROJECT_DETAIL_FORM) Then

        If Not CloseLoadedProjectForm Then
            Exit Function
        End If

    End If

    'Open only the selected ProjectID.
    DoCmd.OpenForm _
        formName:=PROJECT_DETAIL_FORM, _
        View:=acNormal, _
        WhereCondition:= _
            "[" & PROJECT_ID_FIELD & "]=" & _
            CStr(projectIDValue), _
        DataMode:=acFormEdit, _
        WindowMode:=acWindowNormal

    'Confirm that Access loaded the requested project.
    If Not IsFormLoaded(PROJECT_DETAIL_FORM) Then

        Err.Raise _
            vbObjectError + 7601, _
            "OpenSelectedProject", _
            "F_Projects did not open."

    End If

    openedProjectID = Nz( _
        Forms(PROJECT_DETAIL_FORM).Recordset.Fields( _
            PROJECT_ID_FIELD).value, _
        0)

    If openedProjectID <> projectIDValue Then

        Err.Raise _
            vbObjectError + 7602, _
            "OpenSelectedProject", _
            "Access opened Project ID " & _
            CStr(openedProjectID) & _
            " instead of Project ID " & _
            CStr(projectIDValue) & "."

    End If

    Forms(PROJECT_DETAIL_FORM).SetFocus

    OpenSelectedProject = True

Exit_Handler:
    Exit Function

Err_Handler:

    MsgBox _
        "The selected project could not be opened." & _
        vbCrLf & vbCrLf & _
        "Selected ProjectID: " & _
            CStr(projectIDValue) & vbCrLf & vbCrLf & _
        "Error Number: " & Err.Number & vbCrLf & _
        "Details: " & Err.Description, _
        vbExclamation, _
        "Open Project"

    OpenSelectedProject = False
    Resume Exit_Handler

End Function


'=============================================================
' CLOSE AN ALREADY-OPEN PROJECT FORM
'=============================================================

Private Function CloseLoadedProjectForm() As Boolean

    On Error GoTo Err_Handler

    Dim projectForm As Access.Form

    CloseLoadedProjectForm = False

    If Not IsFormLoaded(PROJECT_DETAIL_FORM) Then
        CloseLoadedProjectForm = True
        Exit Function
    End If

    Set projectForm = Forms(PROJECT_DETAIL_FORM)

    If projectForm.Dirty Then

        'F_Projects automatically enters defaults such as the
        'project code, date and status on a blank new record.
        '
        'Do not try to save that blank default record.
        If projectForm.NewRecord And _
           IsBlankNewProject(projectForm) Then

            projectForm.Undo

        Else

            'Save a genuine edited or newly entered project before
            'opening the project selected from the list.
            '
            'This runs the existing validation code in F_Projects.
            projectForm.Dirty = False

        End If

    End If

    DoCmd.close _
        acForm, _
        PROJECT_DETAIL_FORM, _
        acSaveNo

    CloseLoadedProjectForm = _
        Not IsFormLoaded(PROJECT_DETAIL_FORM)

Exit_Handler:
    Set projectForm = Nothing
    Exit Function

Err_Handler:

    MsgBox _
        "The currently open project could not be saved or closed." & _
        vbCrLf & vbCrLf & _
        "Correct or cancel the current project record, then " & _
        "double-click the required project again." & _
        vbCrLf & vbCrLf & _
        "Error Number: " & Err.Number & vbCrLf & _
        "Details: " & Err.Description, _
        vbExclamation, _
        "Current Project Still Open"

    CloseLoadedProjectForm = False
    Resume Exit_Handler

End Function


'=============================================================
' CHECK FOR A BLANK DEFAULT NEW PROJECT
'=============================================================

Private Function IsBlankNewProject( _
    ByVal projectForm As Access.Form _
) As Boolean

    On Error GoTo Not_Blank

    IsBlankNewProject = _
        (Len(Trim$(Nz( _
            projectForm.Controls( _
                "txtProjectName").value, _
            vbNullString))) = 0) And _
        (Nz(projectForm.Controls( _
            "cboCustomerID").value, 0) = 0) And _
        (Nz(projectForm.Controls( _
            "cboAssemblyTemplateID").value, 0) = 0) And _
        (Nz(projectForm.Controls( _
            "cboOutputLocationID").value, 0) = 0) And _
        (Len(Trim$(Nz( _
            projectForm.Controls( _
                "txtCustomerPOReference").value, _
            vbNullString))) = 0) And _
        (Len(Trim$(Nz( _
            projectForm.Controls( _
                "txtWorkInstructions").value, _
            vbNullString))) = 0) And _
        (Len(Trim$(Nz( _
            projectForm.Controls( _
                "txtNOtes").value, _
            vbNullString))) = 0)

Exit_Handler:
    Exit Function

Not_Blank:
    IsBlankNewProject = False
    Resume Exit_Handler

End Function


'=============================================================
' FORM HELPERS
'=============================================================

Private Function FormExists( _
    ByVal formName As String _
) As Boolean

    On Error GoTo Not_Found

    Dim formObject As AccessObject

    Set formObject = CurrentProject.AllForms(formName)

    FormExists = True

Exit_Handler:
    Set formObject = Nothing
    Exit Function

Not_Found:
    FormExists = False
    Resume Exit_Handler

End Function


Private Function IsFormLoaded( _
    ByVal formName As String _
) As Boolean

    On Error GoTo Not_Loaded

    IsFormLoaded = _
        CurrentProject.AllForms(formName).IsLoaded

    Exit Function

Not_Loaded:
    IsFormLoaded = False

End Function
