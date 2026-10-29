Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Compare Database
Option Explicit

'=============================================================
' FORM: SF_TemplateList
'
' PURPOSE:
'   Datasheet list of assembly/project templates.
'
'   Double-clicking a template row opens F_AssemblyTemplate
'   filtered to the selected AssemblyTemplateID.
'
' RECORD SOURCE:
'   Q_TemplateList
'
' REQUIRED FIELD:
'   AssemblyTemplateID
'
' IMPORTANT:
'   On each visible textbox/control in the datasheet, set:
'       On Dbl Click: =OpenSelectedTemplate()
'=============================================================

Private Const DETAIL_FORM_NAME As String = "F_AssemblyTemplate"


Public Function OpenSelectedTemplate() As Boolean

    On Error GoTo Err_Handler

    Dim templateIDValue As Long
    Dim whereText As String

    OpenSelectedTemplate = False

    '---------------------------------------------------------
    ' Make sure current row has a valid template ID
    '---------------------------------------------------------
    If IsNull(Me!AssemblyTemplateID) Then
        Exit Function
    End If

    templateIDValue = CLng(Me!AssemblyTemplateID)

    If templateIDValue <= 0 Then
        Exit Function
    End If

    '---------------------------------------------------------
    ' Confirm the selected template still exists in the table
    '---------------------------------------------------------
    If DCount("*", "T_AssemblyTemplates", "AssemblyTemplateID=" & templateIDValue) = 0 Then

        MsgBox _
            "The selected template no longer exists in T_AssemblyTemplates." & vbCrLf & vbCrLf & _
            "AssemblyTemplateID: " & templateIDValue, _
            vbExclamation, _
            "Template Not Found"

        Exit Function

    End If

    '---------------------------------------------------------
    ' Confirm detail form exists
    '---------------------------------------------------------
    If Not FormExists(DETAIL_FORM_NAME) Then

        MsgBox _
            "The template detail form was not found." & vbCrLf & vbCrLf & _
            "Expected form: " & DETAIL_FORM_NAME, _
            vbExclamation, _
            "Template Form Missing"

        Exit Function

    End If

    whereText = "[AssemblyTemplateID]=" & templateIDValue

    '---------------------------------------------------------
    ' Important:
    ' Close the form first if already open.
    ' This prevents Access from showing the previous/new/blank
    ' record instead of the selected one.
    '---------------------------------------------------------
    If CurrentProject.AllForms(DETAIL_FORM_NAME).IsLoaded Then
        DoCmd.close acForm, DETAIL_FORM_NAME, acSaveNo
    End If

    '---------------------------------------------------------
    ' Open detail form filtered to the selected template
    '---------------------------------------------------------
    DoCmd.OpenForm _
        DETAIL_FORM_NAME, _
        acNormal, _
        , _
        whereText, _
        acFormEdit

    OpenSelectedTemplate = True

Exit_Handler:
    Exit Function

Err_Handler:

    MsgBox _
        "The selected template could not be opened." & vbCrLf & vbCrLf & _
        "AssemblyTemplateID: " & Nz(Me!AssemblyTemplateID, "Null") & vbCrLf & _
        "Filter: [AssemblyTemplateID]=" & Nz(Me!AssemblyTemplateID, "Null") & vbCrLf & vbCrLf & _
        "Error Number: " & Err.Number & vbCrLf & _
        "Details: " & Err.Description, _
        vbExclamation, _
        "Open Template"

    OpenSelectedTemplate = False
    Resume Exit_Handler

End Function


Private Function FormExists(ByVal formName As String) As Boolean

    On Error GoTo Err_Handler

    Dim tmp As String

    tmp = CurrentProject.AllForms(formName).Name

    FormExists = True

Exit_Handler:
    Exit Function

Err_Handler:
    FormExists = False
    Resume Exit_Handler

End Function
