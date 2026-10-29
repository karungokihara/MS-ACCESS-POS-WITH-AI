Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Compare Database
Option Explicit

Public Function OpenSelectedRecord() As Boolean

    On Error GoTo Err_Handler

    If Nz(Me!InvoiceID, 0) <= 0 Then Exit Function

    DoCmd.OpenForm "F_Invoices", acNormal, , "InvoiceID = " & CLng(Me!InvoiceID)

    OpenSelectedRecord = True

Exit_Handler:
    Exit Function

Err_Handler:
    MsgBox "The selected invoice could not be opened." & vbCrLf & vbCrLf & _
           "Error Number: " & Err.Number & vbCrLf & _
           "Details: " & Err.Description, vbExclamation, "Open Invoice"
    Resume Exit_Handler

End Function
