Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Compare Database
Option Explicit

'=============================================================
' FORM: F_NotificationCenter
'
' PURPOSE:
'   Popup notification summary.
'
' IMPORTANT:
'   The form itself should have:
'       On Open: blank
'       On Load: blank
'       On Activate: blank
'
'   Buttons call M_Notifications.
'=============================================================


Private Sub cmdRefresh_Click()

    Notif_RefreshNotificationForm Me

End Sub


Private Sub cmdViewRestockRequired_Click()

    Notif_OpenRestockRequired

End Sub


Private Sub cmdViewLowStock_Click()

    Notif_OpenLowStock

End Sub


Private Sub cmdViewToolOverdue_Click()

    Notif_OpenOverdueTools

End Sub


Private Sub cmdViewToolDueSoon_Click()

    Notif_OpenToolsDueSoon

End Sub


Private Sub cmdViewInventoryList_Click()

    Notif_OpenInventoryList

End Sub


Private Sub cmdClose_Click()

    DoCmd.close acForm, Me.Name

End Sub
