Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Compare Database
Option Explicit


'=============================================================
' SUPPLIER
'=============================================================

Private Sub cboSupplierID_NotInList( _
    NewData As String, _
    Response As Integer _
)

    Response = acDataErrContinue

    If HandleComboNotInList( _
        Me!cboSupplierID, _
        NewData, _
        "F_SupplierSetup", _
        "T_Suppliers", _
        "SupplierID", _
        "SupplierCode", _
        "CompanyName" _
    ) Then

        Response = acDataErrAdded

    Else

        Me!cboSupplierID.Undo

    End If

End Sub


Private Sub cboSupplierID_DblClick(Cancel As Integer)

    Cancel = True

    EditComboEntry _
        Me!cboSupplierID, _
        "F_SupplierSetup", _
        "SupplierID"

End Sub


'=============================================================
' CURRENCY
'=============================================================

Private Sub cboCurrencyID_NotInList( _
    NewData As String, _
    Response As Integer _
)

    Response = acDataErrContinue

    If HandleComboNotInList( _
        Me!cboCurrencyID, _
        NewData, _
        "F_CurrencySetup", _
        "T_Currencies", _
        "CurrencyID", _
        "CurrencyCode", _
        "CurrencyName" _
    ) Then

        Response = acDataErrAdded

    Else

        Me!cboCurrencyID.Undo

    End If

End Sub


Private Sub cboCurrencyID_DblClick(Cancel As Integer)

    Cancel = True

    EditComboEntry _
        Me!cboCurrencyID, _
        "F_CurrencySetup", _
        "CurrencyID"

End Sub
