Attribute VB_Name = "UXT_Test"
Option Compare Database
Option Explicit

'Scratch-copy test harness. Never imported into the delivered database.

Public Function UXT_LocationTest() As String

    Dim f As Form
    Dim report As String
    Dim waited As Single

    On Error GoTo Err_Handler

    CurrentDb.Execute "DELETE FROM T_Locations WHERE LocationName LIKE 'TEST-*'"

    '--- 1. New location, no warehouse, no bin, no shelf
    DoCmd.OpenForm "F_LocationSetup", acNormal, , , acFormAdd
    Set f = Forms("F_LocationSetup")

    f!LocationName = "TEST-AISLE-1"
    f!LocationCode = Null
    f!cboWarehouseID = Null
    f!BinNumber = Null
    f!ShelfNumber = Null
    f.Dirty = False

    report = report & "save=" & DCount("*", "T_Locations", "LocationName='TEST-AISLE-1'") & ";"
    report = report & "code=" & Nz(DLookup("LocationCode", "T_Locations", "LocationName='TEST-AISLE-1'"), "") & ";"
    report = report & "whNull=" & CStr(IsNull(DLookup("WarehouseID", "T_Locations", "LocationName='TEST-AISLE-1'"))) & ";"
    report = report & "binNull=" & CStr(IsNull(DLookup("BinNumber", "T_Locations", "LocationName='TEST-AISLE-1'"))) & ";"
    report = report & "shelfNull=" & CStr(IsNull(DLookup("ShelfNumber", "T_Locations", "LocationName='TEST-AISLE-1'"))) & ";"

    DoCmd.Close acForm, "F_LocationSetup", acSaveNo

    '--- 2. Reopen saved record
    DoCmd.OpenForm "F_LocationSetup", acNormal, , "LocationName='TEST-AISLE-1'"
    Set f = Forms("F_LocationSetup")
    report = report & "reopen=" & Nz(f!LocationName, "") & "/" & Nz(f!LocationCode, "") & ";"
    DoCmd.Close acForm, "F_LocationSetup", acSaveNo

    '--- 3. Quick add with existing name closes itself, no duplicate
    DoCmd.OpenForm "F_LocationSetup", acNormal, , , acFormAdd, acNormal, "QUICKADD|TEST-AISLE-1"
    waited = Timer
    Do While Timer < waited + 1.5
        DoEvents
    Loop
    report = report & "qa1Loaded=" & CStr(CurrentProject.AllForms("F_LocationSetup").IsLoaded) & ";"
    report = report & "qa1Rows=" & DCount("*", "T_Locations", "LocationName='TEST-AISLE-1'") & ";"
    If CurrentProject.AllForms("F_LocationSetup").IsLoaded Then
        DoCmd.Close acForm, "F_LocationSetup", acSaveNo
    End If

    '--- 4. Quick add with new name prefills and generates code
    DoCmd.OpenForm "F_LocationSetup", acNormal, , , acFormAdd, acNormal, "QUICKADD|TEST-AISLE-2"
    waited = Timer
    Do While Timer < waited + 1
        DoEvents
    Loop
    Set f = Forms("F_LocationSetup")
    report = report & "qa2Name=" & Nz(f!LocationName, "") & ";"
    report = report & "qa2Code=" & Nz(f!LocationCode, "") & ";"
    f.Dirty = False
    report = report & "qa2Rows=" & DCount("*", "T_Locations", "LocationName='TEST-AISLE-2'") & ";"
    DoCmd.Close acForm, "F_LocationSetup", acSaveNo

    CurrentDb.Execute "DELETE FROM T_Locations WHERE LocationName LIKE 'TEST-*'"

    UXT_LocationTest = report & "OK"
    Exit Function

Err_Handler:

    UXT_LocationTest = report & "ERR " & Err.Number & ":" & Err.Description

    On Error Resume Next
    If CurrentProject.AllForms("F_LocationSetup").IsLoaded Then
        DoCmd.Close acForm, "F_LocationSetup", acSaveNo
    End If

End Function

Public Function UXT_TemplateTest() As String

    Dim f As Form
    On Error GoTo Err_Handler

    DoCmd.OpenForm "F_AssemblyTemplate", acNormal, , , acFormAdd
    Set f = Forms("F_AssemblyTemplate")
    UXT_TemplateTest = f.UXT_PartNumberTest()
    DoCmd.Close acForm, "F_AssemblyTemplate", acSaveNo
    Exit Function

Err_Handler:
    UXT_TemplateTest = "ERR " & Err.Number & ":" & Err.Description
    On Error Resume Next
    If CurrentProject.AllForms("F_AssemblyTemplate").IsLoaded Then
        DoCmd.Close acForm, "F_AssemblyTemplate", acSaveNo
    End If

End Function

Public Function UXT_DrillTest() As String

    Dim f As Form
    Dim lst As Form
    Dim subF As Form
    Dim report As String
    Dim names As Variant
    Dim n As Variant

    On Error GoTo Err_Handler

    CurrentDb.Execute "DELETE FROM T_Customers WHERE CustomerName='TEST CUSTOMER'"
    CurrentDb.Execute "INSERT INTO T_Customers (CustomerName, CustomerType, Status) " & _
        "VALUES ('TEST CUSTOMER','Retail','In Review')"

    DoCmd.OpenForm "F_SystemSettings"
    Set f = Forms("F_SystemSettings")

    names = Array("lblCustomerStatus", "lblSupplierStatus", "lblCategoryStatus", _
        "lblUnitStatus", "lblItemStatus", "lblTemplateStatus", _
        "lblWarehouseStatus", "lblLocationStatus")

    For Each n In names
        report = report & n & "='" & Nz(f.Controls(n).OnClick, "") & "';"
    Next n

    DoCmd.Close acForm, "F_SystemSettings", acSaveNo

    UX_OpenAreaList "customer", True
    Set lst = Forms("F_CustomerList")
    Set subF = lst.Controls("subCustomerList").Form

    report = report & " filter='" & Nz(subF.Filter, "") & "';"
    report = report & " on=" & CStr(subF.FilterOn) & ";"
    report = report & " rows=" & subF.RecordsetClone.RecordCount & ";"

    DoCmd.Close acForm, "F_CustomerList", acSaveNo
    CurrentDb.Execute "DELETE FROM T_Customers WHERE CustomerName='TEST CUSTOMER'"

    UXT_DrillTest = report
    Exit Function

Err_Handler:
    UXT_DrillTest = report & "ERR " & Err.Number & ":" & Err.Description
    On Error Resume Next
    If CurrentProject.AllForms("F_CustomerList").IsLoaded Then DoCmd.Close acForm, "F_CustomerList", acSaveNo
    If CurrentProject.AllForms("F_SystemSettings").IsLoaded Then DoCmd.Close acForm, "F_SystemSettings", acSaveNo
    CurrentDb.Execute "DELETE FROM T_Customers WHERE CustomerName='TEST CUSTOMER'"

End Function

Public Function UXT_WarehouseTest() As String

    Dim idA As Long
    Dim idB As Long
    Dim report As String

    On Error GoTo Err_Handler

    CurrentDb.Execute "DELETE FROM T_Warehouses WHERE WarehouseName LIKE 'TEST-WH*'"
    CurrentDb.Execute "INSERT INTO T_Warehouses (WarehouseName, WarehouseCode) VALUES ('TEST-WH-A','TWA')"
    idA = DMax("WarehouseID", "T_Warehouses", "WarehouseName='TEST-WH-A'")
    CurrentDb.Execute "INSERT INTO T_Warehouses (WarehouseName, WarehouseCode) VALUES ('TEST-WH-B','TWB')"
    idB = DMax("WarehouseID", "T_Warehouses", "WarehouseName='TEST-WH-B'")

    DoCmd.OpenForm "F_Warehouses", acNormal, , "WarehouseID=" & idB
    report = "openB=" & Nz(Forms("F_Warehouses")!WarehouseName, "?") & ";"
    report = report & "filter='" & Nz(Forms("F_Warehouses").Filter, "") & "';"
    report = report & "on=" & CStr(Forms("F_Warehouses").FilterOn) & ";"
    DoCmd.Close acForm, "F_Warehouses", acSaveNo

    DoCmd.OpenForm "F_Warehouses", acNormal
    report = report & "unfiltered=" & Nz(Forms("F_Warehouses")!WarehouseName, "?") & ";"
    DoCmd.Close acForm, "F_Warehouses", acSaveNo

    CurrentDb.Execute "DELETE FROM T_Warehouses WHERE WarehouseName LIKE 'TEST-WH*'"

    UXT_WarehouseTest = report
    Exit Function

Err_Handler:
    UXT_WarehouseTest = report & "ERR " & Err.Number & ":" & Err.Description
    On Error Resume Next
    If CurrentProject.AllForms("F_Warehouses").IsLoaded Then DoCmd.Close acForm, "F_Warehouses", acSaveNo
    CurrentDb.Execute "DELETE FROM T_Warehouses WHERE WarehouseName LIKE 'TEST-WH*'"

End Function

Public Function UXT_UnitTest() As String

    Dim f As Form
    Dim lst As Form
    Dim newID As Long
    Dim report As String

    On Error GoTo Err_Handler

    CurrentDb.Execute "DELETE FROM T_Units WHERE UnitCode='TESTU'"

    DoCmd.OpenForm "F_UnitSetup", acNormal, , , acFormAdd
    Set f = Forms("F_UnitSetup")
    f!UnitCode = "TESTU"
    f!UnitName = "Test Unit"
    f.Dirty = False
    report = "created=" & DCount("*", "T_Units", "UnitCode='TESTU'") & ";"
    DoCmd.Close acForm, "F_UnitSetup", acSaveNo

    newID = DLookup("UnitID", "T_Units", "UnitCode='TESTU'")

    DoCmd.OpenForm "F_UnitList"
    Set lst = Forms("F_UnitList")
    report = report & "listRows=" & lst.RecordsetClone.RecordCount & ";"
    report = report & "inList=" & DCount("*", "Q_UnitList", "UnitID=" & newID) & ";"
    DoCmd.Close acForm, "F_UnitList", acSaveNo

    DoCmd.OpenForm "F_UnitSetup", acNormal, , "UnitID=" & newID
    Set f = Forms("F_UnitSetup")
    report = report & "editName=" & Nz(f!UnitName, "?") & ";"
    f!UnitName = "Test Unit Edited"
    f.Dirty = False
    report = report & "edited=" & DCount("*", "T_Units", "UnitID=" & newID & " AND UnitName='Test Unit Edited'") & ";"
    DoCmd.Close acForm, "F_UnitSetup", acSaveNo

    CurrentDb.Execute "DELETE FROM T_Units WHERE UnitCode='TESTU'"

    UXT_UnitTest = report
    Exit Function

Err_Handler:
    UXT_UnitTest = report & "ERR " & Err.Number & ":" & Err.Description
    On Error Resume Next
    If CurrentProject.AllForms("F_UnitSetup").IsLoaded Then DoCmd.Close acForm, "F_UnitSetup", acSaveNo
    If CurrentProject.AllForms("F_UnitList").IsLoaded Then DoCmd.Close acForm, "F_UnitList", acSaveNo
    CurrentDb.Execute "DELETE FROM T_Units WHERE UnitCode='TESTU'"

End Function

Public Function UXT_AutoSaveLocation() As Boolean

    On Error Resume Next

    If CurrentProject.AllForms("F_LocationSetup").IsLoaded Then
        Forms("F_LocationSetup").Dirty = False
        DoCmd.Close acForm, "F_LocationSetup", acSaveNo
        UXT_AutoSaveLocation = True
    End If

End Function

Public Function UXT_ItemLocTest() As String

    Dim report As String

    On Error GoTo Err_Handler

    CurrentDb.Execute "DELETE FROM T_Locations WHERE LocationName='TEST-ITEM-LOC'"

    DoCmd.OpenForm "F_ItemSetup", acNormal, , , acFormAdd
    On Error Resume Next
    TempVars.Remove "UXT_DirtyResult"
    On Error GoTo Err_Handler
    TempVars.Add "UXT_AutoSave", True

    report = Forms("F_ItemSetup").UXT_CreateLoc("TEST-ITEM-LOC") & ";"
    report = report & "dirtyResult=" & Nz(TempVars("UXT_DirtyResult"), "-") & ";"
    report = report & "rows=" & DCount("*", "T_Locations", "LocationName='TEST-ITEM-LOC'") & ";"
    report = report & "locID=" & Nz(DLookup("LocationID", "T_Locations", "LocationName='TEST-ITEM-LOC'"), -1)

    DoCmd.Close acForm, "F_ItemSetup", acSaveNo
    CurrentDb.Execute "DELETE FROM T_Locations WHERE LocationName='TEST-ITEM-LOC'"

    UXT_ItemLocTest = report
    Exit Function

Err_Handler:
    UXT_ItemLocTest = report & "ERR " & Err.Number & ":" & Err.Description
    On Error Resume Next
    If CurrentProject.AllForms("F_LocationSetup").IsLoaded Then DoCmd.Close acForm, "F_LocationSetup", acSaveNo
    If CurrentProject.AllForms("F_ItemSetup").IsLoaded Then DoCmd.Close acForm, "F_ItemSetup", acSaveNo
    CurrentDb.Execute "DELETE FROM T_Locations WHERE LocationName='TEST-ITEM-LOC'"

End Function
