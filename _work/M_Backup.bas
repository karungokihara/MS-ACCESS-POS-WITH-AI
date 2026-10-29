Attribute VB_Name = "M_Backup"
Option Compare Database
Option Explicit

'=============================================================
' MODULE: M_Backup
'
' Full ERP master-data backup and restore.
'
' UX_BackupAllData
'   Creates one .accdb file inside .\Backups containing every
'   user table and reports the file path.
'
' UX_RestoreAllData
'   Asks for one backup file, saves a safety copy of the current
'   data, then replaces the current data with the backup.
'=============================================================

Private Const BACKUP_FOLDER_NAME As String = "Backups"


'=============================================================
' BACKUP
'=============================================================

Public Function UX_BackupAllData( _
    Optional ByVal silentBackup As Boolean = False _
) As String

    On Error GoTo Err_Handler

    Dim folderPath As String
    Dim destPath As String
    Dim backupDb As DAO.Database
    Dim sourceTable As DAO.TableDef
    Dim exportedCount As Long

    folderPath = BackupFolderPath()

    destPath = folderPath & "\LLC_DataBackup_" & _
        Format$(Now(), "yyyymmdd_hhnnss") & ".accdb"

    Set backupDb = DBEngine.CreateDatabase(destPath, dbLangGeneral)
    backupDb.Close
    Set backupDb = Nothing

    For Each sourceTable In CurrentDb.TableDefs

        If IsUserTable(sourceTable.Name) Then

            DoCmd.TransferDatabase _
                acExport, _
                "Microsoft Access", _
                destPath, _
                acTable, _
                sourceTable.Name, _
                sourceTable.Name

            exportedCount = exportedCount + 1

        End If

    Next sourceTable

    UX_BackupAllData = destPath

    If Not silentBackup Then
        MsgBox _
            "Backup completed." & vbCrLf & vbCrLf & _
            "Tables exported: " & exportedCount & vbCrLf & _
            "File:" & vbCrLf & destPath, _
            vbInformation, _
            "System Backup"
    End If

Exit_Handler:
    Set sourceTable = Nothing
    Set backupDb = Nothing
    Exit Function

Err_Handler:

    If Not silentBackup Then
        MsgBox _
            "The system backup could not be completed." & vbCrLf & vbCrLf & _
            "Error Number: " & Err.Number & vbCrLf & _
            "Details: " & Err.Description, _
            vbExclamation, _
            "System Backup"
    End If

    Resume Exit_Handler

End Function


Public Sub UX_BackupButton()

    UX_BackupAllData

End Sub


'=============================================================
' RESTORE
'=============================================================

Public Sub UX_RestoreAllData()

    On Error GoTo Err_Handler

    Dim sourcePath As String
    Dim confirmValue As VbMsgBoxResult

    sourcePath = BrowseForBackupFile()

    If Len(sourcePath) = 0 Then Exit Sub

    confirmValue = MsgBox( _
        "Restore will replace the current ERP data with the " & _
        "selected backup." & vbCrLf & vbCrLf & _
        "A safety backup of the current data is saved first." & _
        vbCrLf & vbCrLf & _
        "Continue?", _
        vbQuestion + vbYesNo, _
        "Restore System Data")

    If confirmValue <> vbYes Then Exit Sub

    UX_RestoreFromPath sourcePath, False
    Exit Sub

Err_Handler:

    MsgBox _
        "The restore could not be completed." & vbCrLf & vbCrLf & _
        "Error Number: " & Err.Number & vbCrLf & _
        "Details: " & Err.Description, _
        vbExclamation, _
        "Restore System Data"

End Sub


Public Function UX_RestoreFromPath( _
    ByVal sourcePath As String, _
    Optional ByVal silentRestore As Boolean = False _
) As Long

    Dim sourceDb As DAO.Database
    Dim sourceTable As DAO.TableDef
    Dim tableName As String
    Dim importedCount As Long
    Dim remainingDeletes As Long
    Dim passNumber As Long

    On Error GoTo Err_Handler

    If Len(Trim$(sourcePath)) = 0 Then Exit Function

    If Dir(sourcePath, vbNormal) = vbNullString Then Exit Function

    UX_BackupAllData True

    DoCmd.SetWarnings False

    Set sourceDb = DBEngine.OpenDatabase(sourcePath)

    ClearUserTables remainingDeletes, passNumber

    For Each sourceTable In sourceDb.TableDefs

        tableName = sourceTable.Name

        If IsUserTable(tableName) Then

            If TableExistsLocal(tableName) Then

                ImportRestoredTable sourcePath, tableName
                importedCount = importedCount + 1

            End If

        End If

    Next sourceTable

    sourceDb.Close
    Set sourceDb = Nothing

    DoCmd.SetWarnings True

    Application.RefreshDatabaseWindow

    UX_RestoreFromPath = importedCount

    If Not silentRestore Then
        MsgBox _
            "Restore completed." & vbCrLf & vbCrLf & _
            "Tables restored: " & importedCount & vbCrLf & vbCrLf & _
            "Close and reopen the database before using it.", _
            vbInformation, _
            "Restore System Data"
    End If

Exit_Handler:

    On Error Resume Next

    DoCmd.SetWarnings True

    If Not sourceDb Is Nothing Then sourceDb.Close

    Set sourceTable = Nothing
    Set sourceDb = Nothing

    Exit Function

Err_Handler:

    DoCmd.SetWarnings True

    If Not silentRestore Then
        MsgBox _
            "The restore could not be completed." & vbCrLf & vbCrLf & _
            "Error Number: " & Err.Number & vbCrLf & _
            "Details: " & Err.Description, _
            vbExclamation, _
            "Restore System Data"
    End If

    Resume Exit_Handler

End Function


Public Sub UX_RestoreButton()

    UX_RestoreAllData

End Sub


'=============================================================
' DELETE EVERY USER ROW - RETRIED TO SATISFY FOREIGN KEYS
'=============================================================

Private Sub ClearUserTables( _
    ByRef remainingDeletes As Long, _
    ByRef passNumber As Long _
)

    Dim sourceTable As DAO.TableDef

    remainingDeletes = 1
    passNumber = 0

    Do While remainingDeletes > 0 And passNumber < 30

        remainingDeletes = 0
        passNumber = passNumber + 1

        For Each sourceTable In CurrentDb.TableDefs

            If IsUserTable(sourceTable.Name) Then

                On Error Resume Next

                CurrentDb.Execute _
                    "DELETE FROM [" & sourceTable.Name & "]", _
                    dbFailOnError

                If Err.Number <> 0 Then
                    remainingDeletes = remainingDeletes + 1
                    Err.Clear
                End If

                On Error GoTo 0

            End If

        Next sourceTable

    Loop

End Sub


'=============================================================
' ONE TABLE: TEMP IMPORT, THEN INSERT SHARED FIELDS
'=============================================================

Private Sub ImportRestoredTable( _
    ByVal sourcePath As String, _
    ByVal tableName As String _
)

    On Error GoTo Err_Handler

    Dim tempName As String
    Dim fieldList As String
    Dim db As DAO.Database
    Dim destTable As DAO.TableDef
    Dim srcTable As DAO.TableDef
    Dim destField As DAO.Field
    Dim srcDb As DAO.Database

    tempName = tableName & "_RESTORE_TMP"

    If TableExistsLocal(tempName) Then
        CurrentDb.TableDefs.Delete tempName
        CurrentDb.TableDefs.Refresh
    End If

    DoCmd.TransferDatabase _
        acImport, _
        "Microsoft Access", _
        sourcePath, _
        acTable, _
        tableName, _
        tempName

    Set db = CurrentDb
    Set destTable = db.TableDefs(tableName)
    Set srcDb = DBEngine.OpenDatabase(sourcePath)
    Set srcTable = srcDb.TableDefs(tableName)

    fieldList = vbNullString

    For Each destField In destTable.Fields

        If FieldExistsInTableDef(srcTable, destField.Name) Then

            If Len(fieldList) > 0 Then fieldList = fieldList & ", "
            fieldList = fieldList & "[" & destField.Name & "]"

        End If

    Next destField

    srcDb.Close
    Set srcDb = Nothing

    If Len(fieldList) > 0 Then

        db.Execute _
            "INSERT INTO [" & tableName & "] (" & fieldList & ") " & _
            "SELECT " & fieldList & " FROM [" & tempName & "]", _
            dbFailOnError

    End If

    db.TableDefs.Delete tempName
    db.TableDefs.Refresh

Exit_Handler:

    On Error Resume Next

    If Not srcDb Is Nothing Then srcDb.Close

    Set destField = Nothing
    Set destTable = Nothing
    Set srcTable = Nothing
    Set srcDb = Nothing
    Set db = Nothing

    Exit Sub

Err_Handler:

    MsgBox _
        "Table '" & tableName & "' could not be restored." & _
        vbCrLf & vbCrLf & _
        "Error Number: " & Err.Number & vbCrLf & _
        "Details: " & Err.Description, _
        vbExclamation, _
        "Restore Table"

    Resume Exit_Handler

End Sub


'=============================================================
' HELPERS
'=============================================================

Private Function BackupFolderPath() As String

    Dim folderPath As String

    folderPath = CurrentProject.Path & "\" & BACKUP_FOLDER_NAME

    If Dir(folderPath, vbDirectory) = vbNullString Then
        MkDir folderPath
    End If

    BackupFolderPath = folderPath

End Function


Private Function BrowseForBackupFile() As String

    Dim dialogObject As Object
    Dim selectedPath As String

    Set dialogObject = Application.FileDialog(1)

    With dialogObject

        .Title = "Select LLC data backup"
        .AllowMultiSelect = False
        .Filters.Clear
        .Filters.Add "Access Backup", "*.accdb;*.mdb"

        If .Show = -1 Then
            selectedPath = CStr(.SelectedItems(1))
        End If

    End With

    BrowseForBackupFile = selectedPath
    Set dialogObject = Nothing

End Function


Private Function IsUserTable(ByVal tableName As String) As Boolean

    If Left$(tableName, 4) = "MSys" Then Exit Function
    If Left$(tableName, 4) = "USys" Then Exit Function
    If Left$(tableName, 1) = "~" Then Exit Function
    If InStr(1, tableName, "_RESTORE_TMP", vbTextCompare) > 0 Then Exit Function

    IsUserTable = True

End Function


Private Function TableExistsLocal(ByVal tableName As String) As Boolean

    On Error GoTo Not_Found

    TableExistsLocal = (CurrentDb.TableDefs(tableName).Name = tableName)
    Exit Function

Not_Found:
    TableExistsLocal = False

End Function


Private Function FieldExistsInTableDef( _
    ByVal tableDefObject As DAO.TableDef, _
    ByVal fieldName As String _
) As Boolean

    Dim oneField As DAO.Field

    For Each oneField In tableDefObject.Fields
        If StrComp(oneField.Name, fieldName, vbTextCompare) = 0 Then
            FieldExistsInTableDef = True
            Exit Function
        End If
    Next oneField

End Function
