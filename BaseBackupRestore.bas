Attribute VB_Name = "BaseBackupRestore"
Option Explicit

' This function returns the full path of the main (current) database file.
Public Function GetMainDBPath() As String
    ' Modify this path according to your actual main DB location
    GetMainDBPath = ThisWorkbook.Worksheets("Kartei").Range("X1").value & "\Alarm\KindElternDaten_25_front.accdb"
End Function

' This procedure allows the user to pick a backup file from a dialog
' and then replaces the main DB with the chosen backup file (removing the suffix).
Public Sub RestoreDatabaseFromBackup()
    Dim fD As FileDialog
    Dim selectedFile As String
    Dim fso As Object
    Dim fileDate As Date
    Dim userResponse As VbMsgBoxResult
    Dim mainDBPath As String
    
    mainDBPath = GetMainDBPath
    
    ' Create a FileDialog object
    Set fD = Application.FileDialog(msoFileDialogFilePicker)
    With fD
        .title = "Select a Backup File for KindElternDaten_25_front"
        .AllowMultiSelect = False
        .InitialView = msoFileDialogViewDetails
        .Filters.Clear
        .Filters.Add "Access Database Files", "*.accdb"
        
        ' Optionally set initial folder location (modify as needed)
        .InitialFileName = "C:\KindEltern\2025\Saves\Dump\"
        
        ' Show the dialog and check if the user cancelled
        If .Show <> -1 Then
            MsgBox "Operation cancelled by the user.", vbInformation
            Exit Sub
        End If
        
        ' Get the selected file path
        selectedFile = .SelectedItems(1)
    End With
    
    ' Retrieve the file's last modified date/time
    Set fso = CreateObject("Scripting.FileSystemObject")
    fileDate = fso.GetFile(selectedFile).DateLastModified
    
    ' Ask user for confirmation
    userResponse = MsgBox( _
        "Do you want to restore the database to the state of " & fileDate & "?", _
        vbYesNo + vbQuestion, _
        "Confirm Restoration" _
    )
    
    If userResponse = vbNo Then
        MsgBox "Restoration cancelled.", vbInformation
        Exit Sub
    End If
    
    ' Remove the current main DB file if it exists
    On Error Resume Next
    Kill mainDBPath
    On Error GoTo 0
    
    ' Copy the chosen backup file to the main DB path (suffix removed in final naming)
    fso.CopyFile selectedFile, mainDBPath, True
    
    ' Release the FileSystemObject
    Set fso = Nothing
    
    MsgBox "The database has been successfully restored using the backup from " & fileDate _
            & ". Now this base version will be downloaded to the sheet Kartei", vbInformation
            
    ThisWorkbook.Worksheets("Tabelle8").Range("B1").value = GetIDWithDAO(mainDBPath)
            
    Call ThisWorkbook.Workbook_Open
End Sub


Public Function GetIDWithDAO(ByVal dbPath As String) As Long
    Dim engine As DAO.DBEngine
    Set engine = New DAO.DBEngine
    
    Dim wsDao As DAO.Workspace
    Set wsDao = engine.Workspaces(0)

    Dim db As DAO.Database
    Set db = wsDao.OpenDatabase(dbPath)
    
    Dim rs As DAO.Recordset
    Dim nextID As Long
    Set rs = db.OpenRecordset("tblKartei", dbOpenDynaset, dbSeeChanges)
    
    wsDao.BeginTrans
    rs.MoveLast
    rs.MoveFirst

    rs.AddNew
    rs("Value1") = "TemporaryData"
    rs.Update
    
    rs.MoveLast
    Debug.Print "Record count is: "; rs.RecordCount
    
    Dim FluxID As Long
    FluxID = rs("ID")
    
    wsDao.Rollback
    
    rs.Close

    Set rs = Nothing
    Set db = Nothing
    
    GetIDWithDAO = FluxID
    
End Function

