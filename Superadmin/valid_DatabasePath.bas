Attribute VB_Name = "valid_DatabasePath"
'==========================
'   Module: valid_DatabasePath
'   Purpose: Centralized database path management for Superadmin
'   - Validates database file existence
'   - Prompts user to select folder if not found
'   - Stores path in Kartei!X1
'==========================

Option Explicit

Private Const DB_FILENAME As String = "KindElternDaten_25_front.accdb"
Private Const DB_SUBFOLDER As String = "Alarm"

' Get validated database path, prompting user if file not found
Public Function GetValidatedDatabasePath() As String
    Dim dbPath As String
    dbPath = BuildDatabasePath()
    
    ' Check if database file exists
    If Not FileExists(dbPath) Then
        ' Database not found - prompt user to select folder
        Dim userSelectedPath As String
        userSelectedPath = PromptForDatabaseFolder(dbPath)
        
        If userSelectedPath = "" Then
            ' User cancelled - return empty string to signal abort
            GetValidatedDatabasePath = ""
            Exit Function
        End If
        
        dbPath = userSelectedPath
    End If
    
    GetValidatedDatabasePath = dbPath
End Function

' Build database path from Kartei!X1 or ThisWorkbook.Path
Private Function BuildDatabasePath() As String
    On Error Resume Next
    
    Dim basePath As String
    basePath = ""
    
    ' Try to read from Kartei!X1
    Dim wsKartei As Worksheet
    Set wsKartei = ThisWorkbook.Worksheets("Kartei")
    
    If Not wsKartei Is Nothing Then
        basePath = Trim(CStr(wsKartei.Range("X1").Value))
    End If
    
    ' Fallback to workbook path if X1 is empty
    If basePath = "" Then
        basePath = ThisWorkbook.Path
    End If
    
    BuildDatabasePath = basePath & "\" & DB_SUBFOLDER & "\" & DB_FILENAME
    On Error GoTo 0
End Function

' Prompt user to select the folder containing the database
Private Function PromptForDatabaseFolder(ByVal currentPath As String) As String
    PromptForDatabaseFolder = ""
    
    ' Show message explaining the problem
    Dim msgResult As VbMsgBoxResult
    msgResult = MsgBox("Datenbankdatei nicht gefunden:" & vbCrLf & vbCrLf & _
                       currentPath & vbCrLf & vbCrLf & _
                       "Moechten Sie den Ordner mit der Datenbank auswaehlen?" & vbCrLf & _
                       "(Der Ordner sollte '" & DB_SUBFOLDER & "\" & DB_FILENAME & "' enthalten)", _
                       vbYesNo + vbQuestion, "Datenbank nicht gefunden")
    
    If msgResult = vbNo Then
        Exit Function
    End If
    
    ' Show folder picker dialog
    Dim folderPath As String
    folderPath = BrowseForFolder("Waehlen Sie den Ordner mit der Datenbank (uebergeordneter Ordner von " & DB_SUBFOLDER & ")")
    
    If folderPath = "" Then
        MsgBox "Kein Ordner ausgewaehlt. Vorgang abgebrochen.", vbInformation, "Abgebrochen"
        Exit Function
    End If
    
    ' Build and validate the new path
    Dim newDbPath As String
    newDbPath = folderPath & "\" & DB_SUBFOLDER & "\" & DB_FILENAME
    
    If Not FileExists(newDbPath) Then
        MsgBox "Datenbankdatei immer noch nicht gefunden unter:" & vbCrLf & vbCrLf & _
               newDbPath & vbCrLf & vbCrLf & _
               "Bitte stellen Sie sicher, dass der ausgewaehlte Ordner '" & DB_SUBFOLDER & "\" & DB_FILENAME & "' enthaelt.", _
               vbExclamation, "Datenbank nicht gefunden"
        Exit Function
    End If
    
    ' Save the selected path to Kartei!X1
    SaveDatabaseBasePath folderPath
    
    MsgBox "Datenbankpfad erfolgreich aktualisiert!" & vbCrLf & vbCrLf & _
           "Neuer Pfad: " & newDbPath, vbInformation, "Pfad aktualisiert"
    
    PromptForDatabaseFolder = newDbPath
End Function

' Save base path to Kartei!X1
Private Sub SaveDatabaseBasePath(ByVal basePath As String)
    On Error Resume Next
    
    Dim wsKartei As Worksheet
    Set wsKartei = ThisWorkbook.Worksheets("Kartei")
    
    If Not wsKartei Is Nothing Then
        wsKartei.Range("X1").Value = basePath
    End If
    
    On Error GoTo 0
End Sub

' Check if file exists
Private Function FileExists(ByVal filePath As String) As Boolean
    On Error Resume Next
    FileExists = (Dir(filePath) <> "")
    On Error GoTo 0
End Function

' Browse for folder using FileDialog
Private Function BrowseForFolder(ByVal dialogTitle As String) As String
    BrowseForFolder = ""
    
    Dim fd As FileDialog
    Set fd = Application.FileDialog(msoFileDialogFolderPicker)
    
    With fd
        .Title = dialogTitle
        .InitialFileName = ThisWorkbook.Path & "\"
        .AllowMultiSelect = False
        
        If .Show = -1 Then
            BrowseForFolder = .SelectedItems(1)
        End If
    End With
    
    Set fd = Nothing
End Function

' Public procedure to manually select database folder
Public Sub SelectDatabaseFolder()
    Dim currentPath As String
    currentPath = BuildDatabasePath()
    
    Dim newPath As String
    newPath = PromptForDatabaseFolder(currentPath)
    
    If newPath <> "" Then
        MsgBox "Datenbankpfad ist jetzt gesetzt auf:" & vbCrLf & vbCrLf & newPath, vbInformation, "Datenbankpfad"
    End If
End Sub
