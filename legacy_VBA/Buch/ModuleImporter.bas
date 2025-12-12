Attribute VB_Name = "ModuleImporter"
'==========================
'   Module Importer
'   Imports .bas and .cls modules from user-selected files into current workbook
'   Supports files with or without VB_Name attribute
'==========================
Option Explicit

' ========================================
' Main Import Procedure
' ========================================

Public Sub ImportModules()
    ' Imports selected .bas and .cls files into current workbook
    ' User can select multiple files in a dialog
    ' Supports files without VB_Name attribute - uses filename as module name
    
    On Error GoTo ErrorHandler
    
    ' Check if VBA project access is allowed
    If Not IsVBAAccessTrusted() Then
        MsgBox "Zugriff auf das VBA-Projekt ist nicht erlaubt." & vbCrLf & vbCrLf & _
               "Bitte aktivieren Sie in Excel:" & vbCrLf & _
               "Datei > Optionen > Trust Center > Einstellungen fuer das Trust Center > " & vbCrLf & _
               "Makroeinstellungen > 'Zugriff auf das VBA-Projektobjektmodell vertrauen'", _
               vbCritical, "VBA-Zugriff verweigert"
        Exit Sub
    End If
    
    ' Show file picker dialog
    Dim selectedFiles As Variant
    selectedFiles = SelectModuleFiles()
    
    If Not IsArray(selectedFiles) Then
        ' User cancelled
        Exit Sub
    End If
    
    Application.ScreenUpdating = False
    
    Dim importedCount As Long
    Dim updatedCount As Long
    Dim skippedCount As Long
    Dim errorCount As Long
    importedCount = 0
    updatedCount = 0
    skippedCount = 0
    errorCount = 0
    
    Dim vbProj As Object
    Set vbProj = ThisWorkbook.VBProject
    
    Dim i As Long
    For i = LBound(selectedFiles) To UBound(selectedFiles)
        Dim filePath As String
        filePath = selectedFiles(i)
        
        Dim result As Long
        result = ImportSingleModule(vbProj, filePath)
        
        Select Case result
            Case 1
                importedCount = importedCount + 1
            Case 2
                updatedCount = updatedCount + 1
            Case 0
                skippedCount = skippedCount + 1
            Case -1
                errorCount = errorCount + 1
        End Select
    Next i
    
    Application.ScreenUpdating = True
    
    ' Show results
    Dim msg As String
    msg = "Import abgeschlossen!" & vbCrLf & vbCrLf & _
          "Neu importiert: " & importedCount & " Modul(e)" & vbCrLf & _
          "Aktualisiert: " & updatedCount & " Modul(e)"
    
    If skippedCount > 0 Then
        msg = msg & vbCrLf & "Uebersprungen: " & skippedCount & " Datei(en)"
    End If
    
    If errorCount > 0 Then
        msg = msg & vbCrLf & "Fehler: " & errorCount & " Datei(en)"
    End If
    
    MsgBox msg, vbInformation, "Module importiert"
    
    Exit Sub
    
ErrorHandler:
    Application.ScreenUpdating = True
    MsgBox "Fehler beim Importieren der Module: " & Err.Description, vbCritical, "Importfehler"
End Sub

Public Sub ImportModulesFromFolder()
    ' Imports all .bas and .cls files from a user-selected folder
    ' Provides option to include subfolders
    
    On Error GoTo ErrorHandler
    
    ' Check if VBA project access is allowed
    If Not IsVBAAccessTrusted() Then
        MsgBox "Zugriff auf das VBA-Projekt ist nicht erlaubt." & vbCrLf & vbCrLf & _
               "Bitte aktivieren Sie in Excel:" & vbCrLf & _
               "Datei > Optionen > Trust Center > Einstellungen fuer das Trust Center > " & vbCrLf & _
               "Makroeinstellungen > 'Zugriff auf das VBA-Projektobjektmodell vertrauen'", _
               vbCritical, "VBA-Zugriff verweigert"
        Exit Sub
    End If
    
    ' Show folder picker
    Dim folderPath As String
    folderPath = SelectImportFolder()
    
    If folderPath = "" Then
        Exit Sub
    End If
    
    If Right(folderPath, 1) <> "\" Then
        folderPath = folderPath & "\"
    End If
    
    ' Collect all .bas and .cls files from folder
    Dim files As Collection
    Set files = New Collection
    
    Dim fso As Object
    Set fso = CreateObject("Scripting.FileSystemObject")
    
    Dim folder As Object
    Set folder = fso.GetFolder(folderPath)
    
    Dim file As Object
    For Each file In folder.files
        Dim ext As String
        ext = LCase(fso.GetExtensionName(file.name))
        
        If ext = "bas" Or ext = "cls" Then
            files.Add file.Path
        End If
    Next file
    
    If files.Count = 0 Then
        MsgBox "Keine .bas oder .cls Dateien im ausgewaehlten Ordner gefunden.", _
               vbExclamation, "Keine Module"
        Exit Sub
    End If
    
    ' Confirm import
    Dim response As VbMsgBoxResult
    response = MsgBox("Es wurden " & files.Count & " Moduldatei(en) gefunden." & vbCrLf & vbCrLf & _
                      "Moechten Sie alle importieren?", _
                      vbYesNo + vbQuestion, "Import bestaetigen")
    
    If response <> vbYes Then
        Exit Sub
    End If
    
    Application.ScreenUpdating = False
    
    Dim importedCount As Long
    Dim updatedCount As Long
    Dim skippedCount As Long
    Dim errorCount As Long
    importedCount = 0
    updatedCount = 0
    skippedCount = 0
    errorCount = 0
    
    Dim vbProj As Object
    Set vbProj = ThisWorkbook.VBProject
    
    Dim filePath As Variant
    For Each filePath In files
        Dim result As Long
        result = ImportSingleModule(vbProj, CStr(filePath))
        
        Select Case result
            Case 1
                importedCount = importedCount + 1
            Case 2
                updatedCount = updatedCount + 1
            Case 0
                skippedCount = skippedCount + 1
            Case -1
                errorCount = errorCount + 1
        End Select
    Next filePath
    
    Application.ScreenUpdating = True
    
    Dim msg As String
    msg = "Import abgeschlossen!" & vbCrLf & vbCrLf & _
          "Neu importiert: " & importedCount & " Modul(e)" & vbCrLf & _
          "Aktualisiert: " & updatedCount & " Modul(e)"
    
    If skippedCount > 0 Then
        msg = msg & vbCrLf & "Uebersprungen: " & skippedCount & " Datei(en)"
    End If
    
    If errorCount > 0 Then
        msg = msg & vbCrLf & "Fehler: " & errorCount & " Datei(en)"
    End If
    
    MsgBox msg, vbInformation, "Module importiert"
    
    Exit Sub
    
ErrorHandler:
    Application.ScreenUpdating = True
    MsgBox "Fehler beim Importieren: " & Err.Description, vbCritical, "Importfehler"
End Sub

' ========================================
' Core Import Function
' ========================================

Private Function ImportSingleModule(ByVal vbProj As Object, ByVal filePath As String) As Long
    ' Imports a single module file into the VBA project
    ' Returns: 1 = new import, 2 = updated, 0 = skipped, -1 = error
    
    On Error GoTo ImportError
    
    Dim fso As Object
    Set fso = CreateObject("Scripting.FileSystemObject")
    
    ' Verify file exists
    If Not fso.FileExists(filePath) Then
        Debug.Print "File not found: " & filePath
        ImportSingleModule = 0
        Exit Function
    End If
    
    ' Get file info
    Dim fileName As String
    Dim fileExt As String
    fileName = fso.GetBaseName(filePath)
    fileExt = LCase(fso.GetExtensionName(filePath))
    
    ' Validate extension
    If fileExt <> "bas" And fileExt <> "cls" Then
        Debug.Print "Invalid extension: " & filePath
        ImportSingleModule = 0
        Exit Function
    End If
    
    ' Read VB_Name attribute from file if present
    Dim moduleName As String
    moduleName = ExtractVBNameFromFile(filePath)
    
    ' If no VB_Name attribute found, use filename
    If moduleName = "" Then
        moduleName = fileName
    End If
    
    ' Sanitize module name (remove invalid characters)
    moduleName = SanitizeModuleName(moduleName)
    
    If moduleName = "" Then
        Debug.Print "Invalid module name for: " & filePath
        ImportSingleModule = 0
        Exit Function
    End If
    
    ' Check if module already exists
    Dim existingModule As Object
    Dim moduleExists As Boolean
    moduleExists = False
    
    On Error Resume Next
    Set existingModule = vbProj.VBComponents(moduleName)
    If Err.Number = 0 And Not existingModule Is Nothing Then
        moduleExists = True
    End If
    Err.Clear
    On Error GoTo ImportError
    
    ' Prepare file for import (ensure VB_Name attribute exists)
    Dim importPath As String
    importPath = PrepareFileForImport(filePath, moduleName, fileExt)
    
    If moduleExists Then
        ' Remove existing module first
        Dim expectedType As Long
        If fileExt = "bas" Then
            expectedType = 1  ' vbext_ct_StdModule
        Else
            expectedType = 2  ' vbext_ct_ClassModule
        End If
        
        ' Only remove if type matches or if replacing
        If existingModule.Type = expectedType Then
            vbProj.VBComponents.Remove existingModule
            
            ' Small delay to ensure module is removed
            DoEvents
        Else
            ' Type mismatch - skip with warning
            Debug.Print "Type mismatch for " & moduleName & ": existing=" & existingModule.Type & ", file=" & expectedType
            CleanupTempFile importPath, filePath
            ImportSingleModule = 0
            Exit Function
        End If
    End If
    
    ' Import the module
    vbProj.VBComponents.Import importPath
    
    ' Cleanup temp file if created
    CleanupTempFile importPath, filePath
    
    If moduleExists Then
        ImportSingleModule = 2  ' Updated
    Else
        ImportSingleModule = 1  ' New import
    End If
    
    Exit Function
    
ImportError:
    Debug.Print "Error importing " & filePath & ": " & Err.Description
    CleanupTempFile importPath, filePath
    ImportSingleModule = -1
End Function

' ========================================
' File Processing Functions
' ========================================

Private Function ExtractVBNameFromFile(ByVal filePath As String) As String
    ' Reads the Attribute VB_Name line from a .bas or .cls file
    ' Returns the module name or empty string if not found
    
    On Error GoTo ExtractError
    
    Dim fso As Object
    Set fso = CreateObject("Scripting.FileSystemObject")
    
    Dim ts As Object
    Set ts = fso.OpenTextFile(filePath, 1)  ' ForReading
    
    Dim lineText As String
    Dim foundName As String
    foundName = ""
    
    ' Read first 10 lines looking for VB_Name attribute
    Dim lineCount As Long
    lineCount = 0
    
    Do While Not ts.AtEndOfStream And lineCount < 10
        lineText = ts.ReadLine
        lineCount = lineCount + 1
        
        ' Check for Attribute VB_Name = "ModuleName"
        If InStr(1, lineText, "Attribute VB_Name", vbTextCompare) > 0 Then
            ' Extract name between quotes
            Dim startPos As Long
            Dim endPos As Long
            
            startPos = InStr(1, lineText, """")
            If startPos > 0 Then
                endPos = InStr(startPos + 1, lineText, """")
                If endPos > startPos Then
                    foundName = Mid(lineText, startPos + 1, endPos - startPos - 1)
                End If
            End If
            Exit Do
        End If
    Loop
    
    ts.Close
    ExtractVBNameFromFile = foundName
    Exit Function
    
ExtractError:
    ExtractVBNameFromFile = ""
End Function

Private Function PrepareFileForImport(ByVal originalPath As String, _
                                      ByVal moduleName As String, _
                                      ByVal fileExt As String) As String
    ' Prepares a file for import by ensuring it has proper VB_Name attribute
    ' If file already has VB_Name, returns original path
    ' If file lacks VB_Name, creates a temp file with attribute added
    
    On Error GoTo PrepareError
    
    Dim fso As Object
    Set fso = CreateObject("Scripting.FileSystemObject")
    
    ' Read file content
    Dim ts As Object
    Set ts = fso.OpenTextFile(originalPath, 1)  ' ForReading
    Dim content As String
    content = ts.ReadAll
    ts.Close
    
    ' Check if VB_Name already exists
    If InStr(1, content, "Attribute VB_Name", vbTextCompare) > 0 Then
        ' File already has VB_Name, use as-is
        PrepareFileForImport = originalPath
        Exit Function
    End If
    
    ' Create temp file with VB_Name attribute prepended
    Dim tempPath As String
    tempPath = fso.GetParentFolderName(originalPath) & "\" & "_temp_import_" & moduleName & "." & fileExt
    
    Dim tsOut As Object
    Set tsOut = fso.CreateTextFile(tempPath, True)  ' Overwrite if exists
    
    ' Write VB_Name attribute first
    tsOut.WriteLine "Attribute VB_Name = """ & moduleName & """"
    
    ' For class modules, add required attributes if missing
    If fileExt = "cls" Then
        If InStr(1, content, "VB_GlobalNameSpace", vbTextCompare) = 0 Then
            tsOut.WriteLine "Attribute VB_GlobalNameSpace = False"
        End If
        If InStr(1, content, "VB_Creatable", vbTextCompare) = 0 Then
            tsOut.WriteLine "Attribute VB_Creatable = False"
        End If
        If InStr(1, content, "VB_PredeclaredId", vbTextCompare) = 0 Then
            tsOut.WriteLine "Attribute VB_PredeclaredId = False"
        End If
        If InStr(1, content, "VB_Exposed", vbTextCompare) = 0 Then
            tsOut.WriteLine "Attribute VB_Exposed = False"
        End If
    End If
    
    ' Write original content
    tsOut.Write content
    tsOut.Close
    
    PrepareFileForImport = tempPath
    Exit Function
    
PrepareError:
    PrepareFileForImport = originalPath
End Function

Private Sub CleanupTempFile(ByVal tempPath As String, ByVal originalPath As String)
    ' Deletes temp file if it was created (different from original)
    
    If tempPath <> originalPath And tempPath <> "" Then
        On Error Resume Next
        Dim fso As Object
        Set fso = CreateObject("Scripting.FileSystemObject")
        
        If fso.FileExists(tempPath) Then
            fso.DeleteFile tempPath
        End If
        On Error GoTo 0
    End If
End Sub

Private Function SanitizeModuleName(ByVal name As String) As String
    ' Removes or replaces invalid characters from module name
    ' VBA module names must start with letter and contain only alphanumeric and underscore
    
    Dim result As String
    Dim i As Long
    Dim c As String
    
    result = ""
    
    For i = 1 To Len(name)
        c = Mid(name, i, 1)
        
        If i = 1 Then
            ' First character must be a letter
            If c Like "[A-Za-z]" Then
                result = result & c
            ElseIf c Like "[0-9_]" Then
                ' Prepend underscore if starts with number or underscore
                result = "_" & c
            End If
        Else
            ' Subsequent characters can be letters, numbers, or underscore
            If c Like "[A-Za-z0-9_]" Then
                result = result & c
            End If
        End If
    Next i
    
    ' Ensure name is not too long (max 31 characters for VBA)
    If Len(result) > 31 Then
        result = Left(result, 31)
    End If
    
    SanitizeModuleName = result
End Function

' ========================================
' Dialog Functions
' ========================================

Private Function SelectModuleFiles() As Variant
    ' Shows file picker dialog for selecting .bas and .cls files
    ' Returns array of selected file paths or False if cancelled
    
    Dim fileDialog As fileDialog
    Set fileDialog = Application.fileDialog(msoFileDialogFilePicker)
    
    With fileDialog
        .Title = "VBA-Module zum Importieren waehlen"
        .AllowMultiSelect = True
        .InitialFileName = ThisWorkbook.Path & "\"
        
        ' Clear existing filters and add new ones
        .Filters.Clear
        .Filters.Add "VBA Module", "*.bas;*.cls"
        .Filters.Add "Standard Module (*.bas)", "*.bas"
        .Filters.Add "Class Module (*.cls)", "*.cls"
        .Filters.Add "Alle Dateien", "*.*"
        .FilterIndex = 1
        
        If .Show = -1 Then
            ' User selected files
            Dim fileCount As Long
            fileCount = .SelectedItems.Count
            
            If fileCount > 0 Then
                Dim files() As String
                ReDim files(1 To fileCount)
                
                Dim i As Long
                For i = 1 To fileCount
                    files(i) = .SelectedItems(i)
                Next i
                
                SelectModuleFiles = files
            Else
                SelectModuleFiles = False
            End If
        Else
            SelectModuleFiles = False
        End If
    End With
End Function

Private Function SelectImportFolder() As String
    ' Shows folder picker dialog and returns selected path
    
    Dim folderDialog As fileDialog
    Set folderDialog = Application.fileDialog(msoFileDialogFolderPicker)
    
    With folderDialog
        .Title = "Ordner mit VBA-Modulen waehlen"
        .AllowMultiSelect = False
        .InitialFileName = ThisWorkbook.Path & "\"
        
        If .Show = -1 Then
            SelectImportFolder = .SelectedItems(1)
        Else
            SelectImportFolder = ""
        End If
    End With
End Function

Private Function IsVBAAccessTrusted() As Boolean
    ' Checks if VBA project access is allowed in Trust Center settings
    
    On Error Resume Next
    
    Dim testAccess As Long
    testAccess = ThisWorkbook.VBProject.VBComponents.Count
    
    If Err.Number = 0 Then
        IsVBAAccessTrusted = True
    Else
        IsVBAAccessTrusted = False
    End If
    
    Err.Clear
    On Error GoTo 0
End Function
