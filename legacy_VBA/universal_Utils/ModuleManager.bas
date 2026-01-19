Attribute VB_Name = "ModuleManager"
'==========================
'   Module: ModuleManager
'   Purpose: Import/Export VBA modules (.bas, .cls) with user-friendly dialogs
'
'   Features:
'   - Import multiple .bas/.cls files at once
'   - Warning if modules with same names already exist
'   - Handles files without proper Attribute headers
'   - Export all modules to selected folder
'   - Export selected modules from a list
'
'   Entry Points:
'   - ImportModules()        - Select and import multiple .bas/.cls files
'   - ExportAllModules()     - Export all modules to selected folder
'   - ExportSelectedModules() - Show list and export selected modules
'
'   Requirements:
'   - Trust access to VBA project object model must be enabled
'     (File > Options > Trust Center > Trust Center Settings > Macro Settings)
'==========================

Option Explicit

' Module type constants
Private Const vbext_ct_StdModule As Long = 1
Private Const vbext_ct_ClassModule As Long = 2
Private Const vbext_ct_MSForm As Long = 3
Private Const vbext_ct_Document As Long = 100

' ============================================================
' IMPORT MODULES
' ============================================================

' Main entry point: Import multiple .bas/.cls files
Public Sub ImportModules()
    On Error GoTo ErrorHandler
    
    ' Check VBA project access
    If Not HasVBProjectAccess() Then
        MsgBox "Zugriff auf VBA-Projekt nicht moeglich." & vbCrLf & vbCrLf & _
               "Bitte aktivieren Sie:" & vbCrLf & _
               "Datei > Optionen > Trust Center > Einstellungen fuer das Trust Center > Makroeinstellungen" & vbCrLf & _
               "> 'Zugriff auf das VBA-Projektobjektmodell vertrauen'", _
               vbCritical, "VBA-Projektzugriff erforderlich"
        Exit Sub
    End If
    
    ' Open file dialog for multiple selection
    Dim fd As FileDialog
    Set fd = Application.FileDialog(msoFileDialogFilePicker)
    
    With fd
        .Title = "VBA-Module zum Importieren auswaehlen"
        .AllowMultiSelect = True
        .Filters.Clear
        .Filters.Add "VBA-Module", "*.bas;*.cls"
        .Filters.Add "Standard-Module (*.bas)", "*.bas"
        .Filters.Add "Klassen-Module (*.cls)", "*.cls"
        .Filters.Add "Alle Dateien", "*.*"
        .FilterIndex = 1
        
        If .Show <> -1 Then
            ' User cancelled
            Exit Sub
        End If
        
        If .SelectedItems.Count = 0 Then
            Exit Sub
        End If
        
        ' Collect files to import
        Dim filePaths() As String
        ReDim filePaths(1 To .SelectedItems.Count)
        
        Dim i As Long
        For i = 1 To .SelectedItems.Count
            filePaths(i) = .SelectedItems(i)
        Next i
    End With
    
    ' Check for existing modules and confirm
    Dim existingModules As String
    existingModules = CheckExistingModules(filePaths)
    
    If existingModules <> "" Then
        Dim result As VbMsgBoxResult
        result = MsgBox("Folgende Module existieren bereits und werden ueberschrieben:" & vbCrLf & vbCrLf & _
                       existingModules & vbCrLf & _
                       "Moechten Sie fortfahren?", _
                       vbYesNo + vbExclamation, "Module existieren bereits")
        
        If result = vbNo Then
            Exit Sub
        End If
    End If
    
    ' Import files
    Dim importedCount As Long
    Dim failedCount As Long
    Dim failedFiles As String
    
    Application.ScreenUpdating = False
    
    For i = LBound(filePaths) To UBound(filePaths)
        If ImportSingleModule(filePaths(i)) Then
            importedCount = importedCount + 1
        Else
            failedCount = failedCount + 1
            failedFiles = failedFiles & vbCrLf & "  - " & GetFileName(filePaths(i))
        End If
    Next i
    
    Application.ScreenUpdating = True
    
    ' Show result
    Dim msg As String
    msg = "Import abgeschlossen:" & vbCrLf & vbCrLf & _
          "Erfolgreich importiert: " & importedCount & " Modul(e)"
    
    If failedCount > 0 Then
        msg = msg & vbCrLf & "Fehlgeschlagen: " & failedCount & " Modul(e)" & failedFiles
    End If
    
    MsgBox msg, IIf(failedCount > 0, vbExclamation, vbInformation), "Import-Ergebnis"
    Exit Sub
    
ErrorHandler:
    Application.ScreenUpdating = True
    MsgBox "Fehler beim Importieren: " & Err.Description, vbCritical, "Import-Fehler"
End Sub

' Check which modules already exist
Private Function CheckExistingModules(ByRef filePaths() As String) As String
    Dim result As String
    result = ""
    
    Dim i As Long
    For i = LBound(filePaths) To UBound(filePaths)
        Dim moduleName As String
        moduleName = GetModuleNameFromFile(filePaths(i))
        
        If ModuleExists(moduleName) Then
            result = result & vbCrLf & "  - " & moduleName
        End If
    Next i
    
    CheckExistingModules = result
End Function

' Get module name from file (reads Attribute VB_Name or uses filename)
Private Function GetModuleNameFromFile(ByVal filePath As String) As String
    On Error GoTo UseFileName
    
    Dim fso As Object
    Set fso = CreateObject("Scripting.FileSystemObject")
    
    Dim ts As Object
    Set ts = fso.OpenTextFile(filePath, 1, False) ' 1 = ForReading
    
    Dim line As String
    Dim moduleName As String
    moduleName = ""
    
    ' Read first 20 lines looking for Attribute VB_Name
    Dim lineCount As Long
    lineCount = 0
    
    Do While Not ts.AtEndOfStream And lineCount < 20
        line = ts.ReadLine
        lineCount = lineCount + 1
        
        If InStr(1, line, "Attribute VB_Name", vbTextCompare) > 0 Then
            ' Extract name from: Attribute VB_Name = "ModuleName"
            Dim startPos As Long, endPos As Long
            startPos = InStr(line, """")
            If startPos > 0 Then
                endPos = InStr(startPos + 1, line, """")
                If endPos > startPos Then
                    moduleName = Mid(line, startPos + 1, endPos - startPos - 1)
                End If
            End If
            Exit Do
        End If
    Loop
    
    ts.Close
    
    If moduleName = "" Then
        GoTo UseFileName
    End If
    
    GetModuleNameFromFile = moduleName
    Exit Function
    
UseFileName:
    ' Use filename without extension as module name
    GetModuleNameFromFile = GetFileNameWithoutExtension(filePath)
End Function

' Check if module with given name exists
Private Function ModuleExists(ByVal moduleName As String) As Boolean
    On Error Resume Next
    
    Dim vbComp As Object
    Set vbComp = ThisWorkbook.VBProject.VBComponents(moduleName)
    
    ModuleExists = Not (vbComp Is Nothing)
    
    On Error GoTo 0
End Function

' Import a single module file
Private Function ImportSingleModule(ByVal filePath As String) As Boolean
    On Error GoTo ErrorHandler
    
    Dim moduleName As String
    moduleName = GetModuleNameFromFile(filePath)
    
    ' Remove existing module if present
    If ModuleExists(moduleName) Then
        ThisWorkbook.VBProject.VBComponents.Remove _
            ThisWorkbook.VBProject.VBComponents(moduleName)
    End If
    
    ' Check if file has proper Attribute header
    If Not HasAttributeHeader(filePath) Then
        ' File doesn't have Attribute header - need to fix it
        Dim tempPath As String
        tempPath = CreateTempFileWithHeader(filePath, moduleName)
        
        If tempPath <> "" Then
            ThisWorkbook.VBProject.VBComponents.Import tempPath
            ' Delete temp file
            On Error Resume Next
            Kill tempPath
            On Error GoTo ErrorHandler
        Else
            ImportSingleModule = False
            Exit Function
        End If
    Else
        ' Normal import
        ThisWorkbook.VBProject.VBComponents.Import filePath
    End If
    
    ImportSingleModule = True
    Exit Function
    
ErrorHandler:
    Debug.Print "Error importing " & filePath & ": " & Err.Description
    ImportSingleModule = False
End Function

' Check if file has Attribute VB_Name header
Private Function HasAttributeHeader(ByVal filePath As String) As Boolean
    On Error GoTo NoHeader
    
    Dim fso As Object
    Set fso = CreateObject("Scripting.FileSystemObject")
    
    Dim ts As Object
    Set ts = fso.OpenTextFile(filePath, 1, False)
    
    Dim line As String
    Dim lineCount As Long
    lineCount = 0
    
    ' Check first 10 lines for Attribute VB_Name
    Do While Not ts.AtEndOfStream And lineCount < 10
        line = ts.ReadLine
        lineCount = lineCount + 1
        
        If InStr(1, line, "Attribute VB_Name", vbTextCompare) > 0 Then
            ts.Close
            HasAttributeHeader = True
            Exit Function
        End If
    Loop
    
    ts.Close
    HasAttributeHeader = False
    Exit Function
    
NoHeader:
    HasAttributeHeader = False
End Function

' Create temp file with proper Attribute header
Private Function CreateTempFileWithHeader(ByVal sourcePath As String, ByVal moduleName As String) As String
    On Error GoTo ErrorHandler
    
    Dim fso As Object
    Set fso = CreateObject("Scripting.FileSystemObject")
    
    ' Determine module type from extension
    Dim ext As String
    ext = LCase(fso.GetExtensionName(sourcePath))
    
    Dim isClass As Boolean
    isClass = (ext = "cls")
    
    ' Read source file
    Dim ts As Object
    Set ts = fso.OpenTextFile(sourcePath, 1, False)
    Dim content As String
    content = ts.ReadAll
    ts.Close
    
    ' Create temp file path
    Dim tempPath As String
    tempPath = Environ("TEMP") & "\" & moduleName & "." & ext
    
    ' Build header
    Dim header As String
    If isClass Then
        header = "VERSION 1.0 CLASS" & vbCrLf & _
                 "BEGIN" & vbCrLf & _
                 "  MultiUse = -1  'True" & vbCrLf & _
                 "END" & vbCrLf & _
                 "Attribute VB_Name = """ & moduleName & """" & vbCrLf & _
                 "Attribute VB_GlobalNameSpace = False" & vbCrLf & _
                 "Attribute VB_Creatable = False" & vbCrLf & _
                 "Attribute VB_PredeclaredId = False" & vbCrLf & _
                 "Attribute VB_Exposed = False" & vbCrLf
    Else
        header = "Attribute VB_Name = """ & moduleName & """" & vbCrLf
    End If
    
    ' Write temp file with header
    Set ts = fso.CreateTextFile(tempPath, True)
    ts.Write header & content
    ts.Close
    
    CreateTempFileWithHeader = tempPath
    Exit Function
    
ErrorHandler:
    Debug.Print "Error creating temp file: " & Err.Description
    CreateTempFileWithHeader = ""
End Function

' ============================================================
' EXPORT ALL MODULES
' ============================================================

' Export all .bas and .cls modules to selected folder
Public Sub ExportAllModules()
    On Error GoTo ErrorHandler
    
    ' Check VBA project access
    If Not HasVBProjectAccess() Then
        MsgBox "Zugriff auf VBA-Projekt nicht moeglich." & vbCrLf & vbCrLf & _
               "Bitte aktivieren Sie 'Zugriff auf das VBA-Projektobjektmodell vertrauen'", _
               vbCritical, "VBA-Projektzugriff erforderlich"
        Exit Sub
    End If
    
    ' Select target folder
    Dim targetFolder As String
    targetFolder = BrowseForFolder("Zielordner fuer Export auswaehlen")
    
    If targetFolder = "" Then
        Exit Sub
    End If
    
    ' Ensure trailing backslash
    If Right(targetFolder, 1) <> "\" Then
        targetFolder = targetFolder & "\"
    End If
    
    ' Export all eligible modules
    Dim exportedCount As Long
    exportedCount = 0
    
    Application.ScreenUpdating = False
    
    Dim vbComp As Object
    For Each vbComp In ThisWorkbook.VBProject.VBComponents
        If IsExportableModule(vbComp) Then
            Dim filePath As String
            filePath = targetFolder & vbComp.Name & GetModuleExtension(vbComp)
            
            vbComp.Export filePath
            exportedCount = exportedCount + 1
        End If
    Next vbComp
    
    Application.ScreenUpdating = True
    
    MsgBox "Export abgeschlossen:" & vbCrLf & vbCrLf & _
           "Exportiert: " & exportedCount & " Modul(e)" & vbCrLf & _
           "Zielordner: " & targetFolder, _
           vbInformation, "Export-Ergebnis"
    Exit Sub
    
ErrorHandler:
    Application.ScreenUpdating = True
    MsgBox "Fehler beim Exportieren: " & Err.Description, vbCritical, "Export-Fehler"
End Sub

' ============================================================
' EXPORT SELECTED MODULES
' ============================================================

' Show list of modules and export selected ones
Public Sub ExportSelectedModules()
    On Error GoTo ErrorHandler
    
    ' Check VBA project access
    If Not HasVBProjectAccess() Then
        MsgBox "Zugriff auf VBA-Projekt nicht moeglich." & vbCrLf & vbCrLf & _
               "Bitte aktivieren Sie 'Zugriff auf das VBA-Projektobjektmodell vertrauen'", _
               vbCritical, "VBA-Projektzugriff erforderlich"
        Exit Sub
    End If
    
    ' Get list of exportable modules
    Dim moduleList As String
    Dim moduleCount As Long
    moduleCount = 0
    
    Dim vbComp As Object
    For Each vbComp In ThisWorkbook.VBProject.VBComponents
        If IsExportableModule(vbComp) Then
            moduleCount = moduleCount + 1
            moduleList = moduleList & moduleCount & ". " & vbComp.Name & _
                        " (" & GetModuleTypeName(vbComp) & ")" & vbCrLf
        End If
    Next vbComp
    
    If moduleCount = 0 Then
        MsgBox "Keine exportierbaren Module gefunden.", vbInformation, "Keine Module"
        Exit Sub
    End If
    
    ' Ask user to select modules (by number, comma-separated)
    Dim selection As String
    selection = InputBox("Verfuegbare Module:" & vbCrLf & vbCrLf & _
                        moduleList & vbCrLf & _
                        "Geben Sie die Nummern der zu exportierenden Module ein" & vbCrLf & _
                        "(kommagetrennt, z.B. '1,3,5' oder 'alle' fuer alle):", _
                        "Module zum Exportieren auswaehlen", "alle")
    
    If selection = "" Then
        Exit Sub
    End If
    
    ' Select target folder
    Dim targetFolder As String
    targetFolder = BrowseForFolder("Zielordner fuer Export auswaehlen")
    
    If targetFolder = "" Then
        Exit Sub
    End If
    
    ' Ensure trailing backslash
    If Right(targetFolder, 1) <> "\" Then
        targetFolder = targetFolder & "\"
    End If
    
    ' Parse selection
    Dim selectedNumbers() As String
    Dim exportAll As Boolean
    exportAll = (LCase(Trim(selection)) = "alle" Or LCase(Trim(selection)) = "all")
    
    If Not exportAll Then
        selectedNumbers = Split(selection, ",")
    End If
    
    ' Export selected modules
    Dim exportedCount As Long
    exportedCount = 0
    Dim currentIndex As Long
    currentIndex = 0
    
    Application.ScreenUpdating = False
    
    For Each vbComp In ThisWorkbook.VBProject.VBComponents
        If IsExportableModule(vbComp) Then
            currentIndex = currentIndex + 1
            
            Dim shouldExport As Boolean
            shouldExport = False
            
            If exportAll Then
                shouldExport = True
            Else
                Dim j As Long
                For j = LBound(selectedNumbers) To UBound(selectedNumbers)
                    If Val(Trim(selectedNumbers(j))) = currentIndex Then
                        shouldExport = True
                        Exit For
                    End If
                Next j
            End If
            
            If shouldExport Then
                Dim filePath As String
                filePath = targetFolder & vbComp.Name & GetModuleExtension(vbComp)
                
                vbComp.Export filePath
                exportedCount = exportedCount + 1
            End If
        End If
    Next vbComp
    
    Application.ScreenUpdating = True
    
    MsgBox "Export abgeschlossen:" & vbCrLf & vbCrLf & _
           "Exportiert: " & exportedCount & " Modul(e)" & vbCrLf & _
           "Zielordner: " & targetFolder, _
           vbInformation, "Export-Ergebnis"
    Exit Sub
    
ErrorHandler:
    Application.ScreenUpdating = True
    MsgBox "Fehler beim Exportieren: " & Err.Description, vbCritical, "Export-Fehler"
End Sub

' ============================================================
' UTILITY FUNCTIONS
' ============================================================

' Check if VBA project access is enabled
Private Function HasVBProjectAccess() As Boolean
    On Error GoTo NoAccess
    
    Dim testAccess As Long
    testAccess = ThisWorkbook.VBProject.VBComponents.Count
    
    HasVBProjectAccess = True
    Exit Function
    
NoAccess:
    HasVBProjectAccess = False
End Function

' Check if module is exportable (.bas or .cls, not ThisWorkbook/Sheet)
Private Function IsExportableModule(ByVal vbComp As Object) As Boolean
    Select Case vbComp.Type
        Case vbext_ct_StdModule   ' Standard module (.bas)
            IsExportableModule = True
        Case vbext_ct_ClassModule ' Class module (.cls)
            IsExportableModule = True
        Case Else
            ' Document modules (ThisWorkbook, Sheets) and Forms - not exported
            IsExportableModule = False
    End Select
End Function

' Get file extension for module type
Private Function GetModuleExtension(ByVal vbComp As Object) As String
    Select Case vbComp.Type
        Case vbext_ct_StdModule
            GetModuleExtension = ".bas"
        Case vbext_ct_ClassModule
            GetModuleExtension = ".cls"
        Case vbext_ct_MSForm
            GetModuleExtension = ".frm"
        Case Else
            GetModuleExtension = ".bas"
    End Select
End Function

' Get human-readable module type name
Private Function GetModuleTypeName(ByVal vbComp As Object) As String
    Select Case vbComp.Type
        Case vbext_ct_StdModule
            GetModuleTypeName = "Standard"
        Case vbext_ct_ClassModule
            GetModuleTypeName = "Klasse"
        Case vbext_ct_MSForm
            GetModuleTypeName = "UserForm"
        Case vbext_ct_Document
            GetModuleTypeName = "Dokument"
        Case Else
            GetModuleTypeName = "Unbekannt"
    End Select
End Function

' Get filename from path
Private Function GetFileName(ByVal filePath As String) As String
    Dim pos As Long
    pos = InStrRev(filePath, "\")
    
    If pos > 0 Then
        GetFileName = Mid(filePath, pos + 1)
    Else
        GetFileName = filePath
    End If
End Function

' Get filename without extension
Private Function GetFileNameWithoutExtension(ByVal filePath As String) As String
    Dim fileName As String
    fileName = GetFileName(filePath)
    
    Dim pos As Long
    pos = InStrRev(fileName, ".")
    
    If pos > 0 Then
        GetFileNameWithoutExtension = Left(fileName, pos - 1)
    Else
        GetFileNameWithoutExtension = fileName
    End If
End Function

' Browse for folder using FileDialog
Private Function BrowseForFolder(ByVal dialogTitle As String) As String
    BrowseForFolder = ""
    
    Dim fd As FileDialog
    Set fd = Application.FileDialog(msoFileDialogFolderPicker)
    
    With fd
        .Title = dialogTitle
        .AllowMultiSelect = False
        
        If .Show = -1 Then
            BrowseForFolder = .SelectedItems(1)
        End If
    End With
End Function

' ============================================================
' QUICK ACCESS MACROS
' ============================================================

' List all modules (for information)
Public Sub ListAllModules()
    On Error GoTo ErrorHandler
    
    If Not HasVBProjectAccess() Then
        MsgBox "Zugriff auf VBA-Projekt nicht moeglich.", vbCritical, "Fehler"
        Exit Sub
    End If
    
    Dim moduleList As String
    moduleList = "Module in dieser Arbeitsmappe:" & vbCrLf & vbCrLf
    
    Dim vbComp As Object
    For Each vbComp In ThisWorkbook.VBProject.VBComponents
        moduleList = moduleList & "  " & vbComp.Name & _
                    " (" & GetModuleTypeName(vbComp) & ")" & vbCrLf
    Next vbComp
    
    MsgBox moduleList, vbInformation, "Modulliste"
    Exit Sub
    
ErrorHandler:
    MsgBox "Fehler: " & Err.Description, vbCritical, "Fehler"
End Sub
