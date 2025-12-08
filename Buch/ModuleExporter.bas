Attribute VB_Name = "ModuleExporter"
'==========================
'   Module Exporter
'   Exports all .bas modules from current workbook to user-selected folder
'==========================
Option Explicit

' ========================================
' Main Export Procedure
' ========================================

Public Sub ExportAllModules()
    ' Exports all VBA modules (.bas) from current workbook to a user-selected folder
    ' Displays a folder picker dialog and exports each module as separate .bas file
    
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
    
    ' Show folder picker dialog
    Dim exportPath As String
    exportPath = SelectExportFolder()
    
    If exportPath = "" Then
        ' User cancelled
        Exit Sub
    End If
    
    ' Ensure path ends with backslash
    If Right(exportPath, 1) <> "\" Then
        exportPath = exportPath & "\"
    End If
    
    Application.ScreenUpdating = False
    
    ' Export all standard modules
    Dim exportCount As Long
    Dim skippedCount As Long
    exportCount = 0
    skippedCount = 0
    
    Dim vbProj As Object
    Set vbProj = ThisWorkbook.VBProject
    
    Dim vbComp As Object
    For Each vbComp In vbProj.VBComponents
        ' Export only standard modules (vbext_ct_StdModule = 1)
        If vbComp.Type = 1 Then
            Dim fileName As String
            fileName = exportPath & vbComp.name & ".bas"
            
            On Error Resume Next
            vbComp.Export fileName
            
            If Err.Number = 0 Then
                exportCount = exportCount + 1
            Else
                skippedCount = skippedCount + 1
                Debug.Print "Warning: Could not export " & vbComp.name & ": " & Err.Description
                Err.Clear
            End If
            On Error GoTo ErrorHandler
        End If
    Next vbComp
    
    Application.ScreenUpdating = True
    
    ' Show results
    Dim msg As String
    msg = "Export abgeschlossen!" & vbCrLf & vbCrLf & _
          "Exportiert: " & exportCount & " Modul(e)" & vbCrLf & _
          "Zielordner: " & exportPath
    
    If skippedCount > 0 Then
        msg = msg & vbCrLf & vbCrLf & "Uebersprungen: " & skippedCount & " Modul(e)"
    End If
    
    MsgBox msg, vbInformation, "Module exportiert"
    
    Exit Sub
    
ErrorHandler:
    Application.ScreenUpdating = True
    MsgBox "Fehler beim Exportieren der Module: " & Err.Description, vbCritical, "Exportfehler"
End Sub

Public Sub ExportAllModulesAndClasses()
    ' Exports all VBA modules (.bas) and class modules (.cls) from current workbook
    ' to a user-selected folder
    
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
    
    ' Show folder picker dialog
    Dim exportPath As String
    exportPath = SelectExportFolder()
    
    If exportPath = "" Then
        Exit Sub
    End If
    
    If Right(exportPath, 1) <> "\" Then
        exportPath = exportPath & "\"
    End If
    
    Application.ScreenUpdating = False
    
    Dim basCount As Long
    Dim clsCount As Long
    Dim skippedCount As Long
    basCount = 0
    clsCount = 0
    skippedCount = 0
    
    Dim vbProj As Object
    Set vbProj = ThisWorkbook.VBProject
    
    Dim vbComp As Object
    For Each vbComp In vbProj.VBComponents
        Dim fileName As String
        Dim exported As Boolean
        exported = False
        
        Select Case vbComp.Type
            Case 1  ' vbext_ct_StdModule - Standard module
                fileName = exportPath & vbComp.name & ".bas"
                On Error Resume Next
                vbComp.Export fileName
                If Err.Number = 0 Then
                    basCount = basCount + 1
                    exported = True
                End If
                Err.Clear
                On Error GoTo ErrorHandler
                
            Case 2  ' vbext_ct_ClassModule - Class module
                fileName = exportPath & vbComp.name & ".cls"
                On Error Resume Next
                vbComp.Export fileName
                If Err.Number = 0 Then
                    clsCount = clsCount + 1
                    exported = True
                End If
                Err.Clear
                On Error GoTo ErrorHandler
        End Select
        
        If Not exported And (vbComp.Type = 1 Or vbComp.Type = 2) Then
            skippedCount = skippedCount + 1
            Debug.Print "Warning: Could not export " & vbComp.name
        End If
    Next vbComp
    
    Application.ScreenUpdating = True
    
    Dim msg As String
    msg = "Export abgeschlossen!" & vbCrLf & vbCrLf & _
          "Standard-Module (.bas): " & basCount & vbCrLf & _
          "Klassen-Module (.cls): " & clsCount & vbCrLf & _
          "Zielordner: " & exportPath
    
    If skippedCount > 0 Then
        msg = msg & vbCrLf & vbCrLf & "Uebersprungen: " & skippedCount
    End If
    
    MsgBox msg, vbInformation, "Module exportiert"
    
    Exit Sub
    
ErrorHandler:
    Application.ScreenUpdating = True
    MsgBox "Fehler beim Exportieren: " & Err.Description, vbCritical, "Exportfehler"
End Sub

' ========================================
' Helper Functions
' ========================================

Private Function SelectExportFolder() As String
    ' Shows folder picker dialog and returns selected path
    ' Returns empty string if user cancels
    
    Dim folderDialog As fileDialog
    Set folderDialog = Application.fileDialog(msoFileDialogFolderPicker)
    
    With folderDialog
        .Title = "Zielordner fuer Modulexport waehlen"
        .AllowMultiSelect = False
        .InitialFileName = ThisWorkbook.Path & "\"
        
        If .Show = -1 Then
            SelectExportFolder = .SelectedItems(1)
        Else
            SelectExportFolder = ""
        End If
    End With
End Function

Private Function IsVBAAccessTrusted() As Boolean
    ' Checks if VBA project access is allowed in Trust Center settings
    ' Returns True if access is granted, False otherwise
    
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
