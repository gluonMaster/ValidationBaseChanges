Attribute VB_Name = "nj_NewYearEntry"
Option Explicit

' ============================================================================
' Module: nj_NewYearEntry
' Purpose: Top-level orchestrator for new year preparation scenario.
'          Coordinates data transformation, workbook saving, and Access
'          database creation for the transition to the new year.
' ============================================================================

Private Const NEW_YEAR As Long = 26
Private Const FULL_YEAR As Long = 2026

' ============================================================================
' PUBLIC ENTRY POINT
' ============================================================================

Public Sub ImNeuenJahr_26()
    ' Main entry point for new year 2026 preparation.
    ' Orchestrates the complete workflow:
    '   1. User confirmation
    '   2. Data format normalization
    '   3. Kartei data transformation
    '   4. Header updates
    '   5. Sorting
    '   6. Workbook save as new file
    '   7. Access database creation
    
    Dim ws As Worksheet
    
    ' Activate Kartei sheet
    Set ws = ThisWorkbook.Worksheets("Kartei")
    ws.Activate
    
    ' Check if there is data to process
    If Not nj_HasDataToProcess(ws) Then
        MsgBox "Keine Daten zum Verarbeiten gefunden (weniger als 3 Zeilen).", _
               vbInformation, "Neues Jahr Vorbereitung"
        Exit Sub
    End If
    
    ' Step 1: User confirmation
    If Not nj_ConfirmNewYearPreparation() Then
        Exit Sub
    End If
    
    ' Enable performance optimizations
    nj_SetPerformanceMode True
    
    On Error GoTo CleanupOnError
    
    ' Step 2: Normalize formats before transformation
    Call ConvertAndFormatCellsOptimized
    
    ' Step 3: Transform Kartei data for new year
    Call nj_PrepareKarteiDataForNewYear(NEW_YEAR)
    
    ' Step 4: Update headers
    nj_UpdateKarteiHeaders ws
    
    ' Step 5: Sort Kartei
    Call SortNameZ
    
    ' Restore normal mode before user dialogs
    nj_SetPerformanceMode False
    
    ' Step 6: Save workbook as new year file and get the save folder path
    Dim saveFolderPath As String
    If Not nj_SaveWorkbookAsNewYear(NEW_YEAR, saveFolderPath) Then
        MsgBox "Speichern abgebrochen. Die Vorbereitung wurde nicht abgeschlossen.", _
               vbExclamation, "Neues Jahr Vorbereitung"
        Exit Sub
    End If
    
    ' Step 7: Update X1 with the new base path for database connection
    nj_UpdateBasePath ws, saveFolderPath
    
    ' Re-enable performance mode for database export
    nj_SetPerformanceMode True
    
    ' Step 8: Create Access database in [saveFolderPath]\Alarm\
    Call nj_CreateAccessDbForYear(NEW_YEAR, saveFolderPath)
    
    ' Step 8.5: Rebuild Kartei_Original baseline for the new year
    ' This prevents CompareAndSyncKartei from treating all rows as changed
    Call RebuildKarteiOriginal
    
    ' Restore normal mode
    nj_SetPerformanceMode False
    
    ' Step 9: Save workbook again to persist X1 change and new baseline
    ThisWorkbook.Save
    
    ' Final confirmation
    MsgBox "Vorbereitung fuer das Jahr " & FULL_YEAR & " wurde erfolgreich abgeschlossen.", _
           vbInformation, "Neues Jahr Vorbereitung"
    Exit Sub
    
CleanupOnError:
    ' Ensure Excel is restored to normal state even on error
    nj_SetPerformanceMode False
    MsgBox "Fehler bei der Vorbereitung:" & vbCrLf & Err.Description, _
           vbCritical, "Neues Jahr Vorbereitung"
End Sub

' ============================================================================
' PRIVATE HELPERS: PERFORMANCE
' ============================================================================

Private Sub nj_SetPerformanceMode(ByVal enableOptimization As Boolean)
    ' Toggles Excel performance settings for faster processing.
    ' Call with True before heavy operations, False after completion.
    '
    ' Settings affected:
    '   - ScreenUpdating: Prevents screen refresh during operations
    '   - Calculation: Switches to manual to avoid recalc on each cell change
    '   - EnableEvents: Prevents event triggers during bulk operations
    
    If enableOptimization Then
        Application.ScreenUpdating = False
        Application.Calculation = xlCalculationManual
        Application.EnableEvents = False
    Else
        Application.EnableEvents = True
        Application.Calculation = xlCalculationAutomatic
        Application.ScreenUpdating = True
    End If
End Sub

' ============================================================================
' PRIVATE HELPERS
' ============================================================================

Private Function nj_HasDataToProcess(ws As Worksheet) As Boolean
    ' Checks if Kartei has data rows (at least row 3 with content).
    
    Dim lastRow As Long
    lastRow = ws.Cells(ws.Rows.Count, 1).End(xlUp).Row
    
    nj_HasDataToProcess = (lastRow >= 3)
End Function

Private Function nj_ConfirmNewYearPreparation() As Boolean
    ' Displays confirmation dialog before starting the transformation.
    ' Returns True if user confirms, False otherwise.
    
    Dim response As VbMsgBoxResult
    
    response = MsgBox("Moechten Sie die Daten fuer das Jahr " & FULL_YEAR & " vorbereiten?" & vbCrLf & vbCrLf & _
                      "Diese Aktion kann nicht rueckgaengig gemacht werden.", _
                      vbYesNo + vbQuestion, "Neues Jahr Vorbereitung")
    
    nj_ConfirmNewYearPreparation = (response = vbYes)
End Function

Private Sub nj_UpdateKarteiHeaders(ws As Worksheet)
    ' Updates header cells with new year labels.
    
    ws.Range("A1").Value = 1
    ws.Range("J2").Value = "Gruppe " & NEW_YEAR & "/1"
    ws.Range("O2").Value = "Gruppe " & NEW_YEAR & "/2"
    ws.Range("M2").Value = "Preis " & NEW_YEAR & "/1"
    ws.Range("R2").Value = "Preis " & NEW_YEAR & "/2"
End Sub

Private Function nj_SaveWorkbookAsNewYear(ByVal targetYear As Long, ByRef outFolderPath As String) As Boolean
    ' Saves the current workbook with new year filename.
    ' Shows folder picker dialog for user to select destination.
    ' Returns True on success, False if cancelled or failed.
    ' outFolderPath: Returns the selected folder path (without trailing backslash).
    
    Dim fileName As String
    Dim folderPath As String
    Dim fullPath As String
    Dim fd As FileDialog
    
    nj_SaveWorkbookAsNewYear = False
    outFolderPath = ""
    
    ' Build target filename
    fileName = "KindElternDaten_" & targetYear & "_Admin.xlsm"
    
    ' Show folder picker dialog
    Set fd = Application.FileDialog(msoFileDialogFolderPicker)
    
    With fd
        .Title = "Ordner fuer die neue Datei auswaehlen"
        .ButtonName = "Auswaehlen"
        .InitialFileName = ThisWorkbook.Path & "\"
        
        If .Show = -1 Then
            folderPath = .SelectedItems(1)
        Else
            ' User cancelled
            Exit Function
        End If
    End With
    
    ' Normalize folder path (remove trailing backslash for consistency)
    If Right(folderPath, 1) = "\" Then
        folderPath = Left(folderPath, Len(folderPath) - 1)
    End If
    
    ' Build full path for file
    fullPath = folderPath & "\" & fileName
    
    ' Save workbook
    On Error GoTo SaveError
    ThisWorkbook.SaveAs Filename:=fullPath, FileFormat:=xlOpenXMLWorkbookMacroEnabled
    On Error GoTo 0
    
    ' Return the folder path to caller
    outFolderPath = folderPath
    nj_SaveWorkbookAsNewYear = True
    Exit Function
    
SaveError:
    MsgBox "Fehler beim Speichern der Datei:" & vbCrLf & Err.Description, _
           vbCritical, "Speicherfehler"
    nj_SaveWorkbookAsNewYear = False
End Function

Private Sub nj_UpdateBasePath(ws As Worksheet, ByVal folderPath As String)
    ' Updates Kartei!X1 with the new base path.
    ' This ensures the migrated file will find its Access database on next open.
    
    ws.Range("X1").Value = folderPath
End Sub
