Attribute VB_Name = "Export_ManualImport"
'==========================
'   Manual Database Import Module
'   Allows Admin to manually refresh Kartei data from Access database
'   without closing and reopening the file
'==========================
Option Explicit

' ========================================
' Main Public Procedure for Manual Import
' ========================================

Public Sub ManualImportFromDatabase()
    ' Manually imports data from Access database to Kartei sheet
    ' Checks for unsaved changes before proceeding
    ' Does NOT show frmListNachHilfe form after import
    
    On Error GoTo ErrorHandler
    
    Dim wsKartei As Worksheet
    Set wsKartei = ThisWorkbook.Worksheets("Kartei")
    
    ' Step 1: Check for unsaved changes
    If HasUnsavedChanges(wsKartei) Then
        MsgBox "Es gibt ungespeicherte Aenderungen auf dem Kartei-Blatt." & vbCrLf & vbCrLf & _
               "Bitte speichern Sie Ihre Aenderungen zuerst durch Klicken auf die 'zuBase'-Schaltflaeche, " & _
               "bevor Sie Daten aus der Datenbank importieren.", _
               vbExclamation, "Ungespeicherte Aenderungen erkannt"
        Exit Sub
    End If
    
    ' Step 2: Confirm import action
    Dim userResponse As VbMsgBoxResult
    userResponse = MsgBox("Dies laedt alle Daten aus der Access-Datenbank neu." & vbCrLf & vbCrLf & _
                          "Die lokale Anzeige wird durch den aktuellen Datenbankzustand ersetzt " & _
                          "(einschliesslich ausstehender und abgelehnter Datensaetze vom Superadmin)." & vbCrLf & vbCrLf & _
                          "Moechten Sie fortfahren?", _
                          vbQuestion + vbYesNo, "Datenbankimport bestaetigen")
    
    If userResponse <> vbYes Then
        Exit Sub
    End If
    
    ' Step 3: Perform import
    Call PerformDatabaseImport(wsKartei)
    
    MsgBox "Datenbankimport erfolgreich abgeschlossen." & vbCrLf & vbCrLf & _
           "Das Kartei-Blatt zeigt jetzt den aktuellen Stand der Datenbank.", _
           vbInformation, "Import abgeschlossen"
    
    ' Navigate to Kartei sheet
    wsKartei.Activate
    wsKartei.Cells(3, 1).Select
    
    Exit Sub
    
ErrorHandler:
    Application.Calculation = xlCalculationAutomatic
    Application.ScreenUpdating = True
    Application.EnableEvents = True
    
    MsgBox "Beim Datenbankimport ist ein Fehler aufgetreten:" & vbCrLf & vbCrLf & _
           "Fehler " & Err.Number & ": " & Err.Description, _
           vbCritical, "Importfehler"
End Sub

' ========================================
' Check for Unsaved Changes
' ========================================

Private Function HasUnsavedChanges(ByVal wsKartei As Worksheet) As Boolean
    ' Compares Kartei with Kartei_Original to detect changes
    ' Returns True if there are unsaved changes
    
    On Error GoTo ErrorHandler
    
    Dim wsOriginal As Worksheet
    
    ' Check if Kartei_Original exists
    On Error Resume Next
    Set wsOriginal = ThisWorkbook.Worksheets("Kartei_Original")
    On Error GoTo ErrorHandler
    
    If wsOriginal Is Nothing Then
        ' No Kartei_Original means no baseline to compare
        ' Assume no changes (or first run)
        HasUnsavedChanges = False
        Exit Function
    End If
    
    ' Read both sheets into dictionaries using ID (column 48 = AV) as key
    Dim dictLocal As Scripting.Dictionary
    Dim dictOriginal As Scripting.Dictionary
    
    Set dictLocal = ReadSheetIntoDictionary_ID(wsKartei, 3, 51)
    Set dictOriginal = ReadSheetIntoDictionary_ID(wsOriginal, 3, 51)
    
    ' Find changed IDs
    Dim changedIDs As Collection
    Set changedIDs = FindChangedIDs(dictLocal, dictOriginal)
    
    ' If there are any changed IDs, we have unsaved changes
    HasUnsavedChanges = (changedIDs.Count > 0)
    
    Exit Function
    
ErrorHandler:
    ' In case of error, assume there might be changes (safer approach)
    HasUnsavedChanges = True
End Function

' ========================================
' Perform the Actual Import
' ========================================

Private Sub PerformDatabaseImport(ByVal wsKartei As Worksheet)
    ' Performs the actual database import
    ' Replicates Workbook_Open import logic but without showing frmListNachHilfe
    
    On Error GoTo ErrorHandler
    
    Application.ScreenUpdating = False
    Application.Calculation = xlCalculationManual
    Application.EnableEvents = False
    
    ' Step 0: Reset sheet state (filters, hidden rows/columns, etc.)
    Call ResetSheetStateBeforeImport(wsKartei)
    
    ' Step 1: Import main Kartei data from tblKartei
    ' Note: ImportKarteiAndFormat_Optimized is in DieseArbeitsmappe (ThisWorkbook class)
    Call DieseArbeitsmappe.ImportKarteiAndFormat_Optimized
    
    ' Step 2: Overlay pending and declined records
    Call OverlayPendingAndDeclined(wsKartei)
    
    ' Step 3: Apply cell formatting conversion
    Call ConvertAndFormatCellsOptimized
    
    ' Step 4: Rebuild Kartei_Original with new data
    ' Note: RebuildKarteiOriginalSheet is in DieseArbeitsmappe (ThisWorkbook class)
    Call DieseArbeitsmappe.RebuildKarteiOriginalSheet(wsKartei)
    
Cleanup:
    Application.EnableEvents = True
    Application.Calculation = xlCalculationAutomatic
    Application.ScreenUpdating = True
    Exit Sub
    
ErrorHandler:
    Application.EnableEvents = True
    Application.Calculation = xlCalculationAutomatic
    Application.ScreenUpdating = True
    Err.Raise Err.Number, Err.Source, Err.Description
End Sub

' ========================================
' Reset Sheet State Before Import
' ========================================

Private Sub ResetSheetStateBeforeImport(ByVal ws As Worksheet)
    ' Resets the Kartei sheet to a clean state before importing data
    ' This prevents mixing old and new data due to filters, hidden rows/columns, etc.
    
    On Error Resume Next
    
    ' 1. Clear any active AutoFilter (remove filter criteria, not the arrows)
    If ws.AutoFilterMode Then
        ' Check if any filter is actually applied
        If ws.FilterMode Then
            ws.ShowAllData
        End If
    End If
    
    On Error GoTo ErrorHandler
    
    ' 2. Unhide all rows in the data area (from row 3 onwards)
    Dim lastRowToCheck As Long
    lastRowToCheck = ws.Rows.Count
    
    ' Only process rows that might contain data (optimize for large sheets)
    Dim usedLastRow As Long
    usedLastRow = ws.Cells(ws.Rows.Count, 1).End(xlUp).Row
    If usedLastRow < 3 Then usedLastRow = 1000  ' Default range if sheet is empty
    
    ' Unhide all rows from 3 to usedLastRow + buffer
    ws.Rows("3:" & (usedLastRow + 100)).Hidden = False
    
    ' 3. Unhide all columns in the data area (A to BA = columns 1 to 53)
    ws.Columns("A:BA").Hidden = False
    
    ' 4. Clear all data from row 3 onwards (content, formats, colors)
    '    This ensures no old data remains that could mix with new import
    Dim lastRow As Long
    lastRow = ws.Cells(ws.Rows.Count, 1).End(xlUp).Row
    
    If lastRow >= 3 Then
        ' Clear contents
        ws.Range("A3:BA" & lastRow).ClearContents
        
        ' Reset interior colors
        ws.Range("A3:BA" & lastRow).Interior.ColorIndex = xlColorIndexNone
        
        ' Reset font colors to automatic
        ws.Range("A3:BA" & lastRow).Font.ColorIndex = xlColorIndexAutomatic
        
        ' Reset row heights to default
        ws.Range("A3:BA" & lastRow).RowHeight = 16
    End If
    
    ' 5. Remove any grouping/outline (if present)
    On Error Resume Next
    ws.Outline.ShowLevels RowLevels:=8, ColumnLevels:=8
    On Error GoTo ErrorHandler
    
    ' 6. Scroll to top-left to ensure consistent view
    ws.Activate
    Application.Goto ws.Range("A3"), True
    
    Exit Sub
    
ErrorHandler:
    ' Log error but continue - sheet reset is best-effort
    Debug.Print "ResetSheetStateBeforeImport error: " & Err.Description
    Err.Clear
End Sub

' ========================================
' Quick Check Function (for external use)
' ========================================

Public Function CheckForUnsavedChanges() As Boolean
    ' Public wrapper to check for unsaved changes
    ' Can be called from other modules or UI buttons
    
    Dim wsKartei As Worksheet
    Set wsKartei = ThisWorkbook.Worksheets("Kartei")
    
    CheckForUnsavedChanges = HasUnsavedChanges(wsKartei)
End Function

' ========================================
' Button Handler (for ribbon/sheet button)
' ========================================

Public Sub btnImportDatabase_Click()
    ' Handler for Import Database button
    ' Simply calls the main import procedure
    
    Call ManualImportFromDatabase
End Sub
