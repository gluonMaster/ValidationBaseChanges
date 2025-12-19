Attribute VB_Name = "Kind_CleanupKinder"
Option Explicit

' =============================================================================
' Module: Kind_CleanupKinder
' Purpose: Main cleanup routine - removes Kinder rows that have KN status in Kartei
' Prefix: Kind_ (all modules/classes in this project use this prefix)
' =============================================================================

' -----------------------------------------------------------------------------
' Kind_CleanupKinder
' Purpose: Main entry point for Kinder cleanup based on Kartei KN status
' Process:
'   1. Verify Admin workbook is open
'   2. Verify Kinder and ErrorLog sheets exist
'   3. Build Kartei index once
'   4. Iterate Kinder rows bottom-up
'   5. Delete rows where Kartei match has KN status
'   6. Log errors/ambiguities to ErrorLog sheet
' -----------------------------------------------------------------------------
Public Sub Kind_CleanupKinder()
    Dim appState As Kind_AppState
    Dim wbAdmin As Workbook
    Dim wsKinder As Worksheet
    Dim wsErrorLog As Worksheet
    Dim karteiIndex As Object
    
    Dim lastRow As Long
    Dim r As Long
    Dim familyID As String
    Dim lastName As String
    Dim firstName As String
    Dim subject As String
    Dim subjectNorm As String
    
    Dim matchCount As Long
    Dim anyKN As Boolean
    Dim details As String
    Dim keyExists As Boolean
    
    Dim countChecked As Long
    Dim countDeleted As Long
    Dim countLogged As Long
    
    On Error GoTo ErrHandler
    
    ' Initialize counters
    countChecked = 0
    countDeleted = 0
    countLogged = 0
    
    ' ---------------------------------------------------------------------
    ' Step 1: Verify Admin workbook is open
    ' ---------------------------------------------------------------------
    If Not Kind_TryGetOpenWorkbookByName(KIND_ADMIN_WB_NAME, wbAdmin) Then
        MsgBox "Die Admin-Datei '" & KIND_ADMIN_WB_NAME & "' ist nicht geoeffnet." & vbCrLf & _
               "Bitte oeffnen Sie die Datei zuerst.", _
               vbExclamation, "Admin-Datei nicht gefunden"
        Exit Sub
    End If
    
    ' ---------------------------------------------------------------------
    ' Step 2: Verify Kinder and ErrorLog sheets exist in ThisWorkbook
    ' ---------------------------------------------------------------------
    If Not Kind_TryGetWorksheet(ThisWorkbook, KIND_THIS_SHEET_KINDER, wsKinder) Then
        MsgBox "Das Blatt '" & KIND_THIS_SHEET_KINDER & "' wurde in dieser Datei nicht gefunden.", _
               vbExclamation, "Blatt nicht gefunden"
        Exit Sub
    End If
    
    If Not Kind_TryGetWorksheet(ThisWorkbook, KIND_THIS_SHEET_ERRORLOG, wsErrorLog) Then
        MsgBox "Das Blatt '" & KIND_THIS_SHEET_ERRORLOG & "' wurde in dieser Datei nicht gefunden.", _
               vbExclamation, "Blatt nicht gefunden"
        Exit Sub
    End If
    
    ' ---------------------------------------------------------------------
    ' Step 3: Initialize application state and optimize for run
    ' ---------------------------------------------------------------------
    Set appState = New Kind_AppState
    appState.Capture
    appState.OptimizeForRun
    
    ' Ensure ErrorLog headers exist
    Kind_EnsureErrorLogHeaders wsErrorLog
    
    ' ---------------------------------------------------------------------
    ' Step 4: Build Kartei index once
    ' ---------------------------------------------------------------------
    Application.StatusBar = "Kartei-Index wird aufgebaut..."
    Set karteiIndex = Kind_BuildKarteiIndex(wbAdmin)
    
    If karteiIndex Is Nothing Then
        appState.Restore
        MsgBox "Fehler beim Aufbau des Kartei-Index." & vbCrLf & _
               "Bitte pruefen Sie, ob das Blatt 'Kartei' in der Admin-Datei existiert.", _
               vbExclamation, "Index-Fehler"
        Exit Sub
    End If
    
    ' ---------------------------------------------------------------------
    ' Step 5: Iterate Kinder rows bottom-up
    ' ---------------------------------------------------------------------
    lastRow = wsKinder.Cells(wsKinder.Rows.Count, KIND_KINDER_COL_FAMILYID).End(xlUp).Row
    
    ' Process from bottom to top (safe deletion)
    For r = lastRow To KIND_KINDER_FIRST_DATA_ROW Step -1
        ' Update status bar periodically
        If (lastRow - r + 1) Mod 50 = 0 Then
            Application.StatusBar = "Verarbeite Zeile " & r & " von " & lastRow & "..."
        End If
        
        ' Read Kinder data
        familyID = CStr(wsKinder.Range(KIND_KINDER_COL_FAMILYID & r).Value)
        lastName = CStr(wsKinder.Range(KIND_KINDER_COL_NACHNAME & r).Value)
        firstName = CStr(wsKinder.Range(KIND_KINDER_COL_VORNAME & r).Value)
        subject = CStr(wsKinder.Range(KIND_KINDER_COL_FACH & r).Value)
        
        ' Skip rows without FamilyID
        If Len(Trim$(familyID)) = 0 Then
            GoTo NextRow
        End If
        
        countChecked = countChecked + 1
        
        ' Normalize subject for comparison
        subjectNorm = Kind_NormalizeSubjectKinder(subject)
        
        ' Search in Kartei index
        keyExists = Kind_FindMatches(karteiIndex, familyID, lastName, firstName, subject, _
                                     matchCount, anyKN, details)
        
        ' Decision logic
        If Not keyExists Then
            ' Key not found - no FamilyID/Subject combination in Kartei
            Kind_LogError familyID, lastName, firstName, subject, "NO_KEY", _
                          "Kein Treffer in Kartei fuer FamilyID/Fach Kombination"
            countLogged = countLogged + 1
            
        ElseIf matchCount = 0 Then
            ' Key exists but no child name match
            Kind_LogError familyID, lastName, firstName, subject, "NO_CHILD", _
                          "FamilyID/Fach gefunden, aber Kind-Name stimmt nicht ueberein"
            countLogged = countLogged + 1
            
        ElseIf matchCount > 1 Then
            ' Multiple matches - ambiguous, do not delete
            Kind_LogError familyID, lastName, firstName, subject, "AMBIGUOUS", _
                          "Mehrere Treffer gefunden: " & details
            countLogged = countLogged + 1
            
        ElseIf matchCount = 1 Then
            ' Exactly one match found
            If anyKN Then
                ' Status is KN - delete the row
                wsKinder.Rows(r).Delete
                countDeleted = countDeleted + 1
            End If
            ' If not KN, simply leave the row (no action needed)
        End If
        
NextRow:
    Next r
    
    ' ---------------------------------------------------------------------
    ' Step 6: Restore application state and show summary
    ' ---------------------------------------------------------------------
CleanExit:
    ' Restore application state (safe even if not captured)
    On Error Resume Next
    If Not appState Is Nothing Then
        appState.Restore
    End If
    On Error GoTo 0
    
    MsgBox "Verarbeitung abgeschlossen." & vbCrLf & vbCrLf & _
           "Geprueft: " & countChecked & " Zeilen" & vbCrLf & _
           "Geloescht: " & countDeleted & " Zeilen (Status KN)" & vbCrLf & _
           "Protokolliert: " & countLogged & " Eintraege (siehe ErrorLog)", _
           vbInformation, "Kinder-Bereinigung"
    
    Exit Sub
    
ErrHandler:
    ' Capture error details before any further operations
    Dim errNum As Long
    Dim errDesc As String
    errNum = Err.Number
    errDesc = Err.Description
    
    ' Log runtime error to ErrorLog sheet
    On Error Resume Next
    Kind_LogRuntimeError errNum, errDesc, "Kind_CleanupKinder (Zeile " & r & ")"
    
    ' Ensure state is restored even on error
    If Not appState Is Nothing Then
        appState.Restore
    End If
    On Error GoTo 0
    
    MsgBox "Ein Fehler ist aufgetreten:" & vbCrLf & _
           "Fehler " & errNum & ": " & errDesc & vbCrLf & vbCrLf & _
           "Zeile: " & r & vbCrLf & _
           "Geprueft: " & countChecked & vbCrLf & _
           "Geloescht: " & countDeleted & vbCrLf & _
           "Protokolliert: " & countLogged, _
           vbCritical, "Laufzeitfehler"
End Sub
