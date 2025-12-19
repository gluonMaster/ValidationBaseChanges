Attribute VB_Name = "indBlank_Main"
Option Explicit

' =============================================================================
' Module: indBlank_Main
' Purpose: Main entry point for individual blank (Einzelmeldung) generation
'          from filtered Kartei rows in KindElternDaten_25_Admin.xlsm
' =============================================================================

' =============================================================================
' Sub: indBlank_GenerateIndBlanksFromFilteredKartei
' Purpose: Main entry point - generates individual blanks for visible Kartei rows
' =============================================================================
Public Sub indBlank_GenerateIndBlanksFromFilteredKartei()
    Dim wsKartei As Worksheet
    Dim wbTemplate As Workbook
    Dim wsMuster As Worksheet
    Dim appState As indBlank_AppState
    Dim targetFolder As String
    Dim yearValue As Long
    Dim monthValue As Long
    Dim createdCount As Long
    Dim skippedCount As Long
    Dim errorCount As Long
    Dim errMsg As String
    
    On Error GoTo ErrorHandler
    
    ' -------------------------------------------------------------------------
    ' Step 1: Find Kartei worksheet in ThisWorkbook
    ' -------------------------------------------------------------------------
    If Not indBlank_TryGetWorksheet(ThisWorkbook, indBlank_Config.KARTEI_SHEET, wsKartei) Then
        MsgBox "Blatt '" & indBlank_Config.KARTEI_SHEET & "' wurde nicht gefunden.", _
               vbCritical, "Fehler"
        Exit Sub
    End If
    
    ' -------------------------------------------------------------------------
    ' Step 2: Confirm if no filter is active (optional safety check)
    ' -------------------------------------------------------------------------
    If Not indBlank_ConfirmIfNoFilter(wsKartei) Then
        Exit Sub
    End If
    
    ' -------------------------------------------------------------------------
    ' Step 3: Check if template file exists
    ' -------------------------------------------------------------------------
    If Dir(indBlank_Config.TEMPLATE_PATH) = "" Then
        MsgBox "Vorlage nicht gefunden:" & vbCrLf & indBlank_Config.TEMPLATE_PATH, _
               vbCritical, "Fehler"
        Exit Sub
    End If
    
    ' -------------------------------------------------------------------------
    ' Step 4: Ask for year
    ' -------------------------------------------------------------------------
    If Not indBlank_AskYear(yearValue) Then
        Exit Sub
    End If
    
    ' -------------------------------------------------------------------------
    ' Step 5: Ask for month
    ' -------------------------------------------------------------------------
    If Not indBlank_AskMonth(monthValue) Then
        Exit Sub
    End If
    
    ' -------------------------------------------------------------------------
    ' Step 6: Select target folder
    ' -------------------------------------------------------------------------
    targetFolder = indBlank_SelectTargetFolder()
    If Len(targetFolder) = 0 Then
        Exit Sub
    End If
    
    ' -------------------------------------------------------------------------
    ' Step 7: Optimize Excel state for performance
    ' -------------------------------------------------------------------------
    Set appState = New indBlank_AppState
    appState.Capture
    appState.OptimizeForRun "Generiere Einzelmeldungen, bitte warten..."
    
    ' -------------------------------------------------------------------------
    ' Step 8: Open template workbook (ReadOnly)
    ' -------------------------------------------------------------------------
    Set wbTemplate = Workbooks.Open( _
        Filename:=indBlank_Config.TEMPLATE_PATH, _
        ReadOnly:=True, _
        UpdateLinks:=False)
    
    ' Get Muster worksheet from template
    If Not indBlank_TryGetWorksheet(wbTemplate, indBlank_Config.TEMPLATE_SHEET, wsMuster) Then
        wbTemplate.Close SaveChanges:=False
        appState.Restore
        MsgBox "Blatt '" & indBlank_Config.TEMPLATE_SHEET & "' nicht in Vorlage gefunden.", _
               vbCritical, "Fehler"
        Exit Sub
    End If
    
    ' -------------------------------------------------------------------------
    ' Step 9: Process visible Kartei rows
    ' -------------------------------------------------------------------------
    If Not indBlank_ProcessVisibleKarteiRows( _
            wsKartei, wbTemplate, wsMuster, targetFolder, yearValue, _
            monthValue, _
            createdCount, skippedCount, errorCount) Then
        ' Fatal error during processing
        errMsg = "Ein schwerwiegender Fehler ist aufgetreten."
    End If
    
    ' -------------------------------------------------------------------------
    ' Step 10: Cleanup - close template workbook
    ' -------------------------------------------------------------------------
    On Error Resume Next
    wbTemplate.Close SaveChanges:=False
    Set wsMuster = Nothing
    Set wbTemplate = Nothing
    On Error GoTo ErrorHandler
    
    ' -------------------------------------------------------------------------
    ' Step 11: Restore Excel state and show summary
    ' -------------------------------------------------------------------------
    appState.Restore
    Set appState = Nothing
    
    indBlank_ShowSummary createdCount, skippedCount, errorCount
    
    Exit Sub
    
ErrorHandler:
    ' Emergency cleanup
    On Error Resume Next
    
    If Not wbTemplate Is Nothing Then
        wbTemplate.Close SaveChanges:=False
    End If
    Set wsMuster = Nothing
    Set wbTemplate = Nothing
    
    If Not appState Is Nothing Then
        appState.Restore
    End If
    Set appState = Nothing
    
    On Error GoTo 0
    
    MsgBox "Ein unerwarteter Fehler ist aufgetreten:" & vbCrLf & _
           Err.Description, vbCritical, "Fehler"
End Sub
