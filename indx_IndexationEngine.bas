Attribute VB_Name = "indx_IndexationEngine"
Option Explicit

' =============================================================================
' indx_IndexationEngine.bas
' General indexation engine supporting parameterized monthly salary corrections
' =============================================================================

' Configuration Type for indexation scenarios
Public Type IndxConfig
    ' Month identification
    MonthName As String          ' e.g., "August", "December"
    
    ' Kartei column mappings (1-based column numbers)
    SourceValueCol As Long       ' Column to read "old" value (War)
    BaseValueCol As Long         ' Column used as base for calculation
    TargetWriteCol As Long       ' Column to write result back
    
    ' Calculation parameters
    Coefficient As Double        ' Multiplier (e.g., 0.75)
    
    ' Indexation sheet headers (columns F, G, H)
    HeaderWar As String          ' e.g., "Aug War", "Dez War"
    HeaderWird As String         ' e.g., "Aug Wird", "Dez Wird"
    HeaderBase As String         ' e.g., "Sept", "" (empty for December)
    
    ' Flags
    FillBaseColumn As Boolean    ' Whether to fill column H with base values
End Type

' -----------------------------------------------------------------------------
' indx_CreateAugustConfig
' Creates configuration for August indexation scenario
' -----------------------------------------------------------------------------
Public Function indx_CreateAugustConfig() As IndxConfig
    Dim cfg As IndxConfig
    
    cfg.MonthName = "August"
    cfg.SourceValueCol = 28      ' AB - August current value
    cfg.BaseValueCol = 29        ' AC - September (base for calculation)
    cfg.TargetWriteCol = 28      ' AB - write result back
    cfg.Coefficient = 0.75
    cfg.HeaderWar = "Aug War"
    cfg.HeaderWird = "Aug Wird"
    cfg.HeaderBase = "Sept"
    cfg.FillBaseColumn = True
    
    indx_CreateAugustConfig = cfg
End Function

' -----------------------------------------------------------------------------
' indx_CreateDecemberConfig
' Creates configuration for December indexation scenario
' -----------------------------------------------------------------------------
Public Function indx_CreateDecemberConfig() As IndxConfig
    Dim cfg As IndxConfig
    
    cfg.MonthName = "December"
    cfg.SourceValueCol = 32      ' AF - December current value
    cfg.BaseValueCol = 32        ' AF - December is also the base
    cfg.TargetWriteCol = 32      ' AF - write result back
    cfg.Coefficient = 0.75
    cfg.HeaderWar = "Dez War"
    cfg.HeaderWird = "Dez Wird"
    cfg.HeaderBase = ""          ' Empty - no base column header
    cfg.FillBaseColumn = False   ' Do not fill column H
    
    indx_CreateDecemberConfig = cfg
End Function

' -----------------------------------------------------------------------------
' indx_RunIndexation
' Main entry point for running indexation with given configuration
' -----------------------------------------------------------------------------
Public Sub indx_RunIndexation(cfg As IndxConfig)
    Dim originalScreenUpdating As Boolean
    Dim originalCalculation As Long
    
    ' Store original settings and optimize performance
    originalScreenUpdating = Application.ScreenUpdating
    originalCalculation = Application.Calculation
    
    Application.ScreenUpdating = False
    Application.Calculation = xlCalculationManual
    
    On Error GoTo ErrorHandler
    
    ' Step 1: Navigate to Kartei sheet
    Worksheets("Kartei").Activate
    
    ' Step 2: Remove existing Indexation sheet if exists
    Call indx_Engine_RemoveExistingSheet
    
    ' Step 3: Clear active filters if any
    Call indx_ClearActiveFilters
    
    ' Step 4: Execute SortNameZ macro
    Call SortNameZ
    
    ' Step 5: Get selected data using extended function
    Dim selectedData As Variant
    selectedData = indx_GetSelectedDataExtended(cfg)
    
    If IsEmpty(selectedData) Then
        MsgBox "Keine passenden Datensaetze gefunden.", vbInformation, "Information"
        GoTo Cleanup
    End If
    
    ' Step 6: Create and setup Indexation sheet
    Call indx_Engine_CreateSheet(selectedData, cfg)
    
    ' Step 7: Create OK button with config reference
    Call indx_Engine_CreateButton(cfg)
    
    ' Fall through to Cleanup
    
Cleanup:
    ' Restore original settings (guaranteed execution)
    Application.ScreenUpdating = originalScreenUpdating
    Application.Calculation = originalCalculation
    Exit Sub
    
ErrorHandler:
    MsgBox "Fehler beim Ausfuehren der " & cfg.MonthName & "-Indexierung: " & Err.Description, _
           vbCritical, "Fehler"
    Resume Cleanup
End Sub

' -----------------------------------------------------------------------------
' indx_GetSelectedDataExtended
' Extends the base selection with additional columns based on config
' Uses indx_GetSelectedRows() as SINGLE SOURCE OF TRUTH for row selection
' -----------------------------------------------------------------------------
Public Function indx_GetSelectedDataExtended(cfg As IndxConfig) As Variant
    Dim ws As Worksheet
    Dim extendedData As Variant
    Dim i As Long
    Dim rowCount As Long
    Dim selectedRows As Collection
    
    Set ws = Worksheets("Kartei")
    
    ' Use indx_GetSelectedRows as single source of truth for row selection
    Set selectedRows = indx_GetSelectedRows()
    
    ' If no matching rows found
    If selectedRows.Count = 0 Then
        indx_GetSelectedDataExtended = Empty
        Exit Function
    End If
    
    ' Create result array: 8 columns (A, B, D, O, S, SourceValue, CalcValue, BaseValue)
    ReDim extendedData(1 To selectedRows.Count, 1 To 8)
    
    ' Fill result array with data from selected rows
    For rowCount = 1 To selectedRows.Count
        i = selectedRows(rowCount)
        extendedData(rowCount, 1) = ws.Cells(i, 1).Value   ' Column A
        extendedData(rowCount, 2) = ws.Cells(i, 2).Value   ' Column B
        extendedData(rowCount, 3) = ws.Cells(i, 4).Value   ' Column D
        extendedData(rowCount, 4) = ws.Cells(i, 15).Value  ' Column O
        extendedData(rowCount, 5) = ws.Cells(i, 19).Value  ' Column S
        
        ' Config-specific columns
        extendedData(rowCount, 6) = ws.Cells(i, cfg.SourceValueCol).Value  ' Source/War value
        extendedData(rowCount, 7) = ws.Cells(i, cfg.BaseValueCol).Value    ' Base for calculation
        extendedData(rowCount, 8) = ws.Cells(i, cfg.BaseValueCol).Value    ' Base display value
    Next rowCount
    
    indx_GetSelectedDataExtended = extendedData
End Function

' -----------------------------------------------------------------------------
' indx_Engine_RemoveExistingSheet
' Remove existing Indexation sheet if it exists
' -----------------------------------------------------------------------------
Private Sub indx_Engine_RemoveExistingSheet()
    Dim ws As Worksheet
    
    On Error Resume Next
    Set ws = Worksheets("Indexation")
    If Not ws Is Nothing Then
        Application.DisplayAlerts = False
        ws.Delete
        Application.DisplayAlerts = True
    End If
    On Error GoTo 0
End Sub

' -----------------------------------------------------------------------------
' indx_Engine_CreateSheet
' Create Indexation sheet with data using configuration
' -----------------------------------------------------------------------------
Private Sub indx_Engine_CreateSheet(dataArray As Variant, cfg As IndxConfig)
    Dim wsKartei As Worksheet
    Dim wsIndexation As Worksheet
    
    Set wsKartei = Worksheets("Kartei")
    
    ' Create new Indexation sheet right after Kartei
    Set wsIndexation = Worksheets.Add(After:=wsKartei)
    wsIndexation.Name = "Indexation"
    
    ' Setup header row
    Call indx_Engine_SetupHeader(wsIndexation, cfg)
    
    ' Fill data
    Call indx_Engine_FillData(wsIndexation, dataArray, cfg)
    
    ' Apply formatting
    Call indx_Engine_ApplyFormatting(wsIndexation, UBound(dataArray, 1), cfg)
    
    ' Optimize column widths
    Call indx_SetOptimalColumnWidths(wsIndexation, 1, 8)
    
    ' Freeze first row
    wsIndexation.Activate
    wsIndexation.Rows("2:2").Select
    ActiveWindow.FreezePanes = True
    
    ' Select first data cell
    wsIndexation.Cells(2, 1).Select
End Sub

' -----------------------------------------------------------------------------
' indx_Engine_SetupHeader
' Setup header row with config-specific labels
' -----------------------------------------------------------------------------
Private Sub indx_Engine_SetupHeader(ws As Worksheet, cfg As IndxConfig)
    With ws
        .Cells(1, 1).Value = "A"
        .Cells(1, 2).Value = "B"
        .Cells(1, 3).Value = "D"
        .Cells(1, 4).Value = "O"
        .Cells(1, 5).Value = "S"
        .Cells(1, 6).Value = cfg.HeaderWar
        .Cells(1, 7).Value = cfg.HeaderWird
        .Cells(1, 8).Value = cfg.HeaderBase  ' May be empty for December
    End With
End Sub

' -----------------------------------------------------------------------------
' indx_Engine_FillData
' Fill data from selected records into Indexation sheet using configuration
' -----------------------------------------------------------------------------
Private Sub indx_Engine_FillData(ws As Worksheet, dataArray As Variant, cfg As IndxConfig)
    Dim i As Long
    Dim calcValue As Double
    Dim baseValue As Double
    
    For i = 1 To UBound(dataArray, 1)
        ' Fill columns A through E with identifier data
        ws.Cells(i + 1, 1).Value = dataArray(i, 1) ' A
        ws.Cells(i + 1, 2).Value = dataArray(i, 2) ' B
        ws.Cells(i + 1, 3).Value = dataArray(i, 3) ' D
        ws.Cells(i + 1, 4).Value = dataArray(i, 4) ' O
        ws.Cells(i + 1, 5).Value = dataArray(i, 5) ' S
        
        ' Fill column F with source/War value
        ws.Cells(i + 1, 6).Value = dataArray(i, 6)
        
        ' Calculate and fill column G (Wird)
        baseValue = indx_SafeNumericValue(dataArray(i, 7), 0)
        calcValue = Round(cfg.Coefficient * baseValue, 2)
        ws.Cells(i + 1, 7).Value = calcValue
        
        ' Fill column H with base value only if configured
        If cfg.FillBaseColumn Then
            ws.Cells(i + 1, 8).Value = dataArray(i, 8)
        End If
    Next i
End Sub

' -----------------------------------------------------------------------------
' indx_Engine_ApplyFormatting
' Apply formatting to header and data cells
' -----------------------------------------------------------------------------
Private Sub indx_Engine_ApplyFormatting(ws As Worksheet, dataRows As Long, cfg As IndxConfig)
    Dim headerRange As Range
    Dim dataRange As Range
    Dim headerEndCol As Long
    
    ' Determine header end column based on whether we have base column
    If cfg.FillBaseColumn Then
        headerEndCol = 8
    Else
        headerEndCol = 7
    End If
    
    ' Format header row
    Set headerRange = ws.Range(ws.Cells(1, 1), ws.Cells(1, headerEndCol))
    With headerRange
        .Interior.Color = RGB(192, 192, 192) ' Light gray background
        .Font.Bold = True
    End With
    
    ' Format data columns
    If dataRows > 0 Then
        ' Format column F (green background) for War values
        Set dataRange = ws.Range("F2:F" & (dataRows + 1))
        dataRange.Interior.Color = RGB(144, 238, 144) ' Light green
        
        ' Format column G (pink background, bold text) for Wird values
        Set dataRange = ws.Range("G2:G" & (dataRows + 1))
        With dataRange
            .Interior.Color = RGB(255, 182, 193) ' Light pink
            .Font.Bold = True
        End With
    End If
End Sub

' -----------------------------------------------------------------------------
' indx_Engine_CreateButton
' Create OK button with reference to the indexation type
' -----------------------------------------------------------------------------
Private Sub indx_Engine_CreateButton(cfg As IndxConfig)
    Dim ws As Worksheet
    Dim lastRow As Long
    Dim btnTop As Double
    Dim btnLeft As Double
    Dim btn As Button
    
    Set ws = Worksheets("Indexation")
    
    ' Find last row with data
    lastRow = ws.Cells(ws.Rows.Count, 1).End(xlUp).Row
    
    ' Position button below data with some spacing
    btnTop = ws.Cells(lastRow + 2, 1).Top
    btnLeft = ws.Cells(lastRow + 2, 4).Left
    
    ' Create button
    Set btn = ws.Buttons.Add(btnLeft, btnTop, 80, 25)
    
    ' Configure button properties
    ' Use different action based on month type
    With btn
        .Caption = "OK"
        .Name = "btnApplyChanges"
        
        ' Set appropriate handler based on configuration
        Select Case cfg.MonthName
            Case "August"
                .OnAction = "indx_ApplyChangesAugust"
            Case "December"
                .OnAction = "indx_ApplyChangesDecember"
            Case Else
                .OnAction = "indx_ApplyChangesAugust" ' Default to August
        End Select
    End With
End Sub

' -----------------------------------------------------------------------------
' indx_ApplyChangesWithConfig
' Generic apply changes procedure using configuration
' -----------------------------------------------------------------------------
Public Sub indx_ApplyChangesWithConfig(cfg As IndxConfig)
    Dim response As VbMsgBoxResult
    Dim originalScreenUpdating As Boolean
    Dim originalCalculation As Long
    
    ' Confirm action with user
    response = MsgBox("Moechten Sie die " & cfg.MonthName & "-Aenderungen wirklich anwenden? " & _
                     vbCrLf & "Alle Werte in der Zielspalte werden ersetzt.", _
                     vbYesNo + vbQuestion, "Bestaetigung")
    
    If response = vbNo Then
        Exit Sub
    End If
    
    ' Store original settings
    originalScreenUpdating = Application.ScreenUpdating
    originalCalculation = Application.Calculation
    
    Application.ScreenUpdating = False
    Application.Calculation = xlCalculationManual
    
    On Error GoTo ErrorHandler
    
    ' Apply changes to Kartei
    Call indx_Engine_UpdateKarteiData(cfg)
    
    ' Clean up - delete Indexation sheet
    Call indx_Engine_DeleteIndexationSheet
    
    ' Return to Kartei sheet
    Worksheets("Kartei").Activate
    
    ' Show completion message
    MsgBox cfg.MonthName & "-Aenderungen wurden erfolgreich angewendet.", _
           vbInformation, "Fertig"
    
    ' Fall through to Cleanup
    
Cleanup:
    ' Restore original settings (guaranteed execution)
    Application.ScreenUpdating = originalScreenUpdating
    Application.Calculation = originalCalculation
    Exit Sub
    
ErrorHandler:
    MsgBox "Fehler beim Anwenden der " & cfg.MonthName & "-Aenderungen: " & Err.Description, _
           vbCritical, "Fehler"
    Resume Cleanup
End Sub

' -----------------------------------------------------------------------------
' indx_Engine_DeleteIndexationSheet
' Safely delete Indexation sheet with alerts suppressed
' -----------------------------------------------------------------------------
Private Sub indx_Engine_DeleteIndexationSheet()
    On Error Resume Next
    Application.DisplayAlerts = False
    Worksheets("Indexation").Delete
    Application.DisplayAlerts = True
    On Error GoTo 0
End Sub

' -----------------------------------------------------------------------------
' indx_Engine_UpdateKarteiData
' Update target column in Kartei with values from Indexation sheet column G
' -----------------------------------------------------------------------------
Private Sub indx_Engine_UpdateKarteiData(cfg As IndxConfig)
    Dim wsKartei As Worksheet
    Dim wsIndexation As Worksheet
    Dim lastRow As Long
    Dim i As Long
    Dim karteiRow As Long
    Dim searchValues(1 To 4) As Variant
    Dim newValue As Variant
    
    Set wsKartei = Worksheets("Kartei")
    Set wsIndexation = Worksheets("Indexation")
    
    ' Find last row in Indexation sheet
    lastRow = wsIndexation.Cells(wsIndexation.Rows.Count, 1).End(xlUp).Row
    
    ' Loop through each data row in Indexation sheet
    For i = 2 To lastRow ' Start from row 2 (skip header)
        ' Get identifier values (A, B, D, O) from Indexation sheet
        searchValues(1) = wsIndexation.Cells(i, 1).Value ' Column A
        searchValues(2) = wsIndexation.Cells(i, 2).Value ' Column B
        searchValues(3) = wsIndexation.Cells(i, 3).Value ' Column D (from Kartei)
        searchValues(4) = wsIndexation.Cells(i, 4).Value ' Column O (from Kartei)
        
        ' Get new value from column G (Wird)
        newValue = wsIndexation.Cells(i, 7).Value
        
        ' Find corresponding row in Kartei sheet
        karteiRow = indx_Engine_FindKarteiRow(wsKartei, searchValues)
        
        If karteiRow > 0 Then
            ' Update target column with new value
            wsKartei.Cells(karteiRow, cfg.TargetWriteCol).Value = newValue
        End If
    Next i
End Sub

' -----------------------------------------------------------------------------
' indx_Engine_FindKarteiRow
' Find row in Kartei sheet that matches identifier values (A, B, D, O)
' -----------------------------------------------------------------------------
Private Function indx_Engine_FindKarteiRow(ws As Worksheet, searchValues() As Variant) As Long
    Dim lastRow As Long
    Dim i As Long
    Dim isMatch As Boolean
    
    lastRow = indx_FindLastDataRow(ws)
    
    ' Search through data rows (starting from row 3)
    For i = 3 To lastRow
        isMatch = True
        
        ' Check if all four identifier columns match
        If ws.Cells(i, 1).Value <> searchValues(1) Then isMatch = False ' Column A
        If ws.Cells(i, 2).Value <> searchValues(2) Then isMatch = False ' Column B
        If ws.Cells(i, 4).Value <> searchValues(3) Then isMatch = False ' Column D
        If ws.Cells(i, 15).Value <> searchValues(4) Then isMatch = False ' Column O
        
        If isMatch Then
            indx_Engine_FindKarteiRow = i
            Exit Function
        End If
    Next i
    
    ' If no match found
    indx_Engine_FindKarteiRow = 0
End Function

' -----------------------------------------------------------------------------
' indx_ApplyChangesAugust
' Button handler for August indexation
' -----------------------------------------------------------------------------
Public Sub indx_ApplyChangesAugust()
    Dim cfg As IndxConfig
    cfg = indx_CreateAugustConfig()
    Call indx_ApplyChangesWithConfig(cfg)
End Sub

' -----------------------------------------------------------------------------
' indx_ApplyChangesDecember
' Button handler for December indexation
' -----------------------------------------------------------------------------
Public Sub indx_ApplyChangesDecember()
    Dim cfg As IndxConfig
    cfg = indx_CreateDecemberConfig()
    Call indx_ApplyChangesWithConfig(cfg)
End Sub
