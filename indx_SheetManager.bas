Attribute VB_Name = "indx_SheetManager"
Option Explicit

' =============================================================================
' indx_SheetManager.bas
' Legacy sheet management for Indexation
' New code should use indx_IndexationEngine.bas procedures
' These procedures are kept for backward compatibility
' =============================================================================

Sub indx_CreateIndexationSheet(dataArray As Variant)
    ' Legacy procedure - creates August-style Indexation sheet
    ' For new code, use indx_Engine_CreateSheet in indx_IndexationEngine
    
    Dim wsKartei As Worksheet
    Dim wsIndexation As Worksheet
    Dim karteiIndex As Integer
    
    Set wsKartei = Worksheets("Kartei")
    
    ' Find Kartei sheet index
    karteiIndex = wsKartei.Index
    
    ' Create new Indexation sheet right after Kartei
    Set wsIndexation = Worksheets.Add(After:=wsKartei)
    wsIndexation.Name = "Indexation"
    
    ' Setup header row
    Call indx_SetupHeader(wsIndexation)
    
    ' Fill data
    Call indx_FillData(wsIndexation, dataArray)
    
    ' Apply formatting
    Call indx_ApplyFormatting(wsIndexation, UBound(dataArray, 1))
    
    ' Optimize column widths
    Call indx_SetOptimalColumnWidths(wsIndexation, 1, 8)
    
    ' Freeze first row
    wsIndexation.Activate
    wsIndexation.Rows("2:2").Select
    ActiveWindow.FreezePanes = True
    
    ' Select first data cell
    wsIndexation.Cells(2, 1).Select
End Sub

Private Sub indx_SetupHeader(ws As Worksheet)
    ' Setup header row with August-specific labels (legacy)
    
    With ws
        .Cells(1, 1).Value = "A"
        .Cells(1, 2).Value = "B"
        .Cells(1, 3).Value = "D"
        .Cells(1, 4).Value = "O"
        .Cells(1, 5).Value = "S"
        .Cells(1, 6).Value = "Aug War"
        .Cells(1, 7).Value = "Aug Wird"
        .Cells(1, 8).Value = "Sept"
    End With
End Sub

Private Sub indx_FillData(ws As Worksheet, dataArray As Variant)
    ' Fill data from selected records into Indexation sheet (legacy August version)
    ' Mapping: A->A, B->B, D->C, O->D, S->E, AB->F, AC->H
    
    Dim i As Long
    Dim calcValue As Double
    
    For i = 1 To UBound(dataArray, 1)
        ' Fill columns A through E with original data
        ws.Cells(i + 1, 1).Value = dataArray(i, 1) ' A -> A
        ws.Cells(i + 1, 2).Value = dataArray(i, 2) ' B -> B
        ws.Cells(i + 1, 3).Value = dataArray(i, 3) ' D -> C
        ws.Cells(i + 1, 4).Value = dataArray(i, 4) ' O -> D
        ws.Cells(i + 1, 5).Value = dataArray(i, 5) ' S -> E
        
        ' Fill column F with AB data
        ws.Cells(i + 1, 6).Value = dataArray(i, 6) ' AB -> F
        
        ' Calculate and fill column G (0.75 * AC value, rounded to 2 decimal places)
        calcValue = Round(0.75 * indx_SafeNumericValue(dataArray(i, 7), 0), 2)
        ws.Cells(i + 1, 7).Value = calcValue
        
        ' Fill column H with AC data
        ws.Cells(i + 1, 8).Value = dataArray(i, 7) ' AC -> H
    Next i
End Sub

Private Sub indx_ApplyFormatting(ws As Worksheet, dataRows As Long)
    ' Apply formatting to header and data cells
    
    Dim headerRange As Range
    Dim dataRange As Range
    
    ' Format header row (A1:H1)
    Set headerRange = ws.Range("A1:H1")
    With headerRange
        .Interior.Color = RGB(192, 192, 192) ' Light gray background
        .Font.Bold = True
    End With
    
    ' Format column F (green background) for data rows
    If dataRows > 0 Then
        Set dataRange = ws.Range("F2:F" & (dataRows + 1))
        dataRange.Interior.Color = RGB(144, 238, 144) ' Light green
        
        ' Format column G (pink background, bold text) for data rows
        Set dataRange = ws.Range("G2:G" & (dataRows + 1))
        With dataRange
            .Interior.Color = RGB(255, 182, 193) ' Light pink
            .Font.Bold = True
        End With
    End If
End Sub
