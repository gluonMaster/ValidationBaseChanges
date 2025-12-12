Attribute VB_Name = "indx_ButtonHandler"
Option Explicit

' =============================================================================
' indx_ButtonHandler.bas
' Button creation and legacy handlers for Indexation sheet
' New handlers (indx_ApplyChangesAugust, indx_ApplyChangesDecember)
' are defined in indx_IndexationEngine.bas
' =============================================================================

Sub indx_CreateOKButton()
    ' Create OK button at the bottom of data in Indexation sheet
    ' NOTE: This is the legacy version for backward compatibility
    ' New code should use indx_Engine_CreateButton in indx_IndexationEngine
    
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
    btnLeft = ws.Cells(lastRow + 2, 4).Left ' Position around column D
    
    ' Create button
    Set btn = ws.Buttons.Add(btnLeft, btnTop, 80, 25)
    
    ' Configure button properties
    With btn
        .Caption = "OK"
        .Name = "btnApplyChanges"
        .OnAction = "indx_ApplyChanges"
    End With
End Sub

Sub indx_ApplyChanges()
    ' Legacy apply changes - redirects to August handler
    ' This procedure is kept for backward compatibility
    ' New code uses indx_ApplyChangesAugust or indx_ApplyChangesDecember
    ' defined in indx_IndexationEngine.bas
    
    Call indx_ApplyChangesAugust
End Sub

Private Sub indx_UpdateKarteiData()
    ' Legacy procedure - kept for reference
    ' New implementation is indx_Engine_UpdateKarteiData in indx_IndexationEngine
    ' Update AB column in Kartei sheet with values from G column in Indexation sheet
    
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
        
        ' Get new value from column G
        newValue = wsIndexation.Cells(i, 7).Value
        
        ' Find corresponding row in Kartei sheet
        karteiRow = indx_FindMatchingKarteiRow(wsKartei, searchValues)
        
        If karteiRow > 0 Then
            ' Update AB column (column 28) with new value
            wsKartei.Cells(karteiRow, 28).Value = newValue
        End If
    Next i
End Sub

Private Function indx_FindMatchingKarteiRow(ws As Worksheet, searchValues() As Variant) As Long
    ' Find row in Kartei sheet that matches the identifier values (A, B, D, O)
    ' NOTE: This is legacy code, new implementation is indx_Engine_FindKarteiRow
    
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
            indx_FindMatchingKarteiRow = i
            Exit Function
        End If
    Next i
    
    ' If no match found
    indx_FindMatchingKarteiRow = 0
End Function

