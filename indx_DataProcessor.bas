Attribute VB_Name = "indx_DataProcessor"
Option Explicit

' -----------------------------------------------------------------------------
' indx_GetSelectedRows
' Returns Collection of row numbers that meet selection criteria
' This is the SINGLE SOURCE OF TRUTH for row selection logic
' - Column S contains "O/V" (case insensitive)
' - Column T does NOT contain "KN" (case insensitive)
' -----------------------------------------------------------------------------
Public Function indx_GetSelectedRows() As Collection
    Dim ws As Worksheet
    Dim lastRow As Long
    Dim i As Long
    Dim selectedRows As Collection
    
    Set selectedRows = New Collection
    Set ws = Worksheets("Kartei")
    
    ' Find last row with data (starting from row 3)
    lastRow = indx_FindLastDataRow(ws)
    
    If lastRow < 3 Then
        Set indx_GetSelectedRows = selectedRows
        Exit Function
    End If
    
    ' Loop through data rows (starting from row 3)
    For i = 3 To lastRow
        If indx_MeetsSelectionCriteria(ws, i) Then
            selectedRows.Add i
        End If
    Next i
    
    Set indx_GetSelectedRows = selectedRows
End Function

Function indx_GetSelectedData() As Variant
    ' Extract selected records from Kartei sheet based on criteria:
    ' - Column S contains "O/V" (case insensitive)
    ' - Column T does NOT contain "KN" (case insensitive)
    ' Returns array with columns A, B, D, O, S, AB, AC
    ' Uses indx_GetSelectedRows as single source of truth for row selection
    
    Dim ws As Worksheet
    Dim selectedRows As Collection
    Dim resultArray As Variant
    Dim rowCount As Long
    Dim i As Long
    
    Set ws = Worksheets("Kartei")
    
    ' Get selected rows using single source of truth
    Set selectedRows = indx_GetSelectedRows()
    
    ' If no matching rows found
    If selectedRows.Count = 0 Then
        indx_GetSelectedData = Empty
        Exit Function
    End If
    
    ' Create result array: 7 columns (A, B, D, O, S, AB, AC) + number of selected rows
    ReDim resultArray(1 To selectedRows.Count, 1 To 7)
    
    ' Fill result array with data from selected rows
    For rowCount = 1 To selectedRows.Count
        i = selectedRows(rowCount)
        resultArray(rowCount, 1) = ws.Cells(i, 1).Value   ' Column A
        resultArray(rowCount, 2) = ws.Cells(i, 2).Value   ' Column B
        resultArray(rowCount, 3) = ws.Cells(i, 4).Value   ' Column D
        resultArray(rowCount, 4) = ws.Cells(i, 15).Value  ' Column O
        resultArray(rowCount, 5) = ws.Cells(i, 19).Value  ' Column S
        resultArray(rowCount, 6) = ws.Cells(i, 28).Value  ' Column AB
        resultArray(rowCount, 7) = ws.Cells(i, 29).Value  ' Column AC
    Next rowCount
    
    indx_GetSelectedData = resultArray
End Function

Private Function indx_MeetsSelectionCriteria(ws As Worksheet, rowNum As Long) As Boolean
    ' Check if row meets selection criteria:
    ' - Column S contains "O/V" (case insensitive)
    ' - Column T does NOT contain "KN" (case insensitive)
    
    Dim colS As String, colT As String
    
    ' Get cell values
    colS = CStr(ws.Cells(rowNum, 19).value) ' Column S
    colT = CStr(ws.Cells(rowNum, 20).value) ' Column T
    
    ' Check if S contains "O/V" using utility function
    Dim hasOV As Boolean
    hasOV = indx_ContainsText(colS, "O/V")
    
    ' Check if T does NOT contain "KN" using utility function
    Dim doesNotHaveKN As Boolean
    doesNotHaveKN = indx_NotContainsText(colT, "KN")
    
    ' Return True if both conditions are met
    indx_MeetsSelectionCriteria = hasOV And doesNotHaveKN
End Function
