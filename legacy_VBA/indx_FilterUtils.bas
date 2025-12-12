Attribute VB_Name = "indx_FilterUtils"
Option Explicit

Sub indx_ClearActiveFilters()
    ' Clear active filters on Kartei sheet while preserving AutoFilter arrows
    ' This procedure only removes applied filters, not the AutoFilter itself
    
    Dim ws As Worksheet
    Set ws = Worksheets("Kartei")
    
    ' Check if AutoFilter is enabled
    If ws.AutoFilterMode Then
        ' Check if any filters are currently applied
        If indx_HasActiveFilters(ws) Then
            ' Clear all filters but keep AutoFilter arrows
            ws.AutoFilter.ShowAllData
        End If
    End If
End Sub

Private Function indx_HasActiveFilters(ws As Worksheet) As Boolean
    ' Check if there are any active filters applied
    ' Returns True if any filter criteria are currently applied
    
    Dim i As Integer
    indx_HasActiveFilters = False
    
    If ws.AutoFilterMode And Not ws.AutoFilter Is Nothing Then
        ' Check each filter field for applied criteria
        For i = 1 To ws.AutoFilter.Filters.count
            If ws.AutoFilter.Filters(i).On Then
                indx_HasActiveFilters = True
                Exit Function
            End If
        Next i
    End If
End Function
