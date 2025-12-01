Attribute VB_Name = "indx_Utils"
Option Explicit

Function indx_ContainsText(sourceText As String, searchText As String) As Boolean
    ' Case-insensitive text search function
    ' Returns True if searchText is found within sourceText
    
    Dim normalizedSource As String
    Dim normalizedSearch As String
    
    ' Normalize text to uppercase and trim whitespace
    normalizedSource = UCase(Trim(CStr(sourceText)))
    normalizedSearch = UCase(Trim(searchText))
    
    ' Check if search text is contained in source text
    indx_ContainsText = (InStr(normalizedSource, normalizedSearch) > 0)
End Function

Function indx_NotContainsText(sourceText As String, searchText As String) As Boolean
    ' Case-insensitive text search function (negative)
    ' Returns True if searchText is NOT found within sourceText
    
    indx_NotContainsText = Not indx_ContainsText(sourceText, searchText)
End Function

Function indx_IsValidNumeric(value As Variant) As Boolean
    ' Check if value can be safely converted to a number
    ' Returns True if value is numeric or can be converted to numeric
    
    If IsEmpty(value) Or IsNull(value) Then
        indx_IsValidNumeric = False
    ElseIf IsNumeric(value) Then
        indx_IsValidNumeric = True
    Else
        indx_IsValidNumeric = False
    End If
End Function

Function indx_SafeNumericValue(value As Variant, defaultValue As Double) As Double
    ' Safely convert value to numeric, return default if conversion fails
    
    If indx_IsValidNumeric(value) Then
        indx_SafeNumericValue = CDbl(value)
    Else
        indx_SafeNumericValue = defaultValue
    End If
End Function

Sub indx_SetOptimalColumnWidths(ws As Worksheet, startCol As Integer, endCol As Integer)
    ' Auto-fit column widths for better readability
    
    Dim i As Integer
    
    For i = startCol To endCol
        ws.Columns(i).AutoFit
        
        ' Set maximum width to prevent extremely wide columns
        If ws.Columns(i).ColumnWidth > 25 Then
            ws.Columns(i).ColumnWidth = 25
        End If
        
        ' Set minimum width for readability
        If ws.Columns(i).ColumnWidth < 8 Then
            ws.Columns(i).ColumnWidth = 8
        End If
    Next i
End Sub

Function indx_GetSafeSheetReference(sheetName As String) As Worksheet
    ' Safely get worksheet reference with error handling
    
    Dim ws As Worksheet
    
    On Error Resume Next
    Set ws = Worksheets(sheetName)
    On Error GoTo 0
    
    If ws Is Nothing Then
        Err.Raise 9, "indx_Utils", "Arbeitsblatt '" & sheetName & "' wurde nicht gefunden."
    End If
    
    Set indx_GetSafeSheetReference = ws
End Function

Function indx_FindLastDataRow(ws As Worksheet) As Long
    ' Find last row containing data based on columns A or B being non-empty
    ' This function is used across multiple modules
    
    Dim lastRowA As Long, lastRowB As Long
    
    lastRowA = ws.Cells(ws.Rows.count, 1).End(xlUp).row
    lastRowB = ws.Cells(ws.Rows.count, 2).End(xlUp).row
    
    ' Return the maximum of both
    indx_FindLastDataRow = Application.WorksheetFunction.Max(lastRowA, lastRowB)
End Function
