Attribute VB_Name = "valid_FormatMonths"
'==========================
'   Module: valid_FormatMonths
'   Purpose: Format monthly columns (U-AF, 21-32) in Superadmin Kartei
'   Ensures numeric values with proper decimal separators and formatting
'   Adapted from alt/FormatCellsReal.bas for Superadmin self-containment
'==========================

Option Explicit

' Format monthly columns U-AF (21-32) after loading pending records
' Safe to call on empty ranges - will not error if no data present
Public Sub FormatMonthlyColumns()
    On Error GoTo ErrorHandler
    
    Dim ws As Worksheet
    Set ws = ThisWorkbook.Worksheets("Kartei")
    
    ' Find last used row based on column A
    Dim lastRow As Long
    lastRow = ws.Cells(ws.Rows.Count, 1).End(xlUp).Row
    
    ' If no data rows (only headers or empty), exit safely
    If lastRow < 3 Then
        Exit Sub
    End If
    
    ' Disable updates for performance
    Application.ScreenUpdating = False
    Application.Calculation = xlCalculationManual
    
    ' Get system decimal separator
    Dim decimalSeparator As String
    decimalSeparator = Application.International(xlDecimalSeparator)
    
    ' Process each monthly column (U-AF = columns 21-32)
    Dim colIndex As Long
    For colIndex = 21 To 32
        Call FormatSingleMonthColumn(ws, colIndex, lastRow, decimalSeparator)
    Next colIndex
    
    ' Re-enable updates
    Application.Calculation = xlCalculationAutomatic
    Application.ScreenUpdating = True
    Exit Sub
    
ErrorHandler:
    Application.Calculation = xlCalculationAutomatic
    Application.ScreenUpdating = True
    ' Silent error handling - formatting is not critical to data integrity
    Debug.Print "Error in FormatMonthlyColumns: " & Err.Description
End Sub

' Format a single monthly column
Private Sub FormatSingleMonthColumn(ByVal ws As Worksheet, _
                                    ByVal colIndex As Long, _
                                    ByVal lastRow As Long, _
                                    ByVal decimalSeparator As String)
    On Error Resume Next
    
    Dim rng As Range
    Set rng = ws.Range(ws.Cells(3, colIndex), ws.Cells(lastRow, colIndex))
    
    ' Step 1: Remove extra spaces
    rng.Replace What:=" ", Replacement:="", LookAt:=xlPart
    
    ' Step 2: Normalize decimal separators to match system settings
    If decimalSeparator = "," Then
        ' System uses comma - convert any dots to commas
        rng.Replace What:=".", Replacement:=",", LookAt:=xlPart
    Else
        ' System uses dot - convert any commas to dots
        rng.Replace What:=",", Replacement:=".", LookAt:=xlPart
    End If
    
    ' Step 3: Convert text to columns (forces Excel to parse as numeric)
    rng.TextToColumns Destination:=rng.Cells(1, 1), _
                      DataType:=xlDelimited, _
                      TextQualifier:=xlDoubleQuote, _
                      ConsecutiveDelimiter:=False, _
                      Tab:=False, _
                      Semicolon:=False, _
                      Comma:=False, _
                      Space:=False, _
                      Other:=False
    
    ' Step 4: Set number format to two decimal places
    rng.NumberFormat = "0.00"
    
    On Error GoTo 0
End Sub
