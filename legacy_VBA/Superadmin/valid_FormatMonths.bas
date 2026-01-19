Attribute VB_Name = "valid_FormatMonths"
'==========================
'   Module: valid_FormatMonths
'   Purpose: Format monthly columns (U-AF, 21-32) in Superadmin Kartei
'   Ensures numeric values with proper decimal separators and formatting
'   Adapted from alt/FormatCellsReal.bas for Superadmin self-containment
'
'   MULTI-YEAR SUPPORT (2024, 2025, 2026):
'   FormatMonthlyColumnsForSheet(ws) - Works on any Kartei sheet
'   FormatMonthlyColumns - Legacy (defaults to "Kartei" sheet)
'==========================

Option Explicit

' ============================================================
' MULTI-YEAR API
' ============================================================

' Format monthly columns U-AF (21-32) for a specific worksheet
' Safe to call on empty ranges - will not error if no data present
' @param ws - The worksheet to format (e.g., Kartei24, Kartei25, Kartei26)
Public Sub FormatMonthlyColumnsForSheet(ByVal ws As Worksheet)
    On Error GoTo ErrorHandler
    
    If ws Is Nothing Then
        Exit Sub
    End If
    
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
        Call FormatSingleMonthColumnOnSheet(ws, colIndex, lastRow, decimalSeparator)
    Next colIndex
    
    ' Re-enable updates
    Application.Calculation = xlCalculationAutomatic
    Application.ScreenUpdating = True
    Exit Sub
    
ErrorHandler:
    Application.Calculation = xlCalculationAutomatic
    Application.ScreenUpdating = True
    ' Silent error handling - formatting is not critical to data integrity
    Debug.Print "Error in FormatMonthlyColumnsForSheet: " & Err.Description
End Sub

' ============================================================
' LEGACY API - Backward Compatibility
' ============================================================

' Format monthly columns U-AF (21-32) after loading pending records (legacy - "Kartei" sheet)
' For new code, use FormatMonthlyColumnsForSheet(ws) instead.
' Safe to call on empty ranges - will not error if no data present
Public Sub FormatMonthlyColumns()
    On Error Resume Next
    
    ' Try to find legacy "Kartei" sheet first, then fall back to Kartei25
    Dim ws As Worksheet
    Set ws = ThisWorkbook.Worksheets("Kartei")
    
    If ws Is Nothing Then
        Set ws = ThisWorkbook.Worksheets("Kartei25")
    End If
    
    If ws Is Nothing Then
        Exit Sub
    End If
    
    On Error GoTo 0
    
    ' Delegate to sheet-parameterized version
    FormatMonthlyColumnsForSheet ws
End Sub

' ============================================================
' INTERNAL HELPERS
' ============================================================

' Format a single monthly column on a specific sheet
Private Sub FormatSingleMonthColumnOnSheet(ByVal ws As Worksheet, _
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
