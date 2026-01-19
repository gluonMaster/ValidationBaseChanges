Attribute VB_Name = "pay_Utils"
'==========================
'   Module: pay_Utils
'   Purpose: Utility functions for family payment exports
'
'   Ported from Data-file/zahl_Utils.bas to Superadmin
'   Changes:
'   - Made sheet-agnostic (accepts worksheet reference)
'   - Integrates with valid_YearConfig for year handling
'
'   Used by: pay_Main, pay_DataProcessor, pay_FileGenerator
'==========================

Option Explicit

' ============================================================
' PERFORMANCE OPTIMIZATION
' ============================================================

Public Sub OptimizePerformance(restore As Boolean)
    ' Optimize or restore Excel performance settings
    If restore Then
        ' Restore normal settings
        Application.ScreenUpdating = True
        Application.Calculation = xlCalculationAutomatic
        Application.EnableEvents = True
        Application.DisplayAlerts = True
    Else
        ' Optimize for performance
        Application.ScreenUpdating = False
        Application.Calculation = xlCalculationManual
        Application.EnableEvents = False
        Application.DisplayAlerts = False
    End If
End Sub

' ============================================================
' ROW/COLUMN UTILITIES
' ============================================================

Public Function LetzteNr(ws As Worksheet, ErsteRow As Integer, Spalte As String) As Integer
    ' Find the last non-empty row in specified column
    ' Sheet-agnostic version - accepts worksheet reference
    Dim lastRow As Long
    
    On Error GoTo ErrorHandler
    
    ' Use Excel's built-in method to find last row (more efficient)
    lastRow = ws.Cells(ws.Rows.count, Spalte).End(xlUp).row
    
    ' Ensure we don't return a value less than ErsteRow
    If lastRow < ErsteRow Then
        LetzteNr = ErsteRow
    ElseIf lastRow > 10000 Then
        ' Safety check - limit to reasonable number
        LetzteNr = 10000
    Else
        LetzteNr = CInt(lastRow)
    End If
    
    Exit Function
    
ErrorHandler:
    ' Fallback to original method if Excel method fails
    LetzteNr = LetzteNrFallback(ws, ErsteRow, Spalte)
End Function

Private Function LetzteNrFallback(ws As Worksheet, ErsteRow As Integer, Spalte As String) As Integer
    ' Original method as fallback
    Dim i As Integer
    
    On Error GoTo ErrorHandler
    
    For i = ErsteRow To 10000
        If ws.Cells(i, Spalte) = "" Then
            LetzteNrFallback = i - 1
            Exit Function
        End If
    Next i
    
    ' If we reach here, all 10000 rows have data
    LetzteNrFallback = 10000
    Exit Function
    
ErrorHandler:
    ' Return minimum value if error occurs
    LetzteNrFallback = ErsteRow
End Function

' ============================================================
' MONTH NAME UTILITIES
' ============================================================

Public Function NumMonat(namm As String) As Integer
    ' Convert German month name to number
    Dim monthName As String
    monthName = UCase(Trim(namm))
    
    Select Case monthName
        Case "JANUAR"
            NumMonat = 1
        Case "FEBRUAR"
            NumMonat = 2
        Case "MAERZ", "MARZ", "MÄRZ"
            NumMonat = 3
        Case "APRIL"
            NumMonat = 4
        Case "MAI"
            NumMonat = 5
        Case "JUNI"
            NumMonat = 6
        Case "JULI"
            NumMonat = 7
        Case "AUGUST"
            NumMonat = 8
        Case "SEPTEMBER"
            NumMonat = 9
        Case "OKTOBER"
            NumMonat = 10
        Case "NOVEMBER"
            NumMonat = 11
        Case "DEZEMBER"
            NumMonat = 12
        Case Else
            NumMonat = 0 ' Invalid month name
    End Select
End Function

' ============================================================
' FILE SYSTEM UTILITIES
' ============================================================

Public Function CreateDirectory(path As String) As Boolean
    ' Create directory if it doesn't exist
    On Error GoTo ErrorHandler
    
    ' Check if directory already exists
    If Len(Dir(path, vbDirectory)) > 0 Then
        CreateDirectory = True
        Exit Function
    End If
    
    ' Try to create directory
    MkDir path
    CreateDirectory = True
    Exit Function
    
ErrorHandler:
    CreateDirectory = False
End Function

' ============================================================
' CELL VALUE UTILITIES
' ============================================================

Public Function SafeGetCellValue(ws As Worksheet, row As Integer, col As Variant) As String
    ' Safely get cell value with error handling
    On Error GoTo ErrorHandler
    
    Dim cellValue As Variant
    cellValue = ws.Cells(row, col).Value
    
    If IsNull(cellValue) Or IsEmpty(cellValue) Then
        SafeGetCellValue = ""
    Else
        SafeGetCellValue = CStr(cellValue)
    End If
    Exit Function
    
ErrorHandler:
    SafeGetCellValue = ""
End Function

Public Function SafeGetNumericValue(ws As Worksheet, row As Integer, col As Variant) As Double
    ' Safely get numeric cell value with error handling
    On Error GoTo ErrorHandler
    
    Dim cellValue As Variant
    cellValue = ws.Cells(row, col).Value
    
    If IsNumeric(cellValue) Then
        SafeGetNumericValue = CDbl(cellValue)
    Else
        SafeGetNumericValue = 0
    End If
    Exit Function
    
ErrorHandler:
    SafeGetNumericValue = 0
End Function

' ============================================================
' WORKSHEET UTILITIES
' ============================================================

Public Function ValidateWorksheet(wsName As String) As Boolean
    ' Check if worksheet exists in ThisWorkbook
    Dim ws As Worksheet
    
    On Error GoTo ErrorHandler
    Set ws = ThisWorkbook.Worksheets(wsName)
    ValidateWorksheet = True
    Exit Function
    
ErrorHandler:
    ValidateWorksheet = False
End Function

Public Function GetWorksheetSafe(wsName As String) As Worksheet
    ' Get worksheet reference, returns Nothing if not found
    On Error GoTo ErrorHandler
    Set GetWorksheetSafe = ThisWorkbook.Worksheets(wsName)
    Exit Function
    
ErrorHandler:
    Set GetWorksheetSafe = Nothing
End Function

' ============================================================
' LOGGING AND DISPLAY UTILITIES
' ============================================================

Public Sub LogMessage(message As String, Optional logType As String = "INFO")
    ' Simple logging function for debugging
    Debug.Print Format(Now, "yyyy-mm-dd hh:mm:ss") & " [" & logType & "] " & message
End Sub

Public Function FormatExecutionTime(seconds As Double) As String
    ' Format execution time for display
    If seconds < 1 Then
        FormatExecutionTime = Format(seconds * 1000, "0") & " ms"
    ElseIf seconds < 60 Then
        FormatExecutionTime = Format(seconds, "0.0") & " Sekunden"
    Else
        Dim minutes As Integer
        minutes = Int(seconds / 60)
        FormatExecutionTime = minutes & " Min " & Format(seconds - minutes * 60, "0") & " Sek"
    End If
End Function

' ============================================================
' COLOR AND FORMATTING UTILITIES
' ============================================================

Public Function GetColorName(colorValue As Integer) As String
    ' Convert color code to readable name for debugging
    Select Case colorValue
        Case 0
            GetColorName = "Normal"
        Case 2
            GetColorName = "Gelb (Konflikt)"
        Case Else
            GetColorName = "Unbekannt"
    End Select
End Function

' ============================================================
' YEAR SUFFIX UTILITIES
' ============================================================

Public Function GetYearSuffix(year2 As Integer) As String
    ' Get year suffix for filenames (e.g., "_24")
    GetYearSuffix = "_" & Format(year2, "00")
End Function

Public Function BuildYearFilename(baseName As String, year2 As Integer, extension As String) As String
    ' Build filename with year suffix
    ' e.g., BuildYearFilename("EltKosten", 24, ".xlsx") => "EltKosten_24.xlsx"
    BuildYearFilename = baseName & GetYearSuffix(year2) & extension
End Function

