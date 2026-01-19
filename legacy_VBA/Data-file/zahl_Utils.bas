Attribute VB_Name = "zahl_Utils"
' Utility functions module for family payments
' Contains helper functions and performance optimization routines

Option Explicit

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

Public Function LetzteNr(TabName As String, ErsteRow As Integer, Spalte As String) As Integer
    ' Find the last non-empty row in specified column
    ' Enhanced version with better error handling
    Dim ws As Worksheet
    Dim lastRow As Long
    
    On Error GoTo ErrorHandler
    
    ' Get worksheet reference
    Set ws = Worksheets(TabName)
    
    ' Use Excel's built-in method to find last row (more efficient)
    lastRow = ws.Cells(ws.Rows.count, Spalte).End(xlUp).row
    
    ' Ensure we don't return a value less than ErsteRow
    If lastRow < ErsteRow Then
        LetzteNr = ErsteRow
    ElseIf lastRow > 10000 Then
        ' Safety check - limit to reasonable number
        LetzteNr = 10000
    Else
        LetzteNr = lastRow
    End If
    
    Exit Function
    
ErrorHandler:
    ' Fallback to original method if Excel method fails
    LetzteNr = LetzteNrFallback(TabName, ErsteRow, Spalte)
End Function

Private Function LetzteNrFallback(TabName As String, ErsteRow As Integer, Spalte As String) As Integer
    ' Original method as fallback
    Dim i As Integer
    
    On Error GoTo ErrorHandler
    
    For i = ErsteRow To 10000
        If Worksheets(TabName).Cells(i, Spalte) = "" Then
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

Public Function numMonat(namm As String) As Integer
    ' Convert German month name to number
    ' Enhanced with better string handling
    Dim monthName As String
    monthName = UCase(Trim(namm))
    
    Select Case monthName
        Case "JANUAR"
            numMonat = 1
        Case "FEBRUAR"
            numMonat = 2
        Case "MAERZ", "MARZ"
            numMonat = 3
        Case "APRIL"
            numMonat = 4
        Case "MAI"
            numMonat = 5
        Case "JUNI"
            numMonat = 6
        Case "JULI"
            numMonat = 7
        Case "AUGUST"
            numMonat = 8
        Case "SEPTEMBER"
            numMonat = 9
        Case "OKTOBER"
            numMonat = 10
        Case "NOVEMBER"
            numMonat = 11
        Case "DEZEMBER"
            numMonat = 12
        Case Else
            numMonat = 0 ' Invalid month name
    End Select
End Function

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

Public Function SafeGetCellValue(ws As Worksheet, row As Integer, col As Variant) As String
    ' Safely get cell value with error handling
    On Error GoTo ErrorHandler
    
    SafeGetCellValue = CStr(ws.Cells(row, col))
    Exit Function
    
ErrorHandler:
    SafeGetCellValue = ""
End Function

Public Function SafeGetNumericValue(ws As Worksheet, row As Integer, col As Variant) As Double
    ' Safely get numeric cell value with error handling
    On Error GoTo ErrorHandler
    
    Dim cellValue As Variant
    cellValue = ws.Cells(row, col)
    
    If IsNumeric(cellValue) Then
        SafeGetNumericValue = CDbl(cellValue)
    Else
        SafeGetNumericValue = 0
    End If
    Exit Function
    
ErrorHandler:
    SafeGetNumericValue = 0
End Function

Public Sub LogMessage(message As String, Optional logType As String = "INFO")
    ' Simple logging function for debugging (optional)
    ' Can be enhanced to write to file if needed
    Debug.Print Format(Now, "yyyy-mm-dd hh:mm:ss") & " [" & logType & "] " & message
End Sub

Public Function ValidateWorksheet(wsName As String) As Boolean
    ' Check if worksheet exists
    Dim ws As Worksheet
    
    On Error GoTo ErrorHandler
    Set ws = Worksheets(wsName)
    ValidateWorksheet = True
    Exit Function
    
ErrorHandler:
    ValidateWorksheet = False
End Function

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

Public Sub ClearArray(ByRef arr As Variant, arraySize As Integer)
    ' Generic function to clear arrays
    ' Note: This is a template - specific implementation depends on array type
    Dim i As Integer
    For i = 1 To arraySize
        ' Implementation depends on array type
        ' This would need to be customized for each array type
    Next i
End Sub

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

Public Function CreateBackupFilename(originalFilename As String) As String
    ' Create backup filename with timestamp
    Dim fileExtension As String
    Dim baseName As String
    Dim timestamp As String
    
    fileExtension = Right(originalFilename, 4) ' .xlsx
    baseName = Left(originalFilename, Len(originalFilename) - 4)
    timestamp = Format(Now, "yyyymmdd_hhmmss")
    
    CreateBackupFilename = baseName & "_backup_" & timestamp & fileExtension
End Function
