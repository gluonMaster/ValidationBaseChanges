Attribute VB_Name = "NHblank_Diagnostics"
Option Explicit

' Comprehensive diagnostic check for the entire system
Public Sub RunFullDiagnostics()
    Dim msg As String
    Dim issues As Long
    
    issues = 0
    msg = "=== SYSTEM DIAGNOSTICS ===" & vbCrLf & vbCrLf
    
    ' 1. Check if Kinder worksheet exists
    msg = msg & "1. Worksheet 'Kinder':" & vbCrLf
    If CheckKinderWorksheet(msg) Then
        msg = msg & "   OK - Worksheet exists" & vbCrLf
    Else
        msg = msg & "   ERROR - Worksheet not found!" & vbCrLf
        issues = issues + 1
    End If
    msg = msg & vbCrLf
    
    ' 2. Check current position
    msg = msg & "2. Current Position:" & vbCrLf
    If ActiveSheet.name = "Kinder" Then
        msg = msg & "   OK - On 'Kinder' sheet" & vbCrLf
        msg = msg & "   Current row: " & ActiveCell.Row & vbCrLf
        msg = msg & "   Current column: " & ActiveCell.Column & vbCrLf
    Else
        msg = msg & "   WARNING - Not on 'Kinder' sheet!" & vbCrLf
        msg = msg & "   Current sheet: " & ActiveSheet.name & vbCrLf
        issues = issues + 1
    End If
    msg = msg & vbCrLf
    
    ' 3. Check active records
    msg = msg & "3. Active Records Analysis:" & vbCrLf
    CheckActiveRecords msg, issues
    msg = msg & vbCrLf
    
    ' 4. Check current row data
    If ActiveSheet.name = "Kinder" And ActiveCell.Row >= 5 Then
        msg = msg & "4. Current Row Data:" & vbCrLf
        CheckCurrentRowData msg, issues
        msg = msg & vbCrLf
    End If
    
    ' 5. Summary
    msg = msg & "=== SUMMARY ===" & vbCrLf
    If issues = 0 Then
        msg = msg & "No critical issues found." & vbCrLf
    Else
        msg = msg & "Found " & issues & " issue(s) - see details above." & vbCrLf
    End If
    
    MsgBox msg, vbInformation, "System Diagnostics Report"
End Sub

' Check if Kinder worksheet exists
Private Function CheckKinderWorksheet(ByRef msg As String) As Boolean
    Dim ws As Worksheet
    
    On Error Resume Next
    Set ws = ThisWorkbook.Worksheets("Kinder")
    On Error GoTo 0
    
    CheckKinderWorksheet = Not (ws Is Nothing)
End Function

' Check and count active records
Private Sub CheckActiveRecords(ByRef msg As String, ByRef issues As Long)
    Dim ws As Worksheet
    Dim lastRow As Long
    Dim i As Long
    Dim activeCount As Long
    Dim inactiveCount As Long
    Dim emptyCount As Long
    
    On Error Resume Next
    Set ws = ThisWorkbook.Worksheets("Kinder")
    On Error GoTo 0
    
    If ws Is Nothing Then
        msg = msg & "   Cannot check - worksheet not found" & vbCrLf
        Exit Sub
    End If
    
    lastRow = ws.Cells(ws.Rows.Count, "C").End(xlUp).Row
    msg = msg & "   Last row with data: " & lastRow & vbCrLf
    
    activeCount = 0
    inactiveCount = 0
    emptyCount = 0
    
    For i = 5 To lastRow
        If ws.Cells(i, "C").value = "" Then
            emptyCount = emptyCount + 1
        ElseIf ws.Cells(i, "C").Font.ColorIndex = 15 Then
            inactiveCount = inactiveCount + 1
        Else
            activeCount = activeCount + 1
        End If
    Next i
    
    msg = msg & "   Active records (ColorIndex <> 15): " & activeCount & vbCrLf
    msg = msg & "   Inactive records (ColorIndex = 15): " & inactiveCount & vbCrLf
    msg = msg & "   Empty records: " & emptyCount & vbCrLf
    
    If activeCount = 0 Then
        msg = msg & "   WARNING - No active records found!" & vbCrLf
        issues = issues + 1
    End If
End Sub

' Check current row data completeness
Private Sub CheckCurrentRowData(ByRef msg As String, ByRef issues As Long)
    Dim ws As Worksheet
    Dim rowNum As Long
    Dim lastName As String
    Dim firstName As String
    Dim discipline As String
    Dim dateFrom As Variant
    Dim dateTo As Variant
    Dim referenceDate As Variant
    Dim hasIssues As Boolean
    
    Set ws = ActiveSheet
    rowNum = ActiveCell.Row
    hasIssues = False
    
    lastName = Trim(ws.Cells(rowNum, "C").value)
    firstName = Trim(ws.Cells(rowNum, "D").value)
    discipline = Trim(ws.Cells(rowNum, "E").value)
    dateFrom = ws.Cells(rowNum, "G").value
    dateTo = ws.Cells(rowNum, "H").value
    referenceDate = ws.Range("T2").value
    
    msg = msg & "   Row: " & rowNum & vbCrLf
    msg = msg & "   ColorIndex: " & ws.Cells(rowNum, "C").Font.ColorIndex & vbCrLf
    
    If lastName = "" Then
        msg = msg & "   ERROR - Last name (column C) is empty!" & vbCrLf
        hasIssues = True
    Else
        msg = msg & "   Last name: " & lastName & vbCrLf
    End If
    
    If firstName = "" Then
        msg = msg & "   ERROR - First name (column D) is empty!" & vbCrLf
        hasIssues = True
    Else
        msg = msg & "   First name: " & firstName & vbCrLf
    End If
    
    If discipline = "" Then
        msg = msg & "   ERROR - Discipline (column E) is empty!" & vbCrLf
        hasIssues = True
    Else
        msg = msg & "   Discipline: " & discipline & vbCrLf
    End If
    
    ' Check dates
    If IsEmpty(dateFrom) Then
        msg = msg & "   Date From (G): Empty" & vbCrLf
    ElseIf Not IsDate(dateFrom) Then
        msg = msg & "   ERROR - Date From (G) is not a valid date: " & dateFrom & vbCrLf
        hasIssues = True
    Else
        msg = msg & "   Date From (G): " & Format(dateFrom, "dd.mm.yyyy") & vbCrLf
    End If
    
    If IsEmpty(dateTo) Then
        msg = msg & "   Date To (H): Empty" & vbCrLf
    ElseIf Not IsDate(dateTo) Then
        msg = msg & "   ERROR - Date To (H) is not a valid date: " & dateTo & vbCrLf
        hasIssues = True
    Else
        msg = msg & "   Date To (H): " & Format(dateTo, "dd.mm.yyyy") & vbCrLf
    End If
    
    If Not IsDate(referenceDate) Then
        msg = msg & "   ERROR - Reference Date (T2) is not a valid date: " & referenceDate & vbCrLf
        hasIssues = True
    Else
        msg = msg & "   Reference Date (T2): " & Format(referenceDate, "dd.mm.yyyy") & vbCrLf
    End If
    
    If hasIssues Then
        issues = issues + 1
    End If
End Sub

' Check if template file is valid
Public Function CheckTemplateFile(ByVal templatePath As String, ByRef errorMsg As String) As Boolean
    Dim wbTemplate As Workbook
    Dim wsTemplate As Worksheet
    
    On Error GoTo ErrorHandler
    
    ' Check if file exists
    If Dir(templatePath) = "" Then
        errorMsg = "Template file not found: " & templatePath
        CheckTemplateFile = False
        Exit Function
    End If
    
    ' Try to open template
    Set wbTemplate = Workbooks.Open(templatePath, ReadOnly:=True)
    
    ' Check if Muster worksheet exists
    On Error Resume Next
    Set wsTemplate = wbTemplate.Worksheets("Muster")
    On Error GoTo ErrorHandler
    
    If wsTemplate Is Nothing Then
        errorMsg = "Worksheet 'Muster' not found in template file!"
        wbTemplate.Close SaveChanges:=False
        CheckTemplateFile = False
        Exit Function
    End If
    
    ' All checks passed
    wbTemplate.Close SaveChanges:=False
    CheckTemplateFile = True
    Exit Function
    
ErrorHandler:
    errorMsg = "Error checking template: " & Err.Description
    If Not wbTemplate Is Nothing Then
        wbTemplate.Close SaveChanges:=False
    End If
    CheckTemplateFile = False
End Function

' Check if target folder is writable
Public Function CheckTargetFolder(ByVal folderPath As String, ByRef errorMsg As String) As Boolean
    Dim testFile As String
    Dim fso As Object
    
    On Error GoTo ErrorHandler
    
    ' Create FileSystemObject
    Set fso = CreateObject("Scripting.FileSystemObject")
    
    ' Check if folder exists
    If Not fso.FolderExists(folderPath) Then
        errorMsg = "Target folder does not exist: " & folderPath
        CheckTargetFolder = False
        Exit Function
    End If
    
    ' Try to create a test file to check write permissions
    testFile = folderPath & "\_test_write_" & Format(Now, "yyyymmddhhnnss") & ".tmp"
    
    Dim fileNum As Integer
    fileNum = FreeFile
    Open testFile For Output As #fileNum
    Print #fileNum, "test"
    Close #fileNum
    
    ' Delete test file
    Kill testFile
    
    CheckTargetFolder = True
    Exit Function
    
ErrorHandler:
    errorMsg = "Cannot write to target folder: " & Err.Description
    CheckTargetFolder = False
End Function

' Detailed check before processing
Public Function PreProcessCheck(ByVal templatePath As String, ByVal targetFolder As String, ByRef errorMsg As String) As Boolean
    Dim ws As Worksheet
    
    ' Check Kinder worksheet
    On Error Resume Next
    Set ws = ThisWorkbook.Worksheets("Kinder")
    On Error GoTo 0
    
    If ws Is Nothing Then
        errorMsg = "Worksheet 'Kinder' not found in current workbook!"
        PreProcessCheck = False
        Exit Function
    End If
    
    ' Check template file
    If Not CheckTemplateFile(templatePath, errorMsg) Then
        PreProcessCheck = False
        Exit Function
    End If
    
    ' Check target folder
    If Not CheckTargetFolder(targetFolder, errorMsg) Then
        PreProcessCheck = False
        Exit Function
    End If
    
    PreProcessCheck = True
End Function
