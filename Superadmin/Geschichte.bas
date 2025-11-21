Attribute VB_Name = "Geschichte"
'==========================
'   Module: Geschichte
'   Purpose: Parse and display history for individual records in Superadmin file
'   Updated: Support for new history formats (Address, Subject1/2, Decl_n:)
'==========================

Option Explicit

Sub GeshichteMachen()
    Dim result As Collection
    Dim ws As Worksheet
    Dim currentRow As Long
    Dim operName As String
    Dim recordID As String
    
    On Error GoTo Cleanup

    ' Disable screen updating and automatic calculations for better performance
    Application.ScreenUpdating = False
    Application.Calculation = xlCalculationManual

    
    Set ws = ThisWorkbook.ActiveSheet
    currentRow = ActiveCell.Row
    
    Dim strGeshichte As String
    strGeshichte = ws.Cells(currentRow, 52).Value
    operName = ws.Cells(currentRow, 49).Value
    recordID = CStr(ws.Cells(currentRow, 48).Value) ' Get ID from column AV (48)
    
    If strGeshichte = "" Then
        MsgBox "There is no history of changes detected for this record (ID: " & recordID & ")", vbInformation
        GoTo Cleanup
    End If
    
    Set result = ParseHistory(strGeshichte)
    
    Dim i As Long
    For i = 1 To result.Count
        Dim evt As Object
        Set evt = result(i)
        
        Call CreateGeshichteSheet(i, evt("IsRuck"), evt("Reason"), CStr(evt("ChangeDate")), evt("Changes"), currentRow, operName, recordID)
        
        Debug.Print "Event #" & i & " (Record ID: " & recordID & ")"
        Debug.Print "  IsRuck: " & evt("IsRuck")
        Debug.Print "  Reason: " & Chr(34) & evt("Reason") & Chr(34)
        Debug.Print "  ChangeDate: " & evt("ChangeDate")
        
        Dim chDict As Object
        Set chDict = evt("Changes")
        
        Dim changeKey As Variant
        For Each changeKey In chDict.Keys
            Debug.Print "    Field: " & changeKey & _
                        "; War: " & chDict(changeKey)("War") & _
                        "; Ist: " & chDict(changeKey)("Ist")
        Next changeKey
        
        Debug.Print "---------------------------------------"
    Next i
    
Cleanup:
    ' Re-enable automatic calculations and screen updating
    Application.Calculation = xlCalculationAutomatic
    Application.ScreenUpdating = True
    If Err.Number <> 0 Then
        MsgBox "There is an error: " & Err.Description, vbCritical
    End If
End Sub

Public Sub CreateGeshichteSheet(ByVal j As Long, ByVal isRuck As Boolean, _
                                ByVal Reason As String, ByVal changeDate As String, _
                                ByVal Changes As Object, currentRow As Long, operName As String, _
                                recordID As String)
    Dim ws As Worksheet
    Dim sheetName As String
    sheetName = "Geschichte"
        
    On Error Resume Next
    Set ws = ThisWorkbook.Worksheets(sheetName)
    On Error GoTo 0
        
    ' Insert column headers
    If j = 1 Then
        ' If the sheet does not exist, create it
        If ws Is Nothing Then
            Set ws = ThisWorkbook.Worksheets.Add(After:=ThisWorkbook.Worksheets(ThisWorkbook.Worksheets.Count))
            ws.Name = sheetName
        Else
            ' If the sheet exists, clear its content, formats, and fills
            ws.Cells.Clear
        End If
    
        Dim headers As Variant
        headers = Array("ID", "Eltern", "Kind", "Gruppe I", "Gruppe II", "Jan", "Feb", "Mrz", "Apr", "Mai", "Jun", "Jul", "Aug", "Sep", "Okt", "Nov", "Dez", "Date", "Operator", "Comments")
        
        Dim i As Long
        For i = LBound(headers) To UBound(headers)
            With ws.Cells(1, i + 1) ' Start from column A (1)
                .Value = headers(i)
                .Font.Bold = True
                .HorizontalAlignment = xlCenter
            End With
        Next i
        
        ' Set the width of column T (comments) to 40 characters
        ws.Columns("T").ColumnWidth = 40
        
        ' Set column widths
        ws.Columns("A").ColumnWidth = 8
        ws.Columns("B").ColumnWidth = 20
        ws.Columns("C").ColumnWidth = 20
        ws.Columns("D").ColumnWidth = 22
        ws.Columns("E").ColumnWidth = 22
        ws.Columns("F:Q").ColumnWidth = 6
        ws.Columns("R").ColumnWidth = 12
        ws.Columns("S").ColumnWidth = 9
        
        ws.Activate

    End If
        
    ' Set the delimiter
    ws.Rows(3 * j + 1).RowHeight = ws.StandardHeight * 1 / 4
    ws.Range("A" & (3 * j + 1) & ":T" & (3 * j + 1)).Interior.Color = RGB(192, 192, 192)
    
    ' Format cell borders
    With ws.Range("A" & (2 * j + (j - 1)) & ":S" & (3 * j)).Borders
        .LineStyle = xlContinuous
        .Color = vbBlack
        .Weight = xlThin
    End With
        
    If isRuck Then
        ws.Range("T" & (3 * j)).Interior.Color = RGB(255, 153, 204)
        ws.Range("T" & (3 * j)).Value = Reason
    Else
        ws.Range("T" & (3 * j)).Interior.Color = RGB(204, 255, 153)
        ws.Range("T" & (3 * j)).Value = Reason
    End If
    
    ws.Range("S" & (3 * j)).Value = operName
    ws.Range("R" & (3 * j)).Value = changeDate
    ws.Range("A" & (3 * j)).Value = recordID ' Use passed ID instead of reading from Kartei
    ws.Range("B" & (3 * j)).Value = ThisWorkbook.Worksheets("Kartei").Cells(currentRow, 2).Value
    ws.Range("C" & (3 * j)).Value = ThisWorkbook.Worksheets("Kartei").Cells(currentRow, 4).Value
    ws.Range("D" & (3 * j)).Value = ThisWorkbook.Worksheets("Kartei").Cells(currentRow, 10).Value
    ws.Range("E" & (3 * j)).Value = ThisWorkbook.Worksheets("Kartei").Cells(currentRow, 15).Value
    
    ' Process changes - now supports both month numbers and field names (Address, Subject1, Subject2)
    Dim changeKey As Variant
    For Each changeKey In Changes.Keys
        Dim keyStr As String
        keyStr = CStr(changeKey)
        
        ' Check if it's a month number (1-12)
        If IsNumeric(keyStr) Then
            Dim monthNum As Long
            monthNum = CLng(keyStr)
            
            If monthNum >= 1 And monthNum <= 12 Then
                ' Fill month columns (F-Q = columns 6-17 for months 1-12)
                If IsEmpty(Changes(changeKey)("War")) Or IsNull(Changes(changeKey)("War")) Or Changes(changeKey)("War") = "" Then
                    ws.Cells((2 * j + (j - 1)), 5 + monthNum).Value = 0
                Else
                    ws.Cells((2 * j + (j - 1)), 5 + monthNum).Value = CDbl(Changes(changeKey)("War"))
                End If
                
                If IsEmpty(Changes(changeKey)("Ist")) Or IsNull(Changes(changeKey)("Ist")) Or Changes(changeKey)("Ist") = "" Then
                    ws.Cells((3 * j), 5 + monthNum).Value = 0
                Else
                    ws.Cells((3 * j), 5 + monthNum).Value = CDbl(Changes(changeKey)("Ist"))
                End If

                ws.Cells((3 * j), 5 + monthNum).Interior.Color = RGB(255, 192, 203)
            End If
        Else
            ' Non-numeric key - field names like "Address", "Subject1", "Subject2"
            ' Add to comments field with field name
            Dim fieldComment As String
            fieldComment = keyStr & ": Was(" & Changes(changeKey)("War") & ") -> Is(" & Changes(changeKey)("Ist") & ")"
            
            If ws.Range("T" & (3 * j)).Value <> "" Then
                ws.Range("T" & (3 * j)).Value = ws.Range("T" & (3 * j)).Value & vbCrLf & fieldComment
            Else
                ws.Range("T" & (3 * j)).Value = fieldComment
            End If
        End If
    Next changeKey
    
    With ws.Range("F" & (2 * j + (j - 1)) & ":Q" & (3 * j))
        .NumberFormat = "0.00"
    End With

    
End Sub

Sub ConvertAndFormatCells()
    Dim ws As Worksheet
    Dim lastRow As Long
    Dim i As Long
    Dim rng As Range, rngAY As Range
    Dim decimalSeparator As String
    
    ' Work on the active sheet
    Set ws = ActiveSheet
    
    ' Get the system decimal separator (either "," or ".")
    decimalSeparator = Application.International(xlDecimalSeparator)
    
    ' Find the last used row by checking column A (change if needed)
    lastRow = ws.Cells(ws.Rows.Count, 1).End(xlUp).Row
    If lastRow < 2 Then
        Exit Sub
    End If
    
    For i = 6 To 17
        Set rng = ws.Range(ws.Cells(2, i), ws.Cells(lastRow, i))
        
        ' Remove extra spaces
        rng.Replace What:=" ", Replacement:="", LookAt:=xlPart
        
        ' Unify decimal separators so Excel can parse them correctly
        If decimalSeparator = "," Then
            ' Replace any '.' with ',' in case the data was typed with a dot
            rng.Replace What:=".", Replacement:=",", LookAt:=xlPart
        Else
            ' Replace any ',' with '.' if system uses dot
            rng.Replace What:=",", Replacement:=".", LookAt:=xlPart
        End If
        
        ' Now do TextToColumns on one column range
        rng.TextToColumns Destination:=rng.Cells(1, 1), DataType:=xlDelimited, _
                          TextQualifier:=xlDoubleQuote, ConsecutiveDelimiter:=False, _
                          Tab:=False, Semicolon:=False, Comma:=False, Space:=False, _
                          Other:=False
        
        ' Finally, set numeric format with two decimals, no leading zeros
        rng.NumberFormat = "0.00"
    Next i
    
End Sub

' ==============================
' ParseHistory Function
' ==============================
' Parses history string from AZ column and returns collection of event objects
' History format:
'   - Events separated by " || "
'   - Each event can have "Ruck: " prefix for past month changes
'   - Month changes: "Mnt.N: War(X); Ist(Y). "
'   - Field changes: "FieldName: Was(X); Is(Y). " (Address, Subject1, Subject2)
'   - Decl comments: "Decl_N: comment"
'   - Date: dd.mm.yyyy or similar
'   - Comment: /text/
' ==============================
Public Function ParseHistory(ByVal historyStr As String) As Collection
    Dim result As New Collection
    
    If Trim(historyStr) = "" Then
        Set ParseHistory = result
        Exit Function
    End If
    
    ' Split by event separator " || "
    Dim events() As String
    events = Split(historyStr, " || ")
    
    Dim i As Long
    For i = LBound(events) To UBound(events)
        Dim eventStr As String
        eventStr = Trim(events(i))
        
        If eventStr <> "" Then
            Dim evt As Object
            Set evt = ParseSingleEvent(eventStr)
            
            If Not evt Is Nothing Then
                result.Add evt
            End If
        End If
    Next i
    
    Set ParseHistory = result
End Function

' Parse a single event from history string
Private Function ParseSingleEvent(ByVal eventStr As String) As Object
    Dim evt As Object
    Set evt = CreateObject("Scripting.Dictionary")
    
    Dim isRuck As Boolean
    isRuck = False
    
    Dim changeDate As String
    changeDate = ""
    
    Dim reason As String
    reason = ""
    
    Dim changesDict As Object
    Set changesDict = CreateObject("Scripting.Dictionary")
    
    ' Check for "Ruck: " prefix
    If Left(eventStr, 6) = "Ruck: " Then
        isRuck = True
        eventStr = Mid(eventStr, 7) ' Remove "Ruck: "
    End If
    
    ' Extract comment enclosed in /.../ 
    Dim commentStart As Long
    Dim commentEnd As Long
    commentStart = InStr(eventStr, "/")
    
    If commentStart > 0 Then
        commentEnd = InStr(commentStart + 1, eventStr, "/")
        If commentEnd > commentStart Then
            reason = Mid(eventStr, commentStart + 1, commentEnd - commentStart - 1)
        End If
    End If
    
    ' Extract date (dd.mm.yyyy or similar format)
    Dim datePattern As String
    Dim pos As Long
    
    ' Simple date extraction: look for pattern like "dd.mm.yyyy"
    ' We'll search for sequences that look like dates
    pos = 1
    Do While pos <= Len(eventStr)
        Dim char As String
        char = Mid(eventStr, pos, 1)
        
        If IsNumeric(char) Then
            ' Found start of potential date
            Dim dateCandidate As String
            Dim j As Long
            For j = 0 To 10 ' Max length for date
                If pos + j <= Len(eventStr) Then
                    Dim testChar As String
                    testChar = Mid(eventStr, pos + j, 1)
                    If IsNumeric(testChar) Or testChar = "." Or testChar = "/" Or testChar = "-" Then
                        dateCandidate = dateCandidate & testChar
                    Else
                        Exit For
                    End If
                End If
            Next j
            
            ' Check if this looks like a valid date
            On Error Resume Next
            If IsDate(dateCandidate) Then
                changeDate = dateCandidate
                Exit Do
            End If
            On Error GoTo 0
        End If
        
        pos = pos + 1
    Loop
    
    ' Parse change entries: "Mnt.N: War(...); Ist(...)." or "FieldName: Was(...); Is(...)."
    ' Pattern: "Mnt.1: War(123); Ist(456). "
    Dim changes As String
    changes = eventStr
    
    Dim regex As Object
    Set regex = CreateObject("VBScript.RegExp")
    regex.Global = True
    regex.IgnoreCase = True
    
    ' Pattern for month changes: Mnt.N: War(...); Ist(...).
    regex.Pattern = "Mnt\.(\d+):\s*War\(([^)]*)\);\s*Ist\(([^)]*)\)\."
    
    Dim matches As Object
    Set matches = regex.Execute(changes)
    
    Dim match As Object
    For Each match In matches
        If match.SubMatches.Count >= 3 Then
            Dim monthNum As String
            monthNum = match.SubMatches(0)
            
            Dim warValue As String
            warValue = match.SubMatches(1)
            
            Dim istValue As String
            istValue = match.SubMatches(2)
            
            Dim changeInfo As Object
            Set changeInfo = CreateObject("Scripting.Dictionary")
            changeInfo("War") = warValue
            changeInfo("Ist") = istValue
            
            If Not changesDict.Exists(monthNum) Then
                changesDict.Add monthNum, changeInfo
            End If
        End If
    Next match
    
    ' Pattern for field changes: FieldName: Was(...); Is(...).
    ' Supports: Address, Subject1, Subject2, etc.
    regex.Pattern = "(Address|Subject1|Subject2|Decl_\d+):\s*Was\(([^)]*)\);\s*Is\(([^)]*)\)\."
    
    Set matches = regex.Execute(changes)
    
    For Each match In matches
        If match.SubMatches.Count >= 3 Then
            Dim fieldName As String
            fieldName = match.SubMatches(0)
            
            warValue = match.SubMatches(1)
            istValue = match.SubMatches(2)
            
            Set changeInfo = CreateObject("Scripting.Dictionary")
            changeInfo("War") = warValue
            changeInfo("Ist") = istValue
            
            If Not changesDict.Exists(fieldName) Then
                changesDict.Add fieldName, changeInfo
            End If
        End If
    Next match
    
    ' Store parsed data in event dictionary
    evt("IsRuck") = isRuck
    evt("Reason") = reason
    evt("ChangeDate") = changeDate
    evt("Changes") = changesDict
    
    Set ParseSingleEvent = evt
End Function


