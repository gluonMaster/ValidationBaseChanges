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
    
    ' Split history into segments to enrich parsed events with Address/Subject1/Subject2
    Dim segments() As String
    segments = Split(strGeshichte, "||")
    
    Set result = valid_ParseHistory.ParseHistory(strGeshichte)
    
    Dim i As Long
    For i = 1 To result.Count
        Dim evt As Object
        Set evt = result(i)
        
        ' Enrich Changes with Address/Subject1/Subject2 parsed from corresponding raw segment
        Dim segmentText As String
        segmentText = ""
        If (i - 1) >= LBound(segments) And (i - 1) <= UBound(segments) Then
            segmentText = Trim$(segments(i - 1))
        End If
        
        Dim fieldChanges As Object
        Set fieldChanges = ParseFieldChangesFromSegment(segmentText)
        Call MergeFieldChangesIntoChanges(evt("Changes"), fieldChanges)
        
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
            Dim warText As String
            Dim istText As String
            Dim fieldComment As String
            
            warText = Changes(changeKey)("War")
            istText = Changes(changeKey)("Ist")
            
            fieldComment = keyStr & ": Was(" & warText & "); Is(" & istText & ")."
            
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

' Parse Address/Subject1/Subject2 changes from a raw history segment
Private Function ParseFieldChangesFromSegment(ByVal segment As String) As Object
    Dim result As Object
    Set result = CreateObject("Scripting.Dictionary")
    
    If Len(Trim$(segment)) = 0 Then
        Set ParseFieldChangesFromSegment = result
        Exit Function
    End If
    
    Call AddFieldChangeToDict(segment, "Address", result)
    Call AddFieldChangeToDict(segment, "Subject1", result)
    Call AddFieldChangeToDict(segment, "Subject2", result)
    
    Set ParseFieldChangesFromSegment = result
End Function

' Merge additional field changes into the base Changes dictionary
Private Sub MergeFieldChangesIntoChanges(ByVal baseChanges As Object, ByVal fieldChanges As Object)
    If baseChanges Is Nothing Then Exit Sub
    If fieldChanges Is Nothing Then Exit Sub
    
    Dim key As Variant
    For Each key In fieldChanges.Keys
        If Not baseChanges.Exists(key) Then
            baseChanges.Add key, fieldChanges(key)
        Else
            baseChanges(key) = fieldChanges(key)
        End If
    Next key
End Sub

' Helper: extract single "FieldName: Was(...); Is(...)." block into dictionary
Private Sub AddFieldChangeToDict(ByVal segment As String, ByVal fieldName As String, ByVal dict As Object)
    Dim marker As String
    marker = fieldName & ": Was("
    
    Dim pos As Long
    pos = InStr(1, segment, marker, vbTextCompare)
    If pos = 0 Then Exit Sub
    
    Dim startWar As Long
    startWar = pos + Len(marker)
    
    Dim endWar As Long
    endWar = InStr(startWar, segment, ")", vbTextCompare)
    If endWar = 0 Then Exit Sub
    
    Dim warVal As String
    warVal = Mid$(segment, startWar, endWar - startWar)
    
    Dim markerIs As String
    markerIs = "); Is("
    
    Dim posIs As Long
    posIs = InStr(endWar, segment, markerIs, vbTextCompare)
    If posIs = 0 Then Exit Sub
    
    Dim startIst As Long
    startIst = posIs + Len(markerIs)
    
    Dim endIst As Long
    endIst = InStr(startIst, segment, ")", vbTextCompare)
    
    Dim istVal As String
    If endIst > startIst Then
        istVal = Mid$(segment, startIst, endIst - startIst)
    Else
        istVal = Mid$(segment, startIst)
    End If
    
    warVal = Trim$(warVal)
    istVal = Trim$(istVal)
    
    Dim fieldDict As Object
    Set fieldDict = CreateObject("Scripting.Dictionary")
    fieldDict.Add "War", warVal
    fieldDict.Add "Ist", istVal
    
    If dict.Exists(fieldName) Then
        dict(fieldName) = fieldDict
    Else
        dict.Add fieldName, fieldDict
    End If
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

