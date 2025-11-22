Attribute VB_Name = "Geschichte"
Sub GeshichteMachen()
    Dim result As Collection
    Dim ws As Worksheet
    Dim currentRow As Long
    Dim operName As String
    
    On Error GoTo Cleanup

    ' Disable screen updating and automatic calculations for better performance
    Application.ScreenUpdating = False
    Application.Calculation = xlCalculationManual

    
    Set ws = ThisWorkbook.ActiveSheet
    currentRow = ActiveCell.row
    
    Dim strGeshichte As String
    strGeshichte = ws.Cells(currentRow, 52).value
    operName = ws.Cells(currentRow, 49).value
    
    If strGeshichte = "" Then
        MsgBox ("There is no history of changes detected for this record")
        Exit Sub
    End If
    
    Set result = ParseHistory(strGeshichte)
    
    Dim i As Long
    For i = 1 To result.count
        Dim evt As Object
        Set evt = result(i)
        
        Call CreateGeshichteSheet(i, evt("IsRuck"), evt("Reason"), CStr(evt("ChangeDate")), evt("Changes"), currentRow, operName)
        
        Debug.Print "Event #" & i
        Debug.Print "  IsRuck: " & evt("IsRuck")
        Debug.Print "  Reason: " & Chr(34) & evt("Reason") & Chr(34)
        Debug.Print "  ChangeDate: " & evt("ChangeDate")
        
        Dim chDict As Object
        Set chDict = evt("Changes")
        
        Dim monthKey As Variant
        For Each monthKey In chDict.Keys
            Debug.Print "    Month: " & monthKey & _
                        "; War: " & chDict(monthKey)("War") & _
                        "; Ist: " & chDict(monthKey)("Ist")
        Next monthKey
        
        Debug.Print "---------------------------------------"
    Next i
    
    'Call ConvertAndFormatCells
    
    Exit Sub
Cleanup:
    ' Re-enable automatic calculations and screen updating
    Application.Calculation = xlCalculationAutomatic
    Application.ScreenUpdating = True
    MsgBox "There is an error: " & Err.Description
End Sub

Public Sub CreateGeshichteSheet(ByVal j As Long, ByVal isRuck As Boolean, _
                                ByVal Reason As String, ByVal changeDate As String, _
                                ByVal Changes As Object, currentRow As Long, operName As String)
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
            Set ws = ThisWorkbook.Worksheets.Add(After:=ThisWorkbook.Worksheets(ThisWorkbook.Worksheets.count))
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
                .value = headers(i)
                .Font.Bold = True
                .HorizontalAlignment = xlCenter
            End With
        Next i
        
        ' Set the width of column Q to 40 characters
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
        ws.Range("T" & (3 * j)).value = Reason
    Else
        ws.Range("T" & (3 * j)).Interior.Color = RGB(204, 255, 153)
        ws.Range("T" & (3 * j)).value = Reason
    End If
    
    ws.Range("S" & (3 * j)).value = operName
    ws.Range("R" & (3 * j)).value = changeDate
    ws.Range("A" & (3 * j)).value = ThisWorkbook.Worksheets("Kartei").Cells(currentRow, 1).value
    ws.Range("B" & (3 * j)).value = ThisWorkbook.Worksheets("Kartei").Cells(currentRow, 2).value
    ws.Range("C" & (3 * j)).value = ThisWorkbook.Worksheets("Kartei").Cells(currentRow, 4).value
    ws.Range("D" & (3 * j)).value = ThisWorkbook.Worksheets("Kartei").Cells(currentRow, 10).value
    ws.Range("E" & (3 * j)).value = ThisWorkbook.Worksheets("Kartei").Cells(currentRow, 15).value
    
    Dim monthKey As Variant
    For Each monthKey In Changes.Keys
        If IsEmpty(Changes(monthKey)("War")) Or IsNull(Changes(monthKey)("War")) Or Changes(monthKey)("War") = "" Then
            ws.Cells((2 * j + (j - 1)), 5 + monthKey).value = 0
        Else
            ws.Cells((2 * j + (j - 1)), 5 + monthKey).value = CDbl(Changes(monthKey)("War"))
        End If
        
        If IsEmpty(Changes(monthKey)("Ist")) Or IsNull(Changes(monthKey)("Ist")) Or Changes(monthKey)("Ist") = "" Then
            ws.Cells((3 * j), 5 + monthKey).value = 0
        Else
            ws.Cells((3 * j), 5 + monthKey).value = CDbl(Changes(monthKey)("Ist"))
        End If

        ws.Cells((3 * j), 5 + monthKey).Interior.Color = RGB(255, 192, 203)
    Next monthKey
    
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
    lastRow = ws.Cells(ws.Rows.count, 1).End(xlUp).row
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

