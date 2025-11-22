Attribute VB_Name = "grossGeschichte"
Sub GrossGeshichteMachen()
    Dim wsGross As Worksheet
    Dim wsKartei As Worksheet
    Dim startDate As Date
    Dim endDate As Date
    Dim lastRow As Long
    Dim currentRow As Long
    Dim strGeschichte As String
    Dim operName As String
    Dim result As Collection
    Dim outputRow As Long
    
    On Error GoTo Cleanup
    
    ' Disable screen updating and automatic calculations for better performance
    Application.ScreenUpdating = False
    Application.Calculation = xlCalculationManual
    
    ' Get worksheets
    Set wsGross = ThisWorkbook.Worksheets("GrossGeschichte")
    Set wsKartei = ThisWorkbook.Worksheets("Kartei")
    
    ' Read date range from B1 and C1
    If IsDate(wsGross.Range("B1").value) And IsDate(wsGross.Range("C1").value) Then
        startDate = CDate(wsGross.Range("B1").value)
        endDate = CDate(wsGross.Range("C1").value)
    Else
        MsgBox "Please enter valid dates in cells B1 and C1"
        Exit Sub
    End If
    
    ' Clear content from row 2 onwards
    wsGross.Rows("2:" & wsGross.Rows.count).Clear
    
    ' Create headers in row 2
    Call CreateGrossGeschichteHeaders(wsGross)
    
    ' Find last row in Kartei (column A)
    lastRow = wsKartei.Cells(wsKartei.Rows.count, 1).End(xlUp).row
    
    outputRow = 3 ' Start output from row 3
    
    ' Loop through all rows in Kartei
    For currentRow = 2 To lastRow ' Start from row 2 assuming row 1 has headers
        
        ' Get history string and operator name
        strGeschichte = wsKartei.Cells(currentRow, 52).value
        operName = wsKartei.Cells(currentRow, 49).value
        
        ' Skip if no history
        If strGeschichte <> "" Then
            
            On Error Resume Next
            ' Parse history
            Set result = ParseHistory(strGeschichte)
            
            If Err.Number <> 0 Then
                Debug.Print "Error parsing history for row " & currentRow & ": " & Err.Description
                Err.Clear
                GoTo NextRow
            End If
            On Error GoTo Cleanup
            
            ' Process each event in the history
            Dim i As Long
            For i = 1 To result.count
                Dim evt As Object
                Set evt = result(i)
                
                ' Check if event date falls within the specified range
                Dim eventDate As Date
                On Error Resume Next
                eventDate = CDate(evt("ChangeDate"))
                If Err.Number = 0 Then
                    If eventDate >= startDate And eventDate <= endDate Then
                        ' Create entry for this event
                        Call CreateGrossGeschichteEntry(wsGross, wsKartei, outputRow, evt("IsRuck"), _
                                                      evt("Reason"), CStr(evt("ChangeDate")), _
                                                      evt("Changes"), currentRow, operName)
                        outputRow = outputRow + 3 ' Each entry takes 2 rows + 1 separator row
                    End If
                End If
                Err.Clear
                On Error GoTo Cleanup
            Next i
        End If
NextRow:
    Next currentRow
    
    ' Apply final formatting and AutoFilter
    Call FormatGrossGeschichte(wsGross, outputRow - 1)
    
    ' Apply AutoFilter to the entire data range
    If outputRow > 3 Then ' Only if we have data
        ' Clear any existing AutoFilter first
        If wsGross.AutoFilterMode Then
            wsGross.AutoFilterMode = False
        End If
        
        ' Apply AutoFilter to range from headers (row 2) to last data row
        Dim filterRange As Range
        Set filterRange = wsGross.Range("A2:T" & (outputRow - 1))
        filterRange.AutoFilter
    End If
    
    MsgBox "Gross Geschichte generated successfully for date range " & Format(startDate, "dd.mm.yyyy") & " - " & Format(endDate, "dd.mm.yyyy") & vbCrLf & "Total rows: " & (outputRow - 3)
    
    Exit Sub
    
Cleanup:
    ' Re-enable automatic calculations and screen updating
    Application.Calculation = xlCalculationAutomatic
    Application.ScreenUpdating = True
    MsgBox "There is an error: " & Err.Description
End Sub

' Helper function to convert decimal separator based on system settings
Private Function ConvertDecimalSeparator(ByVal value As String, ByVal systemDecimalSeparator As String) As String
    ' Remove extra spaces (like in original ConvertAndFormatCells)
    value = Replace(value, " ", "")
    
    ' Unify decimal separators so Excel can parse them correctly (like in ConvertAndFormatCells)
    If systemDecimalSeparator = "," Then
        ' Replace any '.' with ',' in case the data was typed with a dot
        value = Replace(value, ".", ",")
    Else
        ' Replace any ',' with '.' if system uses dot
        value = Replace(value, ",", ".")
    End If
    
    ConvertDecimalSeparator = value
End Function

Private Sub CreateGrossGeschichteHeaders(ws As Worksheet)
    Dim headers As Variant
    headers = Array("ID", "Eltern", "Kind", "Gruppe I", "Gruppe II", "Jan", "Feb", "Mrz", "Apr", "Mai", "Jun", "Jul", "Aug", "Sep", "Okt", "Nov", "Dez", "Date", "Operator", "Comments")
    
    Dim i As Long
    For i = LBound(headers) To UBound(headers)
        With ws.Cells(2, i + 1) ' Row 2, start from column A (1)
            .value = headers(i)
            .Font.Bold = True
            .HorizontalAlignment = xlCenter
        End With
    Next i
    
    ' Set column widths
    ws.Columns("A").ColumnWidth = 8
    ws.Columns("B").ColumnWidth = 20
    ws.Columns("C").ColumnWidth = 20
    ws.Columns("D").ColumnWidth = 22
    ws.Columns("E").ColumnWidth = 22
    ws.Columns("F:Q").ColumnWidth = 6
    ws.Columns("R").ColumnWidth = 12
    ws.Columns("S").ColumnWidth = 9
    ws.Columns("T").ColumnWidth = 40
End Sub

Private Sub CreateGrossGeschichteEntry(wsGross As Worksheet, wsKartei As Worksheet, _
                                     startRow As Long, isRuck As Boolean, _
                                     Reason As String, changeDate As String, _
                                     Changes As Object, karteiRow As Long, operName As String)
    
    ' Fill basic information for both rows (War and Ist)
    Dim rowWar As Long
    Dim rowIst As Long
    Dim rowSeparator As Long
    Dim i As Long
    
    rowWar = startRow
    rowIst = startRow + 1
    rowSeparator = startRow + 2
    
    ' Create separator row (like in original GeshichteMachen)
    wsGross.Rows(rowSeparator).RowHeight = wsGross.StandardHeight * 1 / 4
    wsGross.Range("A" & rowSeparator & ":T" & rowSeparator).Interior.Color = RGB(192, 192, 192)
    
    ' Fill basic data for both rows
    For i = 0 To 1
        Dim currentRow As Long
        currentRow = startRow + i
        
        wsGross.Range("A" & currentRow).value = wsKartei.Cells(karteiRow, 1).value ' ID
        wsGross.Range("B" & currentRow).value = wsKartei.Cells(karteiRow, 2).value ' Eltern
        wsGross.Range("C" & currentRow).value = wsKartei.Cells(karteiRow, 4).value ' Kind
        wsGross.Range("D" & currentRow).value = wsKartei.Cells(karteiRow, 10).value ' Gruppe I
        wsGross.Range("E" & currentRow).value = wsKartei.Cells(karteiRow, 15).value ' Gruppe II
        wsGross.Range("R" & currentRow).value = changeDate ' Date
        wsGross.Range("S" & currentRow).value = operName ' Operator
    Next i
    
    ' Fill Comments in Ist row with appropriate formatting (like in original)
    If isRuck Then
        wsGross.Range("T" & rowIst).Interior.Color = RGB(255, 153, 204) ' Pink for Ruck
    Else
        wsGross.Range("T" & rowIst).Interior.Color = RGB(204, 255, 153) ' Green for normal
    End If
    wsGross.Range("T" & rowIst).value = Reason
    
    ' Fill monthly data
    Dim monthKey As Variant
    Dim colIndex As Long
    Dim decimalSeparator As String
    
    ' Get the system decimal separator (like in original ConvertAndFormatCells)
    decimalSeparator = Application.International(xlDecimalSeparator)
    
    On Error Resume Next ' Handle potential errors with dictionary access
    
    For Each monthKey In Changes.Keys
        colIndex = 5 + CLng(monthKey) ' Columns F-Q (6-17) for months 1-12
        
        ' Validate column index
        If colIndex >= 6 And colIndex <= 17 Then
            ' War value (previous state)
            Dim warValue As String
            warValue = Changes(monthKey)("War")
            If warValue = "" Or IsEmpty(warValue) Or IsNull(warValue) Then
                wsGross.Cells(rowWar, colIndex).value = 0
            Else
                ' Convert decimal separator based on system settings (like in ConvertAndFormatCells)
                warValue = ConvertDecimalSeparator(warValue, decimalSeparator)
                If IsNumeric(warValue) Then
                    wsGross.Cells(rowWar, colIndex).value = CDbl(warValue)
                Else
                    wsGross.Cells(rowWar, colIndex).value = 0
                End If
            End If
            
            ' Ist value (new state) with highlighting
            Dim istValue As String
            istValue = Changes(monthKey)("Ist")
            If istValue = "" Or IsEmpty(istValue) Or IsNull(istValue) Then
                wsGross.Cells(rowIst, colIndex).value = 0
            Else
                ' Convert decimal separator based on system settings
                istValue = ConvertDecimalSeparator(istValue, decimalSeparator)
                If IsNumeric(istValue) Then
                    wsGross.Cells(rowIst, colIndex).value = CDbl(istValue)
                Else
                    wsGross.Cells(rowIst, colIndex).value = 0
                End If
            End If
            
            ' Highlight the changed cell (like in original GeshichteMachen)
            wsGross.Cells(rowIst, colIndex).Interior.Color = RGB(255, 192, 203) ' Light pink
        End If
    Next monthKey
    
    On Error GoTo 0 ' Reset error handling
    
    ' Format monthly columns as decimal
    With wsGross.Range("F" & rowWar & ":Q" & rowIst)
        .NumberFormat = "0.00"
    End With
    
    ' Add borders (like in original GeshichteMachen)
    With wsGross.Range("A" & rowWar & ":S" & rowIst).Borders
        .LineStyle = xlContinuous
        .Color = vbBlack
        .Weight = xlThin
    End With
End Sub

Private Sub FormatGrossGeschichte(ws As Worksheet, lastRow As Long)
    ' Add borders to headers
    With ws.Range("A2:T2").Borders
        .LineStyle = xlContinuous
        .Color = vbBlack
        .Weight = xlMedium
    End With
    
    ' Format headers background
    ws.Range("A2:T2").Interior.Color = RGB(220, 220, 220) ' Light gray
End Sub

