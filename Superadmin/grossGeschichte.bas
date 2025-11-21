Attribute VB_Name = "grossGeschichte"
'==========================
'   Module: grossGeschichte
'   Purpose: Generate comprehensive history report (GrossGeschichte) for all records
'   Updated: Support for ID-based tracking and new history formats
'==========================

Option Explicit

Sub GrossGeshichteMachen()
    Dim wsGross As Worksheet
    Dim wsKartei As Worksheet
    Dim startDate As Date
    Dim endDate As Date
    Dim lastRow As Long
    Dim currentRow As Long
    Dim strGeschichte As String
    Dim operName As String
    Dim recordID As String
    Dim result As Collection
    Dim outputRow As Long
    
    On Error GoTo Cleanup
    
    ' Disable screen updating and automatic calculations for better performance
    Application.ScreenUpdating = False
    Application.Calculation = xlCalculationManual
    
    ' Get worksheets
    Set wsGross = GetOrCreateGrossGeschichteSheet()
    Set wsKartei = ThisWorkbook.Worksheets("Kartei")
    
    ' Read date range from B1 and C1
    If IsDate(wsGross.Range("B1").Value) And IsDate(wsGross.Range("C1").Value) Then
        startDate = CDate(wsGross.Range("B1").Value)
        endDate = CDate(wsGross.Range("C1").Value)
    Else
        MsgBox "Please enter valid dates in cells B1 and C1", vbExclamation
        GoTo Cleanup
    End If
    
    ' Clear content from row 2 onwards
    wsGross.Rows("2:" & wsGross.Rows.Count).Clear
    
    ' Create headers in row 2
    Call CreateGrossGeschichteHeaders(wsGross)
    
    ' Find last row in Kartei (column A)
    lastRow = wsKartei.Cells(wsKartei.Rows.Count, 1).End(xlUp).Row
    
    outputRow = 3 ' Start output from row 3
    
    ' Loop through all rows in Kartei
    For currentRow = 3 To lastRow ' Start from row 3 (assuming rows 1-2 are headers)
        
        ' Get history string, operator name, and ID
        strGeschichte = wsKartei.Cells(currentRow, 52).Value ' AZ column (history)
        operName = wsKartei.Cells(currentRow, 49).Value ' AW column (operator)
        recordID = CStr(wsKartei.Cells(currentRow, 48).Value) ' AV column (ID)
        
        ' Skip if no history or no ID
        If strGeschichte <> "" And recordID <> "" Then
            
            On Error Resume Next
            ' Parse history
            Set result = ParseHistory(strGeschichte)
            
            If Err.Number <> 0 Then
                Debug.Print "Error parsing history for row " & currentRow & " (ID: " & recordID & "): " & Err.Description
                Err.Clear
                GoTo NextRow
            End If
            On Error GoTo Cleanup
            
            ' Process each event in the history
            Dim i As Long
            For i = 1 To result.Count
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
                                                      evt("Changes"), currentRow, operName, recordID)
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
    
    MsgBox "Gross Geschichte generated successfully for date range " & Format(startDate, "dd.mm.yyyy") & " - " & Format(endDate, "dd.mm.yyyy") & vbCrLf & "Total rows: " & (outputRow - 3), vbInformation
    
Cleanup:
    ' Re-enable automatic calculations and screen updating
    Application.Calculation = xlCalculationAutomatic
    Application.ScreenUpdating = True
    If Err.Number <> 0 Then
        MsgBox "There is an error: " & Err.Description, vbCritical
    End If
End Sub

' Get or create GrossGeschichte worksheet
Private Function GetOrCreateGrossGeschichteSheet() As Worksheet
    Dim ws As Worksheet
    On Error Resume Next
    Set ws = ThisWorkbook.Worksheets("GrossGeschichte")
    On Error GoTo 0
    
    If ws Is Nothing Then
        Set ws = ThisWorkbook.Worksheets.Add(After:=ThisWorkbook.Worksheets(ThisWorkbook.Worksheets.Count))
        ws.Name = "GrossGeschichte"
        
        ' Add date range inputs
        ws.Range("A1").Value = "Start Date:"
        ws.Range("B1").Value = Date - 30 ' Default: last 30 days
        ws.Range("C1").Value = Date
        ws.Range("A1").Font.Bold = True
    End If
    
    Set GetOrCreateGrossGeschichteSheet = ws
End Function

' Helper function to convert decimal separator based on system settings
Private Function ConvertDecimalSeparator(ByVal Value As String, ByVal systemDecimalSeparator As String) As String
    ' Remove extra spaces (like in original ConvertAndFormatCells)
    Value = Replace(Value, " ", "")
    
    ' Unify decimal separators so Excel can parse them correctly (like in ConvertAndFormatCells)
    If systemDecimalSeparator = "," Then
        ' Replace any '.' with ',' in case the data was typed with a dot
        Value = Replace(Value, ".", ",")
    Else
        ' Replace any ',' with '.' if system uses dot
        Value = Replace(Value, ",", ".")
    End If
    
    ConvertDecimalSeparator = Value
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
                                     Changes As Object, karteiRow As Long, operName As String, _
                                     recordID As String)
    
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
        
        wsGross.Range("A" & currentRow).Value = recordID ' Use passed ID
        wsGross.Range("B" & currentRow).Value = wsKartei.Cells(karteiRow, 2).Value ' Eltern
        wsGross.Range("C" & currentRow).Value = wsKartei.Cells(karteiRow, 4).Value ' Kind
        wsGross.Range("D" & currentRow).Value = wsKartei.Cells(karteiRow, 10).Value ' Gruppe I
        wsGross.Range("E" & currentRow).Value = wsKartei.Cells(karteiRow, 15).Value ' Gruppe II
        wsGross.Range("R" & currentRow).Value = changeDate ' Date
        wsGross.Range("S" & currentRow).Value = operName ' Operator
    Next i
    
    ' Fill Comments in Ist row with appropriate formatting (like in original)
    If isRuck Then
        wsGross.Range("T" & rowIst).Interior.Color = RGB(255, 153, 204) ' Pink for Ruck
    Else
        wsGross.Range("T" & rowIst).Interior.Color = RGB(204, 255, 153) ' Green for normal
    End If
    wsGross.Range("T" & rowIst).Value = Reason
    
    ' Fill monthly data and field changes
    Dim changeKey As Variant
    Dim colIndex As Long
    Dim decimalSeparator As String
    
    ' Get the system decimal separator (like in original ConvertAndFormatCells)
    decimalSeparator = Application.International(xlDecimalSeparator)
    
    On Error Resume Next ' Handle potential errors with dictionary access
    
    For Each changeKey In Changes.Keys
        Dim keyStr As String
        keyStr = CStr(changeKey)
        
        ' Check if it's a month number (1-12)
        If IsNumeric(keyStr) Then
            Dim monthNum As Long
            monthNum = CLng(keyStr)
            
            If monthNum >= 1 And monthNum <= 12 Then
                colIndex = 5 + monthNum ' Columns F-Q (6-17) for months 1-12
                
                ' Validate column index
                If colIndex >= 6 And colIndex <= 17 Then
                    ' War value (previous state)
                    Dim warValue As String
                    warValue = Changes(changeKey)("War")
                    If warValue = "" Or IsEmpty(warValue) Or IsNull(warValue) Then
                        wsGross.Cells(rowWar, colIndex).Value = 0
                    Else
                        ' Convert decimal separator based on system settings (like in ConvertAndFormatCells)
                        warValue = ConvertDecimalSeparator(warValue, decimalSeparator)
                        If IsNumeric(warValue) Then
                            wsGross.Cells(rowWar, colIndex).Value = CDbl(warValue)
                        Else
                            wsGross.Cells(rowWar, colIndex).Value = 0
                        End If
                    End If
                    
                    ' Ist value (new state) with highlighting
                    Dim istValue As String
                    istValue = Changes(changeKey)("Ist")
                    If istValue = "" Or IsEmpty(istValue) Or IsNull(istValue) Then
                        wsGross.Cells(rowIst, colIndex).Value = 0
                    Else
                        ' Convert decimal separator based on system settings
                        istValue = ConvertDecimalSeparator(istValue, decimalSeparator)
                        If IsNumeric(istValue) Then
                            wsGross.Cells(rowIst, colIndex).Value = CDbl(istValue)
                        Else
                            wsGross.Cells(rowIst, colIndex).Value = 0
                        End If
                    End If
                    
                    ' Highlight the changed cell (like in original GeshichteMachen)
                    wsGross.Cells(rowIst, colIndex).Interior.Color = RGB(255, 192, 203) ' Light pink
                End If
            End If
        Else
            ' Non-numeric key - field names like "Address", "Subject1", "Subject2"
            ' Add to comments field with field name
            Dim fieldComment As String
            fieldComment = keyStr & ": Was(" & Changes(changeKey)("War") & ") -> Is(" & Changes(changeKey)("Ist") & ")"
            
            If wsGross.Range("T" & rowIst).Value <> "" Then
                wsGross.Range("T" & rowIst).Value = wsGross.Range("T" & rowIst).Value & vbCrLf & fieldComment
            Else
                wsGross.Range("T" & rowIst).Value = fieldComment
            End If
        End If
    Next changeKey
    
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

