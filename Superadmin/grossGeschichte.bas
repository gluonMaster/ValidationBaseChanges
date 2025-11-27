Attribute VB_Name = "grossGeschichte"
'==========================
'   Module: grossGeschichte
'   Purpose: Generate comprehensive history report (GrossGeschichte) for all records
'   Updated: Support for ID-based tracking and new history formats
'
'   New Column Structure (A-AE):
'   A=FamilyID, B=Parent, C=Child, D=Birthdate, E=Address, F=Phone, G=Mobile, H=Email
'   I=Subject1, J=Price1, K=Subject2, L=Price2, M-X=Months 1-12, Y-AA=Extra1-3
'   AB=Comments, AC=Decision, AD=Decline Comment, AE=RecordID (hidden)
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
    Dim showOnlyLastChange As Boolean
    Dim userChoice As VbMsgBoxResult
    Dim segments() As String
    Dim lastEvent As Object
    Dim i As Long
    Dim evt As Object
    Dim segmentText As String
    Dim fieldChanges As Object
    Dim eventDate As Date
    
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
    
    ' Ask user which mode to use
    userChoice = MsgBox("Choose report mode:" & vbCrLf & vbCrLf & _
                        "YES - Show all history events (Mode A)" & vbCrLf & _
                        "NO - Show only last change per ID (Mode B)", _
                        vbYesNoCancel + vbQuestion, "GrossGeschichte Mode")
    
    If userChoice = vbCancel Then
        GoTo Cleanup
    End If
    
    showOnlyLastChange = (userChoice = vbNo)
    
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
            
            ' Split raw history into segments so we can parse Address/Subject1/Subject2 per event
            segments = Split(strGeschichte, "||")
            
            ' DEBUG: No error masking - let errors propagate to Cleanup
            ' Parse history using tested parser from valid_ParseHistory module
            Debug.Print "GrossGeschichte: Parsing row=" & currentRow & ", ID=" & recordID
            Set result = valid_ParseHistory.ParseHistory(strGeschichte)
            
            Debug.Print "GrossGeschichte: ParseHistory returned, checking result"
            If result Is Nothing Then
                Debug.Print "GrossGeschichte: result is Nothing!"
                GoTo NextRow
            End If
            
            Debug.Print "GrossGeschichte: result.Count=" & result.Count & ", showOnlyLastChange=" & showOnlyLastChange
            
            ' If Mode B, find the last event in date range for this ID
            If showOnlyLastChange Then
                Debug.Print "GrossGeschichte: Mode B - calling FindLastEventInRange"
                Set lastEvent = FindLastEventInRange(result, segments, startDate, endDate)
                
                If Not lastEvent Is Nothing Then
                    Debug.Print "GrossGeschichte: Mode B - calling CreateGrossGeschichteEntry"
                    ' Create entry for the last event only
                    Call CreateGrossGeschichteEntry(wsGross, wsKartei, outputRow, lastEvent("IsRuck"), _
                                                  lastEvent("Reason"), CStr(lastEvent("ChangeDate")), _
                                                  lastEvent("Changes"), currentRow, operName, recordID)
                    outputRow = outputRow + 3 ' Each entry takes 2 rows + 1 separator row
                End If
            Else
                ' Mode A: Process each event in the history
                For i = 1 To result.Count
                    Debug.Print "GrossGeschichte: Mode A - processing event " & i & " of " & result.Count
                    Set evt = result(i)
                    
                    Debug.Print "GrossGeschichte: evt keys - IsRuck=" & evt("IsRuck") & ", ChangeDate=" & evt("ChangeDate")
                    
                    ' Enrich Changes with Address/Subject1/Subject2 parsed from the corresponding raw segment
                    segmentText = ""
                    On Error Resume Next
                    If (i - 1) >= LBound(segments) And (i - 1) <= UBound(segments) Then
                        segmentText = Trim$(segments(i - 1))
                    End If
                    On Error GoTo Cleanup
                    
                    Debug.Print "GrossGeschichte: calling ParseFieldChangesFromSegment"
                    Set fieldChanges = ParseFieldChangesFromSegment(segmentText)
                    Debug.Print "GrossGeschichte: calling MergeFieldChangesIntoChanges"
                    Call MergeFieldChangesIntoChanges(evt("Changes"), fieldChanges)
                    
                    ' Check if event date falls within the specified range
                    Debug.Print "GrossGeschichte: checking date range"
                    On Error Resume Next
                    eventDate = CDate(evt("ChangeDate"))
                    If Err.Number = 0 Then
                        Debug.Print "GrossGeschichte: eventDate=" & eventDate & ", startDate=" & startDate & ", endDate=" & endDate
                        If eventDate >= startDate And eventDate <= endDate Then
                            Debug.Print "GrossGeschichte: IN RANGE - calling CreateGrossGeschichteEntry for event " & i
                            ' Create entry for this event
                            Call CreateGrossGeschichteEntry(wsGross, wsKartei, outputRow, evt("IsRuck"), _
                                                          evt("Reason"), CStr(evt("ChangeDate")), _
                                                          evt("Changes"), currentRow, operName, recordID)
                            Debug.Print "GrossGeschichte: CreateGrossGeschichteEntry completed for event " & i
                            outputRow = outputRow + 3 ' Each entry takes 2 rows + 1 separator row
                        Else
                            Debug.Print "GrossGeschichte: OUT OF RANGE - skipping event " & i
                        End If
                    Else
                        Debug.Print "GrossGeschichte: CDate error for event " & i & ": " & Err.Description
                    End If
                    Err.Clear
                    On Error GoTo Cleanup
                Next i
            End If
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
        Set filterRange = wsGross.Range("A2:AD" & (outputRow - 1))
        filterRange.AutoFilter
    End If
    
    ' CRITICAL: Restore date format in B1:C1 after all operations
    ' This ensures dates display as dd.mm.yyyy, not as numbers
    wsGross.Range("B1:C1").NumberFormat = "dd.mm.yyyy"
    
    Dim modeText As String
    If showOnlyLastChange Then
        modeText = "Mode B (Last change per ID only)"
    Else
        modeText = "Mode A (All events)"
    End If
    
    MsgBox "Gross Geschichte generated successfully for date range " & Format(startDate, "dd.mm.yyyy") & " - " & Format(endDate, "dd.mm.yyyy") & vbCrLf & _
           "Mode: " & modeText & vbCrLf & _
           "Total rows: " & (outputRow - 3), vbInformation
    
Cleanup:
    ' Re-enable automatic calculations and screen updating
    Application.Calculation = xlCalculationAutomatic
    Application.ScreenUpdating = True
    If Err.Number <> 0 Then
        Debug.Print "GrossGeschichte ERROR: row=" & currentRow & ", ID=" & recordID & ", Err=" & Err.Number & " - " & Err.Description
        MsgBox "There is an error: " & Err.Description & vbCrLf & _
               "Row: " & currentRow & ", ID: " & recordID & vbCrLf & _
               "Error Number: " & Err.Number, vbCritical
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
        ws.Range("A1").Font.Bold = True
        
        ' Set date format dd.mm.yyyy BEFORE setting values
        ws.Range("B1:C1").NumberFormat = "dd.mm.yyyy"
        ws.Range("B1").Value = Date - 30 ' Default: last 30 days
        ws.Range("C1").Value = Date
    Else
        ' Sheet exists - ensure date format is correct and C1 = today
        ws.Range("B1:C1").NumberFormat = "dd.mm.yyyy"
        
        ' Set B1 to 30 days ago only if not a valid date
        If Not IsDate(ws.Range("B1").Value) Or IsEmpty(ws.Range("B1").Value) Then
            ws.Range("B1").Value = Date - 30
        End If
        
        ' Always set C1 to current date
        ws.Range("C1").Value = Date
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
    ' New column structure: A-AD (30 columns) + AE (hidden ID)
    ' A=FamilyID, B=Parent, C=Child, D=Birthdate, E=Address, F=Phone, G=Mobile, H=Email
    ' I=Subject1, J=Price1, K=Subject2, L=Price2, M-X=Months 1-12, Y-AA=Extra1-3
    ' AB=Comments, AC=Decision, AD=Decline Comment, AE=RecordID (hidden)
    
    Dim headers As Variant
    headers = Array("FamilyID", "Parent", "Child", "Birthdate", "Address", "Phone", "Mobile", "Email", _
                    "Subject1", "Price1", "Subject2", "Price2", _
                    "Jan", "Feb", "Mrz", "Apr", "Mai", "Jun", "Jul", "Aug", "Sep", "Okt", "Nov", "Dez", _
                    "Extra1", "Extra2", "Extra3", "Comments", "Decision", "Decline Comment", "RecordID")
    
    Dim i As Long
    For i = LBound(headers) To UBound(headers)
        With ws.Cells(2, i + 1) ' Row 2, start from column A (1)
            .Value = headers(i)
            .Font.Bold = True
            .HorizontalAlignment = xlCenter
        End With
    Next i
    
    ' Set column widths for optimal display
    ws.Columns("A").ColumnWidth = 10  ' FamilyID
    ws.Columns("B").ColumnWidth = 20  ' Parent
    ws.Columns("C").ColumnWidth = 18  ' Child
    ws.Columns("D").ColumnWidth = 12  ' Birthdate
    ws.Columns("E").ColumnWidth = 25  ' Address
    ws.Columns("F").ColumnWidth = 14  ' Phone
    ws.Columns("G").ColumnWidth = 14  ' Mobile
    ws.Columns("H").ColumnWidth = 22  ' Email
    ws.Columns("I").ColumnWidth = 18  ' Subject1
    ws.Columns("J").ColumnWidth = 8   ' Price1
    ws.Columns("K").ColumnWidth = 18  ' Subject2
    ws.Columns("L").ColumnWidth = 8   ' Price2
    ws.Columns("M:X").ColumnWidth = 6 ' Months 1-12
    ws.Columns("Y:AA").ColumnWidth = 15 ' Extra1-3
    ws.Columns("AB").ColumnWidth = 35 ' Comments
    ws.Columns("AC").ColumnWidth = 12 ' Decision
    ws.Columns("AD").ColumnWidth = 35 ' Decline Comment
    
    ' Hide column AE (RecordID) - used for internal ID tracking
    ws.Columns("AE").Hidden = True
    
    ' Set text format for columns that should not be auto-converted by Excel
    ' (Phone, Mobile, Birthdate, etc.)
    ' IMPORTANT: Apply text format only to data rows (2 and below), not to row 1 with dates
    ws.Range("A2:L" & ws.Rows.Count).NumberFormat = "@" ' Text format
    ws.Range("Y2:AE" & ws.Rows.Count).NumberFormat = "@" ' Text format for Extra1-3, Comments, Decision, Decline Comment, RecordID
    
    ' Ensure date format is preserved in B1:C1 (row 1 contains date inputs)
    ws.Range("B1:C1").NumberFormat = "dd.mm.yyyy"
    
    ' Months columns M-X will be formatted as numbers in CreateGrossGeschichteEntry
End Sub

Private Sub CreateGrossGeschichteEntry(wsGross As Worksheet, wsKartei As Worksheet, _
                                     startRow As Long, isRuck As Boolean, _
                                     Reason As String, changeDate As String, _
                                     Changes As Object, karteiRow As Long, operName As String, _
                                     recordID As String)
    
    ' New column mapping:
    ' A=FamilyID(1), B=Parent(2), C=Child(4), D=Birthdate(5), E=Address(6), F=Phone(7), G=Mobile(8), H=Email(9)
    ' I=Subject1(10), J=Price1(13), K=Subject2(15), L=Price2(18), M-X=Months(21-32), Y-AA=Extra(37-39)
    ' AB=Comments, AC=Decision, AD=Decline Comment
    
    Dim rowWar As Long
    Dim rowIst As Long
    Dim rowSeparator As Long
    Dim i As Long
    
    rowWar = startRow
    rowIst = startRow + 1
    rowSeparator = startRow + 2
    
    ' Create separator row
    wsGross.Rows(rowSeparator).RowHeight = wsGross.StandardHeight * 1 / 4
    wsGross.Range("A" & rowSeparator & ":AD" & rowSeparator).Interior.Color = RGB(192, 192, 192)
    
    ' Fill basic data for both rows (War and Ist)
    For i = 0 To 1
        Dim currentRow As Long
        currentRow = startRow + i
        
        ' Store values as text to prevent Excel auto-conversion
        wsGross.Range("A" & currentRow).Value = CStr(wsKartei.Cells(karteiRow, 1).Value) ' FamilyID from column A(1)
        wsGross.Range("B" & currentRow).Value = CStr(wsKartei.Cells(karteiRow, 2).Value) ' Parent from column B(2)
        wsGross.Range("C" & currentRow).Value = CStr(wsKartei.Cells(karteiRow, 4).Value) ' Child from column D(4)
        wsGross.Range("D" & currentRow).Value = FormatAsText(wsKartei.Cells(karteiRow, 5).Value) ' Birthdate from column E(5)
        wsGross.Range("E" & currentRow).Value = CStr(wsKartei.Cells(karteiRow, 6).Value) ' Address from column F(6)
        wsGross.Range("F" & currentRow).Value = FormatAsText(wsKartei.Cells(karteiRow, 7).Value) ' Phone from column G(7)
        wsGross.Range("G" & currentRow).Value = FormatAsText(wsKartei.Cells(karteiRow, 8).Value) ' Mobile from column H(8)
        wsGross.Range("H" & currentRow).Value = CStr(wsKartei.Cells(karteiRow, 9).Value) ' Email from column I(9)
        wsGross.Range("I" & currentRow).Value = CStr(wsKartei.Cells(karteiRow, 10).Value) ' Subject1 from column J(10)
        wsGross.Range("J" & currentRow).Value = CStr(wsKartei.Cells(karteiRow, 13).Value) ' Price1 from column M(13)
        wsGross.Range("K" & currentRow).Value = CStr(wsKartei.Cells(karteiRow, 15).Value) ' Subject2 from column O(15)
        wsGross.Range("L" & currentRow).Value = CStr(wsKartei.Cells(karteiRow, 18).Value) ' Price2 from column R(18)
        
        ' Extra subjects 1-3 from columns AK-AM (37-39)
        wsGross.Range("Y" & currentRow).Value = CStr(wsKartei.Cells(karteiRow, 37).Value) ' Extra1
        wsGross.Range("Z" & currentRow).Value = CStr(wsKartei.Cells(karteiRow, 38).Value) ' Extra2
        wsGross.Range("AA" & currentRow).Value = CStr(wsKartei.Cells(karteiRow, 39).Value) ' Extra3
        
        ' Store RecordID in hidden column AE (31) for decision processing
        wsGross.Range("AE" & currentRow).Value = recordID
    Next i
    
    ' Fill Comments in Ist row (column AB) with appropriate formatting
    If isRuck Then
        wsGross.Range("AB" & rowIst).Interior.Color = RGB(255, 153, 204) ' Pink for Ruck
    Else
        wsGross.Range("AB" & rowIst).Interior.Color = RGB(204, 255, 153) ' Green for normal
    End If
    wsGross.Range("AB" & rowIst).Value = Reason
    
    ' Highlight SEPA rows explicitly in comments (now column AB)
    Dim sepaMarker As String
    sepaMarker = Trim$(UCase$(CStr(wsKartei.Cells(karteiRow, 47).Value)))
    If sepaMarker = "SEPA" Then
        Dim sepaText As String
        sepaText = "SEPA"
        
        Dim commentCell As Range
        Set commentCell = wsGross.Range("AB" & rowIst)
        
        Dim baseText As String
        baseText = CStr(commentCell.Value)
        
        Dim startPos As Long
        If baseText <> "" Then
            commentCell.Value = baseText & vbCrLf & sepaText
            startPos = Len(baseText) + 2 ' account for vbCrLf
        Else
            commentCell.Value = sepaText
            startPos = 1
        End If
        
        With commentCell.Characters(startPos, Len(sepaText)).Font
            .Color = vbRed
            .Bold = True
        End With
    End If
    
    ' Fill monthly data and field changes
    ' New column mapping: Months are in M-X (columns 13-24)
    Dim changeKey As Variant
    Dim colIndex As Long
    Dim decimalSeparator As String
    
    ' Variables for tracking field changes from history
    Dim addressWar As String, addressIst As String
    Dim subject1War As String, subject1Ist As String
    Dim subject2War As String, subject2Ist As String
    Dim phoneWar As String, phoneIst As String
    Dim mobileWar As String, mobileIst As String
    Dim emailWar As String, emailIst As String
    Dim birthdateWar As String, birthdateIst As String
    Dim price1War As String, price1Ist As String
    Dim price2War As String, price2Ist As String
    Dim extra1War As String, extra1Ist As String
    Dim extra2War As String, extra2Ist As String
    Dim extra3War As String, extra3Ist As String
    Dim familyIDWar As String, familyIDIst As String
    Dim parentWar As String, parentIst As String
    Dim childWar As String, childIst As String
    
    ' Initialize all to empty
    addressWar = "": addressIst = ""
    subject1War = "": subject1Ist = ""
    subject2War = "": subject2Ist = ""
    phoneWar = "": phoneIst = ""
    mobileWar = "": mobileIst = ""
    emailWar = "": emailIst = ""
    birthdateWar = "": birthdateIst = ""
    price1War = "": price1Ist = ""
    price2War = "": price2Ist = ""
    extra1War = "": extra1Ist = ""
    extra2War = "": extra2Ist = ""
    extra3War = "": extra3Ist = ""
    familyIDWar = "": familyIDIst = ""
    parentWar = "": parentIst = ""
    childWar = "": childIst = ""
    
    ' Get the system decimal separator
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
                ' New mapping: Months 1-12 are in columns M-X (13-24)
                colIndex = 12 + monthNum ' M=13 for month 1, X=24 for month 12
                
                ' Validate column index
                If colIndex >= 13 And colIndex <= 24 Then
                    ' War value (previous state)
                    Dim warValue As String
                    warValue = Changes(changeKey)("War")
                    If warValue = "" Or IsEmpty(warValue) Or IsNull(warValue) Then
                        wsGross.Cells(rowWar, colIndex).Value = 0
                    Else
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
                        istValue = ConvertDecimalSeparator(istValue, decimalSeparator)
                        If IsNumeric(istValue) Then
                            wsGross.Cells(rowIst, colIndex).Value = CDbl(istValue)
                        Else
                            wsGross.Cells(rowIst, colIndex).Value = 0
                        End If
                    End If
                    
                    ' Highlight the changed cell
                    wsGross.Cells(rowIst, colIndex).Interior.Color = RGB(255, 192, 203) ' Light pink
                End If
            End If
        Else
            ' Non-numeric key - capture values for dedicated columns
            Dim warText As String
            Dim istText As String
            
            warText = Changes(changeKey)("War")
            istText = Changes(changeKey)("Ist")
            
            Select Case keyStr
                Case "Address"
                    addressWar = warText
                    addressIst = istText
                Case "Subject1", "SB1"
                    subject1War = warText
                    subject1Ist = istText
                Case "Subject2", "SB2"
                    subject2War = warText
                    subject2Ist = istText
                Case "TEL"
                    phoneWar = warText
                    phoneIst = istText
                Case "MOB"
                    mobileWar = warText
                    mobileIst = istText
                Case "EML"
                    emailWar = warText
                    emailIst = istText
                Case "DOB"
                    birthdateWar = warText
                    birthdateIst = istText
                Case "PR1"
                    price1War = warText
                    price1Ist = istText
                Case "PR2"
                    price2War = warText
                    price2Ist = istText
                Case "EX1"
                    extra1War = warText
                    extra1Ist = istText
                Case "EX2"
                    extra2War = warText
                    extra2Ist = istText
                Case "EX3"
                    extra3War = warText
                    extra3Ist = istText
                Case "FID"
                    familyIDWar = warText
                    familyIDIst = istText
                Case "PAR"
                    parentWar = warText
                    parentIst = istText
                Case "CHD"
                    childWar = warText
                    childIst = istText
            End Select
        End If
    Next changeKey
    
    On Error GoTo 0 ' Reset error handling
    
    ' Apply field changes to dedicated columns with highlighting
    ' Column E = Address (5)
    If addressWar <> "" Or addressIst <> "" Then
        wsGross.Cells(rowWar, 5).Value = addressWar
        wsGross.Cells(rowIst, 5).Value = addressIst
        wsGross.Cells(rowIst, 5).Interior.Color = RGB(255, 192, 203)
    End If
    
    ' Column I = Subject1 (9)
    If subject1War <> "" Or subject1Ist <> "" Then
        wsGross.Cells(rowWar, 9).Value = subject1War
        wsGross.Cells(rowIst, 9).Value = subject1Ist
        wsGross.Cells(rowIst, 9).Interior.Color = RGB(255, 192, 203)
    End If
    
    ' Column K = Subject2 (11)
    If subject2War <> "" Or subject2Ist <> "" Then
        wsGross.Cells(rowWar, 11).Value = subject2War
        wsGross.Cells(rowIst, 11).Value = subject2Ist
        wsGross.Cells(rowIst, 11).Interior.Color = RGB(255, 192, 203)
    End If
    
    ' Column F = Phone (6)
    If phoneWar <> "" Or phoneIst <> "" Then
        wsGross.Cells(rowWar, 6).Value = phoneWar
        wsGross.Cells(rowIst, 6).Value = phoneIst
        wsGross.Cells(rowIst, 6).Interior.Color = RGB(255, 192, 203)
    End If
    
    ' Column G = Mobile (7)
    If mobileWar <> "" Or mobileIst <> "" Then
        wsGross.Cells(rowWar, 7).Value = mobileWar
        wsGross.Cells(rowIst, 7).Value = mobileIst
        wsGross.Cells(rowIst, 7).Interior.Color = RGB(255, 192, 203)
    End If
    
    ' Column H = Email (8)
    If emailWar <> "" Or emailIst <> "" Then
        wsGross.Cells(rowWar, 8).Value = emailWar
        wsGross.Cells(rowIst, 8).Value = emailIst
        wsGross.Cells(rowIst, 8).Interior.Color = RGB(255, 192, 203)
    End If
    
    ' Column D = Birthdate (4)
    If birthdateWar <> "" Or birthdateIst <> "" Then
        wsGross.Cells(rowWar, 4).Value = birthdateWar
        wsGross.Cells(rowIst, 4).Value = birthdateIst
        wsGross.Cells(rowIst, 4).Interior.Color = RGB(255, 192, 203)
    End If
    
    ' Column J = Price1 (10)
    If price1War <> "" Or price1Ist <> "" Then
        wsGross.Cells(rowWar, 10).Value = price1War
        wsGross.Cells(rowIst, 10).Value = price1Ist
        wsGross.Cells(rowIst, 10).Interior.Color = RGB(255, 192, 203)
    End If
    
    ' Column L = Price2 (12)
    If price2War <> "" Or price2Ist <> "" Then
        wsGross.Cells(rowWar, 12).Value = price2War
        wsGross.Cells(rowIst, 12).Value = price2Ist
        wsGross.Cells(rowIst, 12).Interior.Color = RGB(255, 192, 203)
    End If
    
    ' Column Y = Extra1 (25)
    If extra1War <> "" Or extra1Ist <> "" Then
        wsGross.Cells(rowWar, 25).Value = extra1War
        wsGross.Cells(rowIst, 25).Value = extra1Ist
        wsGross.Cells(rowIst, 25).Interior.Color = RGB(255, 192, 203)
    End If
    
    ' Column Z = Extra2 (26)
    If extra2War <> "" Or extra2Ist <> "" Then
        wsGross.Cells(rowWar, 26).Value = extra2War
        wsGross.Cells(rowIst, 26).Value = extra2Ist
        wsGross.Cells(rowIst, 26).Interior.Color = RGB(255, 192, 203)
    End If
    
    ' Column AA = Extra3 (27)
    If extra3War <> "" Or extra3Ist <> "" Then
        wsGross.Cells(rowWar, 27).Value = extra3War
        wsGross.Cells(rowIst, 27).Value = extra3Ist
        wsGross.Cells(rowIst, 27).Interior.Color = RGB(255, 192, 203)
    End If
    
    ' Column A = FamilyID (1)
    If familyIDWar <> "" Or familyIDIst <> "" Then
        wsGross.Cells(rowWar, 1).Value = familyIDWar
        wsGross.Cells(rowIst, 1).Value = familyIDIst
        wsGross.Cells(rowIst, 1).Interior.Color = RGB(255, 192, 203)
    End If
    
    ' Column B = Parent (2)
    If parentWar <> "" Or parentIst <> "" Then
        wsGross.Cells(rowWar, 2).Value = parentWar
        wsGross.Cells(rowIst, 2).Value = parentIst
        wsGross.Cells(rowIst, 2).Interior.Color = RGB(255, 192, 203)
    End If
    
    ' Column C = Child (3)
    If childWar <> "" Or childIst <> "" Then
        wsGross.Cells(rowWar, 3).Value = childWar
        wsGross.Cells(rowIst, 3).Value = childIst
        wsGross.Cells(rowIst, 3).Interior.Color = RGB(255, 192, 203)
    End If
    
    ' Format monthly columns M-X (13-24) as decimal with 2 places
    With wsGross.Range("M" & rowWar & ":X" & rowIst)
        .NumberFormat = "0.00"
    End With
    
    ' Add borders to the full row range A-AD
    With wsGross.Range("A" & rowWar & ":AD" & rowIst).Borders
        .LineStyle = xlContinuous
        .Color = vbBlack
        .Weight = xlThin
    End With
End Sub

' Helper function to format value as text (prevents Excel auto-conversion)
Private Function FormatAsText(ByVal Value As Variant) As String
    If IsEmpty(Value) Or IsNull(Value) Then
        FormatAsText = ""
    ElseIf IsDate(Value) Then
        ' Format dates as dd.mm.yyyy to prevent conversion
        FormatAsText = Format(Value, "dd.mm.yyyy")
    Else
        FormatAsText = CStr(Value)
    End If
End Function

Private Sub FormatGrossGeschichte(ws As Worksheet, lastRow As Long)
    ' Add borders to headers (now A2:AD2)
    With ws.Range("A2:AD2").Borders
        .LineStyle = xlContinuous
        .Color = vbBlack
        .Weight = xlMedium
    End With
    
    ' Format headers background
    ws.Range("A2:AD2").Interior.Color = RGB(220, 220, 220) ' Light gray
End Sub

' Parse Address/Subject1/Subject2 changes from a raw history segment
' Supports both formats:
'   - Legacy: Address: Was(X); Is(Y). Subject1: Was(X); Is(Y). Subject2: Was(X); Is(Y).
'   - New: ADR(X->Y) SB1(X->Y) SB2(X->Y)
Function ParseFieldChangesFromSegment(ByVal segment As String) As Object
    On Error GoTo ErrHandler
    
    Dim result As Object
    Set result = CreateObject("Scripting.Dictionary")
    
    Debug.Print "ParseFieldChangesFromSegment: ENTER, segment='" & Left(segment, 50) & "...'"
    
    If Len(Trim$(segment)) = 0 Then
        Debug.Print "ParseFieldChangesFromSegment: empty segment, returning empty dict"
        Set ParseFieldChangesFromSegment = result
        Exit Function
    End If
    
    ' First try legacy format: FieldName: Was(X); Is(Y).
    Debug.Print "ParseFieldChangesFromSegment: trying legacy format"
    Call AddFieldChangeToDict(segment, "Address", result)
    Call AddFieldChangeToDict(segment, "Subject1", result)
    Call AddFieldChangeToDict(segment, "Subject2", result)
    
    ' Then try new format: TAG(X->Y)
    Debug.Print "ParseFieldChangesFromSegment: trying new format"
    Call AddNewFormatFieldChangeToDict(segment, "ADR", "Address", result)
    Call AddNewFormatFieldChangeToDict(segment, "SB1", "Subject1", result)
    Call AddNewFormatFieldChangeToDict(segment, "SB2", "Subject2", result)
    
    Debug.Print "ParseFieldChangesFromSegment: returning dict with Count=" & result.Count
    Set ParseFieldChangesFromSegment = result
    Exit Function
    
ErrHandler:
    Debug.Print "ERROR in ParseFieldChangesFromSegment: Err=" & Err.Number & " - " & Err.Description
    Err.Raise Err.Number, "grossGeschichte.ParseFieldChangesFromSegment", Err.Description
End Function

' Merge additional field changes into the base Changes dictionary used by GrossGeschichte
Sub MergeFieldChangesIntoChanges(ByVal baseChanges As Object, ByVal fieldChanges As Object)
    On Error GoTo ErrHandler
    
    If baseChanges Is Nothing Then
        Debug.Print "MergeFieldChanges: baseChanges is Nothing, exiting"
        Exit Sub
    End If
    If fieldChanges Is Nothing Then
        Debug.Print "MergeFieldChanges: fieldChanges is Nothing, exiting"
        Exit Sub
    End If
    
    Debug.Print "MergeFieldChanges: baseChanges.Count=" & baseChanges.Count & ", fieldChanges.Count=" & fieldChanges.Count
    
    Dim key As Variant
    For Each key In fieldChanges.Keys
        Debug.Print "MergeFieldChanges: processing key='" & key & "', exists in base=" & baseChanges.Exists(key)
        If Not baseChanges.Exists(key) Then
            Debug.Print "MergeFieldChanges: adding key='" & key & "'"
            baseChanges.Add key, fieldChanges.Item(key)
        Else
            Debug.Print "MergeFieldChanges: updating key='" & key & "'"
            ' Use Set for object assignment (Dictionary values are objects)
            Set baseChanges.Item(key) = fieldChanges.Item(key)
        End If
    Next key
    
    Debug.Print "MergeFieldChanges: completed successfully"
    Exit Sub
    
ErrHandler:
    Debug.Print "ERROR in MergeFieldChanges: Err=" & Err.Number & " - " & Err.Description & " | key=" & key
    Err.Raise Err.Number, "grossGeschichte.MergeFieldChangesIntoChanges", Err.Description
End Sub

' Helper: extract single "FieldName: Was(...); Is(...)." block into dictionary (legacy format)
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

' Helper: extract single TAG(OLD->NEW) block into dictionary (new format)
Private Sub AddNewFormatFieldChangeToDict(ByVal segment As String, ByVal tag As String, ByVal fieldName As String, ByVal dict As Object)
    ' If already parsed by legacy format, skip
    If dict.Exists(fieldName) Then Exit Sub
    
    ' Use regex to find TAG(value->value)
    Dim regex As Object
    Set regex = CreateObject("VBScript.RegExp")
    
    With regex
        .Pattern = tag & "\(([^)]*)->([^)]*)\)"
        .IgnoreCase = True
        .Global = False
    End With
    
    Dim matches As Object
    Set matches = regex.Execute(segment)
    
    If matches.Count = 0 Then Exit Sub
    
    Dim warVal As String
    Dim istVal As String
    
    warVal = Trim$(matches(0).SubMatches(0))
    istVal = Trim$(matches(0).SubMatches(1))
    
    Dim fieldDict As Object
    Set fieldDict = CreateObject("Scripting.Dictionary")
    fieldDict.Add "War", warVal
    fieldDict.Add "Ist", istVal
    
    dict.Add fieldName, fieldDict
End Sub

' Find the last event in date range for a given ID
' Returns the event with the latest date, and if multiple events have the same date, returns the last one in the collection
Function FindLastEventInRange(ByVal events As Collection, ByRef segments() As String, ByVal startDate As Date, ByVal endDate As Date) As Object
    Dim lastEvt As Object
    Set lastEvt = Nothing
    
    Dim maxDate As Date
    maxDate = 0
    
    Dim maxDateIndex As Long
    maxDateIndex = 0
    
    Dim i As Long
    For i = 1 To events.Count
        Dim evt As Object
        Set evt = events(i)
        
        ' Enrich this event with field changes
        Dim segmentText As String
        segmentText = ""
        On Error Resume Next
        If (i - 1) >= LBound(segments) And (i - 1) <= UBound(segments) Then
            segmentText = Trim$(segments(i - 1))
        End If
        On Error GoTo 0
        
        Dim fieldChanges As Object
        Set fieldChanges = ParseFieldChangesFromSegment(segmentText)
        Call MergeFieldChangesIntoChanges(evt("Changes"), fieldChanges)
        
        ' Check if event date falls within range
        Dim eventDate As Date
        On Error Resume Next
        eventDate = CDate(evt("ChangeDate"))
        
        If Err.Number = 0 Then
            If eventDate >= startDate And eventDate <= endDate Then
                ' Check if this is the latest date or same date but later position
                If lastEvt Is Nothing Then
                    Set lastEvt = evt
                    maxDate = eventDate
                    maxDateIndex = i
                ElseIf eventDate > maxDate Then
                    ' Found a later date
                    Set lastEvt = evt
                    maxDate = eventDate
                    maxDateIndex = i
                ElseIf eventDate = maxDate And i > maxDateIndex Then
                    ' Same date but later in the collection (last in order)
                    Set lastEvt = evt
                    maxDateIndex = i
                End If
            End If
        End If
        Err.Clear
    Next i
    
    Set FindLastEventInRange = lastEvt
End Function
