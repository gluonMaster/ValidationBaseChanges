Attribute VB_Name = "grossGeschichte"
'==========================
'   Module: grossGeschichte
'   Purpose: Generate pure history report (Geschichte) for viewing only
'   Updated: Now outputs to "Geschichte" sheet (not GrossGeschichte)
'           Decision columns (AC/AD/AE) removed - decisions handled by valid_GrossGeschichteDecision
'
'   MULTI-YEAR SUPPORT:
'   - GrossGeshichteMachen now auto-detects the year from user's target sheet
'   - Routes to valid_HistoryPerYear.GrossGeschichteMachenForYear for year-specific behavior
'   - Falls back to legacy single-sheet behavior for backward compatibility
'
'   For explicit year-specific calls, use valid_HistoryPerYear module:
'   - GrossGeschichteMachen24, GrossGeschichteMachen25, GrossGeschichteMachen26
'
'   Column Structure (A-AB):
'   A=FamilyID, B=Parent, C=Child, D=Birthdate, E=Address, F=Phone, G=Mobile, H=Email
'   I=Subject1, J=Price1, K=Subject2, L=Price2, M-X=Months 1-12, Y-AA=Extra1-3
'   AB=Comments (no decision columns - this is a view-only history report)
'==========================

Option Explicit

' Main entry point: Generate a view-only history report
' Auto-detects year from active sheet or prompts user for year selection
' Mode A (JA): All history events
' Mode B (NEIN): Last change per ID only
' NOTE: This is a VIEW-ONLY report. For pending decisions, use valid_GrossGeschichteDecision.BuildPendingDecisionSheet
Sub GrossGeshichteMachen()
    ' Check if we're on a year-specific Geschichte_Alle sheet
    Dim activeSheetName As String
    activeSheetName = ActiveSheet.Name
    
    ' Try to detect year from sheet name (e.g., Geschichte24_Alle, Kartei25)
    Dim year2 As Integer
    year2 = valid_YearConfig.GetYearFromSheetName(activeSheetName)
    
    If year2 > 0 Then
        ' Route to year-specific implementation
        valid_HistoryPerYear.GrossGeschichteMachenForYear year2
        Exit Sub
    End If
    
    ' Check for legacy sheet names (Geschichte, Kartei)
    If activeSheetName = "Geschichte" Or activeSheetName = "Kartei" Then
        GrossGeshichteMachenLegacy
        Exit Sub
    End If
    
    ' Unknown sheet - ask user which year to use
    Dim userChoice As Integer
    userChoice = MsgBox("Welches Jahr moechten Sie fuer den Geschichte-Bericht verwenden?" & vbCrLf & vbCrLf & _
                        "JA = 2025" & vbCrLf & _
                        "NEIN = Abbrechen (bitte erst Kartei-Blatt waehlen)", _
                        vbYesNo + vbQuestion, "Jahr waehlen")
    
    If userChoice = vbYes Then
        valid_HistoryPerYear.GrossGeschichteMachenForYear 25
    End If
End Sub

' Legacy single-sheet implementation (for backward compatibility with "Kartei"/"Geschichte" sheets)
Private Sub GrossGeshichteMachenLegacy()
    Dim wsHistory As Worksheet
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
    
    ' Get worksheets - output to "Geschichte" sheet (view-only history report)
    Set wsHistory = GetOrCreateHistorySheet("Geschichte")
    Set wsKartei = ThisWorkbook.Worksheets("Kartei")
    
    ' Read date range from B1 and C1
    If IsDate(wsHistory.Range("B1").Value) And IsDate(wsHistory.Range("C1").Value) Then
        startDate = CDate(wsHistory.Range("B1").Value)
        endDate = CDate(wsHistory.Range("C1").Value)
    Else
        MsgBox "Bitte gueltige Datumsangaben in den Zellen B1 und C1 eingeben", vbExclamation
        GoTo Cleanup
    End If
    
    ' Ask user which mode to use
    ' Both modes output to "Geschichte" sheet as view-only reports
    userChoice = MsgBox("Berichtsmodus waehlen:" & vbCrLf & vbCrLf & _
                        "JA - Alle Verlaufsereignisse anzeigen (Modus A)" & vbCrLf & _
                        "NEIN - Nur letzte Aenderung pro ID anzeigen (Modus B)" & vbCrLf & vbCrLf & _
                        "Hinweis: Dies ist ein reiner Ansichtsbericht ohne Entscheidungsspalten." & vbCrLf & _
                        "Fuer Pending-Entscheidungen verwenden Sie bitte 'Ausstehende Aenderungen laden'.", _
                        vbYesNoCancel + vbQuestion, "Geschichte Bericht")
    
    If userChoice = vbCancel Then
        GoTo Cleanup
    End If
    
    showOnlyLastChange = (userChoice = vbNo)
    
    ' Clear content from row 2 onwards
    wsHistory.Rows("2:" & wsHistory.Rows.Count).Clear
    
    ' Create headers in row 2 (no decision columns - view-only report)
    Call CreateHistoryHeaders(wsHistory)
    
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
                Debug.Print "Geschichte: Mode B - calling FindLastEventInRange"
                Set lastEvent = FindLastEventInRange(result, segments, startDate, endDate)
                
                If Not lastEvent Is Nothing Then
                    Debug.Print "Geschichte: Mode B - calling CreateHistoryEntry"
                    ' Create entry for the last event only (view-only, no decision columns)
                    Call CreateHistoryEntry(wsHistory, wsKartei, outputRow, lastEvent("IsRuck"), _
                                           lastEvent("Reason"), CStr(lastEvent("ChangeDate")), _
                                           lastEvent("Changes"), currentRow, operName, recordID)
                    outputRow = outputRow + 3 ' Each entry takes 2 rows + 1 separator row
                End If
            Else
                ' Mode A: Process each event in the history
                For i = 1 To result.Count
                    Debug.Print "Geschichte: Mode A - processing event " & i & " of " & result.Count
                    Set evt = result(i)
                    
                    Debug.Print "Geschichte: evt keys - IsRuck=" & evt("IsRuck") & ", ChangeDate=" & evt("ChangeDate")
                    
                    ' Enrich Changes with Address/Subject1/Subject2 parsed from the corresponding raw segment
                    segmentText = ""
                    On Error Resume Next
                    If (i - 1) >= LBound(segments) And (i - 1) <= UBound(segments) Then
                        segmentText = Trim$(segments(i - 1))
                    End If
                    On Error GoTo Cleanup
                    
                    Debug.Print "Geschichte: calling ParseFieldChangesFromSegment"
                    Set fieldChanges = ParseFieldChangesFromSegment(segmentText)
                    Debug.Print "Geschichte: calling MergeFieldChangesIntoChanges"
                    Call MergeFieldChangesIntoChanges(evt("Changes"), fieldChanges)
                    
                    ' Check if event date falls within the specified range
                    Debug.Print "Geschichte: checking date range"
                    On Error Resume Next
                    eventDate = CDate(evt("ChangeDate"))
                    If Err.Number = 0 Then
                        Debug.Print "Geschichte: eventDate=" & eventDate & ", startDate=" & startDate & ", endDate=" & endDate
                        If eventDate >= startDate And eventDate <= endDate Then
                            Debug.Print "Geschichte: IN RANGE - calling CreateHistoryEntry for event " & i
                            ' Create entry for this event (view-only, no decision columns)
                            Call CreateHistoryEntry(wsHistory, wsKartei, outputRow, evt("IsRuck"), _
                                                   evt("Reason"), CStr(evt("ChangeDate")), _
                                                   evt("Changes"), currentRow, operName, recordID)
                            Debug.Print "Geschichte: CreateHistoryEntry completed for event " & i
                            outputRow = outputRow + 3 ' Each entry takes 2 rows + 1 separator row
                        Else
                            Debug.Print "Geschichte: OUT OF RANGE - skipping event " & i
                        End If
                    Else
                        Debug.Print "Geschichte: CDate error for event " & i & ": " & Err.Description
                    End If
                    Err.Clear
                    On Error GoTo Cleanup
                Next i
            End If
        End If
NextRow:
    Next currentRow
    
    ' Apply final formatting and AutoFilter (no decision columns - A:AB only)
    Call FormatHistorySheet(wsHistory, outputRow - 1)
    
    ' Apply AutoFilter to the entire data range (A:AB, no decision columns)
    If outputRow > 3 Then ' Only if we have data
        ' Clear any existing AutoFilter first
        If wsHistory.AutoFilterMode Then
            wsHistory.AutoFilterMode = False
        End If
        
        ' Apply AutoFilter to range from headers (row 2) to last data row
        Dim filterRange As Range
        Set filterRange = wsHistory.Range("A2:AB" & (outputRow - 1))
        filterRange.AutoFilter
    End If
    
    ' CRITICAL: Restore date format in B1:C1 after all operations
    ' This ensures dates display as dd.mm.yyyy, not as numbers
    wsHistory.Range("B1:C1").NumberFormat = "dd.mm.yyyy"
    
    Dim modeText As String
    If showOnlyLastChange Then
        modeText = "Modus B (Nur letzte Aenderung pro ID)"
    Else
        modeText = "Modus A (Alle Ereignisse)"
    End If
    
    MsgBox "Geschichte-Bericht erfolgreich generiert fuer Zeitraum " & Format(startDate, "dd.mm.yyyy") & " - " & Format(endDate, "dd.mm.yyyy") & vbCrLf & _
           "Modus: " & modeText & vbCrLf & _
           "Gesamtzahl Zeilen: " & (outputRow - 3) & vbCrLf & vbCrLf & _
           "Hinweis: Dies ist ein reiner Ansichtsbericht (Blatt 'Geschichte')." & vbCrLf & _
           "Fuer Pending-Entscheidungen verwenden Sie 'Ausstehende Aenderungen laden'.", vbInformation
    
Cleanup:
    ' Re-enable automatic calculations and screen updating
    Application.Calculation = xlCalculationAutomatic
    Application.ScreenUpdating = True
    If Err.Number <> 0 Then
        Debug.Print "Geschichte ERROR: row=" & currentRow & ", ID=" & recordID & ", Err=" & Err.Number & " - " & Err.Description
        MsgBox "Es ist ein Fehler aufgetreten: " & Err.Description & vbCrLf & _
               "Zeile: " & currentRow & ", ID: " & recordID & vbCrLf & _
               "Fehlernummer: " & Err.Number, vbCritical
    End If
End Sub

' Get or create a history worksheet by name (e.g., "Geschichte")
' This is a view-only sheet without decision columns
Private Function GetOrCreateHistorySheet(ByVal sheetName As String) As Worksheet
    Dim ws As Worksheet
    On Error Resume Next
    Set ws = ThisWorkbook.Worksheets(sheetName)
    On Error GoTo 0
    
    If ws Is Nothing Then
        Set ws = ThisWorkbook.Worksheets.Add(After:=ThisWorkbook.Worksheets(ThisWorkbook.Worksheets.Count))
        ws.Name = sheetName
        
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
    
    Set GetOrCreateHistorySheet = ws
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

' Create headers for view-only history sheet (NO decision columns AC/AD/AE)
Private Sub CreateHistoryHeaders(ws As Worksheet)
    ' Column structure: A-AB (28 columns) - NO decision columns
    ' A=FamilyID, B=Parent, C=Child, D=Birthdate, E=Address, F=Phone, G=Mobile, H=Email
    ' I=Subject1, J=Price1, K=Subject2, L=Price2, M-X=Months 1-12, Y-AA=Extra1-3
    ' AB=Comments (last column - no AC/AD/AE)
    
    Dim headers As Variant
    headers = Array("FamilyID", "Parent", "Child", "Birthdate", "Address", "Phone", "Mobile", "Email", _
                    "Subject1", "Price1", "Subject2", "Price2", _
                    "Jan", "Feb", "Mrz", "Apr", "Mai", "Jun", "Jul", "Aug", "Sep", "Okt", "Nov", "Dez", _
                    "Extra1", "Extra2", "Extra3", "Comments")
    
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
    ws.Columns("AB").ColumnWidth = 40 ' Comments (wider since it's the last column)
    
    ' NO columns AC/AD/AE - this is a view-only history report
    
    ' Set text format for columns that should not be auto-converted by Excel
    ' (Phone, Mobile, Birthdate, etc.)
    ' IMPORTANT: Apply text format only to data rows (2 and below), not to row 1 with dates
    ws.Range("A2:L" & ws.Rows.Count).NumberFormat = "@" ' Text format
    ws.Range("Y2:AB" & ws.Rows.Count).NumberFormat = "@" ' Text format for Extra1-3, Comments
    
    ' Ensure date format is preserved in B1:C1 (row 1 contains date inputs)
    ws.Range("B1:C1").NumberFormat = "dd.mm.yyyy"
    
    ' Months columns M-X will be formatted as numbers in CreateHistoryEntry
End Sub

' Create a single history entry (War/Ist rows) - view-only, no decision columns
Private Sub CreateHistoryEntry(wsHistory As Worksheet, wsKartei As Worksheet, _
                               startRow As Long, isRuck As Boolean, _
                               Reason As String, changeDate As String, _
                               Changes As Object, karteiRow As Long, operName As String, _
                               recordID As String)
    
    ' Column mapping (NO decision columns AC/AD/AE):
    ' A=FamilyID(1), B=Parent(2), C=Child(4), D=Birthdate(5), E=Address(6), F=Phone(7), G=Mobile(8), H=Email(9)
    ' I=Subject1(10), J=Price1(13), K=Subject2(15), L=Price2(18), M-X=Months(21-32), Y-AA=Extra(37-39)
    ' AB=Comments (last column)
    
    Dim rowWar As Long
    Dim rowIst As Long
    Dim rowSeparator As Long
    Dim i As Long
    
    rowWar = startRow
    rowIst = startRow + 1
    rowSeparator = startRow + 2
    
    ' Create separator row (A:AB only, no decision columns)
    wsHistory.Rows(rowSeparator).RowHeight = wsHistory.StandardHeight * 1 / 4
    wsHistory.Range("A" & rowSeparator & ":AB" & rowSeparator).Interior.Color = RGB(192, 192, 192)
    
    ' Fill basic data for both rows (War and Ist)
    For i = 0 To 1
        Dim currentRow As Long
        currentRow = startRow + i
        
        ' Store values as text to prevent Excel auto-conversion
        wsHistory.Range("A" & currentRow).Value = CStr(wsKartei.Cells(karteiRow, 1).Value) ' FamilyID from column A(1)
        wsHistory.Range("B" & currentRow).Value = CStr(wsKartei.Cells(karteiRow, 2).Value) ' Parent from column B(2)
        wsHistory.Range("C" & currentRow).Value = CStr(wsKartei.Cells(karteiRow, 4).Value) ' Child from column D(4)
        wsHistory.Range("D" & currentRow).Value = FormatAsText(wsKartei.Cells(karteiRow, 5).Value) ' Birthdate from column E(5)
        wsHistory.Range("E" & currentRow).Value = CStr(wsKartei.Cells(karteiRow, 6).Value) ' Address from column F(6)
        ' Phone columns (7=Tel., 8=Handy): use .Text to preserve leading zeros and prevent
        ' scientific notation. .Text returns the displayed string, not the internal value.
        wsHistory.Range("F" & currentRow).Value = wsKartei.Cells(karteiRow, 7).Text ' Phone from column G(7)
        wsHistory.Range("G" & currentRow).Value = wsKartei.Cells(karteiRow, 8).Text ' Mobile from column H(8)
        wsHistory.Range("H" & currentRow).Value = CStr(wsKartei.Cells(karteiRow, 9).Value) ' Email from column I(9)
        wsHistory.Range("I" & currentRow).Value = CStr(wsKartei.Cells(karteiRow, 10).Value) ' Subject1 from column J(10)
        wsHistory.Range("J" & currentRow).Value = CStr(wsKartei.Cells(karteiRow, 13).Value) ' Price1 from column M(13)
        wsHistory.Range("K" & currentRow).Value = CStr(wsKartei.Cells(karteiRow, 15).Value) ' Subject2 from column O(15)
        wsHistory.Range("L" & currentRow).Value = CStr(wsKartei.Cells(karteiRow, 18).Value) ' Price2 from column R(18)
        
        ' Extra subjects 1-3 from columns AK-AM (37-39)
        wsHistory.Range("Y" & currentRow).Value = CStr(wsKartei.Cells(karteiRow, 37).Value) ' Extra1
        wsHistory.Range("Z" & currentRow).Value = CStr(wsKartei.Cells(karteiRow, 38).Value) ' Extra2
        wsHistory.Range("AA" & currentRow).Value = CStr(wsKartei.Cells(karteiRow, 39).Value) ' Extra3
        
        ' NO RecordID column AE - this is a view-only history report
    Next i
    
    ' Fill Comments in Ist row (column AB) with appropriate formatting
    If isRuck Then
        wsHistory.Range("AB" & rowIst).Interior.Color = RGB(255, 153, 204) ' Pink for Ruck
    Else
        wsHistory.Range("AB" & rowIst).Interior.Color = RGB(204, 255, 153) ' Green for normal
    End If
    wsHistory.Range("AB" & rowIst).Value = Reason
    
    ' Highlight SEPA rows explicitly in comments (now column AB)
    Dim sepaMarker As String
    sepaMarker = Trim$(UCase$(CStr(wsKartei.Cells(karteiRow, 47).Value)))
    If sepaMarker = "SEPA" Then
        Dim sepaText As String
        sepaText = "SEPA"
        
        Dim commentCell As Range
        Set commentCell = wsHistory.Range("AB" & rowIst)
        
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
                        wsHistory.Cells(rowWar, colIndex).Value = 0
                    Else
                        warValue = ConvertDecimalSeparator(warValue, decimalSeparator)
                        If IsNumeric(warValue) Then
                            wsHistory.Cells(rowWar, colIndex).Value = CDbl(warValue)
                        Else
                            wsHistory.Cells(rowWar, colIndex).Value = 0
                        End If
                    End If
                    
                    ' Ist value (new state) with highlighting
                    Dim istValue As String
                    istValue = Changes(changeKey)("Ist")
                    If istValue = "" Or IsEmpty(istValue) Or IsNull(istValue) Then
                        wsHistory.Cells(rowIst, colIndex).Value = 0
                    Else
                        istValue = ConvertDecimalSeparator(istValue, decimalSeparator)
                        If IsNumeric(istValue) Then
                            wsHistory.Cells(rowIst, colIndex).Value = CDbl(istValue)
                        Else
                            wsHistory.Cells(rowIst, colIndex).Value = 0
                        End If
                    End If
                    
                    ' Highlight the changed cell
                    wsHistory.Cells(rowIst, colIndex).Interior.Color = RGB(255, 192, 203) ' Light pink
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
        wsHistory.Cells(rowWar, 5).Value = addressWar
        wsHistory.Cells(rowIst, 5).Value = addressIst
        wsHistory.Cells(rowIst, 5).Interior.Color = RGB(255, 192, 203)
    End If
    
    ' Column I = Subject1 (9)
    If subject1War <> "" Or subject1Ist <> "" Then
        wsHistory.Cells(rowWar, 9).Value = subject1War
        wsHistory.Cells(rowIst, 9).Value = subject1Ist
        wsHistory.Cells(rowIst, 9).Interior.Color = RGB(255, 192, 203)
    End If
    
    ' Column K = Subject2 (11)
    If subject2War <> "" Or subject2Ist <> "" Then
        wsHistory.Cells(rowWar, 11).Value = subject2War
        wsHistory.Cells(rowIst, 11).Value = subject2Ist
        wsHistory.Cells(rowIst, 11).Interior.Color = RGB(255, 192, 203)
    End If
    
    ' Column F = Phone (6)
    If phoneWar <> "" Or phoneIst <> "" Then
        wsHistory.Cells(rowWar, 6).Value = phoneWar
        wsHistory.Cells(rowIst, 6).Value = phoneIst
        wsHistory.Cells(rowIst, 6).Interior.Color = RGB(255, 192, 203)
    End If
    
    ' Column G = Mobile (7)
    If mobileWar <> "" Or mobileIst <> "" Then
        wsHistory.Cells(rowWar, 7).Value = mobileWar
        wsHistory.Cells(rowIst, 7).Value = mobileIst
        wsHistory.Cells(rowIst, 7).Interior.Color = RGB(255, 192, 203)
    End If
    
    ' Column H = Email (8)
    If emailWar <> "" Or emailIst <> "" Then
        wsHistory.Cells(rowWar, 8).Value = emailWar
        wsHistory.Cells(rowIst, 8).Value = emailIst
        wsHistory.Cells(rowIst, 8).Interior.Color = RGB(255, 192, 203)
    End If
    
    ' Column D = Birthdate (4)
    If birthdateWar <> "" Or birthdateIst <> "" Then
        wsHistory.Cells(rowWar, 4).Value = birthdateWar
        wsHistory.Cells(rowIst, 4).Value = birthdateIst
        wsHistory.Cells(rowIst, 4).Interior.Color = RGB(255, 192, 203)
    End If
    
    ' Column J = Price1 (10)
    If price1War <> "" Or price1Ist <> "" Then
        wsHistory.Cells(rowWar, 10).Value = price1War
        wsHistory.Cells(rowIst, 10).Value = price1Ist
        wsHistory.Cells(rowIst, 10).Interior.Color = RGB(255, 192, 203)
    End If
    
    ' Column L = Price2 (12)
    If price2War <> "" Or price2Ist <> "" Then
        wsHistory.Cells(rowWar, 12).Value = price2War
        wsHistory.Cells(rowIst, 12).Value = price2Ist
        wsHistory.Cells(rowIst, 12).Interior.Color = RGB(255, 192, 203)
    End If
    
    ' Column Y = Extra1 (25)
    If extra1War <> "" Or extra1Ist <> "" Then
        wsHistory.Cells(rowWar, 25).Value = extra1War
        wsHistory.Cells(rowIst, 25).Value = extra1Ist
        wsHistory.Cells(rowIst, 25).Interior.Color = RGB(255, 192, 203)
    End If
    
    ' Column Z = Extra2 (26)
    If extra2War <> "" Or extra2Ist <> "" Then
        wsHistory.Cells(rowWar, 26).Value = extra2War
        wsHistory.Cells(rowIst, 26).Value = extra2Ist
        wsHistory.Cells(rowIst, 26).Interior.Color = RGB(255, 192, 203)
    End If
    
    ' Column AA = Extra3 (27)
    If extra3War <> "" Or extra3Ist <> "" Then
        wsHistory.Cells(rowWar, 27).Value = extra3War
        wsHistory.Cells(rowIst, 27).Value = extra3Ist
        wsHistory.Cells(rowIst, 27).Interior.Color = RGB(255, 192, 203)
    End If
    
    ' Column A = FamilyID (1)
    If familyIDWar <> "" Or familyIDIst <> "" Then
        wsHistory.Cells(rowWar, 1).Value = familyIDWar
        wsHistory.Cells(rowIst, 1).Value = familyIDIst
        wsHistory.Cells(rowIst, 1).Interior.Color = RGB(255, 192, 203)
    End If
    
    ' Column B = Parent (2)
    If parentWar <> "" Or parentIst <> "" Then
        wsHistory.Cells(rowWar, 2).Value = parentWar
        wsHistory.Cells(rowIst, 2).Value = parentIst
        wsHistory.Cells(rowIst, 2).Interior.Color = RGB(255, 192, 203)
    End If
    
    ' Column C = Child (3)
    If childWar <> "" Or childIst <> "" Then
        wsHistory.Cells(rowWar, 3).Value = childWar
        wsHistory.Cells(rowIst, 3).Value = childIst
        wsHistory.Cells(rowIst, 3).Interior.Color = RGB(255, 192, 203)
    End If
    
    ' Format monthly columns M-X (13-24) as decimal with 2 places
    With wsHistory.Range("M" & rowWar & ":X" & rowIst)
        .NumberFormat = "0.00"
    End With
    
    ' Add borders to the full row range A-AB (no decision columns)
    With wsHistory.Range("A" & rowWar & ":AB" & rowIst).Borders
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

' Format history sheet headers and borders (view-only, no decision columns)
Private Sub FormatHistorySheet(ws As Worksheet, lastRow As Long)
    ' Add borders to headers (A2:AB2 - no decision columns)
    With ws.Range("A2:AB2").Borders
        .LineStyle = xlContinuous
        .Color = vbBlack
        .Weight = xlMedium
    End With
    
    ' Format headers background (A2:AB2 - no decision columns)
    ws.Range("A2:AB2").Interior.Color = RGB(220, 220, 220) ' Light gray
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
