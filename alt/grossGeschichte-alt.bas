Attribute VB_Name = "grossGeschichte"
'==========================
'   Module: grossGeschichte (Data file version)
'   Purpose: Generate history report for ALL records on Kartei sheet
'   View-only report with extended field support (new format compatible)
'
'   User interface:
'   - User sets date range in B1/C1 on Geschichte_Alle sheet
'   - Runs GrossGeshichteMachen macro
'   - System prompts for mode:
'       Mode A (JA): All history events for all IDs in date range
'       Mode B (NEIN): Only last change per ID in date range
'   - System creates/updates "Geschichte_Alle" sheet with history
'
'   Column Structure (A-AB) - same as Superadmin Geschichte:
'   A=FamilyID, B=Parent, C=Child, D=Birthdate, E=Address, F=Phone, G=Mobile, H=Email
'   I=Subject1, J=Price1, K=Subject2, L=Price2, M-X=Months 1-12, Y-AA=Extra1-3
'   AB=Comments
'
'   NO decision columns (AC/AD/AE) - this is a VIEW-ONLY report
'
'   Uses: History_ParseData.ParseHistory for parsing both legacy and new formats
'   SELF-CONTAINED for Data file - does not depend on Admin or Superadmin modules
'==========================

Option Explicit

Private Const SHEET_NAME As String = "Geschichte_Alle"
Private Const HISTORY_COL As Long = 52       ' AZ - history column in Kartei
Private Const ID_COL As Long = 48            ' AV - record ID column in Kartei
Private Const OPERATOR_COL As Long = 49      ' AW - operator column in Kartei
Private Const SEPA_COL As Long = 47          ' AU - SEPA marker column in Kartei

' Main entry point: Generate history report for all records
' Mode A (JA): All history events for all IDs in date range
' Mode B (NEIN): Only last change per ID in date range
' NOTE: This is a VIEW-ONLY report without decision columns
Sub GrossGeshichteMachen()
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
    Dim processedCount As Long
    
    On Error GoTo Cleanup
    
    ' Disable screen updating and automatic calculations for better performance
    Application.ScreenUpdating = False
    Application.Calculation = xlCalculationManual
    
    ' Get worksheets
    Set wsHistory = GetOrCreateHistorySheet()
    Set wsKartei = ThisWorkbook.Worksheets("Kartei")
    
    ' Read date range from B1 and C1
    If IsDate(wsHistory.Range("B1").Value) And IsDate(wsHistory.Range("C1").Value) Then
        startDate = CDate(wsHistory.Range("B1").Value)
        endDate = CDate(wsHistory.Range("C1").Value)
    Else
        MsgBox "Bitte gueltige Datumsangaben in den Zellen B1 und C1 eingeben", vbExclamation, "Geschichte Alle"
        GoTo Cleanup
    End If
    
    ' Ask user which mode to use
    userChoice = MsgBox("Berichtsmodus waehlen:" & vbCrLf & vbCrLf & _
                        "JA - Alle Verlaufsereignisse anzeigen (Modus A)" & vbCrLf & _
                        "NEIN - Nur letzte Aenderung pro ID anzeigen (Modus B)" & vbCrLf & vbCrLf & _
                        "Zeitraum: " & Format(startDate, "dd.mm.yyyy") & " - " & Format(endDate, "dd.mm.yyyy"), _
                        vbYesNoCancel + vbQuestion, "Geschichte Alle - Modus")
    
    If userChoice = vbCancel Then
        GoTo Cleanup
    End If
    
    showOnlyLastChange = (userChoice = vbNo)
    
    ' Clear content from row 2 onwards (preserve row 1 with date inputs)
    wsHistory.Rows("2:" & wsHistory.Rows.Count).Clear
    
    ' Create headers in row 2 (no decision columns - view-only report)
    Call CreateHistoryHeaders(wsHistory)
    
    ' Find last row in Kartei (column A)
    lastRow = wsKartei.Cells(wsKartei.Rows.Count, 1).End(xlUp).Row
    
    outputRow = 3 ' Start output from row 3
    processedCount = 0
    
    ' Loop through all rows in Kartei
    For currentRow = 3 To lastRow ' Start from row 3 (rows 1-2 are headers)
        
        ' Get history string, operator name, and ID
        strGeschichte = wsKartei.Cells(currentRow, HISTORY_COL).Value
        operName = wsKartei.Cells(currentRow, OPERATOR_COL).Value
        recordID = CStr(wsKartei.Cells(currentRow, ID_COL).Value)
        
        ' Skip if no history or no ID
        If strGeschichte <> "" And recordID <> "" Then
            
            ' Split raw history into segments for field change parsing
            segments = Split(strGeschichte, "||")
            
            ' Parse history using History_ParseData module
            Set result = History_ParseData.ParseHistory(strGeschichte)
            
            If result Is Nothing Then GoTo NextRow
            If result.Count = 0 Then GoTo NextRow
            
            ' If Mode B, find the last event in date range for this ID
            If showOnlyLastChange Then
                Set lastEvent = FindLastEventInRange(result, segments, startDate, endDate)
                
                If Not lastEvent Is Nothing Then
                    ' Create entry for the last event only
                    Call CreateHistoryEntry(wsHistory, wsKartei, outputRow, _
                                           lastEvent("IsRuck"), lastEvent("Reason"), _
                                           CStr(lastEvent("ChangeDate")), lastEvent("Changes"), _
                                           currentRow, operName, recordID)
                    outputRow = outputRow + 3 ' Each entry takes 2 rows + 1 separator
                    processedCount = processedCount + 1
                End If
            Else
                ' Mode A: Process each event in the history
                For i = 1 To result.Count
                    Set evt = result(i)
                    
                    ' Enrich Changes with field changes from raw segment
                    segmentText = ""
                    On Error Resume Next
                    If (i - 1) >= LBound(segments) And (i - 1) <= UBound(segments) Then
                        segmentText = Trim$(segments(i - 1))
                    End If
                    On Error GoTo Cleanup
                    
                    Set fieldChanges = ParseFieldChangesFromSegment(segmentText)
                    Call MergeFieldChangesIntoChanges(evt("Changes"), fieldChanges)
                    
                    ' Check if event date falls within the specified range
                    On Error Resume Next
                    eventDate = CDate(evt("ChangeDate"))
                    If Err.Number = 0 Then
                        If eventDate >= startDate And eventDate <= endDate Then
                            ' Create entry for this event
                            Call CreateHistoryEntry(wsHistory, wsKartei, outputRow, _
                                                   evt("IsRuck"), evt("Reason"), _
                                                   CStr(evt("ChangeDate")), evt("Changes"), _
                                                   currentRow, operName, recordID)
                            outputRow = outputRow + 3
                            processedCount = processedCount + 1
                        End If
                    End If
                    Err.Clear
                    On Error GoTo Cleanup
                Next i
            End If
        End If
NextRow:
    Next currentRow
    
    ' Apply final formatting
    Call FormatHistorySheet(wsHistory, outputRow - 1)
    
    ' Apply AutoFilter if we have data
    If outputRow > 3 Then
        If wsHistory.AutoFilterMode Then
            wsHistory.AutoFilterMode = False
        End If
        wsHistory.Range("A2:AB" & (outputRow - 1)).AutoFilter
    End If
    
    ' Ensure date format in B1:C1
    wsHistory.Range("B1:C1").NumberFormat = "dd.mm.yyyy"
    
    ' Activate history sheet
    wsHistory.Activate
    
    ' Show summary
    Dim modeText As String
    If showOnlyLastChange Then
        modeText = "Modus B (Nur letzte Aenderung pro ID)"
    Else
        modeText = "Modus A (Alle Ereignisse)"
    End If
    
    MsgBox "Geschichte-Bericht erfolgreich generiert." & vbCrLf & vbCrLf & _
           "Zeitraum: " & Format(startDate, "dd.mm.yyyy") & " - " & Format(endDate, "dd.mm.yyyy") & vbCrLf & _
           "Modus: " & modeText & vbCrLf & _
           "Anzahl Eintraege: " & processedCount, vbInformation, "Geschichte Alle"
    
Cleanup:
    Application.Calculation = xlCalculationAutomatic
    Application.ScreenUpdating = True
    
    If Err.Number <> 0 Then
        MsgBox "Es ist ein Fehler aufgetreten: " & Err.Description & vbCrLf & _
               "Zeile: " & currentRow & ", ID: " & recordID, vbCritical, "Geschichte Alle - Fehler"
    End If
End Sub

' ========================================
' Sheet Management
' ========================================

' Get or create the history sheet
Private Function GetOrCreateHistorySheet() As Worksheet
    Dim ws As Worksheet
    
    On Error Resume Next
    Set ws = ThisWorkbook.Worksheets(SHEET_NAME)
    On Error GoTo 0
    
    If ws Is Nothing Then
        ' Create new sheet after Kartei
        Dim wsKartei As Worksheet
        On Error Resume Next
        Set wsKartei = ThisWorkbook.Worksheets("Kartei")
        On Error GoTo 0
        
        If wsKartei Is Nothing Then
            Set ws = ThisWorkbook.Worksheets.Add(After:=ThisWorkbook.Worksheets(ThisWorkbook.Worksheets.Count))
        Else
            Set ws = ThisWorkbook.Worksheets.Add(After:=wsKartei)
        End If
        ws.Name = SHEET_NAME
        
        ' Add date range inputs
        ws.Range("A1").Value = "Zeitraum:"
        ws.Range("A1").Font.Bold = True
        ws.Range("B1:C1").NumberFormat = "dd.mm.yyyy"
        ws.Range("B1").Value = Date - 30 ' Default: last 30 days
        ws.Range("C1").Value = Date
    Else
        ' Sheet exists - ensure date format is correct
        ws.Range("B1:C1").NumberFormat = "dd.mm.yyyy"
        
        ' Set default dates if not valid
        If Not IsDate(ws.Range("B1").Value) Or IsEmpty(ws.Range("B1").Value) Then
            ws.Range("B1").Value = Date - 30
        End If
        If Not IsDate(ws.Range("C1").Value) Or IsEmpty(ws.Range("C1").Value) Then
            ws.Range("C1").Value = Date
        End If
    End If
    
    Set GetOrCreateHistorySheet = ws
End Function

' ========================================
' Headers and Formatting
' ========================================

' Create headers for view-only history sheet (columns A-AB, NO decision columns)
Private Sub CreateHistoryHeaders(ws As Worksheet)
    ' Column structure: A-AB (28 columns) - NO decision columns AC/AD/AE
    ' A=FamilyID, B=Parent, C=Child, D=Birthdate, E=Address, F=Phone, G=Mobile, H=Email
    ' I=Subject1, J=Price1, K=Subject2, L=Price2, M-X=Months 1-12, Y-AA=Extra1-3
    ' AB=Comments (last column)
    
    Dim headers As Variant
    headers = Array("FamilyID", "Eltern", "Kind", "Geburtsdatum", "Adresse", "Telefon", "Mobil", "Email", _
                    "Fach1", "Preis1", "Fach2", "Preis2", _
                    "Jan", "Feb", "Mrz", "Apr", "Mai", "Jun", "Jul", "Aug", "Sep", "Okt", "Nov", "Dez", _
                    "Extra1", "Extra2", "Extra3", "Kommentar")
    
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
    ws.Columns("B").ColumnWidth = 20  ' Eltern
    ws.Columns("C").ColumnWidth = 18  ' Kind
    ws.Columns("D").ColumnWidth = 12  ' Geburtsdatum
    ws.Columns("E").ColumnWidth = 25  ' Adresse
    ws.Columns("F").ColumnWidth = 14  ' Telefon
    ws.Columns("G").ColumnWidth = 14  ' Mobil
    ws.Columns("H").ColumnWidth = 22  ' Email
    ws.Columns("I").ColumnWidth = 18  ' Fach1
    ws.Columns("J").ColumnWidth = 8   ' Preis1
    ws.Columns("K").ColumnWidth = 18  ' Fach2
    ws.Columns("L").ColumnWidth = 8   ' Preis2
    ws.Columns("M:X").ColumnWidth = 6 ' Months 1-12
    ws.Columns("Y:AA").ColumnWidth = 15 ' Extra1-3
    ws.Columns("AB").ColumnWidth = 40 ' Kommentar
    
    ' Set text format for non-numeric columns (to prevent auto-conversion)
    ' Apply only to data rows (row 2 and below), not to row 1 with dates
    ws.Range("A2:L" & ws.Rows.Count).NumberFormat = "@"
    ws.Range("Y2:AB" & ws.Rows.Count).NumberFormat = "@"
    
    ' Ensure date format is preserved in B1:C1
    ws.Range("B1:C1").NumberFormat = "dd.mm.yyyy"
End Sub

' Apply final formatting to history sheet
Private Sub FormatHistorySheet(ws As Worksheet, lastRow As Long)
    If lastRow < 3 Then Exit Sub
    
    ' Format header row with borders
    With ws.Range("A2:AB2").Borders
        .LineStyle = xlContinuous
        .Color = vbBlack
        .Weight = xlMedium
    End With
    
    ' Header background
    ws.Range("A2:AB2").Interior.Color = RGB(220, 220, 220)
End Sub

' ========================================
' History Entry Creation
' ========================================

' Create a single history entry (War/Ist rows + separator)
' View-only: no decision columns AC/AD/AE
Private Sub CreateHistoryEntry(wsHistory As Worksheet, wsKartei As Worksheet, _
                               startRow As Long, isRuck As Boolean, _
                               reason As String, changeDate As String, _
                               Changes As Object, karteiRow As Long, _
                               operName As String, recordID As String)
    
    ' Kartei column mapping:
    ' A(1)=FamilyID, B(2)=Parent, D(4)=Child, E(5)=Birthdate, F(6)=Address
    ' G(7)=Phone, H(8)=Mobile, I(9)=Email, J(10)=Subject1, M(13)=Price1
    ' O(15)=Subject2, R(18)=Price2, U-AF(21-32)=Months, AK-AM(37-39)=Extra1-3
    
    Dim rowWar As Long, rowIst As Long, rowSeparator As Long
    Dim i As Long
    Dim decimalSeparator As String
    
    rowWar = startRow
    rowIst = startRow + 1
    rowSeparator = startRow + 2
    
    decimalSeparator = Application.International(xlDecimalSeparator)
    
    ' Create separator row (thin gray line)
    wsHistory.Rows(rowSeparator).RowHeight = wsHistory.StandardHeight * 0.25
    wsHistory.Range("A" & rowSeparator & ":AB" & rowSeparator).Interior.Color = RGB(192, 192, 192)
    
    ' Fill current values for both rows (from Kartei)
    For i = 0 To 1
        Dim currentRow As Long
        currentRow = startRow + i
        
        wsHistory.Range("A" & currentRow).Value = CStr(wsKartei.Cells(karteiRow, 1).Value)   ' FamilyID
        wsHistory.Range("B" & currentRow).Value = CStr(wsKartei.Cells(karteiRow, 2).Value)   ' Parent
        wsHistory.Range("C" & currentRow).Value = CStr(wsKartei.Cells(karteiRow, 4).Value)   ' Child
        wsHistory.Range("D" & currentRow).Value = FormatDateAsText(wsKartei.Cells(karteiRow, 5).Value) ' Birthdate
        wsHistory.Range("E" & currentRow).Value = CStr(wsKartei.Cells(karteiRow, 6).Value)   ' Address
        wsHistory.Range("F" & currentRow).Value = FormatAsText(wsKartei.Cells(karteiRow, 7).Value)  ' Phone
        wsHistory.Range("G" & currentRow).Value = FormatAsText(wsKartei.Cells(karteiRow, 8).Value)  ' Mobile
        wsHistory.Range("H" & currentRow).Value = CStr(wsKartei.Cells(karteiRow, 9).Value)   ' Email
        wsHistory.Range("I" & currentRow).Value = CStr(wsKartei.Cells(karteiRow, 10).Value)  ' Subject1
        wsHistory.Range("J" & currentRow).Value = CStr(wsKartei.Cells(karteiRow, 13).Value)  ' Price1
        wsHistory.Range("K" & currentRow).Value = CStr(wsKartei.Cells(karteiRow, 15).Value)  ' Subject2
        wsHistory.Range("L" & currentRow).Value = CStr(wsKartei.Cells(karteiRow, 18).Value)  ' Price2
        
        ' Extra subjects
        wsHistory.Range("Y" & currentRow).Value = CStr(wsKartei.Cells(karteiRow, 37).Value)  ' Extra1
        wsHistory.Range("Z" & currentRow).Value = CStr(wsKartei.Cells(karteiRow, 38).Value)  ' Extra2
        wsHistory.Range("AA" & currentRow).Value = CStr(wsKartei.Cells(karteiRow, 39).Value) ' Extra3
    Next i
    
    ' Fill comment cell (AB) with Ruck highlighting
    If isRuck Then
        wsHistory.Range("AB" & rowIst).Interior.Color = RGB(255, 153, 204) ' Pink for Ruck
    Else
        wsHistory.Range("AB" & rowIst).Interior.Color = RGB(204, 255, 153) ' Green for normal
    End If
    
    ' Build comment text: date + operator + reason
    Dim commentText As String
    commentText = changeDate
    If Len(operName) > 0 Then
        commentText = commentText & " [" & operName & "]"
    End If
    If Len(reason) > 0 Then
        commentText = commentText & " - " & reason
    End If
    wsHistory.Range("AB" & rowIst).Value = commentText
    
    ' Add SEPA marker if applicable
    Dim sepaMarker As String
    sepaMarker = Trim$(UCase$(CStr(wsKartei.Cells(karteiRow, SEPA_COL).Value)))
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
        
        ' Highlight SEPA text in red bold
        With commentCell.Characters(startPos, Len(sepaText)).Font
            .Color = vbRed
            .Bold = True
        End With
    End If
    
    ' Process changes from history - apply to War/Ist rows
    Call ApplyChangesToRows(wsHistory, Changes, rowWar, rowIst, decimalSeparator)
    
    ' Format month columns as numbers with 2 decimals
    With wsHistory.Range("M" & rowWar & ":X" & rowIst)
        .NumberFormat = "0.00"
    End With
    
    ' Add borders to the entry
    With wsHistory.Range("A" & rowWar & ":AB" & rowIst).Borders
        .LineStyle = xlContinuous
        .Color = vbBlack
        .Weight = xlThin
    End With
End Sub

' Apply field changes to War/Ist rows with highlighting
Private Sub ApplyChangesToRows(wsHistory As Worksheet, Changes As Object, _
                               rowWar As Long, rowIst As Long, decimalSeparator As String)
    
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
    
    On Error Resume Next ' Handle potential errors with dictionary access
    
    Dim changeKey As Variant
    For Each changeKey In Changes.Keys
        Dim keyStr As String
        keyStr = CStr(changeKey)
        
        ' Check if it's a month number (1-12)
        If IsNumeric(keyStr) Then
            Dim monthNum As Long
            monthNum = CLng(keyStr)
            
            If monthNum >= 1 And monthNum <= 12 Then
                ' Months 1-12 are in columns M-X (13-24)
                Dim colIndex As Long
                colIndex = 12 + monthNum ' M=13 for month 1, X=24 for month 12
                
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
        Else
            ' Non-numeric key - capture values for dedicated columns
            Dim warText As String
            Dim istText As String
            
            warText = Changes(changeKey)("War")
            istText = Changes(changeKey)("Ist")
            
            Select Case keyStr
                Case "Address", "ADR"
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
    
    On Error GoTo 0
    
    ' Apply field changes to dedicated columns with highlighting
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
    
    ' Column D = Birthdate (4)
    If birthdateWar <> "" Or birthdateIst <> "" Then
        wsHistory.Cells(rowWar, 4).Value = birthdateWar
        wsHistory.Cells(rowIst, 4).Value = birthdateIst
        wsHistory.Cells(rowIst, 4).Interior.Color = RGB(255, 192, 203)
    End If
    
    ' Column E = Address (5)
    If addressWar <> "" Or addressIst <> "" Then
        wsHistory.Cells(rowWar, 5).Value = addressWar
        wsHistory.Cells(rowIst, 5).Value = addressIst
        wsHistory.Cells(rowIst, 5).Interior.Color = RGB(255, 192, 203)
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
    
    ' Column I = Subject1 (9)
    If subject1War <> "" Or subject1Ist <> "" Then
        wsHistory.Cells(rowWar, 9).Value = subject1War
        wsHistory.Cells(rowIst, 9).Value = subject1Ist
        wsHistory.Cells(rowIst, 9).Interior.Color = RGB(255, 192, 203)
    End If
    
    ' Column J = Price1 (10)
    If price1War <> "" Or price1Ist <> "" Then
        wsHistory.Cells(rowWar, 10).Value = price1War
        wsHistory.Cells(rowIst, 10).Value = price1Ist
        wsHistory.Cells(rowIst, 10).Interior.Color = RGB(255, 192, 203)
    End If
    
    ' Column K = Subject2 (11)
    If subject2War <> "" Or subject2Ist <> "" Then
        wsHistory.Cells(rowWar, 11).Value = subject2War
        wsHistory.Cells(rowIst, 11).Value = subject2Ist
        wsHistory.Cells(rowIst, 11).Interior.Color = RGB(255, 192, 203)
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
End Sub

' ========================================
' Helper Functions
' ========================================

' Convert decimal separator based on system settings
Private Function ConvertDecimalSeparator(ByVal Value As String, ByVal systemDecimalSeparator As String) As String
    ' Remove extra spaces
    Value = Replace(Value, " ", "")
    
    ' Unify decimal separators so Excel can parse them correctly
    If systemDecimalSeparator = "," Then
        Value = Replace(Value, ".", ",")
    Else
        Value = Replace(Value, ",", ".")
    End If
    
    ConvertDecimalSeparator = Value
End Function

' Format value as text to prevent Excel auto-conversion
Private Function FormatAsText(ByVal Value As Variant) As String
    If IsEmpty(Value) Or IsNull(Value) Then
        FormatAsText = ""
    Else
        FormatAsText = "'" & CStr(Value)
    End If
End Function

' Format date as text in dd.mm.yyyy format
Private Function FormatDateAsText(ByVal Value As Variant) As String
    If IsEmpty(Value) Or IsNull(Value) Then
        FormatDateAsText = ""
    ElseIf IsDate(Value) Then
        FormatDateAsText = Format(Value, "dd.mm.yyyy")
    Else
        FormatDateAsText = CStr(Value)
    End If
End Function

' ========================================
' Field Change Parsing (from raw segments)
' ========================================

' Parse Address/Subject1/Subject2 changes from a raw history segment
' Supports both formats:
'   - Legacy: Address: Was(X); Is(Y). Subject1: Was(X); Is(Y). Subject2: Was(X); Is(Y).
'   - New: ADR(X->Y) SB1(X->Y) SB2(X->Y)
Private Function ParseFieldChangesFromSegment(ByVal segment As String) As Object
    On Error GoTo ErrHandler
    
    Dim result As Object
    Set result = CreateObject("Scripting.Dictionary")
    
    If Len(Trim$(segment)) = 0 Then
        Set ParseFieldChangesFromSegment = result
        Exit Function
    End If
    
    ' Try legacy format: FieldName: Was(...); Is(...).
    Call AddFieldChangeToDict(segment, "Address", result)
    Call AddFieldChangeToDict(segment, "Subject1", result)
    Call AddFieldChangeToDict(segment, "Subject2", result)
    
    ' Try new format: TAG(X->Y)
    Call AddNewFormatFieldChangeToDict(segment, "ADR", "Address", result)
    Call AddNewFormatFieldChangeToDict(segment, "SB1", "Subject1", result)
    Call AddNewFormatFieldChangeToDict(segment, "SB2", "Subject2", result)
    
    Set ParseFieldChangesFromSegment = result
    Exit Function
    
ErrHandler:
    Debug.Print "ERROR in ParseFieldChangesFromSegment: " & Err.Description
    Set ParseFieldChangesFromSegment = CreateObject("Scripting.Dictionary")
End Function

' Merge additional field changes into the base Changes dictionary
Private Sub MergeFieldChangesIntoChanges(ByVal baseChanges As Object, ByVal fieldChanges As Object)
    On Error GoTo ErrHandler
    
    If baseChanges Is Nothing Then Exit Sub
    If fieldChanges Is Nothing Then Exit Sub
    
    Dim key As Variant
    For Each key In fieldChanges.Keys
        If Not baseChanges.Exists(key) Then
            baseChanges.Add key, fieldChanges.Item(key)
        Else
            ' Update existing key
            Set baseChanges.Item(key) = fieldChanges.Item(key)
        End If
    Next key
    
    Exit Sub
    
ErrHandler:
    Debug.Print "ERROR in MergeFieldChangesIntoChanges: " & Err.Description
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
        Set dict(fieldName) = fieldDict
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

' ========================================
' Mode B: Find Last Event in Date Range
' ========================================

' Find the last event in date range for a given ID
' Returns the event with the latest date, and if multiple events have the same date,
' returns the last one in the collection (most recent in order)
Private Function FindLastEventInRange(ByVal events As Collection, ByRef segments() As String, _
                                      ByVal startDate As Date, ByVal endDate As Date) As Object
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
        
        ' Enrich this event with field changes from raw segment
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

