Attribute VB_Name = "valid_HistoryPerYear"
'==========================
'   Module: valid_HistoryPerYear
'   Purpose: Year-aware history report generation for Superadmin multi-year workbook
'
'   Provides:
'   - GeschichteMachenForYear(year2): Single-record history for selected row on KarteiYY -> GeschichteYY
'   - GrossGeschichteMachenForYear(year2): Date-range history for all records (DB tblKartei or KarteiYY) -> GeschichteYY_Alle
'
'   Wrapper macros for UI:
'   - GeschichteMachen24/25/26: Single-record history
'   - GrossGeschichteMachen24/25/26: Date-range history for all records
'
'   Data Source:
'   - Reads from KarteiYY sheets (populated by valid_ImportPending from pre_tblKartei)
'   - For full database history, GrossGeschichteMachenForYear supports a direct Access read of tblKartei
'
'   Uses: valid_ParseHistory.ParseHistory for parsing legacy + new history formats
'   Uses: valid_YearConfig for sheet name conventions and year validation
'==========================

Option Explicit

' Column indices in Kartei sheets
Private Const HISTORY_COL As Long = 52       ' AZ - history string column
Private Const ID_COL As Long = 48            ' AV - record ID column
Private Const OPERATOR_COL As Long = 49      ' AW - operator name column
Private Const SEPA_COL As Long = 47          ' AU - SEPA marker column

' ============================================================
' WRAPPER MACROS - Single-Record History (GeschichteYY)
' ============================================================

' Single-record history for year 2024
Public Sub GeschichteMachen24()
    GeschichteMachenForYear 24
End Sub

' Single-record history for year 2025
Public Sub GeschichteMachen25()
    GeschichteMachenForYear 25
End Sub

' Single-record history for year 2026
Public Sub GeschichteMachen26()
    GeschichteMachenForYear 26
End Sub

' ============================================================
' WRAPPER MACROS - Date-Range History for All Records (GeschichteYY_Alle)
' ============================================================

' Date-range history for all records - year 2024
Public Sub GrossGeschichteMachen24()
    GrossGeschichteMachenForYear 24
End Sub

' Date-range history for all records - year 2025
Public Sub GrossGeschichteMachen25()
    GrossGeschichteMachenForYear 25
End Sub

' Date-range history for all records - year 2026
Public Sub GrossGeschichteMachen26()
    GrossGeschichteMachenForYear 26
End Sub

' ============================================================
' MAIN IMPLEMENTATION - Single-Record History
' ============================================================

' Generate history report for a single selected record on KarteiYY sheet
' User selects any cell in a row on KarteiYY, then runs this macro
' Output goes to GeschichteYY sheet
'
' @param year2 - Two-digit year (24, 25, or 26)
Public Sub GeschichteMachenForYear(ByVal year2 As Integer)
    Dim wsKartei As Worksheet
    Dim wsHistory As Worksheet
    Dim karteiSheetName As String
    Dim historySheetName As String
    Dim currentRow As Long
    Dim recordID As String
    Dim strGeschichte As String
    Dim operName As String
    Dim result As Collection
    Dim startDate As Date
    Dim endDate As Date
    Dim useDateFilter As Boolean
    Dim outputRow As Long
    Dim evt As Object
    Dim eventDate As Date
    Dim i As Long
    Dim segments() As String
    Dim segmentText As String
    Dim fieldChanges As Object
    
    On Error GoTo Cleanup
    
    ' Validate year
    If Not IsValidYear(year2) Then
        MsgBox "Ungueltiges Jahr: " & year2 & ". Unterstuetzte Jahre: 24, 25, 26.", _
               vbExclamation, "Geschichte - Jahresfehler"
        Exit Sub
    End If
    
    ' Get sheet names for this year
    karteiSheetName = valid_YearConfig.GetKarteiSheetName(year2)
    historySheetName = valid_YearConfig.GetGeschichteSheetName(year2)
    
    ' Validate active sheet is the correct Kartei sheet
    If ActiveSheet.Name <> karteiSheetName Then
        MsgBox "Bitte waehlen Sie zuerst eine Zeile auf dem Blatt '" & karteiSheetName & "' aus.", _
               vbExclamation, "Geschichte " & Format(year2, "00")
        Exit Sub
    End If
    
    Set wsKartei = ActiveSheet
    currentRow = ActiveCell.Row
    
    ' Validate row selection (skip header rows)
    If currentRow < 3 Then
        MsgBox "Bitte waehlen Sie eine Datenzeile aus (Zeile 3 oder hoeher).", _
               vbExclamation, "Geschichte " & Format(year2, "00")
        Exit Sub
    End If
    
    ' Get record ID, operator, and history
    recordID = CStr(wsKartei.Cells(currentRow, ID_COL).Value)
    operName = CStr(wsKartei.Cells(currentRow, OPERATOR_COL).Value)
    strGeschichte = wsKartei.Cells(currentRow, HISTORY_COL).Value
    
    If recordID = "" Then
        MsgBox "Die ausgewaehlte Zeile hat keine gueltige ID.", _
               vbExclamation, "Geschichte " & Format(year2, "00")
        Exit Sub
    End If
    
    If strGeschichte = "" Then
        MsgBox "Fuer diesen Datensatz (ID: " & recordID & ") wurde keine Aenderungshistorie gefunden.", _
               vbInformation, "Geschichte " & Format(year2, "00")
        Exit Sub
    End If
    
    ' Disable screen updating for performance
    Application.ScreenUpdating = False
    Application.Calculation = xlCalculationManual
    
    ' Get or create history sheet for this year
    Set wsHistory = GetOrCreateSingleHistorySheet(historySheetName)
    
    ' Check if date filter should be applied
    useDateFilter = False
    If IsDate(wsHistory.Range("B1").Value) And IsDate(wsHistory.Range("C1").Value) Then
        startDate = CDate(wsHistory.Range("B1").Value)
        endDate = CDate(wsHistory.Range("C1").Value)
        useDateFilter = True
    End If
    
    ' Clear previous content (keep row 1 with date inputs and row 2 with headers)
    wsHistory.Rows("3:" & wsHistory.Rows.Count).Clear
    
    ' Create headers
    Call CreateSingleHistoryHeaders(wsHistory)
    
    ' Display current record info in D1
    wsHistory.Range("D1").Value = "ID: " & recordID
    wsHistory.Range("D1").Font.Bold = True
    
    ' Split history into segments for field change parsing
    segments = Split(strGeschichte, "||")
    
    ' Parse history using valid_ParseHistory module
    Set result = valid_ParseHistory.ParseHistory(strGeschichte)
    
    If result.Count = 0 Then
        MsgBox "Die Historie konnte nicht geparst werden oder enthaelt keine gueltigen Ereignisse.", _
               vbInformation, "Geschichte " & Format(year2, "00")
        GoTo Cleanup
    End If
    
    outputRow = 3 ' Start output from row 3
    
    ' Process each event
    For i = 1 To result.Count
        Set evt = result(i)
        
        ' Enrich Changes with field changes from raw segment
        segmentText = ""
        If (i - 1) >= LBound(segments) And (i - 1) <= UBound(segments) Then
            segmentText = Trim$(segments(i - 1))
        End If
        
        Set fieldChanges = ParseFieldChangesFromSegment(segmentText)
        Call MergeFieldChangesIntoChanges(evt("Changes"), fieldChanges)
        
        ' Apply date filter if enabled
        If useDateFilter Then
            On Error Resume Next
            eventDate = CDate(evt("ChangeDate"))
            If Err.Number = 0 Then
                If eventDate < startDate Or eventDate > endDate Then
                    Err.Clear
                    GoTo NextSingleEvent
                End If
            Else
                Err.Clear
                GoTo NextSingleEvent
            End If
            On Error GoTo Cleanup
        End If
        
        ' Create history entry
        Call CreateSingleHistoryEntry(wsHistory, wsKartei, outputRow, _
                               evt("IsRuck"), evt("Reason"), CStr(evt("ChangeDate")), _
                               evt("Changes"), currentRow, operName, recordID)
        outputRow = outputRow + 3 ' Each entry: War + Ist + Separator
        
NextSingleEvent:
    Next i
    
    ' Apply final formatting
    Call FormatSingleHistorySheet(wsHistory, outputRow - 1)
    
    ' Apply AutoFilter if we have data
    If outputRow > 3 Then
        If wsHistory.AutoFilterMode Then
            wsHistory.AutoFilterMode = False
        End If
        wsHistory.Range("A2:T" & (outputRow - 1)).AutoFilter
    End If
    
    ' Ensure date format in B1:C1
    wsHistory.Range("B1:C1").NumberFormat = "dd.mm.yyyy"
    
    ' Activate history sheet
    wsHistory.Activate
    
    ' Show summary
    Dim filterInfo As String
    If useDateFilter Then
        filterInfo = vbCrLf & "Zeitraum: " & Format(startDate, "dd.mm.yyyy") & " - " & Format(endDate, "dd.mm.yyyy")
    Else
        filterInfo = vbCrLf & "(Kein Datumsfilter - alle Ereignisse angezeigt)"
    End If
    
    MsgBox "Geschichte fuer ID " & recordID & " erfolgreich generiert." & filterInfo & vbCrLf & _
           "Anzahl Ereignisse: " & ((outputRow - 3) / 3) & vbCrLf & _
           "Ausgabe: " & historySheetName, vbInformation, "Geschichte " & Format(year2, "00")
    
Cleanup:
    Application.Calculation = xlCalculationAutomatic
    Application.ScreenUpdating = True
    
    If Err.Number <> 0 Then
        MsgBox "Es ist ein Fehler aufgetreten: " & Err.Description, _
               vbCritical, "Geschichte " & Format(year2, "00") & " - Fehler"
    End If
End Sub

' ============================================================
' MAIN IMPLEMENTATION - Date-Range History for All Records
' ============================================================

' Generate date-range history report for all records (either from tblKartei (DB) or KarteiYY)
' Mode A (JA): All history events for all IDs in date range
' Mode B (NEIN): Only last change per ID in date range
' Output goes to GeschichteYY_Alle sheet
'
' NOTE:
' - In Superadmin, KarteiYY typically contains only pending records (from pre_tblKartei).
' - This macro can also read the full year database directly (tblKartei) to cover all records.
'
' @param year2 - Two-digit year (24, 25, or 26)
Public Sub GrossGeschichteMachenForYear(ByVal year2 As Integer)
    Dim wsHistory As Worksheet
    Dim wsKartei As Worksheet
    Dim karteiSheetName As String
    Dim historySheetName As String
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
    Dim sourceChoice As VbMsgBoxResult
    Dim useDatabase As Boolean
    Dim dbPath As String
    Dim db As DAO.Database
    Dim rs As DAO.Recordset
    Dim recordValues As Variant
    Dim segments() As String
    Dim lastEvent As Object
    Dim i As Long
    Dim evt As Object
    Dim segmentText As String
    Dim fieldChanges As Object
    Dim eventDate As Date
    Dim processedCount As Long
    
    On Error GoTo Cleanup
    
    ' Validate year
    If Not IsValidYear(year2) Then
        MsgBox "Ungueltiges Jahr: " & year2 & ". Unterstuetzte Jahre: 24, 25, 26.", _
               vbExclamation, "Geschichte Alle - Jahresfehler"
        Exit Sub
    End If
    
    ' Get sheet names for this year
    karteiSheetName = valid_YearConfig.GetKarteiSheetName(year2)
    historySheetName = valid_YearConfig.GetGeschichteAlleSheetName(year2)
    
    ' Disable screen updating for performance
    Application.ScreenUpdating = False
    Application.Calculation = xlCalculationManual
    
    ' Get output worksheet
    Set wsHistory = GetOrCreateAlleHistorySheet(historySheetName)
    
    ' Read date range from B1 and C1
    If IsDate(wsHistory.Range("B1").Value) And IsDate(wsHistory.Range("C1").Value) Then
        startDate = CDate(wsHistory.Range("B1").Value)
        endDate = CDate(wsHistory.Range("C1").Value)
    Else
        Application.Calculation = xlCalculationAutomatic
        Application.ScreenUpdating = True
        MsgBox "Bitte gueltige Datumsangaben in den Zellen B1 und C1 des Blattes '" & historySheetName & "' eingeben", _
               vbExclamation, "Geschichte Alle " & Format(year2, "00")
        Exit Sub
    End If
    
    ' Ask user which data source to use
    Application.ScreenUpdating = True ' Temporarily show screen for MsgBox
    sourceChoice = MsgBox("Datenquelle waehlen:" & vbCrLf & vbCrLf & _
                          "JA - Vollstaendig aus Datenbank (tblKartei)" & vbCrLf & _
                          "NEIN - Nur aus '" & karteiSheetName & "' (pending subset)" & vbCrLf & vbCrLf & _
                          "Zeitraum: " & Format(startDate, "dd.mm.yyyy") & " - " & Format(endDate, "dd.mm.yyyy") & vbCrLf & vbCrLf & _
                          "Hinweis:" & vbCrLf & _
                          "  - '" & karteiSheetName & "' enthaelt in Superadmin typischerweise nur pre_tblKartei (pending)." & vbCrLf & _
                          "  - Datenbankmodus liest direkt tblKartei.Value52 (Historie) fuer ALLE Datensaetze.", _
                          vbYesNoCancel + vbQuestion, "Geschichte Alle " & Format(year2, "00") & " - Datenquelle")
    Application.ScreenUpdating = False
    
    If sourceChoice = vbCancel Then
        GoTo Cleanup
    End If
    
    useDatabase = (sourceChoice = vbYes)
    
    ' Ask user which mode to use
    Application.ScreenUpdating = True ' Temporarily show screen for MsgBox
    userChoice = MsgBox("Berichtsmodus waehlen:" & vbCrLf & vbCrLf & _
                        "JA - Alle Verlaufsereignisse anzeigen (Modus A)" & vbCrLf & _
                        "NEIN - Nur letzte Aenderung pro ID anzeigen (Modus B)" & vbCrLf & vbCrLf & _
                        "Zeitraum: " & Format(startDate, "dd.mm.yyyy") & " - " & Format(endDate, "dd.mm.yyyy") & vbCrLf & _
                        "Datenquelle: " & IIf(useDatabase, "Datenbank (tblKartei)", karteiSheetName) & vbCrLf & vbCrLf & _
                        IIf(useDatabase, _
                            "Hinweis: Datenbankmodus kann je nach Datenmenge laenger dauern.", _
                            "Hinweis: Diese Ansicht zeigt nur Datensaetze aus dem Blatt '" & karteiSheetName & "'."), _
                        vbYesNoCancel + vbQuestion, "Geschichte Alle " & Format(year2, "00") & " - Modus")
    Application.ScreenUpdating = False
    
    If userChoice = vbCancel Then
        GoTo Cleanup
    End If
    
    showOnlyLastChange = (userChoice = vbNo)
    
    ' Clear content from row 2 onwards (preserve row 1 with date inputs)
    wsHistory.Rows("2:" & wsHistory.Rows.Count).Clear
    
    ' Create headers in row 2
    Call CreateAlleHistoryHeaders(wsHistory)
    
    outputRow = 3 ' Start output from row 3
    processedCount = 0
    
    If useDatabase Then
        dbPath = valid_DatabasePath.GetValidatedDatabasePathForYear(year2)
        If dbPath = "" Then GoTo Cleanup
        
        Set db = OpenDatabaseSafe(dbPath)
        If db Is Nothing Then
            MsgBox "Datenbank konnte nicht geoeffnet werden: " & dbPath, vbCritical, "Geschichte Alle " & Format(year2, "00")
            GoTo Cleanup
        End If
        
        Set rs = db.OpenRecordset( _
            "SELECT ID, Value1, Value2, Value4, Value5, Value6, Value7, Value8, Value9, " & _
            "Value10, Value13, Value15, Value18, Value37, Value38, Value39, Value47, Value52 " & _
            "FROM tblKartei ORDER BY ID", dbOpenSnapshot)
        
        Do While Not rs.EOF
            recordID = Trim$(CStr(Nz(rs.Fields("ID").Value, "")))
            strGeschichte = CStr(Nz(rs.Fields("Value52").Value, ""))
            
            If recordID <> "" And strGeschichte <> "" Then
                ' Prepare baseline values in Kartei-column index space (1..52)
                recordValues = BuildKarteiStyleValuesFromRecordset(rs)
                
                segments = Split(strGeschichte, "||")
                Set result = valid_ParseHistory.ParseHistory(strGeschichte)
                
                If Not result Is Nothing Then
                    If result.Count > 0 Then
                        If showOnlyLastChange Then
                            Set lastEvent = FindLastEventInRange(result, segments, startDate, endDate)
                            If Not lastEvent Is Nothing Then
                                Call CreateAlleHistoryEntryFromValues(wsHistory, recordValues, outputRow, _
                                    lastEvent("IsRuck"), lastEvent("Reason"), CStr(lastEvent("ChangeDate")), lastEvent("Changes"))
                                outputRow = outputRow + 3
                                processedCount = processedCount + 1
                            End If
                        Else
                            For i = 1 To result.Count
                                Set evt = result(i)
                                
                                segmentText = ""
                                On Error Resume Next
                                If (i - 1) >= LBound(segments) And (i - 1) <= UBound(segments) Then
                                    segmentText = Trim$(segments(i - 1))
                                End If
                                On Error GoTo Cleanup
                                
                                Set fieldChanges = ParseFieldChangesFromSegment(segmentText)
                                Call MergeFieldChangesIntoChanges(evt("Changes"), fieldChanges)
                                
                                On Error Resume Next
                                eventDate = CDate(evt("ChangeDate"))
                                If Err.Number = 0 Then
                                    If eventDate >= startDate And eventDate <= endDate Then
                                        Call CreateAlleHistoryEntryFromValues(wsHistory, recordValues, outputRow, _
                                            evt("IsRuck"), evt("Reason"), CStr(evt("ChangeDate")), evt("Changes"))
                                        outputRow = outputRow + 3
                                        processedCount = processedCount + 1
                                    End If
                                End If
                                Err.Clear
                                On Error GoTo Cleanup
                            Next i
                        End If
                    End If
                End If
            End If
            
            rs.MoveNext
        Loop
        
        rs.Close
        db.Close
        Set rs = Nothing
        Set db = Nothing
    Else
        ' Sheet-based mode (pending subset)
        On Error Resume Next
        Set wsKartei = ThisWorkbook.Worksheets(karteiSheetName)
        On Error GoTo Cleanup
        
        If wsKartei Is Nothing Then
            MsgBox "Das Blatt '" & karteiSheetName & "' existiert nicht." & vbCrLf & _
                   "Bitte importieren Sie zuerst die Daten fuer Jahr 20" & Format(year2, "00") & ".", _
                   vbExclamation, "Geschichte Alle " & Format(year2, "00")
            GoTo Cleanup
        End If
        
        ' Find last row in Kartei using ID column (more reliable than column A)
        lastRow = wsKartei.Cells(wsKartei.Rows.Count, ID_COL).End(xlUp).Row
        
        For currentRow = 3 To lastRow
            strGeschichte = wsKartei.Cells(currentRow, HISTORY_COL).Value
            operName = wsKartei.Cells(currentRow, OPERATOR_COL).Value
            recordID = CStr(wsKartei.Cells(currentRow, ID_COL).Value)
            
            If strGeschichte <> "" And recordID <> "" Then
                segments = Split(strGeschichte, "||")
                Set result = valid_ParseHistory.ParseHistory(strGeschichte)
                
                If result Is Nothing Then GoTo NextAlleRow
                If result.Count = 0 Then GoTo NextAlleRow
                
                If showOnlyLastChange Then
                    Set lastEvent = FindLastEventInRange(result, segments, startDate, endDate)
                    
                    If Not lastEvent Is Nothing Then
                        Call CreateAlleHistoryEntry(wsHistory, wsKartei, outputRow, _
                            lastEvent("IsRuck"), lastEvent("Reason"), _
                            CStr(lastEvent("ChangeDate")), lastEvent("Changes"), _
                            currentRow, operName, recordID)
                        outputRow = outputRow + 3
                        processedCount = processedCount + 1
                    End If
                Else
                    For i = 1 To result.Count
                        Set evt = result(i)
                        
                        segmentText = ""
                        On Error Resume Next
                        If (i - 1) >= LBound(segments) And (i - 1) <= UBound(segments) Then
                            segmentText = Trim$(segments(i - 1))
                        End If
                        On Error GoTo Cleanup
                        
                        Set fieldChanges = ParseFieldChangesFromSegment(segmentText)
                        Call MergeFieldChangesIntoChanges(evt("Changes"), fieldChanges)
                        
                        On Error Resume Next
                        eventDate = CDate(evt("ChangeDate"))
                        If Err.Number = 0 Then
                            If eventDate >= startDate And eventDate <= endDate Then
                                Call CreateAlleHistoryEntry(wsHistory, wsKartei, outputRow, _
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
NextAlleRow:
        Next currentRow
    End If
    
    ' Apply final formatting
    Call FormatAlleHistorySheet(wsHistory, outputRow - 1)
    
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
    
    MsgBox "Geschichte-Bericht erfolgreich generiert!" & vbCrLf & vbCrLf & _
           "Zeitraum: " & Format(startDate, "dd.mm.yyyy") & " - " & Format(endDate, "dd.mm.yyyy") & vbCrLf & _
           "Modus: " & modeText & vbCrLf & _
           "Verarbeitete Eintraege: " & processedCount & vbCrLf & _
           "Ausgabe: " & historySheetName & vbCrLf & vbCrLf & _
           "Datenquelle: " & IIf(useDatabase, "Datenbank (tblKartei)", karteiSheetName), _
           vbInformation, "Geschichte Alle " & Format(year2, "00")
    
Cleanup:
    On Error Resume Next
    If Not rs Is Nothing Then rs.Close
    If Not db Is Nothing Then db.Close
    Set rs = Nothing
    Set db = Nothing
    On Error GoTo 0
    
    Application.Calculation = xlCalculationAutomatic
    Application.ScreenUpdating = True
    
    If Err.Number <> 0 Then
        MsgBox "Es ist ein Fehler aufgetreten: " & Err.Description & vbCrLf & _
               "Zeile: " & currentRow & ", ID: " & recordID, _
               vbCritical, "Geschichte Alle " & Format(year2, "00") & " - Fehler"
    End If
End Sub

' ============================================================
' SHEET MANAGEMENT - Single Record History
' ============================================================

' Get or create the single-record history sheet for a year
Private Function GetOrCreateSingleHistorySheet(ByVal sheetName As String) As Worksheet
    Dim ws As Worksheet
    
    On Error Resume Next
    Set ws = ThisWorkbook.Worksheets(sheetName)
    On Error GoTo 0
    
    If ws Is Nothing Then
        ' Create new sheet
        Set ws = ThisWorkbook.Worksheets.Add(After:=ThisWorkbook.Worksheets(ThisWorkbook.Worksheets.Count))
        ws.Name = sheetName
        
        ' Add date range inputs
        ws.Range("A1").Value = "Zeitraum:"
        ws.Range("A1").Font.Bold = True
        ws.Range("B1:C1").NumberFormat = "dd.mm.yyyy"
        
        Dim year2 As Integer
        year2 = valid_YearConfig.GetYearFromSheetName(sheetName)
        ApplyDefaultDateRangeByYear ws, year2, True
    Else
        ' Ensure date format
        ws.Range("B1:C1").NumberFormat = "dd.mm.yyyy"
    End If
    
    Set GetOrCreateSingleHistorySheet = ws
End Function

' Create headers for single-record history sheet (compact layout like old Geschichte.bas)
Private Sub CreateSingleHistoryHeaders(ws As Worksheet)
    Dim headers As Variant
    headers = Array("ID", "Eltern", "Kind", "Gruppe I", "Gruppe II", _
                    "Jan", "Feb", "Mrz", "Apr", "Mai", "Jun", _
                    "Jul", "Aug", "Sep", "Okt", "Nov", "Dez", _
                    "Date", "Operator", "Comments")
    
    Dim i As Long
    For i = LBound(headers) To UBound(headers)
        With ws.Cells(2, i + 1)
            .Value = headers(i)
            .Font.Bold = True
            .HorizontalAlignment = xlCenter
        End With
    Next i
    
    ' Set column widths
    ws.Columns("A").ColumnWidth = 8   ' ID
    ws.Columns("B").ColumnWidth = 20  ' Eltern
    ws.Columns("C").ColumnWidth = 20  ' Kind
    ws.Columns("D").ColumnWidth = 22  ' Gruppe I
    ws.Columns("E").ColumnWidth = 22  ' Gruppe II
    ws.Columns("F:Q").ColumnWidth = 6 ' Months
    ws.Columns("R").ColumnWidth = 12  ' Date
    ws.Columns("S").ColumnWidth = 9   ' Operator
    ws.Columns("T").ColumnWidth = 40  ' Comments
End Sub

' Create a single history entry for single-record view
Private Sub CreateSingleHistoryEntry(wsHistory As Worksheet, wsKartei As Worksheet, _
                                    startRow As Long, isRuck As Boolean, _
                                    Reason As String, changeDate As String, _
                                    Changes As Object, karteiRow As Long, _
                                    operName As String, recordID As String)
    Dim j As Long
    j = (startRow - 1) / 3 + 1 ' Calculate entry number
    
    Dim rowWar As Long
    Dim rowIst As Long
    Dim rowSeparator As Long
    
    rowWar = startRow
    rowIst = startRow + 1
    rowSeparator = startRow + 2
    
    ' Set the separator row
    wsHistory.Rows(rowSeparator).RowHeight = wsHistory.StandardHeight * 1 / 4
    wsHistory.Range("A" & rowSeparator & ":T" & rowSeparator).Interior.Color = RGB(192, 192, 192)
    
    ' Format cell borders
    With wsHistory.Range("A" & rowWar & ":S" & rowIst).Borders
        .LineStyle = xlContinuous
        .Color = vbBlack
        .Weight = xlThin
    End With
    
    ' Fill comments and formatting
    If isRuck Then
        wsHistory.Range("T" & rowIst).Interior.Color = RGB(255, 153, 204) ' Pink for Ruck
    Else
        wsHistory.Range("T" & rowIst).Interior.Color = RGB(204, 255, 153) ' Green for normal
    End If
    wsHistory.Range("T" & rowIst).Value = Reason
    
    wsHistory.Range("S" & rowIst).Value = operName
    wsHistory.Range("R" & rowIst).Value = changeDate
    wsHistory.Range("A" & rowIst).Value = recordID
    wsHistory.Range("B" & rowIst).Value = wsKartei.Cells(karteiRow, 2).Value  ' Parent
    wsHistory.Range("C" & rowIst).Value = wsKartei.Cells(karteiRow, 4).Value  ' Child
    wsHistory.Range("D" & rowIst).Value = wsKartei.Cells(karteiRow, 10).Value ' Subject1
    wsHistory.Range("E" & rowIst).Value = wsKartei.Cells(karteiRow, 15).Value ' Subject2
    
    ' Process month changes
    Dim changeKey As Variant
    For Each changeKey In Changes.Keys
        Dim keyStr As String
        keyStr = CStr(changeKey)
        
        If IsNumeric(keyStr) Then
            Dim monthNum As Long
            monthNum = CLng(keyStr)
            
            If monthNum >= 1 And monthNum <= 12 Then
                Dim colIndex As Long
                colIndex = 5 + monthNum ' F=6 for month 1, Q=17 for month 12
                
                ' War value
                Dim warVal As String
                warVal = Changes(changeKey)("War")
                If warVal = "" Then
                    wsHistory.Cells(rowWar, colIndex).Value = 0
                Else
                    On Error Resume Next
                    wsHistory.Cells(rowWar, colIndex).Value = CDbl(warVal)
                    If Err.Number <> 0 Then
                        wsHistory.Cells(rowWar, colIndex).Value = 0
                        Err.Clear
                    End If
                    On Error GoTo 0
                End If
                
                ' Ist value
                Dim istVal As String
                istVal = Changes(changeKey)("Ist")
                If istVal = "" Then
                    wsHistory.Cells(rowIst, colIndex).Value = 0
                Else
                    On Error Resume Next
                    wsHistory.Cells(rowIst, colIndex).Value = CDbl(istVal)
                    If Err.Number <> 0 Then
                        wsHistory.Cells(rowIst, colIndex).Value = 0
                        Err.Clear
                    End If
                    On Error GoTo 0
                End If
                
                ' Highlight changed cell
                wsHistory.Cells(rowIst, colIndex).Interior.Color = RGB(255, 192, 203)
            End If
        Else
            ' Non-numeric key - add to comments
            Dim fieldComment As String
            fieldComment = keyStr & ": War(" & Changes(changeKey)("War") & "); Ist(" & Changes(changeKey)("Ist") & ")."
            
            If wsHistory.Range("T" & rowIst).Value <> "" Then
                wsHistory.Range("T" & rowIst).Value = wsHistory.Range("T" & rowIst).Value & vbCrLf & fieldComment
            Else
                wsHistory.Range("T" & rowIst).Value = fieldComment
            End If
        End If
    Next changeKey
    
    ' Format months columns
    With wsHistory.Range("F" & rowWar & ":Q" & rowIst)
        .NumberFormat = "0.00"
    End With
End Sub

' Format single-record history sheet
Private Sub FormatSingleHistorySheet(ws As Worksheet, lastRow As Long)
    ' Add borders to headers
    With ws.Range("A2:T2").Borders
        .LineStyle = xlContinuous
        .Color = vbBlack
        .Weight = xlMedium
    End With
    
    ' Format headers background
    ws.Range("A2:T2").Interior.Color = RGB(220, 220, 220)
End Sub

' ============================================================
' SHEET MANAGEMENT - All Records History (GeschichteYY_Alle)
' ============================================================

' Get or create the all-records history sheet for a year
Private Function GetOrCreateAlleHistorySheet(ByVal sheetName As String) As Worksheet
    Dim ws As Worksheet
    
    On Error Resume Next
    Set ws = ThisWorkbook.Worksheets(sheetName)
    On Error GoTo 0
    
    If ws Is Nothing Then
        ' Create new sheet
        Set ws = ThisWorkbook.Worksheets.Add(After:=ThisWorkbook.Worksheets(ThisWorkbook.Worksheets.Count))
        ws.Name = sheetName
        
        ' Add date range inputs
        ws.Range("A1").Value = "Zeitraum:"
        ws.Range("A1").Font.Bold = True
        ws.Range("B1:C1").NumberFormat = "dd.mm.yyyy"
        
        Dim year2 As Integer
        year2 = valid_YearConfig.GetYearFromSheetName(sheetName)
        ApplyDefaultDateRangeByYear ws, year2, False
    Else
        ' Ensure date format
        ws.Range("B1:C1").NumberFormat = "dd.mm.yyyy"
        
        ' Set defaults only if invalid/empty (do NOT overwrite user input)
        If Not IsDate(ws.Range("B1").Value) Or IsEmpty(ws.Range("B1").Value) Then
            ws.Range("B1").Value = Date - 30
        End If
        If Not IsDate(ws.Range("C1").Value) Or IsEmpty(ws.Range("C1").Value) Then
            ws.Range("C1").Value = Date
        End If
    End If
    
    Set GetOrCreateAlleHistorySheet = ws
End Function

' ============================================================
' DATABASE SUPPORT (tblKartei) FOR FULL HISTORY
' ============================================================

' Open Access database safely (returns Nothing on failure)
Private Function OpenDatabaseSafe(ByVal dbPath As String) As DAO.Database
    On Error GoTo ErrorHandler
    
    Dim engine As DAO.DBEngine
    Set engine = New DAO.DBEngine
    
    Set OpenDatabaseSafe = engine.Workspaces(0).OpenDatabase(dbPath)
    Exit Function
    
ErrorHandler:
    Set OpenDatabaseSafe = Nothing
End Function

' Build a Kartei-like 1..52 array (only required indices populated) from tblKartei recordset.
' This allows reusing the same rendering logic as the sheet-based mode.
Private Function BuildKarteiStyleValuesFromRecordset(ByVal rs As DAO.Recordset) As Variant
    Dim values As Variant
    ReDim values(1 To 52)
    
    values(1) = Nz(rs.Fields("Value1").Value, "")
    values(2) = Nz(rs.Fields("Value2").Value, "")
    values(4) = Nz(rs.Fields("Value4").Value, "")
    values(5) = Nz(rs.Fields("Value5").Value, "")
    values(6) = Nz(rs.Fields("Value6").Value, "")
    ' Phone columns (7=Tel., 8=Handy): must be stored as strings to preserve
    ' leading zeros and prevent scientific notation (e.g. "0176..." not "1.76E+9")
    ' Also normalize any existing scientific-notation strings (e.g. "1,76E+9" -> "176000000")
    values(7) = phone_Normalize.NormalizePhoneText(rs.Fields("Value7").Value)
    values(8) = phone_Normalize.NormalizePhoneText(rs.Fields("Value8").Value)
    values(9) = Nz(rs.Fields("Value9").Value, "")
    values(10) = Nz(rs.Fields("Value10").Value, "")
    values(13) = Nz(rs.Fields("Value13").Value, "")
    values(15) = Nz(rs.Fields("Value15").Value, "")
    values(18) = Nz(rs.Fields("Value18").Value, "")
    values(37) = Nz(rs.Fields("Value37").Value, "")
    values(38) = Nz(rs.Fields("Value38").Value, "")
    values(39) = Nz(rs.Fields("Value39").Value, "")
    values(47) = Nz(rs.Fields("Value47").Value, "")
    
    BuildKarteiStyleValuesFromRecordset = values
End Function

' Nz helper (Null/Empty -> defaultValue)
Private Function Nz(ByVal Value As Variant, Optional ByVal defaultValue As Variant = "") As Variant
    If IsNull(Value) Or IsEmpty(Value) Then
        Nz = defaultValue
    Else
        Nz = Value
    End If
End Function

' Default date range helper for year-specific sheets.
' For past years: whole year. For current year: last 30 days (clamped to year start).
Private Sub ApplyDefaultDateRangeByYear(ByVal ws As Worksheet, ByVal year2 As Integer, ByVal forSingleRecord As Boolean)
    Dim fullYear As Long
    If year2 > 0 Then
        fullYear = valid_YearConfig.GetFullYear(year2)
    Else
        fullYear = Year(Date)
    End If
    
    Dim yearStart As Date
    Dim yearEnd As Date
    yearStart = DateSerial(fullYear, 1, 1)
    yearEnd = DateSerial(fullYear, 12, 31)
    
    If fullYear < Year(Date) Then
        ws.Range("B1").Value = yearStart
        ws.Range("C1").Value = yearEnd
        Exit Sub
    End If
    
    If fullYear = Year(Date) Then
        Dim startDate As Date
        startDate = Date - 30
        If startDate < yearStart Then startDate = yearStart
        
        ws.Range("B1").Value = startDate
        ws.Range("C1").Value = Date
        Exit Sub
    End If
    
    ' Future year / fallback
    ws.Range("B1").Value = yearStart
    ws.Range("C1").Value = yearEnd
End Sub

' Create a history entry for all-records view (database mode).
' Uses a Kartei-style values array instead of a worksheet row.
Private Sub CreateAlleHistoryEntryFromValues(ByVal wsHistory As Worksheet, ByVal values As Variant, _
                                            ByVal startRow As Long, ByVal isRuck As Boolean, _
                                            ByVal Reason As String, ByVal changeDate As String, _
                                            ByVal Changes As Object)
    
    Dim rowWar As Long
    Dim rowIst As Long
    Dim rowSeparator As Long
    
    rowWar = startRow
    rowIst = startRow + 1
    rowSeparator = startRow + 2
    
    wsHistory.Rows(rowSeparator).RowHeight = wsHistory.StandardHeight * 1 / 4
    wsHistory.Range("A" & rowSeparator & ":AB" & rowSeparator).Interior.Color = RGB(192, 192, 192)
    
    Dim i As Long
    For i = 0 To 1
        Dim currentRow As Long
        currentRow = startRow + i
        
        wsHistory.Range("A" & currentRow).Value = CStr(values(1)) ' FamilyID
        wsHistory.Range("B" & currentRow).Value = CStr(values(2)) ' Parent
        wsHistory.Range("C" & currentRow).Value = CStr(values(4)) ' Child
        wsHistory.Range("D" & currentRow).Value = FormatAsText(values(5)) ' Birthdate
        wsHistory.Range("E" & currentRow).Value = CStr(values(6)) ' Address
        ' Phone columns (7=Tel., 8=Handy): values are already strings from BuildKarteiStyleValuesFromRecordset
        ' Use CStr to ensure no scientific notation
        wsHistory.Range("F" & currentRow).Value = CStr(values(7)) ' Phone
        wsHistory.Range("G" & currentRow).Value = CStr(values(8)) ' Mobile
        wsHistory.Range("H" & currentRow).Value = CStr(values(9)) ' Email
        wsHistory.Range("I" & currentRow).Value = CStr(values(10)) ' Subject1
        wsHistory.Range("J" & currentRow).Value = CStr(values(13)) ' Price1
        wsHistory.Range("K" & currentRow).Value = CStr(values(15)) ' Subject2
        wsHistory.Range("L" & currentRow).Value = CStr(values(18)) ' Price2
        wsHistory.Range("Y" & currentRow).Value = CStr(values(37)) ' Extra1
        wsHistory.Range("Z" & currentRow).Value = CStr(values(38)) ' Extra2
        wsHistory.Range("AA" & currentRow).Value = CStr(values(39)) ' Extra3
    Next i
    
    If isRuck Then
        wsHistory.Range("AB" & rowIst).Interior.Color = RGB(255, 153, 204)
    Else
        wsHistory.Range("AB" & rowIst).Interior.Color = RGB(204, 255, 153)
    End If
    wsHistory.Range("AB" & rowIst).Value = Reason
    
    Dim sepaMarker As String
    sepaMarker = Trim$(UCase$(CStr(values(47))))
    If sepaMarker = "SEPA" Then
        Dim commentCell As Range
        Set commentCell = wsHistory.Range("AB" & rowIst)
        
        Dim baseText As String
        baseText = CStr(commentCell.Value)
        
        Dim startPos As Long
        If baseText <> "" Then
            commentCell.Value = baseText & vbCrLf & "SEPA"
            startPos = Len(baseText) + 2
        Else
            commentCell.Value = "SEPA"
            startPos = 1
        End If
        
        With commentCell.Characters(startPos, 4).Font
            .Color = vbRed
            .Bold = True
        End With
    End If
    
    ' Reuse the same change application logic as sheet-mode by delegating to the existing routine:
    ' We can call the sheet-mode routine by temporarily emulating the parts below, but to avoid
    ' deep refactors we keep the core logic duplicated here.
    '
    ' NOTE: changeDate is not rendered as a separate column in the A-AB layout (matches Data-file behavior).
    
    Dim changeKey As Variant
    Dim decimalSeparator As String
    decimalSeparator = Application.International(xlDecimalSeparator)
    
    Dim addressWar As String, addressIst As String
    Dim subject1War As String, subject1Ist As String
    Dim subject2War As String, subject2Ist As String
    Dim phoneWar As String, phoneIst As String
    Dim mobileWar As String, mobileIst As String
    Dim emailWar As String, emailIst As String
    
    addressWar = "": addressIst = ""
    subject1War = "": subject1Ist = ""
    subject2War = "": subject2Ist = ""
    phoneWar = "": phoneIst = ""
    mobileWar = "": mobileIst = ""
    emailWar = "": emailIst = ""
    
    On Error Resume Next
    
    For Each changeKey In Changes.Keys
        Dim keyStr As String
        keyStr = CStr(changeKey)
        
        If IsNumeric(keyStr) Then
            Dim monthNum As Long
            monthNum = CLng(keyStr)
            
            If monthNum >= 1 And monthNum <= 12 Then
                Dim colIndex As Long
                colIndex = 12 + monthNum ' M=13 for month 1, X=24 for month 12
                
                Dim warValue As String
                warValue = Changes(changeKey)("War")
                If warValue = "" Then
                    wsHistory.Cells(rowWar, colIndex).Value = 0
                Else
                    warValue = ConvertDecimalSeparator(warValue, decimalSeparator)
                    If IsNumeric(warValue) Then
                        wsHistory.Cells(rowWar, colIndex).Value = CDbl(warValue)
                    Else
                        wsHistory.Cells(rowWar, colIndex).Value = 0
                    End If
                End If
                
                Dim istValue As String
                istValue = Changes(changeKey)("Ist")
                If istValue = "" Then
                    wsHistory.Cells(rowIst, colIndex).Value = 0
                Else
                    istValue = ConvertDecimalSeparator(istValue, decimalSeparator)
                    If IsNumeric(istValue) Then
                        wsHistory.Cells(rowIst, colIndex).Value = CDbl(istValue)
                    Else
                        wsHistory.Cells(rowIst, colIndex).Value = 0
                    End If
                End If
                
                wsHistory.Cells(rowIst, colIndex).Interior.Color = RGB(255, 192, 203)
            End If
        Else
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
            End Select
        End If
    Next changeKey
    
    On Error GoTo 0
    
    If addressWar <> "" Or addressIst <> "" Then
        wsHistory.Cells(rowWar, 5).Value = addressWar
        wsHistory.Cells(rowIst, 5).Value = addressIst
        wsHistory.Cells(rowIst, 5).Interior.Color = RGB(255, 192, 203)
    End If
    
    If subject1War <> "" Or subject1Ist <> "" Then
        wsHistory.Cells(rowWar, 9).Value = subject1War
        wsHistory.Cells(rowIst, 9).Value = subject1Ist
        wsHistory.Cells(rowIst, 9).Interior.Color = RGB(255, 192, 203)
    End If
    
    If subject2War <> "" Or subject2Ist <> "" Then
        wsHistory.Cells(rowWar, 11).Value = subject2War
        wsHistory.Cells(rowIst, 11).Value = subject2Ist
        wsHistory.Cells(rowIst, 11).Interior.Color = RGB(255, 192, 203)
    End If
    
    If phoneWar <> "" Or phoneIst <> "" Then
        wsHistory.Cells(rowWar, 6).Value = phoneWar
        wsHistory.Cells(rowIst, 6).Value = phoneIst
        wsHistory.Cells(rowIst, 6).Interior.Color = RGB(255, 192, 203)
    End If
    
    If mobileWar <> "" Or mobileIst <> "" Then
        wsHistory.Cells(rowWar, 7).Value = mobileWar
        wsHistory.Cells(rowIst, 7).Value = mobileIst
        wsHistory.Cells(rowIst, 7).Interior.Color = RGB(255, 192, 203)
    End If
    
    If emailWar <> "" Or emailIst <> "" Then
        wsHistory.Cells(rowWar, 8).Value = emailWar
        wsHistory.Cells(rowIst, 8).Value = emailIst
        wsHistory.Cells(rowIst, 8).Interior.Color = RGB(255, 192, 203)
    End If
    
    With wsHistory.Range("M" & rowWar & ":X" & rowIst)
        .NumberFormat = "0.00"
    End With
    
    With wsHistory.Range("A" & rowWar & ":AB" & rowIst).Borders
        .LineStyle = xlContinuous
        .Color = vbBlack
        .Weight = xlThin
    End With
End Sub

' Create headers for all-records history sheet (expanded layout like grossGeschichte)
' Column structure: A-AB (28 columns) - NO decision columns
Private Sub CreateAlleHistoryHeaders(ws As Worksheet)
    Dim headers As Variant
    headers = Array("FamilyID", "Parent", "Child", "Birthdate", "Address", "Phone", "Mobile", "Email", _
                    "Subject1", "Price1", "Subject2", "Price2", _
                    "Jan", "Feb", "Mrz", "Apr", "Mai", "Jun", "Jul", "Aug", "Sep", "Okt", "Nov", "Dez", _
                    "Extra1", "Extra2", "Extra3", "Comments")
    
    Dim i As Long
    For i = LBound(headers) To UBound(headers)
        With ws.Cells(2, i + 1)
            .Value = headers(i)
            .Font.Bold = True
            .HorizontalAlignment = xlCenter
        End With
    Next i
    
    ' Set column widths
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
    ws.Columns("AB").ColumnWidth = 40 ' Comments
    
    ' Set text format for columns that should not be auto-converted
    ws.Range("A2:L" & ws.Rows.Count).NumberFormat = "@"
    ws.Range("Y2:AB" & ws.Rows.Count).NumberFormat = "@"
    ws.Range("B1:C1").NumberFormat = "dd.mm.yyyy"
End Sub

' Create a history entry for all-records view
Private Sub CreateAlleHistoryEntry(wsHistory As Worksheet, wsKartei As Worksheet, _
                               startRow As Long, isRuck As Boolean, _
                               Reason As String, changeDate As String, _
                               Changes As Object, karteiRow As Long, _
                               operName As String, recordID As String)
    
    Dim rowWar As Long
    Dim rowIst As Long
    Dim rowSeparator As Long
    
    rowWar = startRow
    rowIst = startRow + 1
    rowSeparator = startRow + 2
    
    ' Create separator row
    wsHistory.Rows(rowSeparator).RowHeight = wsHistory.StandardHeight * 1 / 4
    wsHistory.Range("A" & rowSeparator & ":AB" & rowSeparator).Interior.Color = RGB(192, 192, 192)
    
    ' Fill basic data for both rows
    Dim i As Long
    For i = 0 To 1
        Dim currentRow As Long
        currentRow = startRow + i
        
        wsHistory.Range("A" & currentRow).Value = CStr(wsKartei.Cells(karteiRow, 1).Value) ' FamilyID
        wsHistory.Range("B" & currentRow).Value = CStr(wsKartei.Cells(karteiRow, 2).Value) ' Parent
        wsHistory.Range("C" & currentRow).Value = CStr(wsKartei.Cells(karteiRow, 4).Value) ' Child
        wsHistory.Range("D" & currentRow).Value = FormatAsText(wsKartei.Cells(karteiRow, 5).Value) ' Birthdate
        wsHistory.Range("E" & currentRow).Value = CStr(wsKartei.Cells(karteiRow, 6).Value) ' Address
        ' Phone columns (7=Tel., 8=Handy): use .Text to preserve leading zeros and prevent
        ' scientific notation. .Text returns the displayed string, not the internal numeric value.
        wsHistory.Range("F" & currentRow).Value = wsKartei.Cells(karteiRow, 7).Text ' Phone
        wsHistory.Range("G" & currentRow).Value = wsKartei.Cells(karteiRow, 8).Text ' Mobile
        wsHistory.Range("H" & currentRow).Value = CStr(wsKartei.Cells(karteiRow, 9).Value) ' Email
        wsHistory.Range("I" & currentRow).Value = CStr(wsKartei.Cells(karteiRow, 10).Value) ' Subject1
        wsHistory.Range("J" & currentRow).Value = CStr(wsKartei.Cells(karteiRow, 13).Value) ' Price1
        wsHistory.Range("K" & currentRow).Value = CStr(wsKartei.Cells(karteiRow, 15).Value) ' Subject2
        wsHistory.Range("L" & currentRow).Value = CStr(wsKartei.Cells(karteiRow, 18).Value) ' Price2
        wsHistory.Range("Y" & currentRow).Value = CStr(wsKartei.Cells(karteiRow, 37).Value) ' Extra1
        wsHistory.Range("Z" & currentRow).Value = CStr(wsKartei.Cells(karteiRow, 38).Value) ' Extra2
        wsHistory.Range("AA" & currentRow).Value = CStr(wsKartei.Cells(karteiRow, 39).Value) ' Extra3
    Next i
    
    ' Fill Comments in Ist row with appropriate formatting
    If isRuck Then
        wsHistory.Range("AB" & rowIst).Interior.Color = RGB(255, 153, 204) ' Pink for Ruck
    Else
        wsHistory.Range("AB" & rowIst).Interior.Color = RGB(204, 255, 153) ' Green for normal
    End If
    wsHistory.Range("AB" & rowIst).Value = Reason
    
    ' Add SEPA marker if present
    Dim sepaMarker As String
    sepaMarker = Trim$(UCase$(CStr(wsKartei.Cells(karteiRow, SEPA_COL).Value)))
    If sepaMarker = "SEPA" Then
        Dim commentCell As Range
        Set commentCell = wsHistory.Range("AB" & rowIst)
        
        Dim baseText As String
        baseText = CStr(commentCell.Value)
        
        Dim startPos As Long
        If baseText <> "" Then
            commentCell.Value = baseText & vbCrLf & "SEPA"
            startPos = Len(baseText) + 2
        Else
            commentCell.Value = "SEPA"
            startPos = 1
        End If
        
        With commentCell.Characters(startPos, 4).Font
            .Color = vbRed
            .Bold = True
        End With
    End If
    
    ' Process changes from the Changes dictionary
    Dim changeKey As Variant
    Dim decimalSeparator As String
    decimalSeparator = Application.International(xlDecimalSeparator)
    
    ' Variables for field changes
    Dim addressWar As String, addressIst As String
    Dim subject1War As String, subject1Ist As String
    Dim subject2War As String, subject2Ist As String
    Dim phoneWar As String, phoneIst As String
    Dim mobileWar As String, mobileIst As String
    Dim emailWar As String, emailIst As String
    
    addressWar = "": addressIst = ""
    subject1War = "": subject1Ist = ""
    subject2War = "": subject2Ist = ""
    phoneWar = "": phoneIst = ""
    mobileWar = "": mobileIst = ""
    emailWar = "": emailIst = ""
    
    On Error Resume Next
    
    For Each changeKey In Changes.Keys
        Dim keyStr As String
        keyStr = CStr(changeKey)
        
        If IsNumeric(keyStr) Then
            Dim monthNum As Long
            monthNum = CLng(keyStr)
            
            If monthNum >= 1 And monthNum <= 12 Then
                Dim colIndex As Long
                colIndex = 12 + monthNum ' M=13 for month 1, X=24 for month 12
                
                ' War value
                Dim warValue As String
                warValue = Changes(changeKey)("War")
                If warValue = "" Then
                    wsHistory.Cells(rowWar, colIndex).Value = 0
                Else
                    warValue = ConvertDecimalSeparator(warValue, decimalSeparator)
                    If IsNumeric(warValue) Then
                        wsHistory.Cells(rowWar, colIndex).Value = CDbl(warValue)
                    Else
                        wsHistory.Cells(rowWar, colIndex).Value = 0
                    End If
                End If
                
                ' Ist value
                Dim istValue As String
                istValue = Changes(changeKey)("Ist")
                If istValue = "" Then
                    wsHistory.Cells(rowIst, colIndex).Value = 0
                Else
                    istValue = ConvertDecimalSeparator(istValue, decimalSeparator)
                    If IsNumeric(istValue) Then
                        wsHistory.Cells(rowIst, colIndex).Value = CDbl(istValue)
                    Else
                        wsHistory.Cells(rowIst, colIndex).Value = 0
                    End If
                End If
                
                ' Highlight changed cell
                wsHistory.Cells(rowIst, colIndex).Interior.Color = RGB(255, 192, 203)
            End If
        Else
            ' Non-numeric key - capture for dedicated columns
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
            End Select
        End If
    Next changeKey
    
    On Error GoTo 0
    
    ' Apply field changes to dedicated columns with highlighting
    If addressWar <> "" Or addressIst <> "" Then
        wsHistory.Cells(rowWar, 5).Value = addressWar
        wsHistory.Cells(rowIst, 5).Value = addressIst
        wsHistory.Cells(rowIst, 5).Interior.Color = RGB(255, 192, 203)
    End If
    
    If subject1War <> "" Or subject1Ist <> "" Then
        wsHistory.Cells(rowWar, 9).Value = subject1War
        wsHistory.Cells(rowIst, 9).Value = subject1Ist
        wsHistory.Cells(rowIst, 9).Interior.Color = RGB(255, 192, 203)
    End If
    
    If subject2War <> "" Or subject2Ist <> "" Then
        wsHistory.Cells(rowWar, 11).Value = subject2War
        wsHistory.Cells(rowIst, 11).Value = subject2Ist
        wsHistory.Cells(rowIst, 11).Interior.Color = RGB(255, 192, 203)
    End If
    
    If phoneWar <> "" Or phoneIst <> "" Then
        wsHistory.Cells(rowWar, 6).Value = phoneWar
        wsHistory.Cells(rowIst, 6).Value = phoneIst
        wsHistory.Cells(rowIst, 6).Interior.Color = RGB(255, 192, 203)
    End If
    
    If mobileWar <> "" Or mobileIst <> "" Then
        wsHistory.Cells(rowWar, 7).Value = mobileWar
        wsHistory.Cells(rowIst, 7).Value = mobileIst
        wsHistory.Cells(rowIst, 7).Interior.Color = RGB(255, 192, 203)
    End If
    
    If emailWar <> "" Or emailIst <> "" Then
        wsHistory.Cells(rowWar, 8).Value = emailWar
        wsHistory.Cells(rowIst, 8).Value = emailIst
        wsHistory.Cells(rowIst, 8).Interior.Color = RGB(255, 192, 203)
    End If
    
    ' Format monthly columns
    With wsHistory.Range("M" & rowWar & ":X" & rowIst)
        .NumberFormat = "0.00"
    End With
    
    ' Add borders
    With wsHistory.Range("A" & rowWar & ":AB" & rowIst).Borders
        .LineStyle = xlContinuous
        .Color = vbBlack
        .Weight = xlThin
    End With
End Sub

' Format all-records history sheet
Private Sub FormatAlleHistorySheet(ws As Worksheet, lastRow As Long)
    ' Add borders to headers
    With ws.Range("A2:AB2").Borders
        .LineStyle = xlContinuous
        .Color = vbBlack
        .Weight = xlMedium
    End With
    
    ' Format headers background
    ws.Range("A2:AB2").Interior.Color = RGB(220, 220, 220)
End Sub

' ============================================================
' HELPER FUNCTIONS
' ============================================================

' Check if year is valid (24, 25, or 26)
Private Function IsValidYear(ByVal year2 As Integer) As Boolean
    IsValidYear = (year2 = 24 Or year2 = 25 Or year2 = 26)
End Function

' Convert decimal separator based on system settings
Private Function ConvertDecimalSeparator(ByVal Value As String, ByVal systemDecimalSeparator As String) As String
    Value = Replace(Value, " ", "")
    If systemDecimalSeparator = "," Then
        Value = Replace(Value, ".", ",")
    Else
        Value = Replace(Value, ",", ".")
    End If
    ConvertDecimalSeparator = Value
End Function

' Format value as text (prevents Excel auto-conversion)
Private Function FormatAsText(ByVal Value As Variant) As String
    If IsEmpty(Value) Or IsNull(Value) Then
        FormatAsText = ""
    ElseIf IsDate(Value) Then
        FormatAsText = Format(Value, "dd.mm.yyyy")
    Else
        FormatAsText = CStr(Value)
    End If
End Function

' Parse Address/Subject1/Subject2 changes from a raw history segment
Private Function ParseFieldChangesFromSegment(ByVal segment As String) As Object
    Dim result As Object
    Set result = CreateObject("Scripting.Dictionary")
    
    If Len(Trim$(segment)) = 0 Then
        Set ParseFieldChangesFromSegment = result
        Exit Function
    End If
    
    ' Try legacy format: FieldName: Was(X); Is(Y).
    Call AddFieldChangeToDict(segment, "Address", result)
    Call AddFieldChangeToDict(segment, "Subject1", result)
    Call AddFieldChangeToDict(segment, "Subject2", result)
    
    ' Try new format: TAG(X->Y)
    Call AddNewFormatFieldChangeToDict(segment, "ADR", "Address", result)
    Call AddNewFormatFieldChangeToDict(segment, "SB1", "Subject1", result)
    Call AddNewFormatFieldChangeToDict(segment, "SB2", "Subject2", result)
    
    Set ParseFieldChangesFromSegment = result
End Function

' Merge field changes into the base Changes dictionary
Private Sub MergeFieldChangesIntoChanges(ByVal baseChanges As Object, ByVal fieldChanges As Object)
    If baseChanges Is Nothing Then Exit Sub
    If fieldChanges Is Nothing Then Exit Sub
    
    Dim key As Variant
    For Each key In fieldChanges.Keys
        If Not baseChanges.Exists(key) Then
            baseChanges.Add key, fieldChanges.Item(key)
        Else
            Set baseChanges.Item(key) = fieldChanges.Item(key)
        End If
    Next key
End Sub

' Helper: extract legacy format field change
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

' Helper: extract new format field change TAG(OLD->NEW)
Private Sub AddNewFormatFieldChangeToDict(ByVal segment As String, ByVal tag As String, ByVal fieldName As String, ByVal dict As Object)
    If dict.Exists(fieldName) Then Exit Sub
    
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

' Find the last event in date range for a given ID (Mode B helper)
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
                If lastEvt Is Nothing Then
                    Set lastEvt = evt
                    maxDate = eventDate
                    maxDateIndex = i
                ElseIf eventDate > maxDate Then
                    Set lastEvt = evt
                    maxDate = eventDate
                    maxDateIndex = i
                ElseIf eventDate = maxDate And i > maxDateIndex Then
                    Set lastEvt = evt
                    maxDateIndex = i
                End If
            End If
        End If
        Err.Clear
    Next i
    
    Set FindLastEventInRange = lastEvt
End Function
