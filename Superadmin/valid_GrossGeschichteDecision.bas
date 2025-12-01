Attribute VB_Name = "valid_GrossGeschichteDecision"
'==========================
'   Module: valid_GrossGeschichteDecision
'   Purpose: Build the GrossGeschichte decision sheet for Superadmin
'   - Creates War/Ist pairs for each pending record
'   - War = values from tblKartei (original), Ist = values from pre_tblKartei (pending)
'   - Highlights differences in tracked fields
'   - Aggregates comments from history within date range
'   - Self-contained for Superadmin file - no dependencies on Admin modules
'
'   GrossGeschichte Column Structure:
'   A=FamilyID, B=Parent, C=Child, D=Birthdate, E=Address, F=Phone, G=Mobile, H=Email
'   I=Subject1, J=Price1, K=Subject2, L=Price2, M-X=Months 1-12, Y-AA=Extra1-3
'   AB=Comments, AC=Decision, AD=Decline Comment, AE=RecordID (hidden)
'==========================

Option Explicit

' ============================================================
' CONSTANTS - Column Mappings
' ============================================================

' Kartei sheet columns (source data from pre_tblKartei / tblKartei)
Private Const K_FAMILY_ID As Long = 1      ' A
Private Const K_PARENT As Long = 2         ' B
Private Const K_CHILD As Long = 4          ' D
Private Const K_BIRTHDATE As Long = 5      ' E
Private Const K_ADDRESS As Long = 6        ' F
Private Const K_PHONE As Long = 7          ' G
Private Const K_MOBILE As Long = 8         ' H
Private Const K_EMAIL As Long = 9          ' I
Private Const K_SUBJECT1 As Long = 10      ' J
Private Const K_PRICE1 As Long = 13        ' M
Private Const K_SUBJECT2 As Long = 15      ' O
Private Const K_PRICE2 As Long = 18        ' R
Private Const K_MONTH_START As Long = 21   ' U (Month 1)
Private Const K_MONTH_END As Long = 32     ' AF (Month 12)
Private Const K_EXTRA1 As Long = 37        ' AK
Private Const K_EXTRA2 As Long = 38        ' AL
Private Const K_EXTRA3 As Long = 39        ' AM
Private Const K_ID As Long = 48            ' AV
Private Const K_HISTORY As Long = 52       ' AZ

' GrossGeschichte columns (output)
Private Const G_FAMILY_ID As Long = 1      ' A
Private Const G_PARENT As Long = 2         ' B
Private Const G_CHILD As Long = 3          ' C
Private Const G_BIRTHDATE As Long = 4      ' D
Private Const G_ADDRESS As Long = 5        ' E
Private Const G_PHONE As Long = 6          ' F
Private Const G_MOBILE As Long = 7         ' G
Private Const G_EMAIL As Long = 8          ' H
Private Const G_SUBJECT1 As Long = 9       ' I
Private Const G_PRICE1 As Long = 10        ' J
Private Const G_SUBJECT2 As Long = 11      ' K
Private Const G_PRICE2 As Long = 12        ' L
Private Const G_MONTH_START As Long = 13   ' M (Month 1)
Private Const G_MONTH_END As Long = 24     ' X (Month 12)
Private Const G_EXTRA1 As Long = 25        ' Y
Private Const G_EXTRA2 As Long = 26        ' Z
Private Const G_EXTRA3 As Long = 27        ' AA
Private Const G_COMMENTS As Long = 28      ' AB
Private Const G_DECISION As Long = 29      ' AC
Private Const G_DECLINE_COMMENT As Long = 30 ' AD
Private Const G_RECORD_ID As Long = 31     ' AE (hidden)

' Highlight color for changed cells (light pink - same as grossGeschichte.bas)
Private Const COLOR_CHANGED As Long = 16761024  ' RGB(255, 192, 203)
' Color for new record indicator (light green)
Private Const COLOR_NEW_RECORD As Long = 10092441 ' RGB(153, 255, 153)

' ============================================================
' PUBLIC ENTRY POINT
' ============================================================

' Main entry point: Builds the GrossGeschichte decision sheet
' Reads pending records from Kartei, compares with tblKartei,
' creates War/Ist pairs with highlighted differences
'
' Database optimization: Opens the Access database ONCE at the start
' and passes the connection to ProcessAllPendingRecords, which reuses it
' for all record lookups. This avoids the overhead of opening/closing
' the database for each pending ID.
Public Sub BuildPendingDecisionSheet()
    On Error GoTo ErrorHandler
    
    Application.ScreenUpdating = False
    Application.Calculation = xlCalculationManual
    
    ' Database connection - opened once, closed in Cleanup
    Dim db As DAO.Database
    Set db = Nothing
    
    ' Step 1: Get or create GrossGeschichte sheet
    Dim wsGross As Worksheet
    Set wsGross = GetOrCreateGrossGeschichteSheet()
    
    ' Step 2: Read and validate date range from B1:C1
    Dim startDate As Date
    Dim endDate As Date
    If Not ReadDateRange(wsGross, startDate, endDate) Then
        GoTo Cleanup
    End If
    
    ' Step 3: Get reference to Kartei sheet (contains pending data)
    Dim wsKartei As Worksheet
    On Error Resume Next
    Set wsKartei = ThisWorkbook.Worksheets("Kartei")
    On Error GoTo ErrorHandler
    
    If wsKartei Is Nothing Then
        MsgBox "Kartei-Blatt nicht gefunden. Bitte zuerst ausstehende Aenderungen laden.", _
               vbExclamation, "Fehler"
        GoTo Cleanup
    End If
    
    ' Step 4: Get database path for War values lookup
    Dim dbPath As String
    dbPath = valid_GrossGeschichteData.GetDatabasePath()
    
    If dbPath = "" Then
        MsgBox "Datenbankpfad nicht festgelegt. Vorgang abgebrochen.", vbExclamation, "Fehler"
        GoTo Cleanup
    End If
    
    ' Step 5: Open database connection ONCE for all record lookups
    ' This is the key optimization - avoiding repeated open/close for each ID
    Dim engine As DAO.DBEngine
    Set engine = New DAO.DBEngine
    
    On Error Resume Next
    Set db = engine.Workspaces(0).OpenDatabase(dbPath)
    On Error GoTo ErrorHandler
    
    If db Is Nothing Then
        MsgBox "Datenbank konnte nicht geoeffnet werden: " & dbPath, vbCritical, "Fehler"
        GoTo Cleanup
    End If
    
    ' Step 6: Clear existing data and create headers
    ClearAndPrepareSheet wsGross
    
    ' Step 7: Process all pending records using the single open database connection
    Dim processedCount As Long
    processedCount = ProcessAllPendingRecords(wsGross, wsKartei, db, startDate, endDate)
    
    ' Step 8: Apply final formatting
    ApplyFinalFormatting wsGross
    
    ' Step 9: Show completion message
    If processedCount > 0 Then
        MsgBox "GrossGeschichte erfolgreich erstellt." & vbCrLf & vbCrLf & _
               "Verarbeitete Eintraege: " & processedCount & vbCrLf & _
               "Zeitraum: " & Format(startDate, "dd.mm.yyyy") & " - " & Format(endDate, "dd.mm.yyyy") & vbCrLf & vbCrLf & _
               "Bitte ueberpruefen Sie die Aenderungen und markieren Sie:" & vbCrLf & _
               "  - 'Approved' in Spalte AC fuer Genehmigung" & vbCrLf & _
               "  - 'Declined' in Spalte AC fuer Ablehnung", _
               vbInformation, "GrossGeschichte erstellt"
    Else
        MsgBox "Keine ausstehenden Aenderungen zum Verarbeiten gefunden.", _
               vbInformation, "GrossGeschichte"
    End If
    
Cleanup:
    ' Close database connection (opened once in Step 5)
    On Error Resume Next
    If Not db Is Nothing Then
        db.Close
        Set db = Nothing
    End If
    On Error GoTo 0
    
    Application.Calculation = xlCalculationAutomatic
    Application.ScreenUpdating = True
    Exit Sub
    
ErrorHandler:
    ' Ensure database is closed even on error
    On Error Resume Next
    If Not db Is Nothing Then
        db.Close
        Set db = Nothing
    End If
    On Error GoTo 0
    
    Application.Calculation = xlCalculationAutomatic
    Application.ScreenUpdating = True
    MsgBox "Fehler beim Erstellen des GrossGeschichte-Blatts: " & Err.Description, _
           vbCritical, "Fehler"
End Sub

' ============================================================
' SHEET MANAGEMENT
' ============================================================

' Gets or creates the GrossGeschichte worksheet with proper structure
Private Function GetOrCreateGrossGeschichteSheet() As Worksheet
    Dim ws As Worksheet
    
    On Error Resume Next
    Set ws = ThisWorkbook.Worksheets("grossGeschichte")
    On Error GoTo 0
    
    If ws Is Nothing Then
        ' Create new sheet
        Set ws = ThisWorkbook.Worksheets.Add(After:=ThisWorkbook.Worksheets(ThisWorkbook.Worksheets.Count))
        ws.Name = "grossGeschichte"
        
        ' Set up date range inputs in row 1
        ws.Range("A1").Value = "Start Date:"
        ws.Range("A1").Font.Bold = True
        ws.Range("B1:C1").NumberFormat = "dd.mm.yyyy"
        ws.Range("B1").Value = Date - 30  ' Default: 30 days ago
        ws.Range("C1").Value = Date       ' Default: today
    Else
        ' Ensure date format is preserved
        ws.Range("B1:C1").NumberFormat = "dd.mm.yyyy"
        
        ' Set B1 to 30 days ago if empty or invalid
        If Not IsDate(ws.Range("B1").Value) Or IsEmpty(ws.Range("B1").Value) Then
            ws.Range("B1").Value = Date - 30
        End If
        
        ' Always update C1 to today
        ws.Range("C1").Value = Date
    End If
    
    Set GetOrCreateGrossGeschichteSheet = ws
End Function

' Reads and validates the date range from B1:C1
' Returns False if dates are invalid (shows MsgBox to user)
Private Function ReadDateRange(ByVal ws As Worksheet, ByRef startDate As Date, ByRef endDate As Date) As Boolean
    ReadDateRange = False
    
    If Not IsDate(ws.Range("B1").Value) Or Not IsDate(ws.Range("C1").Value) Then
        MsgBox "Bitte gueltige Datumsangaben in B1 (Startdatum) und C1 (Enddatum) eingeben.", _
               vbExclamation, "Ungueltige Datumsangaben"
        Exit Function
    End If
    
    startDate = CDate(ws.Range("B1").Value)
    endDate = CDate(ws.Range("C1").Value)
    
    If startDate > endDate Then
        MsgBox "Startdatum (B1) darf nicht nach Enddatum (C1) liegen.", _
               vbExclamation, "Ungueltige Datumsangaben"
        Exit Function
    End If
    
    ReadDateRange = True
End Function

' Clears existing data and creates headers
Private Sub ClearAndPrepareSheet(ByVal ws As Worksheet)
    ' Clear everything from row 2 onwards (keep row 1 with dates)
    ws.Rows("2:" & ws.Rows.Count).Clear
    
    ' Create headers in row 2
    CreateHeaders ws
End Sub

' Creates header row with all column labels
Private Sub CreateHeaders(ByVal ws As Worksheet)
    Dim headers As Variant
    headers = Array("FamilyID", "Parent", "Child", "Birthdate", "Address", "Phone", "Mobile", "Email", _
                    "Subject1", "Price1", "Subject2", "Price2", _
                    "Jan", "Feb", "Mrz", "Apr", "Mai", "Jun", "Jul", "Aug", "Sep", "Okt", "Nov", "Dez", _
                    "Extra1", "Extra2", "Extra3", "Comments", "Decision", "Decline Comment", "RecordID")
    
    Dim i As Long
    For i = LBound(headers) To UBound(headers)
        With ws.Cells(2, i + 1)
            .Value = headers(i)
            .Font.Bold = True
            .HorizontalAlignment = xlCenter
        End With
    Next i
    
    ' Format headers
    With ws.Range("A2:AE2")
        .Interior.Color = RGB(220, 220, 220)  ' Light gray background
        .Borders.LineStyle = xlContinuous
        .Borders.Weight = xlMedium
    End With
    
    ' Set column widths
    SetColumnWidths ws
    
    ' Hide RecordID column (AE)
    ws.Columns("AE").Hidden = True
    
    ' Set text format for non-numeric columns (data rows start at 3)
    ws.Range("A3:L" & ws.Rows.Count).NumberFormat = "@"
    ws.Range("Y3:AE" & ws.Rows.Count).NumberFormat = "@"
    
    ' Ensure date format in row 1
    ws.Range("B1:C1").NumberFormat = "dd.mm.yyyy"
End Sub

' Sets appropriate column widths
Private Sub SetColumnWidths(ByVal ws As Worksheet)
    ws.Columns("A").ColumnWidth = 10   ' FamilyID
    ws.Columns("B").ColumnWidth = 20   ' Parent
    ws.Columns("C").ColumnWidth = 18   ' Child
    ws.Columns("D").ColumnWidth = 12   ' Birthdate
    ws.Columns("E").ColumnWidth = 25   ' Address
    ws.Columns("F").ColumnWidth = 14   ' Phone
    ws.Columns("G").ColumnWidth = 14   ' Mobile
    ws.Columns("H").ColumnWidth = 22   ' Email
    ws.Columns("I").ColumnWidth = 18   ' Subject1
    ws.Columns("J").ColumnWidth = 8    ' Price1
    ws.Columns("K").ColumnWidth = 18   ' Subject2
    ws.Columns("L").ColumnWidth = 8    ' Price2
    ws.Columns("M:X").ColumnWidth = 6  ' Months
    ws.Columns("Y:AA").ColumnWidth = 15 ' Extra1-3
    ws.Columns("AB").ColumnWidth = 40  ' Comments
    ws.Columns("AC").ColumnWidth = 12  ' Decision
    ws.Columns("AD").ColumnWidth = 35  ' Decline Comment
End Sub

' ============================================================
' RECORD PROCESSING
' ============================================================

' Processes all pending records from Kartei sheet
' Returns the count of processed records
'
' @param wsGross - Output worksheet for War/Ist pairs
' @param wsKartei - Source worksheet with pending records
' @param db - Already-open DAO.Database connection for War record lookups
' @param startDate - Start of date range for comment aggregation
' @param endDate - End of date range for comment aggregation
' @return Long - Number of records processed
Private Function ProcessAllPendingRecords(ByVal wsGross As Worksheet, _
                                          ByVal wsKartei As Worksheet, _
                                          ByVal db As DAO.Database, _
                                          ByVal startDate As Date, _
                                          ByVal endDate As Date) As Long
    Dim lastRow As Long
    lastRow = wsKartei.Cells(wsKartei.Rows.Count, K_ID).End(xlUp).Row
    
    If lastRow < 3 Then
        ProcessAllPendingRecords = 0
        Exit Function
    End If
    
    Dim outputRow As Long
    outputRow = 3  ' Start output from row 3 (row 2 is headers)
    
    Dim processedCount As Long
    processedCount = 0
    
    Dim r As Long
    For r = 3 To lastRow
        ' Get ID from current row
        Dim idValue As Long
        idValue = valid_GrossGeschichteData.GetIDFromKarteiRow(wsKartei, r)
        
        If idValue > 0 Then
            ' Process this record using the shared database connection
            ProcessSingleRecord wsGross, wsKartei, db, r, idValue, startDate, endDate, outputRow
            outputRow = outputRow + 3  ' Each record takes 3 rows (War, Ist, separator)
            processedCount = processedCount + 1
        End If
    Next r
    
    ProcessAllPendingRecords = processedCount
End Function

' Processes a single pending record: creates War/Ist rows and highlights differences
'
' @param wsGross - Output worksheet for War/Ist pairs
' @param wsKartei - Source worksheet with pending records
' @param db - Already-open DAO.Database connection (avoids open/close per record)
' @param karteiRow - Row index in Kartei sheet
' @param idValue - Record ID to process
' @param startDate - Start of date range for comment aggregation
' @param endDate - End of date range for comment aggregation
' @param outputRow - Row index in output sheet to write to
Private Sub ProcessSingleRecord(ByVal wsGross As Worksheet, _
                                ByVal wsKartei As Worksheet, _
                                ByVal db As DAO.Database, _
                                ByVal karteiRow As Long, _
                                ByVal idValue As Long, _
                                ByVal startDate As Date, _
                                ByVal endDate As Date, _
                                ByVal outputRow As Long)
    
    ' Get pending (Ist) data from Kartei sheet
    Dim arrIst As Variant
    arrIst = valid_GrossGeschichteData.GetPendingRecordByRow(wsKartei, karteiRow)
    
    If IsEmpty(arrIst) Then Exit Sub
    
    ' Get original (War) data from tblKartei using the shared database connection
    ' This uses GetMainRecordByIDFromDatabase which does NOT open/close the database
    Dim arrWar As Variant
    arrWar = valid_GrossGeschichteData.GetMainRecordByIDFromDatabase(db, idValue)
    
    Dim isNewRecord As Boolean
    isNewRecord = IsEmpty(arrWar)
    
    ' Calculate row positions
    Dim rowWar As Long
    Dim rowIst As Long
    Dim rowSeparator As Long
    rowWar = outputRow
    rowIst = outputRow + 1
    rowSeparator = outputRow + 2
    
    ' Write War row (empty if new record)
    If Not isNewRecord Then
        WriteDataRow wsGross, rowWar, arrWar, CStr(idValue)
    Else
        ' Just write the RecordID for the War row (rest stays empty)
        wsGross.Cells(rowWar, G_RECORD_ID).Value = CStr(idValue)
    End If
    
    ' Write Ist row
    WriteDataRow wsGross, rowIst, arrIst, CStr(idValue)
    
    ' Build and write comment for Ist row
    Dim historyText As String
    historyText = SafeString(arrIst(1, K_HISTORY))
    
    Dim commentText As String
    commentText = BuildCommentForID(historyText, startDate, endDate, isNewRecord)
    wsGross.Cells(rowIst, G_COMMENTS).Value = commentText
    
    ' Highlight differences (only if not a new record)
    If Not isNewRecord Then
        HighlightDifferences wsGross, rowWar, rowIst, arrWar, arrIst
    Else
        ' For new records, highlight the entire Ist row with a different color
        HighlightNewRecord wsGross, rowIst
    End If
    
    ' Create separator row
    CreateSeparatorRow wsGross, rowSeparator
    
    ' Add decision dropdown to Ist row
    AddDecisionDropdown wsGross, rowIst
    
    ' Apply borders
    ApplyRowBorders wsGross, rowWar, rowIst
End Sub

' ============================================================
' DATA WRITING
' ============================================================

' Writes data from array to a GrossGeschichte row
Private Sub WriteDataRow(ByVal ws As Worksheet, _
                         ByVal rowNum As Long, _
                         ByVal arr As Variant, _
                         ByVal recordID As String)
    
    ' Map Kartei columns to GrossGeschichte columns
    ws.Cells(rowNum, G_FAMILY_ID).Value = SafeString(arr(1, K_FAMILY_ID))
    ws.Cells(rowNum, G_PARENT).Value = SafeString(arr(1, K_PARENT))
    ws.Cells(rowNum, G_CHILD).Value = SafeString(arr(1, K_CHILD))
    ws.Cells(rowNum, G_BIRTHDATE).Value = FormatAsText(arr(1, K_BIRTHDATE))
    ws.Cells(rowNum, G_ADDRESS).Value = SafeString(arr(1, K_ADDRESS))
    ws.Cells(rowNum, G_PHONE).Value = FormatAsText(arr(1, K_PHONE))
    ws.Cells(rowNum, G_MOBILE).Value = FormatAsText(arr(1, K_MOBILE))
    ws.Cells(rowNum, G_EMAIL).Value = SafeString(arr(1, K_EMAIL))
    ws.Cells(rowNum, G_SUBJECT1).Value = SafeString(arr(1, K_SUBJECT1))
    ws.Cells(rowNum, G_PRICE1).Value = SafeString(arr(1, K_PRICE1))
    ws.Cells(rowNum, G_SUBJECT2).Value = SafeString(arr(1, K_SUBJECT2))
    ws.Cells(rowNum, G_PRICE2).Value = SafeString(arr(1, K_PRICE2))
    
    ' Write months (K_MONTH_START to K_MONTH_END -> G_MONTH_START to G_MONTH_END)
    Dim monthOffset As Long
    For monthOffset = 0 To 11
        Dim karteiMonthCol As Long
        Dim grossMonthCol As Long
        karteiMonthCol = K_MONTH_START + monthOffset
        grossMonthCol = G_MONTH_START + monthOffset
        
        Dim monthVal As Variant
        monthVal = arr(1, karteiMonthCol)
        
        ' Convert to number if possible, otherwise write as-is
        If IsNumeric(monthVal) And Len(Trim(CStr(monthVal))) > 0 Then
            ws.Cells(rowNum, grossMonthCol).Value = CDbl(monthVal)
        Else
            ws.Cells(rowNum, grossMonthCol).Value = monthVal
        End If
    Next monthOffset
    
    ' Write extras
    ws.Cells(rowNum, G_EXTRA1).Value = SafeString(arr(1, K_EXTRA1))
    ws.Cells(rowNum, G_EXTRA2).Value = SafeString(arr(1, K_EXTRA2))
    ws.Cells(rowNum, G_EXTRA3).Value = SafeString(arr(1, K_EXTRA3))
    
    ' Write RecordID (hidden column)
    ws.Cells(rowNum, G_RECORD_ID).Value = recordID
End Sub

' ============================================================
' COMMENT AGGREGATION
' ============================================================

' Builds aggregated comment from history within date range
' If isNew=True, prepends "Neue Schreibung" to indicate a new record
'
' Comment structure:
'   - For new records: "Neue Schreibung; <aggregated reasons>"
'   - For existing records: "<aggregated reasons>"
'   - Reasons are joined with "; "
Private Function BuildCommentForID(ByVal historyText As String, _
                                   ByVal startDate As Date, _
                                   ByVal endDate As Date, _
                                   ByVal isNew As Boolean) As String
    Dim result As String
    result = ""
    
    ' Parse history using valid_ParseHistory
    If Len(Trim(historyText)) > 0 Then
        Dim events As Collection
        On Error Resume Next
        Set events = valid_ParseHistory.ParseHistory(historyText)
        On Error GoTo 0
        
        If Not events Is Nothing Then
            result = AggregateReasonsInRange(events, startDate, endDate)
        End If
    End If
    
    ' Prepend "Neue Schreibung" for new records
    ' This clearly indicates to Superadmin that this is a brand new entry
    If isNew Then
        If Len(result) > 0 Then
            result = "Neue Schreibung; " & result
        Else
            result = "Neue Schreibung"
        End If
    End If
    
    BuildCommentForID = result
End Function

' Aggregates Reason fields from events that fall within the date range
Private Function AggregateReasonsInRange(ByVal events As Collection, _
                                         ByVal startDate As Date, _
                                         ByVal endDate As Date) As String
    Dim reasons As Collection
    Set reasons = New Collection
    
    Dim i As Long
    For i = 1 To events.Count
        Dim evt As Object
        Set evt = events(i)
        
        ' Try to parse the event date
        Dim eventDateStr As String
        eventDateStr = ""
        
        On Error Resume Next
        eventDateStr = CStr(evt("ChangeDate"))
        On Error GoTo 0
        
        If Len(eventDateStr) > 0 Then
            Dim eventDate As Date
            On Error Resume Next
            eventDate = CDate(eventDateStr)
            
            If Err.Number = 0 Then
                ' Check if event falls within range
                If eventDate >= startDate And eventDate <= endDate Then
                    ' Get the reason
                    Dim Reason As String
                    Reason = ""
                    On Error Resume Next
                    Reason = Trim(CStr(evt("Reason")))
                    On Error GoTo 0
                    
                    ' Add non-empty reasons to collection
                    If Len(Reason) > 0 Then
                        reasons.Add Reason
                    End If
                End If
            End If
            Err.Clear
        End If
    Next i
    
    ' Join all reasons with "; "
    If reasons.Count = 0 Then
        AggregateReasonsInRange = ""
        Exit Function
    End If
    
    Dim joined As String
    joined = ""
    
    For i = 1 To reasons.Count
        If i > 1 Then
            joined = joined & "; "
        End If
        joined = joined & reasons(i)
    Next i
    
    AggregateReasonsInRange = joined
End Function

' ============================================================
' DIFFERENCE HIGHLIGHTING
' ============================================================

' Highlights cells in Ist row where values differ from War row
' Only compares tracked fields (same as Export_RiskClassification)
Private Sub HighlightDifferences(ByVal ws As Worksheet, _
                                 ByVal rowWar As Long, _
                                 ByVal rowIst As Long, _
                                 ByVal arrWar As Variant, _
                                 ByVal arrIst As Variant)
    
    ' Compare and highlight each tracked field
    ' FamilyID
    CompareAndHighlight ws, rowIst, G_FAMILY_ID, arrWar(1, K_FAMILY_ID), arrIst(1, K_FAMILY_ID)
    
    ' Parent
    CompareAndHighlight ws, rowIst, G_PARENT, arrWar(1, K_PARENT), arrIst(1, K_PARENT)
    
    ' Child
    CompareAndHighlight ws, rowIst, G_CHILD, arrWar(1, K_CHILD), arrIst(1, K_CHILD)
    
    ' Birthdate
    CompareAndHighlight ws, rowIst, G_BIRTHDATE, arrWar(1, K_BIRTHDATE), arrIst(1, K_BIRTHDATE)
    
    ' Address
    CompareAndHighlight ws, rowIst, G_ADDRESS, arrWar(1, K_ADDRESS), arrIst(1, K_ADDRESS)
    
    ' Phone
    CompareAndHighlight ws, rowIst, G_PHONE, arrWar(1, K_PHONE), arrIst(1, K_PHONE)
    
    ' Mobile
    CompareAndHighlight ws, rowIst, G_MOBILE, arrWar(1, K_MOBILE), arrIst(1, K_MOBILE)
    
    ' Email
    CompareAndHighlight ws, rowIst, G_EMAIL, arrWar(1, K_EMAIL), arrIst(1, K_EMAIL)
    
    ' Subject1
    CompareAndHighlight ws, rowIst, G_SUBJECT1, arrWar(1, K_SUBJECT1), arrIst(1, K_SUBJECT1)
    
    ' Price1
    CompareAndHighlight ws, rowIst, G_PRICE1, arrWar(1, K_PRICE1), arrIst(1, K_PRICE1)
    
    ' Subject2
    CompareAndHighlight ws, rowIst, G_SUBJECT2, arrWar(1, K_SUBJECT2), arrIst(1, K_SUBJECT2)
    
    ' Price2
    CompareAndHighlight ws, rowIst, G_PRICE2, arrWar(1, K_PRICE2), arrIst(1, K_PRICE2)
    
    ' Months (12 columns)
    Dim monthOffset As Long
    For monthOffset = 0 To 11
        Dim karteiCol As Long
        Dim grossCol As Long
        karteiCol = K_MONTH_START + monthOffset
        grossCol = G_MONTH_START + monthOffset
        
        CompareAndHighlightNumeric ws, rowIst, grossCol, arrWar(1, karteiCol), arrIst(1, karteiCol)
    Next monthOffset
    
    ' Extra1-3
    CompareAndHighlight ws, rowIst, G_EXTRA1, arrWar(1, K_EXTRA1), arrIst(1, K_EXTRA1)
    CompareAndHighlight ws, rowIst, G_EXTRA2, arrWar(1, K_EXTRA2), arrIst(1, K_EXTRA2)
    CompareAndHighlight ws, rowIst, G_EXTRA3, arrWar(1, K_EXTRA3), arrIst(1, K_EXTRA3)
End Sub

' Compares two string values and highlights the cell if different
Private Sub CompareAndHighlight(ByVal ws As Worksheet, _
                                ByVal rowNum As Long, _
                                ByVal colNum As Long, _
                                ByVal warValue As Variant, _
                                ByVal istValue As Variant)
    
    Dim warStr As String
    Dim istStr As String
    
    warStr = Trim(SafeString(warValue))
    istStr = Trim(SafeString(istValue))
    
    If warStr <> istStr Then
        ws.Cells(rowNum, colNum).Interior.Color = COLOR_CHANGED
    End If
End Sub

' Compares two numeric values and highlights the cell if different
' Handles empty/null values as 0 for comparison
Private Sub CompareAndHighlightNumeric(ByVal ws As Worksheet, _
                                       ByVal rowNum As Long, _
                                       ByVal colNum As Long, _
                                       ByVal warValue As Variant, _
                                       ByVal istValue As Variant)
    
    Dim warNum As Double
    Dim istNum As Double
    
    warNum = SafeNumeric(warValue)
    istNum = SafeNumeric(istValue)
    
    ' Use small tolerance for floating point comparison
    If Abs(warNum - istNum) > 0.001 Then
        ws.Cells(rowNum, colNum).Interior.Color = COLOR_CHANGED
    End If
End Sub

' Highlights entire Ist row for new records (no War data)
Private Sub HighlightNewRecord(ByVal ws As Worksheet, ByVal rowIst As Long)
    ' Highlight the comment cell with green to indicate new record
    ws.Cells(rowIst, G_COMMENTS).Interior.Color = COLOR_NEW_RECORD
    
    ' Also highlight all data cells (since everything is "new")
    ws.Range(ws.Cells(rowIst, G_FAMILY_ID), ws.Cells(rowIst, G_EXTRA3)).Interior.Color = COLOR_CHANGED
End Sub

' ============================================================
' FORMATTING
' ============================================================

' Creates a thin separator row between record blocks
Private Sub CreateSeparatorRow(ByVal ws As Worksheet, ByVal rowNum As Long)
    ws.Rows(rowNum).RowHeight = ws.StandardHeight * 0.25
    ws.Range(ws.Cells(rowNum, 1), ws.Cells(rowNum, G_RECORD_ID)).Interior.Color = RGB(192, 192, 192)
End Sub

' Adds Approved/Declined dropdown to decision cell
Private Sub AddDecisionDropdown(ByVal ws As Worksheet, ByVal rowIst As Long)
    With ws.Cells(rowIst, G_DECISION).Validation
        .Delete
        .Add Type:=xlValidateList, AlertStyle:=xlValidAlertStop, _
             Formula1:="Approved,Declined"
        .IgnoreBlank = True
        .InCellDropdown = True
    End With
    
    ' Style the decision cell
    ws.Cells(rowIst, G_DECISION).Interior.Color = RGB(255, 255, 204)  ' Light yellow
    
    ' Style the decline comment cell
    ws.Cells(rowIst, G_DECLINE_COMMENT).Interior.Color = RGB(255, 230, 230)  ' Light pink
End Sub

' Applies borders to War/Ist row pair
Private Sub ApplyRowBorders(ByVal ws As Worksheet, ByVal rowWar As Long, ByVal rowIst As Long)
    With ws.Range(ws.Cells(rowWar, 1), ws.Cells(rowIst, G_DECLINE_COMMENT)).Borders
        .LineStyle = xlContinuous
        .Color = vbBlack
        .Weight = xlThin
    End With
End Sub

' Applies final formatting to the entire sheet
Private Sub ApplyFinalFormatting(ByVal ws As Worksheet)
    ' Find last row with data
    Dim lastRow As Long
    lastRow = ws.Cells(ws.Rows.Count, G_RECORD_ID).End(xlUp).Row
    
    If lastRow < 3 Then Exit Sub
    
    ' Format month columns as numbers with 2 decimal places
    ws.Range("M3:X" & lastRow).NumberFormat = "0.00"
    
    ' Apply AutoFilter to headers
    If ws.AutoFilterMode Then
        ws.AutoFilterMode = False
    End If
    
    Dim filterRange As Range
    Set filterRange = ws.Range("A2:AD" & lastRow)
    filterRange.AutoFilter
    
    ' Ensure date format is preserved in B1:C1
    ws.Range("B1:C1").NumberFormat = "dd.mm.yyyy"
End Sub

' ============================================================
' UTILITY FUNCTIONS
' ============================================================

' Safely converts a value to string, handling Null/Empty
Private Function SafeString(ByVal v As Variant) As String
    If IsNull(v) Or IsEmpty(v) Then
        SafeString = ""
    Else
        SafeString = CStr(v)
    End If
End Function

' Safely converts a value to Double, handling Null/Empty/non-numeric
Private Function SafeNumeric(ByVal v As Variant) As Double
    If IsNull(v) Or IsEmpty(v) Then
        SafeNumeric = 0
        Exit Function
    End If
    
    Dim strVal As String
    strVal = Trim(CStr(v))
    
    If Len(strVal) = 0 Then
        SafeNumeric = 0
        Exit Function
    End If
    
    ' Handle decimal separator conversion
    strVal = Replace(strVal, " ", "")
    
    Dim decSep As String
    decSep = Application.International(xlDecimalSeparator)
    
    If decSep = "," Then
        strVal = Replace(strVal, ".", ",")
    Else
        strVal = Replace(strVal, ",", ".")
    End If
    
    If IsNumeric(strVal) Then
        SafeNumeric = CDbl(strVal)
    Else
        SafeNumeric = 0
    End If
End Function

' Formats value as text, preventing Excel auto-conversion
' Handles dates specially to preserve dd.mm.yyyy format
Private Function FormatAsText(ByVal Value As Variant) As String
    If IsEmpty(Value) Or IsNull(Value) Then
        FormatAsText = ""
    ElseIf IsDate(Value) Then
        FormatAsText = Format(Value, "dd.mm.yyyy")
    Else
        FormatAsText = CStr(Value)
    End If
End Function
