Attribute VB_Name = "valid_Dashboard"
'==========================
'   Module: valid_Dashboard
'   Purpose: Dashboard UI helpers for multi-year Superadmin workbook
'
'   Provides button-friendly macros and a unified interface for:
'   - Year selection (24/25/26)
'   - Loading pending changes
'   - Building decision sheets
'   - Syncing decisions
'   - Running history reports
'   - Running payment exports
'   - Refreshing Neu list (2026 only)
'
'   Usage:
'   1. Create a "Dashboard" sheet in the workbook
'   2. Add buttons or use the Developer > Macros menu
'   3. Assign macros from this module to buttons
'
'   Sheet layout suggestion:
'     Row 1: Title "Superadmin Multi-Year Dashboard"
'     Row 3-5: Year selection area (B3:D5 with 24/25/26 buttons)
'     Row 7+: Action buttons in a grid layout
'
'   Dependencies:
'     valid_ApproveFlow, valid_HistoryPerYear, pay_Main, valid_NeuList, valid_YearConfig
'==========================

Option Explicit

' ============================================================
' DASHBOARD CONSTANTS
' ============================================================

Private Const DASHBOARD_SHEET_NAME As String = "Dashboard"
Private Const DASHBOARD_TITLE As String = "Superadmin Multi-Jahren Dashboard"
Private Const SELECTED_YEAR_CELL As String = "B2"  ' Cell storing selected year (24/25/26)

' Date range input cells for Dashboard_DateRangeHistory
Private Const DATE_START_DAY As String = "E4"
Private Const DATE_START_MONTH As String = "F4"
Private Const DATE_START_YEAR As String = "G4"
Private Const DATE_END_DAY As String = "H4"
Private Const DATE_END_MONTH As String = "I4"
Private Const DATE_END_YEAR As String = "J4"

' ============================================================
' DASHBOARD DATE RANGE HELPERS
' ============================================================

' Try to read and validate date range from Dashboard cells E4:J4.
' Returns True if valid dates were obtained, False otherwise.
'
' Behavior:
'   - If inputs are empty: returns False with empty errorMessage (optional mode)
'   - If inputs are invalid: returns False with errorMessage set
'   - If inputs are valid: returns True with startDate/endDate populated
'
' @param startDate - Output: validated start date
' @param endDate - Output: validated end date
' @param errorMessage - Output: error message if validation fails (empty if inputs are blank)
' @return True if valid date range obtained, False if empty or invalid
Private Function TryGetDashboardDateRange(ByRef startDate As Date, ByRef endDate As Date, ByRef errorMessage As String) As Boolean
    On Error GoTo ErrorHandler
    
    TryGetDashboardDateRange = False
    errorMessage = ""
    
    Dim wsDashboard As Worksheet
    On Error Resume Next
    Set wsDashboard = ThisWorkbook.Worksheets(DASHBOARD_SHEET_NAME)
    On Error GoTo ErrorHandler
    
    If wsDashboard Is Nothing Then
        errorMessage = "Dashboard-Blatt nicht gefunden."
        Exit Function
    End If
    
    ' Read date components from Dashboard
    Dim startDay As Variant, startMonth As Variant, startYear As Variant
    Dim endDay As Variant, endMonth As Variant, endYear As Variant
    
    startDay = wsDashboard.Range(DATE_START_DAY).Value
    startMonth = wsDashboard.Range(DATE_START_MONTH).Value
    startYear = wsDashboard.Range(DATE_START_YEAR).Value
    endDay = wsDashboard.Range(DATE_END_DAY).Value
    endMonth = wsDashboard.Range(DATE_END_MONTH).Value
    endYear = wsDashboard.Range(DATE_END_YEAR).Value
    
    ' Check if all inputs are empty (optional mode - no error)
    Dim allEmpty As Boolean
    allEmpty = (IsEmpty(startDay) Or startDay = "") And _
               (IsEmpty(startMonth) Or startMonth = "") And _
               (IsEmpty(startYear) Or startYear = "") And _
               (IsEmpty(endDay) Or endDay = "") And _
               (IsEmpty(endMonth) Or endMonth = "") And _
               (IsEmpty(endYear) Or endYear = "")
    
    If allEmpty Then
        ' Empty inputs = optional mode, no error
        Exit Function
    End If
    
    ' Validate all components are numeric
    If Not IsNumeric(startDay) Or Not IsNumeric(startMonth) Or Not IsNumeric(startYear) Then
        errorMessage = "Bitte gueltige Zahlenwerte fuer das Startdatum eingeben (E4=Tag, F4=Monat, G4=Jahr)." & vbCrLf & vbCrLf & _
                       "Aktuell: Tag=" & CStr(startDay) & ", Monat=" & CStr(startMonth) & ", Jahr=" & CStr(startYear)
        Exit Function
    End If
    
    If Not IsNumeric(endDay) Or Not IsNumeric(endMonth) Or Not IsNumeric(endYear) Then
        errorMessage = "Bitte gueltige Zahlenwerte fuer das Enddatum eingeben (H4=Tag, I4=Monat, J4=Jahr)." & vbCrLf & vbCrLf & _
                       "Aktuell: Tag=" & CStr(endDay) & ", Monat=" & CStr(endMonth) & ", Jahr=" & CStr(endYear)
        Exit Function
    End If
    
    ' Convert to integers
    Dim iStartDay As Integer, iStartMonth As Integer, iStartYear As Integer
    Dim iEndDay As Integer, iEndMonth As Integer, iEndYear As Integer
    
    iStartDay = CInt(startDay)
    iStartMonth = CInt(startMonth)
    iStartYear = CInt(startYear)
    iEndDay = CInt(endDay)
    iEndMonth = CInt(endMonth)
    iEndYear = CInt(endYear)
    
    ' Validate date components are in valid ranges
    If iStartMonth < 1 Or iStartMonth > 12 Then
        errorMessage = "Startmonat (F4) muss zwischen 1 und 12 liegen. Aktuell: " & iStartMonth
        Exit Function
    End If
    
    If iEndMonth < 1 Or iEndMonth > 12 Then
        errorMessage = "Endmonat (I4) muss zwischen 1 und 12 liegen. Aktuell: " & iEndMonth
        Exit Function
    End If
    
    If iStartDay < 1 Or iStartDay > 31 Then
        errorMessage = "Starttag (E4) muss zwischen 1 und 31 liegen. Aktuell: " & iStartDay
        Exit Function
    End If
    
    If iEndDay < 1 Or iEndDay > 31 Then
        errorMessage = "Endtag (H4) muss zwischen 1 und 31 liegen. Aktuell: " & iEndDay
        Exit Function
    End If
    
    ' Build dates using DateSerial
    Dim dStart As Date, dEnd As Date
    dStart = DateSerial(iStartYear, iStartMonth, iStartDay)
    dEnd = DateSerial(iEndYear, iEndMonth, iEndDay)
    
    ' Check for DateSerial rollover (invalid date input, e.g. Feb 30)
    If Day(dStart) <> iStartDay Or Month(dStart) <> iStartMonth Or Year(dStart) <> iStartYear Then
        errorMessage = "Das Startdatum ist ungueltig (z.B. 31. Februar)." & vbCrLf & vbCrLf & _
                       "Eingabe: " & iStartDay & "." & iStartMonth & "." & iStartYear & vbCrLf & _
                       "Interpretiert als: " & Format(dStart, "dd.mm.yyyy")
        Exit Function
    End If
    
    If Day(dEnd) <> iEndDay Or Month(dEnd) <> iEndMonth Or Year(dEnd) <> iEndYear Then
        errorMessage = "Das Enddatum ist ungueltig (z.B. 31. Februar)." & vbCrLf & vbCrLf & _
                       "Eingabe: " & iEndDay & "." & iEndMonth & "." & iEndYear & vbCrLf & _
                       "Interpretiert als: " & Format(dEnd, "dd.mm.yyyy")
        Exit Function
    End If
    
    ' Validate startDate <= endDate (do not swap silently)
    If dStart > dEnd Then
        errorMessage = "Das Startdatum muss vor oder gleich dem Enddatum liegen." & vbCrLf & vbCrLf & _
                       "Startdatum: " & Format(dStart, "dd.mm.yyyy") & vbCrLf & _
                       "Enddatum: " & Format(dEnd, "dd.mm.yyyy")
        Exit Function
    End If
    
    ' All validations passed
    startDate = dStart
    endDate = dEnd
    TryGetDashboardDateRange = True
    Exit Function
    
ErrorHandler:
    errorMessage = "Fehler beim Lesen des Datumsbereichs: " & Err.Description
End Function

' Apply Dashboard date range to all relevant year sheets (B1/C1 cells).
' Sets and formats dates on:
'   - grossGeschichteYY (decision sheet / comment aggregation window)
'   - GeschichteYY (single-record history date filter)
'   - GeschichteYY_Alle (all-records history date filter)
'
' @param year2 - Two-digit year (24, 25, or 26)
' @param startDate - Start date to apply
' @param endDate - End date to apply
Private Sub ApplyDashboardDateRangeToYearSheets(ByVal year2 As Integer, ByVal startDate As Date, ByVal endDate As Date)
    On Error Resume Next
    
    ' Ensure year sheets exist
    valid_YearConfig.EnsureYearSheetsExist year2
    
    Dim ws As Worksheet
    
    ' 1. grossGeschichteYY
    Set ws = ThisWorkbook.Worksheets(valid_YearConfig.GetGrossGeschichteSheetName(year2))
    If Not ws Is Nothing Then
        ws.Range("B1").NumberFormat = "dd.mm.yyyy"
        ws.Range("C1").NumberFormat = "dd.mm.yyyy"
        ws.Range("B1").Value = startDate
        ws.Range("C1").Value = endDate
    End If
    
    ' 2. GeschichteYY (single-record history)
    Set ws = ThisWorkbook.Worksheets(valid_YearConfig.GetGeschichteSheetName(year2))
    If Not ws Is Nothing Then
        ws.Range("B1").NumberFormat = "dd.mm.yyyy"
        ws.Range("C1").NumberFormat = "dd.mm.yyyy"
        ws.Range("B1").Value = startDate
        ws.Range("C1").Value = endDate
    End If
    
    ' 3. GeschichteYY_Alle (all-records history)
    Set ws = ThisWorkbook.Worksheets(valid_YearConfig.GetGeschichteAlleSheetName(year2))
    If Not ws Is Nothing Then
        ws.Range("B1").NumberFormat = "dd.mm.yyyy"
        ws.Range("C1").NumberFormat = "dd.mm.yyyy"
        ws.Range("B1").Value = startDate
        ws.Range("C1").Value = endDate
    End If
    
    On Error GoTo 0
End Sub

' ============================================================
' YEAR SELECTION
' ============================================================

' Get the currently selected year from Dashboard
' Returns 25 as default if not set or invalid
Public Function GetSelectedYear() As Integer
    On Error Resume Next
    
    Dim wsDashboard As Worksheet
    Set wsDashboard = ThisWorkbook.Worksheets(DASHBOARD_SHEET_NAME)
    
    If wsDashboard Is Nothing Then
        ' No Dashboard sheet - default to 25
        GetSelectedYear = 25
        Exit Function
    End If
    
    Dim yearVal As Variant
    yearVal = wsDashboard.Range(SELECTED_YEAR_CELL).Value
    
    If IsNumeric(yearVal) Then
        Dim year2 As Integer
        year2 = CInt(yearVal)
        
        If valid_YearConfig.IsValidYear(year2) Then
            GetSelectedYear = year2
            Exit Function
        End If
    End If
    
    ' Default to 25 if invalid
    GetSelectedYear = 25
    On Error GoTo 0
End Function

' Set the selected year on Dashboard
Public Sub SetSelectedYear(ByVal year2 As Integer)
    On Error Resume Next
    
    If Not valid_YearConfig.IsValidYear(year2) Then
        Exit Sub
    End If
    
    Dim wsDashboard As Worksheet
    Set wsDashboard = GetOrCreateDashboard()
    
    If Not wsDashboard Is Nothing Then
        wsDashboard.Range(SELECTED_YEAR_CELL).Value = year2
        UpdateYearDisplay wsDashboard, year2
        
        ' Auto-fill year cells for date range (G4 and J4)
        Dim fullYear As Long
        fullYear = 2000 + year2
        wsDashboard.Range(DATE_START_YEAR).Value = fullYear
        wsDashboard.Range(DATE_END_YEAR).Value = fullYear
    End If
    
    On Error GoTo 0
End Sub

' Button handler: Select year 2024
Public Sub SelectYear24()
    SetSelectedYear 24
    MsgBox "Jahr 2024 ausgewaehlt." & vbCrLf & vbCrLf & _
           "Alle Aktionen werden jetzt auf Kartei24, grossGeschichte24, etc. angewendet.", _
           vbInformation, "Jahr 2024"
End Sub

' Button handler: Select year 2025
Public Sub SelectYear25()
    SetSelectedYear 25
    MsgBox "Jahr 2025 ausgewaehlt." & vbCrLf & vbCrLf & _
           "Alle Aktionen werden jetzt auf Kartei25, grossGeschichte25, etc. angewendet.", _
           vbInformation, "Jahr 2025"
End Sub

' Button handler: Select year 2026
Public Sub SelectYear26()
    SetSelectedYear 26
    MsgBox "Jahr 2026 ausgewaehlt." & vbCrLf & vbCrLf & _
           "Alle Aktionen werden jetzt auf Kartei26, grossGeschichte26, etc. angewendet.", _
           vbInformation, "Jahr 2026"
End Sub

' ============================================================
' ACTION DISPATCHERS (use selected year)
' ============================================================

' Load pending changes for selected year
' Applies Dashboard date range to year sheets if valid, then loads pending and builds decision sheet
Public Sub Dashboard_LoadPending()
    Dim year2 As Integer
    year2 = GetSelectedYear()
    
    ' Try to read Dashboard date range
    Dim startDate As Date, endDate As Date, errorMessage As String
    Dim hasRange As Boolean
    hasRange = TryGetDashboardDateRange(startDate, endDate, errorMessage)
    
    ' If invalid (not empty), show error and stop
    If errorMessage <> "" Then
        MsgBox errorMessage, vbExclamation, "Ungueltiger Datumsbereich"
        Exit Sub
    End If
    
    ' If valid date range provided, apply to year sheets
    If hasRange Then
        ApplyDashboardDateRangeToYearSheets year2, startDate, endDate
    End If
    
    valid_ApproveFlow.LoadPendingAndBuildDecisionForYear year2
    
    ' Ensure AutoFilter arrows on KarteiYY header row
    Dim wsKartei As Worksheet
    On Error Resume Next
    Set wsKartei = ThisWorkbook.Worksheets(valid_YearConfig.GetKarteiSheetName(year2))
    On Error GoTo 0
    If Not wsKartei Is Nothing Then
        EnsureKarteiAutoFilter wsKartei
    End If
End Sub

' Build decision sheet for selected year
' Applies Dashboard date range to year sheets if valid, then builds decision sheet
Public Sub Dashboard_BuildDecisionSheet()
    Dim year2 As Integer
    year2 = GetSelectedYear()
    
    ' Try to read Dashboard date range
    Dim startDate As Date, endDate As Date, errorMessage As String
    Dim hasRange As Boolean
    hasRange = TryGetDashboardDateRange(startDate, endDate, errorMessage)
    
    ' If invalid (not empty), show error and stop
    If errorMessage <> "" Then
        MsgBox errorMessage, vbExclamation, "Ungueltiger Datumsbereich"
        Exit Sub
    End If
    
    ' If valid date range provided, apply to year sheets
    If hasRange Then
        ApplyDashboardDateRangeToYearSheets year2, startDate, endDate
    End If
    
    valid_GrossGeschichteDecision.BuildPendingDecisionSheetForYear year2
End Sub

' Sync decisions for selected year
Public Sub Dashboard_SyncDecisions()
    Dim year2 As Integer
    year2 = GetSelectedYear()
    
    valid_ApproveFlow.SyncDecisionsForYear year2
End Sub

' Single-record history for selected year
' Applies Dashboard date range to year sheets if valid, then generates single-record history
Public Sub Dashboard_SingleRecordHistory()
    Dim year2 As Integer
    year2 = GetSelectedYear()
    
    ' Try to read Dashboard date range
    Dim startDate As Date, endDate As Date, errorMessage As String
    Dim hasRange As Boolean
    hasRange = TryGetDashboardDateRange(startDate, endDate, errorMessage)
    
    ' If invalid (not empty), show error and stop
    If errorMessage <> "" Then
        MsgBox errorMessage, vbExclamation, "Ungueltiger Datumsbereich"
        Exit Sub
    End If
    
    ' If valid date range provided, apply to year sheets
    If hasRange Then
        ApplyDashboardDateRangeToYearSheets year2, startDate, endDate
    End If
    
    valid_HistoryPerYear.GeschichteMachenForYear year2
End Sub

' Date-range history for selected year
' Reads date range from Dashboard cells E4:J4, validates, then runs the report
' Uses shared TryGetDashboardDateRange and ApplyDashboardDateRangeToYearSheets helpers
Public Sub Dashboard_DateRangeHistory()
    On Error GoTo ErrorHandler
    
    Dim year2 As Integer
    year2 = GetSelectedYear()
    
    ' Try to read Dashboard date range (required for this action)
    Dim startDate As Date, endDate As Date, errorMessage As String
    Dim hasRange As Boolean
    hasRange = TryGetDashboardDateRange(startDate, endDate, errorMessage)
    
    ' If invalid (not empty), show error and stop
    If errorMessage <> "" Then
        MsgBox errorMessage, vbExclamation, "Ungueltiger Datumsbereich"
        Exit Sub
    End If
    
    ' For date-range history, a date range is REQUIRED
    If Not hasRange Then
        MsgBox "Bitte einen gueltigen Datumsbereich eingeben (E4:J4).", _
               vbExclamation, "Datumsbereich erforderlich"
        Exit Sub
    End If
    
    ' Apply date range to all year sheets
    ApplyDashboardDateRangeToYearSheets year2, startDate, endDate
    
    ' Now call the history function (it will read from B1/C1)
    valid_HistoryPerYear.GrossGeschichteMachenForYear year2
    Exit Sub
    
ErrorHandler:
    MsgBox "Fehler bei Dashboard_DateRangeHistory: " & Err.Description, _
           vbCritical, "Fehler"
End Sub

' Payment export for selected year
Public Sub Dashboard_FamZahlungen()
    Dim year2 As Integer
    year2 = GetSelectedYear()
    
    pay_Main.FamZahlungenForYear year2
End Sub

' Refresh Neu list (always uses 2026)
Public Sub Dashboard_RefreshNeu()
    ' Neu list is only for year 2026 (no new records expected for past years)
    valid_NeuList.RefreshNeuList
End Sub

' Import full base (tblKartei) for selected year - NO decision sheet
Public Sub Dashboard_ImportFromBase()
    Dim year2 As Integer
    year2 = GetSelectedYear()
    
    valid_ImportFromBase.ImportFromBaseForYear year2
    
    ' Ensure AutoFilter arrows on KarteiYY header row
    Dim wsKartei As Worksheet
    On Error Resume Next
    Set wsKartei = ThisWorkbook.Worksheets(valid_YearConfig.GetKarteiSheetName(year2))
    On Error GoTo 0
    If Not wsKartei Is Nothing Then
        EnsureKarteiAutoFilter wsKartei
    End If
End Sub

' ============================================================
' CONFIGURATION ACTIONS
' ============================================================

' Open database configuration for all years
Public Sub Dashboard_ConfigureDatabase()
    valid_YearConfig.ConfigureAllDatabases
End Sub

' Show current configuration status
Public Sub Dashboard_ShowConfig()
    Dim msg As String
    msg = "=== Datenbank-Konfiguration ===" & vbCrLf & vbCrLf
    
    Dim cloudRoot As String
    cloudRoot = valid_YearConfig.GetCloudRoot()
    If cloudRoot <> "" Then
        msg = msg & "CloudRoot: " & cloudRoot & vbCrLf & vbCrLf
    Else
        msg = msg & "CloudRoot: (nicht gesetzt)" & vbCrLf & vbCrLf
    End If
    
    Dim years As Variant
    years = valid_YearConfig.GetSupportedYears()
    
    Dim i As Long
    For i = LBound(years) To UBound(years)
        Dim year2 As Integer
        year2 = CInt(years(i))
        
        Dim overridePath As String
        overridePath = valid_YearConfig.GetYearOverrideFromConfig(year2)
        
        Dim status As String
        Dim expectedDbPath As String
        expectedDbPath = valid_YearConfig.TryGetDbPathForYearNoPrompt(year2, status)
        
        msg = msg & "Jahr 20" & Format(year2, "00") & ":" & vbCrLf
        
        If overridePath <> "" Then
            msg = msg & "  Override: " & overridePath & vbCrLf
        Else
            msg = msg & "  Override: (keine)" & vbCrLf
        End If
        
        Select Case status
            Case "ok"
                msg = msg & "  DB: " & expectedDbPath & vbCrLf
            Case "missing-file"
                msg = msg & "  DB: " & expectedDbPath & vbCrLf & _
                            "  Status: Datei nicht gefunden (bitte konfigurieren)" & vbCrLf
            Case "not-configured"
                msg = msg & "  DB: (nicht konfiguriert)" & vbCrLf
            Case Else
                If expectedDbPath <> "" Then
                    msg = msg & "  DB: " & expectedDbPath & vbCrLf
                Else
                    msg = msg & "  DB: (unbekannt)" & vbCrLf
                End If
                msg = msg & "  Status: " & status & vbCrLf
        End Select
        
        msg = msg & vbCrLf
    Next i
    
    msg = msg & vbCrLf & "=== Blattnamen ===" & vbCrLf & vbCrLf
    msg = msg & "Kartei: Kartei24, Kartei25, Kartei26" & vbCrLf
    msg = msg & "Entscheidung: grossGeschichte24/25/26" & vbCrLf
    msg = msg & "Geschichte (Einzel): Geschichte24/25/26" & vbCrLf
    msg = msg & "Geschichte (Alle): Geschichte24_Alle/25_Alle/26_Alle" & vbCrLf
    msg = msg & "Neu-Liste: Neu (nur 2026)" & vbCrLf
    
    msg = msg & vbCrLf & "=== Aktuell ausgewaehlt ===" & vbCrLf & vbCrLf
    msg = msg & "Jahr: 20" & Format(GetSelectedYear(), "00")
    
    MsgBox msg, vbInformation, "Superadmin Multi-Year Konfiguration"
End Sub

' Open the Russian user guide PDF from the workbook folder
Public Sub Dashboard_OpenGuideRU()
    On Error GoTo ErrorHandler
    
    Dim pdfPath As String
    pdfPath = ThisWorkbook.Path & "\SUPERADMIN_MULTIYEAR_GUIDE_RU.pdf"
    
    ' Check if file exists
    If Dir(pdfPath) = "" Then
        MsgBox "PDF-Datei nicht gefunden:" & vbCrLf & vbCrLf & pdfPath, vbExclamation, "Datei nicht gefunden"
        Exit Sub
    End If
    
    ' Open PDF with default application
    Shell "explorer.exe """ & pdfPath & """", vbNormalFocus
    Exit Sub
    
ErrorHandler:
    MsgBox "Fehler beim Oeffnen der PDF-Datei:" & vbCrLf & Err.Description, vbCritical, "Fehler"
End Sub

' ============================================================
' DASHBOARD SHEET MANAGEMENT
' ============================================================

' Get or create the Dashboard sheet
Public Function GetOrCreateDashboard() As Worksheet
    On Error Resume Next
    
    Dim wsDashboard As Worksheet
    Set wsDashboard = ThisWorkbook.Worksheets(DASHBOARD_SHEET_NAME)
    
    If wsDashboard Is Nothing Then
        On Error GoTo 0
        Set wsDashboard = CreateDashboardSheet()
    Else
        On Error GoTo 0
        ' Ensure layout exists even if the sheet was created manually/empty.
        EnsureDashboardLayout wsDashboard
    End If
    
    Set GetOrCreateDashboard = wsDashboard
    On Error GoTo 0
End Function

' Create the Dashboard sheet with layout and instructions
Private Function CreateDashboardSheet() As Worksheet
    On Error GoTo ErrorHandler
    
    Application.ScreenUpdating = False
    
    Dim ws As Worksheet
    Set ws = ThisWorkbook.Worksheets.Add(After:=ThisWorkbook.Worksheets(ThisWorkbook.Worksheets.Count))
    ws.Name = DASHBOARD_SHEET_NAME
    
    EnsureDashboardLayout ws
    
    Application.ScreenUpdating = True
    
    Set CreateDashboardSheet = ws
    Exit Function
    
ErrorHandler:
    Application.ScreenUpdating = True
    Dim errNumber As Long
    errNumber = Err.Number
    Dim errDescription As String
    errDescription = Err.Description
    
    MsgBox "Fehler beim Erstellen des Dashboards (" & errNumber & "): " & errDescription, vbCritical, "Fehler"
    Set CreateDashboardSheet = Nothing
End Function

' Ensure the Dashboard sheet has the expected layout.
' Safe to run on an existing sheet (does not delete shapes/buttons).
Private Sub EnsureDashboardLayout(ByVal ws As Worksheet)
    Dim stage As String
    stage = "start"
    
    On Error GoTo ErrorHandler
    
    If ws Is Nothing Then Exit Sub
    
    Dim oldScreenUpdating As Boolean
    oldScreenUpdating = Application.ScreenUpdating
    
    Dim oldEnableEvents As Boolean
    oldEnableEvents = Application.EnableEvents
    
    Application.ScreenUpdating = False
    Application.EnableEvents = False
    
    ' If the sheet is protected, initialization may fail (1004). Try to unprotect.
    If ws.ProtectContents Then
        stage = "unprotect"
        On Error Resume Next
        ws.Unprotect
        On Error GoTo ErrorHandler
        
        If ws.ProtectContents Then
            Application.ScreenUpdating = True
            MsgBox "Dashboard kann nicht initialisiert werden, weil das Blatt geschuetzt ist." & vbCrLf & vbCrLf & _
                   "Bitte das Blatt 'Dashboard' entsperren (Review -> Unprotect Sheet) und erneut 'ShowDashboard' starten.", _
                   vbExclamation, "Dashboard gesperrt"
            Exit Sub
        End If
    End If
    
    ' Preserve selected year if already set
    Dim selectedYear As Integer
    selectedYear = 25
    If IsNumeric(ws.Range(SELECTED_YEAR_CELL).Value) Then
        If valid_YearConfig.IsValidYear(CInt(ws.Range(SELECTED_YEAR_CELL).Value)) Then
            selectedYear = CInt(ws.Range(SELECTED_YEAR_CELL).Value)
        End If
    End If
    
    ' Title
    stage = "title"
    Call SetCellValueSafe(ws, "A1", DASHBOARD_TITLE)
    On Error Resume Next
    ws.Range("A1").Font.Bold = True
    ws.Range("A1").Font.Size = 16
    On Error GoTo ErrorHandler
    
    ' Year selection label
    stage = "year-cell"
    Call SetCellValueSafe(ws, "A2", "Ausgewaehltes Jahr:")
    Call SetCellValueSafe(ws, SELECTED_YEAR_CELL, selectedYear)
    On Error Resume Next
    ws.Range("A2").Font.Bold = True
    ws.Range(SELECTED_YEAR_CELL).Font.Size = 14
    ws.Range(SELECTED_YEAR_CELL).Font.Bold = True
    ws.Range(SELECTED_YEAR_CELL).HorizontalAlignment = xlCenter
    On Error GoTo ErrorHandler
    
    ' Try to color-code the year cell (non-critical)
    stage = "year-color"
    On Error Resume Next
    UpdateYearDisplay ws, selectedYear
    On Error GoTo ErrorHandler
    
    ' Instructions
    stage = "instructions-a4"
    Call SetCellValueSafe(ws, "A4", "=== Jahres-Auswahl ===")
    stage = "instructions-a5"
    Call SetCellValueSafe(ws, "A5", "Makro: SelectYear24, SelectYear25, SelectYear26")
    stage = "instructions-a6"
    Call SetCellValueSafe(ws, "A6", "(Waehlen Sie ein Jahr, bevor Sie Aktionen ausfuehren)")
    
    ' Date range section for Dashboard_DateRangeHistory (E4:J4)
    ' Only write labels/headers - do not overwrite user-entered day/month values
    stage = "date-range-labels"
    Call SetCellValueSafe(ws, "D3", "Geschichte-Zeitraum:")
    Call SetCellValueSafe(ws, "E3", "Tag")
    Call SetCellValueSafe(ws, "F3", "Monat")
    Call SetCellValueSafe(ws, "G3", "Jahr")
    Call SetCellValueSafe(ws, "H3", "Tag")
    Call SetCellValueSafe(ws, "I3", "Monat")
    Call SetCellValueSafe(ws, "J3", "Jahr")
    Call SetCellValueSafe(ws, "D4", "Von / Bis:")
    
    ' Set year cells to selected year if empty (G4/J4)
    stage = "date-range-years"
    If IsEmpty(ws.Range(DATE_START_YEAR).Value) Or ws.Range(DATE_START_YEAR).Value = "" Then
        ws.Range(DATE_START_YEAR).Value = 2000 + selectedYear
    End If
    If IsEmpty(ws.Range(DATE_END_YEAR).Value) Or ws.Range(DATE_END_YEAR).Value = "" Then
        ws.Range(DATE_END_YEAR).Value = 2000 + selectedYear
    End If
    
    ' Set default day/month values if empty (E4=1, F4=1, H4=current day, I4=current month)
    stage = "date-range-defaults"
    If IsEmpty(ws.Range(DATE_START_DAY).Value) Or ws.Range(DATE_START_DAY).Value = "" Then
        ws.Range(DATE_START_DAY).Value = 1
    End If
    If IsEmpty(ws.Range(DATE_START_MONTH).Value) Or ws.Range(DATE_START_MONTH).Value = "" Then
        ws.Range(DATE_START_MONTH).Value = 1
    End If
    If IsEmpty(ws.Range(DATE_END_DAY).Value) Or ws.Range(DATE_END_DAY).Value = "" Then
        ws.Range(DATE_END_DAY).Value = Day(Date)
    End If
    If IsEmpty(ws.Range(DATE_END_MONTH).Value) Or ws.Range(DATE_END_MONTH).Value = "" Then
        ws.Range(DATE_END_MONTH).Value = Month(Date)
    End If
    
    ' Format date range cells
    On Error Resume Next
    ws.Range("D3:D4").Font.Bold = True
    ws.Range("E3:J3").Font.Bold = True
    ws.Range("E3:J3").Interior.Color = RGB(220, 220, 220)
    ws.Range("E4:J4").NumberFormat = "0"
    ws.Range("E4:J4").HorizontalAlignment = xlCenter
    ws.Range("E3:J4").Borders.LineStyle = xlContinuous
    ws.Columns("D").ColumnWidth = 18
    ws.Columns("E").ColumnWidth = 6
    ws.Columns("F").ColumnWidth = 6
    ws.Columns("G").ColumnWidth = 6
    ws.Columns("H").ColumnWidth = 6
    ws.Columns("I").ColumnWidth = 6
    ws.Columns("J").ColumnWidth = 6
    On Error GoTo ErrorHandler
    
    stage = "instructions-a8"
    Call SetCellValueSafe(ws, "A8", "=== Hauptaktionen ===")
    
    ' Action macros table (write values in one block)
    stage = "actions-table"
    Dim arrActions As Variant
    arrActions = Array( _
        Array("Aktion", "Dashboard-Makro", "Direktes Makro (Jahr)"), _
        Array("Pending laden + Entscheidungsblatt", "Dashboard_LoadPending", "LoadPendingAndBuildDecision24/25/26"), _
        Array("Aktuelle Basis laden (tblKartei, ohne Entscheidung)", "Dashboard_ImportFromBase", "ImportFromBase24/25/26"), _
        Array("Entscheidungsblatt neu erstellen (ohne Import)", "Dashboard_BuildDecisionSheet", "BuildPendingDecisionSheet24/25/26"), _
        Array("Entscheidungen synchronisieren", "Dashboard_SyncDecisions", "SyncDecisions24/25/26"), _
        Array("Einzeldatensatz-Geschichte", "Dashboard_SingleRecordHistory", "GeschichteMachen24/25/26"), _
        Array("Datums-Geschichte (alle)", "Dashboard_DateRangeHistory", "GrossGeschichteMachen24/25/26"), _
        Array("Zahlungsexport", "Dashboard_FamZahlungen", "FamZahlungen24/25/26"), _
        Array("Neu-Liste aktualisieren (nur 2026)", "Dashboard_RefreshNeu", "RefreshNeuList") _
    )
    
    Dim r As Long, c As Long
    For r = LBound(arrActions) To UBound(arrActions)
        For c = LBound(arrActions(r)) To UBound(arrActions(r))
            Call SetCellValueSafe(ws, ws.Cells(9 + r, 1 + c).Address(False, False), arrActions(r)(c))
        Next c
    Next r
    
    On Error Resume Next
    ws.Range("A4").Font.Bold = True
    ws.Range("A8").Font.Bold = True
    ws.Range("A9:C9").Font.Bold = True
    ws.Range("A9:C9").Interior.Color = RGB(220, 220, 220)
    On Error GoTo ErrorHandler
    
    ' Configuration section
    stage = "config"
    Call SetCellValueSafe(ws, "A18", "=== Konfiguration ===")
    Call SetCellValueSafe(ws, "A19", "Datenbank-Pfade konfigurieren")
    Call SetCellValueSafe(ws, "B19", "Dashboard_ConfigureDatabase")
    Call SetCellValueSafe(ws, "A20", "Konfiguration anzeigen")
    Call SetCellValueSafe(ws, "B20", "Dashboard_ShowConfig")
    On Error Resume Next
    ws.Range("A18").Font.Bold = True
    On Error GoTo ErrorHandler
    
    ' Sheet reference
    stage = "sheet-reference"
    Call SetCellValueSafe(ws, "A22", "=== Blatt-Referenz ===")
    Call SetCellValueSafe(ws, "A23", "Pending-Daten: Kartei24, Kartei25, Kartei26")
    Call SetCellValueSafe(ws, "A24", "Entscheidung: grossGeschichte24, grossGeschichte25, grossGeschichte26")
    Call SetCellValueSafe(ws, "A25", "Geschichte (Einzel): Geschichte24, Geschichte25, Geschichte26")
    Call SetCellValueSafe(ws, "A26", "Geschichte (Alle): Geschichte24_Alle, Geschichte25_Alle, Geschichte26_Alle")
    Call SetCellValueSafe(ws, "A27", "Neue Datensaetze: Neu (nur Jahr 2026)")
    Call SetCellValueSafe(ws, "A28", "DB-Konfiguration: DBConfig (versteckt)")
    On Error Resume Next
    ws.Range("A22").Font.Bold = True
    On Error GoTo ErrorHandler
    
    ' Column widths (non-critical)
    stage = "formatting"
    On Error Resume Next
    ws.Columns("A").ColumnWidth = 40
    ws.Columns("B").ColumnWidth = 30
    ws.Columns("C").ColumnWidth = 35
    ws.Range("A9:C16").Borders.LineStyle = xlContinuous
    ws.Range("A19:B20").Borders.LineStyle = xlContinuous
    On Error GoTo ErrorHandler
    
    Application.EnableEvents = oldEnableEvents
    Application.ScreenUpdating = oldScreenUpdating
    Exit Sub
    
ErrorHandler:
    Dim errNumber As Long
    errNumber = Err.Number
    Dim errDescription As String
    errDescription = Err.Description
    
    On Error Resume Next
    Application.EnableEvents = oldEnableEvents
    Application.ScreenUpdating = oldScreenUpdating
    On Error GoTo 0
    
    MsgBox "Fehler beim Initialisieren des Dashboards (" & errNumber & "): " & errDescription & vbCrLf & _
           "Stage: " & stage, vbCritical, "Fehler"
End Sub

' Set a cell value even if it falls inside a merged cell block.
' If the target is within a merged range, writes to the top-left cell of MergeArea.
' IMPORTANT: This sub uses On Error Resume Next throughout and never raises errors
' to avoid interfering with the caller's error handler.
Private Sub SetCellValueSafe(ByVal ws As Worksheet, ByVal address As String, ByVal value As Variant)
    On Error Resume Next
    
    If ws Is Nothing Then Exit Sub
    If address = "" Then Exit Sub
    
    Dim rng As Range
    Set rng = ws.Range(address)
    
    If rng Is Nothing Then Exit Sub
    
    If rng.MergeCells Then
        rng.MergeArea.Cells(1, 1).Value = value
    Else
        rng.Value = value
    End If
    
    ' Clear any error that might have occurred (don't propagate)
    Err.Clear
End Sub

' Update the year display after selection
Private Sub UpdateYearDisplay(ByVal wsDashboard As Worksheet, ByVal year2 As Integer)
    On Error Resume Next
    
    ' Update the selected year cell with color coding
    With wsDashboard.Range(SELECTED_YEAR_CELL)
        .Value = year2
        
        Select Case year2
            Case 24
                .Interior.Color = RGB(255, 230, 200)  ' Light orange for past
            Case 25
                .Interior.Color = RGB(200, 255, 200)  ' Light green for recent
            Case 26
                .Interior.Color = RGB(200, 220, 255)  ' Light blue for current
        End Select
    End With
    
    On Error GoTo 0
End Sub

' ============================================================
' QUICK ACCESS MACROS (for Ribbon/Quick Access Toolbar)
' ============================================================

' Show Dashboard and allow user to select actions
Public Sub ShowDashboard()
    Dim wsDashboard As Worksheet
    Set wsDashboard = GetOrCreateDashboard()
    
    If Not wsDashboard Is Nothing Then
        wsDashboard.Activate
        wsDashboard.Range("A1").Select
    End If
End Sub

' Show year selection dialog and run specified action
Public Sub QuickAction()
    Dim year2 As Integer
    year2 = AskForYear()
    
    If year2 = 0 Then Exit Sub
    
    Dim action As String
    action = AskForAction()
    
    If action = "" Then Exit Sub
    
    Select Case action
        Case "LOAD"
            valid_ApproveFlow.LoadPendingAndBuildDecisionForYear year2
        Case "BASEIMPORT"
            valid_ImportFromBase.ImportFromBaseForYear year2
        Case "SYNC"
            valid_ApproveFlow.SyncDecisionsForYear year2
        Case "HISTORY1"
            valid_HistoryPerYear.GeschichteMachenForYear year2
        Case "HISTORYALL"
            valid_HistoryPerYear.GrossGeschichteMachenForYear year2
        Case "PAYMENT"
            pay_Main.FamZahlungenForYear year2
        Case "NEU"
            If year2 = 26 Then
                valid_NeuList.RefreshNeuList
            Else
                MsgBox "Neu-Liste ist nur fuer Jahr 2026 verfuegbar.", vbInformation, "Hinweis"
            End If
    End Select
End Sub

' Helper: Ask user to select a year
Private Function AskForYear() As Integer
    Dim yearStr As String
    yearStr = InputBox("Jahr auswaehlen (24, 25, oder 26):", "Jahr auswaehlen", CStr(GetSelectedYear()))
    
    If yearStr = "" Then
        AskForYear = 0
        Exit Function
    End If
    
    If Not IsNumeric(yearStr) Then
        MsgBox "Bitte eine gueltige Jahreszahl eingeben (24, 25, oder 26).", vbExclamation, "Ungueltige Eingabe"
        AskForYear = 0
        Exit Function
    End If
    
    Dim year2 As Integer
    year2 = CInt(yearStr)
    
    If Not valid_YearConfig.IsValidYear(year2) Then
        MsgBox "Ungueltiges Jahr: " & year2 & ". Bitte 24, 25 oder 26 eingeben.", vbExclamation, "Ungueltige Eingabe"
        AskForYear = 0
        Exit Function
    End If
    
    AskForYear = year2
End Function

' Helper: Ask user to select an action
Private Function AskForAction() As String
    Dim actionStr As String
    actionStr = InputBox("Aktion auswaehlen:" & vbCrLf & vbCrLf & _
                         "1 = Pending laden" & vbCrLf & _
                         "2 = Aktuelle Basis laden (tblKartei)" & vbCrLf & _
                         "3 = Entscheidungen sync" & vbCrLf & _
                         "4 = Einzeldatensatz-Geschichte" & vbCrLf & _
                         "5 = Datums-Geschichte (alle)" & vbCrLf & _
                         "6 = Zahlungsexport" & vbCrLf & _
                         "7 = Neu-Liste (nur 2026)", _
                         "Aktion auswaehlen", "1")
    
    Select Case actionStr
        Case "1": AskForAction = "LOAD"
        Case "2": AskForAction = "BASEIMPORT"
        Case "3": AskForAction = "SYNC"
        Case "4": AskForAction = "HISTORY1"
        Case "5": AskForAction = "HISTORYALL"
        Case "6": AskForAction = "PAYMENT"
        Case "7": AskForAction = "NEU"
        Case Else: AskForAction = ""
    End Select
End Function

' ============================================================
' AUTOFILTER HELPER
' ============================================================

' Ensure AutoFilter dropdown arrows are visible on KarteiYY header row (row 2).
' Applies AutoFilter to columns A:AZ (52 columns) from row 2 to the last data row.
' If no data rows exist, filter is applied to header row only (A2:AZ2).
' Does not apply any filter criteria; only enables dropdown arrows.
'
' @param wsKartei - The KarteiYY worksheet to apply AutoFilter to
Private Sub EnsureKarteiAutoFilter(ByVal wsKartei As Worksheet)
    On Error Resume Next
    
    If wsKartei Is Nothing Then Exit Sub
    
    ' Determine last row using ID column AV (48) as anchor
    Dim lastRow As Long
    lastRow = wsKartei.Cells(wsKartei.Rows.Count, 48).End(xlUp).Row
    
    ' Ensure at least row 2 (header row)
    If lastRow < 2 Then lastRow = 2
    
    ' Define range A2:AZ<lastRow> (52 columns)
    Dim rng As Range
    Set rng = wsKartei.Range(wsKartei.Cells(2, 1), wsKartei.Cells(lastRow, 52))
    
    ' Reset any previous filter mode safely
    If wsKartei.AutoFilterMode Then wsKartei.AutoFilterMode = False
    
    ' Apply AutoFilter (no criteria = just show arrows)
    rng.AutoFilter
    
    ' Clear any error that might have occurred
    Err.Clear
End Sub
