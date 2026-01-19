Attribute VB_Name = "valid_ImportPending"
'==========================
'   Module: valid_ImportPending
'   Purpose: Load pending changes from pre_tblKartei into Superadmin Kartei sheet
'   Uses ID column (AV, column 48) as the primary key
'
'   MULTI-YEAR SUPPORT (2024, 2025, 2026):
'   Each year has its own Kartei sheet (Kartei24, Kartei25, Kartei26)
'   and connects to the corresponding year database.
'
'   Entry points:
'     LoadPendingChangesForYear(year2) - Main parameterized entry point
'     LoadPendingChanges24/25/26       - Wrapper macros for UX
'     LoadPendingChangesFromPre        - Legacy (defaults to year 25)
'==========================

Option Explicit

' ============================================================
' MULTI-YEAR API - Entry Points
' ============================================================

' Load pending changes for a specific year into KarteiYY sheet
' @param year2 - Two-digit year (24, 25, or 26)
Public Sub LoadPendingChangesForYear(ByVal year2 As Integer)
    On Error GoTo ErrorHandler
    
    ' Validate year
    If Not IsValidYear(year2) Then
        MsgBox "Ungueltiges Jahr: " & year2 & ". Unterstuetzte Jahre: 24, 25, 26.", _
               vbExclamation, "Jahresfehler"
        Exit Sub
    End If
    
    Application.ScreenUpdating = False
    Application.Calculation = xlCalculationManual
    
    Dim wsKartei As Worksheet
    Set wsKartei = GetOrCreateKarteiSheetForYear(year2)
    
    If wsKartei Is Nothing Then
        MsgBox "Kartei-Blatt fuer Jahr " & year2 & " konnte nicht erstellt/gefunden werden.", _
               vbCritical, "Blattfehler"
        GoTo Cleanup
    End If
    
    ' Clear existing data (keep headers in rows 1-2)
    ClearKarteiData wsKartei
    
    ' Load data from pre_tblKartei for this year
    Dim dictPending As Object
    Set dictPending = ReadPreTableIntoDictionaryForYear(year2)
    
    If dictPending.Count = 0 Then
        MsgBox "Keine ausstehenden Aenderungen in pre_tblKartei (Jahr " & year2 & ") gefunden.", _
               vbInformation, "Ausstehende laden"
        GoTo Cleanup
    End If
    
    ' Write pending data to Kartei sheet
    WritePendingToKartei wsKartei, dictPending
    
    ' Format monthly columns (U-AF) to ensure numeric values with proper decimal separators
    Call valid_FormatMonths.FormatMonthlyColumnsForSheet(wsKartei)
    
    ' Optionally load original values from tblKartei for comparison
    LoadOriginalValuesForYear wsKartei, dictPending, year2
    
    ' Set up date range on grossGeschichteYY sheet (B1=start date, C1=today)
    Call SetupGrossGeschichteDatesForYear(year2)
    
    MsgBox "Erfolgreich " & dictPending.Count & " ausstehende Aenderung(en) aus pre_tblKartei (Jahr " & year2 & ") geladen.", _
           vbInformation, "Ausstehende laden"
    
Cleanup:
    Application.Calculation = xlCalculationAutomatic
    Application.ScreenUpdating = True
    Exit Sub
    
ErrorHandler:
    Application.Calculation = xlCalculationAutomatic
    Application.ScreenUpdating = True
    MsgBox "Fehler beim Laden der ausstehenden Aenderungen (Jahr " & year2 & "): " & Err.Description, _
           vbCritical, "Ladefehler"
End Sub

' Wrapper macro: Load pending changes for year 2024
Public Sub LoadPendingChanges24()
    LoadPendingChangesForYear 24
End Sub

' Wrapper macro: Load pending changes for year 2025
Public Sub LoadPendingChanges25()
    LoadPendingChangesForYear 25
End Sub

' Wrapper macro: Load pending changes for year 2026
Public Sub LoadPendingChanges26()
    LoadPendingChangesForYear 26
End Sub

' ============================================================
' LEGACY API - Backward Compatibility (defaults to year 25)
' ============================================================

' Load all pending changes from pre_tblKartei into Kartei sheet (legacy - year 25)
' For new code, use LoadPendingChangesForYear(year2) instead.
Public Sub LoadPendingChangesFromPre()
    On Error GoTo ErrorHandler
    
    Application.ScreenUpdating = False
    Application.Calculation = xlCalculationManual
    
    ' Prefer legacy single-year sheets if they exist (old workflow).
    ' Fall back to multi-year sheets to keep the code usable after migration.
    Dim wsKartei As Worksheet
    On Error Resume Next
    Set wsKartei = ThisWorkbook.Worksheets("Kartei")
    On Error GoTo 0
    
    If wsKartei Is Nothing Then
        Set wsKartei = GetOrCreateKarteiSheetForYear(25)
    End If
    
    ClearKarteiData wsKartei
    
    Dim dictPending As Object
    Set dictPending = ReadPreTableIntoDictionaryForYear(25)
    
    If dictPending.Count = 0 Then
        MsgBox "Keine ausstehenden Aenderungen in pre_tblKartei (Jahr 25) gefunden.", _
               vbInformation, "Ausstehende laden"
        GoTo Cleanup
    End If
    
    WritePendingToKartei wsKartei, dictPending
    Call valid_FormatMonths.FormatMonthlyColumnsForSheet(wsKartei)
    
    LoadOriginalValues wsKartei, dictPending
    
    ' Legacy date range setup on grossGeschichte (without year suffix).
    Call SetupGrossGeschichteDates
    
    MsgBox "Erfolgreich " & dictPending.Count & " ausstehende Aenderung(en) aus pre_tblKartei (Jahr 25) geladen.", _
           vbInformation, "Ausstehende laden"
    
Cleanup:
    Application.Calculation = xlCalculationAutomatic
    Application.ScreenUpdating = True
    Exit Sub
    
ErrorHandler:
    Application.Calculation = xlCalculationAutomatic
    Application.ScreenUpdating = True
    MsgBox "Fehler beim Laden der ausstehenden Aenderungen (Jahr 25): " & Err.Description, _
           vbCritical, "Ladefehler"
End Sub

' ============================================================
' MULTI-YEAR DATA ACCESS
' ============================================================

' Read all records from pre_tblKartei into dictionary for a specific year (key = ID)
Private Function ReadPreTableIntoDictionaryForYear(ByVal year2 As Integer) As Object
    Dim dict As Object
    Set dict = CreateObject("Scripting.Dictionary")
    
    Dim dbPath As String
    dbPath = GetDatabasePathForYear(year2)
    
    ' Check if user cancelled database selection
    If dbPath = "" Then
        Set ReadPreTableIntoDictionaryForYear = dict
        Exit Function
    End If
    
    ' Check if pre_tblKartei exists, create if needed
    If Not TableExists(dbPath, "pre_tblKartei") Then
        CreatePreTable dbPath
        Set ReadPreTableIntoDictionaryForYear = dict
        Exit Function
    End If
    
    Dim engine As DAO.DBEngine
    Set engine = New DAO.DBEngine
    
    Dim wsDao As DAO.Workspace
    Set wsDao = engine.Workspaces(0)
    
    Dim db As DAO.Database
    Set db = wsDao.OpenDatabase(dbPath)
    
    Dim rs As DAO.Recordset
    Dim sqlStr As String
    sqlStr = "SELECT * FROM pre_tblKartei ORDER BY ID"
    
    Set rs = db.OpenRecordset(sqlStr, dbOpenDynaset)
    
    If Not rs.EOF Then
        Do While Not rs.EOF
            Dim arrRow As Variant
            ReDim arrRow(1 To 1, 1 To 52)
            
            Dim c As Long
            ' Fill array from database fields
            For c = 1 To 51
                If c = 48 Then
                    ' Column 48 = ID field
                    arrRow(1, 48) = NzToEmpty(rs.Fields("ID").Value)
                ElseIf c = 7 Or c = 8 Then
                    ' Phone columns (7=Tel., 8=Handy): must be stored as strings to preserve
                    ' leading zeros and prevent scientific notation (e.g. "0176..." not "1.76E+9")
                    ' Also normalize any existing scientific-notation strings (e.g. "1,76E+9" -> "176000000")
                    arrRow(1, c) = phone_Normalize.NormalizePhoneText(rs.Fields("Value" & c).Value)
                Else
                    ' Value1..Value51
                    arrRow(1, c) = NzToEmpty(rs.Fields("Value" & c).Value)
                End If
            Next c
            
            ' Value52 = AZ column (history)
            arrRow(1, 52) = NzToEmpty(rs.Fields("Value52").Value)
            
            Dim strID As String
            strID = CStr(arrRow(1, 48))
            
            If strID <> "" And Not dict.Exists(strID) Then
                dict.Add strID, arrRow
            End If
            
            rs.MoveNext
        Loop
    End If
    
    rs.Close
    db.Close
    
    Set ReadPreTableIntoDictionaryForYear = dict
End Function

' Legacy wrapper - Read from pre_tblKartei (defaults to year 25)
Private Function ReadPreTableIntoDictionary() As Object
    Set ReadPreTableIntoDictionary = ReadPreTableIntoDictionaryForYear(25)
End Function

' ============================================================
' (Legacy function kept for reference - can be removed after migration)
' ============================================================
' Original ReadPreTableIntoDictionary implementation is now in ReadPreTableIntoDictionaryForYear

' Write pending records to Kartei sheet, starting from row 3
Private Sub WritePendingToKartei(ByVal ws As Worksheet, ByVal dictPending As Object)
    Dim currentRow As Long
    currentRow = 3 ' Start from row 3 (rows 1-2 are headers)
    
    Dim idKey As Variant
    For Each idKey In dictPending.Keys
        Dim arrRow As Variant
        arrRow = dictPending(idKey)
        
        ' Write columns A-AY (1-51)
        Dim c As Long
        For c = 1 To 51
            ' Columns G (7) = Phone, H (8) = Mobile: must be written as TEXT to preserve
            ' leading zeros (e.g. "0176..." should not become 176...).
            ' Without text format, Excel interprets numeric-looking strings as numbers,
            ' which causes false differences in GrossGeschichte War/Ist comparison.
            If c = 7 Or c = 8 Then
                With ws.Cells(currentRow, c)
                    .NumberFormat = "@"
                    .Value = CStr(arrRow(1, c))
                End With
            Else
                ws.Cells(currentRow, c).Value = arrRow(1, c)
            End If
        Next c
        
        ' Write AZ (column 52 = history)
        ws.Cells(currentRow, 52).Value = arrRow(1, 52)
        
        ' Ensure ID is in column AV (48)
        ws.Cells(currentRow, 48).Value = CLng(idKey)
        
        ' Mark as pending with light blue fill in column A (if not "Zahlung")
        If UCase(Trim(CStr(arrRow(1, 4)))) <> "ZAHLUNG" Then
            ws.Cells(currentRow, 1).Interior.Color = RGB(173, 216, 230) ' Light blue
        End If
        
        currentRow = currentRow + 1
    Next idKey
End Sub

' Optionally load original values from tblKartei for comparison
' (can be displayed in adjacent columns or separate sheet)
Private Sub LoadOriginalValues(ByVal ws As Worksheet, ByVal dictPending As Object)
    ' This is a stub - implement if Superadmin needs to see original vs pending side-by-side
    ' For now, we'll just ensure the data is loaded
    ' In future iterations, this could populate columns to the right or a separate comparison sheet
End Sub

' Year-parameterized version of LoadOriginalValues
Private Sub LoadOriginalValuesForYear(ByVal ws As Worksheet, ByVal dictPending As Object, ByVal year2 As Integer)
    ' This is a stub - implement if Superadmin needs to see original vs pending side-by-side
    ' For now, we'll just ensure the data is loaded
    ' In future iterations, this could query tblKartei for the specified year
    ' and populate columns to the right or a separate comparison sheet
End Sub

' ============================================================
' MULTI-YEAR SHEET MANAGEMENT
' ============================================================

' Get or create Kartei worksheet for a specific year (Kartei24, Kartei25, Kartei26)
Private Function GetOrCreateKarteiSheetForYear(ByVal year2 As Integer) As Worksheet
    On Error GoTo ErrorHandler
    
    Call valid_YearConfig.EnsureYearSheetsExist(year2)
    
    Dim ws As Worksheet
    On Error Resume Next
    Set ws = ThisWorkbook.Worksheets(valid_YearConfig.GetKarteiSheetName(year2))
    On Error GoTo 0
    
    Set GetOrCreateKarteiSheetForYear = ws
    Exit Function
    
ErrorHandler:
    Set GetOrCreateKarteiSheetForYear = Nothing
End Function

' Get or create Kartei worksheet in Superadmin file (legacy - year 25)
Private Function GetOrCreateKarteiSheet() As Worksheet
    ' For backward compatibility, return the year 25 sheet
    Set GetOrCreateKarteiSheet = GetOrCreateKarteiSheetForYear(25)
End Function

' Clear data rows in Kartei (keep headers)
Private Sub ClearKarteiData(ByVal ws As Worksheet)
    Dim lastRow As Long
    ' ID column (AV=48) is the most reliable anchor for "last used row" here.
    lastRow = ws.Cells(ws.Rows.Count, 48).End(xlUp).Row
    
    If lastRow > 2 Then
        ws.Rows("3:" & lastRow).ClearContents
        ws.Rows("3:" & lastRow).Interior.Color = xlNone
    End If
End Sub

' ============================================================
' YEAR VALIDATION HELPER
' ============================================================

' Check if year is valid (24, 25, or 26)
Private Function IsValidYear(ByVal year2 As Integer) As Boolean
    IsValidYear = (year2 = 24 Or year2 = 25 Or year2 = 26)
End Function

' ============================================================
' DATABASE ACCESS HELPERS
' ============================================================

' Get database path with validation (prompts user if file not found)
' Defaults to year 25 for backward compatibility
Private Function GetDatabasePath() As String
    GetDatabasePath = valid_DatabasePath.GetValidatedDatabasePath()
End Function

' Get database path for a specific year
Private Function GetDatabasePathForYear(ByVal year2 As Integer) As String
    GetDatabasePathForYear = valid_DatabasePath.GetValidatedDatabasePathForYear(year2)
End Function

' Check if table exists in database
Private Function TableExists(ByVal dbPath As String, ByVal tableName As String) As Boolean
    On Error Resume Next
    
    Dim engine As DAO.DBEngine
    Set engine = New DAO.DBEngine
    
    Dim wsDao As DAO.Workspace
    Set wsDao = engine.Workspaces(0)
    
    Dim db As DAO.Database
    Set db = wsDao.OpenDatabase(dbPath)
    
    Dim tbl As DAO.TableDef
    For Each tbl In db.TableDefs
        If tbl.Name = tableName Then
            TableExists = True
            db.Close
            Exit Function
        End If
    Next tbl
    
    db.Close
    TableExists = False
End Function

' Create pre_tblKartei with same structure as tblKartei
Private Sub CreatePreTable(ByVal dbPath As String)
    Dim engine As DAO.DBEngine
    Set engine = New DAO.DBEngine
    
    Dim wsDao As DAO.Workspace
    Set wsDao = engine.Workspaces(0)
    
    Dim db As DAO.Database
    Set db = wsDao.OpenDatabase(dbPath)
    
    ' Create table with same structure as tblKartei
    Dim tbl As DAO.TableDef
    Set tbl = db.CreateTableDef("pre_tblKartei")
    
    ' Add ID field (regular Long, not AutoNumber - ID will be set explicitly from Kartei!AV)
    Dim fld As DAO.Field
    Set fld = tbl.CreateField("ID", dbLong)
    ' NO fld.Attributes = dbAutoIncrField - ID is manually assigned
    tbl.Fields.Append fld
    
    ' Add Value1..Value51 fields (Text)
    Dim i As Long
    For i = 1 To 51
        Set fld = tbl.CreateField("Value" & i, dbText, 255)
        fld.AllowZeroLength = True
        tbl.Fields.Append fld
    Next i
    
    ' Add Value52 (Text, for history)
    Set fld = tbl.CreateField("Value52", dbMemo)
    tbl.Fields.Append fld
    
    ' Add format fields
    For i = 1 To 51
        Set fld = tbl.CreateField("InteriorColor" & i, dbLong)
        tbl.Fields.Append fld
    Next i
    
    Set fld = tbl.CreateField("FontColor3", dbLong)
    tbl.Fields.Append fld
    
    Set fld = tbl.CreateField("FontColor18", dbLong)
    tbl.Fields.Append fld
    
    db.TableDefs.Append tbl
    db.Close
End Sub

' Helper function to convert Null to empty string
Private Function NzToEmpty(ByVal v As Variant) As Variant
    If IsNull(v) Then
        NzToEmpty = ""
    Else
        NzToEmpty = v
    End If
End Function

' ============================================================
' MULTI-YEAR grossGeschichte DATE SETUP
' ============================================================

' Set up date range on grossGeschichteYY sheet for filtering
' B1 = Start date (default: 30 days ago), C1 = End date (today)
' Both cells formatted as DD.MM.YYYY
'
' CONDITIONAL BEHAVIOR (Prompt 13):
'   - Only sets B1 if it is empty or not a valid date (does not overwrite user/Dashboard values)
'   - Only sets C1 if it is empty or not a valid date (does not overwrite user/Dashboard values)
'   - This allows Dashboard to pre-fill dates before LoadPending is called
Private Sub SetupGrossGeschichteDatesForYear(ByVal year2 As Integer)
    On Error Resume Next
    
    Dim sheetName As String
    sheetName = valid_YearConfig.GetGrossGeschichteSheetName(year2)
    
    Dim wsGross As Worksheet
    Set wsGross = ThisWorkbook.Worksheets(sheetName)
    
    If wsGross Is Nothing Then
        ' Create grossGeschichteYY sheet if it doesn't exist
        Set wsGross = ThisWorkbook.Worksheets.Add(After:=ThisWorkbook.Worksheets(ThisWorkbook.Worksheets.Count))
        wsGross.Name = sheetName
        wsGross.Range("A1").Value = "Start Date:"
        wsGross.Range("A1").Font.Bold = True
    End If
    
    ' Set B1 to 30 days ago ONLY if empty or not a valid date
    ' (Do not overwrite user-configured or Dashboard-provided dates)
    If IsEmpty(wsGross.Range("B1").Value) Or Not IsDate(wsGross.Range("B1").Value) Then
        wsGross.Range("B1").Value = Date - 30
    End If
    
    ' Set C1 to current date ONLY if empty or not a valid date
    ' (Do not overwrite user-configured or Dashboard-provided dates)
    If IsEmpty(wsGross.Range("C1").Value) Or Not IsDate(wsGross.Range("C1").Value) Then
        wsGross.Range("C1").Value = Date
    End If
    
    ' Apply date format dd.mm.yyyy AFTER setting values to ensure it sticks
    wsGross.Range("B1:C1").NumberFormat = "dd.mm.yyyy"
    
    On Error GoTo 0
End Sub

' Set up date range on GrossGeschichte sheet for filtering (legacy - year 25)
' B1 = Start date (default: 30 days ago), C1 = End date (today)
' Both cells formatted as DD.MM.YYYY
'
' CONDITIONAL BEHAVIOR (Prompt 13):
'   - Only sets B1 if it is empty or not a valid date (does not overwrite user/Dashboard values)
'   - Only sets C1 if it is empty or not a valid date (does not overwrite user/Dashboard values)
Private Sub SetupGrossGeschichteDates()
    On Error Resume Next
    
    Dim wsGross As Worksheet
    Set wsGross = ThisWorkbook.Worksheets("grossGeschichte")
    
    If wsGross Is Nothing Then
        ' Create GrossGeschichte sheet if it doesn't exist
        Set wsGross = ThisWorkbook.Worksheets.Add(After:=ThisWorkbook.Worksheets(ThisWorkbook.Worksheets.Count))
        wsGross.Name = "grossGeschichte"
        wsGross.Range("A1").Value = "Start Date:"
        wsGross.Range("A1").Font.Bold = True
    End If
    
    ' Set B1 to 30 days ago ONLY if empty or not a valid date
    ' (Do not overwrite user-configured or Dashboard-provided dates)
    If IsEmpty(wsGross.Range("B1").Value) Or Not IsDate(wsGross.Range("B1").Value) Then
        wsGross.Range("B1").Value = Date - 30
    End If
    
    ' Set C1 to current date ONLY if empty or not a valid date
    ' (Do not overwrite user-configured or Dashboard-provided dates)
    If IsEmpty(wsGross.Range("C1").Value) Or Not IsDate(wsGross.Range("C1").Value) Then
        wsGross.Range("C1").Value = Date
    End If
    
    ' Apply date format dd.mm.yyyy AFTER setting values to ensure it sticks
    wsGross.Range("B1:C1").NumberFormat = "dd.mm.yyyy"
    
    On Error GoTo 0
End Sub
