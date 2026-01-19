Attribute VB_Name = "valid_ImportFromBase"
'==========================
'   Module: valid_ImportFromBase
'   Purpose: Import full tblKartei (current database state) into KarteiYY
'
'   Unlike LoadPendingChanges which imports from pre_tblKartei (pending changes),
'   this module imports the FULL current database state from tblKartei.
'
'   Use cases:
'   - Browse current database state for any year
'   - Run Dashboard_SingleRecordHistory for any record (not just pending)
'
'   IMPORTANT: This import does NOT build grossGeschichteYY (no decision sheet).
'   To return to the pending workflow, use Dashboard_LoadPending.
'
'   MULTI-YEAR SUPPORT (2024, 2025, 2026):
'   Each year has its own Kartei sheet (Kartei24, Kartei25, Kartei26)
'   and connects to the corresponding year database.
'
'   Entry points:
'     ImportFromBaseForYear(year2) - Main parameterized entry point
'     ImportFromBase24/25/26       - Wrapper macros for UX
'==========================

Option Explicit

' ============================================================
' MULTI-YEAR API - Entry Points
' ============================================================

' Import full tblKartei for a specific year into KarteiYY sheet
' @param year2 - Two-digit year (24, 25, or 26)
Public Sub ImportFromBaseForYear(ByVal year2 As Integer)
    On Error GoTo ErrorHandler
    
    ' Validate year
    If Not valid_YearConfig.IsValidYear(year2) Then
        MsgBox "Ungueltiges Jahr: " & year2 & ". Unterstuetzte Jahre: 24, 25, 26.", _
               vbExclamation, "Jahresfehler"
        Exit Sub
    End If
    
    ' Resolve database path (prompts user if not configured)
    Dim dbPath As String
    dbPath = valid_YearConfig.GetDbPathForYear(year2)
    
    If dbPath = "" Then
        MsgBox "Datenbankpfad fuer Jahr " & year2 & " nicht konfiguriert.", _
               vbExclamation, "Pfadfehler"
        Exit Sub
    End If
    
    Application.ScreenUpdating = False
    Application.Calculation = xlCalculationManual
    
    ' Ensure required year sheets exist
    valid_YearConfig.EnsureYearSheetsExist year2
    
    Dim wsKartei As Worksheet
    On Error Resume Next
    Set wsKartei = ThisWorkbook.Worksheets(valid_YearConfig.GetKarteiSheetName(year2))
    On Error GoTo ErrorHandler
    
    If wsKartei Is Nothing Then
        MsgBox "Kartei-Blatt fuer Jahr " & year2 & " konnte nicht erstellt/gefunden werden.", _
               vbCritical, "Blattfehler"
        GoTo Cleanup
    End If
    
    ' Clear existing data (keep headers in rows 1-2)
    ClearKarteiData wsKartei
    
    ' Import from tblKartei
    Dim rowCount As Long
    rowCount = ImportTblKarteiToSheet(dbPath, wsKartei)
    
    If rowCount = 0 Then
        MsgBox "Keine Datensaetze in tblKartei (Jahr " & year2 & ") gefunden.", _
               vbInformation, "Import"
        GoTo Cleanup
    End If
    
    ' Format monthly columns (U-AF) to ensure numeric values with proper decimal separators
    On Error Resume Next
    Call valid_FormatMonths.FormatMonthlyColumnsForSheet(wsKartei)
    On Error GoTo ErrorHandler
    
    ' Activate the Kartei sheet so user can see the result
    wsKartei.Activate
    wsKartei.Range("A3").Select
    
    ' Show success message with clarification about the mode
    MsgBox "Erfolgreich " & rowCount & " Datensaetze aus tblKartei (Jahr " & year2 & ") importiert." & vbCrLf & vbCrLf & _
           "HINWEIS: grossGeschichte" & Format(year2, "00") & " wurde NICHT erstellt." & vbCrLf & _
           "(Dies ist der vollstaendige Datenbankimport, nicht der Pending-Workflow.)" & vbCrLf & vbCrLf & _
           "Um zum Pending-Workflow zurueckzukehren, verwenden Sie 'Dashboard_LoadPending'.", _
           vbInformation, "Import abgeschlossen"
    
Cleanup:
    Application.Calculation = xlCalculationAutomatic
    Application.ScreenUpdating = True
    Exit Sub
    
ErrorHandler:
    Application.Calculation = xlCalculationAutomatic
    Application.ScreenUpdating = True
    MsgBox "Fehler beim Import aus tblKartei (Jahr " & year2 & "): " & Err.Description, _
           vbCritical, "Importfehler"
End Sub

' Wrapper macro: Import full base for year 2024
Public Sub ImportFromBase24()
    ImportFromBaseForYear 24
End Sub

' Wrapper macro: Import full base for year 2025
Public Sub ImportFromBase25()
    ImportFromBaseForYear 25
End Sub

' Wrapper macro: Import full base for year 2026
Public Sub ImportFromBase26()
    ImportFromBaseForYear 26
End Sub

' ============================================================
' CORE IMPORT LOGIC
' ============================================================

' Import all records from tblKartei into the specified worksheet
' Returns the number of rows imported
Private Function ImportTblKarteiToSheet(ByVal dbPath As String, ByVal wsKartei As Worksheet) As Long
    On Error GoTo ErrorHandler
    
    Dim engine As DAO.DBEngine
    Set engine = New DAO.DBEngine
    
    Dim wsDao As DAO.Workspace
    Set wsDao = engine.Workspaces(0)
    
    Dim db As DAO.Database
    Set db = wsDao.OpenDatabase(dbPath)
    
    Dim rs As DAO.Recordset
    Set rs = db.OpenRecordset("SELECT * FROM tblKartei ORDER BY ID", dbOpenSnapshot)
    
    If rs.EOF Then
        rs.Close
        db.Close
        ImportTblKarteiToSheet = 0
        Exit Function
    End If
    
    ' Move to end to get accurate record count
    rs.MoveLast
    rs.MoveFirst
    
    ' Get all data into a 2D array using GetRows (Fields x Records)
    Dim arrRaw As Variant
    arrRaw = rs.GetRows(rs.RecordCount)
    
    rs.Close
    db.Close
    
    ' Calculate dimensions
    Dim rowCount As Long
    rowCount = UBound(arrRaw, 2) + 1  ' +1 because array is 0-based
    
    Dim fieldCount As Long
    fieldCount = UBound(arrRaw, 1) + 1
    
    If rowCount = 0 Then
        ImportTblKarteiToSheet = 0
        Exit Function
    End If
    
    ' Expected field structure (105 fields total):
    '   0 = ID
    '   1..102 => Value1..Value51 & InteriorColor1..InteriorColor51 (alternating: ValueN, InteriorColorN)
    '   103 => FontColor3
    '   104 => FontColor18
    '   (105 => Value52 if present)
    '
    ' Mapping:
    '   arrRaw(0, rec) = ID
    '   arrRaw(2*N - 1, rec) = ValueN (N = 1..51)
    '   arrRaw(2*N, rec) = InteriorColorN (N = 1..51)
    '   arrRaw(103, rec) = FontColor3
    '   arrRaw(104, rec) = FontColor18
    '   arrRaw(105, rec) = Value52 (history string)
    
    ' Build value arrays for bulk write
    Dim arrValues As Variant
    ReDim arrValues(1 To rowCount, 1 To 52)
    
    ' Build interior color array (columns 1..51)
    Dim arrInterior As Variant
    ReDim arrInterior(1 To rowCount, 1 To 51)
    
    ' Font color arrays for columns 3 and 18
    Dim arrFontC As Variant  ' FontColor3
    Dim arrFontR As Variant  ' FontColor18
    ReDim arrFontC(1 To rowCount)
    ReDim arrFontR(1 To rowCount)
    
    Dim rec As Long
    Dim rowDest As Long
    Dim col As Long
    Dim idxValue As Long
    
    For rec = 0 To rowCount - 1
        rowDest = rec + 1
        
        ' Fill columns 1..51 from Value1..Value51
        For col = 1 To 51
            ' ValueN is at index (2*N - 1), InteriorColorN is at index (2*N)
            idxValue = 1 + (col - 1) * 2
            
            ' Phone columns (7=Tel., 8=Handy): must be stored as strings to preserve
            ' leading zeros and prevent scientific notation (e.g. "0176..." not "1.76E+9")
            ' Also normalize any existing scientific-notation strings (e.g. "1,76E+9" -> "176000000")
            If col = 7 Or col = 8 Then
                arrValues(rowDest, col) = phone_Normalize.NormalizePhoneText(arrRaw(idxValue, rec))
            Else
                arrValues(rowDest, col) = NzToEmpty(arrRaw(idxValue, rec))
            End If
            arrInterior(rowDest, col) = NzToLong(arrRaw(idxValue + 1, rec))
        Next col
        
        ' Column 48 (AV) = ID
        arrValues(rowDest, 48) = NzToEmpty(arrRaw(0, rec))
        
        ' Column 52 (AZ) = Value52 (history string)
        If fieldCount > 105 Then
            arrValues(rowDest, 52) = NzToEmpty(arrRaw(105, rec))
        Else
            arrValues(rowDest, 52) = ""
        End If
        
        ' Font colors
        arrFontC(rowDest) = NzToLong(arrRaw(103, rec))
        arrFontR(rowDest) = NzToLong(arrRaw(104, rec))
    Next rec
    
    ' Pre-format phone columns (G=7, H=8) as TEXT before bulk write to prevent
    ' Excel from converting phone numbers to scientific notation.
    ' Must be done BEFORE writing values, otherwise Excel may auto-detect numbers.
    wsKartei.Range(wsKartei.Cells(3, 7), wsKartei.Cells(rowCount + 2, 7)).NumberFormat = "@"
    wsKartei.Range(wsKartei.Cells(3, 8), wsKartei.Cells(rowCount + 2, 8)).NumberFormat = "@"
    
    ' Bulk write values to sheet (A3:AZ<last>)
    wsKartei.Range("A3").Resize(rowCount, 52).Value = arrValues
    
    ' Apply interior colors (columns 1..51)
    Dim i As Long, c As Long
    For i = 1 To rowCount
        For c = 1 To 51
            If arrInterior(i, c) <> 0 Then
                wsKartei.Cells(i + 2, c).Interior.Color = arrInterior(i, c)
            End If
        Next c
        
        ' Apply font colors to columns 3 and 18
        If arrFontC(i) <> 0 Then
            wsKartei.Cells(i + 2, 3).Font.Color = arrFontC(i)
        End If
        If arrFontR(i) <> 0 Then
            wsKartei.Cells(i + 2, 18).Font.Color = arrFontR(i)
        End If
    Next i
    
    ' Sort by name (column B) for usability
    SortByName wsKartei, rowCount
    
    ' Remove AutoFilter leftovers if present
    If wsKartei.AutoFilterMode Then
        wsKartei.AutoFilterMode = False
    End If
    
    ImportTblKarteiToSheet = rowCount
    Exit Function
    
ErrorHandler:
    ImportTblKarteiToSheet = 0
    Debug.Print "ImportTblKarteiToSheet Error: " & Err.Description
End Function

' ============================================================
' HELPER FUNCTIONS
' ============================================================

' Clear data rows in Kartei (keep headers in rows 1-2)
Private Sub ClearKarteiData(ByVal ws As Worksheet)
    On Error Resume Next
    
    Dim lastRow As Long
    ' Check multiple columns to find the actual last row
    Dim lastRowA As Long, lastRowB As Long, lastRowAV As Long
    
    lastRowA = ws.Cells(ws.Rows.Count, 1).End(xlUp).Row
    lastRowB = ws.Cells(ws.Rows.Count, 2).End(xlUp).Row
    lastRowAV = ws.Cells(ws.Rows.Count, 48).End(xlUp).Row  ' ID column
    
    lastRow = Application.WorksheetFunction.Max(lastRowA, lastRowB, lastRowAV)
    
    If lastRow > 2 Then
        With ws.Range(ws.Cells(3, 1), ws.Cells(lastRow, 52))
            .ClearContents
            .Interior.ColorIndex = xlColorIndexNone
            .Font.ColorIndex = xlColorIndexAutomatic
        End With
    End If
    
    On Error GoTo 0
End Sub

' Sort the imported data by name (column B)
Private Sub SortByName(ByVal ws As Worksheet, ByVal rowCount As Long)
    On Error Resume Next
    
    If rowCount < 2 Then Exit Sub
    
    Dim lastRow As Long
    lastRow = rowCount + 2  ' Data starts at row 3
    
    With ws
        .Range(.Cells(2, 1), .Cells(lastRow, 52)).Sort _
            Key1:=.Range("B2"), Order1:=xlAscending, Header:=xlYes
    End With
    
    On Error GoTo 0
End Sub

' Convert Null to empty string
Private Function NzToEmpty(ByVal v As Variant) As Variant
    If IsNull(v) Then
        NzToEmpty = ""
    Else
        NzToEmpty = v
    End If
End Function

' Convert Null to 0 (for color values)
Private Function NzToLong(ByVal v As Variant) As Long
    If IsNull(v) Then
        NzToLong = 0
    ElseIf IsNumeric(v) Then
        NzToLong = CLng(v)
    Else
        NzToLong = 0
    End If
End Function
