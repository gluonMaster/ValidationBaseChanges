Attribute VB_Name = "valid_NeuList"
'==========================
'   Module: valid_NeuList
'   Purpose: Manage new records list (Neu) for Superadmin
'   - Tracks new records in tblKartei by ID
'   - LastSeenID stored in hidden NeuConfig sheet
'   - Neu sheet displays all records with ID > LastSeenID
'==========================

Option Explicit

'===========================================
'  Public Procedures
'===========================================

' Refresh the Neu list with all records having ID > LastSeenID
Public Sub RefreshNeuList()
    On Error GoTo ErrorHandler
    
    Application.ScreenUpdating = False
    Application.Calculation = xlCalculationManual
    
    Dim dbPath As String
    dbPath = GetDatabasePath()
    
    ' Check if user cancelled database selection
    If dbPath = "" Then
        GoTo Cleanup
    End If
    
    Dim lastSeenID As Long
    lastSeenID = GetLastSeenID()
    
    ' Open database via DAO
    Dim engine As DAO.DBEngine
    Dim wsDao As DAO.Workspace
    Dim db As DAO.Database
    
    Set engine = New DAO.DBEngine
    Set wsDao = engine.Workspaces(0)
    Set db = wsDao.OpenDatabase(dbPath)
    
    ' Open Recordset with filter by ID
    Dim rs As DAO.Recordset
    Dim sqlStr As String
    sqlStr = "SELECT * FROM tblKartei WHERE ID > " & CLng(lastSeenID) & " ORDER BY ID"
    Set rs = db.OpenRecordset(sqlStr, dbOpenSnapshot)
    
    ' Get or create Neu sheet
    Dim wsNeu As Worksheet
    Set wsNeu = GetOrCreateNeuSheet()
    
    ' Copy header from Kartei to Neu (row 2)
    Dim wsK As Worksheet
    Set wsK = ThisWorkbook.Worksheets("Kartei")
    wsK.Rows(2).Copy Destination:=wsNeu.Rows(2)
    
    ' Ensure AutoFilter is enabled on row 2
    If Not wsNeu.AutoFilterMode Then
        wsNeu.Range("A2:BA2").AutoFilter
    End If
    
    ' Check if there are new records
    If rs.EOF Then
        rs.Close
        db.Close
        ClearNeuData wsNeu
        MsgBox "Keine neuen Eintraege gefunden (ID > " & lastSeenID & ").", vbInformation, "Neu-Liste aktualisieren"
        GoTo Cleanup
    End If
    
    ' Get all data into a 2D array: arrRaw(Fields, Records)
    rs.MoveLast
    rs.MoveFirst
    Dim arrRaw As Variant
    arrRaw = rs.GetRows(rs.RecordCount)
    
    rs.Close
    db.Close
    
    ' Calculate row count
    Dim rowCount As Long
    rowCount = UBound(arrRaw, 2) + 1  ' +1 because array is 0-based
    
    If rowCount = 0 Then
        ClearNeuData wsNeu
        MsgBox "Keine neuen Eintraege gefunden (ID > " & lastSeenID & ").", vbInformation, "Neu-Liste aktualisieren"
        GoTo Cleanup
    End If
    
    ' Build arrays for values and formats
    ' Structure matches ImportKarteiAndFormat_Optimized:
    '   0=ID
    '   1..102 => ValueX & InteriorColorX (51 * 2)
    '   103 => FontColor3
    '   104 => FontColor18
    '   105 => Value52 (History)
    
    Dim arrValues As Variant
    ReDim arrValues(1 To rowCount, 1 To 52)
    
    Dim arrInterior As Variant
    ReDim arrInterior(1 To rowCount, 1 To 51)
    
    Dim arrFontC As Variant    ' for col=3 (Child)
    Dim arrFontR As Variant    ' for col=18 (Price2)
    ReDim arrFontC(1 To rowCount)
    ReDim arrFontR(1 To rowCount)
    
    Dim rec As Long
    Dim rowDest As Long
    Dim col As Long
    Dim idxValue As Long
    
    For rec = 0 To rowCount - 1
        rowDest = rec + 1
        
        ' Fill columns 1..51
        For col = 1 To 51
            ' Value index: col=1 => idx=1, col=2 => idx=3, col=51 => idx=101
            idxValue = 1 + (col - 1) * 2
            arrValues(rowDest, col) = arrRaw(idxValue, rec)
            arrInterior(rowDest, col) = arrRaw(idxValue + 1, rec)
        Next col
        
        ' Value52 = AZ column (history) at index 105
        arrValues(rowDest, 52) = arrRaw(105, rec)
        
        ' Font colors
        arrFontC(rowDest) = arrRaw(103, rec)   ' FontColor3
        arrFontR(rowDest) = arrRaw(104, rec)   ' FontColor18
    Next rec
    
    ' Clear old data before writing new
    ClearNeuData wsNeu
    
    ' Write values to Neu sheet starting from row 3
    wsNeu.Range("A3").Resize(rowCount, 52).Value = arrValues
    
    ' Apply formats (Interior.Color and Font.Color)
    Dim i As Long, c As Long
    For i = 1 To rowCount
        For c = 1 To 51
            If Not IsEmpty(arrInterior(i, c)) And arrInterior(i, c) <> 0 Then
                wsNeu.Cells(i + 2, c).Interior.Color = arrInterior(i, c)
            End If
        Next c
        
        ' Set font color in col=3 (Child)
        If Not IsEmpty(arrFontC(i)) And arrFontC(i) <> 0 Then
            wsNeu.Cells(i + 2, 3).Font.Color = arrFontC(i)
        End If
        
        ' Set font color in col=18 (Price2)
        If Not IsEmpty(arrFontR(i)) And arrFontR(i) <> 0 Then
            wsNeu.Cells(i + 2, 18).Font.Color = arrFontR(i)
        End If
    Next i
    
    ' Sort by name
    SortNeuByName wsNeu
    
    MsgBox rowCount & " neue(n) Eintrag/Eintraege mit ID > " & lastSeenID & " geladen.", vbInformation, "Neu-Liste aktualisieren"
    
Cleanup:
    Application.Calculation = xlCalculationAutomatic
    Application.ScreenUpdating = True
    Exit Sub
    
ErrorHandler:
    Application.Calculation = xlCalculationAutomatic
    Application.ScreenUpdating = True
    MsgBox "Fehler beim Aktualisieren der Neu-Liste: " & Err.Description, vbCritical, "Aktualisierungsfehler"
End Sub

' Mark all records as "seen" by setting LastSeenID to current max ID in database
Public Sub GetOffNeu()
    On Error GoTo ErrorHandler
    
    Dim wsNeu As Worksheet
    Set wsNeu = GetOrCreateNeuSheet()
    
    ' Get database path and open database via DAO
    Dim dbPath As String
    dbPath = GetDatabasePath()
    
    ' Check if user cancelled database selection
    If dbPath = "" Then
        Exit Sub
    End If
    
    Dim engine As DAO.DBEngine
    Dim wsDao As DAO.Workspace
    Dim db As DAO.Database
    
    Set engine = New DAO.DBEngine
    Set wsDao = engine.Workspaces(0)
    Set db = wsDao.OpenDatabase(dbPath)
    
    ' Get max ID from tblKartei
    Dim rsMax As DAO.Recordset
    Set rsMax = db.OpenRecordset("SELECT MAX(ID) AS MaxID FROM tblKartei", dbOpenSnapshot)
    
    Dim maxID As Long
    maxID = 0
    
    If Not rsMax.EOF Then
        If Not IsNull(rsMax.Fields("MaxID").Value) Then
            maxID = CLng(rsMax.Fields("MaxID").Value)
        End If
    End If
    
    rsMax.Close
    db.Close
    
    If maxID = 0 Then
        MsgBox "Keine Eintraege in tblKartei gefunden.", vbExclamation, "GetOffNeu"
        Exit Sub
    End If
    
    ' Get old value for info message
    Dim oldLastSeenID As Long
    oldLastSeenID = GetLastSeenID()
    
    ' Update LastSeenID to max ID from database
    SetLastSeenID maxID
    
    ' Clear Neu sheet data (keep header and filters)
    ClearNeuData wsNeu
    
    ' Save workbook to persist LastSeenID
    ThisWorkbook.Save
    
    MsgBox "Neu geloescht, LastSeenID aktualisiert von " & oldLastSeenID & " auf " & maxID & ".", _
           vbInformation, "GetOffNeu"
    
    Exit Sub
    
ErrorHandler:
    MsgBox "Fehler in GetOffNeu: " & Err.Description, vbCritical, "GetOffNeu Fehler"
End Sub

'===========================================
'  Private Helper Functions
'===========================================

' Get database path with validation (prompts user if file not found)
Private Function GetDatabasePath() As String
    GetDatabasePath = valid_DatabasePath.GetValidatedDatabasePath()
End Function

' Get LastSeenID from NeuConfig sheet
Private Function GetLastSeenID() As Long
    On Error Resume Next
    
    Dim wsConfig As Worksheet
    Set wsConfig = GetOrCreateConfigSheet()
    
    Dim val As Variant
    val = wsConfig.Range("A1").Value
    
    If IsEmpty(val) Or Not IsNumeric(val) Then
        GetLastSeenID = 0
    Else
        GetLastSeenID = CLng(val)
    End If
    
    On Error GoTo 0
End Function

' Set LastSeenID in NeuConfig sheet
Private Sub SetLastSeenID(ByVal newValue As Long)
    On Error Resume Next
    
    Dim wsConfig As Worksheet
    Set wsConfig = GetOrCreateConfigSheet()
    
    wsConfig.Range("A1").Value = newValue
    wsConfig.Range("B1").Value = "Max ID seen in Neu list"
    
    On Error GoTo 0
End Sub

' Get or create the hidden NeuConfig sheet
Private Function GetOrCreateConfigSheet() As Worksheet
    Dim ws As Worksheet
    
    On Error Resume Next
    Set ws = ThisWorkbook.Worksheets("NeuConfig")
    On Error GoTo 0
    
    If ws Is Nothing Then
        ' Create new sheet at the end
        Set ws = ThisWorkbook.Worksheets.Add(After:=ThisWorkbook.Worksheets(ThisWorkbook.Worksheets.Count))
        ws.Name = "NeuConfig"
        ws.Visible = xlSheetVeryHidden
    End If
    
    Set GetOrCreateConfigSheet = ws
End Function

' Get or create the Neu sheet
Private Function GetOrCreateNeuSheet() As Worksheet
    Dim ws As Worksheet
    
    On Error Resume Next
    Set ws = ThisWorkbook.Worksheets("Neu")
    On Error GoTo 0
    
    If ws Is Nothing Then
        ' Try to create after Kartei sheet
        Dim wsKartei As Worksheet
        On Error Resume Next
        Set wsKartei = ThisWorkbook.Worksheets("Kartei")
        On Error GoTo 0
        
        If Not wsKartei Is Nothing Then
            Set ws = ThisWorkbook.Worksheets.Add(After:=wsKartei)
        Else
            Set ws = ThisWorkbook.Worksheets.Add(After:=ThisWorkbook.Worksheets(ThisWorkbook.Worksheets.Count))
        End If
        ws.Name = "Neu"
    End If
    
    Set GetOrCreateNeuSheet = ws
End Function

' Clear data rows on Neu sheet (keep header in row 2)
Private Sub ClearNeuData(ByVal wsNeu As Worksheet)
    Dim lastRow As Long
    lastRow = wsNeu.Cells(wsNeu.Rows.Count, 1).End(xlUp).Row
    
    If lastRow >= 3 Then
        wsNeu.Range("A3:BA" & lastRow).ClearContents
        wsNeu.Range("A3:BA" & lastRow).Interior.ColorIndex = xlColorIndexNone
        wsNeu.Range("A3:BA" & lastRow).Font.ColorIndex = xlColorIndexAutomatic
    End If
End Sub

' Sort Neu sheet by name (similar to SortNameZ)
Private Sub SortNeuByName(ByVal wsNeu As Worksheet)
    Dim letzte As Long
    Dim i As Long
    
    ' Find last row with data in column A
    For i = 3 To 3000
        If wsNeu.Range("A" & i).Value = "" Then
            letzte = i - 1
            Exit For
        End If
    Next i
    
    ' If no data found or only header
    If letzte < 3 Then Exit Sub
    
    ' Set row height
    wsNeu.Range("A3:BA" & letzte).RowHeight = 16
    
    ' Clear existing sort fields and set up new sort
    wsNeu.Sort.SortFields.Clear
    
    ' Key 1: Column B (Parent) - ascending
    wsNeu.Sort.SortFields.Add Key:=wsNeu.Range("B3:B" & letzte), _
        SortOn:=xlSortOnValues, Order:=xlAscending, DataOption:=xlSortNormal
    
    ' Key 2: Column A (FamilyID) - ascending
    wsNeu.Sort.SortFields.Add Key:=wsNeu.Range("A3:A" & letzte), _
        SortOn:=xlSortOnValues, Order:=xlAscending, DataOption:=xlSortNormal
    
    ' Key 3: Column D (Child/Kind) - ascending
    wsNeu.Sort.SortFields.Add Key:=wsNeu.Range("D3:D" & letzte), _
        SortOn:=xlSortOnValues, Order:=xlAscending, DataOption:=xlSortNormal
    
    ' Apply sort
    With wsNeu.Sort
        .SetRange wsNeu.Range("A2:BA" & letzte)
        .Header = xlYes
        .MatchCase = False
        .Orientation = xlTopToBottom
        .SortMethod = xlPinYin
        .Apply
    End With
End Sub
