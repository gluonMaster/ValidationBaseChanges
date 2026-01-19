Attribute VB_Name = "valid_GrossGeschichteData"
'==========================
'   Module: valid_GrossGeschichteData
'   Purpose: Data access layer for GrossGeschichte comparison view
'   - Retrieves "War" (original) records from tblKartei in Access database
'   - Retrieves "Ist" (pending) records from Kartei worksheet (loaded from pre_tblKartei)
'   - Self-contained for Superadmin file - no dependencies on Admin modules
'
'   Array Layout (1 To 1, 1 To 52):
'     Positions 1-47: Value1..Value47
'     Position 48: ID (from database ID field)
'     Positions 49-51: Value49..Value51
'     Position 52: Value52 (history/AZ column)
'==========================

Option Explicit

' Column index for ID in Kartei sheet and arrays
Private Const COL_ID As Long = 48
Private Const TOTAL_COLUMNS As Long = 52

' ============================================================
' PUBLIC FUNCTIONS
' ============================================================

' Retrieves a record from tblKartei (main/production table) by ID.
' Returns a 2D array arr(1 To 1, 1 To 52) with the record data.
' If the record is not found, returns Empty.
'
' This function opens and closes the database connection internally.
' Use this when you need a single lookup. For batch operations (multiple IDs),
' use GetMainRecordByIDFromDatabase() instead to avoid repeated open/close overhead.
'
' Design decision: Returns Empty (not Nothing) for "not found" because:
'   - Variant/Empty is easy to check with IsEmpty()
'   - Calling code can distinguish "not found" from errors
'   - Arrays cannot be Nothing, so Empty is more semantically correct
'
' @param dbPath - Full path to the Access database file
' @param idValue - The ID of the record to retrieve
' @return Variant - 2D array (1 To 1, 1 To 52) or Empty if not found
Public Function GetMainRecordByID(ByVal dbPath As String, ByVal idValue As Long) As Variant
    On Error GoTo ErrorHandler
    
    ' Validate inputs
    If Len(Trim(dbPath)) = 0 Then
        GetMainRecordByID = Empty
        Exit Function
    End If
    
    If idValue <= 0 Then
        GetMainRecordByID = Empty
        Exit Function
    End If
    
    ' Open database and query the record
    Dim db As DAO.Database
    Set db = OpenDatabaseSafe(dbPath)
    
    If db Is Nothing Then
        GetMainRecordByID = Empty
        Exit Function
    End If
    
    ' Execute query
    Dim rs As DAO.Recordset
    Set rs = OpenRecordsetForID(db, "tblKartei", idValue)
    
    If rs Is Nothing Then
        db.Close
        GetMainRecordByID = Empty
        Exit Function
    End If
    
    ' Check if record exists
    If rs.EOF Then
        rs.Close
        db.Close
        GetMainRecordByID = Empty
        Exit Function
    End If
    
    ' Map recordset to array
    Dim resultArray As Variant
    resultArray = MapRecordsetToArray(rs)
    
    ' Cleanup
    rs.Close
    db.Close
    
    GetMainRecordByID = resultArray
    Exit Function
    
ErrorHandler:
    On Error Resume Next
    If Not rs Is Nothing Then rs.Close
    If Not db Is Nothing Then db.Close
    On Error GoTo 0
    GetMainRecordByID = Empty
End Function

' Retrieves a record from tblKartei using an already-open database connection.
' Returns a 2D array arr(1 To 1, 1 To 52) with the record data.
' If the record is not found, returns Empty.
'
' This function does NOT open or close the database - caller is responsible
' for managing the database lifecycle. Use this for batch operations where
' multiple records need to be retrieved to avoid repeated open/close overhead.
'
' @param db - Already-open DAO.Database connection (must not be Nothing)
' @param idValue - The ID of the record to retrieve
' @return Variant - 2D array (1 To 1, 1 To 52) or Empty if not found
Public Function GetMainRecordByIDFromDatabase(ByVal db As DAO.Database, ByVal idValue As Long) As Variant
    On Error GoTo ErrorHandler
    
    ' Validate inputs
    If db Is Nothing Then
        GetMainRecordByIDFromDatabase = Empty
        Exit Function
    End If
    
    If idValue <= 0 Then
        GetMainRecordByIDFromDatabase = Empty
        Exit Function
    End If
    
    ' Execute query using the provided open database connection
    Dim rs As DAO.Recordset
    Set rs = OpenRecordsetForID(db, "tblKartei", idValue)
    
    If rs Is Nothing Then
        GetMainRecordByIDFromDatabase = Empty
        Exit Function
    End If
    
    ' Check if record exists
    If rs.EOF Then
        rs.Close
        GetMainRecordByIDFromDatabase = Empty
        Exit Function
    End If
    
    ' Map recordset to array
    Dim resultArray As Variant
    resultArray = MapRecordsetToArray(rs)
    
    ' Cleanup recordset only - do NOT close the database
    rs.Close
    
    GetMainRecordByIDFromDatabase = resultArray
    Exit Function
    
ErrorHandler:
    On Error Resume Next
    If Not rs Is Nothing Then rs.Close
    ' Do NOT close the database here - caller manages the connection
    On Error GoTo 0
    GetMainRecordByIDFromDatabase = Empty
End Function

' Retrieves a pending record from the Kartei worksheet by row index.
' The Kartei sheet contains data loaded from pre_tblKartei.
' Returns a 2D array arr(1 To 1, 1 To 52) with the record data.
'
' @param wsKartei - Reference to the Kartei worksheet
' @param rowIndex - The row number to read (1-based, should be >= 3 since rows 1-2 are headers)
' @return Variant - 2D array (1 To 1, 1 To 52) or Empty if invalid
Public Function GetPendingRecordByRow(ByVal wsKartei As Worksheet, ByVal rowIndex As Long) As Variant
    On Error GoTo ErrorHandler
    
    ' Validate inputs
    If wsKartei Is Nothing Then
        GetPendingRecordByRow = Empty
        Exit Function
    End If
    
    If rowIndex < 3 Then
        ' Rows 1-2 are headers, data starts at row 3
        GetPendingRecordByRow = Empty
        Exit Function
    End If
    
    ' Check if row has data (ID in column AV/48 should not be empty)
    If IsEmpty(wsKartei.Cells(rowIndex, COL_ID).Value) Then
        GetPendingRecordByRow = Empty
        Exit Function
    End If
    
    ' Map worksheet row to array
    Dim resultArray As Variant
    resultArray = MapWorksheetRowToArray(wsKartei, rowIndex)
    
    GetPendingRecordByRow = resultArray
    Exit Function
    
ErrorHandler:
    GetPendingRecordByRow = Empty
End Function

' Safely extracts and validates the ID from a Kartei worksheet row.
' Returns the ID as Long, or 0 if invalid/not found.
'
' @param wsKartei - Reference to the Kartei worksheet
' @param rowIndex - The row number to read
' @return Long - The validated ID value, or 0 if invalid
Public Function GetIDFromKarteiRow(ByVal wsKartei As Worksheet, ByVal rowIndex As Long) As Long
    On Error GoTo ErrorHandler
    
    GetIDFromKarteiRow = 0
    
    If wsKartei Is Nothing Then Exit Function
    If rowIndex < 3 Then Exit Function
    
    Dim cellValue As Variant
    cellValue = wsKartei.Cells(rowIndex, COL_ID).Value
    
    ' Validate the value
    If IsEmpty(cellValue) Or IsNull(cellValue) Then Exit Function
    If Len(Trim(CStr(cellValue))) = 0 Then Exit Function
    
    ' Try to convert to Long
    If IsNumeric(cellValue) Then
        Dim idVal As Long
        idVal = CLng(cellValue)
        
        ' ID must be positive
        If idVal > 0 Then
            GetIDFromKarteiRow = idVal
        End If
    End If
    
    Exit Function
    
ErrorHandler:
    GetIDFromKarteiRow = 0
End Function

' Returns a collection of all valid IDs from the Kartei worksheet.
' Useful for iterating over all pending records.
'
' @param wsKartei - Reference to the Kartei worksheet
' @return Collection - Collection of Long ID values
Public Function GetAllPendingIDs(ByVal wsKartei As Worksheet) As Collection
    Dim result As Collection
    Set result = New Collection
    
    On Error GoTo ErrorHandler
    
    If wsKartei Is Nothing Then
        Set GetAllPendingIDs = result
        Exit Function
    End If
    
    ' Find last row with data
    Dim lastRow As Long
    lastRow = wsKartei.Cells(wsKartei.Rows.Count, COL_ID).End(xlUp).Row
    
    If lastRow < 3 Then
        ' No data rows
        Set GetAllPendingIDs = result
        Exit Function
    End If
    
    ' Iterate through data rows and collect valid IDs
    Dim r As Long
    Dim idVal As Long
    
    For r = 3 To lastRow
        idVal = GetIDFromKarteiRow(wsKartei, r)
        If idVal > 0 Then
            result.Add idVal
        End If
    Next r
    
    Set GetAllPendingIDs = result
    Exit Function
    
ErrorHandler:
    Set GetAllPendingIDs = result
End Function

' ============================================================
' PRIVATE HELPER FUNCTIONS - Database Access
' ============================================================

' Safely opens a DAO database connection.
' Returns Nothing if the database cannot be opened.
'
' @param dbPath - Full path to the Access database file
' @return DAO.Database - Open database connection or Nothing
Private Function OpenDatabaseSafe(ByVal dbPath As String) As DAO.Database
    On Error GoTo ErrorHandler
    
    Dim engine As DAO.DBEngine
    Set engine = New DAO.DBEngine
    
    Dim wsDao As DAO.Workspace
    Set wsDao = engine.Workspaces(0)
    
    Set OpenDatabaseSafe = wsDao.OpenDatabase(dbPath)
    Exit Function
    
ErrorHandler:
    Set OpenDatabaseSafe = Nothing
End Function

' Opens a recordset for a specific ID from a table.
' Returns Nothing if the query fails.
'
' @param db - Open database connection
' @param tableName - Name of the table to query
' @param idValue - The ID value to search for
' @return DAO.Recordset - Open recordset or Nothing
Private Function OpenRecordsetForID(ByVal db As DAO.Database, _
                                    ByVal tableName As String, _
                                    ByVal idValue As Long) As DAO.Recordset
    On Error GoTo ErrorHandler
    
    Dim sqlStr As String
    sqlStr = "SELECT * FROM [" & tableName & "] WHERE ID = " & idValue
    
    Set OpenRecordsetForID = db.OpenRecordset(sqlStr, dbOpenSnapshot)
    Exit Function
    
ErrorHandler:
    Set OpenRecordsetForID = Nothing
End Function

' ============================================================
' PRIVATE HELPER FUNCTIONS - Data Mapping
' ============================================================

' Maps a DAO recordset row to a 2D array (1 To 1, 1 To 52).
' Field mapping:
'   - ID field -> position 48
'   - Value1..Value47 -> positions 1..47
'   - Value49..Value51 -> positions 49..51
'   - Value52 -> position 52 (history)
'
' @param rs - Open recordset positioned at the desired row
' @return Variant - 2D array (1 To 1, 1 To 52)
Private Function MapRecordsetToArray(ByVal rs As DAO.Recordset) As Variant
    Dim arr As Variant
    ReDim arr(1 To 1, 1 To TOTAL_COLUMNS)
    
    On Error Resume Next
    
    Dim c As Long
    
    ' Map Value1..Value47 to positions 1..47
    For c = 1 To 47
        ' Phone columns (7=Tel., 8=Handy): must be stored as strings to preserve
        ' leading zeros and prevent scientific notation (e.g. "0176..." not "1.76E+9")
        ' Also normalize any existing scientific-notation strings (e.g. "1,76E+9" -> "176000000")
        If c = 7 Or c = 8 Then
            arr(1, c) = phone_Normalize.NormalizePhoneText(rs.Fields("Value" & c).Value)
        Else
            arr(1, c) = NullToEmpty(rs.Fields("Value" & c).Value)
        End If
    Next c
    
    ' Map ID to position 48
    arr(1, COL_ID) = NullToEmpty(rs.Fields("ID").Value)
    
    ' Map Value49..Value51 to positions 49..51
    For c = 49 To 51
        arr(1, c) = NullToEmpty(rs.Fields("Value" & c).Value)
    Next c
    
    ' Map Value52 (history) to position 52
    arr(1, TOTAL_COLUMNS) = NullToEmpty(rs.Fields("Value52").Value)
    
    On Error GoTo 0
    
    MapRecordsetToArray = arr
End Function

' Maps a worksheet row to a 2D array (1 To 1, 1 To 52).
' Reads columns A-AZ (1-52) directly.
'
' @param ws - Worksheet to read from
' @param rowIndex - Row number to read
' @return Variant - 2D array (1 To 1, 1 To 52)
Private Function MapWorksheetRowToArray(ByVal ws As Worksheet, ByVal rowIndex As Long) As Variant
    Dim arr As Variant
    ReDim arr(1 To 1, 1 To TOTAL_COLUMNS)
    
    On Error Resume Next
    
    Dim c As Long
    For c = 1 To TOTAL_COLUMNS
        ' Phone columns (7=Tel., 8=Handy): use .Text to preserve leading zeros
        ' and prevent scientific notation. .Text returns the displayed string.
        ' Also normalize any existing scientific-notation strings (e.g. "1,76E+9" -> "176000000")
        If c = 7 Or c = 8 Then
            arr(1, c) = phone_Normalize.NormalizePhoneText(ws.Cells(rowIndex, c).Text)
        Else
            arr(1, c) = EmptyToEmpty(ws.Cells(rowIndex, c).Value)
        End If
    Next c
    
    On Error GoTo 0
    
    MapWorksheetRowToArray = arr
End Function

' ============================================================
' PRIVATE HELPER FUNCTIONS - Value Conversion
' ============================================================

' Converts Null to empty string.
' Used for database field values.
'
' @param v - Value to convert
' @return Variant - Original value or empty string if Null
Private Function NullToEmpty(ByVal v As Variant) As Variant
    If IsNull(v) Then
        NullToEmpty = ""
    Else
        NullToEmpty = v
    End If
End Function

' Converts Empty to empty string for consistency.
' Used for worksheet cell values.
'
' @param v - Value to convert
' @return Variant - Original value or empty string if Empty
Private Function EmptyToEmpty(ByVal v As Variant) As Variant
    If IsEmpty(v) Then
        EmptyToEmpty = ""
    ElseIf IsNull(v) Then
        EmptyToEmpty = ""
    Else
        EmptyToEmpty = v
    End If
End Function

' ============================================================
' UTILITY FUNCTIONS
' ============================================================

' Gets the validated database path using valid_DatabasePath module.
' Convenience wrapper for external callers.
' Defaults to year 25 for backward compatibility.
'
' @return String - Full path to database or empty string if cancelled
Public Function GetDatabasePath() As String
    GetDatabasePath = valid_DatabasePath.GetValidatedDatabasePath()
End Function

' Gets the validated database path for a specific year.
' Use this for multi-year operations.
'
' @param year2 - Two-digit year (24, 25, or 26)
' @return String - Full path to database or empty string if cancelled
Public Function GetDatabasePathForYear(ByVal year2 As Integer) As String
    GetDatabasePathForYear = valid_DatabasePath.GetValidatedDatabasePathForYear(year2)
End Function

' Checks if a record with the given ID exists in tblKartei.
' Useful for determining if a pending record is new or an update.
'
' @param dbPath - Full path to the Access database file
' @param idValue - The ID to check
' @return Boolean - True if record exists, False otherwise
Public Function RecordExistsInMain(ByVal dbPath As String, ByVal idValue As Long) As Boolean
    On Error GoTo ErrorHandler
    
    RecordExistsInMain = False
    
    If Len(Trim(dbPath)) = 0 Or idValue <= 0 Then Exit Function
    
    Dim db As DAO.Database
    Set db = OpenDatabaseSafe(dbPath)
    
    If db Is Nothing Then Exit Function
    
    Dim rs As DAO.Recordset
    Dim sqlStr As String
    sqlStr = "SELECT ID FROM tblKartei WHERE ID = " & idValue
    
    Set rs = db.OpenRecordset(sqlStr, dbOpenSnapshot)
    
    RecordExistsInMain = Not rs.EOF
    
    rs.Close
    db.Close
    Exit Function
    
ErrorHandler:
    On Error Resume Next
    If Not rs Is Nothing Then rs.Close
    If Not db Is Nothing Then db.Close
    On Error GoTo 0
    RecordExistsInMain = False
End Function
