Attribute VB_Name = "Kind_ErrorLog"
Option Explicit

' =============================================================================
' Module: Kind_ErrorLog
' Purpose: Error logging utilities - writes to ErrorLog sheet starting at column E
' Prefix: Kind_ (all modules/classes in this project use this prefix)
' Note: Columns A:C are reserved, logging starts at column E
' =============================================================================

' -----------------------------------------------------------------------------
' ErrorLog Column Layout (starting at E)
' -----------------------------------------------------------------------------
Private Const ERRLOG_COL_TIMESTAMP As String = "E"
Private Const ERRLOG_COL_FAMILYID As String = "F"
Private Const ERRLOG_COL_NACHNAME As String = "G"
Private Const ERRLOG_COL_VORNAME As String = "H"
Private Const ERRLOG_COL_FACH As String = "I"
Private Const ERRLOG_COL_TYP As String = "J"
Private Const ERRLOG_COL_DETAILS As String = "K"

Private Const ERRLOG_HEADER_ROW As Long = 1


' -----------------------------------------------------------------------------
' Kind_EnsureErrorLogHeaders
' Purpose: Ensures header row exists in ErrorLog sheet (E1:K1)
'          Only writes if E1 is empty
' Parameters:
'   ws - ErrorLog worksheet reference
' -----------------------------------------------------------------------------
Public Sub Kind_EnsureErrorLogHeaders(ByVal ws As Worksheet)
    On Error GoTo ErrHandler
    
    ' Check if ws is valid
    If ws Is Nothing Then
        Exit Sub
    End If
    
    ' Only write headers if E1 is empty
    If Len(Trim$(CStr(ws.Range(ERRLOG_COL_TIMESTAMP & ERRLOG_HEADER_ROW).Value))) = 0 Then
        ws.Range(ERRLOG_COL_TIMESTAMP & ERRLOG_HEADER_ROW).Value = "Timestamp"
        ws.Range(ERRLOG_COL_FAMILYID & ERRLOG_HEADER_ROW).Value = "FamilyID"
        ws.Range(ERRLOG_COL_NACHNAME & ERRLOG_HEADER_ROW).Value = "Nachname"
        ws.Range(ERRLOG_COL_VORNAME & ERRLOG_HEADER_ROW).Value = "Vorname"
        ws.Range(ERRLOG_COL_FACH & ERRLOG_HEADER_ROW).Value = "Fach"
        ws.Range(ERRLOG_COL_TYP & ERRLOG_HEADER_ROW).Value = "Typ"
        ws.Range(ERRLOG_COL_DETAILS & ERRLOG_HEADER_ROW).Value = "Details"
        
        ' Format headers bold
        ws.Range(ERRLOG_COL_TIMESTAMP & ERRLOG_HEADER_ROW & ":" & ERRLOG_COL_DETAILS & ERRLOG_HEADER_ROW).Font.Bold = True
    End If
    
    Exit Sub
    
ErrHandler:
    ' Silent fail - logging should not crash main process
End Sub


' -----------------------------------------------------------------------------
' Kind_GetNextErrorLogRow
' Purpose: Determines the next available row for error logging
'          Calculates as: 1 + Max(lastRow(A), lastRow(C), lastRow(E))
' Parameters:
'   ws - ErrorLog worksheet reference
' Returns: Next available row number (minimum 2 to preserve header)
' -----------------------------------------------------------------------------
Public Function Kind_GetNextErrorLogRow(ByVal ws As Worksheet) As Long
    Dim lastRowA As Long
    Dim lastRowC As Long
    Dim lastRowE As Long
    Dim maxRow As Long
    
    On Error GoTo ErrHandler
    
    ' Default to row 2 (after header)
    Kind_GetNextErrorLogRow = 2
    
    If ws Is Nothing Then
        Exit Function
    End If
    
    ' Find last used row in each column
    lastRowA = GetLastRowInColumn(ws, "A")
    lastRowC = GetLastRowInColumn(ws, "C")
    lastRowE = GetLastRowInColumn(ws, "E")
    
    ' Calculate max
    maxRow = lastRowA
    If lastRowC > maxRow Then maxRow = lastRowC
    If lastRowE > maxRow Then maxRow = lastRowE
    
    ' Next row is max + 1, but at least 2 (preserve header row)
    Kind_GetNextErrorLogRow = maxRow + 1
    If Kind_GetNextErrorLogRow < 2 Then
        Kind_GetNextErrorLogRow = 2
    End If
    
    Exit Function
    
ErrHandler:
    Kind_GetNextErrorLogRow = 2
End Function


' -----------------------------------------------------------------------------
' GetLastRowInColumn (Private Helper)
' Purpose: Finds last used row in a specific column
' Parameters:
'   ws     - Worksheet reference
'   colLtr - Column letter
' Returns: Last used row number, or 0 if column is empty
' -----------------------------------------------------------------------------
Private Function GetLastRowInColumn(ByVal ws As Worksheet, ByVal colLtr As String) As Long
    Dim lastCell As Range
    
    On Error GoTo ErrHandler
    
    GetLastRowInColumn = 0
    
    Set lastCell = ws.Cells(ws.Rows.Count, colLtr).End(xlUp)
    If Not lastCell Is Nothing Then
        GetLastRowInColumn = lastCell.Row
    End If
    
    Exit Function
    
ErrHandler:
    GetLastRowInColumn = 0
End Function


' -----------------------------------------------------------------------------
' Kind_LogRuntimeError
' Purpose: Logs a runtime error to the ErrorLog sheet
' Parameters:
'   errNumber     - Error number (Err.Number)
'   errDesc       - Error description (Err.Description)
'   errSource     - Source/location of error (e.g. procedure name)
' Note: Silently fails if ErrorLog sheet not accessible
' -----------------------------------------------------------------------------
Public Sub Kind_LogRuntimeError(ByVal errNumber As Long, _
                                ByVal errDesc As String, _
                                ByVal errSource As String)
    Dim ws As Worksheet
    Dim nextRow As Long
    
    On Error Resume Next
    
    ' Try to get ErrorLog worksheet from ThisWorkbook
    If Not Kind_TryGetWorksheet(ThisWorkbook, KIND_THIS_SHEET_ERRORLOG, ws) Then
        Exit Sub
    End If
    
    ' Ensure headers exist
    Kind_EnsureErrorLogHeaders ws
    
    ' Get next available row
    nextRow = Kind_GetNextErrorLogRow(ws)
    
    ' Write error entry
    ws.Range(ERRLOG_COL_TIMESTAMP & nextRow).Value = Now
    ws.Range(ERRLOG_COL_FAMILYID & nextRow).Value = ""
    ws.Range(ERRLOG_COL_NACHNAME & nextRow).Value = ""
    ws.Range(ERRLOG_COL_VORNAME & nextRow).Value = ""
    ws.Range(ERRLOG_COL_FACH & nextRow).Value = errSource
    ws.Range(ERRLOG_COL_TYP & nextRow).Value = "RUNTIME_ERR"
    ws.Range(ERRLOG_COL_DETAILS & nextRow).Value = "Fehler " & errNumber & ": " & errDesc
    
    On Error GoTo 0
End Sub


' -----------------------------------------------------------------------------
' Kind_LogError
' Purpose: Logs an error entry to the ErrorLog sheet
' Parameters:
'   familyID  - Family ID
'   lastName  - Last name (Nachname)
'   firstName - First name (Vorname)
'   subject   - Subject (Fach)
'   errType   - Error type code
'   details   - Error details (German without umlauts)
' Note: Silently fails if ErrorLog sheet not accessible
' -----------------------------------------------------------------------------
Public Sub Kind_LogError(ByVal familyID As String, _
                         ByVal lastName As String, _
                         ByVal firstName As String, _
                         ByVal subject As String, _
                         ByVal errType As String, _
                         ByVal details As String)
    Dim ws As Worksheet
    Dim nextRow As Long
    
    On Error GoTo ErrHandler
    
    ' Try to get ErrorLog worksheet from ThisWorkbook
    If Not Kind_TryGetWorksheet(ThisWorkbook, KIND_THIS_SHEET_ERRORLOG, ws) Then
        ' Cannot log - sheet not found
        Exit Sub
    End If
    
    ' Ensure headers exist
    Kind_EnsureErrorLogHeaders ws
    
    ' Get next available row
    nextRow = Kind_GetNextErrorLogRow(ws)
    
    ' Write error entry
    ws.Range(ERRLOG_COL_TIMESTAMP & nextRow).Value = Now
    ws.Range(ERRLOG_COL_FAMILYID & nextRow).Value = familyID
    ws.Range(ERRLOG_COL_NACHNAME & nextRow).Value = lastName
    ws.Range(ERRLOG_COL_VORNAME & nextRow).Value = firstName
    ws.Range(ERRLOG_COL_FACH & nextRow).Value = subject
    ws.Range(ERRLOG_COL_TYP & nextRow).Value = errType
    ws.Range(ERRLOG_COL_DETAILS & nextRow).Value = details
    
    Exit Sub
    
ErrHandler:
    ' Silent fail - logging should not crash main process
End Sub
