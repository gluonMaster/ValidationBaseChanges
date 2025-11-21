Attribute VB_Name = "Export_OverlayPending"
'==========================
'   Overlay Pending/Declined Module
'   Loads pre_tblKartei and decl_tblKartei records and overlays them
'   on the Kartei sheet when Admin file opens
'==========================
Option Explicit

' Status column - BA (53)
Private Const STATUS_COL As Integer = 53

' Colors for visual marking
Private Const COLOR_PENDING As Long = 15849925  ' Light blue (RGB 149, 179, 242)
Private Const COLOR_DECLINED As Long = 5855743  ' Light red (RGB 255, 127, 89)

' ========================================
' Main Overlay Procedure
' ========================================

Public Sub OverlayPendingAndDeclined(ByVal ws As Worksheet)
    ' Main procedure to overlay pending and declined records on Kartei sheet
    ' Called after loading tblKartei data, before creating Kartei_Original
    
    On Error GoTo ErrorHandler
    
    Application.ScreenUpdating = False
    Application.Calculation = xlCalculationManual
    
    ' Open database connection
    Dim engine As DAO.DBEngine
    Set engine = New DAO.DBEngine
    
    Dim wsDao As DAO.Workspace
    Set wsDao = engine.Workspaces(0)
    
    Dim dbPath As String
    dbPath = ws.Range("X1").value & "\Alarm\KindElternDaten_25_front.accdb"
    
    Dim db As DAO.Database
    Set db = wsDao.OpenDatabase(dbPath)
    
    ' Collections to track IDs processed
    Dim pendingIDs As New Collection
    Dim declinedIDs As New Collection
    Dim conflictIDs As New Collection
    
    ' Process pre_tblKartei records
    Call ProcessPreTable(db, ws, pendingIDs)
    
    ' Process decl_tblKartei records
    Call ProcessDeclTable(db, ws, declinedIDs, pendingIDs, conflictIDs)
    
    ' Close database
    db.Close
    Set db = Nothing
    
    ' Notify user if there are conflicts (same ID in both pre_ and decl_)
    If conflictIDs.count > 0 Then
        Dim conflictMsg As String
        conflictMsg = "Data consistency error detected:" & vbCrLf & _
                      "The following IDs exist in both pending and declined tables:" & vbCrLf & vbCrLf
        
        Dim conflictID As Variant
        For Each conflictID In conflictIDs
            conflictMsg = conflictMsg & "  - ID " & conflictID & vbCrLf
        Next conflictID
        
        conflictMsg = conflictMsg & vbCrLf & _
                      "Declined version takes priority and is displayed." & vbCrLf & _
                      "Please contact Superadmin to resolve this inconsistency."
        
        MsgBox conflictMsg, vbExclamation, "Data Consistency Error"
    End If
    
    ' Notify user if there are pending or declined records
    If pendingIDs.count > 0 Or declinedIDs.count > 0 Then
        Dim summaryMsg As String
        summaryMsg = "The Kartei sheet contains records requiring attention:" & vbCrLf & vbCrLf & _
                     "Pending (awaiting approval): " & pendingIDs.count & " record(s)" & vbCrLf & _
                     "Declined (rejected): " & declinedIDs.count & " record(s)" & vbCrLf & vbCrLf & _
                     "Pending records are highlighted in light blue." & vbCrLf & _
                     "Declined records are highlighted in light red."
        
        MsgBox summaryMsg, vbInformation, "Pending/Declined Records"
    End If
    
Cleanup:
    Application.Calculation = xlCalculationAutomatic
    Application.ScreenUpdating = True
    Exit Sub
    
ErrorHandler:
    MsgBox "Error during overlay operation: " & Err.Description, vbCritical, "Overlay Error"
    Resume Cleanup
End Sub

' ========================================
' Processing Functions
' ========================================

Private Sub ProcessPreTable(ByVal db As DAO.Database, _
                            ByVal ws As Worksheet, _
                            ByVal pendingIDs As Collection)
    ' Reads all records from pre_tblKartei and overlays them on Kartei sheet
    ' Marks them as PENDING and colors column A light blue
    
    On Error Resume Next
    Dim rs As DAO.Recordset
    Set rs = db.OpenRecordset("SELECT * FROM pre_tblKartei", dbOpenSnapshot)
    
    If Err.Number <> 0 Then
        ' Table doesn't exist yet, skip
        On Error GoTo 0
        Exit Sub
    End If
    On Error GoTo 0
    
    If rs.EOF Then
        rs.Close
        Exit Sub
    End If
    
    ' Process each record
    Do While Not rs.EOF
        Dim recordID As Long
        recordID = rs.Fields("ID").value
        
        ' Find row on sheet by ID (column AV = 48)
        Dim targetRow As Long
        targetRow = FindRowByID(ws, recordID)
        
        If targetRow > 0 Then
            ' Overlay record data
            Call OverlayRecordData(ws, targetRow, rs)
            
            ' Mark as PENDING
            ws.Cells(targetRow, STATUS_COL).value = "PENDING"
            
            ' Color column A light blue if D <> "Zahlung"
            Dim cellD As String
            cellD = Trim(CStr(ws.Cells(targetRow, 4).value))
            
            If cellD <> "Zahlung" Then
                ws.Cells(targetRow, 1).Interior.Color = COLOR_PENDING
            End If
            
            ' Track this ID
            pendingIDs.Add recordID
        End If
        
        rs.MoveNext
    Loop
    
    rs.Close
End Sub

Private Sub ProcessDeclTable(ByVal db As DAO.Database, _
                             ByVal ws As Worksheet, _
                             ByVal declinedIDs As Collection, _
                             ByVal pendingIDs As Collection, _
                             ByVal conflictIDs As Collection)
    ' Reads all records from decl_tblKartei and overlays them on Kartei sheet
    ' Marks them as DECLINED and colors column A light red
    ' Checks for conflicts with pending records
    
    On Error Resume Next
    Dim rs As DAO.Recordset
    Set rs = db.OpenRecordset("SELECT * FROM decl_tblKartei", dbOpenSnapshot)
    
    If Err.Number <> 0 Then
        ' Table doesn't exist yet, skip
        On Error GoTo 0
        Exit Sub
    End If
    On Error GoTo 0
    
    If rs.EOF Then
        rs.Close
        Exit Sub
    End If
    
    ' Process each record
    Do While Not rs.EOF
        Dim recordID As Long
        recordID = rs.Fields("ID").value
        
        ' Check if this ID was already in pending
        If IDExistsInCollection(pendingIDs, recordID) Then
            conflictIDs.Add recordID
        End If
        
        ' Find row on sheet by ID (column AV = 48)
        Dim targetRow As Long
        targetRow = FindRowByID(ws, recordID)
        
        If targetRow > 0 Then
            ' Overlay record data (declined takes priority)
            Call OverlayRecordData(ws, targetRow, rs)
            
            ' Mark as DECLINED
            ws.Cells(targetRow, STATUS_COL).value = "DECLINED"
            
            ' Color column A light red if D <> "Zahlung"
            Dim cellD As String
            cellD = Trim(CStr(ws.Cells(targetRow, 4).value))
            
            If cellD <> "Zahlung" Then
                ws.Cells(targetRow, 1).Interior.Color = COLOR_DECLINED
            End If
            
            ' Track this ID
            declinedIDs.Add recordID
        End If
        
        rs.MoveNext
    Loop
    
    rs.Close
End Sub

Private Sub OverlayRecordData(ByVal ws As Worksheet, _
                              ByVal targetRow As Long, _
                              ByVal rs As DAO.Recordset)
    ' Overlays all field values and formats from recordset to sheet row
    ' Columns 1-51: Value1..Value51 with InteriorColor1..InteriorColor51
    ' Column 52: Value52 (history in AZ)
    ' Formats: FontColor3, FontColor18
    
    Dim c As Long
    
    ' Overlay values and interior colors for columns 1-51
    For c = 1 To 51
        Dim fieldValue As Variant
        fieldValue = rs.Fields("Value" & c).value
        
        If Not IsNull(fieldValue) Then
            ws.Cells(targetRow, c).value = fieldValue
        Else
            ws.Cells(targetRow, c).value = ""
        End If
        
        ' Interior color
        Dim interiorColor As Long
        interiorColor = rs.Fields("InteriorColor" & c).value
        
        If Not IsNull(interiorColor) Then
            ws.Cells(targetRow, c).Interior.Color = interiorColor
        End If
    Next c
    
    ' Overlay Value52 (history in AZ, column 52)
    Dim value52 As Variant
    value52 = rs.Fields("Value52").value
    
    If Not IsNull(value52) Then
        ws.Cells(targetRow, 52).value = value52
    Else
        ws.Cells(targetRow, 52).value = ""
    End If
    
    ' Overlay font colors
    Dim fontColor3 As Long
    fontColor3 = rs.Fields("FontColor3").value
    
    If Not IsNull(fontColor3) Then
        ws.Cells(targetRow, 3).Font.Color = fontColor3
    End If
    
    Dim fontColor18 As Long
    fontColor18 = rs.Fields("FontColor18").value
    
    If Not IsNull(fontColor18) Then
        ws.Cells(targetRow, 18).Font.Color = fontColor18
    End If
End Sub

' ========================================
' Helper Functions
' ========================================

Private Function FindRowByID(ByVal ws As Worksheet, ByVal recordID As Long) As Long
    ' Finds the row on sheet where column AV (48) matches the given ID
    ' Returns row number or 0 if not found
    
    Dim lastRow As Long
    lastRow = ws.Cells(ws.Rows.count, 1).End(xlUp).row
    
    Dim r As Long
    For r = 3 To lastRow
        Dim cellID As Variant
        cellID = ws.Cells(r, 48).value  ' Column AV
        
        If Not IsEmpty(cellID) And IsNumeric(cellID) Then
            If CLng(cellID) = recordID Then
                FindRowByID = r
                Exit Function
            End If
        End If
    Next r
    
    ' Not found
    FindRowByID = 0
End Function

Private Function IDExistsInCollection(ByVal col As Collection, ByVal id As Long) As Boolean
    ' Checks if an ID exists in a collection
    
    Dim item As Variant
    For Each item In col
        If CLng(item) = id Then
            IDExistsInCollection = True
            Exit Function
        End If
    Next item
    
    IDExistsInCollection = False
End Function

' ========================================
' Status Check Functions
' ========================================

Public Function IsPendingRow(ByVal ws As Worksheet, ByVal row As Long) As Boolean
    ' Checks if a row is marked as PENDING in status column BA (53)
    Dim status As String
    status = Trim(CStr(ws.Cells(row, STATUS_COL).value))
    IsPendingRow = (status = "PENDING")
End Function

Public Function IsDeclinedRow(ByVal ws As Worksheet, ByVal row As Long) As Boolean
    ' Checks if a row is marked as DECLINED in status column BA (53)
    Dim status As String
    status = Trim(CStr(ws.Cells(row, STATUS_COL).value))
    IsDeclinedRow = (status = "DECLINED")
End Function

Public Function GetRowStatus(ByVal ws As Worksheet, ByVal row As Long) As String
    ' Returns the status of a row: "PENDING", "DECLINED", or ""
    GetRowStatus = Trim(CStr(ws.Cells(row, STATUS_COL).value))
End Function
