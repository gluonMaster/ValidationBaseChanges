Attribute VB_Name = "Export_DeclinedTools"
'==========================
'   Declined Records Management Tools
'   Provides Admin interface to review and fix declined records
'==========================
Option Explicit

Private Const DECLINED_SHEET_NAME As String = "DeclinedOverview"

' ========================================
' Main Interface Procedures
' ========================================

Public Sub ShowDeclinedOverview()
    ' Main procedure to create/update DeclinedOverview sheet
    ' Shows all records from decl_tblKartei with key information
    
    On Error GoTo ErrorHandler
    
    Application.ScreenUpdating = False
    Application.Calculation = xlCalculationManual
    
    ' Get declined records from database
    Dim declinedRecords As Collection
    Set declinedRecords = LoadDeclinedRecords()
    
    If declinedRecords.Count = 0 Then
        MsgBox "No declined records found in the database.", vbInformation, "Declined Overview"
        GoTo Cleanup
    End If
    
    ' Create or clear the overview sheet
    Dim ws As Worksheet
    Set ws = GetOrCreateDeclinedSheet()
    
    ' Build the overview
    Call BuildDeclinedOverview(ws, declinedRecords)
    
    ' Activate and show the sheet
    ws.Activate
    ws.Range("A1").Select
    
    MsgBox "Declined overview created with " & declinedRecords.Count & " record(s)." & vbCrLf & vbCrLf & _
           "Review the records, make corrections, and click 'Apply Fixes' to move them to pending approval.", _
           vbInformation, "Declined Overview"
    
Cleanup:
    Application.Calculation = xlCalculationAutomatic
    Application.ScreenUpdating = True
    Exit Sub
    
ErrorHandler:
    MsgBox "Error creating declined overview: " & Err.Description, vbCritical, "Error"
    Resume Cleanup
End Sub

Public Sub ApplyDeclinedFixes()
    ' Applies fixes from DeclinedOverview sheet
    ' Moves corrected records from decl_tblKartei to pre_tblKartei
    ' Removes them from decl_tblKartei
    
    On Error GoTo ErrorHandler
    
    ' Check if DeclinedOverview exists
    Dim ws As Worksheet
    On Error Resume Next
    Set ws = ThisWorkbook.Worksheets(DECLINED_SHEET_NAME)
    On Error GoTo ErrorHandler
    
    If ws Is Nothing Then
        MsgBox "DeclinedOverview sheet not found. Please run 'Show Declined Overview' first.", _
               vbExclamation, "Sheet Not Found"
        Exit Sub
    End If
    
    ' Confirm action
    Dim response As VbMsgBoxResult
    response = MsgBox("This will move selected/all declined records to pending approval (pre_tblKartei)." & vbCrLf & _
                      "Records will be removed from declined status." & vbCrLf & vbCrLf & _
                      "Continue?", vbYesNo + vbQuestion, "Confirm Apply Fixes")
    
    If response <> vbYes Then
        Exit Sub
    End If
    
    Application.ScreenUpdating = False
    Application.Calculation = xlCalculationManual
    
    ' Collect IDs to process
    Dim idsToProcess As Collection
    Set idsToProcess = CollectIDsFromOverview(ws)
    
    If idsToProcess.Count = 0 Then
        MsgBox "No valid IDs found on DeclinedOverview sheet.", vbExclamation, "No Data"
        GoTo Cleanup
    End If
    
    ' Process the fixes
    Dim movedCount As Long
    movedCount = ProcessDeclinedFixes(ws, idsToProcess)
    
    ' Refresh the overview
    If movedCount > 0 Then
        Call ShowDeclinedOverview
        MsgBox "Successfully moved " & movedCount & " record(s) from declined to pending approval." & vbCrLf & _
               "Records are now in pre_tblKartei awaiting Superadmin review.", _
               vbInformation, "Fixes Applied"
    Else
        MsgBox "No records were processed.", vbInformation, "Apply Fixes"
    End If
    
Cleanup:
    Application.Calculation = xlCalculationAutomatic
    Application.ScreenUpdating = True
    Exit Sub
    
ErrorHandler:
    MsgBox "Error applying fixes: " & Err.Description, vbCritical, "Error"
    Resume Cleanup
End Sub

' ========================================
' Data Loading Functions
' ========================================

Private Function LoadDeclinedRecords() As Collection
    ' Loads all records from decl_tblKartei
    ' Returns collection of arrays containing record data
    
    Dim records As New Collection
    
    Dim engine As DAO.DBEngine
    Set engine = New DAO.DBEngine
    
    Dim wsDao As DAO.Workspace
    Set wsDao = engine.Workspaces(0)
    
    Dim dbPath As String
    dbPath = ThisWorkbook.Worksheets("Kartei").Range("X1").value & "\Alarm\KindElternDaten_25_front.accdb"
    
    Dim db As DAO.Database
    Set db = wsDao.OpenDatabase(dbPath)
    
    On Error Resume Next
    Dim rs As DAO.Recordset
    Set rs = db.OpenRecordset("SELECT * FROM decl_tblKartei ORDER BY ID", dbOpenSnapshot)
    
    If Err.Number <> 0 Then
        ' Table doesn't exist yet
        On Error GoTo 0
        db.Close
        Set LoadDeclinedRecords = records
        Exit Function
    End If
    On Error GoTo 0
    
    If rs.EOF Then
        rs.Close
        db.Close
        Set LoadDeclinedRecords = records
        Exit Function
    End If
    
    ' Read all records
    Do While Not rs.EOF
        ' Create array to hold record data
        ' Index: 0=ID, 1=A, 2=B, 3=C, 4=F, 5=J, 6=O, 7-18=Months, 19=AZ, 20=DeclCount
        Dim recData(0 To 20) As Variant
        
        ' Load key fields
        recData(0) = rs.Fields("ID").value
        recData(1) = NzStr(rs.Fields("Value1").value)
        recData(2) = NzStr(rs.Fields("Value2").value)
        recData(3) = NzStr(rs.Fields("Value3").value)
        recData(4) = NzStr(rs.Fields("Value6").value)   ' Address
        recData(5) = NzStr(rs.Fields("Value10").value)  ' Subject1
        recData(6) = NzStr(rs.Fields("Value15").value)  ' Subject2
        
        ' Load months U-AF (columns 21-32 = Value21-Value32)
        Dim m As Integer
        For m = 1 To 12
            recData(6 + m) = NzStr(rs.Fields("Value" & (20 + m)).value)
        Next m
        
        recData(19) = NzStr(rs.Fields("Value52").value) ' History
        
        ' Count declinations from history
        recData(20) = CountDeclinations(CStr(recData(19)))
        
        records.Add recData
        rs.MoveNext
    Loop
    
    rs.Close
    db.Close
    
    Set LoadDeclinedRecords = records
End Function

Private Function CountDeclinations(ByVal historyText As String) As Integer
    ' Counts how many times "Decl_" appears in history string
    ' Each "Decl_n:" indicates a declination by Superadmin
    
    Dim count As Integer
    count = 0
    
    Dim pos As Long
    pos = 1
    
    Do
        pos = InStr(pos, historyText, "Decl_", vbTextCompare)
        If pos > 0 Then
            count = count + 1
            pos = pos + 5
        End If
    Loop While pos > 0
    
    CountDeclinations = count
End Function

Private Function NzStr(ByVal value As Variant) As String
    ' Converts null/empty to empty string
    If IsNull(value) Or IsEmpty(value) Then
        NzStr = ""
    Else
        NzStr = CStr(value)
    End If
End Function

' ========================================
' Sheet Management Functions
' ========================================

Private Function GetOrCreateDeclinedSheet() As Worksheet
    ' Gets existing DeclinedOverview sheet or creates new one
    
    Dim ws As Worksheet
    On Error Resume Next
    Set ws = ThisWorkbook.Worksheets(DECLINED_SHEET_NAME)
    On Error GoTo 0
    
    If ws Is Nothing Then
        ' Create new sheet
        Set ws = ThisWorkbook.Worksheets.Add(After:=ThisWorkbook.Worksheets(ThisWorkbook.Worksheets.Count))
        ws.Name = DECLINED_SHEET_NAME
    Else
        ' Clear existing content
        ws.Cells.Clear
    End If
    
    Set GetOrCreateDeclinedSheet = ws
End Function

Private Sub BuildDeclinedOverview(ByVal ws As Worksheet, ByVal records As Collection)
    ' Builds the overview table on the worksheet
    
    ' Set up headers
    With ws
        .Range("A1").value = "ID"
        .Range("B1").value = "Value A"
        .Range("C1").value = "Value B"
        .Range("D1").value = "Value C"
        .Range("E1").value = "Address (F)"
        .Range("F1").value = "Subject1 (J)"
        .Range("G1").value = "Subject2 (O)"
        .Range("H1").value = "Month 1"
        .Range("I1").value = "Month 2"
        .Range("J1").value = "Month 3"
        .Range("K1").value = "Month 4"
        .Range("L1").value = "Month 5"
        .Range("M1").value = "Month 6"
        .Range("N1").value = "Month 7"
        .Range("O1").value = "Month 8"
        .Range("P1").value = "Month 9"
        .Range("Q1").value = "Month 10"
        .Range("R1").value = "Month 11"
        .Range("S1").value = "Month 12"
        .Range("T1").value = "Declined Times"
        .Range("U1").value = "History (AZ)"
        
        ' Format header row
        .Range("A1:U1").Font.Bold = True
        .Range("A1:U1").Interior.Color = RGB(200, 200, 200)
        .Range("A1:U1").HorizontalAlignment = xlCenter
    End With
    
    ' Fill data rows
    Dim rowNum As Long
    rowNum = 2
    
    Dim rec As Variant
    For Each rec In records
        With ws
            .Cells(rowNum, 1).value = rec(0)  ' ID
            .Cells(rowNum, 2).value = rec(1)  ' A
            .Cells(rowNum, 3).value = rec(2)  ' B
            .Cells(rowNum, 4).value = rec(3)  ' C
            .Cells(rowNum, 5).value = rec(4)  ' F
            .Cells(rowNum, 6).value = rec(5)  ' J
            .Cells(rowNum, 7).value = rec(6)  ' O
            
            ' Months
            Dim m As Integer
            For m = 1 To 12
                .Cells(rowNum, 7 + m).value = rec(6 + m)
            Next m
            
            .Cells(rowNum, 20).value = rec(20)  ' DeclCount
            .Cells(rowNum, 21).value = rec(19)  ' AZ
            
            ' Highlight rows with multiple declinations
            If CInt(rec(20)) > 1 Then
                .Range(.Cells(rowNum, 1), .Cells(rowNum, 21)).Interior.Color = RGB(255, 200, 200)
            End If
        End With
        
        rowNum = rowNum + 1
    Next rec
    
    ' Auto-fit columns
    ws.Columns("A:U").AutoFit
    
    ' Freeze header row
    ws.Range("A2").Select
    ActiveWindow.FreezePanes = True
End Sub

' ========================================
' Fix Application Functions
' ========================================

Private Function CollectIDsFromOverview(ByVal ws As Worksheet) As Collection
    ' Collects all valid IDs from the overview sheet
    
    Dim ids As New Collection
    
    Dim lastRow As Long
    lastRow = ws.Cells(ws.Rows.Count, 1).End(xlUp).Row
    
    If lastRow < 2 Then
        Set CollectIDsFromOverview = ids
        Exit Function
    End If
    
    Dim r As Long
    For r = 2 To lastRow
        Dim idValue As Variant
        idValue = ws.Cells(r, 1).value
        
        If Not IsEmpty(idValue) And IsNumeric(idValue) Then
            Dim idLong As Long
            idLong = CLng(idValue)
            If idLong > 0 Then
                ids.Add idLong
            End If
        End If
    Next r
    
    Set CollectIDsFromOverview = ids
End Function

Private Function ProcessDeclinedFixes(ByVal ws As Worksheet, ByVal idsToProcess As Collection) As Long
    ' Processes fixes for declined records
    ' Reads data from Kartei sheet (by ID), writes to pre_tblKartei, removes from decl_tblKartei
    
    Dim movedCount As Long
    movedCount = 0
    
    Dim engine As DAO.DBEngine
    Set engine = New DAO.DBEngine
    
    Dim wsDao As DAO.Workspace
    Set wsDao = engine.Workspaces(0)
    
    Dim dbPath As String
    dbPath = ThisWorkbook.Worksheets("Kartei").Range("X1").value & "\Alarm\KindElternDaten_25_front.accdb"
    
    Dim db As DAO.Database
    Set db = wsDao.OpenDatabase(dbPath)
    
    ' Ensure pre_tblKartei exists
    Call EnsurePreTableExists(db)
    
    wsDao.BeginTrans
    
    On Error GoTo RollbackTrans
    
    Dim wsKartei As Worksheet
    Set wsKartei = ThisWorkbook.Worksheets("Kartei")
    
    Dim idToProcess As Variant
    For Each idToProcess In idsToProcess
        Dim targetID As Long
        targetID = CLng(idToProcess)
        
        ' Find row on Kartei sheet by ID (column AV = 48)
        Dim karteiRow As Long
        karteiRow = FindRowByIDInKartei(wsKartei, targetID)
        
        If karteiRow > 0 Then
            ' Read current data from Kartei sheet
            Dim rowData As Variant
            rowData = wsKartei.Range(wsKartei.Cells(karteiRow, 1), wsKartei.Cells(karteiRow, 52)).Value2
            
            ' Read formats
            Dim rowFormats() As Variant
            ReDim rowFormats(1 To 53)
            
            Dim c As Integer
            For c = 1 To 51
                rowFormats(c) = wsKartei.Cells(karteiRow, c).Interior.Color
            Next c
            rowFormats(52) = wsKartei.Cells(karteiRow, 3).Font.Color
            rowFormats(53) = wsKartei.Cells(karteiRow, 18).Font.Color
            
            ' Write to pre_tblKartei
            Call WriteRecordToPreTable(db, targetID, rowData, rowFormats)
            
            ' Delete from decl_tblKartei
            db.Execute "DELETE FROM decl_tblKartei WHERE ID = " & targetID
            
            ' Clear DECLINED status on Kartei sheet
            wsKartei.Cells(karteiRow, 53).ClearContents  ' BA column
            wsKartei.Cells(karteiRow, 1).Interior.ColorIndex = xlColorIndexNone
            
            movedCount = movedCount + 1
        End If
    Next idToProcess
    
    wsDao.CommitTrans
    db.Close
    
    ProcessDeclinedFixes = movedCount
    Exit Function
    
RollbackTrans:
    wsDao.Rollback
    db.Close
    ProcessDeclinedFixes = 0
    Err.Raise Err.Number, Err.Source, Err.Description
End Function

Private Function FindRowByIDInKartei(ByVal ws As Worksheet, ByVal targetID As Long) As Long
    ' Finds row in Kartei sheet by ID (column AV = 48)
    
    Dim lastRow As Long
    lastRow = ws.Cells(ws.Rows.Count, 1).End(xlUp).Row
    
    Dim r As Long
    For r = 3 To lastRow
        Dim cellID As Variant
        cellID = ws.Cells(r, 48).value
        
        If Not IsEmpty(cellID) And IsNumeric(cellID) Then
            If CLng(cellID) = targetID Then
                FindRowByIDInKartei = r
                Exit Function
            End If
        End If
    Next r
    
    FindRowByIDInKartei = 0
End Function

Private Sub WriteRecordToPreTable(ByVal db As DAO.Database, _
                                  ByVal targetID As Long, _
                                  ByVal rowData As Variant, _
                                  ByVal rowFormats As Variant)
    ' Writes a record to pre_tblKartei with explicit ID
    
    ' Check if record already exists
    Dim rsCheck As DAO.Recordset
    Set rsCheck = db.OpenRecordset("SELECT * FROM pre_tblKartei WHERE ID = " & targetID, dbOpenDynaset)
    
    Dim rs As DAO.Recordset
    
    If rsCheck.EOF Then
        ' Create new record
        rsCheck.Close
        Set rs = db.OpenRecordset("pre_tblKartei", dbOpenDynaset)
        rs.AddNew
        rs.Fields("ID").value = targetID
    Else
        ' Update existing record
        Set rs = rsCheck
        rs.Edit
    End If
    
    ' Fill fields
    Dim c As Long
    For c = 1 To 51
        Dim fieldValue As Variant
        fieldValue = rowData(1, c)
        
        If Not IsError(fieldValue) And Not IsEmpty(fieldValue) Then
            rs.Fields("Value" & c).value = fieldValue
        Else
            rs.Fields("Value" & c).value = ""
        End If
        
        rs.Fields("InteriorColor" & c).value = rowFormats(c)
    Next c
    
    ' Value52 (history)
    If Not IsError(rowData(1, 52)) And Not IsEmpty(rowData(1, 52)) Then
        rs.Fields("Value52").value = rowData(1, 52)
    Else
        rs.Fields("Value52").value = ""
    End If
    
    ' Font colors
    rs.Fields("FontColor3").value = rowFormats(52)
    rs.Fields("FontColor18").value = rowFormats(53)
    
    rs.Update
    rs.Close
End Sub

