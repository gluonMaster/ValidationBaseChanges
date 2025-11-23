Attribute VB_Name = "Export_DeclinedTools"
'==========================
'   Declined Records Management Tools
'   Provides Admin interface to review and fix declined records
'==========================
Option Explicit

Private Const DECLINED_SHEET_NAME As String = "DeclinedOverview"
Private Const COLOR_PENDING As Long = 15849925  ' Light blue (RGB 149, 179, 242) - same as Export_OverlayPending
Private Const STATUS_COL As Integer = 53  ' BA column for PENDING/DECLINED status

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
    ' Applies fixes from DeclinedOverview sheet and/or Kartei sheet
    ' Moves corrected records from decl_tblKartei to pre_tblKartei
    ' Supports two correction methods:
    '   1. Admin edits directly on DeclinedOverview (takes precedence)
    '   2. Admin edits directly on Kartei sheet
    
    On Error GoTo ErrorHandler
    
    ' Check if DeclinedOverview exists
    Dim wsOverview As Worksheet
    On Error Resume Next
    Set wsOverview = ThisWorkbook.Worksheets(DECLINED_SHEET_NAME)
    On Error GoTo ErrorHandler
    
    If wsOverview Is Nothing Then
        MsgBox "DeclinedOverview sheet not found. Please run 'Show Declined Overview' first.", _
               vbExclamation, "Sheet Not Found"
        Exit Sub
    End If
    
    ' Confirm action
    Dim response As VbMsgBoxResult
    response = MsgBox("This will move selected/all declined records to pending approval (pre_tblKartei)." & vbCrLf & _
                      "Records will be removed from declined status." & vbCrLf & vbCrLf & _
                      "Corrections will be taken from DeclinedOverview if changed there," & vbCrLf & _
                      "otherwise from corresponding rows on Kartei sheet." & vbCrLf & vbCrLf & _
                      "Continue?", vbYesNo + vbQuestion, "Confirm Apply Fixes")
    
    If response <> vbYes Then
        Exit Sub
    End If
    
    Application.ScreenUpdating = False
    Application.Calculation = xlCalculationManual
    
    ' Collect IDs to process
    Dim idsToProcess As Collection
    Set idsToProcess = CollectIDsFromOverview(wsOverview)
    
    If idsToProcess.Count = 0 Then
        MsgBox "No valid IDs found on DeclinedOverview sheet.", vbExclamation, "No Data"
        GoTo Cleanup
    End If
    
    ' Process the fixes
    Dim movedCount As Long
    Dim skippedCount As Long
    movedCount = ProcessDeclinedFixes(wsOverview, idsToProcess, skippedCount)
    
    ' Show results
    Dim msg As String
    If movedCount > 0 Then
        msg = "Successfully moved " & movedCount & " record(s) from declined to pending approval." & vbCrLf & _
              "Records are now in pre_tblKartei awaiting Superadmin review."
        
        If skippedCount > 0 Then
            msg = msg & vbCrLf & vbCrLf & "Skipped " & skippedCount & " record(s) (not found on Kartei sheet)."
        End If
        
        msg = msg & vbCrLf & vbCrLf & "DeclinedOverview sheet will be deleted."
        
        MsgBox msg, vbInformation, "Fixes Applied"
    Else
        msg = "No records were processed."
        If skippedCount > 0 Then
            msg = msg & vbCrLf & "Skipped: " & skippedCount
        End If
        msg = msg & vbCrLf & vbCrLf & "DeclinedOverview sheet will be deleted."
        
        MsgBox msg, vbInformation, "Apply Fixes"
    End If
    
    ' Delete DeclinedOverview sheet after processing (successful or not)
    Application.DisplayAlerts = False
    wsOverview.Delete
    Application.DisplayAlerts = True
    
Cleanup:
    Application.Calculation = xlCalculationAutomatic
    Application.ScreenUpdating = True
    Exit Sub
    
ErrorHandler:
    Application.Calculation = xlCalculationAutomatic
    Application.ScreenUpdating = True
    MsgBox "Error applying fixes: " & Err.Description, vbCritical, "Error"
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
        recData(1) = Export_DeclinedHelpers.NzStr(rs.Fields("Value1").value)
        recData(2) = Export_DeclinedHelpers.NzStr(rs.Fields("Value2").value)
        recData(3) = Export_DeclinedHelpers.NzStr(rs.Fields("Value3").value)
        recData(4) = Export_DeclinedHelpers.NzStr(rs.Fields("Value6").value)   ' Address
        recData(5) = Export_DeclinedHelpers.NzStr(rs.Fields("Value10").value)  ' Subject1
        recData(6) = Export_DeclinedHelpers.NzStr(rs.Fields("Value15").value)  ' Subject2
        
        ' Load months U-AF (columns 21-32 = Value21-Value32)
        Dim m As Integer
        For m = 1 To 12
            recData(6 + m) = Export_DeclinedHelpers.NzStr(rs.Fields("Value" & (20 + m)).value)
        Next m
        
        recData(19) = Export_DeclinedHelpers.NzStr(rs.Fields("Value52").value) ' History
        
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

Private Function ExtractLastDeclComment(ByVal historyText As String) As String
    ' Extracts the text from the last Decl_n: Was(); Is(...). block in history string
    ' Returns only the Is(...) content from the highest Decl_n marker
    ' If no Decl_ marker found, returns empty string
    
    If Len(Trim(historyText)) = 0 Then
        ExtractLastDeclComment = ""
        Exit Function
    End If
    
    ' Use regex to find all Decl_N: blocks with their numbers
    Dim regex As Object
    Set regex = CreateObject("VBScript.RegExp")
    
    With regex
        .Global = True
        .IgnoreCase = True
        ' Pattern: Decl_(\d+):\s*Was\([^)]*\);\s*Is\(([^)]*)\)
        ' Captures: group 1 = N (number), group 2 = content inside Is(...)
        .Pattern = "Decl_(\d+):\s*Was\([^)]*\);\s*Is\(([^)]*)\)"
    End With
    
    Dim matches As Object
    Set matches = regex.Execute(historyText)
    
    If matches.count = 0 Then
        ' No Decl_ markers found, return empty
        ExtractLastDeclComment = ""
        Exit Function
    End If
    
    ' Find the match with the highest N
    Dim maxN As Long
    maxN = -1
    Dim maxContent As String
    maxContent = ""
    
    Dim match As Object
    For Each match In matches
        Dim currentN As Long
        currentN = CLng(match.submatches(0))  ' First capture group = N
        
        If currentN > maxN Then
            maxN = currentN
            maxContent = match.submatches(1)  ' Second capture group = Is(...) content
        End If
    Next match
    
    ExtractLastDeclComment = Trim(maxContent)
End Function

' ========================================
' Sheet Management Functions
' ========================================

Private Function GetOrCreateDeclinedSheet() As Worksheet
    ' Gets existing DeclinedOverview sheet or creates new one
    ' Positions sheet directly after Kartei for easy navigation
    
    Dim ws As Worksheet
    On Error Resume Next
    Set ws = ThisWorkbook.Worksheets(DECLINED_SHEET_NAME)
    On Error GoTo 0
    
    Dim wsKartei As Worksheet
    Set wsKartei = ThisWorkbook.Worksheets("Kartei")
    
    If ws Is Nothing Then
        ' Create new sheet directly after Kartei
        Set ws = ThisWorkbook.Worksheets.Add(After:=wsKartei)
        ws.Name = DECLINED_SHEET_NAME
    Else
        ' Clear existing content
        ws.Cells.Clear
        
        ' Move to position after Kartei if not already there
        If ws.Index <> wsKartei.Index + 1 Then
            ws.Move After:=wsKartei
        End If
    End If
    
    Set GetOrCreateDeclinedSheet = ws
End Function

Private Sub BuildDeclinedOverview(ByVal ws As Worksheet, ByVal records As Collection)
    ' Builds the overview table on the worksheet
    ' Headers and column widths are copied from Kartei for consistency
    
    Dim wsKartei As Worksheet
    Set wsKartei = ThisWorkbook.Worksheets("Kartei")
    
    ' Copy headers from Kartei for relevant columns
    ' Map: DeclinedOverview col -> Kartei col
    ' A (1) -> AV (48) = ID
    ' B (2) -> A (1)
    ' C (3) -> B (2)
    ' D (4) -> C (3)
    ' E (5) -> F (6) = Address
    ' F (6) -> J (10) = Subject1
    ' G (7) -> O (15) = Subject2
    ' H-S (8-19) -> U-AF (21-32) = Months
    ' T (20) = Declined Times (custom)
    ' U (21) -> AZ (52) = History (last Decl_n only)
    
    With ws
        .Range("A1").value = wsKartei.Cells(2, 48).value  ' ID from AV
        .Range("B1").value = wsKartei.Cells(2, 1).value   ' A
        .Range("C1").value = wsKartei.Cells(2, 2).value   ' B
        .Range("D1").value = wsKartei.Cells(2, 3).value   ' C
        .Range("E1").value = wsKartei.Cells(2, 6).value   ' F (Address)
        .Range("F1").value = wsKartei.Cells(2, 10).value  ' J (Subject1)
        .Range("G1").value = wsKartei.Cells(2, 15).value  ' O (Subject2)
        
        ' Copy month headers (U-AF = columns 21-32)
        Dim m As Integer
        For m = 1 To 12
            .Cells(1, 7 + m).value = wsKartei.Cells(2, 20 + m).value
        Next m
        
        .Range("T1").value = "Declined Times"
        .Range("U1").value = "Last Decline Comment"  ' Changed from full history
        
        ' Format header row
        .Range("A1:U1").Font.Bold = True
        .Range("A1:U1").Interior.Color = RGB(200, 200, 200)
        .Range("A1:U1").HorizontalAlignment = xlCenter
        
        ' Copy column widths from Kartei
        .Columns(1).ColumnWidth = wsKartei.Columns(48).ColumnWidth  ' ID
        .Columns(2).ColumnWidth = wsKartei.Columns(1).ColumnWidth   ' A
        .Columns(3).ColumnWidth = wsKartei.Columns(2).ColumnWidth   ' B
        .Columns(4).ColumnWidth = wsKartei.Columns(3).ColumnWidth   ' C
        .Columns(5).ColumnWidth = wsKartei.Columns(6).ColumnWidth   ' F
        .Columns(6).ColumnWidth = wsKartei.Columns(10).ColumnWidth  ' J
        .Columns(7).ColumnWidth = wsKartei.Columns(15).ColumnWidth  ' O
        
        ' Copy month column widths (U-AF)
        For m = 1 To 12
            .Columns(7 + m).ColumnWidth = wsKartei.Columns(20 + m).ColumnWidth
        Next m
        
        ' Set width for custom columns
        .Columns(20).ColumnWidth = 12  ' Declined Times
        .Columns(21).ColumnWidth = wsKartei.Columns(52).ColumnWidth  ' History (AZ)
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
            For m = 1 To 12
                .Cells(rowNum, 7 + m).value = rec(6 + m)
            Next m
            
            .Cells(rowNum, 20).value = rec(20)  ' DeclCount
            
            ' Extract and show only last Decl_n comment in column U
            Dim lastDeclComment As String
            lastDeclComment = ExtractLastDeclComment(CStr(rec(19)))
            .Cells(rowNum, 21).value = lastDeclComment
            
            ' Highlight rows with multiple declinations
            If CInt(rec(20)) > 1 Then
                .Range(.Cells(rowNum, 1), .Cells(rowNum, 21)).Interior.Color = RGB(255, 200, 200)
            End If
        End With
        
        rowNum = rowNum + 1
    Next rec
    
    ' Apply AutoFilter to the data range
    If rowNum > 2 Then  ' Only if we have data rows
        Dim filterRange As Range
        Set filterRange = ws.Range("A1:U" & (rowNum - 1))
        filterRange.AutoFilter
    End If
    
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

Private Function ProcessDeclinedFixes(ByVal wsOverview As Worksheet, _
                                     ByVal idsToProcess As Collection, _
                                     ByRef skippedCount As Long) As Long
    ' Processes fixes for declined records
    ' Priority logic:
    '   1. If DeclinedOverview row differs from decl_tblKartei -> use DeclinedOverview data and copy to Kartei first
    '   2. Otherwise if Kartei row differs from decl_tblKartei -> use Kartei data
    ' Writes corrected data to pre_tblKartei, deletes from decl_tblKartei
    ' Sets BA="PENDING" and colors column A light blue on Kartei
    
    Dim movedCount As Long
    movedCount = 0
    skippedCount = 0
    
    Dim engine As DAO.DBEngine
    Set engine = New DAO.DBEngine
    
    Dim wsDao As DAO.Workspace
    Set wsDao = engine.Workspaces(0)
    
    Dim dbPath As String
    dbPath = ThisWorkbook.Worksheets("Kartei").Range("X1").value & "\Alarm\KindElternDaten_25_front.accdb"
    
    Dim db As DAO.Database
    Set db = wsDao.OpenDatabase(dbPath)
    
    ' Ensure pre_tblKartei exists
    Call Export_DeclinedHelpers.EnsurePreTableExists(db)
    
    wsDao.BeginTrans
    
    On Error GoTo RollbackTrans
    
    Dim wsKartei As Worksheet
    Set wsKartei = ThisWorkbook.Worksheets("Kartei")
    
    ' Build a dictionary of DeclinedOverview data for quick lookup
    Dim overviewDict As Object
    Set overviewDict = CreateObject("Scripting.Dictionary")
    Call Export_DeclinedHelpers.BuildOverviewDictionary(wsOverview, overviewDict)
    
    ' Build a dictionary of decl_tblKartei data for comparison
    Dim declDict As Object
    Set declDict = CreateObject("Scripting.Dictionary")
    Call Export_DeclinedHelpers.BuildDeclDictionary(db, declDict)
    
    Dim idToProcess As Variant
    For Each idToProcess In idsToProcess
        Dim targetID As Long
        targetID = CLng(idToProcess)
        
        ' Find row on Kartei sheet by ID (column AV = 48)
        Dim karteiRow As Long
        karteiRow = FindRowByIDInKartei(wsKartei, targetID)
        
        If karteiRow = 0 Then
            ' ID not found on Kartei - skip with warning
            skippedCount = skippedCount + 1
            Debug.Print "Warning: ID " & targetID & " not found on Kartei sheet, skipping."
            GoTo NextID
        End If
        
        ' Determine source of corrected data
        Dim useOverview As Boolean
        useOverview = False
        
        If overviewDict.exists(targetID) And declDict.exists(targetID) Then
            ' Check if DeclinedOverview differs from decl_tblKartei
            If Export_DeclinedHelpers.RecordsAreDifferent(overviewDict(targetID), declDict(targetID)) Then
                useOverview = True
            End If
        End If
        
        ' If using overview data, copy it to Kartei first
        If useOverview Then
            Call Export_DeclinedHelpers.CopyOverviewDataToKartei(wsKartei, karteiRow, overviewDict(targetID))
        End If
        
        ' Read current data from Kartei sheet (either updated from overview or original)
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
        
        ' Set PENDING status on Kartei sheet
        wsKartei.Cells(karteiRow, STATUS_COL).value = "PENDING"
        
        ' Color column A light blue if D <> "Zahlung"
        Dim cellD As String
        cellD = Trim(CStr(wsKartei.Cells(karteiRow, 4).value))
        
        If cellD <> "Zahlung" Then
            wsKartei.Cells(karteiRow, 1).Interior.Color = COLOR_PENDING
        End If
        
        movedCount = movedCount + 1
        
NextID:
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

