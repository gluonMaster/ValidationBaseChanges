Attribute VB_Name = "valid_ImportPending"
'==========================
'   Module: valid_ImportPending
'   Purpose: Load pending changes from pre_tblKartei into Superadmin Kartei sheet
'   Uses ID column (AV, column 48) as the primary key
'==========================

Option Explicit

' Load all pending changes from pre_tblKartei into Kartei sheet
Public Sub LoadPendingChangesFromPre()
    On Error GoTo ErrorHandler
    
    Application.ScreenUpdating = False
    Application.Calculation = xlCalculationManual
    
    Dim wsKartei As Worksheet
    Set wsKartei = GetOrCreateKarteiSheet()
    
    ' Clear existing data (keep headers in rows 1-2)
    ClearKarteiData wsKartei
    
    ' Load data from pre_tblKartei
    Dim dictPending As Scripting.Dictionary
    Set dictPending = ReadPreTableIntoDictionary()
    
    If dictPending.Count = 0 Then
        MsgBox "No pending changes found in pre_tblKartei.", vbInformation, "Load Pending"
        GoTo Cleanup
    End If
    
    ' Write pending data to Kartei sheet
    WritePendingToKartei wsKartei, dictPending
    
    ' Optionally load original values from tblKartei for comparison
    LoadOriginalValues wsKartei, dictPending
    
    MsgBox "Successfully loaded " & dictPending.Count & " pending change(s) from pre_tblKartei.", vbInformation, "Load Pending"
    
Cleanup:
    Application.Calculation = xlCalculationAutomatic
    Application.ScreenUpdating = True
    Exit Sub
    
ErrorHandler:
    Application.Calculation = xlCalculationAutomatic
    Application.ScreenUpdating = True
    MsgBox "Error loading pending changes: " & Err.Description, vbCritical, "Load Error"
End Sub

' Read all records from pre_tblKartei into dictionary (key = ID)
Private Function ReadPreTableIntoDictionary() As Scripting.Dictionary
    Dim dict As New Scripting.Dictionary
    
    Dim dbPath As String
    dbPath = GetDatabasePath()
    
    ' Check if pre_tblKartei exists, create if needed
    If Not TableExists(dbPath, "pre_tblKartei") Then
        CreatePreTable dbPath
        Set ReadPreTableIntoDictionary = dict
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
    
    Set ReadPreTableIntoDictionary = dict
End Function

' Write pending records to Kartei sheet, starting from row 3
Private Sub WritePendingToKartei(ByVal ws As Worksheet, ByVal dictPending As Scripting.Dictionary)
    Dim currentRow As Long
    currentRow = 3 ' Start from row 3 (rows 1-2 are headers)
    
    Dim idKey As Variant
    For Each idKey In dictPending.Keys
        Dim arrRow As Variant
        arrRow = dictPending(idKey)
        
        ' Write columns A-AY (1-51)
        Dim c As Long
        For c = 1 To 51
            ws.Cells(currentRow, c).Value = arrRow(1, c)
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
Private Sub LoadOriginalValues(ByVal ws As Worksheet, ByVal dictPending As Scripting.Dictionary)
    ' This is a stub - implement if Superadmin needs to see original vs pending side-by-side
    ' For now, we'll just ensure the data is loaded
    ' In future iterations, this could populate columns to the right or a separate comparison sheet
End Sub

' Get or create Kartei worksheet in Superadmin file
Private Function GetOrCreateKarteiSheet() As Worksheet
    Dim ws As Worksheet
    On Error Resume Next
    Set ws = ThisWorkbook.Worksheets("Kartei")
    On Error GoTo 0
    
    If ws Is Nothing Then
        Set ws = ThisWorkbook.Worksheets.Add(Before:=ThisWorkbook.Worksheets(1))
        ws.Name = "Kartei"
        CreateKarteiHeaders ws
    End If
    
    Set GetOrCreateKarteiSheet = ws
End Function

' Clear data rows in Kartei (keep headers)
Private Sub ClearKarteiData(ByVal ws As Worksheet)
    Dim lastRow As Long
    lastRow = ws.Cells(ws.Rows.Count, 1).End(xlUp).Row
    
    If lastRow > 2 Then
        ws.Rows("3:" & lastRow).ClearContents
        ws.Rows("3:" & lastRow).Interior.Color = xlNone
    End If
End Sub

' Create header rows for Kartei sheet
Private Sub CreateKarteiHeaders(ByVal ws As Worksheet)
    ' Row 1: Basic info
    ws.Cells(1, 1).Value = "Superadmin Kartei - Pending Changes"
    ws.Cells(1, 1).Font.Bold = True
    ws.Cells(1, 1).Font.Size = 14
    
    ' Row 2: Column headers (A-AZ)
    ' You can customize these based on actual column meanings
    Dim colNames As Variant
    colNames = Array("ID", "Eltern", "?", "Kind", "?", "Address", "?", "?", "?", "Subject1", _
                     "?", "?", "?", "?", "Subject2", "?", "?", "?", "?", "?", _
                     "Jan", "Feb", "Mrz", "Apr", "Mai", "Jun", "Jul", "Aug", "Sep", "Okt", "Nov", "Dez", _
                     "?", "?", "?", "?", "?", "?", "?", "?", "?", "?", "?", "?", "?", "?", "?", "ID_Col", "?", "?", "?", "History")
    
    Dim i As Long
    For i = LBound(colNames) To UBound(colNames)
        ws.Cells(2, i + 1).Value = colNames(i)
        ws.Cells(2, i + 1).Font.Bold = True
    Next i
End Sub

' Get database path from Kartei!X1 or use default
Private Function GetDatabasePath() As String
    On Error Resume Next
    Dim basePath As String
    basePath = ThisWorkbook.Worksheets("Kartei").Range("X1").Value
    
    If basePath = "" Then
        ' Fallback: prompt user or use default
        basePath = ThisWorkbook.Path
    End If
    
    GetDatabasePath = basePath & "\Alarm\KindElternDaten_25_front.accdb"
    On Error GoTo 0
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
