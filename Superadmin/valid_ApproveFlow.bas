Attribute VB_Name = "valid_ApproveFlow"
'==========================
'   Module: valid_ApproveFlow
'   Purpose: Superadmin workflow for approving/declining pending changes
'   Operations: Load pending, review GrossGeschichte, approve/decline decisions
'==========================

Option Explicit

' Main workflow: Load pending changes and generate GrossGeschichte for review
Public Sub LoadPendingChanges()
    On Error GoTo ErrorHandler
    
    Application.ScreenUpdating = False
    Application.Calculation = xlCalculationManual
    
    ' Step 1: Load data from pre_tblKartei into Kartei sheet
    Call valid_ImportPending.LoadPendingChangesFromPre
    
    ' Step 2: Generate GrossGeschichte report for review
    Call grossGeschichte.GrossGeshichteMachen
    
    ' Step 3: Add decision column to GrossGeschichte if not exists
    Call PrepareGrossGeschichteForDecisions
    
    MsgBox "Pending changes loaded successfully." & vbCrLf & vbCrLf & _
           "Please review the GrossGeschichte sheet and mark each change as:" & vbCrLf & _
           "  - 'Approved' (column X) to accept and move to tblKartei" & vbCrLf & _
           "  - 'Declined' (column X) to reject and move to decl_tblKartei" & vbCrLf & _
           "  - For declined records, optionally enter comment in column Y" & vbCrLf & vbCrLf & _
           "Then run 'SyncDecisions' to process your decisions.", _
           vbInformation, "Load Pending Changes"
    
    Application.Calculation = xlCalculationAutomatic
    Application.ScreenUpdating = True
    Exit Sub
    
ErrorHandler:
    Application.Calculation = xlCalculationAutomatic
    Application.ScreenUpdating = True
    MsgBox "Error loading pending changes: " & Err.Description, vbCritical, "Load Error"
End Sub

' Prepare GrossGeschichte sheet with decision column (X) and decline comment column (Y)
Private Sub PrepareGrossGeschichteForDecisions()
    Dim ws As Worksheet
    On Error Resume Next
    Set ws = ThisWorkbook.Worksheets("GrossGeschichte")
    On Error GoTo 0
    
    If ws Is Nothing Then Exit Sub
    
    ' Add header in column X (Decision)
    ws.Range("X2").Value = "Decision"
    ws.Range("X2").Font.Bold = True
    ws.Range("X2").HorizontalAlignment = xlCenter
    ws.Range("X2").Interior.Color = RGB(255, 255, 153) ' Light yellow
    
    ' Add validation dropdown (Approved/Declined) to decision cells
    Dim lastRow As Long
    lastRow = ws.Cells(ws.Rows.Count, 1).End(xlUp).Row
    
    If lastRow > 2 Then
        ' Find all "Ist" rows (every 3rd row starting from row 3: 3, 6, 9, ...)
        Dim r As Long
        For r = 4 To lastRow Step 3 ' Ist rows are at positions 4, 7, 10, ...
            ' Add dropdown validation to column X
            With ws.Range("X" & r).Validation
                .Delete
                .Add Type:=xlValidateList, AlertStyle:=xlValidAlertStop, _
                     Formula1:="Approved,Declined"
                .IgnoreBlank = True
                .InCellDropdown = True
            End With
            
            ws.Range("X" & r).Interior.Color = RGB(255, 255, 204) ' Lighter yellow
            
            ' Prepare column Y for decline comments (no validation, just styling)
            ws.Range("Y" & r).Interior.Color = RGB(255, 230, 230) ' Light pink
        Next r
    End If
End Sub

' Process all decisions from GrossGeschichte sheet
Public Sub SyncDecisions()
    On Error GoTo ErrorHandler
    
    Application.ScreenUpdating = False
    Application.Calculation = xlCalculationManual
    
    Dim wsGross As Worksheet
    Set wsGross = ThisWorkbook.Worksheets("GrossGeschichte")
    
    Dim lastRow As Long
    lastRow = wsGross.Cells(wsGross.Rows.Count, 1).End(xlUp).Row
    
    Dim approvedIDs As Collection
    Dim declinedIDs As Collection
    Set approvedIDs = New Collection
    Set declinedIDs = New Collection
    
    ' Collect decisions by ID
    Dim r As Long
    For r = 4 To lastRow Step 3 ' Process "Ist" rows (4, 7, 10, ...)
        Dim decision As String
        decision = UCase(Trim(wsGross.Range("X" & r).Value))
        
        If decision = "APPROVED" Or decision = "DECLINED" Then
            Dim recordID As String
            recordID = CStr(wsGross.Range("A" & r).Value)
            
            If recordID <> "" Then
                ' Check for duplicate decisions
                If decision = "APPROVED" Then
                    If Not CollectionContains(approvedIDs, recordID) Then
                        approvedIDs.Add recordID
                    End If
                ElseIf decision = "DECLINED" Then
                    ' For declined, first check column Y for pre-entered comment
                    Dim declComment As String
                    declComment = Trim(wsGross.Range("Y" & r).Value)
                    
                    ' If column Y is empty, prompt for comment via InputBox
                    If declComment = "" Then
                        declComment = InputBox("Enter reason for declining ID " & recordID & ":" & vbCrLf & _
                                              "(You can also enter comments directly in column Y of GrossGeschichte)", _
                                              "Decline Comment", "Superadmin declined this change")
                    End If
                    
                    If declComment <> "" Then
                        If Not CollectionContains(declinedIDs, recordID) Then
                            declinedIDs.Add Array(recordID, declComment)
                        End If
                    Else
                        MsgBox "Skipping ID " & recordID & " - no comment provided.", vbExclamation
                    End If
                End If
            End If
        End If
    Next r
    
    If approvedIDs.Count = 0 And declinedIDs.Count = 0 Then
        MsgBox "No decisions found. Please mark records as 'Approved' or 'Declined' in column X.", _
               vbExclamation, "No Decisions"
        GoTo Cleanup
    End If
    
    ' Process approved records
    If approvedIDs.Count > 0 Then
        Call ProcessApprovedRecords(approvedIDs)
    End If
    
    ' Process declined records
    If declinedIDs.Count > 0 Then
        Call ProcessDeclinedRecords(declinedIDs)
    End If
    
    MsgBox "Decisions processed successfully:" & vbCrLf & _
           "  Approved: " & approvedIDs.Count & vbCrLf & _
           "  Declined: " & declinedIDs.Count, _
           vbInformation, "Sync Complete"
    
Cleanup:
    Application.Calculation = xlCalculationAutomatic
    Application.ScreenUpdating = True
    Exit Sub
    
ErrorHandler:
    Application.Calculation = xlCalculationAutomatic
    Application.ScreenUpdating = True
    MsgBox "Error syncing decisions: " & Err.Description, vbCritical, "Sync Error"
End Sub

' Move approved records from pre_tblKartei to tblKartei
Private Sub ProcessApprovedRecords(ByVal approvedIDs As Collection)
    Dim dbPath As String
    dbPath = GetDatabasePath()
    
    Dim engine As DAO.DBEngine
    Set engine = New DAO.DBEngine
    
    Dim wsDao As DAO.Workspace
    Set wsDao = engine.Workspaces(0)
    
    Dim db As DAO.Database
    Set db = wsDao.OpenDatabase(dbPath)
    
    Dim fld As DAO.Field ' Declare once at procedure level
    
    wsDao.BeginTrans
    
    Dim idVar As Variant
    For Each idVar In approvedIDs
        Dim strID As String
        strID = CStr(idVar)
        
        ' Read record from pre_tblKartei
        Dim rsPre As DAO.Recordset
        Set rsPre = db.OpenRecordset("SELECT * FROM pre_tblKartei WHERE ID = " & strID, dbOpenDynaset)
        
        If Not rsPre.EOF Then
            ' Check if record exists in tblKartei
            Dim rsMain As DAO.Recordset
            Set rsMain = db.OpenRecordset("SELECT * FROM tblKartei WHERE ID = " & strID, dbOpenDynaset)
            
            If rsMain.EOF Then
                ' Record doesn't exist - add new
                rsMain.Close
                Set rsMain = db.OpenRecordset("tblKartei", dbOpenDynaset)
                rsMain.AddNew
                
                ' Copy all fields INCLUDING ID (to preserve ID from pre_tblKartei)
                For Each fld In rsPre.Fields
                    On Error Resume Next
                    rsMain.Fields(fld.Name).Value = rsPre.Fields(fld.Name).Value
                    On Error GoTo 0
                Next fld
            Else
                ' Record exists - update (skip ID since it already matches)
                rsMain.Edit
                
                For Each fld In rsPre.Fields
                    If fld.Name <> "ID" Then ' ID already matches, update other fields
                        On Error Resume Next
                        rsMain.Fields(fld.Name).Value = rsPre.Fields(fld.Name).Value
                        On Error GoTo 0
                    End If
                Next fld
            End If
            
            rsMain.Update
            rsMain.Close
            
            ' Delete from pre_tblKartei
            rsPre.Delete
        End If
        
        rsPre.Close
    Next idVar
    
    wsDao.CommitTrans
    db.Close
End Sub

' Move declined records from pre_tblKartei to decl_tblKartei
Private Sub ProcessDeclinedRecords(ByVal declinedIDs As Collection)
    Dim dbPath As String
    dbPath = GetDatabasePath()
    
    ' Ensure decl_tblKartei exists
    If Not TableExists(dbPath, "decl_tblKartei") Then
        Call CreateDeclTable(dbPath)
    End If
    
    Dim engine As DAO.DBEngine
    Set engine = New DAO.DBEngine
    
    Dim wsDao As DAO.Workspace
    Set wsDao = engine.Workspaces(0)
    
    Dim db As DAO.Database
    Set db = wsDao.OpenDatabase(dbPath)
    
    wsDao.BeginTrans
    
    Dim idVar As Variant
    For Each idVar In declinedIDs
        Dim arrData As Variant
        arrData = idVar
        
        Dim strID As String
        strID = CStr(arrData(0))
        
        Dim declComment As String
        declComment = CStr(arrData(1))
        
        ' Read record from pre_tblKartei
        Dim rsPre As DAO.Recordset
        Set rsPre = db.OpenRecordset("SELECT * FROM pre_tblKartei WHERE ID = " & strID, dbOpenDynaset)
        
        If Not rsPre.EOF Then
            ' Check if record exists in decl_tblKartei
            Dim rsDecl As DAO.Recordset
            Set rsDecl = db.OpenRecordset("SELECT * FROM decl_tblKartei WHERE ID = " & strID, dbOpenDynaset)
            
            If rsDecl.EOF Then
                ' Record doesn't exist - add new
                rsDecl.Close
                Set rsDecl = db.OpenRecordset("decl_tblKartei", dbOpenDynaset)
                rsDecl.AddNew
            Else
                ' Record exists - update
                rsDecl.Edit
            End If
            
            ' Copy all fields from pre_ to decl_ table, INCLUDING ID
            Dim fld As DAO.Field
            For Each fld In rsPre.Fields
                ' Copy ALL fields including ID (ID is regular Long, not AutoNumber)
                On Error Resume Next
                rsDecl.Fields(fld.Name).Value = rsPre.Fields(fld.Name).Value
                On Error GoTo 0
            Next fld
            
            ' Add decline comment to Value52 (history field)
            Dim currentHistory As String
            currentHistory = Nz(rsDecl.Fields("Value52").Value, "")
            
            ' Count existing Decl_N entries
            Dim declNum As Long
            declNum = CountDeclEntries(currentHistory) + 1
            
            ' Append decline comment in Was()/Is() format to match ParseSingleEvent regex
            ' Format: Decl_N: Was(); Is(<comment + timestamp>).
            Dim declEntry As String
            declEntry = "Decl_" & declNum & ": Was(); Is(" & declComment & " (Declined by Superadmin on " & Format(Date, "dd.mm.yyyy") & ")). || "
            
            rsDecl.Fields("Value52").Value = currentHistory & declEntry
            
            rsDecl.Update
            rsDecl.Close
            
            ' Delete from pre_tblKartei
            rsPre.Delete
        End If
        
        rsPre.Close
    Next idVar
    
    wsDao.CommitTrans
    db.Close
End Sub

' Count existing Decl_N entries in history string
Private Function CountDeclEntries(ByVal historyStr As String) As Long
    Dim count As Long
    count = 0
    
    Dim regex As Object
    Set regex = CreateObject("VBScript.RegExp")
    regex.Global = True
    regex.Pattern = "Decl_(\d+):"
    
    Dim matches As Object
    Set matches = regex.Execute(historyStr)
    
    count = matches.Count
    CountDeclEntries = count
End Function

' Helper: Check if collection contains value
Private Function CollectionContains(ByVal col As Collection, ByVal Value As String) As Boolean
    Dim item As Variant
    On Error Resume Next
    For Each item In col
        If IsArray(item) Then
            If CStr(item(0)) = Value Then
                CollectionContains = True
                Exit Function
            End If
        Else
            If CStr(item) = Value Then
                CollectionContains = True
                Exit Function
            End If
        End If
    Next item
    CollectionContains = False
End Function

' Helper: Get database path
Private Function GetDatabasePath() As String
    On Error Resume Next
    Dim basePath As String
    basePath = ThisWorkbook.Worksheets("Kartei").Range("X1").Value
    
    If basePath = "" Then
        basePath = ThisWorkbook.Path
    End If
    
    GetDatabasePath = basePath & "\Alarm\KindElternDaten_25_front.accdb"
    On Error GoTo 0
End Function

' Helper: Check if table exists
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

' Helper: Create decl_tblKartei with same structure as tblKartei
Private Sub CreateDeclTable(ByVal dbPath As String)
    Dim engine As DAO.DBEngine
    Set engine = New DAO.DBEngine
    
    Dim wsDao As DAO.Workspace
    Set wsDao = engine.Workspaces(0)
    
    Dim db As DAO.Database
    Set db = wsDao.OpenDatabase(dbPath)
    
    ' Create table with same structure as tblKartei
    Dim tbl As DAO.TableDef
    Set tbl = db.CreateTableDef("decl_tblKartei")
    
    ' Add ID field (regular Long, not AutoNumber - ID will be preserved from pre_tblKartei)
    Dim fld As DAO.Field
    Set fld = tbl.CreateField("ID", dbLong)
    ' NO fld.Attributes = dbAutoIncrField - ID is manually assigned
    tbl.Fields.Append fld
    
    ' Add Value1..Value51 fields (Text)
    Dim i As Long
    For i = 1 To 51
        Set fld = tbl.CreateField("Value" & i, dbText, 255)
        fld.AllowZeroLength = True ' Allow empty strings in text fields
        tbl.Fields.Append fld
    Next i
    
    ' Add Value52 (Memo, for history)
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

' Helper: Nz function (Null to Zero/Empty)
Private Function Nz(ByVal Value As Variant, Optional ByVal defaultValue As Variant = "") As Variant
    If IsNull(Value) Or IsEmpty(Value) Then
        Nz = defaultValue
    Else
        Nz = Value
    End If
End Function
