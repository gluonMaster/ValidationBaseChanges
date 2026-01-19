Attribute VB_Name = "Export_RiskClassification"
'==========================
'   Risk Classification Module
'   Handles risky change detection and pre_tblKartei operations
'==========================
Option Explicit

' ========================================
' Risk Policy Constants
' ========================================

' Policy modes:
'   "TRACKED_FIELDS" - Changes in tracked (history-recorded) fields are risky, others are safe
'   "ALL_RISKY"      - All changes to existing records are risky
'   "GRANULAR"       - Original SEPA/past months/discipline logic
Private Const RISK_POLICY_MODE As String = "TRACKED_FIELDS"

' ========================================
' Tracked (History-Recorded) Columns
' These are the columns that get recorded in history (AZ column)
' Changes to these fields WILL appear on GrossGeschichte
' Changes to other fields will NOT appear on GrossGeschichte
' ========================================
Private Const TRACKED_COL_FAMILY_ID As Integer = 1       ' A - FID
Private Const TRACKED_COL_PARENT As Integer = 2          ' B - PAR
Private Const TRACKED_COL_CHILD As Integer = 4           ' D - CHD
Private Const TRACKED_COL_BIRTHDATE As Integer = 5       ' E - DOB
Private Const TRACKED_COL_ADDRESS As Integer = 6         ' F - ADR
Private Const TRACKED_COL_PHONE As Integer = 7           ' G - TEL
Private Const TRACKED_COL_MOBILE As Integer = 8          ' H - MOB
Private Const TRACKED_COL_EMAIL As Integer = 9           ' I - EML
Private Const TRACKED_COL_SUBJECT1 As Integer = 10       ' J - SB1
Private Const TRACKED_COL_PRICE1 As Integer = 13         ' M - PR1
Private Const TRACKED_COL_SUBJECT2 As Integer = 15       ' O - SB2
Private Const TRACKED_COL_PRICE2 As Integer = 18         ' R - PR2
Private Const TRACKED_COL_MONTH_START As Integer = 21    ' U - M01 (Month 1)
Private Const TRACKED_COL_MONTH_END As Integer = 32      ' AF - M12 (Month 12)
Private Const TRACKED_COL_EXTRA1 As Integer = 37         ' AK - EX1
Private Const TRACKED_COL_EXTRA2 As Integer = 38         ' AL - EX2
Private Const TRACKED_COL_EXTRA3 As Integer = 39         ' AM - EX3

' ========================================
' Risk Classification Functions
' ========================================

Public Function IsRiskyChange(ByVal wsKartei As Worksheet, _
                              ByVal wsOriginal As Worksheet, _
                              ByVal rowIndex As Long, _
                              ByVal strID As String, _
                              ByVal arrLocal As Variant, _
                              ByVal arrOriginal As Variant) As Boolean
    ' Determines if a change is risky based on current policy
    ' Returns True if the change should go to pre_tblKartei instead of tblKartei
    ' NOTE: This function is only called for EXISTING records (ID exists in dictOriginal/tblKartei)
    '       New records are always considered safe and go directly to tblKartei
    
    Select Case RISK_POLICY_MODE
        Case "TRACKED_FIELDS"
            ' Changes in tracked (history-recorded) fields are risky
            ' Changes in non-tracked fields are safe (they won't appear on GrossGeschichte anyway)
            IsRiskyChange = HasTrackedFieldChanges(arrLocal, arrOriginal)
            
        Case "ALL_RISKY"
            ' All changes to existing records are risky
            IsRiskyChange = True
            
        Case "GRANULAR"
            ' Original SEPA/past months/discipline logic
            IsRiskyChange = IsRiskyChange_Granular(wsKartei, wsOriginal, rowIndex, strID, arrLocal, arrOriginal)
            
        Case Else
            ' Default to tracked fields policy
            IsRiskyChange = HasTrackedFieldChanges(arrLocal, arrOriginal)
    End Select
End Function

' ========================================
' Tracked Fields Policy Implementation
' ========================================

Public Function HasTrackedFieldChanges(ByVal arrLocal As Variant, _
                                       ByVal arrOriginal As Variant) As Boolean
    ' Checks if any tracked (history-recorded) field has changed
    ' Returns True if at least one tracked field differs between local and original
    ' Returns False if only non-tracked fields changed (those won't appear on GrossGeschichte)
    
    HasTrackedFieldChanges = False
    
    ' Check individual tracked columns
    If HasColumnChanged(arrLocal, arrOriginal, TRACKED_COL_FAMILY_ID) Then
        HasTrackedFieldChanges = True
        Exit Function
    End If
    
    If HasColumnChanged(arrLocal, arrOriginal, TRACKED_COL_PARENT) Then
        HasTrackedFieldChanges = True
        Exit Function
    End If
    
    If HasColumnChanged(arrLocal, arrOriginal, TRACKED_COL_CHILD) Then
        HasTrackedFieldChanges = True
        Exit Function
    End If
    
    If HasColumnChanged(arrLocal, arrOriginal, TRACKED_COL_BIRTHDATE) Then
        HasTrackedFieldChanges = True
        Exit Function
    End If
    
    If HasColumnChanged(arrLocal, arrOriginal, TRACKED_COL_ADDRESS) Then
        HasTrackedFieldChanges = True
        Exit Function
    End If
    
    If HasColumnChanged(arrLocal, arrOriginal, TRACKED_COL_PHONE) Then
        HasTrackedFieldChanges = True
        Exit Function
    End If
    
    If HasColumnChanged(arrLocal, arrOriginal, TRACKED_COL_MOBILE) Then
        HasTrackedFieldChanges = True
        Exit Function
    End If
    
    If HasColumnChanged(arrLocal, arrOriginal, TRACKED_COL_EMAIL) Then
        HasTrackedFieldChanges = True
        Exit Function
    End If
    
    If HasColumnChanged(arrLocal, arrOriginal, TRACKED_COL_SUBJECT1) Then
        HasTrackedFieldChanges = True
        Exit Function
    End If
    
    If HasColumnChanged(arrLocal, arrOriginal, TRACKED_COL_PRICE1) Then
        HasTrackedFieldChanges = True
        Exit Function
    End If
    
    If HasColumnChanged(arrLocal, arrOriginal, TRACKED_COL_SUBJECT2) Then
        HasTrackedFieldChanges = True
        Exit Function
    End If
    
    If HasColumnChanged(arrLocal, arrOriginal, TRACKED_COL_PRICE2) Then
        HasTrackedFieldChanges = True
        Exit Function
    End If
    
    ' Check months (columns 21-32)
    Dim col As Integer
    For col = TRACKED_COL_MONTH_START To TRACKED_COL_MONTH_END
        If HasColumnChanged(arrLocal, arrOriginal, col) Then
            HasTrackedFieldChanges = True
            Exit Function
        End If
    Next col
    
    ' Check extra subjects
    If HasColumnChanged(arrLocal, arrOriginal, TRACKED_COL_EXTRA1) Then
        HasTrackedFieldChanges = True
        Exit Function
    End If
    
    If HasColumnChanged(arrLocal, arrOriginal, TRACKED_COL_EXTRA2) Then
        HasTrackedFieldChanges = True
        Exit Function
    End If
    
    If HasColumnChanged(arrLocal, arrOriginal, TRACKED_COL_EXTRA3) Then
        HasTrackedFieldChanges = True
        Exit Function
    End If
    
    ' No tracked field changes found
    HasTrackedFieldChanges = False
End Function

Private Function HasColumnChanged(ByVal arrLocal As Variant, _
                                  ByVal arrOriginal As Variant, _
                                  ByVal colIndex As Integer) As Boolean
    ' Safely compares a single column between local and original arrays
    ' Returns True if values differ
    
    On Error Resume Next
    
    Dim localVal As String
    Dim origVal As String
    
    If IsError(arrLocal(1, colIndex)) Then
        localVal = ""
    Else
        localVal = CStr(arrLocal(1, colIndex))
    End If
    
    If IsError(arrOriginal(1, colIndex)) Then
        origVal = ""
    Else
        origVal = CStr(arrOriginal(1, colIndex))
    End If
    
    On Error GoTo 0
    
    HasColumnChanged = (localVal <> origVal)
End Function

Private Function IsRiskyChange_Granular(ByVal wsKartei As Worksheet, _
                                        ByVal wsOriginal As Worksheet, _
                                        ByVal rowIndex As Long, _
                                        ByVal strID As String, _
                                        ByVal arrLocal As Variant, _
                                        ByVal arrOriginal As Variant) As Boolean
    ' Original granular risk classification logic
    ' Preserved for easy rollback by changing USE_STRICT_ALL_RISKY_POLICY to False
    
    Dim sepaMarker As String
    sepaMarker = Trim(CStr(arrLocal(1, 47)))  ' Column AU (47)
    
    ' Get data year and calculate past months
    Dim dataYear As Long
    dataYear = GetDataYear()
    
    Dim currentYear As Long
    Dim currentMonth As Long
    currentYear = Year(Date)
    currentMonth = Month(Date)
    
    ' Check if any months U-AF (21-32) were changed
    Dim hasMonthChanges As Boolean
    Dim changedMonths() As Boolean
    ReDim changedMonths(1 To 12)
    hasMonthChanges = False
    
    Dim i As Integer
    For i = 21 To 32
        If CStr(arrLocal(1, i)) <> CStr(arrOriginal(1, i)) Then
            hasMonthChanges = True
            changedMonths(i - 20) = True
        End If
    Next i
    
    If Not hasMonthChanges Then
        ' No month changes, not risky
        IsRiskyChange_Granular = False
        Exit Function
    End If
    
    ' ========================================
    ' Scenario B (Strict): SEPA + any month changes
    ' ========================================
    If sepaMarker = "SEPA" Then
        IsRiskyChange_Granular = True
        Exit Function
    End If
    
    ' ========================================
    ' Scenario A (Non-strict): Non-SEPA + past month changes + no NH/Nachhilfe/Ind./VSpE
    ' ========================================
    
    ' Determine which months are "past"
    Dim isPastMonth() As Boolean
    ReDim isPastMonth(1 To 12)
    
    If currentYear > dataYear Then
        ' All months are past
        Dim m As Integer
        For m = 1 To 12
            isPastMonth(m) = True
        Next m
    ElseIf currentYear = dataYear Then
        ' Months < currentMonth are past
        For m = 1 To 12
            isPastMonth(m) = (m < currentMonth)
        Next m
    Else
        ' currentYear < dataYear: no past months
        For m = 1 To 12
            isPastMonth(m) = False
        Next m
    End If
    
    ' Check if any changed months are past
    Dim hasPastMonthChanges As Boolean
    hasPastMonthChanges = False
    
    For m = 1 To 12
        If changedMonths(m) And isPastMonth(m) Then
            hasPastMonthChanges = True
            Exit For
        End If
    Next m
    
    If Not hasPastMonthChanges Then
        ' Only future/current months changed, not risky
        IsRiskyChange_Granular = False
        Exit Function
    End If
    
    ' ========================================
    ' Depth-based policy: changes more than 1 month in the past are always risky
    ' This applies even when disciplines contain NH/Nachhilfe/Ind./VSpE keywords
    ' ========================================
    Dim maxPastDepth As Long
    maxPastDepth = 0
    
    If currentYear > dataYear Then
        ' Editing previous year's data - always considered deep (>1 month)
        maxPastDepth = 999
    ElseIf currentYear = dataYear Then
        ' Calculate actual depth for each changed past month
        For m = 1 To 12
            If changedMonths(m) And isPastMonth(m) Then
                Dim depth As Long
                depth = currentMonth - m
                If depth > maxPastDepth Then
                    maxPastDepth = depth
                End If
            End If
        Next m
    End If
    
    ' If editing more than 1 month in the past, always risky
    If maxPastDepth > 1 Then
        IsRiskyChange_Granular = True
        Exit Function
    End If
    
    ' ========================================
    ' For edits within 1 month depth, check disciplines (J for months 1-6, O for months 7-12)
    ' ========================================
    Dim disciplineJ As String
    Dim disciplineO As String
    disciplineJ = CStr(arrLocal(1, 10))  ' Column J (10)
    disciplineO = CStr(arrLocal(1, 15))  ' Column O (15)
    
    ' Check which range of months was changed
    Dim needCheckJ As Boolean
    Dim needCheckO As Boolean
    needCheckJ = False
    needCheckO = False
    
    For m = 1 To 6
        If changedMonths(m) And isPastMonth(m) Then
            needCheckJ = True
            Exit For
        End If
    Next m
    
    For m = 7 To 12
        If changedMonths(m) And isPastMonth(m) Then
            needCheckO = True
            Exit For
        End If
    Next m
    
    ' Evaluate safety of each discipline independently
    Dim isSafeJ As Boolean
    Dim isSafeO As Boolean
    isSafeJ = True  ' Default safe if not checked
    isSafeO = True  ' Default safe if not checked
    
    If needCheckJ Then
        ' If months 1-6 changed in the past, check discipline J
        If HasSafeKeywords(disciplineJ) Then
            isSafeJ = True
        Else
            isSafeJ = False
        End If
    End If
    
    If needCheckO Then
        ' If months 7-12 changed in the past, check discipline O
        If HasSafeKeywords(disciplineO) Then
            isSafeO = True
        Else
            isSafeO = False
        End If
    End If
    
    ' Change is risky if ANY required discipline check failed
    IsRiskyChange_Granular = (Not isSafeJ) Or (Not isSafeO)
End Function

Private Function HasSafeKeywords(ByVal disciplineText As String) As Boolean
    ' Returns True if the discipline contains safe keywords (NH, Nachhilfe, Ind., VSpE)
    ' which make the change NOT risky
    
    ' Check for "NH" (case-sensitive)
    If InStr(1, disciplineText, "NH", vbBinaryCompare) > 0 Then
        HasSafeKeywords = True
        Exit Function
    End If
    
    ' Check for "Nachhilfe", "Ind.", "VSpE" (case-insensitive)
    If InStr(1, disciplineText, "Nachhilfe", vbTextCompare) > 0 Then
        HasSafeKeywords = True
        Exit Function
    End If
    
    If InStr(1, disciplineText, "Ind.", vbTextCompare) > 0 Then
        HasSafeKeywords = True
        Exit Function
    End If
    
    If InStr(1, disciplineText, "VSpE", vbTextCompare) > 0 Then
        HasSafeKeywords = True
        Exit Function
    End If
    
    HasSafeKeywords = False
End Function

' ========================================
' Pre-table Management Functions
' ========================================

Public Sub EnsurePreTableExists(ByVal db As DAO.Database)
    ' Creates pre_tblKartei if it doesn't exist, with same structure as tblKartei
    ' ID is a regular Long field (NOT AutoNumber) to match IDs from tblKartei
    
    On Error Resume Next
    Dim tdf As DAO.TableDef
    Set tdf = db.TableDefs("pre_tblKartei")
    
    If Err.Number = 0 Then
        ' Table exists
        On Error GoTo 0
        Exit Sub
    End If
    On Error GoTo 0
    
    ' Create table with same structure as tblKartei
    Set tdf = db.CreateTableDef("pre_tblKartei")
    
    ' 1) Add ID field as regular Long (NOT AutoNumber - we set it manually)
    Dim fldID As DAO.Field
    Set fldID = tdf.CreateField("ID", dbLong)
    ' Do NOT set: fldID.Attributes = dbAutoIncrField
    tdf.Fields.Append fldID
    
    ' 2) For each column 1..51, add ValueX (Text) and InteriorColorX (Long)
    Dim i As Long
    For i = 1 To 51
        Dim fldValue As DAO.Field
        Set fldValue = tdf.CreateField("Value" & i, dbText, 255)
        fldValue.Required = False
        fldValue.AllowZeroLength = True
        tdf.Fields.Append fldValue
        
        Dim fldInterior As DAO.Field
        Set fldInterior = tdf.CreateField("InteriorColor" & i, dbLong)
        fldInterior.Required = False
        tdf.Fields.Append fldInterior
    Next i
    
    ' 3) Add FontColor fields
    Dim fldFont3 As DAO.Field
    Set fldFont3 = tdf.CreateField("FontColor3", dbLong)
    fldFont3.Required = False
    tdf.Fields.Append fldFont3
    
    Dim fldFont18 As DAO.Field
    Set fldFont18 = tdf.CreateField("FontColor18", dbLong)
    fldFont18.Required = False
    tdf.Fields.Append fldFont18
    
    ' 4) Add Value52 (Memo for history)
    Set fldValue = tdf.CreateField("Value52", dbMemo)
    fldValue.Required = False
    fldValue.AllowZeroLength = True
    tdf.Fields.Append fldValue
    
    ' Append table
    db.TableDefs.Append tdf
    
    ' Create primary key index on ID
    db.Execute "CREATE INDEX PK_Pre_ID ON pre_tblKartei(ID) WITH PRIMARY"
End Sub

Public Sub EnsureDeclTableExists(ByVal db As DAO.Database)
    ' Creates decl_tblKartei if it doesn't exist, with same structure as tblKartei
    ' ID is a regular Long field (NOT AutoNumber) to match IDs from tblKartei
    
    On Error Resume Next
    Dim tdf As DAO.TableDef
    Set tdf = db.TableDefs("decl_tblKartei")
    
    If Err.Number = 0 Then
        ' Table exists
        On Error GoTo 0
        Exit Sub
    End If
    On Error GoTo 0
    
    ' Create table with same structure as tblKartei
    Set tdf = db.CreateTableDef("decl_tblKartei")
    
    ' 1) Add ID field as regular Long (NOT AutoNumber - we set it manually)
    Dim fldID As DAO.Field
    Set fldID = tdf.CreateField("ID", dbLong)
    ' Do NOT set: fldID.Attributes = dbAutoIncrField
    tdf.Fields.Append fldID
    
    ' 2) For each column 1..51, add ValueX (Text) and InteriorColorX (Long)
    Dim i As Long
    For i = 1 To 51
        Dim fldValue As DAO.Field
        Set fldValue = tdf.CreateField("Value" & i, dbText, 255)
        fldValue.Required = False
        fldValue.AllowZeroLength = True
        tdf.Fields.Append fldValue
        
        Dim fldInterior As DAO.Field
        Set fldInterior = tdf.CreateField("InteriorColor" & i, dbLong)
        fldInterior.Required = False
        tdf.Fields.Append fldInterior
    Next i
    
    ' 3) Add FontColor fields
    Dim fldFont3 As DAO.Field
    Set fldFont3 = tdf.CreateField("FontColor3", dbLong)
    fldFont3.Required = False
    tdf.Fields.Append fldFont3
    
    Dim fldFont18 As DAO.Field
    Set fldFont18 = tdf.CreateField("FontColor18", dbLong)
    fldFont18.Required = False
    tdf.Fields.Append fldFont18
    
    ' 4) Add Value52 (Memo for history)
    Set fldValue = tdf.CreateField("Value52", dbMemo)
    fldValue.Required = False
    fldValue.AllowZeroLength = True
    tdf.Fields.Append fldValue
    
    ' Append table
    db.TableDefs.Append tdf
    
    ' Create primary key index on ID
    db.Execute "CREATE INDEX PK_Decl_ID ON decl_tblKartei(ID) WITH PRIMARY"
End Sub

' ========================================
' Write Functions
' ========================================

Public Sub WriteRiskyChangesToPreTable(ByVal dictLocal As Scripting.Dictionary, _
                                       ByVal dictLocalFormats As Scripting.Dictionary, _
                                       ByVal riskyIDs As Collection)
    ' Writes risky changes to pre_tblKartei instead of tblKartei
    ' Uses same DAO pattern as WriteDictionaryChangesToAccess_Recordset
    ' NOTE: Risky changes always have valid existing IDs from AV/tblKartei
    
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
    
    Dim varID As Variant
    For Each varID In riskyIDs
        If dictLocal.exists(varID) Then
            Dim arrRow As Variant
            Dim arrFormats As Variant
            arrRow = dictLocal(varID)
            arrFormats = dictLocalFormats(varID)
            
            Dim strID As String
            strID = CStr(arrRow(1, 48))
            
            If IsNumeric(strID) And Val(strID) > 0 Then
                Dim targetID As Long
                targetID = CLng(strID)
                
                ' Check if record exists in pre_tblKartei
                Dim rsCheck As DAO.Recordset
                Set rsCheck = db.OpenRecordset("SELECT * FROM pre_tblKartei WHERE ID = " & targetID, dbOpenDynaset)
                
                If rsCheck.EOF Then
                    ' Record not found - create new with explicit ID
                    rsCheck.Close
                    
                    Dim rsNew As DAO.Recordset
                    Set rsNew = db.OpenRecordset("pre_tblKartei", dbOpenDynaset)
                    rsNew.AddNew
                    ' Explicitly set ID to match AV/tblKartei
                    rsNew.Fields("ID").value = targetID
                    Call FillPreRecordFromArray(rsNew, arrRow, arrFormats)
                    rsNew.Update
                    rsNew.Close
                Else
                    ' Record found - update it (ID already set, don't change it)
                    rsCheck.Edit
                    Call FillPreRecordFromArray(rsCheck, arrRow, arrFormats)
                    rsCheck.Update
                    rsCheck.Close
                End If
            Else
                ' Should not happen: risky changes should always have valid ID
                ' Keep as safety check
                Dim rsEmpty As DAO.Recordset
                Set rsEmpty = db.OpenRecordset("pre_tblKartei", dbOpenDynaset)
                rsEmpty.AddNew
                Call FillPreRecordFromArray(rsEmpty, arrRow, arrFormats)
                rsEmpty.Update
                rsEmpty.Close
            End If
        End If
    Next varID
    
    wsDao.CommitTrans
    db.Close
End Sub

Private Sub FillPreRecordFromArray(ByVal rs As DAO.Recordset, _
                                   ByVal arrRow As Variant, _
                                   ByVal arrFormats As Variant)
    ' Fills recordset fields from array data
    ' Same logic as FillRecordFromArray but for pre_tblKartei
    
    Dim c As Long
    For c = 1 To 51
        Dim fieldName As String
        fieldName = "Value" & c
        If Not IsError(arrRow(1, c)) Then
            ' Phone columns (7=Tel., 8=Handy): must be written as strings to preserve
            ' leading zeros and prevent scientific notation in Access
            ' Also normalize any existing scientific-notation strings (e.g. "1,76E+9" -> "176000000")
            If c = 7 Or c = 8 Then
                rs.Fields(fieldName).Value = phone_Normalize.NormalizePhoneText(arrRow(1, c))
            Else
                rs.Fields(fieldName).Value = arrRow(1, c)
            End If
        Else
            rs.Fields(fieldName).Value = ""
        End If
        
        fieldName = "InteriorColor" & c
        If IsNull(arrFormats(1, c)) Or IsEmpty(arrFormats(1, c)) Then
            rs.Fields(fieldName).Value = 0
        Else
            rs.Fields(fieldName).Value = arrFormats(1, c)
        End If
    Next c
    
    ' Value52 (history)
    If Not IsError(arrRow(1, 52)) Then
        rs.Fields("Value52").value = arrRow(1, 52)
    Else
        rs.Fields("Value52").value = ""
    End If
    
    ' FontColor fields
    If IsNull(arrFormats(1, 52)) Or IsEmpty(arrFormats(1, 52)) Then
        rs.Fields("FontColor3").value = 0
    Else
        rs.Fields("FontColor3").value = arrFormats(1, 52)
    End If
    
    If IsNull(arrFormats(1, 53)) Or IsEmpty(arrFormats(1, 53)) Then
        rs.Fields("FontColor18").value = 0
    Else
        rs.Fields("FontColor18").value = arrFormats(1, 53)
    End If
End Sub
