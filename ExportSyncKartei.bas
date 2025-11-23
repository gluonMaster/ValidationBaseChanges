Attribute VB_Name = "ExportSyncKartei"
'==========================
'   Code Section: modSyncAccess
'==========================
Option Explicit

Public Sub CompareAndSyncKartei_OnSave()
    Dim canProceed As Boolean
    canProceed = ValidateAndFixPastMonths()
    If Not canProceed Then
        ' user got message, cancel
        Exit Sub
    End If
    
    ' else do the real sync
    CompareAndSyncKartei False
End Sub

Public Sub CompareAndSyncKartei_OnBtn()
    Dim canProceed As Boolean
    canProceed = ValidateAndFixPastMonths()
    If Not canProceed Then
        ' user got message, cancel
        Exit Sub
    End If
    
    ' else do the real sync
    CompareAndSyncKartei True
End Sub

Public Sub CompareAndSyncKartei(Optional ByVal manualRun As Boolean = False)
    ' Main sync procedure, now using ID in column AV (48).
    ' 1) Read local Kartei & Kartei_Original
    ' 2) Read Access data into dictCloud
    ' 3) Determine changed or new rows by comparing local vs original
    ' 4) For each changed row, update AW,AX,AY in local
    ' 5) Write changes to Access by ID using DAO Recordset
    ' 6) Rebuild Kartei_Original

    Application.ScreenUpdating = False
    Application.Calculation = xlCalculationManual
    
    On Error GoTo Cleanup
    
    Dim wsLocal As Worksheet, wsOriginal As Worksheet
    Set wsLocal = ThisWorkbook.Worksheets("Kartei")
    Set wsOriginal = ThisWorkbook.Worksheets("Kartei_Original")
    
'    wsLocal.Unprotect password:="1212"
'    wsLocal.Cells.Locked = False
    
    ' Remove all filters and depict all hidden rows and columns
    Call ResetSheetView(wsLocal, wsOriginal)
    
    ' SEPA protection: Operator cannot edit SEPA rows
    Call ValidateOperatorSepaRestrictions(wsLocal, wsOriginal)
    
    ' Make Backup of Kartei sheet in the separate file
    Call BackupKarteiSheet
    Call CopyKindElternDatenFile
    
    ' Read local & original
    Dim dictLocal As Scripting.Dictionary
    Dim dictOriginal As Scripting.Dictionary
    Dim dictLocalFormats As Scripting.Dictionary
    
    ' Add IDs to the new records
    Call AddIDsToKartei
    
    Set dictLocal = ReadSheetIntoDictionary_ID(wsLocal, 3, 52)       ' Reads A..AY (1..51), ID in col 48
    Set dictLocalFormats = ReadLocalFormatsIntoDictionary_ID(wsLocal, 3, 51)
    Set dictOriginal = ReadSheetIntoDictionary_ID(wsOriginal, 3, 51)
    
    ' Read from Access
    Dim dictCloud As Scripting.Dictionary
    Set dictCloud = ReadAccessIntoDictionary_ID()
    
    ' Find changed/new IDs
    Dim changedIDs As Collection
    Set changedIDs = FindChangedIDs(dictLocal, dictOriginal)
    
    ' Categorize changed IDs by status: normal, pending, declined
    Dim normalChangedIDs As New Collection
    Dim pendingChangedIDs As New Collection
    Dim declinedChangedIDs As New Collection
    
    Debug.Print "Categorizing " & changedIDs.Count & " changed IDs by status..."
    
    Dim checkID As Variant
    For Each checkID In changedIDs
        Dim checkRow As Long
        checkRow = FindRowByID_Sync(wsLocal, CStr(checkID))
        If checkRow > 0 Then
            Dim rowStatus As String
            rowStatus = GetRowStatus(wsLocal, checkRow)
            
            ' DIAGNOSTIC: Read BA column directly
            Dim baValue As Variant
            baValue = wsLocal.Cells(checkRow, 53).value
            Dim baText As String
            If IsEmpty(baValue) Then
                baText = "(empty)"
            Else
                baText = "'" & CStr(baValue) & "'"
            End If
            
            Debug.Print "ID " & checkID & " (row " & checkRow & "): status = '" & rowStatus & "', BA(53) raw = " & baText
            
            If rowStatus = "PENDING" Then
                pendingChangedIDs.Add checkID
            ElseIf rowStatus = "DECLINED" Then
                declinedChangedIDs.Add checkID
            Else
                ' Normal record (no special status)
                normalChangedIDs.Add checkID
            End If
        Else
            ' Row not found, include as normal (new record case)
            Debug.Print "ID " & checkID & ": row not found, treating as normal (new)"
            normalChangedIDs.Add checkID
        End If
    Next checkID
    
    Debug.Print "Categorization complete: normal=" & normalChangedIDs.Count & ", pending=" & pendingChangedIDs.Count & ", declined=" & declinedChangedIDs.Count
    
    ' Classify changes into safe and risky
    Dim safeIDs As New Collection
    Dim riskyIDs As New Collection
    
    Dim changedID As Variant
    For Each changedID In normalChangedIDs
        Dim arrLocal As Variant
        arrLocal = dictLocal(changedID)
        
        If Not dictOriginal.exists(changedID) Then
            ' New record (ID not in Kartei_Original/tblKartei) - always safe
            safeIDs.Add changedID
        Else
            ' Existing record - check if risky
            Dim arrOriginal As Variant
            arrOriginal = dictOriginal(changedID)
            
            ' Determine if change is risky
            Dim isRisky As Boolean
            isRisky = IsRiskyChange(wsLocal, wsOriginal, 0, CStr(changedID), arrLocal, arrOriginal)
            
            If isRisky Then
                riskyIDs.Add changedID
            Else
                safeIDs.Add changedID
            End If
        End If
    Next changedID
    
    Dim maxIDOriginal As Long
    maxIDOriginal = MaxID(wsOriginal)
    
    ' Update AW,AX,AY (cols 49..51) in local for all changed records
    Dim userRole As String
    userRole = GetUserRole()  ' "Admin" / "Operator"
    
    ' Collections for IDs that successfully updated (user didn't cancel)
    Dim effectiveSafeIDs As New Collection
    Dim effectiveRiskyIDs As New Collection
    
    ' Process safe changes
    For Each changedID In safeIDs
        arrLocal = dictLocal(changedID)
        
        ' Update AW,AX,AY => userRole, Date, Time
        arrLocal(1, 49) = userRole
        arrLocal(1, 50) = Date
        arrLocal(1, 51) = Format(Time, "HH:MM")
        
        ' Overwrite local sheet row so AW,AX,AY are updated visually
        Dim historyUpdate As String
        historyUpdate = UpdateLocalSheetRowByID(wsLocal, wsOriginal, CStr(changedID), arrLocal, maxIDOriginal)
        
        ' Check if user canceled (empty string returned)
        If historyUpdate <> "" Then
            ' Update successful, save to dict and add to effective collection
            arrLocal(1, 52) = historyUpdate
            dictLocal(changedID) = arrLocal
            effectiveSafeIDs.Add changedID
        End If
        ' If canceled (empty string), skip this ID - don't update dict, don't sync to Access
    Next changedID
    
    ' Process risky changes (update history/Notitzen on sheet, but don't write to tblKartei)
    For Each changedID In riskyIDs
        arrLocal = dictLocal(changedID)
        
        ' Update AW,AX,AY => userRole, Date, Time
        arrLocal(1, 49) = userRole
        arrLocal(1, 50) = Date
        arrLocal(1, 51) = Format(Time, "HH:MM")
        
        ' Update local sheet row with history
        historyUpdate = UpdateLocalSheetRowByID(wsLocal, wsOriginal, CStr(changedID), arrLocal, maxIDOriginal)
        
        ' Check if user canceled
        If historyUpdate <> "" Then
            ' Update successful, save to dict and add to effective collection
            arrLocal(1, 52) = historyUpdate
            dictLocal(changedID) = arrLocal
            effectiveRiskyIDs.Add changedID
        End If
        ' If canceled, skip this ID
    Next changedID
    
    ' Write safe changes to Access tblKartei (only IDs that weren't canceled)
    If effectiveSafeIDs.Count > 0 Then
        WriteDictionaryChangesToAccess_Recordset dictLocal, dictLocalFormats, effectiveSafeIDs
    End If
    
    ' Write risky changes to pre_tblKartei (only IDs that weren't canceled)
    If effectiveRiskyIDs.Count > 0 Then
        WriteRiskyChangesToPreTable dictLocal, dictLocalFormats, effectiveRiskyIDs
        
        ' Mark risky records as PENDING on Kartei sheet immediately
        Call MarkRiskyRowsAsPending(wsLocal, effectiveRiskyIDs)
    End If
    
    ' ========================================
    ' Process PENDING changes
    ' ========================================
    ' Admin edits to existing PENDING records are updated in pre_tblKartei
    ' BA remains "PENDING", color stays pending blue
    Dim effectivePendingIDs As New Collection
    
    For Each changedID In pendingChangedIDs
        arrLocal = dictLocal(changedID)
        
        ' Update AW,AX,AY => userRole, Date, Time
        arrLocal(1, 49) = userRole
        arrLocal(1, 50) = Date
        arrLocal(1, 51) = Format(Time, "HH:MM")
        
        ' Update local sheet row with history
        historyUpdate = UpdateLocalSheetRowByID(wsLocal, wsOriginal, CStr(changedID), arrLocal, maxIDOriginal)
        
        ' Check if user canceled
        If historyUpdate <> "" Then
            ' Update successful, save to dict and add to effective collection
            arrLocal(1, 52) = historyUpdate
            dictLocal(changedID) = arrLocal
            effectivePendingIDs.Add changedID
        End If
        ' If canceled, skip this ID
    Next changedID
    
    ' Write pending changes to pre_tblKartei (update existing records)
    If effectivePendingIDs.Count > 0 Then
        WriteRiskyChangesToPreTable dictLocal, dictLocalFormats, effectivePendingIDs
    End If
    
    ' ========================================
    ' Process DECLINED changes
    ' ========================================
    ' Admin fixes to DECLINED records move them from decl_tblKartei to pre_tblKartei
    ' BA changes to "PENDING", color changes to pending blue
    Dim effectiveDeclinedIDs As New Collection
    
    Debug.Print "Processing DECLINED changes: declinedChangedIDs.Count = " & declinedChangedIDs.Count
    
    For Each changedID In declinedChangedIDs
        Debug.Print "Processing DECLINED ID: " & changedID
        arrLocal = dictLocal(changedID)
        
        ' Update AW,AX,AY => userRole, Date, Time
        arrLocal(1, 49) = userRole
        arrLocal(1, 50) = Date
        arrLocal(1, 51) = Format(Time, "HH:MM")
        
        ' Update local sheet row with history and Notitzen
        historyUpdate = UpdateLocalSheetRowByID(wsLocal, wsOriginal, CStr(changedID), arrLocal, maxIDOriginal)
        
        Debug.Print "UpdateLocalSheetRowByID returned: '" & historyUpdate & "' for ID " & changedID
        
        ' Check if user canceled
        If historyUpdate <> "" Then
            ' Update successful, save to dict and add to effective collection
            arrLocal(1, 52) = historyUpdate
            dictLocal(changedID) = arrLocal
            effectiveDeclinedIDs.Add changedID
            
            Debug.Print "Added ID " & changedID & " to effectiveDeclinedIDs"
            
            ' CRITICAL: Re-read formats from sheet after UpdateLocalSheetRowByID
            ' because formats may have changed during history updates
            Dim strID_Decl As String
            strID_Decl = CStr(changedID)
            Dim rowDecl As Long
            rowDecl = FindRowByID_Sync(wsLocal, strID_Decl)
            
            If rowDecl > 0 Then
                ' Re-read format data from sheet for this ID
                Dim updatedFormats() As Variant
                ReDim updatedFormats(1 To 1, 1 To 53)
                
                Dim colIdx As Integer
                For colIdx = 1 To 51
                    updatedFormats(1, colIdx) = wsLocal.Cells(rowDecl, colIdx).Interior.Color
                Next colIdx
                updatedFormats(1, 52) = wsLocal.Cells(rowDecl, 3).Font.Color
                updatedFormats(1, 53) = wsLocal.Cells(rowDecl, 18).Font.Color
                
                ' Update dictLocalFormats with fresh data
                dictLocalFormats(changedID) = updatedFormats
            End If
        End If
        ' If canceled, skip this ID
    Next changedID
    
    ' Move declined records from decl_tblKartei to pre_tblKartei
    If effectiveDeclinedIDs.Count > 0 Then
        Debug.Print "About to call MoveDeclinedToPending with " & effectiveDeclinedIDs.Count & " IDs"
        Call MoveDeclinedToPending(dictLocal, dictLocalFormats, effectiveDeclinedIDs, wsLocal)
        Debug.Print "Returned from MoveDeclinedToPending"
    Else
        Debug.Print "No declined IDs to process (effectiveDeclinedIDs.Count = 0)"
    End If
    
'    Dim lastRow As Long
'
'    lastRow = wsLocal.Cells(wsLocal.Rows.Count, 1).End(xlUp).row
'
'    Dim rng As Range
'    Set rng = wsLocal.Range("A2:AZ" & lastRow)
'
'    rng.Sort Key1:=rng.Columns(2), Order1:=xlAscending, Header:=xlYes

    Call SortNameZ

    ' Rebuild Kartei_Original
    Call RebuildKarteiOriginal
    
'    ' Call the initialization for protection
'    Call InitializeWorkbookPubl

Cleanup:
    Application.Calculation = xlCalculationAutomatic
    Application.ScreenUpdating = True
    
    If Err.Number <> 0 Then
        MsgBox "Error during synchronization: " & Err.Description, vbCritical
        Exit Sub
    End If
    
    ' Inform user about classification results
    Dim msgText As String
    Dim totalProcessed As Long
    totalProcessed = safeIDs.Count + riskyIDs.Count + effectivePendingIDs.Count + effectiveDeclinedIDs.Count
    
    If totalProcessed > 0 Then
        msgText = "Synchronization completed:" & vbCrLf & _
                  "Safe changes written to database: " & safeIDs.Count & vbCrLf & _
                  "Risky changes sent for approval: " & riskyIDs.Count
        
        If effectivePendingIDs.Count > 0 Then
            msgText = msgText & vbCrLf & "Pending records updated: " & effectivePendingIDs.Count
        End If
        
        If effectiveDeclinedIDs.Count > 0 Then
            msgText = msgText & vbCrLf & "Declined records fixed and moved to pending: " & effectiveDeclinedIDs.Count
        End If
        
        MsgBox msgText, vbInformation, "Sync Summary"
    End If
    
    If manualRun Then
        If totalProcessed = 0 Then
            MsgBox "Synchronization completed (ID-based).", vbInformation
        End If
    Else
        ThisWorkbook.Save
'        Dim newFileName As String
'        newFileName = "KED24_bkp.xlsm"
'        ThisWorkbook.SaveAs fileName:=ThisWorkbook.Path & "\" & newFileName, FileFormat:=xlOpenXMLWorkbookMacroEnabled
    End If
End Sub

' New InitializeWorkbook procedure
Public Sub InitializeWorkbookPubl()
    On Error Resume Next ' Error handling, e.g. if a file is opened as read-only

    Dim ws As Worksheet
    Set ws = ThisWorkbook.Sheets("Kartei")

    ' Checking if a file is opened as read-only
    If ThisWorkbook.ReadOnly Then
        Exit Sub
    End If

    ' Protect the sheet with password
'    ws.Protect password:="1212", _
'               AllowFormattingCells:=True, _
'               AllowFormattingColumns:=True, _
'               AllowFormattingRows:=True, _
'               AllowSorting:=True, _
'               AllowFiltering:=True, _
'               AllowUsingPivotTables:=True
    ' Call the procedure to lock past months
    'Call LockPastMonths(ws)
'    Call LockReportColPubl(ws)

    On Error GoTo 0
End Sub

Public Sub LockReportColPubl(ws As Worksheet)
    On Error Resume Next ' Handling of possible errors

    Application.ScreenUpdating = False
    Application.Calculation = xlCalculationManual
    
    ws.Unprotect password:="1212"
    ws.Cells.Locked = False
    
    ws.Columns("AZ").Locked = True

    ' Re-protect the sheet to apply changes
    ws.Protect password:="1212", _
               AllowFormattingCells:=True, _
               AllowFormattingColumns:=True, _
               AllowFormattingRows:=True, _
               AllowSorting:=True, _
               AllowFiltering:=True, _
               AllowUsingPivotTables:=True

    Application.Calculation = xlCalculationAutomatic
    Application.ScreenUpdating = True

    On Error GoTo 0
End Sub

Private Function FindRowByID_Sync(ByVal ws As Worksheet, ByVal strID As String) As Long
    ' Finds row by ID in column AV (48)
    Dim lastRow As Long
    lastRow = ws.Cells(ws.Rows.count, 1).End(xlUp).row
    
    Dim r As Long
    For r = 3 To lastRow
        If CStr(ws.Cells(r, 48).value) = strID Then
            FindRowByID_Sync = r
            Exit Function
        End If
    Next r
    
    FindRowByID_Sync = 0
End Function

Private Sub ValidateOperatorSepaRestrictions(ByVal wsLocal As Worksheet, ByVal wsOriginal As Worksheet)
    ' Validate that Operator has not modified SEPA rows (columns U-AF)
    ' If modifications found, revert them and notify the user
    
    If Not IsOperator() Then
        Exit Sub  ' Admin can edit anything
    End If
    
    Dim lastRow As Long
    Dim firstRow As Long
    firstRow = 3
    lastRow = wsLocal.Cells(wsLocal.Rows.count, 1).End(xlUp).row
    
    Dim row As Long
    Dim col As Integer
    Dim hasSepaViolations As Boolean
    hasSepaViolations = False
    
    ' Columns U-AF are 21-32
    Dim firstMonthCol As Integer
    Dim lastMonthCol As Integer
    firstMonthCol = 21  ' U
    lastMonthCol = 32   ' AF
    
    For row = firstRow To lastRow
        If IsSepaRow(wsLocal, row) Then
            ' Check if any values in U-AF differ from original
            Dim hasChanges As Boolean
            hasChanges = False
            
            For col = firstMonthCol To lastMonthCol
                Dim localVal As Variant
                Dim origVal As Variant
                localVal = wsLocal.Cells(row, col).value
                origVal = wsOriginal.Cells(row, col).value
                
                If CStr(localVal) <> CStr(origVal) Then
                    hasChanges = True
                    ' Revert to original value
                    wsLocal.Cells(row, col).value = origVal
                End If
            Next col
            
            If hasChanges Then
                hasSepaViolations = True
            End If
        End If
    Next row
    
    If hasSepaViolations Then
        MsgBox "Operator is not allowed to edit SEPA records. Changes have been reverted.", vbExclamation, "SEPA Protection"
    End If
End Sub

Private Sub MoveDeclinedToPending(ByVal dictLocal As Scripting.Dictionary, _
                                  ByVal dictLocalFormats As Scripting.Dictionary, _
                                  ByVal declinedIDs As Collection, _
                                  ByVal wsLocal As Worksheet)
    ' Moves declined records from decl_tblKartei to pre_tblKartei
    ' Updates status to PENDING and changes color on Kartei sheet
    ' Updates Kartei_Original to prevent repeated history requests
    
    ' DIAGNOSTIC: Show that function is called
    Debug.Print "MoveDeclinedToPending called with " & declinedIDs.Count & " IDs"
    
    Const STATUS_COL As Integer = 53  ' BA column
    Const COLOR_PENDING As Long = 15849925  ' Light blue
    
    Dim engine As DAO.DBEngine
    Set engine = New DAO.DBEngine
    
    Dim wsDao As DAO.Workspace
    Set wsDao = engine.Workspaces(0)
    
    Dim dbPath As String
    dbPath = wsLocal.Range("X1").value & "\Alarm\KindElternDaten_25_front.accdb"
    
    Dim db As DAO.Database
    Set db = wsDao.OpenDatabase(dbPath)
    
    ' Ensure pre_tblKartei exists
    Call EnsurePreTableExists(db)
    
    ' Get reference to Kartei_Original for updates
    Dim wsOriginal As Worksheet
    Set wsOriginal = ThisWorkbook.Worksheets("Kartei_Original")
    
    wsDao.BeginTrans
    
    On Error GoTo RollbackTrans
    
    Dim varID As Variant
    Dim processedCount As Long
    processedCount = 0
    
    For Each varID In declinedIDs
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
                
                ' First, re-read current data from Kartei sheet to ensure we have latest values
                Dim karteiRow As Long
                karteiRow = FindRowByID_Sync(wsLocal, strID)
                
                If karteiRow = 0 Then
                    ' Row not found on Kartei - data inconsistency
                    Debug.Print "Warning: MoveDeclinedToPending - ID " & targetID & " not found on Kartei sheet."
                    MsgBox "Warning: Record ID " & targetID & " not found on Kartei sheet. Skipping.", vbExclamation, "Data Inconsistency"
                    GoTo NextDeclinedID
                End If
                
                ' Re-read fresh data from sheet (includes updated history from UpdateLocalSheetRowByID)
                ReDim arrRow(1 To 1, 1 To 52)
                Dim c As Long
                For c = 1 To 52
                    arrRow(1, c) = wsLocal.Cells(karteiRow, c).value
                Next c
                
                ' Re-read fresh formats
                ReDim arrFormats(1 To 1, 1 To 53)
                For c = 1 To 51
                    arrFormats(1, c) = wsLocal.Cells(karteiRow, c).Interior.Color
                Next c
                arrFormats(1, 52) = wsLocal.Cells(karteiRow, 3).Font.Color
                arrFormats(1, 53) = wsLocal.Cells(karteiRow, 18).Font.Color
                
                ' Write to pre_tblKartei (will create or update)
                Dim rsCheck As DAO.Recordset
                Set rsCheck = db.OpenRecordset("SELECT * FROM pre_tblKartei WHERE ID = " & targetID, dbOpenDynaset)
                
                If rsCheck.EOF Then
                    ' Create new record
                    rsCheck.Close
                    
                    Dim rsNew As DAO.Recordset
                    Set rsNew = db.OpenRecordset("pre_tblKartei", dbOpenDynaset)
                    rsNew.AddNew
                    rsNew.Fields("ID").value = targetID
                    Call FillPreRecordFromArray_Sync(rsNew, arrRow, arrFormats)
                    rsNew.Update
                    rsNew.Close
                Else
                    ' Update existing record
                    rsCheck.Edit
                    Call FillPreRecordFromArray_Sync(rsCheck, arrRow, arrFormats)
                    rsCheck.Update
                    rsCheck.Close
                End If
                
                ' Delete from decl_tblKartei with verification
                db.Execute "DELETE FROM decl_tblKartei WHERE ID = " & targetID
                
                ' DIAGNOSTIC: Log deletion (RecordsAffected is a property of Database, not Workspace)
                Debug.Print "Executed DELETE for ID " & targetID & ", RecordsAffected = " & db.RecordsAffected
                
                ' Check if deletion was successful
                If db.RecordsAffected = 0 Then
                    ' Record was not found in decl_tblKartei - data inconsistency
                    Debug.Print "Warning: MoveDeclinedToPending - ID " & targetID & " not found in decl_tblKartei."
                    ' This is a warning but not critical - record is already gone or was never there
                    ' Continue processing but log it
                End If
                
                ' Update status on Kartei sheet to PENDING
                wsLocal.Cells(karteiRow, STATUS_COL).value = "PENDING"
                
                ' DIAGNOSTIC: Log status update
                Debug.Print "Set BA=PENDING for row " & karteiRow & ", ID " & targetID
                
                ' Color column A light blue if D <> "Zahlung"
                Dim cellD As String
                cellD = Trim(CStr(wsLocal.Cells(karteiRow, 4).value))
                
                If cellD <> "Zahlung" Then
                    wsLocal.Cells(karteiRow, 1).Interior.Color = COLOR_PENDING
                    Debug.Print "Set pending color for row " & karteiRow & ", ID " & targetID
                End If
                
                ' Update corresponding row in Kartei_Original to prevent repeated history
                Call UpdateKarteiOriginalForDeclined(wsLocal, wsOriginal, karteiRow, targetID)
                
                processedCount = processedCount + 1
            End If
        End If
        
NextDeclinedID:
    Next varID
    
    wsDao.CommitTrans
    db.Close
    
    ' Debug log and user notification
    Debug.Print "MoveDeclinedToPending: Successfully processed " & processedCount & " of " & declinedIDs.Count & " declined records."
    
    ' TEMPORARY: Show visible confirmation
    If processedCount > 0 Then
        MsgBox "MoveDeclinedToPending completed:" & vbCrLf & _
               "Processed: " & processedCount & " records" & vbCrLf & _
               "Total requested: " & declinedIDs.Count, _
               vbInformation, "Declined Processing Complete"
    End If
    
    Exit Sub
    
RollbackTrans:
    wsDao.Rollback
    db.Close
    MsgBox "Error moving declined records to pending: " & Err.Description & vbCrLf & vbCrLf & _
           "Transaction rolled back. No changes were made.", vbCritical, "Database Error"
    Err.Raise Err.Number, Err.Source, Err.Description
End Sub

Private Sub FillPreRecordFromArray_Sync(ByVal rs As DAO.Recordset, _
                                        ByVal arrRow As Variant, _
                                        ByVal arrFormats As Variant)
    ' Fills recordset fields from array data
    ' Same logic as in Export_RiskClassification but local to this module
    
    Dim c As Long
    For c = 1 To 51
        Dim fieldName As String
        fieldName = "Value" & c
        If Not IsError(arrRow(1, c)) Then
            rs.Fields(fieldName).value = arrRow(1, c)
        Else
            rs.Fields(fieldName).value = ""
        End If
        
        fieldName = "InteriorColor" & c
        If IsNull(arrFormats(1, c)) Or IsEmpty(arrFormats(1, c)) Then
            rs.Fields(fieldName).value = 0
        Else
            rs.Fields(fieldName).value = arrFormats(1, c)
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

Private Sub EnsurePreTableExists(ByVal db As DAO.Database)
    ' Creates pre_tblKartei if it doesn't exist
    ' Delegates to Export_RiskClassification module
    Call Export_RiskClassification.EnsurePreTableExists(db)
End Sub

Private Sub UpdateKarteiOriginalForDeclined(ByVal wsKartei As Worksheet, _
                                           ByVal wsKarteiOriginal As Worksheet, _
                                           ByVal karteiRow As Long, _
                                           ByVal targetID As Long)
    ' Updates the corresponding row in Kartei_Original with current Kartei data
    ' This prevents the fixed declined record from requiring Notitzen again during next CompareAndSyncKartei
    ' Similar logic to UpdateKarteiOriginalRow in Export_DeclinedTools
    
    ' Find row in Kartei_Original by ID (column AV = 48)
    Dim lastRow As Long
    lastRow = wsKarteiOriginal.Cells(wsKarteiOriginal.Rows.Count, 1).End(xlUp).Row
    
    Dim origRow As Long
    origRow = 0
    
    Dim r As Long
    For r = 3 To lastRow
        Dim checkID As Variant
        checkID = wsKarteiOriginal.Cells(r, 48).value
        
        If Not IsEmpty(checkID) And IsNumeric(checkID) Then
            If CLng(checkID) = targetID Then
                origRow = r
                Exit For
            End If
        End If
    Next r
    
    If origRow = 0 Then
        ' ID not found in Original - shouldn't happen, but log it
        Debug.Print "Warning: UpdateKarteiOriginalForDeclined - ID " & targetID & " not found in Kartei_Original."
        Exit Sub
    End If
    
    ' Copy values from Kartei to Kartei_Original (columns 1..52 = A..AZ)
    Dim c As Integer
    For c = 1 To 52
        wsKarteiOriginal.Cells(origRow, c).value = wsKartei.Cells(karteiRow, c).value
        
        ' Copy formats for columns 1-51
        If c <= 51 Then
            wsKarteiOriginal.Cells(origRow, c).Interior.Color = wsKartei.Cells(karteiRow, c).Interior.Color
        End If
    Next c
    
    ' Copy font colors for columns C (3) and R (18)
    wsKarteiOriginal.Cells(origRow, 3).Font.Color = wsKartei.Cells(karteiRow, 3).Font.Color
    wsKarteiOriginal.Cells(origRow, 18).Font.Color = wsKartei.Cells(karteiRow, 18).Font.Color
    
    ' IMPORTANT: Also update status column BA (53) in Original to PENDING
    ' This ensures next sync sees the record as PENDING, not DECLINED
    wsKarteiOriginal.Cells(origRow, 53).value = "PENDING"
End Sub

Private Sub MarkRiskyRowsAsPending(ByVal ws As Worksheet, ByVal riskyIDs As Collection)
    ' Marks risky records as PENDING on Kartei sheet immediately after writing to pre_tblKartei
    ' This ensures visual feedback without waiting for file reopen
    
    Const STATUS_COL As Long = 53           ' BA column
    Const COLOR_PENDING As Long = 15849925  ' Light blue (same as in Export_OverlayPending)
    
    Dim varID As Variant
    Dim r As Long
    Dim cellD As String
    
    For Each varID In riskyIDs
        r = FindRowByID_Sync(ws, CStr(varID))
        If r > 0 Then
            ' Set PENDING status in BA column
            ws.Cells(r, STATUS_COL).value = "PENDING"
            
            ' Color column A light blue if D <> "Zahlung"
            cellD = Trim(CStr(ws.Cells(r, 4).value))
            If cellD <> "Zahlung" Then
                ws.Cells(r, 1).Interior.Color = COLOR_PENDING
            End If
        End If
    Next varID
End Sub



