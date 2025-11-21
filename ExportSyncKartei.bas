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
    
    ' Filter out PENDING/DECLINED rows - they cannot be written directly to tblKartei
    Dim filteredIDs As New Collection
    Dim checkID As Variant
    For Each checkID In changedIDs
        Dim checkRow As Long
        checkRow = FindRowByID_Sync(wsLocal, CStr(checkID))
        If checkRow > 0 Then
            Dim rowStatus As String
            rowStatus = GetRowStatus(wsLocal, checkRow)
            If rowStatus <> "PENDING" And rowStatus <> "DECLINED" Then
                filteredIDs.Add checkID
            End If
        Else
            ' Row not found, include it (new record case)
            filteredIDs.Add checkID
        End If
    Next checkID
    
    ' Use filtered IDs for further processing
    Set changedIDs = filteredIDs
    
    ' Classify changes into safe and risky
    Dim safeIDs As New Collection
    Dim riskyIDs As New Collection
    
    Dim changedID As Variant
    For Each changedID In changedIDs
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
    
    ' Process safe changes
    For Each changedID In safeIDs
        arrLocal = dictLocal(changedID)
        
        ' Update AW,AX,AY => userRole, Date, Time
        arrLocal(1, 49) = userRole
        arrLocal(1, 50) = Date
        arrLocal(1, 51) = Format(Time, "HH:MM")
        
        ' Overwrite local sheet row so AW,AX,AY are updated visually
        arrLocal(1, 52) = UpdateLocalSheetRowByID(wsLocal, wsOriginal, CStr(changedID), arrLocal, maxIDOriginal)
        dictLocal(changedID) = arrLocal
    Next changedID
    
    ' Process risky changes (update history/Notitzen on sheet, but don't write to tblKartei)
    For Each changedID In riskyIDs
        arrLocal = dictLocal(changedID)
        
        ' Update AW,AX,AY => userRole, Date, Time
        arrLocal(1, 49) = userRole
        arrLocal(1, 50) = Date
        arrLocal(1, 51) = Format(Time, "HH:MM")
        
        ' Update local sheet row with history
        arrLocal(1, 52) = UpdateLocalSheetRowByID(wsLocal, wsOriginal, CStr(changedID), arrLocal, maxIDOriginal)
        dictLocal(changedID) = arrLocal
    Next changedID
    
    ' Write safe changes to Access tblKartei
    If safeIDs.count > 0 Then
        WriteDictionaryChangesToAccess_Recordset dictLocal, dictLocalFormats, safeIDs
    End If
    
    ' Write risky changes to pre_tblKartei
    If riskyIDs.count > 0 Then
        WriteRiskyChangesToPreTable dictLocal, dictLocalFormats, riskyIDs
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
    If safeIDs.count > 0 Or riskyIDs.count > 0 Then
        msgText = "Synchronization completed:" & vbCrLf & _
                  "Safe changes written to database: " & safeIDs.count & vbCrLf & _
                  "Risky changes sent for approval: " & riskyIDs.count
        MsgBox msgText, vbInformation, "Sync Summary"
    End If
    
    If manualRun Then
        If safeIDs.count = 0 And riskyIDs.count = 0 Then
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

