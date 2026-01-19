Attribute VB_Name = "valid_ApproveFlow"
'==========================
'   Module: valid_ApproveFlow
'   Purpose: Superadmin workflow for approving/declining pending changes
'   Operations: Load pending, review GrossGeschichte, approve/decline decisions
'
'   MULTI-YEAR SUPPORT (2024, 2025, 2026):
'   Each year has its own decision sheet (grossGeschichte24, grossGeschichte25, grossGeschichte26)
'   and syncs decisions to the corresponding year database.
'
'   Entry points:
'     LoadPendingAndBuildDecisionForYear(year2) - Load + build decision sheet for year
'     LoadPendingAndBuildDecision24/25/26       - Wrapper macros for UX
'     SyncDecisionsForYear(year2)        - Sync decisions for year
'     SyncDecisions24/25/26              - Wrapper macros for UX
'     LoadPendingChanges                 - Legacy (defaults to year 25)
'     SyncDecisions                      - Legacy (defaults to year 25)
'   
'   History Format Constants (must match Export_HistoryBuilder in Admin file):
'     DCL - Decline entry tag
'     || - Session separator
'     -> - Value separator (old->new)
'     /@ @/ - Comment delimiters
'==========================

Option Explicit

' ========================================
' History Format Constants (self-contained for Superadmin)
' ========================================

Private Const HT_DECLINE As String = "DCL"           ' Decline entry tag
Private Const HD_SESSION As String = "||"            ' Session separator
Private Const HD_VALUE As String = "->"              ' Old->New value separator
Private Const HD_COMMENT_START As String = "/@"      ' Comment start marker
Private Const HD_COMMENT_END As String = "@/"        ' Comment end marker

' ============================================================
' MULTI-YEAR API - Load Pending Changes + Build Decision Sheet
' ============================================================

' Main workflow for a specific year: Load pending changes and generate decision sheet
' @param year2 - Two-digit year (24, 25, or 26)
Public Sub LoadPendingAndBuildDecisionForYear(ByVal year2 As Integer)
    On Error GoTo ErrorHandler
    
    ' Validate year
    If Not IsValidYear(year2) Then
        MsgBox "Ungueltiges Jahr: " & year2 & ". Unterstuetzte Jahre: 24, 25, 26.", _
               vbExclamation, "Jahresfehler"
        Exit Sub
    End If
    
    Application.ScreenUpdating = False
    Application.Calculation = xlCalculationManual
    
    ' Step 1: Load data from pre_tblKartei into KarteiYY sheet
    Call valid_ImportPending.LoadPendingChangesForYear(year2)
    
    ' Step 2: Build the War/Ist decision sheet for pending IDs
    Call valid_GrossGeschichteDecision.BuildPendingDecisionSheetForYear(year2)
    
    ' Step 3: Add decision dropdowns and styling to grossGeschichteYY
    Call PrepareGrossGeschichteForDecisionsForYear(year2)
    
    Dim sheetName As String
    sheetName = valid_YearConfig.GetGrossGeschichteSheetName(year2)
    
'    MsgBox "Ausstehende Aenderungen (Jahr " & year2 & ") erfolgreich geladen." & vbCrLf & vbCrLf & _
'           "Das " & sheetName & "-Blatt zeigt nun War/Ist-Vergleiche fuer alle Pending-IDs:" & vbCrLf & _
'           "  - War-Zeile: Originalwerte aus tblKartei (oder leer bei neuen Eintraegen)" & vbCrLf & _
'           "  - Ist-Zeile: Aktuelle Pending-Werte aus pre_tblKartei" & vbCrLf & _
'           "  - Unterschiede sind farblich hervorgehoben" & vbCrLf & vbCrLf & _
'           "Bitte markieren Sie jede Aenderung als:" & vbCrLf & _
'           "  - 'Approved' (Spalte AC) um zu akzeptieren und nach tblKartei zu verschieben" & vbCrLf & _
'           "  - 'Declined' (Spalte AC) um abzulehnen und nach decl_tblKartei zu verschieben" & vbCrLf & _
'           "  - Fuer abgelehnte Eintraege optional einen Kommentar in Spalte AD eingeben" & vbCrLf & vbCrLf & _
'           "Dann 'SyncDecisions" & year2 & "' ausfuehren um Ihre Entscheidungen zu verarbeiten.", _
'           vbInformation, "Ausstehende Aenderungen laden - Jahr " & year2
    
    Application.Calculation = xlCalculationAutomatic
    Application.ScreenUpdating = True
    Exit Sub
    
ErrorHandler:
    Application.Calculation = xlCalculationAutomatic
    Application.ScreenUpdating = True
    MsgBox "Fehler beim Laden der ausstehenden Aenderungen (Jahr " & year2 & "): " & Err.Description, _
           vbCritical, "Ladefehler"
End Sub

' Wrapper macro: Full workflow for year 2024 (load pending + build decision sheet)
Public Sub LoadPendingAndBuildDecision24()
    LoadPendingAndBuildDecisionForYear 24
End Sub

' Wrapper macro: Full workflow for year 2025 (load pending + build decision sheet)
Public Sub LoadPendingAndBuildDecision25()
    LoadPendingAndBuildDecisionForYear 25
End Sub

' Wrapper macro: Full workflow for year 2026 (load pending + build decision sheet)
Public Sub LoadPendingAndBuildDecision26()
    LoadPendingAndBuildDecisionForYear 26
End Sub

' ============================================================
' LEGACY API - Load Pending Changes (defaults to year 25)
' ============================================================

' Main workflow: Load pending changes and generate GrossGeschichte decision sheet for review
' This workflow always builds a War/Ist comparison sheet for pending IDs (no mode selection prompt).
Public Sub LoadPendingChanges()
    On Error GoTo ErrorHandler
    
    Application.ScreenUpdating = False
    Application.Calculation = xlCalculationManual
    
    ' Check if legacy single-year sheets exist; if so, use legacy workflow.
    Dim wsKartei As Worksheet
    On Error Resume Next
    Set wsKartei = ThisWorkbook.Worksheets("Kartei")
    On Error GoTo ErrorHandler
    
    If wsKartei Is Nothing Then
        ' No legacy Kartei sheet - use year 25 multi-year workflow
        Application.Calculation = xlCalculationAutomatic
        Application.ScreenUpdating = True
        LoadPendingAndBuildDecisionForYear 25
        Exit Sub
    End If
    
    ' Legacy workflow: use "Kartei" and "grossGeschichte" sheets
    
    ' Step 1: Load data from pre_tblKartei into Kartei sheet
    ' Note: SetupGrossGeschichteDates is called internally by LoadPendingChangesFromPre
    '       to set B1/C1 dates on GrossGeschichte sheet
    Call valid_ImportPending.LoadPendingChangesFromPre
    
    ' Step 2: Build the War/Ist decision sheet for pending IDs
    ' This replaces the old grossGeschichte.GrossGeshichteMachen call.
    ' BuildPendingDecisionSheet compares pending data (Kartei/pre_tblKartei) with
    ' original data (tblKartei), highlights differences, and aggregates comments.
    Call valid_GrossGeschichteDecision.BuildPendingDecisionSheet
    
    ' Step 3: Add decision dropdowns and styling to GrossGeschichte
    Call PrepareGrossGeschichteForDecisions
    
    MsgBox "Ausstehende Aenderungen erfolgreich geladen." & vbCrLf & vbCrLf & _
           "Das GrossGeschichte-Blatt zeigt nun War/Ist-Vergleiche fuer alle Pending-IDs:" & vbCrLf & _
           "  - War-Zeile: Originalwerte aus tblKartei (oder leer bei neuen Eintraegen)" & vbCrLf & _
           "  - Ist-Zeile: Aktuelle Pending-Werte aus pre_tblKartei" & vbCrLf & _
           "  - Unterschiede sind farblich hervorgehoben" & vbCrLf & vbCrLf & _
           "Bitte markieren Sie jede Aenderung als:" & vbCrLf & _
           "  - 'Approved' (Spalte AC) um zu akzeptieren und nach tblKartei zu verschieben" & vbCrLf & _
           "  - 'Declined' (Spalte AC) um abzulehnen und nach decl_tblKartei zu verschieben" & vbCrLf & _
           "  - Fuer abgelehnte Eintraege optional einen Kommentar in Spalte AD eingeben" & vbCrLf & vbCrLf & _
           "Dann Taste 'Validate' druecken um Ihre Entscheidungen zu verarbeiten.", _
           vbInformation, "Ausstehende Aenderungen laden"
    
    Application.Calculation = xlCalculationAutomatic
    Application.ScreenUpdating = True
    Exit Sub
    
ErrorHandler:
    Application.Calculation = xlCalculationAutomatic
    Application.ScreenUpdating = True
    MsgBox "Fehler beim Laden der ausstehenden Aenderungen: " & Err.Description, vbCritical, "Ladefehler"
End Sub

' Prepare GrossGeschichte sheet with decision column (AC) and decline comment column (AD)
Private Sub PrepareGrossGeschichteForDecisions()
    Dim ws As Worksheet
    On Error Resume Next
    Set ws = ThisWorkbook.Worksheets("grossGeschichte")
    On Error GoTo 0
    
    If ws Is Nothing Then Exit Sub
    
    PrepareDecisionColumnsOnSheet ws
End Sub

' Prepare grossGeschichteYY sheet with decision column (AC) and decline comment column (AD)
Private Sub PrepareGrossGeschichteForDecisionsForYear(ByVal year2 As Integer)
    Dim sheetName As String
    sheetName = valid_YearConfig.GetGrossGeschichteSheetName(year2)
    
    Dim ws As Worksheet
    On Error Resume Next
    Set ws = ThisWorkbook.Worksheets(sheetName)
    On Error GoTo 0
    
    If ws Is Nothing Then Exit Sub
    
    PrepareDecisionColumnsOnSheet ws
End Sub

' Common logic to prepare decision columns on any grossGeschichte sheet
Private Sub PrepareDecisionColumnsOnSheet(ByVal ws As Worksheet)
    ' Header for Decision is already created in CreateGrossGeschichteHeaders (column AC = 29)
    ' Just ensure styling is applied
    ws.Range("AC2").Interior.Color = RGB(255, 255, 153) ' Light yellow
    
    ' Add validation dropdown (Approved/Declined) to decision cells
    Dim lastRow As Long
    lastRow = ws.Cells(ws.Rows.Count, 1).End(xlUp).Row
    
    If lastRow > 2 Then
        ' Find all "Ist" rows (every 3rd row starting from row 4: 4, 7, 10, ...)
        Dim r As Long
        For r = 4 To lastRow Step 3 ' Ist rows are at positions 4, 7, 10, ...
            ' Add dropdown validation to column AC (29)
            With ws.Range("AC" & r).Validation
                .Delete
                .Add Type:=xlValidateList, AlertStyle:=xlValidAlertStop, _
                     Formula1:="Approved,Declined"
                .IgnoreBlank = True
                .InCellDropdown = True
            End With
            
            ws.Range("AC" & r).Interior.Color = RGB(255, 255, 204) ' Lighter yellow
            
            ' Prepare column AD (30) for decline comments (no validation, just styling)
            ws.Range("AD" & r).Interior.Color = RGB(255, 230, 230) ' Light pink
        Next r
    End If
End Sub

' ============================================================
' MULTI-YEAR API - Sync Decisions
' ============================================================

' Process all decisions from grossGeschichteYY sheet for a specific year
' @param year2 - Two-digit year (24, 25, or 26)
Public Sub SyncDecisionsForYear(ByVal year2 As Integer)
    On Error GoTo ErrorHandler
    
    ' Validate year
    If Not IsValidYear(year2) Then
        MsgBox "Ungueltiges Jahr: " & year2 & ". Unterstuetzte Jahre: 24, 25, 26.", _
               vbExclamation, "Jahresfehler"
        Exit Sub
    End If
    
    Application.ScreenUpdating = False
    Application.Calculation = xlCalculationManual
    
    Dim sheetName As String
    sheetName = valid_YearConfig.GetGrossGeschichteSheetName(year2)
    
    Dim wsGross As Worksheet
    On Error Resume Next
    Set wsGross = ThisWorkbook.Worksheets(sheetName)
    On Error GoTo ErrorHandler
    
    If wsGross Is Nothing Then
        MsgBox sheetName & "-Blatt nicht gefunden." & vbCrLf & vbCrLf & _
               "Bitte zuerst:" & vbCrLf & _
               "  - 'LoadPendingAndBuildDecision" & year2 & "' (Empfohlen) oder" & vbCrLf & _
               "  - 'LoadPendingChanges" & year2 & "' + 'BuildPendingDecisionSheet" & year2 & "'", _
               vbExclamation, "Blatt nicht gefunden"
        GoTo Cleanup
    End If
    
    ' Collect and process decisions
    Dim approvedIDs As Collection
    Dim declinedIDs As Collection
    Set approvedIDs = New Collection
    Set declinedIDs = New Collection
    
    CollectDecisionsFromSheet wsGross, approvedIDs, declinedIDs
    
    If approvedIDs.Count = 0 And declinedIDs.Count = 0 Then
        MsgBox "Keine Entscheidungen gefunden. Bitte markieren Sie Eintraege als 'Approved' oder 'Declined' in Spalte AC.", _
               vbExclamation, "Keine Entscheidungen"
        GoTo Cleanup
    End If
    
    ' Process approved records (year-specific)
    If approvedIDs.Count > 0 Then
        Call ProcessApprovedRecordsForYear(approvedIDs, year2)
    End If
    
    ' Process declined records (year-specific)
    If declinedIDs.Count > 0 Then
        Call ProcessDeclinedRecordsForYear(declinedIDs, year2)
    End If
    
    MsgBox "Entscheidungen (Jahr " & year2 & ") erfolgreich verarbeitet:" & vbCrLf & _
           "  Genehmigt: " & approvedIDs.Count & vbCrLf & _
           "  Abgelehnt: " & declinedIDs.Count, _
           vbInformation, "Synchronisierung abgeschlossen - Jahr " & year2
    
Cleanup:
    Application.Calculation = xlCalculationAutomatic
    Application.ScreenUpdating = True
    Exit Sub
    
ErrorHandler:
    Application.Calculation = xlCalculationAutomatic
    Application.ScreenUpdating = True
    MsgBox "Fehler bei der Synchronisierung der Entscheidungen (Jahr " & year2 & "): " & Err.Description, _
           vbCritical, "Synchronisierungsfehler"
End Sub

' Wrapper macro: Sync decisions for year 2024
Public Sub SyncDecisions24()
    SyncDecisionsForYear 24
End Sub

' Wrapper macro: Sync decisions for year 2025
Public Sub SyncDecisions25()
    SyncDecisionsForYear 25
End Sub

' Wrapper macro: Sync decisions for year 2026
Public Sub SyncDecisions26()
    SyncDecisionsForYear 26
End Sub

' ============================================================
' LEGACY API - Sync Decisions (defaults to year 25)
' ============================================================

' Process all decisions from GrossGeschichte sheet
' GrossGeschichte now contains exactly ONE War/Ist block per pending ID
' (built by valid_GrossGeschichteDecision.BuildPendingDecisionSheet).
' The idDecisions dictionary is kept as a defensive measure in case
' future changes introduce multiple blocks per ID.
Public Sub SyncDecisions()
    On Error GoTo ErrorHandler
    
    Application.ScreenUpdating = False
    Application.Calculation = xlCalculationManual
    
    ' Check if legacy single-year sheet exists; if so, use it.
    Dim wsGross As Worksheet
    On Error Resume Next
    Set wsGross = ThisWorkbook.Worksheets("grossGeschichte")
    On Error GoTo ErrorHandler
    
    If wsGross Is Nothing Then
        ' No legacy grossGeschichte sheet - use year 25 multi-year workflow
        Application.Calculation = xlCalculationAutomatic
        Application.ScreenUpdating = True
        SyncDecisionsForYear 25
        Exit Sub
    End If
    
    ' Legacy workflow: use "grossGeschichte" sheet and year 25 DB
    
    Dim approvedIDs As Collection
    Dim declinedIDs As Collection
    Set approvedIDs = New Collection
    Set declinedIDs = New Collection
    
    CollectDecisionsFromSheet wsGross, approvedIDs, declinedIDs
    
    If approvedIDs.Count = 0 And declinedIDs.Count = 0 Then
        MsgBox "Keine Entscheidungen gefunden. Bitte markieren Sie Eintraege als 'Approved' oder 'Declined' in Spalte AC.", _
               vbExclamation, "Keine Entscheidungen"
        GoTo Cleanup
    End If
    
    ' Process approved records (legacy - year 25)
    If approvedIDs.Count > 0 Then
        Call ProcessApprovedRecords(approvedIDs)
    End If
    
    ' Process declined records (legacy - year 25)
    If declinedIDs.Count > 0 Then
        Call ProcessDeclinedRecords(declinedIDs)
    End If
    
    MsgBox "Entscheidungen erfolgreich verarbeitet:" & vbCrLf & _
           "  Genehmigt: " & approvedIDs.Count & vbCrLf & _
           "  Abgelehnt: " & declinedIDs.Count, _
           vbInformation, "Synchronisierung abgeschlossen"
    
Cleanup:
    Application.Calculation = xlCalculationAutomatic
    Application.ScreenUpdating = True
    Exit Sub
    
ErrorHandler:
    Application.Calculation = xlCalculationAutomatic
    Application.ScreenUpdating = True
    MsgBox "Fehler bei der Synchronisierung der Entscheidungen: " & Err.Description, vbCritical, "Synchronisierungsfehler"
End Sub

' ============================================================
' DECISION COLLECTION (Shared by legacy and multi-year)
' ============================================================

' Collect decisions from a grossGeschichte sheet into approved/declined collections
Private Sub CollectDecisionsFromSheet(ByVal wsGross As Worksheet, _
                                       ByRef approvedIDs As Collection, _
                                       ByRef declinedIDs As Collection)
    Dim lastRow As Long
    lastRow = wsGross.Cells(wsGross.Rows.Count, 1).End(xlUp).Row
    
    ' Dictionary to track decisions per ID
    Dim idDecisions As Object
    Set idDecisions = CreateObject("Scripting.Dictionary")
    
    ' Track IDs with conflicting decisions (for warning)
    Dim conflictIDs As Collection
    Set conflictIDs = New Collection
    
    ' Collect decisions by ID
    ' Structure: GrossGeschichte has War/Ist/Separator rows in groups of 3
    ' Ist rows are at positions 4, 7, 10, ... (every 3rd row starting from 4)
    ' RecordID is stored in hidden column AE (31)
    Dim r As Long
    For r = 4 To lastRow Step 3 ' Process "Ist" rows only
        Dim decision As String
        decision = UCase(Trim(wsGross.Range("AC" & r).Value)) ' Column AC (29) = Decision
        
        If decision = "APPROVED" Or decision = "DECLINED" Then
            ' Get RecordID from hidden column AE (31)
            Dim recordID As String
            recordID = Trim(CStr(wsGross.Cells(r, 31).Value))
            
            If recordID <> "" Then
                ' Build decision data: (row, decision type, comment)
                Dim decisionData As Variant
                
                If decision = "APPROVED" Then
                    decisionData = Array(r, "APPROVED", "")
                Else
                    ' For declined, get comment from column AD (30)
                    Dim declComment As String
                    declComment = Trim(wsGross.Range("AD" & r).Value)
                    decisionData = Array(r, "DECLINED", declComment)
                End If
                
                ' Store in dictionary; if ID already exists, check for conflict
                If idDecisions.Exists(recordID) Then
                    Dim existingData As Variant
                    existingData = idDecisions(recordID)
                    
                    ' Check for conflicting decisions on same ID
                    If CStr(existingData(1)) <> CStr(decisionData(1)) Then
                        On Error Resume Next
                        conflictIDs.Add recordID, recordID
                        On Error GoTo 0
                    End If
                    
                    ' Keep the one with higher row number (later block takes precedence)
                    If CLng(decisionData(0)) > CLng(existingData(0)) Then
                        idDecisions(recordID) = decisionData
                    End If
                Else
                    idDecisions.Add recordID, decisionData
                End If
            End If
        End If
    Next r
    
    ' Warn about conflicting decisions
    If conflictIDs.Count > 0 Then
        Dim conflictMsg As String
        conflictMsg = "Warnung: Fuer folgende IDs wurden widersprüchliche Entscheidungen gefunden:" & vbCrLf
        Dim conflictID As Variant
        For Each conflictID In conflictIDs
            conflictMsg = conflictMsg & "  - ID " & conflictID & vbCrLf
        Next conflictID
        conflictMsg = conflictMsg & vbCrLf & "Es wird jeweils die letzte Entscheidung (hoehere Zeile) verwendet."
        MsgBox conflictMsg, vbExclamation, "Widersprüchliche Entscheidungen"
    End If
    
    ' Process unique decisions per ID
    Dim idKey As Variant
    For Each idKey In idDecisions.Keys
        Dim finalDecision As Variant
        finalDecision = idDecisions(idKey)
        
        Dim finalDecisionType As String
        finalDecisionType = CStr(finalDecision(1))
        
        If finalDecisionType = "APPROVED" Then
            approvedIDs.Add CStr(idKey)
        ElseIf finalDecisionType = "DECLINED" Then
            Dim finalComment As String
            finalComment = CStr(finalDecision(2))
            
            ' If comment is still empty, prompt for it
            If finalComment = "" Then
                finalComment = InputBox("Geben Sie den Grund fuer die Ablehnung von ID " & idKey & " ein:" & vbCrLf & _
                                      "(Sie koennen Kommentare auch direkt in Spalte AD von GrossGeschichte eingeben)", _
                                      "Ablehnungskommentar", "Superadmin hat diese Aenderung abgelehnt")
            End If
            
            If finalComment <> "" Then
                declinedIDs.Add Array(CStr(idKey), finalComment)
            Else
                MsgBox "ID " & idKey & " wird uebersprungen - kein Kommentar angegeben.", vbExclamation
            End If
        End If
    Next idKey
End Sub

' Get record ID for a given GrossGeschichte row from hidden column AE (31)
' Returns the ID or empty string if not found
Private Function GetRecordIDForGrossRow(ByVal grossRow As Long) As String
    On Error GoTo ErrHandler
    
    Dim wsGross As Worksheet
    Set wsGross = ThisWorkbook.Worksheets("grossGeschichte")
    
    ' RecordID is stored in hidden column AE (column 31)
    Dim recordID As String
    recordID = Trim(CStr(wsGross.Cells(grossRow, 31).Value))
    
    GetRecordIDForGrossRow = recordID
    Exit Function
    
ErrHandler:
    GetRecordIDForGrossRow = ""
End Function

' Move approved records from pre_tblKartei to tblKartei
Private Sub ProcessApprovedRecords(ByVal approvedIDs As Collection)
    Dim dbPath As String
    dbPath = GetDatabasePath()
    
    ' Check if user cancelled database selection
    If dbPath = "" Then
        MsgBox "Datenbankpfad nicht festgelegt. Vorgang abgebrochen.", vbExclamation, "Fehler"
        Exit Sub
    End If
    
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
    
    ' Check if user cancelled database selection
    If dbPath = "" Then
        MsgBox "Datenbankpfad nicht festgelegt. Vorgang abgebrochen.", vbExclamation, "Fehler"
        Exit Sub
    End If
    
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
            
            ' Count existing decline entries (both old Decl_N and new DCL format)
            Dim declNum As Long
            declNum = CountDeclEntries(currentHistory) + 1
            
            ' Build decline entry using new format: DCL(<N>-><comment + timestamp>)||
            Dim declEntry As String
            declEntry = BuildDeclineHistoryEntry(declNum, declComment)
            
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

' ============================================================
' MULTI-YEAR DATABASE OPERATIONS
' ============================================================

' Move approved records from pre_tblKartei to tblKartei (year-specific)
Private Sub ProcessApprovedRecordsForYear(ByVal approvedIDs As Collection, ByVal year2 As Integer)
    Dim dbPath As String
    dbPath = GetDatabasePathForYear(year2)
    
    ' Check if user cancelled database selection
    If dbPath = "" Then
        MsgBox "Datenbankpfad fuer Jahr " & year2 & " nicht festgelegt. Vorgang abgebrochen.", vbExclamation, "Fehler"
        Exit Sub
    End If
    
    Dim engine As DAO.DBEngine
    Set engine = New DAO.DBEngine
    
    Dim wsDao As DAO.Workspace
    Set wsDao = engine.Workspaces(0)
    
    Dim db As DAO.Database
    Set db = wsDao.OpenDatabase(dbPath)
    
    Dim fld As DAO.Field
    
    wsDao.BeginTrans
    
    On Error GoTo RollbackTrans
    
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
                
                ' Copy all fields INCLUDING ID
                For Each fld In rsPre.Fields
                    On Error Resume Next
                    rsMain.Fields(fld.Name).Value = rsPre.Fields(fld.Name).Value
                    On Error GoTo RollbackTrans
                Next fld
            Else
                ' Record exists - update
                rsMain.Edit
                
                For Each fld In rsPre.Fields
                    If fld.Name <> "ID" Then
                        On Error Resume Next
                        rsMain.Fields(fld.Name).Value = rsPre.Fields(fld.Name).Value
                        On Error GoTo RollbackTrans
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
    Exit Sub
    
RollbackTrans:
    On Error Resume Next
    wsDao.Rollback
    db.Close
    MsgBox "Fehler beim Verarbeiten der genehmigten Eintraege (Jahr " & year2 & "): " & Err.Description, _
           vbCritical, "Datenbankfehler"
End Sub

' Move declined records from pre_tblKartei to decl_tblKartei (year-specific)
Private Sub ProcessDeclinedRecordsForYear(ByVal declinedIDs As Collection, ByVal year2 As Integer)
    Dim dbPath As String
    dbPath = GetDatabasePathForYear(year2)
    
    ' Check if user cancelled database selection
    If dbPath = "" Then
        MsgBox "Datenbankpfad fuer Jahr " & year2 & " nicht festgelegt. Vorgang abgebrochen.", vbExclamation, "Fehler"
        Exit Sub
    End If
    
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
    
    On Error GoTo RollbackTrans
    
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
            
            ' Copy all fields from pre_ to decl_ table
            Dim fld As DAO.Field
            For Each fld In rsPre.Fields
                On Error Resume Next
                rsDecl.Fields(fld.Name).Value = rsPre.Fields(fld.Name).Value
                On Error GoTo RollbackTrans
            Next fld
            
            ' Add decline comment to Value52 (history field)
            Dim currentHistory As String
            currentHistory = Nz(rsDecl.Fields("Value52").Value, "")
            
            ' Count existing decline entries
            Dim declNum As Long
            declNum = CountDeclEntries(currentHistory) + 1
            
            ' Build decline entry
            Dim declEntry As String
            declEntry = BuildDeclineHistoryEntry(declNum, declComment)
            
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
    Exit Sub
    
RollbackTrans:
    On Error Resume Next
    wsDao.Rollback
    db.Close
    MsgBox "Fehler beim Verarbeiten der abgelehnten Eintraege (Jahr " & year2 & "): " & Err.Description, _
           vbCritical, "Datenbankfehler"
End Sub

' ============================================================
' HISTORY FORMAT HELPERS
' ============================================================

' Count existing Decl_N entries in history string
Private Function CountDeclEntries(ByVal historyStr As String) As Long
    ' Counts both old format (Decl_N:) and new format (DCL(N->...))
    Dim count As Long
    count = 0
    
    ' Count new format: DCL(
    Dim pos As Long
    pos = 1
    Do
        pos = InStr(pos, historyStr, HT_DECLINE & "(")
        If pos > 0 Then
            count = count + 1
            pos = pos + Len(HT_DECLINE) + 1
        End If
    Loop While pos > 0
    
    ' Count old format: Decl_
    pos = 1
    Do
        pos = InStr(pos, historyStr, "Decl_")
        If pos > 0 Then
            count = count + 1
            pos = pos + 5
        End If
    Loop While pos > 0
    
    CountDeclEntries = count
End Function

' Build a decline history entry using new format
Private Function BuildDeclineHistoryEntry(ByVal declineNumber As Long, _
                                          ByVal declineComment As String) As String
    ' Format: DCL(<N>-><comment (Declined by Superadmin on DD.MM.YYYY)>)||
    
    Dim sanitizedComment As String
    sanitizedComment = SanitizeHistoryValue(declineComment)
    
    Dim fullComment As String
    fullComment = sanitizedComment & " (Abgelehnt vom Superadmin am " & Format(Date, "dd.mm.yyyy") & ")"
    
    BuildDeclineHistoryEntry = HT_DECLINE & "(" & CStr(declineNumber) & HD_VALUE & fullComment & ")" & HD_SESSION
End Function

' Sanitize a value for inclusion in history string
Private Function SanitizeHistoryValue(ByVal value As String) As String
    Dim result As String
    result = value
    
    ' Remove/replace problematic characters
    result = Replace(result, HD_VALUE, "~>")       ' Escape arrow
    result = Replace(result, HD_SESSION, "|")      ' Escape double pipe
    result = Replace(result, "(", "[")             ' Escape open paren
    result = Replace(result, ")", "]")             ' Escape close paren
    result = Replace(result, HD_COMMENT_START, "/")
    result = Replace(result, HD_COMMENT_END, "/")
    result = Replace(result, vbCrLf, " ")          ' Remove line breaks
    result = Replace(result, vbCr, " ")
    result = Replace(result, vbLf, " ")
    
    SanitizeHistoryValue = Trim(result)
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

' Helper: Get database path with validation (prompts user if file not found)
' Defaults to year 25 for backward compatibility
Private Function GetDatabasePath() As String
    GetDatabasePath = valid_DatabasePath.GetValidatedDatabasePath()
End Function

' Helper: Get database path for a specific year
Private Function GetDatabasePathForYear(ByVal year2 As Integer) As String
    GetDatabasePathForYear = valid_DatabasePath.GetValidatedDatabasePathForYear(year2)
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

' ============================================================
' YEAR VALIDATION HELPER
' ============================================================

' Check if year is valid (24, 25, or 26)
Private Function IsValidYear(ByVal year2 As Integer) As Boolean
    IsValidYear = (year2 = 24 Or year2 = 25 Or year2 = 26)
End Function

' ========================================
' Bulk Decision Tools (Multi-Year)
' ========================================

' Set "Approved" decision for all pending records on grossGeschichteYY sheet
' @param year2 - Two-digit year (24, 25, or 26)
Public Sub ApproveAllPendingForYear(ByVal year2 As Integer)
    On Error GoTo ErrorHandler
    
    If Not IsValidYear(year2) Then
        MsgBox "Ungueltiges Jahr: " & year2, vbExclamation, "Jahresfehler"
        Exit Sub
    End If
    
    Dim sheetName As String
    sheetName = valid_YearConfig.GetGrossGeschichteSheetName(year2)
    
    Dim wsGross As Worksheet
    On Error Resume Next
    Set wsGross = ThisWorkbook.Worksheets(sheetName)
    On Error GoTo ErrorHandler
    
    If wsGross Is Nothing Then
        MsgBox sheetName & "-Blatt nicht gefunden." & vbCrLf & vbCrLf & _
               "Bitte zuerst:" & vbCrLf & _
               "  - 'LoadPendingAndBuildDecision" & year2 & "' (Empfohlen) oder" & vbCrLf & _
               "  - 'LoadPendingChanges" & year2 & "' + 'BuildPendingDecisionSheet" & year2 & "'", _
               vbExclamation, "Blatt nicht gefunden"
        Exit Sub
    End If
    
    ' Confirm action
    Dim confirmResult As VbMsgBoxResult
    confirmResult = MsgBox("Dies setzt 'Approved' fuer ALLE ausstehenden Eintraege auf " & sheetName & "." & vbCrLf & vbCrLf & _
                          "Sind Sie sicher, dass Sie alle Eintraege genehmigen moechten?", _
                          vbYesNo + vbQuestion, "Alle genehmigen bestaetigen - Jahr " & year2)
    
    If confirmResult <> vbYes Then
        Exit Sub
    End If
    
    Dim count As Long
    count = SetAllDecisionsOnSheet(wsGross, "Approved", "")
    
    MsgBox count & " Eintrag/Eintraege als 'Approved' markiert." & vbCrLf & vbCrLf & _
           "Fuehren Sie 'SyncDecisions" & year2 & "' aus, um diese Entscheidungen auf die Datenbank anzuwenden.", _
           vbInformation, "Alle genehmigen abgeschlossen - Jahr " & year2
    
    Exit Sub
    
ErrorHandler:
    Application.ScreenUpdating = True
    MsgBox "Fehler in ApproveAllPendingForYear: " & Err.Description, vbCritical, "Fehler"
End Sub

' Set "Declined" decision for all pending records on grossGeschichteYY sheet
' @param year2 - Two-digit year (24, 25, or 26)
Public Sub DeclineAllPendingForYear(ByVal year2 As Integer)
    On Error GoTo ErrorHandler
    
    If Not IsValidYear(year2) Then
        MsgBox "Ungueltiges Jahr: " & year2, vbExclamation, "Jahresfehler"
        Exit Sub
    End If
    
    Dim sheetName As String
    sheetName = valid_YearConfig.GetGrossGeschichteSheetName(year2)
    
    Dim wsGross As Worksheet
    On Error Resume Next
    Set wsGross = ThisWorkbook.Worksheets(sheetName)
    On Error GoTo ErrorHandler
    
    If wsGross Is Nothing Then
        MsgBox sheetName & "-Blatt nicht gefunden." & vbCrLf & vbCrLf & _
               "Bitte zuerst:" & vbCrLf & _
               "  - 'LoadPendingAndBuildDecision" & year2 & "' (Empfohlen) oder" & vbCrLf & _
               "  - 'LoadPendingChanges" & year2 & "' + 'BuildPendingDecisionSheet" & year2 & "'", _
               vbExclamation, "Blatt nicht gefunden"
        Exit Sub
    End If
    
    ' Prompt for decline comment
    Dim commonComment As String
    commonComment = InputBox("Geben Sie einen gemeinsamen Ablehnungskommentar fuer ALLE Eintraege ein:" & vbCrLf & _
                            "(Dieser Kommentar wird auf alle abgelehnten Eintraege angewendet)", _
                            "Ablehnungskommentar - Jahr " & year2, "Massenablehnung durch Superadmin")
    
    If commonComment = "" Then
        Dim confirmEmpty As VbMsgBoxResult
        confirmEmpty = MsgBox("Kein Kommentar eingegeben. Moechten Sie ohne Kommentar fortfahren?", _
                             vbYesNo + vbQuestion, "Kein Kommentar")
        If confirmEmpty <> vbYes Then
            Exit Sub
        End If
    End If
    
    ' Confirm action
    Dim confirmResult As VbMsgBoxResult
    confirmResult = MsgBox("Dies setzt 'Declined' fuer ALLE ausstehenden Eintraege auf " & sheetName & "." & vbCrLf & vbCrLf & _
                          "Sind Sie sicher, dass Sie alle Eintraege ablehnen moechten?", _
                          vbYesNo + vbQuestion, "Alle ablehnen bestaetigen - Jahr " & year2)
    
    If confirmResult <> vbYes Then
        Exit Sub
    End If
    
    Dim count As Long
    count = SetAllDecisionsOnSheet(wsGross, "Declined", commonComment)
    
    MsgBox count & " Eintrag/Eintraege als 'Declined' markiert." & vbCrLf & vbCrLf & _
           "Fuehren Sie 'SyncDecisions" & year2 & "' aus, um diese Entscheidungen auf die Datenbank anzuwenden.", _
           vbInformation, "Alle ablehnen abgeschlossen - Jahr " & year2
    
    Exit Sub
    
ErrorHandler:
    Application.ScreenUpdating = True
    MsgBox "Fehler in DeclineAllPendingForYear: " & Err.Description, vbCritical, "Fehler"
End Sub

' Clear all decisions on grossGeschichteYY sheet
' @param year2 - Two-digit year (24, 25, or 26)
Public Sub ClearAllDecisionsForYear(ByVal year2 As Integer)
    On Error GoTo ErrorHandler
    
    If Not IsValidYear(year2) Then
        MsgBox "Ungueltiges Jahr: " & year2, vbExclamation, "Jahresfehler"
        Exit Sub
    End If
    
    Dim sheetName As String
    sheetName = valid_YearConfig.GetGrossGeschichteSheetName(year2)
    
    Dim wsGross As Worksheet
    On Error Resume Next
    Set wsGross = ThisWorkbook.Worksheets(sheetName)
    On Error GoTo ErrorHandler
    
    If wsGross Is Nothing Then
        MsgBox sheetName & "-Blatt nicht gefunden.", vbExclamation, "Blatt nicht gefunden"
        Exit Sub
    End If
    
    Dim confirmResult As VbMsgBoxResult
    confirmResult = MsgBox("Dies loescht ALLE Entscheidungen auf " & sheetName & "." & vbCrLf & vbCrLf & _
                          "Sind Sie sicher?", _
                          vbYesNo + vbQuestion, "Alle loeschen bestaetigen - Jahr " & year2)
    
    If confirmResult <> vbYes Then
        Exit Sub
    End If
    
    Dim count As Long
    count = SetAllDecisionsOnSheet(wsGross, "", "")
    
    MsgBox count & " Entscheidung(en) geloescht.", vbInformation, "Alle loeschen abgeschlossen - Jahr " & year2
    
    Exit Sub
    
ErrorHandler:
    Application.ScreenUpdating = True
    MsgBox "Fehler in ClearAllDecisionsForYear: " & Err.Description, vbCritical, "Fehler"
End Sub

' Wrapper macros for bulk decisions
Public Sub ApproveAllPending24(): ApproveAllPendingForYear 24: End Sub
Public Sub ApproveAllPending25(): ApproveAllPendingForYear 25: End Sub
Public Sub ApproveAllPending26(): ApproveAllPendingForYear 26: End Sub
Public Sub DeclineAllPending24(): DeclineAllPendingForYear 24: End Sub
Public Sub DeclineAllPending25(): DeclineAllPendingForYear 25: End Sub
Public Sub DeclineAllPending26(): DeclineAllPendingForYear 26: End Sub
Public Sub ClearAllDecisions24(): ClearAllDecisionsForYear 24: End Sub
Public Sub ClearAllDecisions25(): ClearAllDecisionsForYear 25: End Sub
Public Sub ClearAllDecisions26(): ClearAllDecisionsForYear 26: End Sub

' Common logic to set decisions on a grossGeschichte sheet
' Returns the count of affected rows
Private Function SetAllDecisionsOnSheet(ByVal ws As Worksheet, _
                                        ByVal decision As String, _
                                        ByVal comment As String) As Long
    Application.ScreenUpdating = False
    
    Dim lastRow As Long
    lastRow = ws.Cells(ws.Rows.Count, 1).End(xlUp).Row
    
    Dim count As Long
    count = 0
    
    If lastRow >= 4 Then
        Dim r As Long
        For r = 4 To lastRow Step 3
            If Trim(CStr(ws.Cells(r, 31).Value)) <> "" Then
                ws.Cells(r, 29).Value = decision
                If decision = "Declined" And comment <> "" Then
                    ws.Cells(r, 30).Value = comment
                ElseIf decision = "" Then
                    ws.Cells(r, 30).Value = ""
                End If
                count = count + 1
            End If
        Next r
    End If
    
    Application.ScreenUpdating = True
    SetAllDecisionsOnSheet = count
End Function

' ========================================
' Bulk Decision Tools (Legacy)
' ========================================

' Set "Approved" decision for all pending records on GrossGeschichte sheet
Public Sub ApproveAllPending()
    On Error GoTo ErrorHandler
    
    Dim wsGross As Worksheet
    On Error Resume Next
    Set wsGross = ThisWorkbook.Worksheets("grossGeschichte")
    On Error GoTo ErrorHandler
    
    If wsGross Is Nothing Then
        MsgBox "grossGeschichte-Blatt nicht gefunden. Bitte fuehren Sie zuerst 'LoadPendingChanges' aus.", _
               vbExclamation, "Blatt nicht gefunden"
        Exit Sub
    End If
    
    ' Confirm action
    Dim confirmResult As VbMsgBoxResult
    confirmResult = MsgBox("Dies setzt 'Approved' fuer ALLE ausstehenden Eintraege auf grossGeschichte." & vbCrLf & vbCrLf & _
                          "Sind Sie sicher, dass Sie alle Eintraege genehmigen moechten?", _
                          vbYesNo + vbQuestion, "Alle genehmigen bestaetigen")
    
    If confirmResult <> vbYes Then
        Exit Sub
    End If
    
    Application.ScreenUpdating = False
    
    Dim lastRow As Long
    lastRow = wsGross.Cells(wsGross.Rows.Count, 1).End(xlUp).Row
    
    If lastRow < 4 Then
        MsgBox "No pending records found on grossGeschichte.", vbInformation, "No Records"
        Application.ScreenUpdating = True
        Exit Sub
    End If
    
    Dim count As Long
    count = 0
    
    ' Process "Ist" rows (every 3rd row starting from row 4: 4, 7, 10, ...)
    Dim r As Long
    For r = 4 To lastRow Step 3
        ' Check if row has data (RecordID in column AE should be present)
        If Trim(CStr(wsGross.Cells(r, 31).Value)) <> "" Then
            ' Set Decision to "Approved" (column AC = 29)
            wsGross.Cells(r, 29).Value = "Approved"
            count = count + 1
        End If
    Next r
    
    Application.ScreenUpdating = True
    
    MsgBox count & " Eintrag/Eintraege als 'Approved' markiert." & vbCrLf & vbCrLf & _
           "Fuehren Sie 'SyncDecisions' aus, um diese Entscheidungen auf die Datenbank anzuwenden.", _
           vbInformation, "Alle genehmigen abgeschlossen"
    
    Exit Sub
    
ErrorHandler:
    Application.ScreenUpdating = True
    MsgBox "Fehler in ApproveAllPending: " & Err.Description, vbCritical, "Fehler"
End Sub

' Set "Declined" decision for all pending records on grossGeschichte sheet
Public Sub DeclineAllPending()
    On Error GoTo ErrorHandler
    
    Dim wsGross As Worksheet
    On Error Resume Next
    Set wsGross = ThisWorkbook.Worksheets("grossGeschichte")
    On Error GoTo ErrorHandler
    
    If wsGross Is Nothing Then
        MsgBox "grossGeschichte-Blatt nicht gefunden. Bitte fuehren Sie zuerst 'LoadPendingChanges' aus.", _
               vbExclamation, "Blatt nicht gefunden"
        Exit Sub
    End If
    
    ' Prompt for decline comment (will be applied to all records)
    Dim commonComment As String
    commonComment = InputBox("Geben Sie einen gemeinsamen Ablehnungskommentar fuer ALLE Eintraege ein:" & vbCrLf & _
                            "(Dieser Kommentar wird auf alle abgelehnten Eintraege angewendet)", _
                            "Ablehnungskommentar", "Massenablehnung durch Superadmin")
    
    ' User cancelled
    If commonComment = "" Then
        Dim confirmEmpty As VbMsgBoxResult
        confirmEmpty = MsgBox("Kein Kommentar eingegeben. Moechten Sie ohne Kommentar fortfahren?" & vbCrLf & _
                             "(Jeder Eintrag erfordert einen individuellen Kommentar waehrend SyncDecisions)", _
                             vbYesNo + vbQuestion, "Kein Kommentar")
        If confirmEmpty <> vbYes Then
            Exit Sub
        End If
    End If
    
    ' Confirm action
    Dim confirmResult As VbMsgBoxResult
    confirmResult = MsgBox("Dies setzt 'Declined' fuer ALLE ausstehenden Eintraege auf grossGeschichte." & vbCrLf & vbCrLf & _
                          "Sind Sie sicher, dass Sie alle Eintraege ablehnen moechten?", _
                          vbYesNo + vbQuestion, "Alle ablehnen bestaetigen")
    
    If confirmResult <> vbYes Then
        Exit Sub
    End If
    
    Application.ScreenUpdating = False
    
    Dim lastRow As Long
    lastRow = wsGross.Cells(wsGross.Rows.Count, 1).End(xlUp).Row
    
    If lastRow < 4 Then
        MsgBox "Keine ausstehenden Eintraege auf grossGeschichte gefunden.", vbInformation, "Keine Eintraege"
        Application.ScreenUpdating = True
        Exit Sub
    End If
    
    Dim count As Long
    count = 0
    
    ' Process "Ist" rows (every 3rd row starting from row 4: 4, 7, 10, ...)
    Dim r As Long
    For r = 4 To lastRow Step 3
        ' Check if row has data (RecordID in column AE should be present)
        If Trim(CStr(wsGross.Cells(r, 31).Value)) <> "" Then
            ' Set Decision to "Declined" (column AC = 29)
            wsGross.Cells(r, 29).Value = "Declined"
            
            ' Set Decline Comment if provided (column AD = 30)
            If commonComment <> "" Then
                wsGross.Cells(r, 30).Value = commonComment
            End If
            
            count = count + 1
        End If
    Next r
    
    Application.ScreenUpdating = True
    
    MsgBox count & " Eintrag/Eintraege als 'Declined' markiert." & vbCrLf & vbCrLf & _
           "Fuehren Sie 'SyncDecisions' aus, um diese Entscheidungen auf die Datenbank anzuwenden.", _
           vbInformation, "Alle ablehnen abgeschlossen"
    
    Exit Sub
    
ErrorHandler:
    Application.ScreenUpdating = True
    MsgBox "Fehler in DeclineAllPending: " & Err.Description, vbCritical, "Fehler"
End Sub

' Clear all decisions on grossGeschichte sheet (reset to empty)
Public Sub ClearAllDecisions()
    On Error GoTo ErrorHandler
    
    Dim wsGross As Worksheet
    On Error Resume Next
    Set wsGross = ThisWorkbook.Worksheets("grossGeschichte")
    On Error GoTo ErrorHandler
    
    If wsGross Is Nothing Then
        MsgBox "grossGeschichte-Blatt nicht gefunden.", vbExclamation, "Blatt nicht gefunden"
        Exit Sub
    End If
    
    ' Confirm action
    Dim confirmResult As VbMsgBoxResult
    confirmResult = MsgBox("Dies loescht ALLE Entscheidungen (Approved/Declined) auf grossGeschichte." & vbCrLf & vbCrLf & _
                          "Sind Sie sicher?", _
                          vbYesNo + vbQuestion, "Alle loeschen bestaetigen")
    
    If confirmResult <> vbYes Then
        Exit Sub
    End If
    
    Application.ScreenUpdating = False
    
    Dim lastRow As Long
    lastRow = wsGross.Cells(wsGross.Rows.Count, 1).End(xlUp).Row
    
    If lastRow < 4 Then
        MsgBox "Keine Eintraege auf grossGeschichte gefunden.", vbInformation, "Keine Eintraege"
        Application.ScreenUpdating = True
        Exit Sub
    End If
    
    Dim count As Long
    count = 0
    
    ' Process "Ist" rows (every 3rd row starting from row 4: 4, 7, 10, ...)
    Dim r As Long
    For r = 4 To lastRow Step 3
        ' Check if row has data (RecordID in column AE should be present)
        If Trim(CStr(wsGross.Cells(r, 31).Value)) <> "" Then
            ' Clear Decision (column AC = 29)
            wsGross.Cells(r, 29).Value = ""
            ' Clear Decline Comment (column AD = 30)
            wsGross.Cells(r, 30).Value = ""
            count = count + 1
        End If
    Next r
    
    Application.ScreenUpdating = True
    
    MsgBox count & " Entscheidung(en) geloescht.", vbInformation, "Alle loeschen abgeschlossen"
    
    Exit Sub
    
ErrorHandler:
    Application.ScreenUpdating = True
    MsgBox "Fehler in ClearAllDecisions: " & Err.Description, vbCritical, "Fehler"
End Sub
