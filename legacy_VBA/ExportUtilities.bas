Attribute VB_Name = "ExportUtilities"

'==========================
'   Code Section: modUtilities
'==========================
Option Explicit

Public Function GetUserRole() As String
    Dim ws As Worksheet
    Set ws = ThisWorkbook.Worksheets("Kartei")
    
    Dim roleValue As String
    roleValue = Trim(ws.Range("J1").value)
    If roleValue = "" Then
        roleValue = "Operator"
    End If
    
    GetUserRole = roleValue
End Function

Public Function ReadSheetIntoDictionary_ID(ByVal sh As Worksheet, _
                                           ByVal startRow As Long, _
                                           ByVal endCol As Long) As Scripting.Dictionary
    Dim dict As New Scripting.Dictionary
    Dim lastRow As Long
    lastRow = sh.Cells(sh.Rows.count, 1).End(xlUp).row
    
    Dim i As Long
    For i = startRow To lastRow
        
        Dim rowData As Variant
        rowData = sh.Range(sh.Cells(i, 1), sh.Cells(i, endCol)).Value2  ' 1..51
        
        ' Phone (col 7) and Mobile (col 8): Use .Text to preserve leading zeros.
        ' Excel may store phone numbers as numeric values with custom format (e.g. "00000000000"),
        ' displaying "0176..." but internally holding 176... as a number.
        ' Using .Value2 would return the numeric value, losing leading zeros.
        ' Using .Text returns the displayed string, preserving the format.
        ' This ensures the database receives "0176..." instead of "176...".
        If endCol >= 7 Then
            rowData(1, 7) = CStr(sh.Cells(i, 7).Text)
        End If
        If endCol >= 8 Then
            rowData(1, 8) = CStr(sh.Cells(i, 8).Text)
        End If
        
        ' col 48 => ID
        Dim vID As Variant
        vID = rowData(1, 48)
        
        Dim strID As String
        If IsEmpty(vID) Or IsNull(vID) Then
            strID = ""
        Else
            strID = CStr(vID)
        End If
        
        If strID <> "" Or RowHasData(rowData) Then
            dict(strID) = rowData
        End If
    Next i
    
    Set ReadSheetIntoDictionary_ID = dict
End Function

Public Function ReadLocalFormatsIntoDictionary_ID(ByVal sh As Worksheet, _
                                           ByVal startRow As Long, _
                                           ByVal endCol As Long) As Scripting.Dictionary
    Dim dict As New Scripting.Dictionary
    Dim lastRow As Long
    lastRow = sh.Cells(sh.Rows.count, 1).End(xlUp).row
    
    Dim i As Long
    For i = startRow To lastRow
        Dim formatData() As Variant
        Dim c As Integer
        
        ReDim formatData(1 To 1, 1 To 53)

        For c = 1 To 51
            Dim iColor As Long
            iColor = sh.Cells(i, c).Interior.Color
            formatData(1, c) = iColor
        Next c

        
        ' col 48 => ID
        Dim vID As Variant
        vID = sh.Cells(i, 48).value
        
        ' read FontColor for col 3
        Dim fc3 As Long
        fc3 = sh.Cells(i, 3).Font.Color
        
        ' read FontColor for col 18
        Dim fc18 As Long
        fc18 = sh.Cells(i, 18).Font.Color
        
        formatData(1, 52) = fc3
        formatData(1, 53) = fc18
        
        Dim strID As String
        If IsEmpty(vID) Or IsNull(vID) Then
            strID = ""
        Else
            strID = CStr(vID)
        End If
        
        If strID <> "" Then
            dict(strID) = formatData
        End If
    Next i
    
    Set ReadLocalFormatsIntoDictionary_ID = dict
End Function

Private Function RowHasData(ByVal rowData As Variant) As Boolean
    Dim c As Long
    For c = 1 To UBound(rowData, 2)
        If Not (IsEmpty(rowData(1, c)) Or IsNull(rowData(1, c))) Then
            RowHasData = True
            Exit Function
        End If
    Next c
    RowHasData = False
End Function

Public Function FindChangedIDs(ByVal dictLocal As Scripting.Dictionary, _
                               ByVal dictOriginal As Scripting.Dictionary) As Collection
    Dim result As New Collection
    Dim k As Variant
    
    For Each k In dictLocal.Keys
        Dim arrLocal As Variant
        arrLocal = dictLocal(k)
        
        If dictOriginal.exists(k) Then
            Dim arrOrig As Variant
            arrOrig = dictOriginal(k)
            
            ' Compare columns 1..47,49..51 (ignore col 48 => ID)
            If AreArraysDifferent_ID(arrLocal, arrOrig, 1, 47) Or _
               AreArraysDifferent_ID(arrLocal, arrOrig, 49, 51) Then
                result.Add k
            End If
        Else
            ' new ID => changed
            result.Add k
        End If
    Next k
    
    Set FindChangedIDs = result
End Function

Private Function AreArraysDifferent_ID(ByVal arr1 As Variant, _
                                       ByVal arr2 As Variant, _
                                       ByVal startCol As Long, _
                                       ByVal endCol As Long) As Boolean
    Dim c As Long
    For c = startCol To endCol
        If Not IsSameValue_ID(arr1(1, c), arr2(1, c)) Then
            AreArraysDifferent_ID = True
            Exit Function
        End If
    Next c
    AreArraysDifferent_ID = False
End Function

Private Function IsSameValue_ID(ByVal v1 As Variant, ByVal v2 As Variant) As Boolean
    If (IsEmpty(v1) And IsEmpty(v2)) Or (IsNull(v1) And IsNull(v2)) Then
        IsSameValue_ID = True
    ElseIf Trim(CStr(v1)) = Trim(CStr(v2)) Then
        IsSameValue_ID = True
    Else
        IsSameValue_ID = False
    End If
End Function

Public Function MaxID(ByVal shOriginal As Worksheet) As Long
    Dim lastRow As Long
    lastRow = shOriginal.Cells(shOriginal.Rows.count, 1).End(xlUp).row
    
    Dim rng As Range
    Set rng = shOriginal.Range("AV3:AV" & lastRow)
    MaxID = Application.WorksheetFunction.Max(rng)

End Function

Public Function UpdateLocalSheetRowByID(ByVal sh As Worksheet, ByVal shOriginal As Worksheet, _
                                        ByVal strID As String, ByVal rowData As Variant, ByVal maxIDOriginal As Long) As String
    Dim lastRow As Long
    lastRow = sh.Cells(sh.Rows.count, 1).End(xlUp).row

    Application.ScreenUpdating = False
    Application.Calculation = xlCalculationManual
    
    Dim r As Long
    For r = 3 To lastRow
        Dim checkID As String
        checkID = CStr(sh.Cells(r, 48).value)
        If checkID = strID Then
            sh.Cells(r, 49).value = rowData(1, 49)
            sh.Cells(r, 50).value = rowData(1, 50)
            sh.Cells(r, 51).value = rowData(1, 51)
            
            ' Process history updates based on user role
            Dim userRole As String
            userRole = GetCurrentUserRole()
            
            If userRole = "Admin" Then
                Dim dataYear As Long
                dataYear = GetDataYear()
                
                Dim currentM As Long
                currentM = Month(Date)
                Dim currentY As Long
                currentY = Year(Date)
                
                Dim forbiddenEndCol As Integer
                Dim forbiddenStartCol As Integer
                
                ' Determine range of past months that require "Ruck:" prefix
                If currentY > dataYear Then
                    forbiddenEndCol = 32
                    forbiddenStartCol = 21
                ElseIf currentY < dataYear Then
                    ' Future year, no past months
                    forbiddenEndCol = 0
                    forbiddenStartCol = 1
                Else
                    forbiddenEndCol = 20 + (currentM - 1)
                    forbiddenStartCol = 21
                End If
                
                Dim dataChanged As Boolean
                dataChanged = False
                
                Dim hasPastMonthChanges As Boolean
                hasPastMonthChanges = False
                
                ' Check if there are changes in past months (requiring "RUCK:" prefix)
                If Not IsOperatorAllowedToChange(dataYear) And forbiddenStartCol <= forbiddenEndCol Then
                    Dim i As Integer
                    Dim mo As Integer
                    Dim changesArr() As Integer
                    ReDim changesArr(1 To 12)
                    
                    ' Array to track ALL changed months (both past and current/future)
                    Dim allChangedArr() As Integer
                    ReDim allChangedArr(1 To 12)
                    
                    ' Use Export_HistoryBuilder to detect past month changes
                    Dim pastMonthsArr() As Integer
                    hasPastMonthChanges = Export_HistoryBuilder.HasPastMonthChanges( _
                        sh, shOriginal, r, forbiddenStartCol, forbiddenEndCol, pastMonthsArr)
                    
                    ' Also get all changed months for Notitzen display
                    Dim hasAnyMonthChanges As Boolean
                    hasAnyMonthChanges = Export_HistoryBuilder.GetAllChangedMonths(sh, shOriginal, r, allChangedArr)
                    
                    ' Copy past months to changesArr for tracking
                    For i = 1 To 12
                        changesArr(i) = pastMonthsArr(i)
                    Next i
                    
                    If hasPastMonthChanges Then
                        dataChanged = True
                        
                        ' Build RUCK history entry using new format
                        Dim ruckEntry As String
                        Dim ruckHasChanges As Boolean
                        Dim ruckMonthsArr() As Integer
                        
                        ruckEntry = Export_HistoryBuilder.BuildHistoryEntry( _
                            sh, shOriginal, r, True, "", ruckHasChanges, ruckMonthsArr)
                        
                        ' Store original history length for potential rollback
                        Dim originalHistoryLen As Long
                        originalHistoryLen = Len(sh.Cells(r, 52).value)
                        
                        ' Temporarily append entry (without comment yet)
                        Dim tempEntry As String
                        tempEntry = ruckEntry
                        
                        ' Remove trailing separator and empty comment for now
                        If Right(tempEntry, Len(Export_HistoryBuilder.HD_SESSION)) = Export_HistoryBuilder.HD_SESSION Then
                            tempEntry = Left(tempEntry, Len(tempEntry) - Len(Export_HistoryBuilder.HD_SESSION))
                        End If
                        
                        ' Get user comment and finalize history entry
                        Dim Notitz As String
                        
                        ' Check if bulk comment mode is active
                        If BulkComment.IsBulkCommentModeActive() Then
                            ' Use bulk comment instead of prompting user
                            Notitz = BulkComment.GetBulkCommentText()
                            BulkComment.IncrementBulkCommentCounter
                        Else
                            ' Normal flow: prompt user for comment via Notitzen
                            Call CreateOrClearNotitzenSheet
                            ' Show all changed months in Notitzen window
                            Call FillNotitzenSheet(allChangedArr, r)
                            
                            ' Check if user canceled the input
                            If Notitzen.UserCanceled Then
                                ' User canceled, exit the function without updating
                                Application.Calculation = xlCalculationAutomatic
                                Application.ScreenUpdating = True
                                UpdateLocalSheetRowByID = ""
                                Exit Function
                            End If
                            
                            Notitz = ThisWorkbook.Worksheets("Notitzen").Cells(5, 17).value
                        End If
                        
                        ' Replace empty comment placeholder with actual comment
                        Dim emptyComment As String
                        emptyComment = Export_HistoryBuilder.HD_COMMENT_START & Export_HistoryBuilder.HD_COMMENT_END
                        
                        If InStr(ruckEntry, emptyComment) > 0 Then
                            ruckEntry = Replace(ruckEntry, emptyComment, _
                                Export_HistoryBuilder.HD_COMMENT_START & Notitz & Export_HistoryBuilder.HD_COMMENT_END)
                        End If
                        
                        ' Append finalized entry to history
                        sh.Cells(r, 52).value = Export_HistoryBuilder.AppendHistoryEntry( _
                            CStr(sh.Cells(r, 52).value), ruckEntry)
                    End If
                End If
                
                ' Process all other changes (non-month fields if Ruck was processed, or all fields otherwise)
                ' CRITICAL FIX: If Ruck block was processed (dataChanged=True), DON'T call UpdateHistoryString
                ' because ruckEntry already contains ALL changes (months + other fields)
                ' Only call UpdateHistoryString if no Ruck block was processed
                Dim historyResult As String
                If dataChanged Then
                    ' Ruck block already processed ALL fields including non-month ones
                    ' History is already complete in sh.Cells(r, 52).value
                    ' Just return the current value to signal success
                    historyResult = sh.Cells(r, 52).Value
                Else
                    ' No Ruck block, normal flow with Notitzen for all changes
                    historyResult = UpdateHistoryString(sh, shOriginal, r, strID, maxIDOriginal, dataChanged, _
                                                        requestNotitzen:=True, isRuckBlock:=False)
                End If
                
                ' Check if user canceled (empty string returned)
                If historyResult = "" Then
                    ' User canceled, don't update AZ, return empty to signal cancellation
                    Application.Calculation = xlCalculationAutomatic
                    Application.ScreenUpdating = True
                    UpdateLocalSheetRowByID = ""
                    Exit Function
                End If
                
                ' Update successful, write new history
                sh.Cells(r, 52).Value = historyResult
                UpdateLocalSheetRowByID = historyResult
            Else
                ' Operator: process standard history updates
                Dim historyResultOp As String
                historyResultOp = UpdateHistoryString(sh, shOriginal, r, strID, maxIDOriginal, False, _
                                                      requestNotitzen:=True, isRuckBlock:=False)
                
                ' Check if user canceled
                If historyResultOp = "" Then
                    Application.Calculation = xlCalculationAutomatic
                    Application.ScreenUpdating = True
                    UpdateLocalSheetRowByID = ""
                    Exit Function
                End If
                
                sh.Cells(r, 52).Value = historyResultOp
                UpdateLocalSheetRowByID = historyResultOp
            End If
            Exit For
        End If
    Next r
    
    Application.Calculation = xlCalculationAutomatic
    Application.ScreenUpdating = True

End Function

Function SafeCompare(cell1 As Range, cell2 As Range) As Boolean
    If IsError(cell1.value) And IsError(cell2.value) Then
        ' Both cells have errors. We will consider them as equal
        SafeCompare = True
    ElseIf IsError(cell1.value) Or IsError(cell2.value) Then
        ' Only one cell have error. We will consider them as different
        SafeCompare = False
    Else
        ' Both cells are error free, we perform a normal comparison
        SafeCompare = (cell1.value = cell2.value)
    End If
End Function

Public Sub RebuildKarteiOriginal()
    Dim wsKartei As Worksheet
    Set wsKartei = ThisWorkbook.Worksheets("Kartei")
    
    Dim pseudoName As String
    pseudoName = Trim(wsKartei.Range("J1").value)
    If pseudoName = "" Then
        pseudoName = "Unknown_Operator"
    End If
    
    Dim lastRow As Long
    lastRow = wsKartei.Cells(wsKartei.Rows.count, 1).End(xlUp).row
    
    Dim i As Long
    For i = 3 To lastRow
        Dim valAW As Variant, valAX As Variant, valAY As Variant
        valAW = wsKartei.Cells(i, 49).value
        valAX = wsKartei.Cells(i, 50).value
        valAY = wsKartei.Cells(i, 51).value
        
        If IsEmpty(valAW) And IsEmpty(valAX) And IsEmpty(valAY) Then
            wsKartei.Cells(i, 49).value = pseudoName
            wsKartei.Cells(i, 50).value = Date
            wsKartei.Cells(i, 51).value = Format(Time, "HH:MM")
        End If
    Next i
    
    Dim shOrig As Worksheet
    On Error Resume Next
    Set shOrig = ThisWorkbook.Worksheets("Kartei_Original")
    On Error GoTo 0
    
    If Not shOrig Is Nothing Then
        Application.DisplayAlerts = False
        shOrig.Visible = xlSheetVisible
        shOrig.Delete
        Application.DisplayAlerts = True
    End If
    
    wsKartei.Copy After:=wsKartei
    ActiveSheet.name = "Kartei_Original"
    ActiveSheet.Visible = xlSheetHidden
    
    ThisWorkbook.Worksheets("Kartei").Activate
End Sub

Public Sub ResetSheetView(ByVal ws As Worksheet, ByVal wsOrig As Worksheet)
    
    On Error Resume Next
        
    ' Verify, if there is filters applied on the sheet
    If ws.AutoFilterMode Then
        If ws.FilterMode Then
            ws.ShowAllData
        End If
    End If
    
    ' Depict all hidden rows
    ws.Rows.Hidden = False
    
    ' Depict all hidden columns
    ws.Columns.Hidden = False
    
    Dim lastRow, lastRowOrig As Long
    
    lastRow = ws.Cells(ws.Rows.count, 1).End(xlUp).row
    lastRowOrig = wsOrig.Cells(wsOrig.Rows.count, 1).End(xlUp).row
    
    Dim rng1 As Range
    Dim rngOrig As Range
    Dim sortColumn As Integer
    ' Include column BA (53) to keep status (PENDING/DECLINED) synchronized with row data
    Set rng1 = ws.Range("A2:BA" & lastRow)
    Set rngOrig = wsOrig.Range("A2:BA" & lastRowOrig)
    sortColumn = 48
    
    rngOrig.Sort Key1:=rngOrig.Columns(sortColumn), Order1:=xlAscending, Header:=xlYes
    rng1.Sort Key1:=rng1.Columns(sortColumn), Order1:=xlAscending, Header:=xlYes
        
    On Error GoTo 0
End Sub

Function UpdateHistoryString(sh As Worksheet, shOriginal As Worksheet, row As Long, _
                                    strID As String, maxIDOriginal As Long, _
                                    ByRef dataChanged As Boolean, _
                                    Optional ByVal requestNotitzen As Boolean = True, _
                                    Optional ByVal isRuckBlock As Boolean = False) As String
    ' Updated function using Export_HistoryBuilder for extended field tracking
    ' Tracks changes in: A, B, D, E, F, G, H, I, J, M, O, R, U-AF, AK, AL, AM
    
    Dim updateString As String
    Dim changesArr() As Integer
    ReDim changesArr(1 To 12)
    
    Dim localChanges As Boolean
    localChanges = False
    
    ' Always read existing history from AZ at start
    updateString = sh.Cells(row, 52).Value
    
    ' Check if this is an existing record (not new)
    If CLng(strID) <= maxIDOriginal Then
        
        ' Use Export_HistoryBuilder to build history entry for all tracked fields
        Dim historyEntry As String
        Dim hasChanges As Boolean
        hasChanges = False
        
        ' Build history entry using new format (skip if already processed in Ruck block)
        If Not dataChanged Then
            historyEntry = Export_HistoryBuilder.BuildHistoryEntry( _
                sh, shOriginal, row, isRuckBlock, "", hasChanges, changesArr)
            
            If hasChanges Then
                localChanges = True
                dataChanged = True
            End If
        Else
            ' Ruck block already processed months, but we still need to check other fields
            ' Build entry without months (they were already added in Ruck block)
            historyEntry = BuildNonMonthHistoryEntry(sh, shOriginal, row, hasChanges)
            
            If hasChanges Then
                localChanges = True
            End If
        End If
        
        ' Request comment via Notitzen only if requestNotitzen=True and localChanges exist
        If requestNotitzen And localChanges Then
            ' Check if bulk comment mode is active
            Dim Notitz As String
            
            If BulkComment.IsBulkCommentModeActive() Then
                ' Use bulk comment instead of prompting user
                Notitz = BulkComment.GetBulkCommentText()
                BulkComment.IncrementBulkCommentCounter
            Else
                ' Normal flow: prompt user for comment via Notitzen
                Call CreateOrClearNotitzenSheet
                Call FillNotitzenSheet(changesArr, row)
                
                ' Check if user canceled the input
                If Notitzen.UserCanceled Then
                    ' User canceled, return empty string to indicate cancellation
                    UpdateHistoryString = ""
                    Exit Function
                End If
                
                Notitz = ThisWorkbook.Worksheets("Notitzen").Cells(5, 17).Value
            End If
            
            ' Append history entry with comment
            If Len(historyEntry) > 0 Then
                ' Remove trailing session separator if present (we'll add it with comment)
                If Right(historyEntry, Len(Export_HistoryBuilder.HD_SESSION)) = Export_HistoryBuilder.HD_SESSION Then
                    historyEntry = Left(historyEntry, Len(historyEntry) - Len(Export_HistoryBuilder.HD_SESSION))
                End If
                
                ' Find and replace empty comment placeholder with actual comment
                Dim emptyComment As String
                emptyComment = Export_HistoryBuilder.HD_COMMENT_START & Export_HistoryBuilder.HD_COMMENT_END
                
                If InStr(historyEntry, emptyComment) > 0 Then
                    historyEntry = Replace(historyEntry, emptyComment, _
                        Export_HistoryBuilder.HD_COMMENT_START & Notitz & Export_HistoryBuilder.HD_COMMENT_END)
                End If
                
                updateString = updateString & historyEntry & Export_HistoryBuilder.HD_SESSION
            Else
                ' No structured entry, just add comment in old format for compatibility
                updateString = updateString & Export_HistoryBuilder.HD_COMMENT_START & Notitz & _
                               Export_HistoryBuilder.HD_COMMENT_END & Format(Date, "dd.mm.yyyy") & _
                               Export_HistoryBuilder.HD_SESSION
            End If
        ElseIf localChanges And Not requestNotitzen Then
            ' Changes exist but no comment requested (already handled in Ruck block)
            If Len(historyEntry) > 0 Then
                updateString = updateString & historyEntry
            End If
        End If
    End If
    
    ' Return updated history string or empty string on cancellation
    UpdateHistoryString = updateString
End Function

Private Function BuildNonMonthHistoryEntry(ByVal sh As Worksheet, _
                                           ByVal shOriginal As Worksheet, _
                                           ByVal row As Long, _
                                           ByRef hasChanges As Boolean) As String
    ' Builds history entry for non-month fields only (when months were already processed in Ruck block)
    ' Checks: A, B, D, E, F, G, H, I, J, M, O, R, AK, AL, AM
    
    Dim entryParts As Collection
    Set entryParts = New Collection
    
    hasChanges = False
    
    ' Column mappings for non-month fields
    Dim fieldCols As Variant
    Dim fieldTags As Variant
    
    fieldCols = Array(1, 2, 4, 5, 6, 7, 8, 9, 10, 13, 15, 18, 37, 38, 39)
    fieldTags = Array(Export_HistoryBuilder.HT_FAMILY_ID, Export_HistoryBuilder.HT_PARENT, _
                      Export_HistoryBuilder.HT_CHILD, Export_HistoryBuilder.HT_BIRTHDATE, _
                      Export_HistoryBuilder.HT_ADDRESS, Export_HistoryBuilder.HT_PHONE, _
                      Export_HistoryBuilder.HT_MOBILE, Export_HistoryBuilder.HT_EMAIL, _
                      Export_HistoryBuilder.HT_SUBJECT1, Export_HistoryBuilder.HT_PRICE1, _
                      Export_HistoryBuilder.HT_SUBJECT2, Export_HistoryBuilder.HT_PRICE2, _
                      Export_HistoryBuilder.HT_EXTRA1, Export_HistoryBuilder.HT_EXTRA2, _
                      Export_HistoryBuilder.HT_EXTRA3)
    
    Dim idx As Long
    For idx = LBound(fieldCols) To UBound(fieldCols)
        Dim col As Integer
        col = CInt(fieldCols(idx))
        
        If Not SafeCompare(sh.Cells(row, col), shOriginal.Cells(row, col)) Then
            Dim oldVal As String
            Dim newVal As String
            
            oldVal = SanitizeForHistory(CStr(GetCellValueOrEmpty(shOriginal.Cells(row, col))))
            newVal = SanitizeForHistory(CStr(GetCellValueOrEmpty(sh.Cells(row, col))))
            
            Dim fieldEntry As String
            fieldEntry = CStr(fieldTags(idx)) & "(" & oldVal & Export_HistoryBuilder.HD_VALUE & newVal & ")"
            
            entryParts.Add fieldEntry
            hasChanges = True
        End If
    Next idx
    
    ' Build result string
    If entryParts.Count = 0 Then
        BuildNonMonthHistoryEntry = ""
        Exit Function
    End If
    
    Dim result As String
    result = ""
    
    Dim i As Long
    For i = 1 To entryParts.Count
        If i > 1 Then
            result = result & Export_HistoryBuilder.HD_FIELD
        End If
        result = result & entryParts(i)
    Next i
    
    ' Add empty comment placeholder and date
    result = result & Export_HistoryBuilder.HD_COMMENT_START & Export_HistoryBuilder.HD_COMMENT_END & _
             Format(Date, "dd.mm.yyyy") & Export_HistoryBuilder.HD_SESSION
    
    BuildNonMonthHistoryEntry = result
End Function

Private Function SanitizeForHistory(ByVal value As String) As String
    ' Sanitizes a value for inclusion in history string
    Dim result As String
    result = value
    
    result = Replace(result, Export_HistoryBuilder.HD_VALUE, "~>")
    result = Replace(result, Export_HistoryBuilder.HD_FIELD, ",")
    result = Replace(result, Export_HistoryBuilder.HD_SESSION, "|")
    result = Replace(result, "(", "[")
    result = Replace(result, ")", "]")
    result = Replace(result, vbCrLf, " ")
    result = Replace(result, vbCr, " ")
    result = Replace(result, vbLf, " ")
    
    SanitizeForHistory = Trim(result)
End Function

Private Function GetCellValueOrEmpty(ByVal cell As Range) As Variant
    ' Gets cell value safely, converting errors and nulls to empty string
    On Error Resume Next
    If IsError(cell.Value) Then
        GetCellValueOrEmpty = ""
    ElseIf IsNull(cell.Value) Or IsEmpty(cell.Value) Then
        GetCellValueOrEmpty = ""
    Else
        GetCellValueOrEmpty = cell.Value
    End If
    On Error GoTo 0
End Function

Private Function PreserveHyphens(ByVal text As String) As String
    ' Helper function to preserve hyphens in text (no removal)
    PreserveHyphens = text
End Function

Sub SelectFolder()
    Dim folderPath As String
    Dim fD As FileDialog
    
    Set fD = Application.FileDialog(msoFileDialogFolderPicker)
    
    With fD
        .title = "Chouse the folder with relevant base"
        .InitialFileName = Application.DefaultFilePath ' initial folder
        If .Show = -1 Then ' if user press OK
            folderPath = .SelectedItems(1) ' extract folder path
            ThisWorkbook.Worksheets("Kartei").Range("X1").value = folderPath
        Else
            folderPath = "" ' if user press cansel path is empty
            ThisWorkbook.Worksheets("Kartei").Range("X1").value = folderPath
        End If
    End With
    
    ' clear object
    Set fD = Nothing
    
End Sub

' ============================================================
' PHONE COLUMN ENFORCEMENT HELPERS
' Ensure phone columns (7=Tel., 8=Handy) are always displayed as TEXT
' to prevent scientific notation display (e.g. "1,76E+13")
' ============================================================

' Enforces TEXT format on phone columns and normalizes any scientific-notation values.
' Should be called after any operation that writes data to Kartei sheet.
'
' @param ws - The Kartei worksheet
' @param startRow - First data row (typically 3)
Public Sub EnforcePhoneColumnsAsText(ByVal ws As Worksheet, ByVal startRow As Long)
    On Error Resume Next
    
    ' Determine last row using ID column (48) as anchor
    Dim lastRow As Long
    lastRow = ws.Cells(ws.Rows.Count, 48).End(xlUp).Row
    
    ' Also check column A in case ID column is empty
    Dim lastRowA As Long
    lastRowA = ws.Cells(ws.Rows.Count, 1).End(xlUp).Row
    If lastRowA > lastRow Then lastRow = lastRowA
    
    If lastRow < startRow Then Exit Sub
    
    ' Pre-format phone columns as TEXT
    ws.Range(ws.Cells(startRow, 7), ws.Cells(lastRow, 7)).NumberFormat = "@"
    ws.Range(ws.Cells(startRow, 8), ws.Cells(lastRow, 8)).NumberFormat = "@"
    
    ' Convert any existing numeric/scientific cells to normalized text strings
    Dim r As Long
    Dim val7 As String, val8 As String
    Dim norm7 As String, norm8 As String
    
    For r = startRow To lastRow
        ' Use .Text to capture what the user sees (including scientific notation display)
        val7 = ws.Cells(r, 7).Text
        val8 = ws.Cells(r, 8).Text
        
        ' Normalize using phone_Normalize module
        norm7 = phone_Normalize.NormalizePhoneText(val7)
        norm8 = phone_Normalize.NormalizePhoneText(val8)
        
        ' Write back if value needs correction
        If norm7 <> val7 And Len(norm7) > 0 Then
            ws.Cells(r, 7).Value = norm7
        End If
        If norm8 <> val8 And Len(norm8) > 0 Then
            ws.Cells(r, 8).Value = norm8
        End If
    Next r
    
    On Error GoTo 0
End Sub

' Checks phone columns for any remaining scientific notation values.
' Returns the count of cells containing "E+" or "E-".
' Optionally shows a warning message if count > 0.
'
' @param ws - The Kartei worksheet
' @param startRow - First data row (typically 3)
' @param showWarning - If True, shows a MsgBox warning if issues found
' @return Long - Count of cells with scientific notation
Public Function CheckPhoneColumnsForScientific(ByVal ws As Worksheet, _
                                               ByVal startRow As Long, _
                                               Optional ByVal showWarning As Boolean = False) As Long
    On Error Resume Next
    
    Dim lastRow As Long
    lastRow = ws.Cells(ws.Rows.Count, 48).End(xlUp).Row
    Dim lastRowA As Long
    lastRowA = ws.Cells(ws.Rows.Count, 1).End(xlUp).Row
    If lastRowA > lastRow Then lastRow = lastRowA
    
    If lastRow < startRow Then
        CheckPhoneColumnsForScientific = 0
        Exit Function
    End If
    
    Dim countFound As Long
    countFound = 0
    
    Dim r As Long
    Dim cellText As String
    
    For r = startRow To lastRow
        ' Check column 7
        cellText = UCase$(ws.Cells(r, 7).Text)
        If InStr(1, cellText, "E+") > 0 Or InStr(1, cellText, "E-") > 0 Then
            countFound = countFound + 1
        End If
        
        ' Check column 8
        cellText = UCase$(ws.Cells(r, 8).Text)
        If InStr(1, cellText, "E+") > 0 Or InStr(1, cellText, "E-") > 0 Then
            countFound = countFound + 1
        End If
    Next r
    
    On Error GoTo 0
    
    If showWarning And countFound > 0 Then
        MsgBox "WARNUNG: " & countFound & " Telefon-Zelle(n) enthalten noch wissenschaftliche Notation (z.B. 1,76E+9)." & vbCrLf & vbCrLf & _
               "Empfehlung: Fuehren Sie 'EnforcePhoneColumnsAsText' aus oder ueberpruefen Sie die Daten manuell.", _
               vbExclamation, "Telefon-Formatierung"
    End If
    
    CheckPhoneColumnsForScientific = countFound
End Function

