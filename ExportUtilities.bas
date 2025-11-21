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
                
                ' Check if there are changes in past months (requiring "Ruck:" prefix)
                If Not IsOperatorAllowedToChange(dataYear) And forbiddenStartCol <= forbiddenEndCol Then
                    Dim i As Integer
                    Dim mo As Integer
                    Dim changesArr() As Integer
                    ReDim changesArr(1 To 12)
                    
                    For i = forbiddenStartCol To forbiddenEndCol
                        mo = i - 20
                        If Not SafeCompare(sh.Cells(r, i), shOriginal.Cells(r, i)) Then
                            hasPastMonthChanges = True
                            changesArr(mo) = 1
                            dataChanged = True
                        End If
                    Next i
                    
                    ' If past months changed, add "Ruck:" prefix
                    If hasPastMonthChanges Then
                        sh.Cells(r, 52).value = sh.Cells(r, 52).value & "Ruck: "
                        
                        For i = forbiddenStartCol To forbiddenEndCol
                            mo = i - 20
                            If changesArr(mo) = 1 Then
                                sh.Cells(r, 52).value = sh.Cells(r, 52).value & "Mnt." & CStr(mo) & ": War(" & _
                                                        CStr(shOriginal.Cells(r, i).value) & "); Ist(" & _
                                                        CStr(sh.Cells(r, i).value) & "). "
                            End If
                        Next i
                        
                        Call CreateOrClearNotitzenSheet
                        Call FillNotitzenSheet(changesArr, r)
                        
                        ' Check if user canceled the input
                        If Notitzen.UserCanceled Then
                            ' User canceled, exit the function without updating
                            Application.Calculation = xlCalculationAutomatic
                            Application.ScreenUpdating = True
                            UpdateLocalSheetRowByID = ""
                            Exit Function
                        End If
                        
                        Dim Notitz As String
                        Notitz = ThisWorkbook.Worksheets("Notitzen").Cells(5, 17).value
                        sh.Cells(r, 52).value = sh.Cells(r, 52).value & "/" & Notitz & "/ " & CStr(Date) & " || "
                    End If
                    
                    ' Clean up empty "Ruck: " if no changes were actually added
                    If Right(sh.Cells(r, 52).value, Len("Ruck: ")) = "Ruck: " Then
                        sh.Cells(r, 52).value = Left(sh.Cells(r, 52).value, Len(sh.Cells(r, 52).value) - Len("Ruck: "))
                    End If
                End If
                
                ' Process all other changes (future months, F/J/O fields) without "Ruck:" prefix
                sh.Cells(r, 52).value = UpdateHistoryString(sh, shOriginal, r, strID, maxIDOriginal, dataChanged)
                UpdateLocalSheetRowByID = sh.Cells(r, 52).value
            Else
                ' Operator: process standard history updates
                sh.Cells(r, 52).value = UpdateHistoryString(sh, shOriginal, r, strID, maxIDOriginal, False)
                UpdateLocalSheetRowByID = sh.Cells(r, 52).value
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
    Set rng1 = ws.Range("A2:AZ" & lastRow)
    Set rngOrig = wsOrig.Range("A2:AZ" & lastRowOrig)
    sortColumn = 48
    
    rngOrig.Sort Key1:=rngOrig.Columns(sortColumn), Order1:=xlAscending, Header:=xlYes
    rng1.Sort Key1:=rng1.Columns(sortColumn), Order1:=xlAscending, Header:=xlYes
        
    On Error GoTo 0
End Sub

Function UpdateHistoryString(sh As Worksheet, shOriginal As Worksheet, row As Long, _
                                    strID As String, maxIDOriginal As Long, _
                                    ByRef dataChanged As Boolean) As String
    Dim i As Integer
    Dim monat As Integer
    Dim updateString As String
    Dim changesArr() As Integer
    ReDim changesArr(1 To 12)
    
    ' If this row is not new
    If Not dataChanged And CInt(strID) <= maxIDOriginal Then
        updateString = sh.Cells(row, 52).value
        
        ' Check changes in months U-AF (columns 21-32)
        For i = 21 To 32
            monat = i - 20
            If Not SafeCompare(sh.Cells(row, i), shOriginal.Cells(row, i)) Then
                updateString = updateString & "Mnt." & CStr(monat) & ": War(" & _
                               CStr(shOriginal.Cells(row, i).value) & "); Ist(" & _
                               CStr(sh.Cells(row, i).value) & "). "
                changesArr(monat) = 1
                dataChanged = True
            End If
        Next i
        
        ' Check changes in Address (F, column 6)
        If Not SafeCompare(sh.Cells(row, 6), shOriginal.Cells(row, 6)) Then
            updateString = updateString & "Address: Was(" & _
                           PreserveHyphens(CStr(shOriginal.Cells(row, 6).value)) & "); Is(" & _
                           PreserveHyphens(CStr(sh.Cells(row, 6).value)) & "). "
            dataChanged = True
        End If
        
        ' Check changes in Subject1 (J, column 10)
        If Not SafeCompare(sh.Cells(row, 10), shOriginal.Cells(row, 10)) Then
            updateString = updateString & "Subject1: Was(" & _
                           PreserveHyphens(CStr(shOriginal.Cells(row, 10).value)) & "); Is(" & _
                           PreserveHyphens(CStr(sh.Cells(row, 10).value)) & "). "
            dataChanged = True
        End If
        
        ' Check changes in Subject2 (O, column 15)
        If Not SafeCompare(sh.Cells(row, 15), shOriginal.Cells(row, 15)) Then
            updateString = updateString & "Subject2: Was(" & _
                           PreserveHyphens(CStr(shOriginal.Cells(row, 15).value)) & "); Is(" & _
                           PreserveHyphens(CStr(sh.Cells(row, 15).value)) & "). "
            dataChanged = True
        End If
        
        ' If data has changed, request comment via Notitzen
        If dataChanged Then
            Call CreateOrClearNotitzenSheet
            Call FillNotitzenSheet(changesArr, row)
            
            ' Check if user canceled the input
            If Notitzen.UserCanceled Then
                ' User canceled, return empty string to indicate cancellation
                UpdateHistoryString = ""
                Exit Function
            End If
            
            Dim Notitz As String
            Notitz = ThisWorkbook.Worksheets("Notitzen").Cells(5, 17).value
            updateString = updateString & "/" & Notitz & "/ " & CStr(Date) & " || "
        End If
    End If
    
    ' Return a string with changes or an empty string
    UpdateHistoryString = updateString
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

