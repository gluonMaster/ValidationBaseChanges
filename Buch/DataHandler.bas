Attribute VB_Name = "DataHandler"
' Find the last used row in the specified worksheet
Public Function FindLastRow(TabNameX As String) As Integer
    Dim ws As Worksheet
    Set ws = Worksheets(TabNameX)
  
    FindLastRow = ws.Cells(ws.Rows.Count, "A").End(xlUp).row
    If FindLastRow < 3 Then FindLastRow = 3
End Function

' Filter and copy data based on account number
Public Function FilterAndCopy(TabNameX As String, kontoNr As String) As Boolean
    Dim rng2 As Range
    Dim rng As Range
    Dim lastRow As Integer
    Dim result As Boolean
    
    result = False
    lastRow = FindLastRow(TabNameX)

    Application.ScreenUpdating = False
    Application.Calculation = xlCalculationManual
  
    Worksheets(TabNameX).Range("A3:AU" & lastRow).AutoFilter Field:=1, Criteria1:=kontoNr
       
    On Error Resume Next
    Set rng2 = Worksheets(TabNameX).AutoFilter.Range.Offset(1, 0).Resize(Worksheets(TabNameX).AutoFilter.Range.Rows.Count - 1, 1) _
         .SpecialCells(xlCellTypeVisible)
         
    Worksheets("DienstTab").Range("A1:AU1500").Clear
         
    If rng2 Is Nothing Then
        LogMissingAccountData TabNameX, kontoNr
    Else
        Set rng = Worksheets(TabNameX).AutoFilter.Range
        rng.Offset(1, 0).Resize(rng.Rows.Count - 1).Copy Destination:=Worksheets("DienstTab").Range("A1")
        result = True
    End If
    
    On Error GoTo 0
    
    Worksheets(TabNameX).Range("A3:AU" & lastRow).AutoFilter Field:=1
    Application.Calculation = xlCalculationAutomatic
    Application.ScreenUpdating = True
    
    FilterAndCopy = result
End Function

' Log information about missing account data
Private Sub LogMissingAccountData(TabNameX As String, kontoNr As String)
    Dim wsErrorLog As Worksheet
    Dim logRow As Long
    Dim activeSheetBefore As Worksheet
    
    ' Store the currently active sheet to return to it later
    Set activeSheetBefore = ActiveSheet
    
    On Error Resume Next
    Set wsErrorLog = ThisWorkbook.Worksheets("ErrorsLog")
    If Err.Number <> 0 Then
        ' If the sheet doesn't exist, don't try to log
        On Error GoTo 0
        Exit Sub
    End If
    On Error GoTo 0
    
    ' Find the last row in the ErrorsLog sheet without activating it
    logRow = wsErrorLog.Cells(wsErrorLog.Rows.Count, "A").End(xlUp).row + 1
    
    ' Write log entry
    wsErrorLog.Cells(logRow, "A").Value = Format(Now(), "yyyy-mm-dd hh:mm:ss")
    wsErrorLog.Cells(logRow, "B").Value = "Info: No records found for account " & kontoNr & " in sheet " & TabNameX & " (This is normal if the account wasn't active in this year)"
    
    ' Return to the original sheet without changing focus
    activeSheetBefore.Activate
End Sub

' Extract parent/payer information from the specified worksheet
Public Function ExtractParentInfo(TabNameX As String) As Integer
    Dim j As Long, jj As Long
    Dim dict As Object
    Dim dictDuplicates As Object
    Dim wsSource As Worksheet, wsDest As Worksheet, wsErrorLog As Worksheet
    Dim lastRow As Long, sourceLastRow As Long
    Dim cellValue As Variant, parentName As String, cleanParentName As String
    Dim duplicateFound As Boolean
    Dim logRow As Long
    
    ' Set references to source and destination worksheets
    Set wsSource = Worksheets(TabNameX)
    Set wsDest = Worksheets("DienstTab")
    Set wsErrorLog = Worksheets("ErrorsLog")
    
    ' Create dictionaries to store existing values and track duplicates
    Set dict = CreateObject("Scripting.Dictionary")
    Set dictDuplicates = CreateObject("Scripting.Dictionary")
    duplicateFound = False
    
    ' Find the last used row in DienstTab column A
    lastRow = wsDest.Cells(wsDest.Rows.Count, "A").End(xlUp).row
    
    ' Check if DienstTab is truly empty
    If lastRow = 1 And Trim(wsDest.Cells(1, "A").Value) = "" Then
        lastRow = 0
    End If
    
    ' Find the last row in source sheet
    sourceLastRow = wsSource.Cells(wsSource.Rows.Count, "A").End(xlUp).row
    
    ' Populate the dictionary with existing values from DienstTab column A
    For j = 1 To lastRow
        cellValue = Trim(wsDest.Cells(j, "A").Value)
        If cellValue <> "" Then
            parentName = Trim(wsDest.Cells(j, "B").Value)
            cleanParentName = CleanNameForComparison(parentName)
            
            If Not dict.Exists(cellValue) Then
                dict(cellValue) = cleanParentName ' Store cleaned parent name for duplicate detection
            Else
                ' Check if this is a real duplicate (different parent name)
                If dict(cellValue) <> cleanParentName Then
                    ' Track real duplicates that already exist in the destination
                    If Not dictDuplicates.Exists(cellValue) Then
                        dictDuplicates(cellValue) = "ID: " & cellValue & " linked to both '" & dict(cellValue) & "' and '" & cleanParentName & "'"
                        duplicateFound = True
                    End If
                End If
            End If
        End If
    Next j
    
    ' Initialize the destination row counter
    jj = lastRow + 1
    
    ' Use arrays to speed up data access (minimize worksheet access)
    Dim sourceData As Variant
    sourceData = wsSource.Range("A3:F" & sourceLastRow).Value
    
    ' Process data from arrays
    For j = 1 To UBound(sourceData, 1)
        ' Exit if we've reached an empty cell in column A
        If Trim(sourceData(j, 1)) <> "" Then
            cellValue = Trim(sourceData(j, 1))
            
            ' Check if the value in column D is "Zahlung" (with potential space before)
            If Trim(sourceData(j, 4)) = "Zahlung" Or Trim(sourceData(j, 4)) = " Zahlung" Then
                parentName = Trim(sourceData(j, 2))
                cleanParentName = CleanNameForComparison(parentName)
                
                ' Check if this ID already exists in the dictionary
                If Not dict.Exists(cellValue) Then
                    ' Add to destination sheet
                    wsDest.Cells(jj, "A").Value = cellValue
                    wsDest.Cells(jj, "B").Value = parentName
                    
                    ' Check if the next row exists in the array before accessing it
                    If j + 1 <= UBound(sourceData, 1) Then
                        wsDest.Cells(jj, "C").Value = sourceData(j + 1, 6)
                    End If
                    
                    ' Add to dictionary
                    dict(cellValue) = cleanParentName
                    jj = jj + 1
                Else
                    ' Check if this is a real duplicate (different parent name)
                    If dict(cellValue) <> cleanParentName Then
                        ' Track real duplicates
                        If Not dictDuplicates.Exists(cellValue) Then
                            dictDuplicates(cellValue) = "ID: " & cellValue & " linked to both '" & dict(cellValue) & "' and '" & cleanParentName & "'"
                            duplicateFound = True
                        End If
                    End If
                End If
            End If
        End If
    Next j
    
    ' If real duplicates were found, write them to ErrorsLog
    If duplicateFound Then
        ' Store the currently active sheet to return to it later
        Dim activeSheetBefore As Worksheet
        Set activeSheetBefore = ActiveSheet
        
        ' Find the last row in the ErrorsLog sheet without activating it
        logRow = wsErrorLog.Cells(wsErrorLog.Rows.Count, "A").End(xlUp).row
        If logRow = 1 And Trim(wsErrorLog.Cells(1, "A").Value) = "" Then
            logRow = 1
        Else
            logRow = logRow + 1
        End If
        
        ' Write header for this section
        wsErrorLog.Cells(logRow, "A").Value = "Duplicate IDs found in sheet " & TabNameX & " - " & Format(Now(), "yyyy-mm-dd hh:mm:ss")
        logRow = logRow + 1
        
        ' Write column headers
        wsErrorLog.Cells(logRow, "A").Value = "Account ID"
        wsErrorLog.Cells(logRow, "B").Value = "Issue Description"
        
        ' Format headers
        With wsErrorLog.Range("A" & logRow & ":B" & logRow)
            .Font.Bold = True
            .Interior.colorIndex = 15 ' Light gray background
        End With
        
        logRow = logRow + 1
        
        ' Write all duplicates to the log
        For Each cellValue In dictDuplicates.Keys
            wsErrorLog.Cells(logRow, "A").Value = cellValue
            wsErrorLog.Cells(logRow, "B").Value = dictDuplicates(cellValue)
            logRow = logRow + 1
        Next
        
        ' Add a blank row after the list
        logRow = logRow + 1
        
        ' Make sure the user stays on the original sheet
        activeSheetBefore.Activate
    End If
    
    ' Return the last parent row
    ExtractParentInfo = jj - 1
End Function

' Helper function to clean and normalize parent names for comparison
Private Function CleanNameForComparison(name As String) As String
    Dim result As String
    Dim i As Integer, char As String
    
    ' Convert to lowercase
    result = LCase(Trim(name))
    
    ' Remove punctuation and standardize spaces
    For i = 1 To Len(result)
        char = Mid(result, i, 1)
        
        ' Keep only letters and digits
        If (char >= "a" And char <= "z") Or (char >= "0" And char <= "9") Then
            CleanNameForComparison = CleanNameForComparison & char
        End If
    Next i
End Function

' Reset sheet view by clearing filters and unhiding rows/columns
Public Sub ResetSheetView(ByVal ws As Worksheet)
    On Error Resume Next
        
    ' Verify if there are filters applied on the sheet
    If ws.AutoFilterMode Then
        If ws.FilterMode Then
            ws.ShowAllData
        End If
    End If
    
    ' Unhide all hidden rows
    ws.Rows.Hidden = False
    
    ' Unhide all hidden columns
    ws.Columns.Hidden = False
        
    On Error GoTo 0
End Sub

' Clear ErrorsLog sheet
Public Sub ClearErrorsLog()
    Dim wsErrorLog As Worksheet
    
    On Error Resume Next
    Set wsErrorLog = ThisWorkbook.Worksheets("ErrorsLog")
    
    ' Check if ErrorsLog sheet exists
    If Err.Number <> 0 Then
        ' Create ErrorsLog sheet if it doesn't exist
        ThisWorkbook.Sheets.Add(After:=ThisWorkbook.Sheets(ThisWorkbook.Sheets.Count)).name = "ErrorsLog"
        Set wsErrorLog = ThisWorkbook.Worksheets("ErrorsLog")
        
        ' Format the sheet
        With wsErrorLog
            .Cells(1, "A").Value = "Date/Time"
            .Cells(1, "B").Value = "Description"
            .Range("A1:B1").Font.Bold = True
            .Range("A1:B1").Interior.colorIndex = 15
            .Columns("A:B").AutoFit
        End With
    End If
    On Error GoTo 0
    
    ' Store the currently active sheet to return to it later
    Dim activeSheetBefore As Worksheet
    Set activeSheetBefore = ActiveSheet
    
    ' Clear the existing content but keep the header row if it exists
    If wsErrorLog.Cells(1, "A").Value = "Date/Time" And wsErrorLog.Cells(1, "B").Value = "Description" Then
        ' If header exists, clear from row 2 down
        wsErrorLog.Range("A2:Z" & wsErrorLog.Cells(wsErrorLog.Rows.Count, "A").End(xlUp).row).Clear
    Else
        ' Otherwise clear everything and add header
        wsErrorLog.Cells.Clear
        wsErrorLog.Cells(1, "A").Value = "Date/Time"
        wsErrorLog.Cells(1, "B").Value = "Description"
        wsErrorLog.Range("A1:B1").Font.Bold = True
        wsErrorLog.Range("A1:B1").Interior.colorIndex = 15
    End If
    
    ' Format columns
    wsErrorLog.Columns("A:B").AutoFit
    
    ' Add initialization timestamp
    wsErrorLog.Cells(2, "A").Value = Format(Now(), "yyyy-mm-dd hh:mm:ss")
    wsErrorLog.Cells(2, "B").Value = "Form initialization - Error log cleared"
    
    ' Return to the original sheet
    activeSheetBefore.Activate
End Sub
