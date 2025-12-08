Attribute VB_Name = "gesh_data_operations"
Function SetYesterdayDates(wb As Workbook) As Boolean
    ' Set yesterday's date in cells B1 and C1 of GrossGeschichte sheet
    Dim ws As Worksheet
    Dim yesterdayDate As Date
    Dim dateString As String
    
    On Error GoTo ErrorHandler
    
    Set ws = wb.Worksheets("GrossGeschichte")
    yesterdayDate = Date - 1
    dateString = Format(yesterdayDate, "dd.mm.yyyy")
    
    ws.Range("B1").Value = dateString
    ws.Range("C1").Value = dateString
    
    SetYesterdayDates = True
    Exit Function
    
ErrorHandler:
    MsgBox "Fehler beim Setzen der Datumsangaben: " & Err.Description, _
           vbCritical, "Datumsfehler"
    SetYesterdayDates = False
End Function

Function PrepareGesternGeshichteSheet(wb As Workbook) As Boolean
    ' Create or prepare GesternGeshichte sheet
    Dim ws As Worksheet
    Dim sheetExists As Boolean
    
    On Error Resume Next
    Set ws = wb.Worksheets("GesternGeshichte")
    sheetExists = (Err.Number = 0)
    On Error GoTo ErrorHandler
    
    If sheetExists Then
        ' Sheet exists, clear filters and content
        ws.Activate
        
        ' Remove any autofilters
        If ws.AutoFilterMode Then
            ws.AutoFilterMode = False
        End If
        
        ' Clear all content and formats
        ws.Cells.Clear
        ws.Cells.ClearFormats
    Else
        ' Create new sheet
        Set ws = wb.Worksheets.Add
        ws.name = "GesternGeshichte"
    End If
    
    PrepareGesternGeshichteSheet = True
    Exit Function
    
ErrorHandler:
    MsgBox "Fehler beim Vorbereiten des Blattes GesternGeshichte: " & Err.Description, _
           vbCritical, "Blattfehler"
    PrepareGesternGeshichteSheet = False
End Function

Function CopyDataToGesternGeshichte(sourceWb As Workbook, targetWb As Workbook) As Boolean
    ' Copy data from GrossGeschichte to GesternGeshichte
    Dim sourceWs As Worksheet
    Dim targetWs As Worksheet
    Dim lastRow As Long
    Dim copyRange As Range
    Dim i As Long
    
    On Error GoTo ErrorHandler
    
    Set sourceWs = sourceWb.Worksheets("GrossGeschichte")
    Set targetWs = targetWb.Worksheets("GesternGeshichte")
    
    ' Find last row with data in column S (starting from row 2)
    lastRow = FindLastRowInColumnS(sourceWs)
    
    If lastRow < 2 Then
        MsgBox "Keine Daten zum Kopieren gefunden.", vbInformation, "Information"
        CopyDataToGesternGeshichte = True
        Exit Function
    End If
    
    ' Copy data from row 2 to last row
    Set copyRange = sourceWs.Range("2:" & lastRow)
    
    ' Copy the entire rows (data, formats, row heights)
    copyRange.Copy
    targetWs.Range("A2").PasteSpecial Paste:=xlPasteAll
    
    ' Copy column widths
    CopyColumnWidths sourceWs, targetWs
    
    ' Clear clipboard
    Application.CutCopyMode = False
    
    ' Set up autofilter on row 2 and freeze panes
    SetupAutoFilterAndFreezePanes targetWs, lastRow
    
    CopyDataToGesternGeshichte = True
    Exit Function
    
ErrorHandler:
    MsgBox "Fehler beim Kopieren der Daten: " & Err.Description, _
           vbCritical, "Kopierfehler"
    CopyDataToGesternGeshichte = False
End Function

Function FindLastRowInColumnS(ws As Worksheet) As Long
    ' Find last row with data in column S (considering cell color)
    Dim i As Long
    Dim lastRow As Long
    
    lastRow = 1 ' Start with 1, will be updated if data found
    
    ' Start from row 2 and go down
    For i = 2 To ws.Rows.Count
        ' Check if cell in column S is empty and has no background color
        If ws.Range("S" & i).Value = "" And ws.Range("S" & i).Interior.colorIndex = xlNone Then
            ' Found empty cell without color - this is our stopping point
            Exit For
        Else
            ' Cell has data or color, update last row
            lastRow = i
        End If
        
        ' Safety check to avoid infinite loop
        If i > 100000 Then Exit For
    Next i
    
    FindLastRowInColumnS = lastRow
End Function

Sub CopyColumnWidths(sourceWs As Worksheet, targetWs As Worksheet)
    ' Copy column widths from source to target worksheet
    Dim col As Long
    Dim lastCol As Long
    
    On Error Resume Next
    
    ' Find last used column in source sheet
    lastCol = sourceWs.UsedRange.Columns.Count + sourceWs.UsedRange.Column - 1
    
    ' Copy column widths
    For col = 1 To lastCol
        targetWs.Columns(col).ColumnWidth = sourceWs.Columns(col).ColumnWidth
    Next col
    
    On Error GoTo 0
End Sub

Sub SetupAutoFilterAndFreezePanes(ws As Worksheet, lastRow As Long)
    ' Set up autofilter for entire data range and freeze panes for first two rows
    Dim lastCol As Long
    Dim filterRange As Range
    
    On Error Resume Next
    
    ' Activate the worksheet
    ws.Activate
    
    ' Remove any existing autofilter
    If ws.AutoFilterMode Then
        ws.AutoFilterMode = False
    End If
    
    ' Find last column with data in the used range
    If ws.UsedRange.Rows.Count > 0 Then
        lastCol = ws.UsedRange.Columns.Count + ws.UsedRange.Column - 1
    Else
        lastCol = 1
    End If
    
    ' Set autofilter for entire data range (from row 2 to last row with data)
    If lastRow >= 2 And lastCol > 0 Then
        Set filterRange = ws.Range(ws.Cells(2, 1), ws.Cells(lastRow, lastCol))
        filterRange.AutoFilter
    End If
    
    ' Freeze panes - select cell A3 to freeze first two rows
    ws.Range("A3").Select
    ActiveWindow.FreezePanes = True
    
    ' Select cell A1 to position cursor nicely
    ws.Range("A1").Select
    
    On Error GoTo 0
End Sub

