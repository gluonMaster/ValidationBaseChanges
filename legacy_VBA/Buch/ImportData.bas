Attribute VB_Name = "ImportData"
Option Explicit

Dim frmZeigen As Boolean
Dim arrRows(1 To 100) As Integer

Dim START_ROW As Integer
Dim END_ROW As Integer
Dim WARNING_DAYS As Integer


Public Sub ImportFromBase()
    ' Legacy entry point - now delegates to ImportFromBase_Default
    ' Kept for backward compatibility with existing calls
    Call ImportFromBase_Default
End Sub

Public Sub ImportFromBase_Default()
    ' Default import mode: Access-based import for years 2024-2026 only
    ' Does not touch 2023 data
    
    On Error GoTo ErrHandler
    
    ' Check if the workbook is opened in read-only mode (e.g., OneDrive/Shared scenario)
    If ThisWorkbook.ReadOnly = True Then
        MsgBox "The workbook is opened in read-only mode. No updates will be performed.", vbInformation, "Read-Only Mode"
        Exit Sub
    End If
    
    Call ImportKarteiAndFormat_Optimized(2024)
    Call ImportKarteiAndFormat_Optimized(2025)
    Call ImportKarteiAndFormat_Optimized(2026)
        
    Exit Sub

ErrHandler:
    MsgBox "An error occurred during ImportFromBase_Default: " & Err.Description, vbExclamation, "Error"
End Sub

Public Sub ImportFromBase_Extended()
    ' Extended import mode: 2023-2026
    ' - 2023: copied from Excel file (KindElternDaten_23.xlsm) via CopyToBuch.Ubertragen
    ' - 2024-2026: imported from Access databases
    
    On Error GoTo ErrHandler
    
    ' Check if the workbook is opened in read-only mode (e.g., OneDrive/Shared scenario)
    If ThisWorkbook.ReadOnly = True Then
        MsgBox "The workbook is opened in read-only mode. No updates will be performed.", vbInformation, "Read-Only Mode"
        Exit Sub
    End If
    
    ' Import 2023 from Excel file (same logic as CopyToBuch uses)
    Call CopyToBuch.Ubertragen2023
    
    ' Import 2024-2026 from Access
    Call ImportKarteiAndFormat_Optimized(2024)
    Call ImportKarteiAndFormat_Optimized(2025)
    Call ImportKarteiAndFormat_Optimized(2026)
        
    Exit Sub

ErrHandler:
    MsgBox "An error occurred during ImportFromBase_Extended: " & Err.Description, vbExclamation, "Error"
End Sub

'===========================================
'  Code Section: Import from Access to Excel
'===========================================

Sub ImportKarteiAndFormat_Optimized(ByVal jahr As String)
    Dim db As DAO.Database
    Dim rs As DAO.Recordset

    ' 1) Single query to Access
    ' 2) Load all rows into arrays
    ' 3) One massive write for values
    ' 4) Single loop for cell-by-cell format (interior + font color)
    ' 5) Rebuild Kartei_Original

    Dim dbPath As String
    If jahr = 2024 Then
        dbPath = ThisWorkbook.Worksheets(jahr).Range("I1").Value & "\Alarm\KindElternDaten_24_front.accdb"
    ElseIf jahr = 2025 Then
        dbPath = ThisWorkbook.Worksheets(jahr).Range("I1").Value & "\Alarm\KindElternDaten_25_front.accdb"
    ElseIf jahr = 2026 Then
        dbPath = ThisWorkbook.Worksheets(jahr).Range("I1").Value & "\Alarm\KindElternDaten_26_front.accdb"
    End If

    Dim wsKartei As Worksheet
    Set wsKartei = ThisWorkbook.Worksheets(jahr)
    
    wsKartei.Activate

    ' Reset filters/hidden rows/columns before clearing
    DataHandler.ResetSheetView wsKartei

    ' Clear old data from row 3 down
    ClearKarteiData wsKartei

    ' Read from Access in one query
    Set rs = GetKarteiRecordset(dbPath, db)
    Debug.Print "RecordCount after MoveLast/MoveFirst = "; rs.RecordCount
    If rs.EOF Then
        rs.Close
        Set rs = Nothing
        db.Close
        MsgBox "No rows in tblKartei!", vbInformation
        Exit Sub
    End If

    ' Get all data into a 2D array: arrRaw(Fields, Records)
    Dim arrRaw As Variant
    rs.MoveLast
    rs.MoveFirst
    arrRaw = rs.GetRows(rs.RecordCount)   ' loads entire recordset at once

    rs.Close
    Set rs = Nothing
    db.Close

    ' Calculate how many rows we got
    Dim rowCount As Long
    rowCount = UBound(arrRaw, 2) + 1  ' +1 because array is 0-based
    ' We expect 105 fields: 0..104 => 105
    Dim fieldCount As Long
    fieldCount = UBound(arrRaw, 1) + 1

    If rowCount = 0 Then
        MsgBox "No rows in tblKartei!", vbInformation
        Exit Sub
    End If

    ' Build arrValues (rowCount x 52)
    Dim arrValues As Variant
    ReDim arrValues(1 To rowCount, 1 To 52)

    ' Build arrInterior (rowCount x 51) to store interior colors
    Dim arrInterior As Variant
    ReDim arrInterior(1 To rowCount, 1 To 51)

    ' We'll store font color only for col=3 and col=18
    Dim arrFontC As Variant    ' for col=3
    Dim arrFontR As Variant    ' for col=18

    ReDim arrFontC(1 To rowCount)
    ReDim arrFontR(1 To rowCount)

    Dim rec As Long
    For rec = 0 To rowCount - 1
        Dim rowDest As Long
        rowDest = rec + 1  ' we fill arrValues(1..rowCount)

        ' fill columns 1..51
        Dim col As Long
        For col = 1 To 51
            Dim idxValue As Long
            idxValue = 1 + (col - 1) * 2  ' 1..2..3.. means:
            ' col=1 => idxValue=1,  interior => idxValue+1=2
            ' col=2 => idxValue=3, interior => idxValue+1=4
            ' ...
            ' col=51 => idxValue=1+(51-1)*2 = 1+100=101 => interior=102

            arrValues(rowDest, col) = arrRaw(idxValue, rec)
            arrInterior(rowDest, col) = arrRaw(idxValue + 1, rec)
        Next col
        arrValues(rowDest, 52) = arrRaw(105, rec)

        arrFontC(rowDest) = arrRaw(103, rec)   ' FontColor3
        arrFontR(rowDest) = arrRaw(104, rec)   ' FontColor18
    Next rec

    ' Now we write arrValues to the sheet in one operation
    With wsKartei
        .Range("A3").Resize(rowCount, 52).Value = arrValues
    End With

    ' Now apply formats
    Application.ScreenUpdating = False
    Application.Calculation = xlCalculationManual
    
    Dim i As Long, c As Long
    For i = 1 To rowCount
        For c = 1 To 51
            wsKartei.Cells(i + 2, c).Interior.Color = arrInterior(i, c)
        Next c

        ' set font color in col=3
        wsKartei.Cells(i + 2, 3).Font.Color = arrFontC(i)
        ' set font color in col=18
        wsKartei.Cells(i + 2, 18).Font.Color = arrFontR(i)
    Next i
    
    With ActiveSheet
        Dim lastRow As Long
        Dim lastCol As Long

        lastRow = .Cells(.Rows.Count, "B").End(xlUp).row
        lastCol = 52

        .Range(.Cells(2, 1), .Cells(lastRow, lastCol)).Sort _
            Key1:=.Range("B2"), Order1:=xlAscending, Header:=xlYes
    End With
    
    Application.CutCopyMode = False
    
    With wsKartei.Sort
        .SortFields.Clear
        .SortFields.Add Key:=wsKartei.Range("B3:B" & lastRow), _
            SortOn:=xlSortOnValues, Order:=xlAscending, DataOption:=xlSortNormal
        .SortFields.Add Key:=wsKartei.Range("D3:D" & lastRow), _
            SortOn:=xlSortOnValues, Order:=xlAscending, DataOption:=xlSortNormal
        .SetRange wsKartei.Range("A3:AZ" & lastRow)
        .Header = xlNo
        .MatchCase = False
        .Orientation = xlTopToBottom
        .SortMethod = xlPinYin
        .Apply
    End With
    
    ' Ensure correct sheet is active before ConvertAndFormatCellsOptimized (uses ActiveSheet)
    wsKartei.Activate
    Call ConvertAndFormatCellsOptimized

    Application.Calculation = xlCalculationAutomatic
    Application.ScreenUpdating = True

End Sub


Public Sub ClearKarteiData(ByVal ws As Worksheet)
    Dim lastRow As Long
    lastRow = ws.Cells(ws.Rows.Count, 1).End(xlUp).row
    If lastRow < 3 Then Exit Sub

    ws.Range("A3:AZ" & lastRow).ClearContents
    ws.Range("A3:AZ" & lastRow).Interior.colorIndex = xlColorIndexNone
    ws.Range("A3:AZ" & lastRow).Font.colorIndex = xlColorIndexAutomatic
End Sub

Public Function GetKarteiRecordset(ByVal dbPath As String, _
                                   ByRef db As DAO.Database) As DAO.Recordset
    Dim engine As DAO.DBEngine
    Set engine = New DAO.DBEngine
    
    Dim wsDao As DAO.Workspace
    Set wsDao = engine.Workspaces(0)
    
    Set db = wsDao.OpenDatabase(dbPath)
    Dim rs As DAO.Recordset
    Set rs = db.OpenRecordset("SELECT * FROM tblKartei ORDER BY ID", dbOpenSnapshot)
    rs.MoveLast
    rs.MoveFirst
    
    Set GetKarteiRecordset = rs
    
End Function

' New EnableEditing procedure
Sub EnableEditing()
    Dim password As String
    password = InputBox("Enter password to enable editing:", "Unlock Editing")

    If password = "1212" Then
        On Error Resume Next ' Error handling

        Dim ws As Worksheet
        Set ws = ThisWorkbook.Sheets("Kartei")

        ' Unprotect the sheet
        ws.Unprotect password:="1212"

        ' Allow all columns to be editable
        ws.Cells.Locked = False

        ' Optionally, one can re-lock past months if needed
        'Call LockPastMonths(ws)

        MsgBox "Editing enabled.", vbInformation
    Else
        MsgBox "Incorrect password.", vbExclamation
    End If
End Sub

Public Sub ConvertAndFormatCellsOptimized()
    Dim ws As Worksheet
    Dim lastRow As Long
    Dim i As Long
    Dim rng As Range, rngAY As Range
    Dim decimalSeparator As String
    
    Application.Calculation = xlCalculationManual
    Application.ScreenUpdating = False

    ' Work on the active sheet
    Set ws = ActiveSheet
    
    ' Get the system decimal separator (either "," or ".")
    decimalSeparator = Application.International(xlDecimalSeparator)
    
    ' Find the last used row by checking column A (change if needed)
    lastRow = ws.Cells(ws.Rows.Count, 1).End(xlUp).row
    If lastRow < 3 Then
        MsgBox "No data to process from row 3 onwards.", vbInformation
        Exit Sub
    End If
    
    ' --- 1) Convert columns U through AF individually ---
    '    (U=21, V=22, ... AF=32)
    For i = 21 To 32
        Set rng = ws.Range(ws.Cells(3, i), ws.Cells(lastRow, i))
        
        ' Remove extra spaces
        rng.Replace What:=" ", Replacement:="", LookAt:=xlPart
        ' Remove non-breaking spaces and tabs before numeric conversion
        rng.Replace What:=Chr(160), Replacement:="", LookAt:=xlPart
        rng.Replace What:=vbTab, Replacement:="", LookAt:=xlPart
        
        ' Unify decimal separators so Excel can parse them correctly
        If decimalSeparator = "," Then
            ' Replace any '.' with ',' in case the data was typed with a dot
            rng.Replace What:=".", Replacement:=",", LookAt:=xlPart
        Else
            ' Replace any ',' with '.' if system uses dot
            rng.Replace What:=",", Replacement:=".", LookAt:=xlPart
        End If
        
        ' Now do TextToColumns on one column range
        rng.TextToColumns Destination:=rng.Cells(1, 1), DataType:=xlDelimited, _
                          TextQualifier:=xlDoubleQuote, ConsecutiveDelimiter:=False, _
                          Tab:=False, Semicolon:=False, Comma:=False, Space:=False, _
                          Other:=False
        
        ' Finally, set numeric format with two decimals, no leading zeros
        rng.NumberFormat = "0.00"
    Next i
        
    Application.Calculation = xlCalculationAutomatic
    Application.ScreenUpdating = True

    
'    MsgBox "Conversion and formatting completed successfully!", vbInformation
End Sub

