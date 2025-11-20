VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "DieseArbeitsmappe"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = True
Option Explicit

Dim frmZeigen As Boolean
Dim arrRows(1 To 100) As Integer

Dim START_ROW As Integer
Dim END_ROW As Integer
Dim WARNING_DAYS As Integer

Public Sub Workbook_Open()
    On Error GoTo ErrHandler
    
    ' Check if the workbook is opened in read-only mode (e.g., OneDrive/Shared scenario)
    If ThisWorkbook.ReadOnly = True Then
        MsgBox "The workbook is opened in read-only mode. No updates will be performed.", vbInformation, "Read-Only Mode"
        Exit Sub
    End If
    
    Dim ws As Worksheet
    Set ws = ThisWorkbook.Worksheets("Kartei")
'    ws.Unprotect password:="1212"
'    ws.Cells.Locked = False
    
    Call ImportKarteiAndFormat_Optimized
    Call ConvertAndFormatCellsOptimized
    ' Rebuild Kartei_Original
    Call RebuildKarteiOriginalSheet(ws)

    
'    ' Call the initialization for protection
'    Call InitializeWorkbook
    
    ' Optionally, one can add here a delay or a custom function call to ensure
    ' OneDrive sync has completed before proceeding.
    ' For example: WaitForOneDriveSync (not a built-in function, just a placeholder)
    
    Call StartHourlyMessage
    
    frmZeigen = False
    Call Initialis
    
    'If frmZeigen Then
        'frmListNachHilfe.Show
    'End If

    If frmZeigen Then
        Dim frm As frmListNachHilfe
        Set frm = New frmListNachHilfe
        ' The form's UserForm_Initialize event will populate the list
        frm.Show
    End If

    Exit Sub

ErrHandler:
    MsgBox "An error occurred during Workbook_Open: " & Err.Description, vbExclamation, "Error"
End Sub

Private Sub Workbook_BeforeClose(Cancel As Boolean)
    Dim ws As Worksheet
    Set ws = ThisWorkbook.Worksheets("Kartei")
    
    ws.Activate
    
    Dim lastRow As Integer
    lastRow = ws.Cells(ws.Rows.count, 46).End(xlUp).row
    
    If lastRow > 2 Then
        ws.Range(ws.Cells(3, 53), ws.Cells(lastRow, 53)).Copy
        ws.Range(ws.Cells(3, 46), ws.Cells(lastRow, 46)).PasteSpecial Paste:=xlPasteValues
        ws.Range(ws.Cells(3, 53), ws.Cells(lastRow, 53)).Clear
    End If
    
    Application.CutCopyMode = False
    ws.Cells(3, 1).Select
    
    Call StopHourlyMessage
    Call CompareAndSyncKartei_OnSave
End Sub

'===========================================
'  Code Section: Import from Access to Excel
'===========================================

Sub ImportKarteiAndFormat_Optimized()
    Dim db As DAO.Database
    Dim rs As DAO.Recordset

    ' 1) Single query to Access
    ' 2) Load all rows into arrays
    ' 3) One massive write for values
    ' 4) Single loop for cell-by-cell format (interior + font color)
    ' 5) Rebuild Kartei_Original

    Dim dbPath As String
    dbPath = ThisWorkbook.Worksheets("Kartei").Range("X1").value & "\Alarm\KindElternDaten_25_front.accdb"

    Dim wsKartei As Worksheet
    Set wsKartei = ThisWorkbook.Worksheets("Kartei")

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

    ' We expect fieldCount=105:
    '   0=ID
    '   1..102 => ValueX & InteriorColorX (51 * 2)
    '   103 => FontColor3
    '   104 => FontColor18
    ' If structure is different, adapt the code.

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

    ' Indices in arrRaw for each Value/Interior pair:
    ' ID is arrRaw(0, rec)
    ' Value1 => arrRaw(1, rec)
    ' InteriorColor1 => arrRaw(2, rec)
    ' Value2 => arrRaw(3, rec)
    ' InteriorColor2 => arrRaw(4, rec)
    ' ...
    ' ValueN => arrRaw(2*N-1, rec)
    ' InteriorColorN => arrRaw(2*N, rec)
    '
    ' FontColor3 => arrRaw(103, rec)
    ' FontColor18 => arrRaw(104, rec)

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
        .Range("A3").Resize(rowCount, 52).value = arrValues
    End With

    ' Now apply formats
    ' Turn off screen updating for speed
    Application.ScreenUpdating = False
    ' Optionally turn off calculation
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

        lastRow = .Cells(.Rows.count, "B").End(xlUp).row
        lastCol = 52

        .Range(.Cells(2, 1), .Cells(lastRow, lastCol)).Sort _
            Key1:=.Range("B2"), Order1:=xlAscending, Header:=xlYes
    End With


    Application.Calculation = xlCalculationAutomatic
    Application.ScreenUpdating = True

    ' Rebuild Kartei_Original
    'Call RebuildKarteiOriginalSheet(wsKartei)

    MsgBox "Imported " & rowCount & " rows into Kartei with formatting.", vbInformation
End Sub


Sub RebuildKarteiOriginalSheet(ByVal wsKartei As Worksheet)
    ' Deletes the Kartei_Original sheet if it exists,
    ' copies wsKartei, and renames the new sheet to Kartei_Original.

    Dim shOrig As Worksheet
    On Error Resume Next
    Set shOrig = ThisWorkbook.Worksheets("Kartei_Original")
    On Error GoTo 0
    
    ' Optimize performance
    Application.ScreenUpdating = False
    Application.Calculation = xlCalculationManual

    
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
    
    ' Restore application settings
    Application.Calculation = xlCalculationAutomatic
    Application.ScreenUpdating = True
End Sub


Private Sub ClearKarteiData(ByVal ws As Worksheet)
    Dim lastRow As Long
    lastRow = ws.Cells(ws.Rows.count, 1).End(xlUp).row
    If lastRow < 3 Then Exit Sub

    ws.Range("A3:AZ" & lastRow).ClearContents
    ws.Range("A3:AZ" & lastRow).Interior.ColorIndex = xlColorIndexNone
    ws.Range("A3:AZ" & lastRow).Font.ColorIndex = xlColorIndexAutomatic
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




'===========================================
'  Code Section: Initials
'===========================================


Private Sub Initialis()
    On Error GoTo ErrHandler

    Dim ws As Worksheet
    Dim i As Long
    Dim prufDat As Boolean
    
    START_ROW = 3
    END_ROW = 3000
    WARNING_DAYS = 15

    Set ws = ThisWorkbook.Worksheets("Kartei")

    For i = START_ROW To END_ROW
        If ws.Cells(i, 1).value = "" Then
            ' No more entries, exit loop
            Exit For
        ElseIf ws.Cells(i, 3).value <> "" Then
            prufDat = IsDate(ws.Cells(i, 3).value)
            If prufDat Then
                ' Check if the due date is within WARNING_DAYS and not yet marked
                If DateDiff("d", Now, ws.Cells(i, 3).value) < WARNING_DAYS And ws.Cells(i, 3).Font.ColorIndex <> 5 Then
                    frmZeigen = True
                    Exit For
                End If
            End If
        End If
    Next i

    Exit Sub

ErrHandler:
    MsgBox "An error occurred in Initialis: " & Err.Description, vbExclamation, "Error"
End Sub

' New InitializeWorkbook procedure
Sub InitializeWorkbook()
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

'    Call LockReportCol(ws)

    On Error GoTo 0
End Sub

Sub LockPastMonths(ws As Worksheet)
    On Error Resume Next ' Handling of possible errors

    Application.ScreenUpdating = False
    Application.Calculation = xlCalculationManual

    Dim currentMonth As Integer
    currentMonth = Month(Date)

    Dim monthColumns As Variant
    ' Assuming U=21, V=22, ..., AF=32
    monthColumns = Array("U", "V", "W", "X", "Y", "Z", "AA", "AB", "AC", "AD", "AE", "AF")

    Dim i As Integer
    For i = LBound(monthColumns) To UBound(monthColumns)
        If (i + 1) < currentMonth Then
            ' Lock columns for past months
            ws.Columns(monthColumns(i)).Locked = True
        Else
            ' Unlock columns for current and future months
            ws.Columns(monthColumns(i)).Locked = False
        End If
    Next i

'    ' Re-protect the sheet to apply changes
'    ws.Protect password:="1212", _
'               AllowFormattingCells:=True, _
'               AllowFormattingColumns:=True, _
'               AllowFormattingRows:=True, _
'               AllowSorting:=True, _
'               AllowFiltering:=True, _
'               AllowUsingPivotTables:=True

    Application.Calculation = xlCalculationAutomatic
    Application.ScreenUpdating = True

    On Error GoTo 0
End Sub

 Sub LockReportCol(ws As Worksheet)
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




