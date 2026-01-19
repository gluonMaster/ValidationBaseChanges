Attribute VB_Name = "ImportData"
Option Explicit

Dim frmZeigen As Boolean
Dim arrRows(1 To 100) As Integer

Dim START_ROW As Integer
Dim END_ROW As Integer
Dim WARNING_DAYS As Integer


Sub ImportFromBase()
'    On Error GoTo ErrHandler
    
    ' Check if the workbook is opened in read-only mode (e.g., OneDrive/Shared scenario)
    If ThisWorkbook.ReadOnly = True Then
        MsgBox "The workbook is opened in read-only mode. No updates will be performed.", vbInformation, "Read-Only Mode"
        Exit Sub
    End If
    
    Dim ws As Worksheet
    Set ws = ThisWorkbook.Worksheets("Kartei")
    
    Call ImportKarteiAndFormat_Optimized
    Call ConvertAndFormatCellsOptimized
        
    Exit Sub

'ErrHandler:
'    MsgBox "An error occurred during Workbook_Open: " & Err.Description, vbExclamation, "Error"
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
    dbPath = ThisWorkbook.Worksheets("Kartei").Range("I1").Value & "\Alarm\KindElternDaten_25_front.accdb"
    'C:\Users\Alla\OneDrive - Kinder- und Elternzentrum Kolibri e.V\Datenbank\2025
    'dbPath = "C:\Konst\2024\Kolibri\Valentina\Alla\Release\KindElternDaten_24_front.accdb"

    Dim wsKartei As Worksheet
    Set wsKartei = ThisWorkbook.Worksheets("Kartei")

    ' Clear old data from row 3 down
    ClearKarteiData wsKartei

    ' Read from Access in one query
    Set rs = GetKarteiRecordset(dbPath, db)
    Debug.Print "RecordCount after MoveLast/MoveFirst = "; rs.recordCount
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
    arrRaw = rs.GetRows(rs.recordCount)   ' loads entire recordset at once

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
        .Range("A3").Resize(rowCount, 52).Value = arrValues
    End With

    ' Now apply formats
    ' Turn off screen updating for speed
    Application.ScreenUpdating = False
    ' Optionally turn off calculation
    Application.Calculation = xlCalculationManual
    Dim i As Long, C As Long
    For i = 1 To rowCount
        For C = 1 To 51
            wsKartei.Cells(i + 2, C).Interior.Color = arrInterior(i, C)
        Next C

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


Public Sub ClearKarteiData(ByVal ws As Worksheet)
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


Sub SelectFolder()
    Dim folderPath As String
    Dim fD As FileDialog
    
    Set fD = Application.FileDialog(msoFileDialogFolderPicker)
    
    With fD
        .Title = "Chouse the folder with relevant base"
        .InitialFileName = Application.DefaultFilePath ' initial folder
        If .Show = -1 Then ' if user press OK
            folderPath = .SelectedItems(1) ' extract folder path
            ThisWorkbook.Worksheets("Kartei").Range("I1").Value = folderPath
        Else
            folderPath = "" ' if user press cansel path is empty
            ThisWorkbook.Worksheets("Kartei").Range("I1").Value = folderPath
        End If
    End With
    
    ' clear object
    Set fD = Nothing
    
End Sub



