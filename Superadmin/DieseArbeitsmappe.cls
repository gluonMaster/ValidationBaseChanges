VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "DieseArbeitsmappe"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = True
'==========================
'   Module: DieseArbeitsmappe (Superadmin)
'   Purpose: Workbook events for KindElternDaten_25_Suprime.xlsm
'   
'   Key functionality:
'   - On open: Refresh Neu list with new records
'   - On close: Clear GrossGeschichte and Kartei sheets (from row 3 down)
'   - Activate Kartei sheet before closing
'   - Save without prompting the user
'==========================

Option Explicit

Private Sub Workbook_Open()
    On Error Resume Next
    ' Refresh Neu list to show new records (ID > LastSeenID)
    valid_NeuList.RefreshNeuList
    On Error GoTo 0
End Sub

Private Sub Workbook_BeforeClose(Cancel As Boolean)
    On Error GoTo ErrorHandler
    
    Application.ScreenUpdating = False
    Application.DisplayAlerts = False
    
    ' Clear GrossGeschichte sheet (from row 3 down)
    Call ClearSheetData("GrossGeschichte")
    
    ' Clear Kartei sheet (from row 3 down)
    Call ClearSheetData("Kartei")
    
    ' Activate Kartei sheet
    Call ActivateKarteiSheet
    
    ' Save the workbook without prompting
    ThisWorkbook.Save
    
    Application.DisplayAlerts = True
    Application.ScreenUpdating = True
    Exit Sub
    
ErrorHandler:
    Application.DisplayAlerts = True
    Application.ScreenUpdating = True
    ' Don't show error message on close - just proceed with closing
    Debug.Print "Workbook_BeforeClose Error: " & Err.Description
End Sub

' Clear data and formatting from row 3 down on specified sheet
Private Sub ClearSheetData(ByVal sheetName As String)
    On Error Resume Next
    
    Dim ws As Worksheet
    Set ws = ThisWorkbook.Worksheets(sheetName)
    
    If ws Is Nothing Then Exit Sub
    
    ' Find last row with data (check multiple columns for reliability)
    Dim lastRow As Long
    Dim lastRowA As Long, lastRowB As Long, lastRowC As Long
    
    lastRowA = ws.Cells(ws.Rows.Count, 1).End(xlUp).Row
    lastRowB = ws.Cells(ws.Rows.Count, 2).End(xlUp).Row
    lastRowC = ws.Cells(ws.Rows.Count, 3).End(xlUp).Row
    
    lastRow = Application.WorksheetFunction.Max(lastRowA, lastRowB, lastRowC)
    
    ' Only clear if there is data below row 2
    If lastRow > 2 Then
        ' Clear contents and formatting from row 3 to last row
        ' Use UsedRange columns to determine the extent
        Dim lastCol As Long
        lastCol = ws.UsedRange.Columns.Count
        If lastCol < 1 Then lastCol = 1
        
        With ws.Range(ws.Cells(3, 1), ws.Cells(lastRow, lastCol))
            .ClearContents
            .ClearFormats
            .Interior.ColorIndex = xlColorIndexNone
            .Font.ColorIndex = xlColorIndexAutomatic
            .Borders.LineStyle = xlNone
        End With
    End If
    
    ' Remove AutoFilter if present
    If ws.AutoFilterMode Then
        ws.AutoFilterMode = False
    End If
    
    On Error GoTo 0
End Sub

' Activate the Kartei sheet
Private Sub ActivateKarteiSheet()
    On Error Resume Next
    
    Dim ws As Worksheet
    Set ws = ThisWorkbook.Worksheets("Kartei")
    
    If Not ws Is Nothing Then
        ws.Activate
        ws.Cells(1, 1).Select
    End If
    
    On Error GoTo 0
End Sub
