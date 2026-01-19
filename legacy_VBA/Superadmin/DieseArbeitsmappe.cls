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
'   Purpose: Workbook events for the multi-year Superadmin workbook
'   
'   MULTI-YEAR SUPPORT (2024, 2025, 2026):
'   - On open: Show Dashboard and refresh Neu list (2026 only)
'   - On close: Clear ALL year-specific sheets (Kartei24/25/26, grossGeschichte24/25/26)
'   - Activate Dashboard before closing
'   - Save without prompting the user
'
'   Protected sheets (NOT cleared on close):
'   - Dashboard (user instructions + year selection)
'   - DBConfig (database path configuration)
'   - NeuConfig (LastSeenID storage)
'   - Neu (new records list - kept for reference)
'   - GeschichteYY / GeschichteYY_Alle (history reports - view-only output)
'
'   Cleared sheets (transient data from pre_tblKartei):
'   - KarteiYY (pending import staging)
'   - grossGeschichteYY (decision sheet)
'==========================

Option Explicit

' Whether to auto-refresh Neu list on open (can be disabled if slow)
Private Const AUTO_REFRESH_NEU As Boolean = True

' Whether to clear Kartei/grossGeschichte sheets on close
' Set to False to preserve data across sessions (for debugging)
Private Const CLEAR_ON_CLOSE As Boolean = True

Private Sub Workbook_Open()
    On Error Resume Next
    
    ' Show (and create if missing) Dashboard
    valid_Dashboard.ShowDashboard
    
    ' Refresh Neu list (2026 only) to show new records
    If AUTO_REFRESH_NEU Then
        Dim status As String
        Dim dbPath As String
        dbPath = valid_YearConfig.TryGetDbPathForYearNoPrompt(26, status)
        
        ' Avoid prompting on workbook open. Refresh only if DB is already available.
        If status = "ok" Then
            valid_NeuList.RefreshNeuList
        End If
    End If
    
    On Error GoTo 0
End Sub

Private Sub Workbook_BeforeClose(Cancel As Boolean)
    On Error GoTo ErrorHandler
    
    If Not CLEAR_ON_CLOSE Then
        ' Skip clearing - just save and exit
        ThisWorkbook.Save
        Exit Sub
    End If
    
    Application.ScreenUpdating = False
    Application.DisplayAlerts = False
    
    ' Clear all year-specific transient sheets
    Call ClearAllYearSheets
    
    ' Activate Dashboard (or first sheet if Dashboard doesn't exist)
    Call ActivateDashboardOrFirst
    
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

' Clear all year-specific transient sheets (Kartei24/25/26, grossGeschichte24/25/26)
Private Sub ClearAllYearSheets()
    On Error Resume Next
    
    Dim years As Variant
    years = Array(24, 25, 26)
    
    Dim i As Long
    For i = LBound(years) To UBound(years)
        Dim year2 As Integer
        year2 = CInt(years(i))
        
        ' Clear KarteiYY sheet
        Call ClearSheetData("Kartei" & Format(year2, "00"))
        
        ' Clear grossGeschichteYY sheet
        Call ClearSheetData("grossGeschichte" & Format(year2, "00"))
    Next i
    
    ' Also clear legacy non-suffixed sheets if they exist (backwards compatibility)
    Call ClearSheetData("Kartei")
    Call ClearSheetData("GrossGeschichte")
    Call ClearSheetData("grossGeschichte")
    
    On Error GoTo 0
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

' Activate the Dashboard sheet (or first sheet if not available)
Private Sub ActivateDashboardOrFirst()
    On Error Resume Next
    
    Dim ws As Worksheet
    
    ' Try Dashboard first
    Set ws = ThisWorkbook.Worksheets("Dashboard")
    
    ' Fall back to first visible sheet
    If ws Is Nothing Then
        Dim i As Long
        For i = 1 To ThisWorkbook.Worksheets.Count
            If ThisWorkbook.Worksheets(i).Visible = xlSheetVisible Then
                Set ws = ThisWorkbook.Worksheets(i)
                Exit For
            End If
        Next i
    End If
    
    If Not ws Is Nothing Then
        ws.Activate
        ws.Cells(1, 1).Select
    End If
    
    On Error GoTo 0
End Sub
