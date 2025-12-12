Attribute VB_Name = "kosten_Report"
'==========================
'   Kosten Report Module
'   Provides utilities for managing the Kosten report sheet
'   Orchestrates report building for multiple years
'==========================
Option Explicit

' ========================================
' Report Building Entry Points
' ========================================

Public Sub Kosten_BuildReport_AllYears()
    ' Builds the full Kosten report for years 2023-2026
    ' Uses data from GlobalVariables (nameEltern, Konto, RowFurTabelle)
    ' Delegates to internal BuildReportForYears with full year list
    
    Dim years(0 To 3) As String
    years(0) = "2023"
    years(1) = "2024"
    years(2) = "2025"
    years(3) = "2026"
    
    Call BuildReportForYears(years)
End Sub

Public Sub Kosten_BuildReport_RecentYears()
    ' Builds the Kosten report for years 2024-2026 only
    ' Excludes 2023 data
    ' Uses data from GlobalVariables (nameEltern, Konto, RowFurTabelle)
    
    Dim years(0 To 2) As String
    years(0) = "2024"
    years(1) = "2025"
    years(2) = "2026"
    
    Call BuildReportForYears(years)
End Sub

' ========================================
' Core Report Building Logic
' ========================================

Private Sub BuildReportForYears(ByRef yearsList() As String)
    ' Internal procedure to build report for specified years
    ' Handles: FindLastRow -> FilterAndCopy -> SumTarifMonat for each year
    ' Then applies final calculations and formatting
    
    On Error GoTo ErrorHandler
    
    ' Save current application state
    Dim oldScreenUpdating As Boolean
    Dim oldCalculation As XlCalculation
    Dim oldEnableEvents As Boolean
    
    oldScreenUpdating = Application.ScreenUpdating
    oldCalculation = Application.Calculation
    oldEnableEvents = Application.EnableEvents
    
    ' Disable for performance
    Application.ScreenUpdating = False
    Application.Calculation = xlCalculationManual
    Application.EnableEvents = False
    
    ' Clear report area before building
    Call Kosten_ClearReportArea
    
    ' Process each year
    Dim i As Long
    Dim yearStr As String
    Dim isFirstYear As Boolean
    
    For i = LBound(yearsList) To UBound(yearsList)
        yearStr = yearsList(i)
        isFirstYear = (i = LBound(yearsList))
        
        ' Check if worksheet exists for this year
        If Not WorksheetExists(yearStr) Then
            Debug.Print "Kosten_BuildReport: Worksheet '" & yearStr & "' not found, skipping."
            GoTo NextYear
        End If
        
        ' Find last row in the year's data
        GlobalVariables.letzteRow = DataHandler.FindLastRow(yearStr)
        
        ' Filter and copy data for the account
        Call DataHandler.FilterAndCopy(yearStr, GlobalVariables.Konto)
        
        ' Calculate and display tariff summaries
        ' First year uses vol=True to initialize RowFurTabelle, subsequent years use vol=False
        Call FinanceCalculations.SumTarifMonat(yearStr, isFirstYear, 8, _
            GlobalVariables.nameEltern, GlobalVariables.Konto, GlobalVariables.RowFurTabelle)
        
NextYear:
    Next i
    
    ' Apply final calculations and formatting
    Call ApplyFinalCalculations
    
Cleanup:
    ' Restore application state
    Application.EnableEvents = oldEnableEvents
    Application.Calculation = oldCalculation
    Application.ScreenUpdating = oldScreenUpdating
    Exit Sub
    
ErrorHandler:
    Debug.Print "Kosten_BuildReport error: " & Err.Description
    Resume Cleanup
End Sub

Private Sub ApplyFinalCalculations()
    ' Applies final calculations and formatting to the Kosten report
    ' Sets "Noch zu bezahlen" label and sum formula
    ' Formats the summary cells
    
    On Error Resume Next
    
    Dim wsKosten As Worksheet
    Set wsKosten = ThisWorkbook.Worksheets("Kosten")
    
    If wsKosten Is Nothing Then
        Debug.Print "ApplyFinalCalculations: Worksheet 'Kosten' not found."
        Exit Sub
    End If
    
    Dim summaryRow As Long
    Dim rowsToSum As Long
    
    summaryRow = GlobalVariables.RowFurTabelle + 2
    rowsToSum = GlobalVariables.RowFurTabelle - 3
    
    ' Set "Noch zu bezahlen" label in column O (15)
    wsKosten.Cells(summaryRow, 15).Value = "Noch zu bezahlen"
    
    ' Set sum formula in column S (19)
    wsKosten.Cells(summaryRow, 19).NumberFormat = "0.00"
    wsKosten.Cells(summaryRow, 19).FormulaR1C1 = "=SUM(R[-" & rowsToSum & "]C:R[-2]C)"
    
    ' Format header cell in row 5, column S
    wsKosten.Cells(5, "S").NumberFormat = "0.00"
    wsKosten.Cells(5, "S").Interior.ColorIndex = 40
    
    On Error GoTo 0
End Sub

' ========================================
' Sheet Clearing Utilities
' ========================================

Public Sub Kosten_ClearReportArea()
    ' Clears data and formats in the Kosten report area (A3:T300)
    ' Removes values, interior colors, and font colors
    ' No UI messages displayed
    
    On Error GoTo ErrorHandler
    
    Dim wsKosten As Worksheet
    On Error Resume Next
    Set wsKosten = ThisWorkbook.Worksheets("Kosten")
    On Error GoTo ErrorHandler
    
    If wsKosten Is Nothing Then
        Debug.Print "Kosten_ClearReportArea: Worksheet 'Kosten' not found."
        Exit Sub
    End If
    
    Dim rngReport As Range
    Set rngReport = wsKosten.Range("A3:T300")
    
    ' Clear all contents and formats in one call
    rngReport.Clear
    
    Exit Sub
    
ErrorHandler:
    Debug.Print "Kosten_ClearReportArea error: " & Err.Description
End Sub

' ========================================
' Helper Functions
' ========================================

Private Function WorksheetExists(ByVal sheetName As String) As Boolean
    ' Checks if a worksheet with the given name exists in ThisWorkbook
    ' Returns True if exists, False otherwise
    
    Dim ws As Worksheet
    On Error Resume Next
    Set ws = ThisWorkbook.Worksheets(sheetName)
    On Error GoTo 0
    
    WorksheetExists = Not ws Is Nothing
End Function

