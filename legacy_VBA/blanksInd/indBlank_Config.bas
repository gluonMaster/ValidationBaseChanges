Attribute VB_Name = "indBlank_Config"
Option Explicit

' =============================================================================
' Module: indBlank_Config
' Purpose: Configuration constants and helper functions for individual blank
'          (Einzelmeldung) generation in KindElternDaten_25_Admin.xlsm
' =============================================================================

' -----------------------------------------------------------------------------
' Sheet and Row Constants
' -----------------------------------------------------------------------------
Public Const KARTEI_SHEET As String = "Kartei"
Public Const KARTEI_FIRST_DATA_ROW As Long = 3

' -----------------------------------------------------------------------------
' Column Constants for Kartei sheet
' -----------------------------------------------------------------------------
Public Const KARTEI_COL_CHILD As String = "D"
Public Const KARTEI_COL_FAMILYID As String = "A"
Public Const KARTEI_COL_DISCIPLINE_H1 As String = "J"
Public Const KARTEI_COL_TEACHER_H1 As String = "K"
Public Const KARTEI_COL_DISCIPLINE_H2 As String = "O"
Public Const KARTEI_COL_TEACHER_H2 As String = "P"

' -----------------------------------------------------------------------------
' Template Constants
' -----------------------------------------------------------------------------
Public Const TEMPLATE_PATH As String = "C:\temp\ShablonInd.xlsx"
Public Const TEMPLATE_SHEET As String = "Muster"

' =============================================================================
' Helper: indBlank_GetSemesterColumns
' Purpose: Returns discipline and teacher columns based on current semester
'          January-June (months 1-6) -> H1 columns (J/K)
'          July-December (months 7-12) -> H2 columns (O/P)
' =============================================================================
Public Sub indBlank_GetSemesterColumns(ByRef outDisciplineCol As String, ByRef outTeacherCol As String, Optional ByVal monthValue As Long = 0)
    Dim currentMonth As Long
    
    If monthValue >= 1 And monthValue <= 12 Then
        currentMonth = monthValue
    Else
        currentMonth = Month(Date)
    End If
    
    If currentMonth >= 1 And currentMonth <= 6 Then
        ' First semester (Halbjahr 1): January - June
        outDisciplineCol = KARTEI_COL_DISCIPLINE_H1
        outTeacherCol = KARTEI_COL_TEACHER_H1
    Else
        ' Second semester (Halbjahr 2): July - December
        outDisciplineCol = KARTEI_COL_DISCIPLINE_H2
        outTeacherCol = KARTEI_COL_TEACHER_H2
    End If
End Sub

' =============================================================================
' Helper: indBlank_TryGetWorksheet
' Purpose: Safely attempts to get a worksheet by name from a workbook
' Returns: True if worksheet exists and is assigned to ws, False otherwise
' =============================================================================
Public Function indBlank_TryGetWorksheet(ByVal wb As Workbook, ByVal sheetName As String, ByRef ws As Worksheet) As Boolean
    On Error Resume Next
    Set ws = wb.Worksheets(sheetName)
    On Error GoTo 0
    
    If ws Is Nothing Then
        indBlank_TryGetWorksheet = False
    Else
        indBlank_TryGetWorksheet = True
    End If
End Function

' =============================================================================
' Helper: indBlank_GetGermanMonth
' Purpose: Returns German month name without umlauts
' Param:   monthNumber - month number (1-12)
' Returns: German month name as string (e.g., "Januar", "Maerz")
' =============================================================================
Public Function indBlank_GetGermanMonth(ByVal monthNumber As Long) As String
    Select Case monthNumber
        Case 1
            indBlank_GetGermanMonth = "Januar"
        Case 2
            indBlank_GetGermanMonth = "Februar"
        Case 3
            indBlank_GetGermanMonth = "Maerz"
        Case 4
            indBlank_GetGermanMonth = "April"
        Case 5
            indBlank_GetGermanMonth = "Mai"
        Case 6
            indBlank_GetGermanMonth = "Juni"
        Case 7
            indBlank_GetGermanMonth = "Juli"
        Case 8
            indBlank_GetGermanMonth = "August"
        Case 9
            indBlank_GetGermanMonth = "September"
        Case 10
            indBlank_GetGermanMonth = "Oktober"
        Case 11
            indBlank_GetGermanMonth = "November"
        Case 12
            indBlank_GetGermanMonth = "Dezember"
        Case Else
            indBlank_GetGermanMonth = "Unbekannt"
    End Select
End Function
