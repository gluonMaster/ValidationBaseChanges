Attribute VB_Name = "Kind_Config"
Option Explicit

' =============================================================================
' Module: Kind_Config
' Purpose: Configuration constants and helper functions for Nachhilfe validation
' Prefix: Kind_ (all modules/classes in this project use this prefix)
' Note: No Access database calls - only reads from open Admin xlsm workbook
' =============================================================================

' -----------------------------------------------------------------------------
' Admin Workbook Configuration
' -----------------------------------------------------------------------------
Public Const KIND_ADMIN_WB_NAME As String = "KindElternDaten_25_Admin.xlsm"
Public Const KIND_ADMIN_SHEET_KARTEI As String = "Kartei"

' -----------------------------------------------------------------------------
' This Workbook Sheet Names
' -----------------------------------------------------------------------------
Public Const KIND_THIS_SHEET_KINDER As String = "Kinder"
Public Const KIND_THIS_SHEET_ERRORLOG As String = "ErrorLog"

' -----------------------------------------------------------------------------
' Kinder Sheet Layout
' -----------------------------------------------------------------------------
Public Const KIND_KINDER_FIRST_DATA_ROW As Long = 5

' Kinder sheet column letters
Public Const KIND_KINDER_COL_FAMILYID As String = "B"
Public Const KIND_KINDER_COL_NACHNAME As String = "C"
Public Const KIND_KINDER_COL_VORNAME As String = "D"
Public Const KIND_KINDER_COL_FACH As String = "E"

' -----------------------------------------------------------------------------
' Kartei Sheet Layout (Admin workbook)
' -----------------------------------------------------------------------------
Public Const KIND_KARTEI_COL_FAMILYID As String = "A"
Public Const KIND_KARTEI_COL_CHILD As String = "D"
Public Const KIND_KARTEI_COL_STATUS As String = "T"

' Fach column depends on month: J for months 1-6, O for months 7-12
Public Const KIND_KARTEI_COL_FACH_FIRST_HALF As String = "J"
Public Const KIND_KARTEI_COL_FACH_SECOND_HALF As String = "O"


' =============================================================================
' Helper Functions
' =============================================================================

' -----------------------------------------------------------------------------
' Kind_TryGetOpenWorkbookByName
' Purpose: Attempts to get reference to an already open workbook by name
' Parameters:
'   wbName - Name of the workbook to find
'   wb     - (ByRef) Output workbook reference if found
' Returns: True if workbook is open and reference obtained, False otherwise
' -----------------------------------------------------------------------------
Public Function Kind_TryGetOpenWorkbookByName(ByVal wbName As String, ByRef wb As Workbook) As Boolean
    Dim wbLoop As Workbook
    
    On Error GoTo ErrHandler
    
    Set wb = Nothing
    Kind_TryGetOpenWorkbookByName = False
    
    ' Iterate through all open workbooks
    For Each wbLoop In Application.Workbooks
        If StrComp(wbLoop.Name, wbName, vbTextCompare) = 0 Then
            Set wb = wbLoop
            Kind_TryGetOpenWorkbookByName = True
            Exit Function
        End If
    Next wbLoop
    
    Exit Function
    
ErrHandler:
    Set wb = Nothing
    Kind_TryGetOpenWorkbookByName = False
End Function


' -----------------------------------------------------------------------------
' Kind_GetKarteiSubjectColumnLetter
' Purpose: Returns the correct Fach column letter based on current month
'          Months 1-6 use column J, months 7-12 use column O
' Returns: Column letter as String ("J" or "O")
' -----------------------------------------------------------------------------
Public Function Kind_GetKarteiSubjectColumnLetter() As String
    Dim currentMonth As Long
    
    currentMonth = Month(Date)
    
    If currentMonth >= 1 And currentMonth <= 6 Then
        Kind_GetKarteiSubjectColumnLetter = KIND_KARTEI_COL_FACH_FIRST_HALF
    Else
        Kind_GetKarteiSubjectColumnLetter = KIND_KARTEI_COL_FACH_SECOND_HALF
    End If
End Function


' -----------------------------------------------------------------------------
' Kind_TryGetWorksheet
' Purpose: Attempts to get reference to a worksheet by name within a workbook
' Parameters:
'   wb        - Workbook to search in
'   sheetName - Name of the worksheet to find
'   ws        - (ByRef) Output worksheet reference if found
' Returns: True if worksheet found and reference obtained, False otherwise
' -----------------------------------------------------------------------------
Public Function Kind_TryGetWorksheet(ByVal wb As Workbook, ByVal sheetName As String, ByRef ws As Worksheet) As Boolean
    Dim wsLoop As Worksheet
    
    On Error GoTo ErrHandler
    
    Set ws = Nothing
    Kind_TryGetWorksheet = False
    
    ' Check if workbook reference is valid
    If wb Is Nothing Then
        Exit Function
    End If
    
    ' Iterate through all worksheets in workbook
    For Each wsLoop In wb.Worksheets
        If StrComp(wsLoop.Name, sheetName, vbTextCompare) = 0 Then
            Set ws = wsLoop
            Kind_TryGetWorksheet = True
            Exit Function
        End If
    Next wsLoop
    
    Exit Function
    
ErrHandler:
    Set ws = Nothing
    Kind_TryGetWorksheet = False
End Function
