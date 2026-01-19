Attribute VB_Name = "pay_Main"
'==========================
'   Module: pay_Main
'   Purpose: Main entry point for family payment exports
'
'   Ported from Data-file/zahl_Main.bas to Superadmin
'   Changes:
'   - Made sheet-agnostic via FamZahlungenForYear(year2)
'   - Uses valid_YearConfig for sheet name resolution
'   - Output files are year-suffixed (EltKosten_24.xlsx, etc.)
'   - Provides wrapper macros: FamZahlungen24, FamZahlungen25, FamZahlungen26
'
'   Output folder: C:\FamZahlung
'   Output files: EltKosten_YY.xlsx, EltKostenNH_YY.xlsx, EltKostenInd_YY.xlsx
'
'   Note: This reads from KarteiYY sheets. If the sheet only contains pending
'   records (from pre_tblKartei), the payment report will be incomplete.
'   For full reporting, ensure KarteiYY contains all tblKartei data.
'==========================

Option Explicit

' ============================================================
' PUBLIC GLOBAL VARIABLES (used by pay_DataProcessor, pay_FileGenerator)
' ============================================================

' Current source worksheet reference (set by FamZahlungenForYear)
Public wsKartei As Worksheet

' Current year being processed (2-digit)
Public currentYear2 As Integer

' Global arrays for data storage (up to 2000 unique families)
Public num(1 To 2000) As String
Public nam(1 To 2000) As String

' Regular costs by months (Jan-Dec)
Public sum01(1 To 2000) As Double, sum02(1 To 2000) As Double
Public sum03(1 To 2000) As Double, sum04(1 To 2000) As Double
Public sum05(1 To 2000) As Double, sum06(1 To 2000) As Double
Public sum07(1 To 2000) As Double, sum08(1 To 2000) As Double
Public sum09(1 To 2000) As Double, sum10(1 To 2000) As Double
Public sum11(1 To 2000) As Double, sum12(1 To 2000) As Double

' Nachhilfe costs by months
Public sum01N(1 To 2000) As Double, sum02N(1 To 2000) As Double
Public sum03N(1 To 2000) As Double, sum04N(1 To 2000) As Double
Public sum05N(1 To 2000) As Double, sum06N(1 To 2000) As Double
Public sum07N(1 To 2000) As Double, sum08N(1 To 2000) As Double
Public sum09N(1 To 2000) As Double, sum10N(1 To 2000) As Double
Public sum11N(1 To 2000) As Double, sum12N(1 To 2000) As Double

' Individual lessons costs by months
Public sum01I(1 To 2000) As Double, sum02I(1 To 2000) As Double
Public sum03I(1 To 2000) As Double, sum04I(1 To 2000) As Double
Public sum05I(1 To 2000) As Double, sum06I(1 To 2000) As Double
Public sum07I(1 To 2000) As Double, sum08I(1 To 2000) As Double
Public sum09I(1 To 2000) As Double, sum10I(1 To 2000) As Double
Public sum11I(1 To 2000) As Double, sum12I(1 To 2000) As Double

' Flags for non-empty records
Public nichtNull(1 To 2000) As Boolean
Public nichtNullN(1 To 2000) As Boolean
Public nichtNullI(1 To 2000) As Boolean

' Color marking flags for conflicts
Public colorRegular(1 To 2000, 1 To 12) As Integer ' 0=normal, 2=yellow
Public colorNachhilfe(1 To 2000, 1 To 12) As Integer ' 0=normal, 2=yellow
Public colorIndividual(1 To 2000, 1 To 12) As Integer ' 0=normal, 2=yellow

Public letzRow As Integer
Public conflictCount As Integer
Public dataErrorCount As Integer

' ============================================================
' PUBLIC API - YEAR-SPECIFIC WRAPPERS
' ============================================================

' Wrapper for year 2024
Public Sub FamZahlungen24()
    Call FamZahlungenForYear(24)
End Sub

' Wrapper for year 2025
Public Sub FamZahlungen25()
    Call FamZahlungenForYear(25)
End Sub

' Wrapper for year 2026
Public Sub FamZahlungen26()
    Call FamZahlungenForYear(26)
End Sub

' ============================================================
' PUBLIC API - MAIN ENTRY POINT
' ============================================================

' Main entry point - processes payments for a specific year
' @param year2 - Two-digit year (24, 25, or 26)
Public Sub FamZahlungenForYear(ByVal year2 As Integer)
    Dim startTime As Double
    Dim sheetName As String
    
    startTime = Timer
    
    If Not IsValidYear(year2) Then
        MsgBox "Ungueltiges Jahr: " & year2 & ". Unterstuetzte Jahre: 24, 25, 26.", _
               vbExclamation, "FamZahlungen - Jahresfehler"
        Exit Sub
    End If
    
    currentYear2 = year2
    
    ' Performance optimization - disable screen updating and calculations
    Call pay_Utils.OptimizePerformance(False)
    
    On Error GoTo ErrorHandler
    
    ' Check if output directory exists, create if needed
    If Not pay_Utils.CreateDirectory("C:\FamZahlung") Then
        MsgBox "Fehler beim Erstellen des Ausgabeordners C:\FamZahlung", vbCritical, "Fehler"
        GoTo Cleanup
    End If
    
    ' Get sheet name for the year (e.g., "Kartei24")
    sheetName = valid_YearConfig.GetKarteiSheetName(year2)
    
    ' Ensure year sheets exist (creates KarteiYY if missing)
    Call valid_YearConfig.EnsureYearSheetsExist(year2)
    
    ' Validate that the KarteiYY worksheet exists
    If Not pay_Utils.ValidateWorksheet(sheetName) Then
        MsgBox "Arbeitsblatt '" & sheetName & "' nicht gefunden." & vbCrLf & vbCrLf & _
               "Bitte stellen Sie sicher, dass das Blatt existiert und Daten enthaelt.", _
               vbCritical, "Arbeitsblatt fehlt"
        GoTo Cleanup
    End If
    
    ' Set the global worksheet reference
    Set wsKartei = ThisWorkbook.Worksheets(sheetName)
    
    ' Initialize variables and arrays
    Call InitializeData
    
    ' Find last row in source worksheet
    letzRow = pay_Utils.LetzteNr(wsKartei, 3, "A")
    
    If letzRow < 3 Then
        MsgBox "Keine Daten im Arbeitsblatt '" & sheetName & "' gefunden", vbInformation, "Information"
        GoTo Cleanup
    End If
    
    ' Validate data consistency (check for duplicate IDs with different names)
    Call ValidateDataConsistency
    
    ' Process all data from source worksheet
    Call pay_DataProcessor.ProcessAllData
    
    ' Generate output files (with year suffix)
    Call pay_FileGenerator.GenerateAllReports
    
    ' Show completion message with conflicts if any
    Call ShowCompletionMessage(Timer - startTime, sheetName)
    
Cleanup:
    ' Restore performance settings
    Call pay_Utils.OptimizePerformance(True)
    Set wsKartei = Nothing
    Exit Sub
    
ErrorHandler:
    ' Restore performance settings in case of error
    Call pay_Utils.OptimizePerformance(True)
    Set wsKartei = Nothing
    If Err.Number <> 0 Then
        MsgBox "Fehler beim Verarbeiten der Daten: " & Err.Description, vbCritical, "Fehler"
    End If
End Sub

' Alternative entry point - accepts worksheet directly
' @param ws - Source worksheet to process
' @param year2 - Optional 2-digit year for filename suffix (defaults to extracting from sheet name)
Public Sub FamZahlungenForSheet(ByVal ws As Worksheet, Optional ByVal year2 As Integer = 0)
    ' Extract year from sheet name if not provided
    If year2 = 0 Then
        year2 = ExtractYearFromSheetName(ws.Name)
        If year2 = 0 Then
            MsgBox "Konnte Jahr nicht aus Blattnamen '" & ws.Name & "' ermitteln." & vbCrLf & _
                   "Bitte year2-Parameter angeben.", vbExclamation, "Jahr fehlt"
            Exit Sub
        End If
    End If
    
    If Not IsValidYear(year2) Then
        MsgBox "Ungueltiges Jahr: " & year2 & ". Unterstuetzte Jahre: 24, 25, 26.", _
               vbExclamation, "FamZahlungen - Jahresfehler"
        Exit Sub
    End If
    
    ' Set globals and delegate to main processing
    currentYear2 = year2
    Set wsKartei = ws
    
    ' Re-use main logic but skip sheet lookup
    Dim startTime As Double
    startTime = Timer
    
    Call pay_Utils.OptimizePerformance(False)
    
    On Error GoTo ErrorHandler
    
    If Not pay_Utils.CreateDirectory("C:\FamZahlung") Then
        MsgBox "Fehler beim Erstellen des Ausgabeordners C:\FamZahlung", vbCritical, "Fehler"
        GoTo Cleanup
    End If
    
    Call InitializeData
    
    letzRow = pay_Utils.LetzteNr(wsKartei, 3, "A")
    
    If letzRow < 3 Then
        MsgBox "Keine Daten im Arbeitsblatt '" & ws.Name & "' gefunden", vbInformation, "Information"
        GoTo Cleanup
    End If
    
    Call ValidateDataConsistency
    Call pay_DataProcessor.ProcessAllData
    Call pay_FileGenerator.GenerateAllReports
    Call ShowCompletionMessage(Timer - startTime, ws.Name)
    
Cleanup:
    Call pay_Utils.OptimizePerformance(True)
    Set wsKartei = Nothing
    Exit Sub
    
ErrorHandler:
    Call pay_Utils.OptimizePerformance(True)
    Set wsKartei = Nothing
    If Err.Number <> 0 Then
        MsgBox "Fehler beim Verarbeiten der Daten: " & Err.Description, vbCritical, "Fehler"
    End If
End Sub

' ============================================================
' PRIVATE HELPER FUNCTIONS
' ============================================================

Private Function ExtractYearFromSheetName(sheetName As String) As Integer
    ' Try to extract 2-digit year from sheet name like "Kartei24"
    Dim i As Integer
    Dim lastTwo As String
    
    If Len(sheetName) >= 2 Then
        lastTwo = Right(sheetName, 2)
        If IsNumeric(lastTwo) Then
            ExtractYearFromSheetName = CInt(lastTwo)
            Exit Function
        End If
    End If
    
    ExtractYearFromSheetName = 0
End Function

Private Function IsValidYear(ByVal year2 As Integer) As Boolean
    IsValidYear = (year2 = 24 Or year2 = 25 Or year2 = 26)
End Function

Private Sub InitializeData()
    ' Initialize all arrays and counters
    Dim i As Integer, j As Integer
    
    conflictCount = 0
    dataErrorCount = 0
    
    For i = 1 To 2000
        num(i) = ""
        nam(i) = ""
        
        ' Initialize all sum arrays
        sum01(i) = 0: sum02(i) = 0: sum03(i) = 0: sum04(i) = 0
        sum05(i) = 0: sum06(i) = 0: sum07(i) = 0: sum08(i) = 0
        sum09(i) = 0: sum10(i) = 0: sum11(i) = 0: sum12(i) = 0
        
        sum01N(i) = 0: sum02N(i) = 0: sum03N(i) = 0: sum04N(i) = 0
        sum05N(i) = 0: sum06N(i) = 0: sum07N(i) = 0: sum08N(i) = 0
        sum09N(i) = 0: sum10N(i) = 0: sum11N(i) = 0: sum12N(i) = 0
        
        sum01I(i) = 0: sum02I(i) = 0: sum03I(i) = 0: sum04I(i) = 0
        sum05I(i) = 0: sum06I(i) = 0: sum07I(i) = 0: sum08I(i) = 0
        sum09I(i) = 0: sum10I(i) = 0: sum11I(i) = 0: sum12I(i) = 0
        
        nichtNull(i) = False
        nichtNullN(i) = False
        nichtNullI(i) = False
        
        ' Initialize color arrays
        For j = 1 To 12
            colorRegular(i, j) = 0
            colorNachhilfe(i, j) = 0
            colorIndividual(i, j) = 0
        Next j
    Next i
End Sub

Private Sub ShowCompletionMessage(processingTime As Double, sheetName As String)
    Dim message As String
    Dim yearSuffix As String
    
    yearSuffix = pay_Utils.GetYearSuffix(currentYear2)
    
    message = "Verarbeitung von '" & sheetName & "' abgeschlossen in " & _
              Format(processingTime, "0.0") & " Sekunden." & vbCrLf & vbCrLf & _
              "Ausgabedateien in C:\FamZahlung:" & vbCrLf & _
              "  - EltKosten" & yearSuffix & ".xlsx" & vbCrLf & _
              "  - EltKostenNH" & yearSuffix & ".xlsx" & vbCrLf & _
              "  - EltKostenInd" & yearSuffix & ".xlsx"
    
    ' Check for data consistency errors
    If dataErrorCount > 0 Then
        message = message & vbCrLf & vbCrLf & "FEHLER: Gefunden " & dataErrorCount & _
                 " Familien-IDs mit unterschiedlichen Namen." & _
                 vbCrLf & "Details siehe Datei: C:\FamZahlung\Datenfehler" & yearSuffix & ".xlsx"
    End If
    
    If conflictCount > 0 Then
        message = message & vbCrLf & vbCrLf & "WARNUNG: Gefunden " & conflictCount & _
                 " Datensaetze mit gemischten Markierungen von Individual- und Nachhilfe-Unterricht." & _
                 vbCrLf & "Diese wurden gelb markiert in der Individual-Datei."
    End If
    
    If dataErrorCount > 0 Then
        MsgBox message, vbCritical, "Verarbeitung abgeschlossen mit Fehlern"
    ElseIf conflictCount > 0 Then
        MsgBox message, vbExclamation, "Verarbeitung abgeschlossen"
    Else
        MsgBox message, vbInformation, "Verarbeitung abgeschlossen"
    End If
End Sub

Private Sub ValidateDataConsistency()
    ' Check for families with same ID but different names
    Dim familyMap As Object
    Dim i As Integer
    Dim familyID As String, familyName As String
    Dim existingName As String
    Dim errorReport As String
    Dim wb As Workbook
    Dim ws As Worksheet
    Dim yearSuffix As String
    
    Set familyMap = CreateObject("Scripting.Dictionary")
    errorReport = ""
    dataErrorCount = 0
    yearSuffix = pay_Utils.GetYearSuffix(currentYear2)
    
    ' Scan through all data rows
    For i = 3 To letzRow
        familyID = CStr(wsKartei.Cells(i, "A").Value)
        familyName = CStr(wsKartei.Cells(i, "B").Value)
        
        ' Skip empty rows
        If Len(Trim(familyID)) = 0 Then GoTo NextRow
        
        If familyMap.Exists(familyID) Then
            existingName = familyMap(familyID)
            If existingName <> familyName Then
                ' Found inconsistency
                dataErrorCount = dataErrorCount + 1
                errorReport = errorReport & "Zeile " & i & ": Familie-ID '" & familyID & _
                             "' mit unterschiedlichen Namen: '" & existingName & "' und '" & familyName & "'" & vbCrLf
            End If
        Else
            familyMap.Add familyID, familyName
        End If
        
NextRow:
    Next i
    
    ' Create error report file if errors found
    If dataErrorCount > 0 Then
        Set wb = Workbooks.Add
        Set ws = wb.Sheets(1)
        
        ws.Cells(1, 1) = "Datenfehler-Bericht fuer " & wsKartei.Name
        ws.Cells(2, 1) = "Gefunden am: " & Format(Now, "dd.mm.yyyy hh:mm:ss")
        ws.Cells(3, 1) = "Anzahl Fehler: " & dataErrorCount
        ws.Cells(4, 1) = ""
        ws.Cells(5, 1) = "Details:"
        
        ' Split error report into lines and add to worksheet
        Dim errorLines As Variant
        Dim j As Integer
        errorLines = Split(errorReport, vbCrLf)
        
        For j = 0 To UBound(errorLines) - 1 ' -1 because last element is empty
            ws.Cells(6 + j, 1) = errorLines(j)
        Next j
        
        ' Format the report
        ws.Columns("A:A").AutoFit
        ws.Cells(1, 1).Font.Bold = True
        ws.Cells(1, 1).Font.Size = 14
        
        ' Save error report with year suffix
        On Error GoTo SaveErrorReport
        wb.SaveAs filename:="C:\FamZahlung\Datenfehler" & yearSuffix & ".xlsx", _
                  FileFormat:=xlOpenXMLWorkbook, CreateBackup:=False
        wb.Close
        On Error GoTo 0
        
        Exit Sub
        
SaveErrorReport:
        MsgBox "Fehler beim Speichern des Datenfehler-Berichts: " & Err.Description, vbCritical, "Speicherfehler"
        wb.Close False
        On Error GoTo 0
    End If
End Sub

