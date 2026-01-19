Attribute VB_Name = "zahl_Main"
' Main module for processing family payments
' Handles the main workflow and coordination between modules
'
' Recent changes:
' - Fixed table headers to be in row 1 with proper column names
' - Removed pink cell coloring for mixed cases (kept yellow for conflicts)
' - Added data validation to detect families with same ID but different names
' - Creates error report file (Datenfehler.xlsx) when data inconsistencies found

Option Explicit

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

Sub FamZahlungen()
    Dim startTime As Double
    startTime = Timer
    
    ' Performance optimization - disable screen updating and calculations
    Call zahl_Utils.OptimizePerformance(False)
    
    On Error GoTo ErrorHandler
    
    ' Check if output directory exists, create if needed
    If Not zahl_Utils.CreateDirectory("C:\FamZahlung") Then
        MsgBox "Fehler beim Erstellen des Ausgabeordners C:\FamZahlung", vbCritical, "Fehler"
        GoTo ErrorHandler
    End If
    
    ' Validate that Kartei worksheet exists
    If Not zahl_Utils.ValidateWorksheet("Kartei") Then
        MsgBox "Arbeitsblatt 'Kartei' nicht gefunden", vbCritical, "Fehler"
        GoTo ErrorHandler
    End If
    
    ' Initialize variables and arrays
    Call InitializeData
    
    ' Find last row in Kartei worksheet
    letzRow = zahl_Utils.LetzteNr("Kartei", 3, "A")
    
    If letzRow < 3 Then
        MsgBox "Keine Daten im Arbeitsblatt 'Kartei' gefunden", vbInformation, "Information"
        GoTo ErrorHandler
    End If
    
    ' Validate data consistency (check for duplicate IDs with different names)
    Call ValidateDataConsistency
    
    ' Process all data from Kartei worksheet
    Call zahl_DataProcessor.ProcessAllData
    
    ' Generate output files
    Call zahl_FileGenerator.GenerateAllReports
    
    ' Show completion message with conflicts if any
    Call ShowCompletionMessage(Timer - startTime)
    
    ' Restore performance settings
    Call zahl_Utils.OptimizePerformance(True)
    Exit Sub
    
ErrorHandler:
    ' Restore performance settings in case of error
    Call zahl_Utils.OptimizePerformance(True)
    If Err.Number <> 0 Then
        MsgBox "Fehler beim Verarbeiten der Daten: " & Err.Description, vbCritical, "Fehler"
    End If
End Sub

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

Private Sub ShowCompletionMessage(processingTime As Double)
    Dim message As String
    message = "Verarbeitung abgeschlossen in " & Format(processingTime, "0.0") & " Sekunden."
    
    ' Check for data consistency errors
    If dataErrorCount > 0 Then
        message = message & vbCrLf & vbCrLf & "FEHLER: Gefunden " & dataErrorCount & _
                 " Familien-IDs mit unterschiedlichen Namen." & _
                 vbCrLf & "Details siehe Datei: C:\FamZahlung\Datenfehler.xlsx"
    End If
    
    If conflictCount > 0 Then
        message = message & vbCrLf & vbCrLf & "WARNUNG: Gefunden " & conflictCount & _
                 " Datensatze mit gemischten Markierungen von Individual- und Nachhilfe-Unterricht." & _
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
    
    Set familyMap = CreateObject("Scripting.Dictionary")
    errorReport = ""
    dataErrorCount = 0
    
    ' Scan through all data rows
    For i = 3 To letzRow
        familyID = CStr(Worksheets("Kartei").Cells(i, "A").Value)
        familyName = CStr(Worksheets("Kartei").Cells(i, "B").Value)
        
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
        
        ws.Cells(1, 1) = "Datenfehler-Bericht"
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
        
        ' Save error report
        On Error GoTo SaveErrorReport
        wb.SaveAs filename:="C:\FamZahlung\Datenfehler.xlsx", FileFormat:=xlOpenXMLWorkbook, CreateBackup:=False
        wb.Close
        On Error GoTo 0
        
        Exit Sub
        
SaveErrorReport:
        MsgBox "Fehler beim Speichern des Datenfehler-Berichts: " & Err.Description, vbCritical, "Speicherfehler"
        wb.Close False
        On Error GoTo 0
    End If
End Sub
