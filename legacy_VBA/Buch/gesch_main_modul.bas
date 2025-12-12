Attribute VB_Name = "gesch_main_modul"
Sub GestGesch_MainDataProcessor()
    ' Main macro to process KindElternDaten data and create GesternGeshichte sheet
    
    Dim originalCalculation As XlCalculation
    Dim originalScreenUpdating As Boolean
    Dim originalDisplayAlerts As Boolean
    Dim targetWorkbook As Workbook
    Dim sourceWorkbook As Workbook
    Dim targetFilePath As String
    
    ' Store original Excel settings
    originalCalculation = Application.Calculation
    originalScreenUpdating = Application.ScreenUpdating
    originalDisplayAlerts = Application.DisplayAlerts
    
    ' Disable screen updating, calculation and alerts for performance
    Application.ScreenUpdating = False
    Application.Calculation = xlCalculationManual
    Application.DisplayAlerts = False
    
    On Error GoTo ErrorHandler
    
    Set sourceWorkbook = ThisWorkbook
    targetFilePath = sourceWorkbook.Path & "\KindElternDaten_25_Data.xlsm"
    
    ' Check if target file exists
    If Not FileExists(targetFilePath) Then
        MsgBox "Datei KindElternDaten_25_Data.xlsm wurde im aktuellen Verzeichnis nicht gefunden.", _
               vbCritical, "Datei nicht gefunden"
        GoTo Cleanup
    End If
    
    ' Open or get reference to target workbook
    Set targetWorkbook = OpenOrGetWorkbook(targetFilePath)
    If targetWorkbook Is Nothing Then
        GoTo Cleanup
    End If
    
    ' Activate Kartei sheet before running ImportFromBase macro
    If Not ActivateSheet(targetWorkbook, "Kartei") Then
        GoTo CleanupWithClose
    End If
    
    ' Execute ImportFromBase macro
    If Not ExecuteMacro(targetWorkbook, "ImportFromBase") Then
        GoTo CleanupWithClose
    End If
    
    ' Set yesterday's dates in GrossGeschichte sheet
    If Not SetYesterdayDates(targetWorkbook) Then
        GoTo CleanupWithClose
    End If
    
    ' Execute GrossGeshichteMachen macro
    If Not ExecuteMacro(targetWorkbook, "GrossGeshichteMachen") Then
        GoTo CleanupWithClose
    End If
    
    ' Create or prepare GesternGeshichte sheet in source workbook
    If Not PrepareGesternGeshichteSheet(sourceWorkbook) Then
        GoTo CleanupWithClose
    End If
    
    ' Copy data from GrossGeschichte to GesternGeshichte
    If Not CopyDataToGesternGeshichte(targetWorkbook, sourceWorkbook) Then
        GoTo CleanupWithClose
    End If
    
    ' Close target workbook without saving
    targetWorkbook.Close SaveChanges:=False
    
    MsgBox "Datenverarbeitung erfolgreich abgeschlossen.", vbInformation, "Erfolg"
    GoTo Cleanup
    
CleanupWithClose:
    ' Close target workbook without saving if it's open
    If Not targetWorkbook Is Nothing Then
        targetWorkbook.Close SaveChanges:=False
    End If
    
ErrorHandler:
    If Err.Number <> 0 Then
        MsgBox "Fehler aufgetreten: " & Err.Description, vbCritical, "Fehler"
    End If
    
Cleanup:
    ' Restore original Excel settings
    Application.ScreenUpdating = originalScreenUpdating
    Application.Calculation = originalCalculation
    Application.DisplayAlerts = originalDisplayAlerts
    
    ' Clean up object references
    Set targetWorkbook = Nothing
    Set sourceWorkbook = Nothing
End Sub

