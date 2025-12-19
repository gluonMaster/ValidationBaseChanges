Attribute VB_Name = "NHblank_Main"
Option Explicit

' Main entry point for the blank generation macro
Public Sub GenerateBlanks()
    Dim screenUpdateState As Boolean
    Dim calculationState As XlCalculation
    Dim eventsState As Boolean
    Dim currentRowNumber As Long
    
    On Error GoTo ErrorHandler
    
    ' Save current row number BEFORE any dialogs or operations
    ' This ensures we capture the correct row where user is positioned
    currentRowNumber = ActiveCell.Row
    
    ' Save current Excel state
    screenUpdateState = Application.ScreenUpdating
    calculationState = Application.Calculation
    eventsState = Application.EnableEvents
    
    ' Optimize performance
    Application.ScreenUpdating = False
    Application.Calculation = xlCalculationManual
    Application.EnableEvents = False
    
    ' Pre-flight diagnostic check (optional - comment out in production)
    ' Uncomment next line to run diagnostics before each execution:
    ' NHblank_Diagnostics.RunFullDiagnostics
    
    ' Ask user for processing mode
    Dim processAllRecords As Boolean
    If Not NHblank_UI.AskProcessingMode(processAllRecords) Then
        ' User cancelled
        GoTo CleanExit
    End If
    
    ' Ask user to select target folder
    Dim targetFolder As String
    targetFolder = NHblank_UI.SelectTargetFolder()
    If targetFolder = "" Then
        ' User cancelled folder selection
        GoTo CleanExit
    End If
    
    ' Check if template file exists
    Dim templatePath As String
    Dim workbookLocalPath As String
    Dim errorMsg As String
    
    ' Get local path (handles OneDrive URLs)
    workbookLocalPath = NHblank_Utils.GetLocalPath(ThisWorkbook.Path)
    
    If workbookLocalPath = "" Then
        ' Could not determine local path, ask user to select template manually
        MsgBox "Automatische Erkennung des Template-Pfades fehlgeschlagen." & vbCrLf & _
               "Bitte waehlen Sie die Datei Shablon.xlsx manuell aus.", vbInformation
        
        Dim templateFilePath As String
        templateFilePath = NHblank_UI.SelectTemplateFile()
        If templateFilePath = "" Then
            ' User cancelled template selection
            GoTo CleanExit
        End If
        templatePath = templateFilePath
    Else
        ' Use detected local path
        templatePath = workbookLocalPath & "\Shablon.xlsx"
        If Dir(templatePath) = "" Then
            ' Template not found in detected path, ask user to select manually
            MsgBox "Datei Shablon.xlsx wurde im Verzeichnis nicht gefunden: " & workbookLocalPath & vbCrLf & _
                   "Bitte waehlen Sie die Datei manuell aus.", vbExclamation
            
            Dim manualTemplatePath As String
            manualTemplatePath = NHblank_UI.SelectTemplateFile()
            If manualTemplatePath = "" Then
                GoTo CleanExit
            End If
            templatePath = manualTemplatePath
        End If
    End If
    
    ' Pre-process validation
    If Not NHblank_Diagnostics.PreProcessCheck(templatePath, targetFolder, errorMsg) Then
        NHblank_UI.ShowError "Validierungsfehler: " & errorMsg
        GoTo CleanExit
    End If
    
    ' Process records
    Dim blanksCreated As Long
    Dim activeRecordsFound As Long
    blanksCreated = NHblank_DataProcessor.ProcessRecords( _
        processAllRecords, _
        templatePath, _
        targetFolder, _
        currentRowNumber, _
        activeRecordsFound _
    )
    
    ' Show appropriate message based on results
    If blanksCreated > 0 Then
        NHblank_UI.ShowSuccess blanksCreated
    ElseIf activeRecordsFound = 0 Then
        NHblank_UI.ShowError "Keine aktiven Datensaetze gefunden." & vbCrLf & vbCrLf & _
                             "Hinweis: Aktive Datensaetze muessen:" & vbCrLf & _
                             "- Nicht leere Zellen in Spalte C haben" & vbCrLf & _
                             "- NICHT grauen Text haben (ColorIndex <> 15)"
    Else
        NHblank_UI.ShowError "Es wurden " & activeRecordsFound & " aktive Datensaetze gefunden," & vbCrLf & _
                             "aber keiner konnte verarbeitet werden." & vbCrLf & vbCrLf & _
                             "Moegliche Ursachen:" & vbCrLf & _
                             "- Leere Pflichtfelder (Nachname, Vorname, Disziplin)" & vbCrLf & _
                             "- Fehler beim Speichern der Dateien" & vbCrLf & _
                             "- Template-Datei hat keinen 'Muster'-Sheet" & vbCrLf & vbCrLf & _
                             "Verwenden Sie 'RunFullDiagnostics' fuer Details."
    End If
    
CleanExit:
    ' Restore Excel state
    Application.ScreenUpdating = screenUpdateState
    Application.Calculation = calculationState
    Application.EnableEvents = eventsState
    Exit Sub
    
ErrorHandler:
    ' Restore Excel state even on error
    Application.ScreenUpdating = screenUpdateState
    Application.Calculation = calculationState
    Application.EnableEvents = eventsState
    
    NHblank_UI.ShowError "Fehler beim Ausfuehren des Makros: " & Err.Description
End Sub

' Temporary debugging procedure - call this to see path information
Public Sub DebugPaths()
    NHblank_Utils.ShowPathDebugInfo
End Sub

' Run full system diagnostics
Public Sub RunFullDiagnostics()
    NHblank_Diagnostics.RunFullDiagnostics
End Sub

