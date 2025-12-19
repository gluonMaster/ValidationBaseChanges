Attribute VB_Name = "NHblank_UI"
Option Explicit

' Ask user whether to process current record or all active records
' Returns True if user made a choice, False if cancelled
Public Function AskProcessingMode(ByRef processAllRecords As Boolean) As Boolean
    Dim response As VbMsgBoxResult
    
    response = MsgBox( _
        "Moechten Sie einen Blank fuer den aktuellen Datensatz erstellen?" & vbCrLf & vbCrLf & _
        "Ja - Nur aktueller Datensatz" & vbCrLf & _
        "Nein - Alle aktiven Datensaetze", _
        vbYesNoCancel + vbQuestion, _
        "Modus auswaehlen" _
    )
    
    Select Case response
        Case vbYes
            processAllRecords = False
            AskProcessingMode = True
        Case vbNo
            processAllRecords = True
            AskProcessingMode = True
        Case vbCancel
            AskProcessingMode = False
    End Select
End Function

' Let user select target folder for saving blanks
' Returns folder path or empty string if cancelled
Public Function SelectTargetFolder() As String
    Dim folderDialog As fileDialog
    
    Set folderDialog = Application.fileDialog(msoFileDialogFolderPicker)
    
    With folderDialog
        .title = "Zielordner fuer Blanks auswaehlen"
        .AllowMultiSelect = False
        
        If .Show = -1 Then
            SelectTargetFolder = .SelectedItems(1)
        Else
            SelectTargetFolder = ""
        End If
    End With
    
    Set folderDialog = Nothing
End Function

' Let user select template file (Shablon.xlsx)
' Returns file path or empty string if cancelled
Public Function SelectTemplateFile() As String
    Dim fileDialog As fileDialog
    
    Set fileDialog = Application.fileDialog(msoFileDialogFilePicker)
    
    With fileDialog
        .title = "Bitte waehlen Sie die Datei Shablon.xlsx aus"
        .AllowMultiSelect = False
        .Filters.Clear
        .Filters.Add "Excel Dateien", "*.xlsx;*.xls"
        .Filters.Add "Alle Dateien", "*.*"
        
        ' Try to set initial directory to user's Documents folder
        .InitialFileName = Environ("USERPROFILE") & "\Documents\"
        
        If .Show = -1 Then
            SelectTemplateFile = .SelectedItems(1)
        Else
            SelectTemplateFile = ""
        End If
    End With
    
    Set fileDialog = Nothing
End Function

' Show error message to user
Public Sub ShowError(ByVal message As String)
    MsgBox message, vbCritical, "Fehler"
End Sub

' Show warning about wrong worksheet
Public Sub ShowWrongSheetWarning()
    MsgBox "Bitte waehlen Sie eine Zeile auf dem Blatt 'Kinder' aus.", vbExclamation, "Falsches Blatt"
End Sub

' Show success message with count of created blanks
Public Sub ShowSuccess(ByVal blanksCreated As Long)
    MsgBox _
        "Erfolgreich abgeschlossen!" & vbCrLf & vbCrLf & _
        "Anzahl erstellter Blanks: " & blanksCreated, _
        vbInformation, _
        "Fertig"
End Sub

