Attribute VB_Name = "indBlank_UI"
Option Explicit

' =============================================================================
' Module: indBlank_UI
' Purpose: User interface utilities for individual blank generation
'          (no UserForms, only InputBox/MsgBox/FileDialog)
' =============================================================================

' =============================================================================
' Function: indBlank_AskYear
' Purpose: Prompts user to enter a year via InputBox
' Param:   outYear - output year value
' Returns: True if valid year entered, False if cancelled or invalid
' =============================================================================
Public Function indBlank_AskYear(ByRef outYear As Long) As Boolean
    Dim inputValue As String
    Dim yearValue As Long
    Dim defaultYear As String
    Dim attempts As Long
    Const MAX_ATTEMPTS As Long = 3
    
    defaultYear = CStr(Year(Date))
    indBlank_AskYear = False
    
    For attempts = 1 To MAX_ATTEMPTS
        inputValue = InputBox( _
            "Bitte geben Sie das Schuljahr ein:" & vbCrLf & _
            "(z.B. " & defaultYear & ")", _
            "Schuljahr eingeben", _
            defaultYear)
        
        ' Check for Cancel
        If Len(inputValue) = 0 Then
            Exit Function
        End If
        
        ' Validate numeric
        If Not IsNumeric(inputValue) Then
            MsgBox "Ungueltige Eingabe. Bitte geben Sie eine Jahreszahl ein.", _
                   vbExclamation, "Fehler"
        Else
            yearValue = CLng(inputValue)
            
            ' Validate range
            If yearValue >= 2000 And yearValue <= 2100 Then
                outYear = yearValue
                indBlank_AskYear = True
                Exit Function
            Else
                MsgBox "Das Jahr muss zwischen 2000 und 2100 liegen.", _
                       vbExclamation, "Fehler"
            End If
        End If
    Next attempts
    
    ' Max attempts reached
    MsgBox "Maximale Anzahl an Versuchen erreicht. Vorgang abgebrochen.", _
           vbCritical, "Abbruch"
End Function

' =============================================================================
' Function: indBlank_AskMonth
' Purpose: Prompts user to enter a month (1-12) via InputBox
' Param:   outMonth - output month value (1-12)
' Returns: True if valid month entered, False if cancelled or invalid
' =============================================================================
Public Function indBlank_AskMonth(ByRef outMonth As Long) As Boolean
    Dim inputValue As String
    Dim monthValue As Long
    Dim defaultMonth As String
    Dim attempts As Long
    Const MAX_ATTEMPTS As Long = 3
    
    defaultMonth = CStr(Month(Date))
    indBlank_AskMonth = False
    
    For attempts = 1 To MAX_ATTEMPTS
        inputValue = InputBox( _
            "Bitte geben Sie den Monat ein (1-12):", _
            "Monat eingeben", _
            defaultMonth)
        
        ' Cancel
        If Len(inputValue) = 0 Then
            Exit Function
        End If
        
        If Not IsNumeric(inputValue) Then
            MsgBox "Ungueltige Eingabe. Bitte geben Sie eine Zahl von 1 bis 12 ein.", _
                   vbExclamation, "Fehler"
        Else
            monthValue = CLng(inputValue)
            If monthValue >= 1 And monthValue <= 12 Then
                outMonth = monthValue
                indBlank_AskMonth = True
                Exit Function
            Else
                MsgBox "Der Monat muss zwischen 1 und 12 liegen.", _
                       vbExclamation, "Fehler"
            End If
        End If
    Next attempts
    
    MsgBox "Maximale Anzahl an Versuchen erreicht. Vorgang abgebrochen.", _
           vbCritical, "Abbruch"
End Function

' =============================================================================
' Function: indBlank_SelectTargetFolder
' Purpose: Opens folder picker dialog for user to select target folder
' Returns: Selected folder path (with trailing backslash) or empty string if cancelled
' =============================================================================
Public Function indBlank_SelectTargetFolder() As String
    Dim fd As Object
    Dim selectedPath As String
    
    indBlank_SelectTargetFolder = ""
    
    Set fd = Application.FileDialog(msoFileDialogFolderPicker)
    
    With fd
        .Title = "Zielordner fuer Einzelmeldungen auswaehlen"
        .InitialFileName = Application.DefaultFilePath
        .AllowMultiSelect = False
        
        If .Show = -1 Then
            selectedPath = .SelectedItems(1)
            
            ' Ensure trailing backslash
            If Right$(selectedPath, 1) <> "\" Then
                selectedPath = selectedPath & "\"
            End If
            
            indBlank_SelectTargetFolder = selectedPath
        End If
    End With
    
    Set fd = Nothing
End Function

' =============================================================================
' Sub: indBlank_ShowError
' Purpose: Displays error message to user (German, no umlauts)
' Param:   msg - error message text
' =============================================================================
Public Sub indBlank_ShowError(ByVal msg As String)
    MsgBox msg, vbCritical, "Fehler"
End Sub

' =============================================================================
' Sub: indBlank_ShowSummary
' Purpose: Displays summary of generation results (German, no umlauts)
' Param:   createdCount - number of blanks successfully created
' Param:   skippedCount - number of rows skipped (not Ind/VSpE)
' Param:   errorCount - number of errors encountered
' =============================================================================
Public Sub indBlank_ShowSummary(ByVal createdCount As Long, ByVal skippedCount As Long, ByVal errorCount As Long)
    Dim summaryText As String
    Dim iconType As VbMsgBoxStyle
    Dim title As String
    
    summaryText = "Verarbeitung abgeschlossen." & vbCrLf & vbCrLf & _
                  "Erstellt: " & createdCount & vbCrLf & _
                  "Uebersprungen: " & skippedCount & vbCrLf & _
                  "Fehler: " & errorCount
    
    If errorCount > 0 Then
        iconType = vbExclamation
        title = "Abgeschlossen mit Fehlern"
    ElseIf createdCount = 0 Then
        iconType = vbInformation
        title = "Keine Dokumente erstellt"
    Else
        iconType = vbInformation
        title = "Erfolgreich abgeschlossen"
    End If
    
    MsgBox summaryText, iconType, title
End Sub

' =============================================================================
' Function: indBlank_ConfirmIfNoFilter
' Purpose: Asks for confirmation if worksheet has no active filter
'          (to prevent processing all rows unintentionally)
' Param:   ws - worksheet to check
' Returns: True if user confirms or filter is active, False to cancel
' =============================================================================
Public Function indBlank_ConfirmIfNoFilter(ByVal ws As Worksheet) As Boolean
    Dim hasActiveFilter As Boolean
    
    indBlank_ConfirmIfNoFilter = True
    
    ' Check if AutoFilter exists and is in FilterMode
    On Error Resume Next
    hasActiveFilter = False
    If ws.AutoFilterMode Then
        If ws.FilterMode Then
            hasActiveFilter = True
        End If
    End If
    On Error GoTo 0
    
    ' If no filter is active, ask for confirmation
    If Not hasActiveFilter Then
        Dim response As VbMsgBoxResult
        response = MsgBox( _
            "Es ist kein Filter aktiv." & vbCrLf & _
            "Alle sichtbaren Zeilen werden verarbeitet." & vbCrLf & vbCrLf & _
            "Moechten Sie fortfahren?", _
            vbQuestion + vbYesNo, _
            "Kein Filter aktiv")
        
        indBlank_ConfirmIfNoFilter = (response = vbYes)
    End If
End Function
