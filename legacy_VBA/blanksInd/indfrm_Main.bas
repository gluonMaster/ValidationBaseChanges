Attribute VB_Name = "indfrm_Main"
' Main module for individual form functionality
' Contains global variables and main entry points

Option Explicit

' Global variable to store the active row from Kartei sheet
Public g_ActiveKarteiRow As Long

' Global constants for sheet names
Public Const KARTEI_SHEET As String = "Kartei"
Public Const FORM_SHEET As String = "frmInd"

' Entry point for opening the form (called by button on Kartei sheet)
Public Sub indfrm_OpenForm()
    On Error GoTo ErrorHandler
    
    Call indfrm_Utils_DisableUpdates
    Call indfrm_FormOpen_Execute
    Call indfrm_Utils_HideHelperWorksheet ' Hide helper sheet after opening form
    Call indfrm_Utils_EnableUpdates
    
    Exit Sub
    
ErrorHandler:
    Call indfrm_Utils_EnableUpdates
    MsgBox "Fehler beim Offnen der Form: " & Err.Description, 16, "Fehler" ' 16 = vbCritical
End Sub

' Entry point for processing month selection (called by "-->" button on form)
Public Sub indfrm_ProcessMonth()
    On Error GoTo ErrorHandler
    
    Call indfrm_Utils_DisableUpdates
    Call indfrm_ProcessMonth_Execute
    Call indfrm_Utils_EnableUpdates
    
    Exit Sub
    
ErrorHandler:
    Call indfrm_Utils_EnableUpdates
    MsgBox "Fehler bei der Monatsverarbeitung: " & Err.Description, 16, "Fehler" ' 16 = vbCritical
End Sub

' Entry point for updating Kartei data (called by update button on form)
Public Sub indfrm_UpdateKartei()
    On Error GoTo ErrorHandler
    
    Call indfrm_Utils_DisableUpdates
    Call indfrm_UpdateKartei_Execute
    Call indfrm_Utils_EnableUpdates
    
    Exit Sub
    
ErrorHandler:
    Call indfrm_Utils_EnableUpdates
    MsgBox "Fehler beim Aktualisieren der Kartei: " & Err.Description, 16, "Fehler" ' 16 = vbCritical
End Sub

' Entry point for handling child name selection (called when B4 changes)
Public Sub indfrm_OnChildSelection()
    On Error GoTo ErrorHandler
    
    Call indfrm_Utils_DisableUpdates
    Call indfrm_ProcessMonth_OnChildSelected
    Call indfrm_Utils_EnableUpdates
    
    Exit Sub
    
ErrorHandler:
    Call indfrm_Utils_EnableUpdates
    MsgBox "Fehler bei der Kindauswahl: " & Err.Description, 16, "Fehler" ' 16 = vbCritical
End Sub

' Entry point for handling lesson selection (called when B6 changes)
Public Sub indfrm_OnLessonSelection()
    On Error GoTo ErrorHandler
    
    Call indfrm_Utils_DisableUpdates
    Call indfrm_ProcessMonth_OnLessonSelected
    Call indfrm_Utils_EnableUpdates
    
    Exit Sub
    
ErrorHandler:
    Call indfrm_Utils_EnableUpdates
    MsgBox "Fehler bei der Unterrichtauswahl: " & Err.Description, 16, "Fehler" ' 16 = vbCritical
End Sub

' Setup instructions for worksheet events (one-time setup)
Public Sub indfrm_SetupEvents()
    Call indfrm_SheetEvents_SetupWorksheetEvents
End Sub
