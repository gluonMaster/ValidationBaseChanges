VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "Kind_AppState"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

' =============================================================================
' Class: Kind_AppState
' Purpose: Captures and restores Excel application state for safe batch operations
' Usage:
'   Dim appState As New Kind_AppState
'   appState.Capture
'   appState.OptimizeForRun
'   ' ... do work ...
'   appState.Restore  ' call in cleanup/finally block with On Error Resume Next
' =============================================================================

' -----------------------------------------------------------------------------
' Private State Storage
' -----------------------------------------------------------------------------
Private m_ScreenUpdating As Boolean
Private m_EnableEvents As Boolean
Private m_DisplayStatusBar As Boolean
Private m_StatusBar As Variant
Private m_Calculation As XlCalculation
Private m_IsCaptured As Boolean


' -----------------------------------------------------------------------------
' Class_Initialize
' Purpose: Initialize default state
' -----------------------------------------------------------------------------
Private Sub Class_Initialize()
    m_IsCaptured = False
End Sub


' -----------------------------------------------------------------------------
' Capture
' Purpose: Saves current Excel application state before modification
' Note: Call this before OptimizeForRun to preserve original settings
' -----------------------------------------------------------------------------
Public Sub Capture()
    On Error Resume Next
    
    m_ScreenUpdating = Application.ScreenUpdating
    m_EnableEvents = Application.EnableEvents
    m_DisplayStatusBar = Application.DisplayStatusBar
    m_StatusBar = Application.StatusBar
    m_Calculation = Application.Calculation
    
    m_IsCaptured = True
    
    On Error GoTo 0
End Sub


' -----------------------------------------------------------------------------
' OptimizeForRun
' Purpose: Sets Excel to optimized state for batch processing
'          - Disables screen updating for speed
'          - Disables events to prevent triggers
'          - Sets manual calculation to avoid recalc delays
'          - Shows status message in status bar
' -----------------------------------------------------------------------------
Public Sub OptimizeForRun()
    On Error Resume Next
    
    ' Capture state if not already done
    If Not m_IsCaptured Then
        Capture
    End If
    
    Application.ScreenUpdating = False
    Application.EnableEvents = False
    Application.Calculation = xlCalculationManual
    Application.DisplayStatusBar = True
    Application.StatusBar = "Verarbeitung laeuft, bitte warten..."
    
    On Error GoTo 0
End Sub


' -----------------------------------------------------------------------------
' Restore
' Purpose: Restores Excel application state to captured values
' Note: Safe to call in cleanup/finally block with On Error Resume Next
'       Will not fail even if state was not captured
' -----------------------------------------------------------------------------
Public Sub Restore()
    On Error Resume Next
    
    ' Only restore if state was captured
    If m_IsCaptured Then
        Application.ScreenUpdating = m_ScreenUpdating
        Application.EnableEvents = m_EnableEvents
        Application.Calculation = m_Calculation
        Application.DisplayStatusBar = m_DisplayStatusBar
        Application.StatusBar = m_StatusBar
    Else
        ' Fallback to safe defaults if state was never captured
        Application.ScreenUpdating = True
        Application.EnableEvents = True
        Application.Calculation = xlCalculationAutomatic
        Application.StatusBar = False
    End If
    
    On Error GoTo 0
End Sub


' -----------------------------------------------------------------------------
' Class_Terminate
' Purpose: Ensures state is restored when object is destroyed
' Note: Acts as safety net if Restore was not explicitly called
' -----------------------------------------------------------------------------
Private Sub Class_Terminate()
    On Error Resume Next
    Restore
    On Error GoTo 0
End Sub
