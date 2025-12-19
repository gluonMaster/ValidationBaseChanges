VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "indBlank_AppState"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

' =============================================================================
' Class: indBlank_AppState
' Purpose: Captures and restores Excel Application state for performance
'          optimization during bulk operations
' =============================================================================

' -----------------------------------------------------------------------------
' Private state storage
' -----------------------------------------------------------------------------
Private m_ScreenUpdating As Boolean
Private m_EnableEvents As Boolean
Private m_Calculation As XlCalculation
Private m_DisplayAlerts As Boolean
Private m_StatusBar As Variant
Private m_DisplayStatusBar As Boolean
Private m_IsCaptured As Boolean

' =============================================================================
' Method: Capture
' Purpose: Captures current Excel Application state for later restoration
' =============================================================================
Public Sub Capture()
    m_ScreenUpdating = Application.ScreenUpdating
    m_EnableEvents = Application.EnableEvents
    m_Calculation = Application.Calculation
    m_DisplayAlerts = Application.DisplayAlerts
    m_StatusBar = Application.StatusBar
    m_DisplayStatusBar = Application.DisplayStatusBar
    m_IsCaptured = True
End Sub

' =============================================================================
' Method: OptimizeForRun
' Purpose: Sets Application properties for maximum performance during bulk ops
' Param:   statusText - Optional status bar text to display (German, no umlauts)
' =============================================================================
Public Sub OptimizeForRun(Optional ByVal statusText As String = "Verarbeitung laeuft, bitte warten...")
    ' Ensure state is captured before optimizing
    If Not m_IsCaptured Then
        Capture
    End If
    
    ' Disable screen updates for speed
    Application.ScreenUpdating = False
    
    ' Disable event handlers to prevent cascading triggers
    Application.EnableEvents = False
    
    ' Set calculation to manual to avoid recalc on every change
    Application.Calculation = xlCalculationManual
    
    ' Suppress alert dialogs
    Application.DisplayAlerts = False
    
    ' Show status bar with progress message
    Application.DisplayStatusBar = True
    Application.StatusBar = statusText
End Sub

' =============================================================================
' Method: Restore
' Purpose: Safely restores all captured Application state settings
'          Uses On Error Resume Next for safe cleanup in error scenarios
' =============================================================================
Public Sub Restore()
    On Error Resume Next
    
    ' Only restore if state was previously captured
    If m_IsCaptured Then
        Application.ScreenUpdating = m_ScreenUpdating
        Application.EnableEvents = m_EnableEvents
        Application.Calculation = m_Calculation
        Application.DisplayAlerts = m_DisplayAlerts
        Application.DisplayStatusBar = m_DisplayStatusBar
        
        ' Restore status bar (False clears custom text)
        If VarType(m_StatusBar) = vbBoolean Then
            Application.StatusBar = False
        Else
            Application.StatusBar = m_StatusBar
        End If
        
        m_IsCaptured = False
    Else
        ' Fallback: restore to safe defaults if not captured
        Application.ScreenUpdating = True
        Application.EnableEvents = True
        Application.Calculation = xlCalculationAutomatic
        Application.DisplayAlerts = True
        Application.StatusBar = False
    End If
    
    On Error GoTo 0
End Sub

' =============================================================================
' Class_Terminate
' Purpose: Automatic cleanup - ensures state is restored if object is destroyed
' =============================================================================
Private Sub Class_Terminate()
    If m_IsCaptured Then
        Restore
    End If
End Sub
