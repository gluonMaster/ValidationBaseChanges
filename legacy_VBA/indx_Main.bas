Attribute VB_Name = "indx_Main"
Option Explicit

' =============================================================================
' indx_Main.bas
' Orchestrator for August salary indexation
' Uses indx_IndexationEngine for the heavy lifting
' =============================================================================

Sub indx_Main()
    ' Main procedure for August salary corrections
    ' Delegates to indx_IndexationEngine with August configuration
    
    Dim cfg As IndxConfig
    
    ' Create August-specific configuration
    cfg = indx_CreateAugustConfig()
    
    ' Run indexation using the generic engine
    Call indx_RunIndexation(cfg)
End Sub

' -----------------------------------------------------------------------------
' Legacy compatibility - keep old procedure name working
' -----------------------------------------------------------------------------
Private Sub indx_RemoveExistingIndexationSheet()
    ' Remove existing Indexation sheet if it exists
    ' Note: This is now handled by indx_Engine_RemoveExistingSheet in the engine
    ' Kept for backward compatibility if called directly
    
    Dim ws As Worksheet
    
    On Error Resume Next
    Set ws = Worksheets("Indexation")
    If Not ws Is Nothing Then
        Application.DisplayAlerts = False
        ws.Delete
        Application.DisplayAlerts = True
    End If
    On Error GoTo 0
End Sub
