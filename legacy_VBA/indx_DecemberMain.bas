Attribute VB_Name = "indx_DecemberMain"
Option Explicit

' =============================================================================
' indx_DecemberMain.bas
' Orchestrator for December salary indexation
' Uses indx_IndexationEngine for the heavy lifting
' =============================================================================

Sub indx_DecemberMain()
    ' Main procedure for December salary corrections
    ' Recalculates December (AF) values using coefficient 0.75
    ' Base and source are both AF (December uses its own value as base)
    ' Delegates to indx_IndexationEngine with December configuration
    
    Dim cfg As IndxConfig
    
    ' Create December-specific configuration
    cfg = indx_CreateDecemberConfig()
    
    ' Run indexation using the generic engine
    Call indx_RunIndexation(cfg)
End Sub
