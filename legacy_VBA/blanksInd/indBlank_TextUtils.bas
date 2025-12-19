Attribute VB_Name = "indBlank_TextUtils"
Option Explicit

' =============================================================================
' Module: indBlank_TextUtils
' Purpose: Text manipulation utilities for individual blank generation
' =============================================================================

' =============================================================================
' Function: indBlank_TrimAndCollapseSpaces
' Purpose: Trims leading/trailing spaces and collapses multiple spaces to single
' Param:   s - input string
' Returns: Cleaned string with single spaces only
' =============================================================================
Public Function indBlank_TrimAndCollapseSpaces(ByVal s As String) As String
    Dim result As String
    Dim prevLen As Long
    
    result = Trim$(s)
    
    ' Collapse multiple spaces by repeated replacement until stable
    Do
        prevLen = Len(result)
        result = Replace(result, "  ", " ")
    Loop While Len(result) < prevLen
    
    indBlank_TrimAndCollapseSpaces = result
End Function

' =============================================================================
' Function: indBlank_NormalizeForCompare
' Purpose: Normalizes string for comparison (uppercase + trim/collapse spaces)
' Param:   s - input string
' Returns: Normalized uppercase string for comparison
' =============================================================================
Public Function indBlank_NormalizeForCompare(ByVal s As String) As String
    indBlank_NormalizeForCompare = UCase$(indBlank_TrimAndCollapseSpaces(s))
End Function

' =============================================================================
' Function: indBlank_CreateSafeFileName
' Purpose: Creates a safe file name by removing invalid characters
'          and replacing spaces with underscores
' Param:   s - input string (proposed file name without extension)
' Returns: Safe file name string
' =============================================================================
Public Function indBlank_CreateSafeFileName(ByVal s As String) As String
    Dim result As String
    Dim i As Long
    Dim c As String
    Const INVALID_CHARS As String = "\/:*?""<>|"
    
    result = indBlank_TrimAndCollapseSpaces(s)
    
    ' Remove invalid file name characters
    For i = 1 To Len(INVALID_CHARS)
        c = Mid$(INVALID_CHARS, i, 1)
        result = Replace(result, c, "")
    Next i
    
    ' Replace spaces with underscores
    result = Replace(result, " ", "_")
    
    ' Collapse multiple underscores
    Do While InStr(1, result, "__") > 0
        result = Replace(result, "__", "_")
    Loop
    
    ' Trim leading/trailing underscores
    Do While Len(result) > 0 And Left$(result, 1) = "_"
        result = Mid$(result, 2)
    Loop
    Do While Len(result) > 0 And Right$(result, 1) = "_"
        result = Left$(result, Len(result) - 1)
    Loop
    
    If Len(result) = 0 Then
        result = "Unnamed"
    End If
    
    indBlank_CreateSafeFileName = result
End Function

' =============================================================================
' Function: indBlank_CreateSafeFolderName
' Purpose: Creates a safe folder name by removing invalid characters
'          and replacing spaces with underscores
' Param:   s - input string (proposed folder name)
' Returns: Safe folder name string
' =============================================================================
Public Function indBlank_CreateSafeFolderName(ByVal s As String) As String
    ' Folder name rules are same as file name rules in Windows
    indBlank_CreateSafeFolderName = indBlank_CreateSafeFileName(s)
End Function

' =============================================================================
' Function: indBlank_FormatChildFullName
' Purpose: Formats child name from raw data (replaces separators with spaces,
'          collapses spaces, returns "Nachname Vorname..." format)
' Param:   childRaw - raw child name string (may contain ; , . as separators)
' Returns: Formatted full name "Nachname Vorname..."
' =============================================================================
Public Function indBlank_FormatChildFullName(ByVal childRaw As String) As String
    Dim result As String
    
    result = childRaw
    
    ' Replace common separators with spaces
    result = Replace(result, ";", " ")
    result = Replace(result, ",", " ")
    result = Replace(result, ".", " ")
    
    ' Collapse multiple spaces and trim
    result = indBlank_TrimAndCollapseSpaces(result)
    
    indBlank_FormatChildFullName = result
End Function
