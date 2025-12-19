Attribute VB_Name = "Kind_TextUtils"
Option Explicit

' =============================================================================
' Module: Kind_TextUtils
' Purpose: Text normalization and comparison utilities for Nachhilfe validation
' Prefix: Kind_ (all modules/classes in this project use this prefix)
' =============================================================================

' -----------------------------------------------------------------------------
' Kind_TrimAndCollapseSpaces
' Purpose: Trims leading/trailing spaces and collapses multiple internal spaces
'          to single space
' Parameters:
'   s - Input string
' Returns: Normalized string with collapsed spaces
' -----------------------------------------------------------------------------
Public Function Kind_TrimAndCollapseSpaces(ByVal s As String) As String
    Dim result As String
    Dim prevSpace As Boolean
    Dim i As Long
    Dim ch As String
    
    ' Handle empty/null
    If Len(s) = 0 Then
        Kind_TrimAndCollapseSpaces = ""
        Exit Function
    End If
    
    ' First trim
    s = Trim$(s)
    
    ' Collapse multiple spaces to single
    result = ""
    prevSpace = False
    
    For i = 1 To Len(s)
        ch = Mid$(s, i, 1)
        If ch = " " Then
            If Not prevSpace Then
                result = result & ch
                prevSpace = True
            End If
            ' Skip additional spaces
        Else
            result = result & ch
            prevSpace = False
        End If
    Next i
    
    Kind_TrimAndCollapseSpaces = result
End Function


' -----------------------------------------------------------------------------
' Kind_NormalizeForCompare
' Purpose: Normalizes string for comparison: uppercase, trim, collapse spaces
' Parameters:
'   s - Input string
' Returns: Normalized uppercase string ready for comparison
' -----------------------------------------------------------------------------
Public Function Kind_NormalizeForCompare(ByVal s As String) As String
    Kind_NormalizeForCompare = UCase$(Kind_TrimAndCollapseSpaces(s))
End Function


' -----------------------------------------------------------------------------
' Kind_NormalizeSubjectKinder
' Purpose: Normalizes subject name from Kinder sheet for comparison
' Parameters:
'   s - Raw subject string from Kinder sheet
' Returns: Normalized subject string
' -----------------------------------------------------------------------------
Public Function Kind_NormalizeSubjectKinder(ByVal s As String) As String
    Kind_NormalizeSubjectKinder = Kind_NormalizeForCompare(s)
End Function


' -----------------------------------------------------------------------------
' Kind_TryNormalizeSubjectKartei
' Purpose: Attempts to extract and normalize base subject from Kartei format
'          Only processes subjects with "NACHHILFE" or "NH" prefix
' Parameters:
'   s              - Raw subject string from Kartei sheet (e.g. "Nachhilfe Mathe")
'   outBaseSubject - (ByRef) Output normalized base subject without prefix
' Returns: True if valid Nachhilfe subject found and parsed, False otherwise
' Examples:
'   "Nachhilfe Mathe"    -> True, outBaseSubject = "MATHE"
'   "NACHHILFE  Deutsch" -> True, outBaseSubject = "DEUTSCH"
'   "NH Englisch"        -> True, outBaseSubject = "ENGLISCH"
'   "NHPhysik"           -> True, outBaseSubject = "PHYSIK"
'   "Klavierunterricht"  -> False (not a Nachhilfe subject)
'   "Unterricht Mathe"   -> False (wrong prefix)
' -----------------------------------------------------------------------------
Public Function Kind_TryNormalizeSubjectKartei(ByVal s As String, ByRef outBaseSubject As String) As Boolean
    Const PREFIX_FULL As String = "NACHHILFE"
    Const PREFIX_SHORT As String = "NH"
    
    Dim normalized As String
    Dim baseSubject As String
    Dim prefixLen As Long
    Dim spacePos As Long
    Dim lastCh As String
    
    ' Initialize output
    outBaseSubject = ""
    Kind_TryNormalizeSubjectKartei = False
    
    ' Normalize input: trim and collapse spaces, uppercase for comparison
    normalized = Kind_NormalizeForCompare(s)
    
    ' Check for empty after normalization
    If Len(normalized) = 0 Then
        Exit Function
    End If
    
    ' Check for "NACHHILFE" prefix first (longer match takes priority)
    If Left$(normalized, Len(PREFIX_FULL)) = PREFIX_FULL Then
        prefixLen = Len(PREFIX_FULL)
        baseSubject = Mid$(normalized, prefixLen + 1)
        
    ' Check for "NH" prefix
    ElseIf Left$(normalized, Len(PREFIX_SHORT)) = PREFIX_SHORT Then
        prefixLen = Len(PREFIX_SHORT)
        baseSubject = Mid$(normalized, prefixLen + 1)
        
    Else
        ' No valid prefix found - not a Nachhilfe subject
        Exit Function
    End If
    
    ' Trim any leading spaces from base subject (prefix may be followed by spaces)
    baseSubject = LTrim$(baseSubject)
    
    ' Base subject must not be empty
    If Len(baseSubject) = 0 Then
        Exit Function
    End If

    ' Kartei may contain extra tokens after the subject, e.g. "NH Deutsch Mo 3kl".
    ' For matching, only the first token (subject name) is relevant.
    spacePos = InStr(1, baseSubject, " ", vbTextCompare)
    If spacePos > 0 Then
        baseSubject = Left$(baseSubject, spacePos - 1)
    End If

    ' Remove trailing punctuation from the token (e.g. "DEUTSCH," -> "DEUTSCH").
    Do While Len(baseSubject) > 0
        lastCh = Right$(baseSubject, 1)
        If (lastCh Like "[A-Z0-9]") Then
            Exit Do
        End If
        baseSubject = Left$(baseSubject, Len(baseSubject) - 1)
    Loop

    If Len(baseSubject) = 0 Then
        Exit Function
    End If
    
    ' Success - return normalized base subject
    outBaseSubject = baseSubject
    Kind_TryNormalizeSubjectKartei = True
End Function
