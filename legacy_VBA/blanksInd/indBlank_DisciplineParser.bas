Attribute VB_Name = "indBlank_DisciplineParser"
Option Explicit

' =============================================================================
' Module: indBlank_DisciplineParser
' Purpose: Parses discipline values from Kartei columns (J/O) to extract
'          the discipline name for template generation
' =============================================================================

' =============================================================================
' Function: indBlank_TryParseDiscipline
' Purpose: Attempts to parse a raw discipline value and extract discipline name
' Param:   rawValue - raw value from discipline column (J or O)
' Param:   outDiscipline - output discipline name for template
' Returns: True if this is an Ind/VSpE discipline, False otherwise
'
' Rules:
'   - "VSPE_..." -> outDiscipline = "VSpE", True
'   - "IND..." or "IND. ..." -> parse first token after prefix, map to full name
'   - Anything else -> False (not for Ind blanks)
' =============================================================================
Public Function indBlank_TryParseDiscipline(ByVal rawValue As String, ByRef outDiscipline As String) As Boolean
    Dim normalizedValue As String
    Dim afterPrefix As String
    Dim firstToken As String
    Dim spacePos As Long
    
    outDiscipline = ""
    indBlank_TryParseDiscipline = False
    
    ' Trim and normalize input
    normalizedValue = Trim$(rawValue)
    If Len(normalizedValue) = 0 Then
        Exit Function
    End If
    
    ' Check for VSPE_ prefix (case insensitive)
    If UCase$(Left$(normalizedValue, 5)) = "VSPE_" Then
        outDiscipline = "VSpE"
        indBlank_TryParseDiscipline = True
        Exit Function
    End If
    
    ' Check for IND prefix (case insensitive)
    If Not StartsWithInd(normalizedValue) Then
        ' Not an Ind discipline - skip
        Exit Function
    End If
    
    ' Remove IND prefix and extract remaining part
    afterPrefix = RemoveIndPrefix(normalizedValue)
    
    ' Get first token (before space)
    afterPrefix = Trim$(afterPrefix)
    If Len(afterPrefix) = 0 Then
        Exit Function
    End If
    
    spacePos = InStr(1, afterPrefix, " ")
    If spacePos > 0 Then
        firstToken = Left$(afterPrefix, spacePos - 1)
    Else
        firstToken = afterPrefix
    End If
    
    ' Remove trailing dot if present
    If Right$(firstToken, 1) = "." Then
        firstToken = Left$(firstToken, Len(firstToken) - 1)
    End If
    
    firstToken = Trim$(firstToken)
    If Len(firstToken) = 0 Then
        Exit Function
    End If
    
    ' Map token to full discipline name
    outDiscipline = MapDisciplineToken(firstToken)
    indBlank_TryParseDiscipline = True
End Function

' =============================================================================
' Private Helper: StartsWithInd
' Purpose: Checks if value starts with "IND" (case insensitive)
' =============================================================================
Private Function StartsWithInd(ByVal s As String) As Boolean
    Dim upper As String
    upper = UCase$(s)
    
    ' Check for "IND " or "IND." or just "IND" at start
    If Left$(upper, 4) = "IND " Then
        StartsWithInd = True
    ElseIf Left$(upper, 4) = "IND." Then
        StartsWithInd = True
    ElseIf Left$(upper, 3) = "IND" And Len(s) >= 3 Then
        ' Check if next char is space, dot, or we're at exact "IND"
        If Len(s) = 3 Then
            StartsWithInd = True
        Else
            Dim nextChar As String
            nextChar = Mid$(s, 4, 1)
            ' Accept if next char is space, dot, or letter (for cases like "IndMa")
            StartsWithInd = True
        End If
    Else
        StartsWithInd = False
    End If
End Function

' =============================================================================
' Private Helper: RemoveIndPrefix
' Purpose: Removes "IND", "IND." or "IND " prefix from the string
' =============================================================================
Private Function RemoveIndPrefix(ByVal s As String) As String
    Dim upper As String
    Dim result As String
    
    upper = UCase$(s)
    result = s
    
    ' Remove "IND." first (4 chars)
    If Left$(upper, 4) = "IND." Then
        result = Mid$(s, 5)
    ' Then check "IND " (4 chars)
    ElseIf Left$(upper, 4) = "IND " Then
        result = Mid$(s, 5)
    ' Then just "IND" (3 chars)
    ElseIf Left$(upper, 3) = "IND" Then
        result = Mid$(s, 4)
    End If
    
    ' Remove leading spaces and dots
    Do While Len(result) > 0
        If Left$(result, 1) = " " Or Left$(result, 1) = "." Then
            result = Mid$(result, 2)
        Else
            Exit Do
        End If
    Loop
    
    RemoveIndPrefix = result
End Function

' =============================================================================
' Private Helper: MapDisciplineToken
' Purpose: Maps abbreviated discipline token to full German name
' Param:   token - abbreviated discipline (e.g., "Ma", "Ru")
' Returns: Full discipline name or token as fallback
' =============================================================================
Private Function MapDisciplineToken(ByVal token As String) As String
    Dim upperToken As String
    upperToken = UCase$(Trim$(token))
    
    Select Case upperToken
        Case "MA"
            MapDisciplineToken = "Mathe"
        Case "RU"
            MapDisciplineToken = "Russisch"
        Case "ENG"
            MapDisciplineToken = "Englisch"
        Case "DE"
            MapDisciplineToken = "Deutsch"
        Case Else
            ' Fallback: return token as-is (e.g., "Sachunt")
            MapDisciplineToken = token
    End Select
End Function
