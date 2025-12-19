Attribute VB_Name = "Kind_NameUtils"
Option Explicit

' =============================================================================
' Module: Kind_NameUtils
' Purpose: Name parsing and comparison utilities for matching Kinder to Kartei
' Prefix: Kind_ (all modules/classes in this project use this prefix)
'
' Kartei.Child formats supported:
'   - "Nachname ; Vorname"
'   - "Nachname, Vorname"
'   - "Nachname. Vorname"
'   - "Nachname Vorname" (space-separated fallback)
' =============================================================================

' -----------------------------------------------------------------------------
' Kind_ParseChild
' Purpose: Parses Kartei Child string into Last and First name components
' Parameters:
'   childRaw - Raw child string from Kartei (e.g. "Mueller ; Hans")
'   outLast  - (ByRef) Output last name (normalized)
'   outFirst - (ByRef) Output first name (normalized)
' Returns: True if successfully parsed, False if empty or unparseable
' Logic:
'   1. If delimiter (;  ,  .) found -> split at first occurrence
'   2. Otherwise split at first space (last name = first word, first name = rest)
' -----------------------------------------------------------------------------
Public Function Kind_ParseChild(ByVal childRaw As String, ByRef outLast As String, ByRef outFirst As String) As Boolean
    Dim normalized As String
    Dim delimPos As Long
    Dim delimChar As String
    Dim spacePos As Long
    
    ' Initialize outputs
    outLast = ""
    outFirst = ""
    Kind_ParseChild = False
    
    ' Normalize: trim and collapse spaces
    normalized = Kind_TrimAndCollapseSpaces(childRaw)
    
    ' Check for empty
    If Len(normalized) = 0 Then
        Exit Function
    End If
    
    ' Try to find delimiter: ; or , or .
    delimChar = ""
    delimPos = 0
    
    ' Check for semicolon first (most explicit)
    delimPos = InStr(1, normalized, ";", vbTextCompare)
    If delimPos > 0 Then
        delimChar = ";"
    End If
    
    ' If no semicolon, check comma
    If delimPos = 0 Then
        delimPos = InStr(1, normalized, ",", vbTextCompare)
        If delimPos > 0 Then
            delimChar = ","
        End If
    End If
    
    ' If no comma, check period
    If delimPos = 0 Then
        delimPos = InStr(1, normalized, ".", vbTextCompare)
        If delimPos > 0 Then
            delimChar = "."
        End If
    End If
    
    ' If delimiter found, split by it
    If delimPos > 0 Then
        outLast = Kind_NormalizeForCompare(Left$(normalized, delimPos - 1))
        outFirst = Kind_NormalizeForCompare(Mid$(normalized, delimPos + 1))
        
        ' Both parts must be non-empty for successful parse
        If Len(outLast) > 0 And Len(outFirst) > 0 Then
            Kind_ParseChild = True
        End If
        Exit Function
    End If
    
    ' No delimiter found - fallback to space logic
    ' Last name = first word, First name = everything after first space
    spacePos = InStr(1, normalized, " ", vbTextCompare)
    
    If spacePos > 0 Then
        outLast = Kind_NormalizeForCompare(Left$(normalized, spacePos - 1))
        outFirst = Kind_NormalizeForCompare(Mid$(normalized, spacePos + 1))
        
        ' Both parts must be non-empty
        If Len(outLast) > 0 And Len(outFirst) > 0 Then
            Kind_ParseChild = True
        End If
    Else
        ' Single word - treat as last name only
        outLast = Kind_NormalizeForCompare(normalized)
        outFirst = ""
        ' Return False because we couldn't determine first name
    End If
End Function


' -----------------------------------------------------------------------------
' Kind_IsSameChild
' Purpose: Compares Kartei child string with Kinder (Nachname, Vorname)
' Parameters:
'   karteiChildRaw - Raw child string from Kartei
'   kinderLast     - Last name from Kinder sheet
'   kinderFirst    - First name from Kinder sheet
' Returns: True if names match, False otherwise
' Logic:
'   1. Parse karteiChildRaw using Kind_ParseChild
'   2. If parse successful, compare parsed parts with Kinder names
'   3. If parse fails or uncertain, fallback: check if normalized karteiChild
'      contains both Nachname and Vorname as substrings
' -----------------------------------------------------------------------------
Public Function Kind_IsSameChild(ByVal karteiChildRaw As String, ByVal kinderLast As String, ByVal kinderFirst As String) As Boolean
    Dim parsedLast As String
    Dim parsedFirst As String
    Dim parseOk As Boolean
    Dim normalizedKartei As String
    Dim normalizedKinderLast As String
    Dim normalizedKinderFirst As String
    
    Kind_IsSameChild = False
    
    ' Normalize Kinder inputs
    normalizedKinderLast = Kind_NormalizeForCompare(kinderLast)
    normalizedKinderFirst = Kind_NormalizeForCompare(kinderFirst)
    
    ' Both Kinder name parts must be non-empty
    If Len(normalizedKinderLast) = 0 Or Len(normalizedKinderFirst) = 0 Then
        Exit Function
    End If
    
    ' Try to parse Kartei child
    parseOk = Kind_ParseChild(karteiChildRaw, parsedLast, parsedFirst)
    
    If parseOk Then
        ' Strict comparison: both parts must match exactly
        If parsedLast = normalizedKinderLast And parsedFirst = normalizedKinderFirst Then
            Kind_IsSameChild = True
            Exit Function
        End If
        
        ' Handle multiple first names: parsed first may contain Kinder first or vice versa
        ' e.g. Kinder: "Anna Maria" vs Kartei: "Mueller ; Anna Maria"
        ' or Kinder: "Anna" vs Kartei: "Mueller ; Anna Maria" (partial match)
        If parsedLast = normalizedKinderLast Then
            ' Check if first names are related (one contains the other)
            If parsedFirst = normalizedKinderFirst Then
                Kind_IsSameChild = True
                Exit Function
            ElseIf InStr(1, parsedFirst, normalizedKinderFirst, vbTextCompare) > 0 Then
                ' Kartei first name contains Kinder first name
                Kind_IsSameChild = True
                Exit Function
            ElseIf InStr(1, normalizedKinderFirst, parsedFirst, vbTextCompare) > 0 Then
                ' Kinder first name contains parsed first name
                Kind_IsSameChild = True
                Exit Function
            End If
        End If
    End If
    
    ' Fallback: substring matching
    ' Normalized karteiChild should contain both Nachname and Vorname
    normalizedKartei = Kind_NormalizeForCompare(karteiChildRaw)
    
    If Len(normalizedKartei) = 0 Then
        Exit Function
    End If
    
    ' Check if kartei contains both name parts as substrings
    If InStr(1, normalizedKartei, normalizedKinderLast, vbTextCompare) > 0 Then
        If InStr(1, normalizedKartei, normalizedKinderFirst, vbTextCompare) > 0 Then
            Kind_IsSameChild = True
        End If
    End If
End Function
