Attribute VB_Name = "History_ParseData"
'==========================
'   Module: History_ParseData
'   Purpose: Universal history string parser supporting both legacy and new formats
'   SELF-CONTAINED for Data file - does not depend on Admin or Superadmin modules
'
'   Supported Formats:
'   - Legacy: Mnt.N: War(X); Ist(Y). /Comment/ DD.MM.YYYY ||
'             Address: Was(X); Is(Y). /Comment/ DD.MM.YYYY ||
'             Subject1: Was(X); Is(Y). /Comment/ DD.MM.YYYY ||
'             Subject2: Was(X); Is(Y). /Comment/ DD.MM.YYYY ||
'             Ruck: Mnt.N: War(X); Ist(Y). /Comment/ DD.MM.YYYY ||
'             Decl_N: Was(); Is(comment). ||
'
'   - New: [RUCK:]<TAG>(<OLD>-><NEW>);.../@<COMMENT>@/<DATE>||
'          DCL(<N>-><comment>)||
'
'   Tags: FID, PAR, CHD, DOB, ADR, TEL, MOB, EML, SB1, PR1, SB2, PR2, M01-M12, EX1-EX3, DCL
'
'   Returns: Collection of Dictionary objects with keys:
'       - IsRuck (Boolean): True if retroactive change
'       - Reason (String): Comment/reason for change
'       - ChangeDate (String): Date in DD.MM.YYYY format
'       - Changes (Dictionary): key = field identifier, value = Dictionary with "War"/"Ist" keys
'
'   Field identifiers in Changes dictionary:
'       - Months: "1" to "12" (string numbers)
'       - Text fields: "Address", "Subject1", "Subject2", "FID", "PAR", "CHD", "DOB",
'                      "TEL", "MOB", "EML", "PR1", "PR2", "EX1", "EX2", "EX3"
'==========================

Option Explicit

' ========================================
' Public Interface
' ========================================

' Parse history string into collection of events
' Compatible with legacy Parceing.ParseHistory and Superadmin valid_ParseHistory.ParseHistory
Public Function ParseHistory(ByVal historyString As String) As Collection
    On Error GoTo ErrHandler
    
    Dim result As Collection
    Set result = New Collection
    
    If Len(Trim(historyString)) = 0 Then
        Set ParseHistory = result
        Exit Function
    End If
    
    Dim eventsArray() As String
    Dim i As Long
    
    ' Split the input string by "||" (session separator)
    eventsArray = Split(historyString, "||")
    
    Dim segment As String
    
    For i = LBound(eventsArray) To UBound(eventsArray)
        segment = Trim(eventsArray(i))
        
        ' Skip empty segments
        If Len(segment) = 0 Then GoTo SkipSegment
        
        ' Detect format and parse accordingly
        Dim evt As Object
        If IsNewFormatSegment(segment) Then
            Set evt = ParseNewFormat(segment)
        Else
            Set evt = ParseLegacyFormat(segment)
        End If
        
        ' Add event if valid (has date and changes)
        If Not evt Is Nothing Then
            result.Add evt
        End If
        
SkipSegment:
    Next i
    
    Set ParseHistory = result
    Exit Function
    
ErrHandler:
    Debug.Print "ERROR in History_ParseData.ParseHistory: Err=" & Err.Number & " - " & Err.Description
    Err.Raise Err.Number, "History_ParseData.ParseHistory", Err.Description
End Function

' ========================================
' Format Detection
' ========================================

Private Function IsNewFormatSegment(ByVal segment As String) As Boolean
    ' New format indicators:
    ' - Contains -> (value separator)
    ' - Contains /@ or @/ (comment delimiters)
    ' - Starts with DCL( or RUCK: followed by new-style tag
    ' - Contains field tags like M01(, ADR(, SB1( etc.
    
    ' Check for new format markers
    If InStr(segment, "->") > 0 Then
        IsNewFormatSegment = True
        Exit Function
    End If
    
    If InStr(segment, "/@") > 0 Or InStr(segment, "@/") > 0 Then
        IsNewFormatSegment = True
        Exit Function
    End If
    
    ' Check for DCL( at start (decline entry in new format)
    If Left(segment, 4) = "DCL(" Then
        IsNewFormatSegment = True
        Exit Function
    End If
    
    ' Check for RUCK: prefix with new format after it
    If UCase(Left(segment, 5)) = "RUCK:" Then
        Dim afterRuck As String
        afterRuck = Trim(Mid(segment, 6))
        If InStr(afterRuck, "->") > 0 Then
            IsNewFormatSegment = True
            Exit Function
        End If
    End If
    
    ' Check for new format field tags (M01(, ADR(, etc.)
    Dim newTagPattern As Object
    Set newTagPattern = CreateObject("VBScript.RegExp")
    With newTagPattern
        .Pattern = "(M\d{2}|FID|PAR|CHD|DOB|ADR|TEL|MOB|EML|SB1|PR1|SB2|PR2|EX[123])\("
        .IgnoreCase = True
        .Global = False
    End With
    
    If newTagPattern.Test(segment) Then
        IsNewFormatSegment = True
        Exit Function
    End If
    
    IsNewFormatSegment = False
End Function

' ========================================
' New Format Parser
' ========================================

Private Function ParseNewFormat(ByVal segment As String) As Object
    On Error GoTo ErrHandler
    
    ' Parse new format: [RUCK:]<TAG>(<OLD>-><NEW>);.../@<COMMENT>@/<DATE>
    '                   DCL(<N>-><comment>) - skipped for history view
    
    Dim evt As Object
    Set evt = CreateObject("Scripting.Dictionary")
    
    Dim changesDict As Object
    Set changesDict = CreateObject("Scripting.Dictionary")
    
    Dim isRuck As Boolean
    isRuck = False
    
    Dim reason As String
    reason = ""
    
    Dim changeDate As String
    changeDate = ""
    
    ' Check for RUCK: prefix (case-insensitive)
    If UCase(Left(segment, 5)) = "RUCK:" Then
        isRuck = True
        segment = Trim(Mid(segment, 6))
    End If
    
    ' Check for DCL( decline entry - skip for history view (return Nothing)
    If Left(segment, 4) = "DCL(" Then
        Set ParseNewFormat = Nothing
        Exit Function
    End If
    
    ' Extract comment and date: /@<COMMENT>@/<DATE>
    Dim commentDateRegex As Object
    Set commentDateRegex = CreateObject("VBScript.RegExp")
    With commentDateRegex
        .Pattern = "/@(.*)@/(\d{1,2}\.\d{1,2}\.\d{4})\s*$"
        .IgnoreCase = True
        .Global = False
    End With
    
    Dim cdMatches As Object
    Set cdMatches = commentDateRegex.Execute(segment)
    
    If cdMatches.Count > 0 Then
        reason = cdMatches(0).SubMatches(0)
        changeDate = cdMatches(0).SubMatches(1)
        ' Remove comment+date from segment
        segment = RTrim(Replace(segment, cdMatches(0).Value, ""))
    Else
        ' No date found - invalid segment for history
        Set ParseNewFormat = Nothing
        Exit Function
    End If
    
    ' Parse field changes: TAG(OLD->NEW);TAG(OLD->NEW);...
    Dim fieldRegex As Object
    Set fieldRegex = CreateObject("VBScript.RegExp")
    With fieldRegex
        ' Match: TAG(value->value) where TAG can be M01-M12, FID, PAR, etc.
        .Pattern = "(M\d{2}|FID|PAR|CHD|DOB|ADR|TEL|MOB|EML|SB1|PR1|SB2|PR2|EX[123])\(([^)]*)->([^)]*)\)"
        .IgnoreCase = True
        .Global = True
    End With
    
    Dim fieldMatches As Object
    Set fieldMatches = fieldRegex.Execute(segment)
    
    Dim fieldMatch As Object
    For Each fieldMatch In fieldMatches
        Dim tag As String
        Dim oldVal As String
        Dim newVal As String
        
        tag = fieldMatch.SubMatches(0)
        oldVal = fieldMatch.SubMatches(1)
        newVal = fieldMatch.SubMatches(2)
        
        ' Convert tag to unified key format
        Dim unifiedKey As String
        unifiedKey = ConvertNewTagToKey(tag)
        
        ' Create change dictionary
        Dim singleChange As Object
        Set singleChange = CreateObject("Scripting.Dictionary")
        singleChange.Add "War", oldVal
        singleChange.Add "Ist", newVal
        
        ' Use assignment instead of Add to handle duplicate keys (last value wins)
        If changesDict.Exists(unifiedKey) Then
            Set changesDict(unifiedKey) = singleChange
        Else
            changesDict.Add unifiedKey, singleChange
        End If
    Next fieldMatch
    
    ' Build result event
    evt.Add "IsRuck", isRuck
    evt.Add "Reason", reason
    evt.Add "ChangeDate", changeDate
    evt.Add "Changes", changesDict
    
    Set ParseNewFormat = evt
    Exit Function
    
ErrHandler:
    Debug.Print "ERROR in History_ParseData.ParseNewFormat: Err=" & Err.Number & " - " & Err.Description
    Err.Raise Err.Number, "History_ParseData.ParseNewFormat", Err.Description
End Function

' Convert new format tag to unified key used in Changes dictionary
Private Function ConvertNewTagToKey(ByVal tag As String) As String
    ' Month tags M01-M12 -> "1"-"12" (string numbers for compatibility)
    If UCase(Left(tag, 1)) = "M" And Len(tag) = 3 Then
        Dim monthNum As Integer
        On Error Resume Next
        monthNum = CInt(Mid(tag, 2, 2))
        On Error GoTo 0
        
        If monthNum >= 1 And monthNum <= 12 Then
            ConvertNewTagToKey = CStr(monthNum)
            Exit Function
        End If
    End If
    
    ' Field tags -> descriptive names (for compatibility with Geschichte display)
    Select Case UCase(tag)
        Case "FID"
            ConvertNewTagToKey = "FID"
        Case "PAR"
            ConvertNewTagToKey = "PAR"
        Case "CHD"
            ConvertNewTagToKey = "CHD"
        Case "DOB"
            ConvertNewTagToKey = "DOB"
        Case "ADR"
            ConvertNewTagToKey = "Address"
        Case "TEL"
            ConvertNewTagToKey = "TEL"
        Case "MOB"
            ConvertNewTagToKey = "MOB"
        Case "EML"
            ConvertNewTagToKey = "EML"
        Case "SB1"
            ConvertNewTagToKey = "Subject1"
        Case "PR1"
            ConvertNewTagToKey = "PR1"
        Case "SB2"
            ConvertNewTagToKey = "Subject2"
        Case "PR2"
            ConvertNewTagToKey = "PR2"
        Case "EX1"
            ConvertNewTagToKey = "EX1"
        Case "EX2"
            ConvertNewTagToKey = "EX2"
        Case "EX3"
            ConvertNewTagToKey = "EX3"
        Case Else
            ' Return as-is for unknown tags
            ConvertNewTagToKey = UCase(tag)
    End Select
End Function

' ========================================
' Legacy Format Parser
' ========================================

Private Function ParseLegacyFormat(ByVal segment As String) As Object
    On Error GoTo ErrHandler
    
    ' Parse legacy formats:
    ' - Mnt.N: War(X); Ist(Y). /Comment/ DD.MM.YYYY
    ' - Address: Was(X); Is(Y). /Comment/ DD.MM.YYYY
    ' - Subject1: Was(X); Is(Y). /Comment/ DD.MM.YYYY
    ' - Subject2: Was(X); Is(Y). /Comment/ DD.MM.YYYY
    ' - Ruck: Mnt.N: War(X); Ist(Y). /Comment/ DD.MM.YYYY
    ' - Decl_N: Was(); Is(comment). - skipped for history view
    
    Dim evt As Object
    Set evt = CreateObject("Scripting.Dictionary")
    
    Dim changesDict As Object
    Set changesDict = CreateObject("Scripting.Dictionary")
    
    Dim isRuck As Boolean
    isRuck = False
    
    Dim reason As String
    reason = ""
    
    Dim changeDate As String
    changeDate = ""
    
    ' Check for Decl_ prefix - skip for history view
    If Left(segment, 5) = "Decl_" Then
        Set ParseLegacyFormat = Nothing
        Exit Function
    End If
    
    ' Check for Ruck: prefix
    If Left(segment, 5) = "Ruck:" Then
        isRuck = True
        segment = Trim(Mid(segment, 6))
    End If
    
    ' Extract reason (in slashes) and date using regex
    Dim reasonDateRegex As Object
    Set reasonDateRegex = CreateObject("VBScript.RegExp")
    With reasonDateRegex
        .Pattern = "(?:/(.*?)/\s*)?(\d{1,2}\.\d{1,2}\.\d{4})\s*$"
        .IgnoreCase = True
        .Global = False
    End With
    
    Dim rdMatches As Object
    Set rdMatches = reasonDateRegex.Execute(segment)
    
    If rdMatches.Count > 0 Then
        reason = rdMatches(0).SubMatches(0)
        changeDate = rdMatches(0).SubMatches(1)
        ' Remove reason+date from segment
        segment = RTrim(Replace(segment, rdMatches(0).Value, ""))
    Else
        ' No date found - invalid segment
        Set ParseLegacyFormat = Nothing
        Exit Function
    End If
    
    ' Parse month changes: Mnt.N: War(X); Ist(Y).
    Dim monthRegex As Object
    Set monthRegex = CreateObject("VBScript.RegExp")
    With monthRegex
        .Pattern = "Mnt\.(\d+):\s*War\(([^)]*)\);\s*Ist\(([^)]*)\)\."
        .IgnoreCase = True
        .Global = True
    End With
    
    Dim monthMatches As Object
    Set monthMatches = monthRegex.Execute(segment)
    
    Dim monthMatch As Object
    For Each monthMatch In monthMatches
        Dim monthNum As String
        Dim warVal As String
        Dim istVal As String
        
        monthNum = monthMatch.SubMatches(0)
        warVal = monthMatch.SubMatches(1)
        istVal = monthMatch.SubMatches(2)
        
        ' Create change dictionary for this month
        Dim monthChange As Object
        Set monthChange = CreateObject("Scripting.Dictionary")
        monthChange.Add "War", warVal
        monthChange.Add "Ist", istVal
        
        changesDict.Add monthNum, monthChange
    Next monthMatch
    
    ' Parse field changes: FieldName: Was(X); Is(Y).
    ' Supports: Address, Subject1, Subject2
    Dim fieldRegex As Object
    Set fieldRegex = CreateObject("VBScript.RegExp")
    With fieldRegex
        .Pattern = "(Address|Subject1|Subject2):\s*Was\(([^)]*)\);\s*Is\(([^)]*)\)\."
        .IgnoreCase = True
        .Global = True
    End With
    
    Dim fieldMatches As Object
    Set fieldMatches = fieldRegex.Execute(segment)
    
    Dim fieldMatch As Object
    For Each fieldMatch In fieldMatches
        Dim fieldName As String
        Dim fieldOld As String
        Dim fieldNew As String
        
        fieldName = fieldMatch.SubMatches(0)
        fieldOld = fieldMatch.SubMatches(1)
        fieldNew = fieldMatch.SubMatches(2)
        
        ' Convert to unified key
        Dim fieldKey As String
        fieldKey = ConvertLegacyFieldToKey(fieldName)
        
        ' Create change dictionary for this field
        Dim fieldChange As Object
        Set fieldChange = CreateObject("Scripting.Dictionary")
        fieldChange.Add "War", fieldOld
        fieldChange.Add "Ist", fieldNew
        
        If Not changesDict.Exists(fieldKey) Then
            changesDict.Add fieldKey, fieldChange
        End If
    Next fieldMatch
    
    ' If no changes found, return Nothing
    If changesDict.Count = 0 Then
        Set ParseLegacyFormat = Nothing
        Exit Function
    End If
    
    ' Build result event
    evt.Add "IsRuck", isRuck
    evt.Add "Reason", reason
    evt.Add "ChangeDate", changeDate
    evt.Add "Changes", changesDict
    
    Set ParseLegacyFormat = evt
    Exit Function
    
ErrHandler:
    Debug.Print "ERROR in History_ParseData.ParseLegacyFormat: Err=" & Err.Number & " - " & Err.Description
    Err.Raise Err.Number, "History_ParseData.ParseLegacyFormat", Err.Description
End Function

' Convert legacy field name to unified key
Private Function ConvertLegacyFieldToKey(ByVal fieldName As String) As String
    Select Case LCase(fieldName)
        Case "address"
            ConvertLegacyFieldToKey = "Address"
        Case "subject1"
            ConvertLegacyFieldToKey = "Subject1"
        Case "subject2"
            ConvertLegacyFieldToKey = "Subject2"
        Case Else
            ' Return as-is for unknown fields
            ConvertLegacyFieldToKey = fieldName
    End Select
End Function

' ========================================
' Utility Functions (for testing/debugging)
' ========================================

' Test function to verify parser works correctly
' Usage: Call History_ParseData.TestParser from Immediate window
Public Sub TestParser()
    Dim testCases As Variant
    Dim i As Long
    Dim result As Collection
    
    ' Test cases: legacy and new format samples
    testCases = Array( _
        "Mnt.1: War(100); Ist(150). /Erhoehung/ 15.01.2025", _
        "Ruck: Mnt.3: War(50); Ist(75). /Korrektur/ 20.02.2025", _
        "Address: Was(Alt Strasse 1); Is(Neu Strasse 2). /Umzug/ 01.03.2025", _
        "M01(100->150);M02(200->250);/@Indexation@/15.01.2025", _
        "RUCK:M03(50->75);/@Korrektur@/20.02.2025", _
        "ADR(Alt Strasse->Neu Strasse);TEL(123->456);/@Update@/01.03.2025", _
        "DCL(1->Bitte pruefen)", _
        "Decl_1: Was(); Is(Abgelehnt)." _
    )
    
    For i = LBound(testCases) To UBound(testCases)
        Debug.Print "=== Test " & (i + 1) & " ==="
        Debug.Print "Input: " & testCases(i)
        
        On Error Resume Next
        Set result = ParseHistory(CStr(testCases(i)))
        
        If Err.Number <> 0 Then
            Debug.Print "ERROR: " & Err.Description
            Err.Clear
        ElseIf result.Count = 0 Then
            Debug.Print "Result: (skipped or no valid events)"
        Else
            Dim evt As Object
            Set evt = result(1)
            Debug.Print "IsRuck: " & evt("IsRuck")
            Debug.Print "Reason: " & evt("Reason")
            Debug.Print "ChangeDate: " & evt("ChangeDate")
            Debug.Print "Changes count: " & evt("Changes").Count
            
            Dim key As Variant
            For Each key In evt("Changes").Keys
                Debug.Print "  " & key & ": War=" & evt("Changes")(key)("War") & ", Ist=" & evt("Changes")(key)("Ist")
            Next key
        End If
        On Error GoTo 0
        
        Debug.Print ""
    Next i
End Sub
