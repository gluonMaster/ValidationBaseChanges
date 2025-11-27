Attribute VB_Name = "valid_ParseHistory"
'==========================
'   Module: valid_ParseHistory
'   Purpose: Universal history string parser supporting both legacy and new formats
'   SELF-CONTAINED for Superadmin file - does not depend on Admin modules
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
'==========================

Option Explicit

' Parse history string into collection of events
' Returns: Collection of Dictionary objects with keys: IsRuck, Reason, ChangeDate, Changes
' Changes is a Dictionary: key = field identifier (month number or field name), value = Dictionary with "War"/"Ist" keys
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
        
        Debug.Print "ParseHistory: segment " & i & " = '" & Left(segment, 50) & "...'"
        
        ' Detect format and parse accordingly
        Dim evt As Object
        If IsNewFormatSegment(segment) Then
            Debug.Print "ParseHistory: detected NEW format"
            Set evt = ParseNewFormat(segment)
        Else
            Debug.Print "ParseHistory: detected LEGACY format"
            Set evt = ParseLegacyFormat(segment)
        End If
        
        ' Add event if valid (has date or is decline)
        If Not evt Is Nothing Then
            result.Add evt
        End If
        
SkipSegment:
    Next i
    
    Set ParseHistory = result
    Exit Function
    
ErrHandler:
    Debug.Print "ERROR in ParseHistory: Err=" & Err.Number & " - " & Err.Description & " | segment=" & segment
    Err.Raise Err.Number, "valid_ParseHistory.ParseHistory", Err.Description
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
    
    ' Check for DCL( at start
    If Left(segment, 4) = "DCL(" Then
        IsNewFormatSegment = True
        Exit Function
    End If
    
    ' Check for RUCK: prefix with new format after it
    If Left(segment, 5) = "RUCK:" Then
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
    '                   DCL(<N>-><comment>)
    
    Debug.Print "ParseNewFormat: ENTER, segment='" & Left(segment, 80) & "...'"
    
    Dim evt As Object
    Set evt = CreateObject("Scripting.Dictionary")
    
    Dim changesDict As Object
    Set changesDict = CreateObject("Scripting.Dictionary")
    
    Dim isRuck As Boolean
    isRuck = False
    
    Dim Reason As String
    Reason = ""
    
    Dim changeDate As String
    changeDate = ""
    
    ' Check for RUCK: prefix
    If Left(segment, 5) = "RUCK:" Then
        isRuck = True
        segment = Trim(Mid(segment, 6))
    End If
    
    ' Check for DCL( decline entry - skip for GrossGeschichte (handled separately)
    If Left(segment, 4) = "DCL(" Then
        ' Parse DCL(N->comment) but don't add to regular events
        ' Return Nothing to skip this segment in GrossGeschichte
        Set ParseNewFormat = Nothing
        Exit Function
    End If
    
    ' Extract comment and date: /@<COMMENT>@/<DATE>
    Debug.Print "ParseNewFormat: creating commentDateRegex"
    Dim commentDateRegex As Object
    Set commentDateRegex = CreateObject("VBScript.RegExp")
    With commentDateRegex
        .Pattern = "/@(.*)@/(\d{1,2}\.\d{1,2}\.\d{4})\s*$"
        .IgnoreCase = True
        .Global = False
    End With
    
    Debug.Print "ParseNewFormat: executing commentDateRegex"
    Dim cdMatches As Object
    Set cdMatches = commentDateRegex.Execute(segment)
    
    Debug.Print "ParseNewFormat: cdMatches.Count=" & cdMatches.Count
    If cdMatches.Count > 0 Then
        Debug.Print "ParseNewFormat: accessing SubMatches for comment/date"
        Reason = cdMatches(0).SubMatches(0)
        changeDate = cdMatches(0).SubMatches(1)
        Debug.Print "ParseNewFormat: Reason='" & Reason & "', changeDate='" & changeDate & "'"
        ' Remove comment+date from segment
        segment = RTrim(Replace(segment, cdMatches(0).Value, ""))
    Else
        ' No date found - invalid segment
        Set ParseNewFormat = Nothing
        Exit Function
    End If
    
    ' Parse field changes: TAG(OLD->NEW);TAG(OLD->NEW);...
    Debug.Print "ParseNewFormat: creating fieldRegex"
    Dim fieldRegex As Object
    Set fieldRegex = CreateObject("VBScript.RegExp")
    With fieldRegex
        ' Match: TAG(value->value) where TAG can be M01-M12, FID, PAR, etc.
        .Pattern = "(M\d{2}|FID|PAR|CHD|DOB|ADR|TEL|MOB|EML|SB1|PR1|SB2|PR2|EX[123])\(([^)]*)->([^)]*)\)"
        .IgnoreCase = True
        .Global = True
    End With
    
    Debug.Print "ParseNewFormat: executing fieldRegex"
    Dim fieldMatches As Object
    Set fieldMatches = fieldRegex.Execute(segment)
    
    Debug.Print "ParseNewFormat: fieldMatches.Count=" & fieldMatches.Count
    Dim fieldMatch As Object
    Dim matchIndex As Long
    matchIndex = 0
    For Each fieldMatch In fieldMatches
        Debug.Print "ParseNewFormat: processing match " & matchIndex
        Dim tag As String
        Dim oldVal As String
        Dim newVal As String
        
        Debug.Print "ParseNewFormat: accessing SubMatches(0)"
        tag = fieldMatch.SubMatches(0)
        Debug.Print "ParseNewFormat: accessing SubMatches(1)"
        oldVal = fieldMatch.SubMatches(1)
        Debug.Print "ParseNewFormat: accessing SubMatches(2)"
        newVal = fieldMatch.SubMatches(2)
        
        Debug.Print "ParseNewFormat: tag='" & tag & "', old='" & oldVal & "', new='" & newVal & "'"
        
        ' Convert tag to unified key format
        Dim unifiedKey As String
        unifiedKey = ConvertNewTagToKey(tag)
        
        ' Create change dictionary
        Dim singleChange As Object
        Set singleChange = CreateObject("Scripting.Dictionary")
        singleChange.Add "War", oldVal
        singleChange.Add "Ist", newVal
        
        changesDict.Add unifiedKey, singleChange
        matchIndex = matchIndex + 1
    Next fieldMatch
    
    ' Build result event
    evt.Add "IsRuck", isRuck
    evt.Add "Reason", Reason
    evt.Add "ChangeDate", changeDate
    evt.Add "Changes", changesDict
    
    Debug.Print "ParseNewFormat: SUCCESS, returning evt"
    Set ParseNewFormat = evt
    Exit Function
    
ErrHandler:
    Debug.Print "ERROR in ParseNewFormat: Err=" & Err.Number & " - " & Err.Description & " | segment=" & segment
    Err.Raise Err.Number, "valid_ParseHistory.ParseNewFormat", Err.Description
End Function

' Convert new format tag to unified key used in Changes dictionary
Private Function ConvertNewTagToKey(ByVal tag As String) As String
    ' Month tags M01-M12 -> 1-12 (numeric string for compatibility with GrossGeschichte)
    If Left(UCase(tag), 1) = "M" And Len(tag) = 3 Then
        Dim monthNum As Integer
        On Error Resume Next
        monthNum = CInt(Mid(tag, 2, 2))
        On Error GoTo 0
        
        If monthNum >= 1 And monthNum <= 12 Then
            ConvertNewTagToKey = CStr(monthNum)
            Exit Function
        End If
    End If
    
    ' Field tags -> descriptive names (for Address, Subject1, Subject2)
    Select Case UCase(tag)
        Case "ADR"
            ConvertNewTagToKey = "Address"
        Case "SB1"
            ConvertNewTagToKey = "Subject1"
        Case "SB2"
            ConvertNewTagToKey = "Subject2"
        Case Else
            ' Return as-is for other fields
            ConvertNewTagToKey = UCase(tag)
    End Select
End Function

' ========================================
' Legacy Format Parser
' ========================================

Private Function ParseLegacyFormat(ByVal segment As String) As Object
    On Error GoTo ErrHandler
    
    ' Parse legacy format: Mnt.N: War(X); Ist(Y). /Comment/ DD.MM.YYYY
    '                      Address: Was(X); Is(Y). /Comment/ DD.MM.YYYY
    '                      Decl_N: Was(); Is(comment).
    
    Debug.Print "ParseLegacyFormat: ENTER, segment='" & Left(segment, 80) & "...'"
    
    Dim evt As Object
    Set evt = CreateObject("Scripting.Dictionary")
    
    Dim changesDict As Object
    Set changesDict = CreateObject("Scripting.Dictionary")
    
    Dim isRuck As Boolean
    isRuck = False
    
    ' Check for Ruck: prefix
    If Left(segment, 5) = "Ruck:" Then
        isRuck = True
        segment = Trim(Mid(segment, 6))
    End If
    
    ' Check for Decl_N: legacy decline format - skip for GrossGeschichte
    If Left(segment, 5) = "Decl_" Then
        Set ParseLegacyFormat = Nothing
        Exit Function
    End If
    
    ' Extract reason and date using two-step approach (VBScript.RegExp does NOT support (?:...) non-capturing groups)
    Dim Reason As String
    Reason = ""
    Dim changeDate As String
    changeDate = ""
    
    ' Step 1: Try to match /reason/ date pattern first
    Debug.Print "ParseLegacyFormat: creating regexWithReason"
    Dim regexWithReason As Object
    Set regexWithReason = CreateObject("VBScript.RegExp")
    With regexWithReason
        ' Match: /.../ DD.MM.YYYY at the end
        .Pattern = "/([^/]*)/\s*(\d{1,2}\.\d{1,2}\.\d{4})\s*$"
        .IgnoreCase = True
        .Global = False
    End With
    
    Debug.Print "ParseLegacyFormat: executing regexWithReason"
    Dim reasonMatches As Object
    Set reasonMatches = regexWithReason.Execute(segment)
    
    Debug.Print "ParseLegacyFormat: reasonMatches.Count=" & reasonMatches.Count
    If reasonMatches.Count > 0 Then
        Debug.Print "ParseLegacyFormat: accessing SubMatches for reason/date"
        Reason = reasonMatches(0).SubMatches(0)
        changeDate = reasonMatches(0).SubMatches(1)
        Debug.Print "ParseLegacyFormat: Reason='" & Reason & "', changeDate='" & changeDate & "'"
        ' Remove reason+date from the segment
        segment = RTrim(Replace(segment, reasonMatches(0).Value, ""))
    Else
        ' Step 2: Try to match just date (no reason)
        Debug.Print "ParseLegacyFormat: creating regexDateOnly"
        Dim regexDateOnly As Object
        Set regexDateOnly = CreateObject("VBScript.RegExp")
        With regexDateOnly
            .Pattern = "(\d{1,2}\.\d{1,2}\.\d{4})\s*$"
            .IgnoreCase = True
            .Global = False
        End With
        
        Debug.Print "ParseLegacyFormat: executing regexDateOnly"
        Dim dateMatches As Object
        Set dateMatches = regexDateOnly.Execute(segment)
        
        Debug.Print "ParseLegacyFormat: dateMatches.Count=" & dateMatches.Count
        If dateMatches.Count > 0 Then
            Debug.Print "ParseLegacyFormat: accessing SubMatches(0) for date"
            changeDate = dateMatches(0).SubMatches(0)
            Debug.Print "ParseLegacyFormat: changeDate='" & changeDate & "'"
            ' Remove date from the segment
            segment = RTrim(Replace(segment, dateMatches(0).Value, ""))
        Else
            ' No date found - skip this segment
            Set ParseLegacyFormat = Nothing
            Exit Function
        End If
    End If
    
    ' RegExp to capture Mnt.X: War(Y); Ist(Z).
    ' Note: VBScript.RegExp does NOT support (?:...) so we use capturing groups
    Debug.Print "ParseLegacyFormat: creating regexMnt"
    Dim regexMnt As Object
    Set regexMnt = CreateObject("VBScript.RegExp")
    With regexMnt
        ' Pattern: Mnt.N: War(value); Ist(value).
        .Pattern = "Mnt\.(\d+):\s*War\(([^)]*)\);\s*Ist\(([^)]*)\)\."
        .IgnoreCase = True
        .Global = True
    End With
    
    Debug.Print "ParseLegacyFormat: executing regexMnt"
    Dim mntMatches As Object
    Set mntMatches = regexMnt.Execute(segment)
    
    Debug.Print "ParseLegacyFormat: mntMatches.Count=" & mntMatches.Count
    Dim matchItem As Object
    Dim monthNum As String
    Dim warVal As String
    Dim istVal As String
    Dim mntIndex As Long
    mntIndex = 0
    
    For Each matchItem In mntMatches
        Debug.Print "ParseLegacyFormat: processing mntMatch " & mntIndex
        Debug.Print "ParseLegacyFormat: accessing SubMatches(0)"
        monthNum = matchItem.SubMatches(0)
        Debug.Print "ParseLegacyFormat: accessing SubMatches(1)"
        warVal = matchItem.SubMatches(1)
        Debug.Print "ParseLegacyFormat: accessing SubMatches(2)"
        istVal = matchItem.SubMatches(2)
        
        Debug.Print "ParseLegacyFormat: month=" & monthNum & ", war='" & warVal & "', ist='" & istVal & "'"
        
        Dim singleMonthChanges As Object
        Set singleMonthChanges = CreateObject("Scripting.Dictionary")
        singleMonthChanges.Add "War", warVal
        singleMonthChanges.Add "Ist", istVal
        
        changesDict.Add monthNum, singleMonthChanges
        mntIndex = mntIndex + 1
    Next matchItem
    
    ' Create dictionary for the event
    evt.Add "IsRuck", isRuck
    evt.Add "Reason", Reason
    evt.Add "ChangeDate", changeDate
    evt.Add "Changes", changesDict
    
    Debug.Print "ParseLegacyFormat: SUCCESS, returning evt"
    Set ParseLegacyFormat = evt
    Exit Function
    
ErrHandler:
    Debug.Print "ERROR in ParseLegacyFormat: Err=" & Err.Number & " - " & Err.Description & " | segment=" & segment
    Err.Raise Err.Number, "valid_ParseHistory.ParseLegacyFormat", Err.Description
End Function
