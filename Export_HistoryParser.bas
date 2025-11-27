Attribute VB_Name = "Export_HistoryParser"
'==========================
'   Module: Export_HistoryParser
'   Purpose: Universal history string parser supporting both legacy and new formats
'   Used by Admin file for parsing history strings from AZ column
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

' ========================================
' Public Type for Parsed Event
' ========================================

' Event structure returned by parser
' Changes is a Dictionary: key = field identifier (month number or field name), value = Dictionary with "War"/"Ist" keys
Public Type HistoryEvent
    IsRuck As Boolean
    Reason As String
    ChangeDate As String
    Changes As Object       ' Scripting.Dictionary
    IsDecline As Boolean
    DeclineNumber As Long
    DeclineComment As String
End Type

' ========================================
' Main Parsing Function
' ========================================

' Parse history string into collection of events
' Returns: Collection of Dictionary objects with keys: IsRuck, Reason, ChangeDate, Changes, IsDecline, DeclineNumber, DeclineComment
Public Function ParseHistoryUnified(ByVal historyString As String) As Collection
    Dim result As Collection
    Set result = New Collection
    
    If Len(Trim(historyString)) = 0 Then
        Set ParseHistoryUnified = result
        Exit Function
    End If
    
    ' Split the input string by "||" session separator
    Dim eventsArray() As String
    eventsArray = Split(historyString, "||")
    
    Dim i As Long
    Dim segment As String
    
    For i = LBound(eventsArray) To UBound(eventsArray)
        segment = Trim(eventsArray(i))
        
        ' Skip empty segments
        If Len(segment) = 0 Then GoTo SkipSegment
        
        ' Detect format and parse accordingly
        Dim evt As Object
        If IsNewFormat(segment) Then
            Set evt = ParseNewFormatSegment(segment)
        Else
            Set evt = ParseLegacyFormatSegment(segment)
        End If
        
        ' Add event if valid
        If Not evt Is Nothing Then
            result.Add evt
        End If
        
SkipSegment:
    Next i
    
    Set ParseHistoryUnified = result
End Function

' ========================================
' Format Detection
' ========================================

Private Function IsNewFormat(ByVal segment As String) As Boolean
    ' New format indicators:
    ' - Contains -> (value separator)
    ' - Contains /@ or @/ (comment delimiters)
    ' - Starts with DCL( or RUCK: followed by tag
    ' - Contains field tags like M01(, ADR(, SB1( etc.
    
    ' Check for new format markers
    If InStr(segment, "->") > 0 Then
        IsNewFormat = True
        Exit Function
    End If
    
    If InStr(segment, "/@") > 0 Or InStr(segment, "@/") > 0 Then
        IsNewFormat = True
        Exit Function
    End If
    
    ' Check for DCL( at start
    If Left(segment, 4) = "DCL(" Then
        IsNewFormat = True
        Exit Function
    End If
    
    ' Check for RUCK: prefix with new format after it
    If Left(segment, 5) = "RUCK:" Then
        Dim afterRuck As String
        afterRuck = Trim(Mid(segment, 6))
        ' Check if followed by new-style tag
        If InStr(afterRuck, "->") > 0 Then
            IsNewFormat = True
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
        IsNewFormat = True
        Exit Function
    End If
    
    IsNewFormat = False
End Function

' ========================================
' New Format Parser
' ========================================

Private Function ParseNewFormatSegment(ByVal segment As String) As Object
    ' Parse new format: [RUCK:]<TAG>(<OLD>-><NEW>);.../@<COMMENT>@/<DATE>
    '                   DCL(<N>-><comment>)
    
    Dim evt As Object
    Set evt = CreateObject("Scripting.Dictionary")
    
    Dim changesDict As Object
    Set changesDict = CreateObject("Scripting.Dictionary")
    
    Dim isRuck As Boolean
    isRuck = False
    
    Dim isDecline As Boolean
    isDecline = False
    
    Dim declNumber As Long
    declNumber = 0
    
    Dim declComment As String
    declComment = ""
    
    Dim Reason As String
    Reason = ""
    
    Dim changeDate As String
    changeDate = ""
    
    ' Check for RUCK: prefix
    If Left(segment, 5) = "RUCK:" Then
        isRuck = True
        segment = Trim(Mid(segment, 6))
    End If
    
    ' Check for DCL( decline entry
    If Left(segment, 4) = "DCL(" Then
        isDecline = True
        
        ' Parse DCL(N->comment)
        Dim dclRegex As Object
        Set dclRegex = CreateObject("VBScript.RegExp")
        With dclRegex
            .Pattern = "DCL\((\d+)->([^)]+)\)"
            .IgnoreCase = True
            .Global = False
        End With
        
        Dim dclMatches As Object
        Set dclMatches = dclRegex.Execute(segment)
        
        If dclMatches.Count > 0 Then
            declNumber = CLng(dclMatches(0).SubMatches(0))
            declComment = dclMatches(0).SubMatches(1)
        End If
        
        evt.Add "IsRuck", False
        evt.Add "Reason", ""
        evt.Add "ChangeDate", ""
        evt.Add "Changes", changesDict
        evt.Add "IsDecline", True
        evt.Add "DeclineNumber", declNumber
        evt.Add "DeclineComment", declComment
        
        Set ParseNewFormatSegment = evt
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
        Reason = cdMatches(0).SubMatches(0)
        changeDate = cdMatches(0).SubMatches(1)
        ' Remove comment+date from segment
        segment = RTrim(Replace(segment, cdMatches(0).Value, ""))
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
        unifiedKey = ConvertTagToUnifiedKey(tag)
        
        ' Create change dictionary
        Dim singleChange As Object
        Set singleChange = CreateObject("Scripting.Dictionary")
        singleChange.Add "War", oldVal
        singleChange.Add "Ist", newVal
        
        changesDict.Add unifiedKey, singleChange
    Next fieldMatch
    
    ' Build result event
    evt.Add "IsRuck", isRuck
    evt.Add "Reason", Reason
    evt.Add "ChangeDate", changeDate
    evt.Add "Changes", changesDict
    evt.Add "IsDecline", False
    evt.Add "DeclineNumber", 0
    evt.Add "DeclineComment", ""
    
    Set ParseNewFormatSegment = evt
End Function

' Convert new format tag to unified key used in Changes dictionary
Private Function ConvertTagToUnifiedKey(ByVal tag As String) As String
    ' Month tags M01-M12 -> 1-12 (numeric string for compatibility)
    If Left(UCase(tag), 1) = "M" And Len(tag) = 3 Then
        Dim monthNum As Integer
        On Error Resume Next
        monthNum = CInt(Mid(tag, 2, 2))
        On Error GoTo 0
        
        If monthNum >= 1 And monthNum <= 12 Then
            ConvertTagToUnifiedKey = CStr(monthNum)
            Exit Function
        End If
    End If
    
    ' Field tags -> descriptive names (for Address, Subject1, Subject2)
    Select Case UCase(tag)
        Case "ADR"
            ConvertTagToUnifiedKey = "Address"
        Case "SB1"
            ConvertTagToUnifiedKey = "Subject1"
        Case "SB2"
            ConvertTagToUnifiedKey = "Subject2"
        Case Else
            ' Return as-is for other fields (FID, PAR, CHD, DOB, TEL, MOB, EML, PR1, PR2, EX1-3)
            ConvertTagToUnifiedKey = UCase(tag)
    End Select
End Function

' ========================================
' Legacy Format Parser
' ========================================

Private Function ParseLegacyFormatSegment(ByVal segment As String) As Object
    ' Parse legacy format: Mnt.N: War(X); Ist(Y). /Comment/ DD.MM.YYYY
    '                      Address: Was(X); Is(Y). /Comment/ DD.MM.YYYY
    '                      Decl_N: Was(); Is(comment).
    
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
    
    ' Check for Decl_N: legacy decline format
    If Left(segment, 5) = "Decl_" Then
        ' Parse Decl_N: Was(); Is(comment).
        Dim declRegex As Object
        Set declRegex = CreateObject("VBScript.RegExp")
        With declRegex
            .Pattern = "Decl_(\d+):\s*Was\([^)]*\);\s*Is\(([^)]*)\)\."
            .IgnoreCase = True
            .Global = False
        End With
        
        Dim declMatches As Object
        Set declMatches = declRegex.Execute(segment)
        
        If declMatches.Count > 0 Then
            evt.Add "IsRuck", False
            evt.Add "Reason", ""
            evt.Add "ChangeDate", ""
            evt.Add "Changes", changesDict
            evt.Add "IsDecline", True
            evt.Add "DeclineNumber", CLng(declMatches(0).SubMatches(0))
            evt.Add "DeclineComment", declMatches(0).SubMatches(1)
            
            Set ParseLegacyFormatSegment = evt
            Exit Function
        End If
    End If
    
    ' Extract /Comment/ and date using two-step approach
    ' VBScript.RegExp does NOT support (?:...) non-capturing groups
    Dim Reason As String
    Reason = ""
    Dim changeDate As String
    changeDate = ""
    
    ' Step 1: Try to match /reason/ date pattern first
    Dim regexWithReason As Object
    Set regexWithReason = CreateObject("VBScript.RegExp")
    With regexWithReason
        .Pattern = "/([^/]*)/\s*(\d{1,2}\.\d{1,2}\.\d{4})\s*$"
        .IgnoreCase = True
        .Global = False
    End With
    
    Dim reasonMatches As Object
    Set reasonMatches = regexWithReason.Execute(segment)
    
    If reasonMatches.Count > 0 Then
        Reason = reasonMatches(0).SubMatches(0)
        changeDate = reasonMatches(0).SubMatches(1)
        segment = RTrim(Replace(segment, reasonMatches(0).Value, ""))
    Else
        ' Step 2: Try to match just date (no reason)
        Dim regexDateOnly As Object
        Set regexDateOnly = CreateObject("VBScript.RegExp")
        With regexDateOnly
            .Pattern = "(\d{1,2}\.\d{1,2}\.\d{4})\s*$"
            .IgnoreCase = True
            .Global = False
        End With
        
        Dim dateMatches As Object
        Set dateMatches = regexDateOnly.Execute(segment)
        
        If dateMatches.Count > 0 Then
            changeDate = dateMatches(0).SubMatches(0)
            segment = RTrim(Replace(segment, dateMatches(0).Value, ""))
        Else
            ' No date found - skip this segment
            Set ParseLegacyFormatSegment = Nothing
            Exit Function
        End If
    End If
    
    ' Parse Mnt.N: War(X); Ist(Y).
    ' Note: VBScript.RegExp does NOT support (?:...) so we use capturing groups
    Dim regexMnt As Object
    Set regexMnt = CreateObject("VBScript.RegExp")
    With regexMnt
        ' Pattern: Mnt.N: War(value); Ist(value). where value can be -?digit(s) optionally with ,digit(s)
        .Pattern = "Mnt\.(\d+):\s*War\((-?\d+(,\d+)?)?\);\s*Ist\((-?\d+(,\d+)?)?\)\."
        .IgnoreCase = True
        .Global = True
    End With
    
    Dim mntMatches As Object
    Set mntMatches = regexMnt.Execute(segment)
    
    Dim matchItem As Object
    For Each matchItem In mntMatches
        Dim monthNum As String
        Dim warVal As String
        Dim istVal As String
        
        monthNum = matchItem.SubMatches(0)
        ' SubMatches(1) is the full War value (including optional decimal part)
        ' SubMatches(2) is just the decimal part (,\d+) - we ignore it
        ' SubMatches(3) is the full Ist value
        ' SubMatches(4) is just the decimal part for Ist - we ignore it
        warVal = matchItem.SubMatches(1)
        istVal = matchItem.SubMatches(3)
        
        Dim singleMonthChange As Object
        Set singleMonthChange = CreateObject("Scripting.Dictionary")
        singleMonthChange.Add "War", warVal
        singleMonthChange.Add "Ist", istVal
        
        changesDict.Add monthNum, singleMonthChange
    Next matchItem
    
    ' Parse Address: Was(X); Is(Y).
    Call AddLegacyFieldChange(segment, "Address", changesDict)
    
    ' Parse Subject1: Was(X); Is(Y).
    Call AddLegacyFieldChange(segment, "Subject1", changesDict)
    
    ' Parse Subject2: Was(X); Is(Y).
    Call AddLegacyFieldChange(segment, "Subject2", changesDict)
    
    ' Build result event
    evt.Add "IsRuck", isRuck
    evt.Add "Reason", Reason
    evt.Add "ChangeDate", changeDate
    evt.Add "Changes", changesDict
    evt.Add "IsDecline", False
    evt.Add "DeclineNumber", 0
    evt.Add "DeclineComment", ""
    
    Set ParseLegacyFormatSegment = evt
End Function

' Helper to extract legacy field changes like Address: Was(X); Is(Y).
Private Sub AddLegacyFieldChange(ByVal segment As String, ByVal fieldName As String, ByVal dict As Object)
    Dim marker As String
    marker = fieldName & ": Was("
    
    Dim pos As Long
    pos = InStr(1, segment, marker, vbTextCompare)
    If pos = 0 Then Exit Sub
    
    Dim startWar As Long
    startWar = pos + Len(marker)
    
    Dim endWar As Long
    endWar = InStr(startWar, segment, ")", vbTextCompare)
    If endWar = 0 Then Exit Sub
    
    Dim warVal As String
    warVal = Mid(segment, startWar, endWar - startWar)
    
    Dim markerIs As String
    markerIs = "); Is("
    
    Dim posIs As Long
    posIs = InStr(endWar, segment, markerIs, vbTextCompare)
    If posIs = 0 Then Exit Sub
    
    Dim startIst As Long
    startIst = posIs + Len(markerIs)
    
    Dim endIst As Long
    endIst = InStr(startIst, segment, ")", vbTextCompare)
    
    Dim istVal As String
    If endIst > startIst Then
        istVal = Mid(segment, startIst, endIst - startIst)
    Else
        istVal = Mid(segment, startIst)
    End If
    
    warVal = Trim(warVal)
    istVal = Trim(istVal)
    
    Dim fieldDict As Object
    Set fieldDict = CreateObject("Scripting.Dictionary")
    fieldDict.Add "War", warVal
    fieldDict.Add "Ist", istVal
    
    If Not dict.Exists(fieldName) Then
        dict.Add fieldName, fieldDict
    End If
End Sub

' ========================================
' Utility Functions
' ========================================

' Get the last decline entry from history
' Returns Dictionary with keys: Number, Comment, or Nothing if no decline found
Public Function GetLastDeclineEntry(ByVal historyString As String) As Object
    Dim events As Collection
    Set events = ParseHistoryUnified(historyString)
    
    Dim maxN As Long
    maxN = -1
    Dim lastDecline As Object
    Set lastDecline = Nothing
    
    Dim evt As Object
    Dim i As Long
    For i = 1 To events.Count
        Set evt = events(i)
        
        If evt("IsDecline") Then
            If CLng(evt("DeclineNumber")) > maxN Then
                maxN = CLng(evt("DeclineNumber"))
                
                Set lastDecline = CreateObject("Scripting.Dictionary")
                lastDecline.Add "Number", maxN
                lastDecline.Add "Comment", evt("DeclineComment")
            End If
        End If
    Next i
    
    Set GetLastDeclineEntry = lastDecline
End Function

' Count total decline entries in history
Public Function CountDeclineEntries(ByVal historyString As String) As Long
    Dim events As Collection
    Set events = ParseHistoryUnified(historyString)
    
    Dim count As Long
    count = 0
    
    Dim evt As Object
    Dim i As Long
    For i = 1 To events.Count
        Set evt = events(i)
        
        If evt("IsDecline") Then
            count = count + 1
        End If
    Next i
    
    CountDeclineEntries = count
End Function

' Get events filtered by date range
Public Function GetEventsInDateRange(ByVal historyString As String, _
                                     ByVal startDate As Date, _
                                     ByVal endDate As Date) As Collection
    Dim allEvents As Collection
    Set allEvents = ParseHistoryUnified(historyString)
    
    Dim result As Collection
    Set result = New Collection
    
    Dim evt As Object
    Dim i As Long
    For i = 1 To allEvents.Count
        Set evt = allEvents(i)
        
        ' Skip decline entries (they don't have dates)
        If evt("IsDecline") Then GoTo NextEvent
        
        Dim evtDate As Date
        On Error Resume Next
        evtDate = CDate(evt("ChangeDate"))
        If Err.Number <> 0 Then
            Err.Clear
            GoTo NextEvent
        End If
        On Error GoTo 0
        
        If evtDate >= startDate And evtDate <= endDate Then
            result.Add evt
        End If
        
NextEvent:
    Next i
    
    Set GetEventsInDateRange = result
End Function

