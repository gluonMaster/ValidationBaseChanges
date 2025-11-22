Attribute VB_Name = "valid_ParseHistory"
'==========================
'   Module: valid_ParseHistory
'   Purpose: Parse history strings from AZ column (tested implementation from alt/Parceing.bas)
'   Uses exact working logic - DO NOT MODIFY REGEX PATTERNS
'==========================

Option Explicit

' Parse history string into collection of events
' Returns: Collection of Dictionary objects with keys: IsRuck, Reason, ChangeDate, Changes
Public Function ParseHistory(ByVal historyString As String) As Collection
    Dim result As Collection
    Set result = New Collection
    
    Dim eventsArray() As String
    Dim i As Long
    
    ' Split the input string by "||" (exact match from alt/Parceing.bas)
    eventsArray = Split(historyString, "||")
    
    ' RegExp to capture reason (if present) and date
    Dim regexReasonDate As Object
    Set regexReasonDate = CreateObject("VBScript.Regexp")
    With regexReasonDate
        .Pattern = "(?:/(.*?)/\s*)?(\d{1,2}\.\d{1,2}\.\d{4})\s*$"
        .IgnoreCase = True
        .Global = False
    End With
    
    ' RegExp to capture Mnt.X: War(Y); Ist(Z).
    Dim regexMnt As Object
    Set regexMnt = CreateObject("VBScript.Regexp")
    With regexMnt
        .Pattern = "Mnt\.(\d+):\s*War\((\-?\d+(?:,\d+)?)?\);\s*Ist\((\-?\d+(?:,\d+)?)?\)\."
        .IgnoreCase = True
        .Global = True
    End With
    
    Dim segment As String
    Dim isRuck As Boolean
    
    Dim reasonDateMatches As Object
    Dim mntMatches As Object
    Dim matchItem As Object
    
    For i = LBound(eventsArray) To UBound(eventsArray)
        
        segment = Trim(eventsArray(i))
        
        ' Skip empty segments
        If Len(segment) = 0 Then GoTo SkipSegment
        
        ' Check if segment starts with "Ruck:"
        If Left(segment, 5) = "Ruck:" Then
            isRuck = True
            segment = Trim(Mid(segment, 6))
        Else
            isRuck = False
        End If
        
        Dim Reason As String
        Reason = ""
        
        Dim changeDate As String
        changeDate = ""
        
        Set reasonDateMatches = regexReasonDate.Execute(segment)
        
        If reasonDateMatches.Count > 0 Then
            Reason = reasonDateMatches(0).submatches(0)
            changeDate = reasonDateMatches(0).submatches(1)
            
            ' Remove reason+date from the segment
            segment = RTrim(Replace(segment, reasonDateMatches(0).Value, ""))
        Else
            GoTo SkipSegment
        End If
        
        ' Parse Mnt.X: War(Y); Ist(Z).
        Set mntMatches = regexMnt.Execute(segment)
        
        Dim changesDict As Object
        Set changesDict = CreateObject("Scripting.Dictionary")
        
        Dim monthNum As String
        Dim warVal As String
        Dim istVal As String
        
        For Each matchItem In mntMatches
            monthNum = matchItem.submatches(0)
            warVal = matchItem.submatches(1)
            istVal = matchItem.submatches(2)
            
            Dim singleMonthChanges As Object
            Set singleMonthChanges = CreateObject("Scripting.Dictionary")
            singleMonthChanges.Add "War", warVal
            singleMonthChanges.Add "Ist", istVal
            
            changesDict.Add monthNum, singleMonthChanges
        Next matchItem
        
        ' Create dictionary for the event
        Dim singleEvent As Object
        Set singleEvent = CreateObject("Scripting.Dictionary")
        
        singleEvent.Add "IsRuck", isRuck
        singleEvent.Add "Reason", Reason
        singleEvent.Add "ChangeDate", changeDate
        singleEvent.Add "Changes", changesDict
        
        result.Add singleEvent
        
SkipSegment:
    Next i
    
    Set ParseHistory = result
End Function
