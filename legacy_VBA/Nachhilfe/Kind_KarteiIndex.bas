Attribute VB_Name = "Kind_KarteiIndex"
Option Explicit

' =============================================================================
' Module: Kind_KarteiIndex
' Purpose: Builds and queries an index of Kartei records for fast lookup
' Prefix: Kind_ (all modules/classes in this project use this prefix)
' Index Structure:
'   Key: "FamilyIDNorm|SubjectNorm" (SubjectNorm = base subject without prefix)
'   Value: Collection of Variant Arrays [karteiRow, childRaw, isKN]
' =============================================================================

' -----------------------------------------------------------------------------
' Kartei layout constants (from Kind_Config)
' First data row in Kartei
' -----------------------------------------------------------------------------
Private Const KARTEI_FIRST_DATA_ROW As Long = 2


' -----------------------------------------------------------------------------
' Kind_BuildKarteiIndex
' Purpose: Reads Kartei sheet from Admin workbook and builds lookup index
' Parameters:
'   wbAdmin - Admin workbook containing Kartei sheet
' Returns: Scripting.Dictionary object (late bound), or Nothing on failure
' Index Key: "FamilyIDNorm|SubjectNorm"
' Index Value: Collection of Variant Array(0 To 2):
'   [0] = karteiRow (Long)
'   [1] = childRaw (String)
'   [2] = isKN (Boolean)
' -----------------------------------------------------------------------------
Public Function Kind_BuildKarteiIndex(ByVal wbAdmin As Workbook) As Object
    Dim dict As Object
    Dim wsKartei As Worksheet
    Dim lastRow As Long
    Dim r As Long
    Dim subjectColLetter As String
    
    Dim familyIDRaw As String
    Dim familyIDNorm As String
    Dim childRaw As String
    Dim statusRaw As String
    Dim subjectRaw As String
    Dim subjectNorm As String
    Dim isKN As Boolean
    Dim dictKey As String
    
    Dim entry As Variant
    Dim coll As Collection
    
    On Error GoTo ErrHandler
    
    Set Kind_BuildKarteiIndex = Nothing
    
    ' Validate input
    If wbAdmin Is Nothing Then
        Exit Function
    End If
    
    ' Get Kartei worksheet
    If Not Kind_TryGetWorksheet(wbAdmin, KIND_ADMIN_SHEET_KARTEI, wsKartei) Then
        Exit Function
    End If
    
    ' Create dictionary (late binding)
    Set dict = CreateObject("Scripting.Dictionary")
    dict.CompareMode = 1  ' vbTextCompare - case insensitive keys
    
    ' Determine subject column based on current month
    subjectColLetter = Kind_GetKarteiSubjectColumnLetter()
    
    ' Find last row in FamilyID column
    lastRow = wsKartei.Cells(wsKartei.Rows.Count, KIND_KARTEI_COL_FAMILYID).End(xlUp).Row
    
    ' Iterate through Kartei rows
    For r = KARTEI_FIRST_DATA_ROW To lastRow
        ' Read FamilyID
        familyIDRaw = CStr(wsKartei.Range(KIND_KARTEI_COL_FAMILYID & r).Value)
        familyIDNorm = Kind_NormalizeForCompare(familyIDRaw)
        
        ' Skip rows without FamilyID
        If Len(familyIDNorm) = 0 Then
            GoTo NextRow
        End If
        
        ' Read subject and try to normalize (extract base subject)
        subjectRaw = CStr(wsKartei.Range(subjectColLetter & r).Value)
        If Not Kind_TryNormalizeSubjectKartei(subjectRaw, subjectNorm) Then
            ' Not a Nachhilfe subject - skip
            GoTo NextRow
        End If
        
        ' Read child name
        childRaw = CStr(wsKartei.Range(KIND_KARTEI_COL_CHILD & r).Value)
        
        ' Read status and check for KN
        statusRaw = CStr(wsKartei.Range(KIND_KARTEI_COL_STATUS & r).Value)
        isKN = (UCase$(Trim$(statusRaw)) = "KN")
        
        ' Build dictionary key
        dictKey = familyIDNorm & "|" & subjectNorm
        
        ' Create a fresh entry array for this row
        ' NOTE: Using Array(...) avoids reusing the same array instance across iterations.
        entry = Array(r, childRaw, isKN)
        
        ' Add to dictionary
        If dict.Exists(dictKey) Then
            ' Key exists - add to existing collection
            Set coll = dict.Item(dictKey)
            coll.Add entry
        Else
            ' New key - create new collection
            Set coll = New Collection
            coll.Add entry
            dict.Add dictKey, coll
        End If
        
NextRow:
    Next r
    
    Set Kind_BuildKarteiIndex = dict
    Exit Function
    
ErrHandler:
    Set Kind_BuildKarteiIndex = Nothing
End Function


' -----------------------------------------------------------------------------
' Kind_FindMatches
' Purpose: Searches index for matching Kartei records by FamilyID + Subject,
'          then filters by child name
' Parameters:
'   index         - Dictionary index built by Kind_BuildKarteiIndex
'   familyID      - FamilyID to search for
'   kinderLast    - Last name from Kinder sheet
'   kinderFirst   - First name from Kinder sheet
'   kinderSubject - Subject from Kinder sheet (raw, will be normalized)
'   outMatchCount - (ByRef) Number of child name matches found
'   outAnyKN      - (ByRef) True if any matching record has KN status
'   outDetails    - (ByRef) Details string for logging
' Returns: True if key exists in index (FamilyID+Subject combination found),
'          False if key not found
' -----------------------------------------------------------------------------
Public Function Kind_FindMatches(ByVal index As Object, _
                                 ByVal familyID As String, _
                                 ByVal kinderLast As String, _
                                 ByVal kinderFirst As String, _
                                 ByVal kinderSubject As String, _
                                 ByRef outMatchCount As Long, _
                                 ByRef outAnyKN As Boolean, _
                                 ByRef outDetails As String) As Boolean
    Dim dictKey As String
    Dim familyIDNorm As String
    Dim subjectNorm As String
    Dim coll As Collection
    Dim entry As Variant
    Dim childRaw As String
    Dim isKN As Boolean
    Dim karteiRow As Long
    Dim matchRows As String
    Dim i As Long
    
    ' Initialize outputs
    outMatchCount = 0
    outAnyKN = False
    outDetails = ""
    Kind_FindMatches = False
    
    ' Validate index
    If index Is Nothing Then
        outDetails = "Index nicht verfuegbar"
        Exit Function
    End If
    
    ' Normalize inputs
    familyIDNorm = Kind_NormalizeForCompare(familyID)
    subjectNorm = Kind_NormalizeSubjectKinder(kinderSubject)
    
    ' Build key
    dictKey = familyIDNorm & "|" & subjectNorm
    
    ' Check if key exists
    If Not index.Exists(dictKey) Then
        ' Key not found - no FamilyID + Subject combination in Kartei
        outDetails = "Keine Kartei-Eintraege fuer FamilyID/Fach Kombination"
        Exit Function
    End If
    
    ' Key exists - return True
    Kind_FindMatches = True
    
    ' Get collection of entries for this key
    Set coll = index.Item(dictKey)
    
    ' Search for child name matches within collection
    matchRows = ""
    
    For i = 1 To coll.Count
        entry = coll.Item(i)
        
        karteiRow = CLng(entry(0))
        childRaw = CStr(entry(1))
        isKN = CBool(entry(2))
        
        ' Check if child name matches
        If Kind_IsSameChild(childRaw, kinderLast, kinderFirst) Then
            outMatchCount = outMatchCount + 1
            
            ' Track if any match has KN status
            If isKN Then
                outAnyKN = True
            End If
            
            ' Build list of matching rows for details
            If Len(matchRows) > 0 Then
                matchRows = matchRows & ", "
            End If
            matchRows = matchRows & CStr(karteiRow)
        End If
    Next i
    
    ' Build details string based on match count
    Select Case outMatchCount
        Case 0
            outDetails = "FamilyID/Fach gefunden, aber Kind-Name passt nicht (" & coll.Count & " Eintraege geprueft)"
        Case 1
            If outAnyKN Then
                outDetails = "Treffer in Zeile " & matchRows & " (Status: KN)"
            Else
                outDetails = "Treffer in Zeile " & matchRows
            End If
        Case Else
            outDetails = "Mehrere Treffer in Zeilen: " & matchRows
            If outAnyKN Then
                outDetails = outDetails & " (mindestens ein KN)"
            End If
    End Select
End Function
