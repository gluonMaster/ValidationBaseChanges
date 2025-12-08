Attribute VB_Name = "Export_ValidationKartei"
'=====================================
'   Module: Export_ValidationKartei
'   Purpose: Validate Kartei sheet data before synchronization
'   - Detect duplicate FamilyIDs with different Parent names
'   - Detect empty FamilyID with non-empty Parent
'=====================================
Option Explicit

' Global flag to signal validation failure during OnSave scenario
' Used by Workbook_BeforeClose to cancel closing if validation fails
Public g_KarteiValidationFailed As Boolean

' Color constant for highlighting problematic rows
Private Const HIGHLIGHT_COLOR As Long = 65535  ' Yellow (RGB 255, 255, 0)

'=====================================
'   Main validation function
'=====================================
Public Function ValidateKarteiBeforeSync() As Boolean
    ' Main entry point for Kartei validation before sync
    ' Returns True if no problems found, False otherwise
    ' Side effects: highlights problematic rows, shows MsgBox to user
    
    Dim wsKartei As Worksheet
    Set wsKartei = ThisWorkbook.Worksheets("Kartei")
    
    Dim lastRow As Long
    lastRow = wsKartei.Cells(wsKartei.Rows.Count, 1).End(xlUp).Row
    
    ' If no data rows, validation passes
    If lastRow < 3 Then
        ValidateKarteiBeforeSync = True
        Exit Function
    End If
    
    ' Run both validations
    Dim duplicateProblems As String
    Dim emptyFamilyIDProblems As String
    Dim duplicateRows As Collection
    Dim emptyFamilyIDRows As Collection
    
    Set duplicateRows = New Collection
    Set emptyFamilyIDRows = New Collection
    
    ' Check for duplicate FamilyIDs with different Parents
    duplicateProblems = CheckDuplicateFamilyIDs(wsKartei, lastRow, duplicateRows)
    
    ' Check for empty FamilyID with non-empty Parent
    emptyFamilyIDProblems = CheckEmptyFamilyIDWithParent(wsKartei, lastRow, emptyFamilyIDRows)
    
    ' Highlight problematic rows
    If duplicateRows.Count > 0 Then
        HighlightProblematicRows wsKartei, duplicateRows
    End If
    
    If emptyFamilyIDRows.Count > 0 Then
        HighlightProblematicRows wsKartei, emptyFamilyIDRows
    End If
    
    ' Show messages and determine result
    Dim hasProblems As Boolean
    hasProblems = False
    
    If duplicateProblems <> "" Then
        MsgBox duplicateProblems, vbExclamation, "Validierungsfehler - Doppelte FamilyID"
        hasProblems = True
    End If
    
    If emptyFamilyIDProblems <> "" Then
        MsgBox emptyFamilyIDProblems, vbExclamation, "Validierungsfehler - Leere FamilyID"
        hasProblems = True
    End If
    
    ValidateKarteiBeforeSync = Not hasProblems
End Function

'=====================================
'   Check for duplicate FamilyIDs with different Parent names
'=====================================
Private Function CheckDuplicateFamilyIDs(ByVal ws As Worksheet, _
                                         ByVal lastRow As Long, _
                                         ByRef problemRows As Collection) As String
    ' Checks for FamilyIDs that appear with different Parent names
    ' Returns error message string (empty if no problems)
    ' Populates problemRows collection with row numbers to highlight
    
    ' Structure using nested dictionaries (to avoid VBA Collection assignment issues):
    ' dictFamilyID: Key = FamilyID (uppercase)
    ' dictRows: Key = FamilyID (uppercase), Value = semicolon-separated row numbers
    ' dictParents: Key = FamilyID (uppercase), Value = Dictionary of Parent names
    
    Dim dictRows As Object      ' FamilyID -> "row1;row2;row3"
    Dim dictParents As Object   ' FamilyID -> Dictionary(ParentKey -> ParentOriginal)
    
    Set dictRows = CreateObject("Scripting.Dictionary")
    Set dictParents = CreateObject("Scripting.Dictionary")
    
    Dim i As Long
    Dim familyID As String
    Dim parentName As String
    Dim familyIDKey As String
    Dim parentKey As String
    
    ' First pass: collect all FamilyID -> Parent mappings
    For i = 3 To lastRow
        familyID = Trim(CStr(ws.Cells(i, 1).Value))   ' Column A - FamilyID
        parentName = Trim(CStr(ws.Cells(i, 2).Value)) ' Column B - Parent
        
        ' Skip rows where both FamilyID and Parent are empty
        If familyID = "" And parentName = "" Then
            GoTo NextRow
        End If
        
        ' Skip rows with empty FamilyID (handled by other validation)
        If familyID = "" Then
            GoTo NextRow
        End If
        
        familyIDKey = UCase(familyID)
        parentKey = UCase(parentName)
        
        If Not dictRows.Exists(familyIDKey) Then
            ' First occurrence of this FamilyID
            dictRows(familyIDKey) = CStr(i)
            
            ' Create parents dictionary for this FamilyID
            Dim newParentDict As Object
            Set newParentDict = CreateObject("Scripting.Dictionary")
            If parentName <> "" Then
                newParentDict(parentKey) = parentName ' Store original case
            End If
            Set dictParents(familyIDKey) = newParentDict
        Else
            ' FamilyID already seen - append row number
            dictRows(familyIDKey) = dictRows(familyIDKey) & ";" & CStr(i)
            
            ' Add parent to distinct parents dict if non-empty
            If parentName <> "" Then
                Dim existingParentDict As Object
                Set existingParentDict = dictParents(familyIDKey)
                If Not existingParentDict.Exists(parentKey) Then
                    existingParentDict(parentKey) = parentName
                End If
            End If
        End If
NextRow:
    Next i
    
    ' Second pass: find FamilyIDs with multiple different Parents
    Dim conflictFamilyIDs As Collection
    Set conflictFamilyIDs = New Collection
    
    Dim key As Variant
    For Each key In dictRows.Keys
        Dim parentDict As Object
        Set parentDict = dictParents(key)
        
        ' Conflict exists if more than one distinct non-empty Parent name
        If parentDict.Count > 1 Then
            conflictFamilyIDs.Add key
            
            ' Add all rows for this FamilyID to problem rows
            Dim rowsStr As String
            rowsStr = dictRows(key)
            Dim rowArr() As String
            rowArr = Split(rowsStr, ";")
            Dim idx As Long
            For idx = LBound(rowArr) To UBound(rowArr)
                On Error Resume Next  ' Avoid duplicates in collection
                problemRows.Add CLng(rowArr(idx)), rowArr(idx)
                On Error GoTo 0
            Next idx
        End If
    Next key
    
    ' Build error message if conflicts found
    If conflictFamilyIDs.Count = 0 Then
        CheckDuplicateFamilyIDs = ""
        Exit Function
    End If
    
    Dim msg As String
    msg = "Es wurden doppelte FamilyID-Eintraege mit unterschiedlichen Elternnamen gefunden:" & vbCrLf & vbCrLf
    
    Dim conflictKey As Variant
    For Each conflictKey In conflictFamilyIDs
        Dim conflictRowsStr As String
        conflictRowsStr = dictRows(conflictKey)
        
        Dim conflictParentDict As Object
        Set conflictParentDict = dictParents(conflictKey)
        
        ' Get original FamilyID from first row
        Dim firstRowNum As Long
        Dim conflictRowArr() As String
        conflictRowArr = Split(conflictRowsStr, ";")
        firstRowNum = CLng(conflictRowArr(0))
        
        Dim originalFamilyID As String
        originalFamilyID = Trim(CStr(ws.Cells(firstRowNum, 1).Value))
        
        msg = msg & "FamilyID '" & originalFamilyID & "':" & vbCrLf
        msg = msg & "  Zeilen: " & Replace(conflictRowsStr, ";", ", ") & vbCrLf
        msg = msg & "  Elternnamen: " & DictValuesToString(conflictParentDict) & vbCrLf & vbCrLf
    Next conflictKey
    
    msg = msg & "Synchronisation abgebrochen." & vbCrLf
    msg = msg & "Bitte korrigieren Sie die Daten im Blatt Kartei."
    
    CheckDuplicateFamilyIDs = msg
End Function

'=====================================
'   Check for empty FamilyID with non-empty Parent
'=====================================
Private Function CheckEmptyFamilyIDWithParent(ByVal ws As Worksheet, _
                                               ByVal lastRow As Long, _
                                               ByRef problemRows As Collection) As String
    ' Checks for rows where FamilyID is empty but Parent is not
    ' Returns error message string (empty if no problems)
    ' Populates problemRows collection with row numbers to highlight
    
    Dim i As Long
    Dim familyID As String
    Dim parentName As String
    
    For i = 3 To lastRow
        familyID = Trim(CStr(ws.Cells(i, 1).Value))   ' Column A - FamilyID
        parentName = Trim(CStr(ws.Cells(i, 2).Value)) ' Column B - Parent
        
        ' Problem: FamilyID empty but Parent not empty
        If familyID = "" And parentName <> "" Then
            On Error Resume Next  ' Avoid duplicates
            problemRows.Add i, CStr(i)
            On Error GoTo 0
        End If
    Next i
    
    ' Build error message if problems found
    If problemRows.Count = 0 Then
        CheckEmptyFamilyIDWithParent = ""
        Exit Function
    End If
    
    Dim msg As String
    msg = "Es wurden Eintraege mit leerer FamilyID und ausgefuelltem Elternnamen gefunden:" & vbCrLf & vbCrLf
    msg = msg & "Zeilen: " & CollectionToString(problemRows) & vbCrLf & vbCrLf
    msg = msg & "Synchronisation abgebrochen." & vbCrLf
    msg = msg & "Bitte fuellen Sie die FamilyID aus oder entfernen/korrigieren Sie diese Zeilen im Blatt Kartei."
    
    CheckEmptyFamilyIDWithParent = msg
End Function

'=====================================
'   Helper: Highlight problematic rows
'=====================================
Private Sub HighlightProblematicRows(ByVal ws As Worksheet, ByVal rows As Collection)
    ' Highlights column A cells for the specified rows with yellow color
    
    Dim r As Variant
    For Each r In rows
        ws.Cells(CLng(r), 1).Interior.Color = HIGHLIGHT_COLOR
    Next r
End Sub

'=====================================
'   Helper: Convert Collection to comma-separated string
'=====================================
Private Function CollectionToString(ByVal coll As Collection) As String
    Dim result As String
    Dim item As Variant
    Dim first As Boolean
    
    first = True
    For Each item In coll
        If first Then
            result = CStr(item)
            first = False
        Else
            result = result & ", " & CStr(item)
        End If
    Next item
    
    CollectionToString = result
End Function

'=====================================
'   Helper: Convert Dictionary values to comma-separated string
'=====================================
Private Function DictValuesToString(ByVal dict As Object) As String
    Dim result As String
    Dim key As Variant
    Dim first As Boolean
    
    first = True
    For Each key In dict.Keys
        If first Then
            result = "'" & dict(key) & "'"
            first = False
        Else
            result = result & ", '" & dict(key) & "'"
        End If
    Next key
    
    DictValuesToString = result
End Function

'=====================================
'   Utility: Clear validation highlights
'=====================================
Public Sub ClearValidationHighlights()
    ' Clears yellow highlighting from column A that was added by validation
    ' Can be called manually if user wants to clear highlights after fixing issues
    
    Dim ws As Worksheet
    Set ws = ThisWorkbook.Worksheets("Kartei")
    
    Dim lastRow As Long
    lastRow = ws.Cells(ws.Rows.Count, 1).End(xlUp).Row
    
    If lastRow < 3 Then Exit Sub
    
    Dim i As Long
    For i = 3 To lastRow
        If ws.Cells(i, 1).Interior.Color = HIGHLIGHT_COLOR Then
            ws.Cells(i, 1).Interior.ColorIndex = xlNone
        End If
    Next i
End Sub
