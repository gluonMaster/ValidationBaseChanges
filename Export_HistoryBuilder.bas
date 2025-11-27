Attribute VB_Name = "Export_HistoryBuilder"
'==========================
'   Module: Export_HistoryBuilder
'   Purpose: Extended history string formatting for change tracking
'   Contains: Constants, tag definitions, history building functions
'   
'   History Format: [RUCK:]<TAG>(<OLD>-><NEW>);<TAG>(<OLD>-><NEW>);.../@<COMMENT>@/<DATE>||
'   
'   Field Tags:
'     FID - Family ID (A)          PAR - Parent name (B)
'     CHD - Child name (D)         DOB - Date of birth (E)
'     ADR - Address (F)            TEL - Phone (G)
'     MOB - Mobile (H)             EML - Email (I)
'     SB1 - Subject 1-6 (J)        PR1 - Price 1-6 (M)
'     SB2 - Subject 7-12 (O)       PR2 - Price 7-12 (R)
'     M01..M12 - Months (U-AF)     EX1..EX3 - Extra subjects (AK-AM)
'     DCL - Decline entry (Superadmin rejection)
'==========================

Option Explicit

' ========================================
' History Field Tags (Public Constants)
' ========================================

Public Const HT_FAMILY_ID As String = "FID"      ' Column A (1)
Public Const HT_PARENT As String = "PAR"         ' Column B (2)
Public Const HT_CHILD As String = "CHD"          ' Column D (4)
Public Const HT_BIRTHDATE As String = "DOB"      ' Column E (5)
Public Const HT_ADDRESS As String = "ADR"        ' Column F (6)
Public Const HT_PHONE As String = "TEL"          ' Column G (7)
Public Const HT_MOBILE As String = "MOB"         ' Column H (8)
Public Const HT_EMAIL As String = "EML"          ' Column I (9)
Public Const HT_SUBJECT1 As String = "SB1"       ' Column J (10)
Public Const HT_PRICE1 As String = "PR1"         ' Column M (13)
Public Const HT_SUBJECT2 As String = "SB2"       ' Column O (15)
Public Const HT_PRICE2 As String = "PR2"         ' Column R (18)
Public Const HT_MONTH_PREFIX As String = "M"     ' M01-M12, Columns U-AF (21-32)
Public Const HT_EXTRA1 As String = "EX1"         ' Column AK (37)
Public Const HT_EXTRA2 As String = "EX2"         ' Column AL (38)
Public Const HT_EXTRA3 As String = "EX3"         ' Column AM (39)
Public Const HT_DECLINE As String = "DCL"        ' Decline entry marker

' ========================================
' History Delimiters (Public Constants)
' ========================================

Public Const HD_SESSION As String = "||"         ' Session separator
Public Const HD_FIELD As String = ";"            ' Field separator within session
Public Const HD_VALUE As String = "->"           ' Old->New value separator
Public Const HD_COMMENT_START As String = "/@"   ' Comment start marker
Public Const HD_COMMENT_END As String = "@/"     ' Comment end marker
Public Const HD_RUCK_PREFIX As String = "RUCK:"  ' Retroactive change prefix

' ========================================
' Column Mappings (Private Constants)
' ========================================

Private Const COL_FAMILY_ID As Integer = 1       ' A
Private Const COL_PARENT As Integer = 2          ' B
Private Const COL_CHILD As Integer = 4           ' D
Private Const COL_BIRTHDATE As Integer = 5       ' E
Private Const COL_ADDRESS As Integer = 6         ' F
Private Const COL_PHONE As Integer = 7           ' G
Private Const COL_MOBILE As Integer = 8          ' H
Private Const COL_EMAIL As Integer = 9           ' I
Private Const COL_SUBJECT1 As Integer = 10       ' J
Private Const COL_PRICE1 As Integer = 13         ' M
Private Const COL_SUBJECT2 As Integer = 15       ' O
Private Const COL_PRICE2 As Integer = 18         ' R
Private Const COL_MONTH_START As Integer = 21    ' U (Month 1)
Private Const COL_MONTH_END As Integer = 32      ' AF (Month 12)
Private Const COL_EXTRA1 As Integer = 37         ' AK
Private Const COL_EXTRA2 As Integer = 38         ' AL
Private Const COL_EXTRA3 As Integer = 39         ' AM
Private Const COL_HISTORY As Integer = 52        ' AZ

' ========================================
' Field Definition Type
' ========================================

Private Type FieldDef
    Tag As String
    Column As Integer
    Description As String
End Type

' ========================================
' Main History Building Functions
' ========================================

Public Function BuildHistoryEntry(ByVal sh As Worksheet, _
                                  ByVal shOriginal As Worksheet, _
                                  ByVal row As Long, _
                                  ByVal isRuck As Boolean, _
                                  ByVal userComment As String, _
                                  ByRef hasChanges As Boolean, _
                                  ByRef changedMonthsArr() As Integer) As String
    ' Builds a complete history entry for all changed fields in a row
    ' Returns formatted history string ready to append to existing history
    '
    ' Parameters:
    '   sh          - Current Kartei worksheet
    '   shOriginal  - Kartei_Original worksheet for comparison
    '   row         - Row number to process
    '   isRuck      - True if this is a retroactive (past month) change
    '   userComment - Comment from Notitzen (already cleaned)
    '   hasChanges  - OUT: Set to True if any changes detected
    '   changedMonthsArr - OUT: Array(1 To 12) with 1 for changed months
    '
    ' Returns:
    '   Formatted history entry string, or empty string if no changes
    
    Dim entryParts As Collection
    Set entryParts = New Collection
    
    hasChanges = False
    ReDim changedMonthsArr(1 To 12)
    
    ' Check all tracked fields for changes
    Call CheckAndAddFieldChange(sh, shOriginal, row, HT_FAMILY_ID, COL_FAMILY_ID, entryParts, hasChanges)
    Call CheckAndAddFieldChange(sh, shOriginal, row, HT_PARENT, COL_PARENT, entryParts, hasChanges)
    Call CheckAndAddFieldChange(sh, shOriginal, row, HT_CHILD, COL_CHILD, entryParts, hasChanges)
    Call CheckAndAddFieldChange(sh, shOriginal, row, HT_BIRTHDATE, COL_BIRTHDATE, entryParts, hasChanges)
    Call CheckAndAddFieldChange(sh, shOriginal, row, HT_ADDRESS, COL_ADDRESS, entryParts, hasChanges)
    Call CheckAndAddFieldChange(sh, shOriginal, row, HT_PHONE, COL_PHONE, entryParts, hasChanges)
    Call CheckAndAddFieldChange(sh, shOriginal, row, HT_MOBILE, COL_MOBILE, entryParts, hasChanges)
    Call CheckAndAddFieldChange(sh, shOriginal, row, HT_EMAIL, COL_EMAIL, entryParts, hasChanges)
    Call CheckAndAddFieldChange(sh, shOriginal, row, HT_SUBJECT1, COL_SUBJECT1, entryParts, hasChanges)
    Call CheckAndAddFieldChange(sh, shOriginal, row, HT_PRICE1, COL_PRICE1, entryParts, hasChanges)
    Call CheckAndAddFieldChange(sh, shOriginal, row, HT_SUBJECT2, COL_SUBJECT2, entryParts, hasChanges)
    Call CheckAndAddFieldChange(sh, shOriginal, row, HT_PRICE2, COL_PRICE2, entryParts, hasChanges)
    
    ' Check months (U-AF, columns 21-32)
    Dim m As Integer
    For m = 1 To 12
        Dim monthCol As Integer
        monthCol = COL_MONTH_START + m - 1
        
        Dim monthTag As String
        monthTag = HT_MONTH_PREFIX & Format(m, "00")  ' M01, M02, ..., M12
        
        Dim monthChanged As Boolean
        monthChanged = False
        Call CheckAndAddFieldChange(sh, shOriginal, row, monthTag, monthCol, entryParts, monthChanged)
        
        If monthChanged Then
            changedMonthsArr(m) = 1
            hasChanges = True
        End If
    Next m
    
    ' Check extra subjects (AK, AL, AM)
    Call CheckAndAddFieldChange(sh, shOriginal, row, HT_EXTRA1, COL_EXTRA1, entryParts, hasChanges)
    Call CheckAndAddFieldChange(sh, shOriginal, row, HT_EXTRA2, COL_EXTRA2, entryParts, hasChanges)
    Call CheckAndAddFieldChange(sh, shOriginal, row, HT_EXTRA3, COL_EXTRA3, entryParts, hasChanges)
    
    ' If no changes, return empty string
    If entryParts.Count = 0 Then
        BuildHistoryEntry = ""
        Exit Function
    End If
    
    ' Build the entry string
    Dim result As String
    result = ""
    
    ' Add RUCK prefix if retroactive change
    If isRuck Then
        result = HD_RUCK_PREFIX
    End If
    
    ' Join all field changes with semicolon separator
    Dim i As Long
    For i = 1 To entryParts.Count
        If i > 1 Then
            result = result & HD_FIELD
        End If
        result = result & entryParts(i)
    Next i
    
    ' Add comment and date
    result = result & HD_COMMENT_START & SanitizeComment(userComment) & HD_COMMENT_END & Format(Date, "dd.mm.yyyy") & HD_SESSION
    
    BuildHistoryEntry = result
End Function

Public Function BuildDeclineEntry(ByVal declineNumber As Long, _
                                  ByVal declineComment As String) As String
    ' Builds a decline history entry in the new format
    ' Format: DCL(<N>-><comment (Declined by Superadmin on DD.MM.YYYY)>)||
    '
    ' Parameters:
    '   declineNumber  - Sequential decline number (Decl_N)
    '   declineComment - Superadmin's decline comment
    '
    ' Returns:
    '   Formatted decline entry string
    
    Dim fullComment As String
    fullComment = SanitizeComment(declineComment) & " (Declined by Superadmin on " & Format(Date, "dd.mm.yyyy") & ")"
    
    BuildDeclineEntry = HT_DECLINE & "(" & CStr(declineNumber) & HD_VALUE & fullComment & ")" & HD_SESSION
End Function

Public Function AppendHistoryEntry(ByVal existingHistory As String, _
                                   ByVal newEntry As String) As String
    ' Appends a new history entry to existing history string
    ' Handles proper spacing and formatting
    
    If Len(Trim(existingHistory)) = 0 Then
        AppendHistoryEntry = newEntry
    Else
        ' Ensure existing history ends properly before appending
        existingHistory = Trim(existingHistory)
        
        ' If existing doesn't end with session separator, add space
        If Right(existingHistory, Len(HD_SESSION)) <> HD_SESSION Then
            existingHistory = existingHistory & " "
        End If
        
        AppendHistoryEntry = existingHistory & newEntry
    End If
End Function

' ========================================
' Field Change Detection
' ========================================

Private Sub CheckAndAddFieldChange(ByVal sh As Worksheet, _
                                   ByVal shOriginal As Worksheet, _
                                   ByVal row As Long, _
                                   ByVal tag As String, _
                                   ByVal col As Integer, _
                                   ByVal entryParts As Collection, _
                                   ByRef hasChanges As Boolean)
    ' Checks if a specific field has changed and adds it to the entry parts collection
    
    If Not SafeCompareValues(sh.Cells(row, col), shOriginal.Cells(row, col)) Then
        Dim oldVal As String
        Dim newVal As String
        
        oldVal = SanitizeFieldValue(CStr(GetCellValueSafe(shOriginal.Cells(row, col))))
        newVal = SanitizeFieldValue(CStr(GetCellValueSafe(sh.Cells(row, col))))
        
        Dim fieldEntry As String
        fieldEntry = tag & "(" & oldVal & HD_VALUE & newVal & ")"
        
        entryParts.Add fieldEntry
        hasChanges = True
    End If
End Sub

Private Function SafeCompareValues(ByVal cell1 As Range, ByVal cell2 As Range) As Boolean
    ' Safely compares two cell values, handling errors
    
    If IsError(cell1.Value) And IsError(cell2.Value) Then
        SafeCompareValues = True
    ElseIf IsError(cell1.Value) Or IsError(cell2.Value) Then
        SafeCompareValues = False
    Else
        SafeCompareValues = (CStr(GetCellValueSafe(cell1)) = CStr(GetCellValueSafe(cell2)))
    End If
End Function

Private Function GetCellValueSafe(ByVal cell As Range) As Variant
    ' Gets cell value safely, converting errors and nulls to empty string
    
    On Error Resume Next
    If IsError(cell.Value) Then
        GetCellValueSafe = ""
    ElseIf IsNull(cell.Value) Or IsEmpty(cell.Value) Then
        GetCellValueSafe = ""
    Else
        GetCellValueSafe = cell.Value
    End If
    On Error GoTo 0
End Function

' ========================================
' Sanitization Functions
' ========================================

Private Function SanitizeFieldValue(ByVal value As String) As String
    ' Sanitizes a field value for inclusion in history string
    ' Removes/escapes characters that could break parsing
    
    Dim result As String
    result = value
    
    ' Remove/replace problematic characters
    result = Replace(result, HD_VALUE, "~>")       ' Escape arrow
    result = Replace(result, HD_FIELD, ",")        ' Escape semicolon
    result = Replace(result, HD_SESSION, "|")      ' Escape double pipe
    result = Replace(result, "(", "[")             ' Escape open paren
    result = Replace(result, ")", "]")             ' Escape close paren
    result = Replace(result, vbCrLf, " ")          ' Remove line breaks
    result = Replace(result, vbCr, " ")
    result = Replace(result, vbLf, " ")
    
    SanitizeFieldValue = Trim(result)
End Function

Private Function SanitizeComment(ByVal comment As String) As String
    ' Sanitizes user comment for inclusion in history string
    
    Dim result As String
    result = comment
    
    ' Remove line breaks
    result = Replace(result, vbCrLf, " ")
    result = Replace(result, vbCr, " ")
    result = Replace(result, vbLf, " ")
    
    ' Remove/replace markers that could break parsing
    result = Replace(result, HD_COMMENT_START, "/")
    result = Replace(result, HD_COMMENT_END, "/")
    result = Replace(result, HD_SESSION, "|")
    
    SanitizeComment = Trim(result)
End Function

' ========================================
' Month Analysis Functions
' ========================================

Public Function HasPastMonthChanges(ByVal sh As Worksheet, _
                                    ByVal shOriginal As Worksheet, _
                                    ByVal row As Long, _
                                    ByVal forbiddenStartCol As Integer, _
                                    ByVal forbiddenEndCol As Integer, _
                                    ByRef changedPastMonths() As Integer) As Boolean
    ' Checks if there are changes in past (forbidden) months
    ' Used to determine if RUCK prefix is needed
    '
    ' Parameters:
    '   sh, shOriginal  - Worksheets for comparison
    '   row             - Row to check
    '   forbiddenStartCol - Start column of forbidden range (21-32)
    '   forbiddenEndCol   - End column of forbidden range (21-32)
    '   changedPastMonths - OUT: Array(1 To 12) with 1 for changed past months
    '
    ' Returns:
    '   True if any past month has changes
    
    ReDim changedPastMonths(1 To 12)
    HasPastMonthChanges = False
    
    If forbiddenStartCol < COL_MONTH_START Or forbiddenEndCol > COL_MONTH_END Then
        Exit Function
    End If
    
    If forbiddenStartCol > forbiddenEndCol Then
        Exit Function
    End If
    
    Dim col As Integer
    For col = forbiddenStartCol To forbiddenEndCol
        If Not SafeCompareValues(sh.Cells(row, col), shOriginal.Cells(row, col)) Then
            Dim monthNum As Integer
            monthNum = col - COL_MONTH_START + 1
            changedPastMonths(monthNum) = 1
            HasPastMonthChanges = True
        End If
    Next col
End Function

Public Function GetAllChangedMonths(ByVal sh As Worksheet, _
                                    ByVal shOriginal As Worksheet, _
                                    ByVal row As Long, _
                                    ByRef changedMonths() As Integer) As Boolean
    ' Gets all changed months (1-12) regardless of past/future
    '
    ' Returns:
    '   True if any month has changes
    
    ReDim changedMonths(1 To 12)
    GetAllChangedMonths = False
    
    Dim col As Integer
    For col = COL_MONTH_START To COL_MONTH_END
        If Not SafeCompareValues(sh.Cells(row, col), shOriginal.Cells(row, col)) Then
            Dim monthNum As Integer
            monthNum = col - COL_MONTH_START + 1
            changedMonths(monthNum) = 1
            GetAllChangedMonths = True
        End If
    Next col
End Function

' ========================================
' Decline Counter Function
' ========================================

Public Function CountDeclineEntries(ByVal historyStr As String) As Long
    ' Counts existing DCL (decline) entries in history string
    ' Supports both old format (Decl_N:) and new format (DCL(N->...))
    
    Dim count As Long
    count = 0
    
    ' Count new format: DCL(
    Dim pos As Long
    pos = 1
    Do
        pos = InStr(pos, historyStr, HT_DECLINE & "(")
        If pos > 0 Then
            count = count + 1
            pos = pos + Len(HT_DECLINE) + 1
        End If
    Loop While pos > 0
    
    ' Count old format: Decl_
    pos = 1
    Do
        pos = InStr(pos, historyStr, "Decl_")
        If pos > 0 Then
            count = count + 1
            pos = pos + 5
        End If
    Loop While pos > 0
    
    CountDeclineEntries = count
End Function

' ========================================
' Utility Functions for External Use
' ========================================

Public Function GetMonthTag(ByVal monthNumber As Integer) As String
    ' Returns the tag for a specific month (1-12)
    ' E.g., 1 -> "M01", 12 -> "M12"
    
    If monthNumber < 1 Or monthNumber > 12 Then
        GetMonthTag = ""
    Else
        GetMonthTag = HT_MONTH_PREFIX & Format(monthNumber, "00")
    End If
End Function

Public Function GetColumnForTag(ByVal tag As String) As Integer
    ' Returns the column number for a given tag
    ' Returns 0 if tag is not recognized
    
    Select Case tag
        Case HT_FAMILY_ID: GetColumnForTag = COL_FAMILY_ID
        Case HT_PARENT: GetColumnForTag = COL_PARENT
        Case HT_CHILD: GetColumnForTag = COL_CHILD
        Case HT_BIRTHDATE: GetColumnForTag = COL_BIRTHDATE
        Case HT_ADDRESS: GetColumnForTag = COL_ADDRESS
        Case HT_PHONE: GetColumnForTag = COL_PHONE
        Case HT_MOBILE: GetColumnForTag = COL_MOBILE
        Case HT_EMAIL: GetColumnForTag = COL_EMAIL
        Case HT_SUBJECT1: GetColumnForTag = COL_SUBJECT1
        Case HT_PRICE1: GetColumnForTag = COL_PRICE1
        Case HT_SUBJECT2: GetColumnForTag = COL_SUBJECT2
        Case HT_PRICE2: GetColumnForTag = COL_PRICE2
        Case HT_EXTRA1: GetColumnForTag = COL_EXTRA1
        Case HT_EXTRA2: GetColumnForTag = COL_EXTRA2
        Case HT_EXTRA3: GetColumnForTag = COL_EXTRA3
        Case Else
            ' Check for month tags M01-M12
            If Left(tag, 1) = HT_MONTH_PREFIX And Len(tag) = 3 Then
                Dim monthNum As Integer
                On Error Resume Next
                monthNum = CInt(Mid(tag, 2, 2))
                On Error GoTo 0
                
                If monthNum >= 1 And monthNum <= 12 Then
                    GetColumnForTag = COL_MONTH_START + monthNum - 1
                Else
                    GetColumnForTag = 0
                End If
            Else
                GetColumnForTag = 0
            End If
    End Select
End Function

Public Function GetTagDescription(ByVal tag As String) As String
    ' Returns human-readable description for a tag
    
    Select Case tag
        Case HT_FAMILY_ID: GetTagDescription = "Family ID"
        Case HT_PARENT: GetTagDescription = "Parent Name"
        Case HT_CHILD: GetTagDescription = "Child Name"
        Case HT_BIRTHDATE: GetTagDescription = "Date of Birth"
        Case HT_ADDRESS: GetTagDescription = "Address"
        Case HT_PHONE: GetTagDescription = "Phone"
        Case HT_MOBILE: GetTagDescription = "Mobile"
        Case HT_EMAIL: GetTagDescription = "Email"
        Case HT_SUBJECT1: GetTagDescription = "Subject (Months 1-6)"
        Case HT_PRICE1: GetTagDescription = "Price (Months 1-6)"
        Case HT_SUBJECT2: GetTagDescription = "Subject (Months 7-12)"
        Case HT_PRICE2: GetTagDescription = "Price (Months 7-12)"
        Case HT_EXTRA1: GetTagDescription = "Extra Subject 1"
        Case HT_EXTRA2: GetTagDescription = "Extra Subject 2"
        Case HT_EXTRA3: GetTagDescription = "Extra Subject 3"
        Case HT_DECLINE: GetTagDescription = "Decline Entry"
        Case Else
            ' Check for month tags
            If Left(tag, 1) = HT_MONTH_PREFIX And Len(tag) = 3 Then
                Dim monthNum As Integer
                On Error Resume Next
                monthNum = CInt(Mid(tag, 2, 2))
                On Error GoTo 0
                
                If monthNum >= 1 And monthNum <= 12 Then
                    Dim monthNames As Variant
                    monthNames = Array("January", "February", "March", "April", "May", "June", _
                                       "July", "August", "September", "October", "November", "December")
                    GetTagDescription = "Month " & monthNum & " (" & monthNames(monthNum - 1) & ")"
                Else
                    GetTagDescription = "Unknown"
                End If
            Else
                GetTagDescription = "Unknown"
            End If
    End Select
End Function

