Attribute VB_Name = "Export_HistoryConverter"
'==========================
'   Module: Export_HistoryConverter
'   Purpose: Converts legacy history strings to new structured format
'   
'   Legacy Format Examples:
'     Mnt.8: War(12); Ist(24,6). /Comment/ 11.06.2025 ||
'     Ruck: Mnt.9: War(32,4); Ist(47). /Comment/ 23.11.2025 ||
'     Address: Was(old); Is(new). /Comment/ 01.01.2025 ||
'     Subject1: Was(old); Is(new). Subject2: Was(old); Is(new). /Comment/ 01.01.2025 ||
'     Decl_1: Was(); Is(comment). ||
'   
'   New Format:
'     M08(12->24,6)/@Comment@/11.06.2025||
'     RUCK:M09(32,4->47)/@Comment@/23.11.2025||
'     ADR(old->new)/@Comment@/01.01.2025||
'     SB1(old->new);SB2(old->new)/@Comment@/01.01.2025||
'     DCL(1->comment)||
'==========================

Option Explicit

' ========================================
' Legacy Pattern Constants
' ========================================

Private Const LEGACY_MONTH_PATTERN As String = "Mnt\.\d+:"
Private Const LEGACY_WAR_IST As String = "War\(([^)]*)\);\s*Ist\(([^)]*)\)\."
Private Const LEGACY_WAS_IS As String = "Was\(([^)]*)\);\s*Is\(([^)]*)\)\."
Private Const LEGACY_RUCK_PREFIX As String = "Ruck:"
Private Const LEGACY_ADDRESS As String = "Address:"
Private Const LEGACY_SUBJECT1 As String = "Subject1:"
Private Const LEGACY_SUBJECT2 As String = "Subject2:"
Private Const LEGACY_DECL_PATTERN As String = "Decl_(\d+):"
Private Const LEGACY_COMMENT_PATTERN As String = "/([^/]+)/\s*(\d{1,2}\.\d{1,2}\.\d{4})"
Private Const LEGACY_SESSION_SEP As String = " || "

' ========================================
' Main Conversion Functions
' ========================================

Public Function ConvertHistoryToNewFormat(ByVal legacyHistory As String) As String
    ' Converts entire legacy history string to new format
    ' Preserves session structure and all field data
    '
    ' Parameters:
    '   legacyHistory - Original history string in legacy format
    '
    ' Returns:
    '   Converted history string in new format
    
    If Len(Trim(legacyHistory)) = 0 Then
        ConvertHistoryToNewFormat = ""
        Exit Function
    End If
    
    ' Split by session separator
    Dim sessions As Variant
    sessions = Split(legacyHistory, "||")
    
    Dim result As String
    result = ""
    
    Dim i As Long
    For i = LBound(sessions) To UBound(sessions)
        Dim session As String
        session = Trim(CStr(sessions(i)))
        
        If Len(session) > 0 Then
            Dim convertedSession As String
            convertedSession = ConvertSingleSession(session)
            
            If Len(convertedSession) > 0 Then
                result = result & convertedSession
                
                ' Ensure session ends with separator
                If Right(result, Len(Export_HistoryBuilder.HD_SESSION)) <> Export_HistoryBuilder.HD_SESSION Then
                    result = result & Export_HistoryBuilder.HD_SESSION
                End If
            End If
        End If
    Next i
    
    ConvertHistoryToNewFormat = result
End Function

Private Function ConvertSingleSession(ByVal session As String) As String
    ' Converts a single history session to new format
    
    Dim result As String
    result = ""
    
    Dim isRuck As Boolean
    isRuck = False
    
    Dim fieldParts As Collection
    Set fieldParts = New Collection
    
    ' Check for Ruck prefix
    If InStr(1, session, "Ruck:", vbTextCompare) > 0 Then
        isRuck = True
        ' Remove "Ruck: " or "Ruck:" from session for processing
        session = Trim(Replace(session, "Ruck:", "", 1, 1, vbTextCompare))
    End If
    
    ' Extract comment and date first
    Dim commentText As String
    Dim dateText As String
    commentText = ""
    dateText = ""
    
    Call ExtractCommentAndDate(session, commentText, dateText)
    
    ' Remove comment/date part from session for field parsing
    If Len(commentText) > 0 Or Len(dateText) > 0 Then
        Dim commentPattern As String
        commentPattern = "/" & commentText & "/"
        
        Dim commentPos As Long
        commentPos = InStr(session, commentPattern)
        
        If commentPos > 0 Then
            session = Trim(Left(session, commentPos - 1))
        End If
    End If
    
    ' Parse month changes: Mnt.N: War(X); Ist(Y).
    Call ParseMonthChanges(session, fieldParts)
    
    ' Parse Address: Was(X); Is(Y).
    Call ParseFieldChange(session, "Address:", Export_HistoryBuilder.HT_ADDRESS, fieldParts)
    
    ' Parse Subject1: Was(X); Is(Y).
    Call ParseFieldChange(session, "Subject1:", Export_HistoryBuilder.HT_SUBJECT1, fieldParts)
    
    ' Parse Subject2: Was(X); Is(Y).
    Call ParseFieldChange(session, "Subject2:", Export_HistoryBuilder.HT_SUBJECT2, fieldParts)
    
    ' Parse Decl_N: Was(); Is(comment).
    Dim declResult As String
    declResult = ParseDeclineEntry(session)
    
    If Len(declResult) > 0 Then
        ' Decline entries are standalone, return directly
        ConvertSingleSession = declResult
        Exit Function
    End If
    
    ' Build result from field parts
    If fieldParts.Count = 0 Then
        ' No recognized fields, but might have comment only
        If Len(commentText) > 0 And Len(dateText) > 0 Then
            result = Export_HistoryBuilder.HD_COMMENT_START & SanitizeValue(commentText) & _
                     Export_HistoryBuilder.HD_COMMENT_END & dateText
        Else
            ConvertSingleSession = ""
            Exit Function
        End If
    Else
        ' Add RUCK prefix if needed
        If isRuck Then
            result = Export_HistoryBuilder.HD_RUCK_PREFIX
        End If
        
        ' Join field parts
        Dim j As Long
        For j = 1 To fieldParts.Count
            If j > 1 Then
                result = result & Export_HistoryBuilder.HD_FIELD
            End If
            result = result & fieldParts(j)
        Next j
        
        ' Add comment and date
        If Len(dateText) > 0 Then
            result = result & Export_HistoryBuilder.HD_COMMENT_START & _
                     SanitizeValue(commentText) & Export_HistoryBuilder.HD_COMMENT_END & dateText
        End If
    End If
    
    ConvertSingleSession = result
End Function

' ========================================
' Parsing Functions
' ========================================

Private Sub ParseMonthChanges(ByVal session As String, ByVal fieldParts As Collection)
    ' Parses all Mnt.N: War(X); Ist(Y). patterns and adds to collection
    
    Dim regex As Object
    Set regex = CreateObject("VBScript.RegExp")
    regex.Global = True
    regex.IgnoreCase = True
    regex.Pattern = "Mnt\.(\d+):\s*War\(([^)]*)\);\s*Ist\(([^)]*)\)\."
    
    Dim matches As Object
    Set matches = regex.Execute(session)
    
    Dim match As Object
    For Each match In matches
        If match.SubMatches.Count >= 3 Then
            Dim monthNum As Integer
            monthNum = CInt(match.SubMatches(0))
            
            Dim oldVal As String
            oldVal = SanitizeValue(CStr(match.SubMatches(1)))
            
            Dim newVal As String
            newVal = SanitizeValue(CStr(match.SubMatches(2)))
            
            ' Format month tag: M01, M02, ..., M12
            Dim monthTag As String
            monthTag = Export_HistoryBuilder.HT_MONTH_PREFIX & Format(monthNum, "00")
            
            Dim fieldEntry As String
            fieldEntry = monthTag & "(" & oldVal & Export_HistoryBuilder.HD_VALUE & newVal & ")"
            
            fieldParts.Add fieldEntry
        End If
    Next match
End Sub

Private Sub ParseFieldChange(ByVal session As String, _
                             ByVal legacyPrefix As String, _
                             ByVal newTag As String, _
                             ByVal fieldParts As Collection)
    ' Parses a single field change: FieldName: Was(X); Is(Y).
    
    Dim startPos As Long
    startPos = InStr(1, session, legacyPrefix, vbTextCompare)
    
    If startPos = 0 Then Exit Sub
    
    ' Extract the Was()/Is() part after the prefix
    Dim afterPrefix As String
    afterPrefix = Mid(session, startPos + Len(legacyPrefix))
    
    Dim regex As Object
    Set regex = CreateObject("VBScript.RegExp")
    regex.Global = False
    regex.IgnoreCase = True
    regex.Pattern = "\s*Was\(([^)]*)\);\s*Is\(([^)]*)\)\."
    
    Dim matches As Object
    Set matches = regex.Execute(afterPrefix)
    
    If matches.Count > 0 Then
        Dim match As Object
        Set match = matches(0)
        
        If match.SubMatches.Count >= 2 Then
            Dim oldVal As String
            oldVal = SanitizeValue(CStr(match.SubMatches(0)))
            
            Dim newVal As String
            newVal = SanitizeValue(CStr(match.SubMatches(1)))
            
            Dim fieldEntry As String
            fieldEntry = newTag & "(" & oldVal & Export_HistoryBuilder.HD_VALUE & newVal & ")"
            
            fieldParts.Add fieldEntry
        End If
    End If
End Sub

Private Function ParseDeclineEntry(ByVal session As String) As String
    ' Parses Decl_N: Was(); Is(comment). and converts to DCL(N->comment)
    
    Dim regex As Object
    Set regex = CreateObject("VBScript.RegExp")
    regex.Global = False
    regex.IgnoreCase = True
    regex.Pattern = "Decl_(\d+):\s*Was\(([^)]*)\);\s*Is\(([^)]*)\)\."
    
    Dim matches As Object
    Set matches = regex.Execute(session)
    
    If matches.Count > 0 Then
        Dim match As Object
        Set match = matches(0)
        
        If match.SubMatches.Count >= 3 Then
            Dim declNum As String
            declNum = CStr(match.SubMatches(0))
            
            ' SubMatches(1) is the "Was" part (usually empty)
            ' SubMatches(2) is the "Is" part (the comment)
            Dim declComment As String
            declComment = SanitizeValue(CStr(match.SubMatches(2)))
            
            ParseDeclineEntry = Export_HistoryBuilder.HT_DECLINE & "(" & declNum & _
                                Export_HistoryBuilder.HD_VALUE & declComment & ")"
            Exit Function
        End If
    End If
    
    ParseDeclineEntry = ""
End Function

Private Sub ExtractCommentAndDate(ByVal session As String, _
                                  ByRef commentText As String, _
                                  ByRef dateText As String)
    ' Extracts /Comment/ DD.MM.YYYY from session
    
    commentText = ""
    dateText = ""
    
    Dim regex As Object
    Set regex = CreateObject("VBScript.RegExp")
    regex.Global = False
    regex.IgnoreCase = False
    regex.Pattern = "/([^/]+)/\s*(\d{1,2}\.\d{1,2}\.\d{4})"
    
    Dim matches As Object
    Set matches = regex.Execute(session)
    
    If matches.Count > 0 Then
        Dim match As Object
        Set match = matches(0)
        
        If match.SubMatches.Count >= 2 Then
            commentText = CStr(match.SubMatches(0))
            dateText = CStr(match.SubMatches(1))
        End If
    End If
End Sub

' ========================================
' Sanitization
' ========================================

Private Function SanitizeValue(ByVal value As String) As String
    ' Sanitizes a value for inclusion in new format history string
    
    Dim result As String
    result = value
    
    ' Remove/replace problematic characters
    result = Replace(result, Export_HistoryBuilder.HD_VALUE, "~>")
    result = Replace(result, Export_HistoryBuilder.HD_FIELD, ",")
    result = Replace(result, Export_HistoryBuilder.HD_SESSION, "|")
    result = Replace(result, "(", "[")
    result = Replace(result, ")", "]")
    result = Replace(result, Export_HistoryBuilder.HD_COMMENT_START, "/")
    result = Replace(result, Export_HistoryBuilder.HD_COMMENT_END, "/")
    result = Replace(result, vbCrLf, " ")
    result = Replace(result, vbCr, " ")
    result = Replace(result, vbLf, " ")
    
    SanitizeValue = Trim(result)
End Function

' ========================================
' Batch Conversion Functions
' ========================================

Public Sub ConvertAllHistoriesInKartei()
    ' Converts all non-empty history strings in Kartei sheet to new format
    ' History is in column AZ (52)
    
    On Error GoTo ErrorHandler
    
    Dim ws As Worksheet
    Set ws = ThisWorkbook.Worksheets("Kartei")
    
    Application.ScreenUpdating = False
    Application.Calculation = xlCalculationManual
    
    Dim lastRow As Long
    lastRow = ws.Cells(ws.Rows.Count, 1).End(xlUp).Row
    
    If lastRow < 3 Then
        MsgBox "No data rows found in Kartei sheet.", vbInformation, "Conversion"
        GoTo Cleanup
    End If
    
    Dim convertedCount As Long
    Dim skippedCount As Long
    Dim errorCount As Long
    convertedCount = 0
    skippedCount = 0
    errorCount = 0
    
    Dim r As Long
    For r = 3 To lastRow
        Dim historyCell As Range
        Set historyCell = ws.Cells(r, 52)
        
        Dim currentHistory As String
        currentHistory = Trim(CStr(historyCell.Value))
        
        If Len(currentHistory) > 0 Then
            ' Check if already in new format (contains -> or /@)
            If IsAlreadyNewFormat(currentHistory) Then
                skippedCount = skippedCount + 1
            Else
                On Error Resume Next
                Dim newHistory As String
                newHistory = ConvertHistoryToNewFormat(currentHistory)
                
                If Err.Number <> 0 Then
                    errorCount = errorCount + 1
                    Debug.Print "Error converting row " & r & ": " & Err.Description
                    Err.Clear
                Else
                    If Len(newHistory) > 0 Then
                        historyCell.Value = newHistory
                        convertedCount = convertedCount + 1
                    Else
                        skippedCount = skippedCount + 1
                    End If
                End If
                On Error GoTo ErrorHandler
            End If
        End If
    Next r
    
    MsgBox "History conversion complete:" & vbCrLf & vbCrLf & _
           "  Converted: " & convertedCount & vbCrLf & _
           "  Skipped (already new format or empty): " & skippedCount & vbCrLf & _
           "  Errors: " & errorCount, _
           vbInformation, "Conversion Complete"
    
Cleanup:
    Application.Calculation = xlCalculationAutomatic
    Application.ScreenUpdating = True
    Exit Sub
    
ErrorHandler:
    MsgBox "Error during conversion: " & Err.Description, vbCritical, "Conversion Error"
    Resume Cleanup
End Sub

Public Sub ConvertHistoriesInAccessTable(ByVal tableName As String)
    ' Converts all non-empty history strings in specified Access table
    ' History is in Value52 field
    '
    ' Parameters:
    '   tableName - Name of the table (tblKartei, pre_tblKartei, or decl_tblKartei)
    
    On Error GoTo ErrorHandler
    
    Dim ws As Worksheet
    Set ws = ThisWorkbook.Worksheets("Kartei")
    
    Dim dbPath As String
    dbPath = ws.Range("X1").Value & "\Alarm\KindElternDaten_25_front.accdb"
    
    Dim engine As DAO.DBEngine
    Set engine = New DAO.DBEngine
    
    Dim wsDao As DAO.Workspace
    Set wsDao = engine.Workspaces(0)
    
    Dim db As DAO.Database
    Set db = wsDao.OpenDatabase(dbPath)
    
    ' Check if table exists
    Dim tblExists As Boolean
    tblExists = False
    
    Dim tbl As DAO.TableDef
    For Each tbl In db.TableDefs
        If tbl.Name = tableName Then
            tblExists = True
            Exit For
        End If
    Next tbl
    
    If Not tblExists Then
        MsgBox "Table '" & tableName & "' does not exist.", vbExclamation, "Table Not Found"
        db.Close
        Exit Sub
    End If
    
    Application.ScreenUpdating = False
    
    Dim rs As DAO.Recordset
    Set rs = db.OpenRecordset(tableName, dbOpenDynaset)
    
    Dim convertedCount As Long
    Dim skippedCount As Long
    Dim errorCount As Long
    convertedCount = 0
    skippedCount = 0
    errorCount = 0
    
    wsDao.BeginTrans
    
    Do While Not rs.EOF
        Dim currentHistory As String
        
        If IsNull(rs.Fields("Value52").Value) Then
            currentHistory = ""
        Else
            currentHistory = Trim(CStr(rs.Fields("Value52").Value))
        End If
        
        If Len(currentHistory) > 0 Then
            If IsAlreadyNewFormat(currentHistory) Then
                skippedCount = skippedCount + 1
            Else
                On Error Resume Next
                Dim newHistory As String
                newHistory = ConvertHistoryToNewFormat(currentHistory)
                
                If Err.Number <> 0 Then
                    errorCount = errorCount + 1
                    Debug.Print "Error converting ID " & rs.Fields("ID").Value & ": " & Err.Description
                    Err.Clear
                Else
                    If Len(newHistory) > 0 Then
                        rs.Edit
                        rs.Fields("Value52").Value = newHistory
                        rs.Update
                        convertedCount = convertedCount + 1
                    Else
                        skippedCount = skippedCount + 1
                    End If
                End If
                On Error GoTo ErrorHandler
            End If
        End If
        
        rs.MoveNext
    Loop
    
    wsDao.CommitTrans
    
    rs.Close
    db.Close
    
    MsgBox "History conversion in '" & tableName & "' complete:" & vbCrLf & vbCrLf & _
           "  Converted: " & convertedCount & vbCrLf & _
           "  Skipped (already new format or empty): " & skippedCount & vbCrLf & _
           "  Errors: " & errorCount, _
           vbInformation, "Conversion Complete"
    
    Application.ScreenUpdating = True
    Exit Sub
    
ErrorHandler:
    On Error Resume Next
    wsDao.Rollback
    rs.Close
    db.Close
    Application.ScreenUpdating = True
    MsgBox "Error during Access table conversion: " & Err.Description, vbCritical, "Conversion Error"
End Sub

Public Sub ConvertAllHistoriesEverywhere()
    ' Master procedure to convert histories in:
    ' 1. Current Kartei sheet
    ' 2. tblKartei in Access
    ' 3. pre_tblKartei in Access (if exists)
    ' 4. decl_tblKartei in Access (if exists)
    
    Dim result As VbMsgBoxResult
    result = MsgBox("This will convert all history strings to the new format in:" & vbCrLf & vbCrLf & _
                    "  - Kartei sheet (current workbook)" & vbCrLf & _
                    "  - tblKartei (Access database)" & vbCrLf & _
                    "  - pre_tblKartei (if exists)" & vbCrLf & _
                    "  - decl_tblKartei (if exists)" & vbCrLf & vbCrLf & _
                    "This operation cannot be undone. Continue?", _
                    vbYesNo + vbQuestion, "Convert All Histories")
    
    If result <> vbYes Then Exit Sub
    
    ' Convert Kartei sheet
    MsgBox "Step 1/4: Converting Kartei sheet...", vbInformation, "Progress"
    Call ConvertAllHistoriesInKartei
    
    ' Convert tblKartei
    MsgBox "Step 2/4: Converting tblKartei...", vbInformation, "Progress"
    Call ConvertHistoriesInAccessTable("tblKartei")
    
    ' Convert pre_tblKartei
    MsgBox "Step 3/4: Converting pre_tblKartei...", vbInformation, "Progress"
    On Error Resume Next
    Call ConvertHistoriesInAccessTable("pre_tblKartei")
    On Error GoTo 0
    
    ' Convert decl_tblKartei
    MsgBox "Step 4/4: Converting decl_tblKartei...", vbInformation, "Progress"
    On Error Resume Next
    Call ConvertHistoriesInAccessTable("decl_tblKartei")
    On Error GoTo 0
    
    MsgBox "All history conversions complete!", vbInformation, "Done"
End Sub

' ========================================
' Format Detection
' ========================================

Public Function IsAlreadyNewFormat(ByVal historyStr As String) As Boolean
    ' Checks if history string is already in new format
    ' New format indicators: contains -> or /@ or @/ or tags like M01( or ADR( etc.
    
    If Len(historyStr) = 0 Then
        IsAlreadyNewFormat = False
        Exit Function
    End If
    
    ' Check for new format delimiters
    If InStr(historyStr, Export_HistoryBuilder.HD_VALUE) > 0 Then
        ' Contains -> which is the new format value separator
        ' But need to check it's not just text containing ->
        ' Look for pattern like TAG(...)
        If InStr(historyStr, "M01(") > 0 Or InStr(historyStr, "M02(") > 0 Or _
           InStr(historyStr, "ADR(") > 0 Or InStr(historyStr, "SB1(") > 0 Or _
           InStr(historyStr, "DCL(") > 0 Or InStr(historyStr, "FID(") > 0 Or _
           InStr(historyStr, Export_HistoryBuilder.HD_COMMENT_START) > 0 Then
            IsAlreadyNewFormat = True
            Exit Function
        End If
    End If
    
    ' Check for new format comment markers
    If InStr(historyStr, Export_HistoryBuilder.HD_COMMENT_START) > 0 And _
       InStr(historyStr, Export_HistoryBuilder.HD_COMMENT_END) > 0 Then
        IsAlreadyNewFormat = True
        Exit Function
    End If
    
    IsAlreadyNewFormat = False
End Function

' ========================================
' Testing / Preview Functions
' ========================================

Public Function PreviewConversion(ByVal legacyHistory As String) As String
    ' Preview function for testing conversion without modifying data
    ' Use in Immediate Window: ?Export_HistoryConverter.PreviewConversion("...")
    
    PreviewConversion = ConvertHistoryToNewFormat(legacyHistory)
End Function

Public Sub TestConversionWithSample()
    ' Test conversion with sample legacy history string
    
    Dim sample As String
    sample = "Mnt.8: War(12); Ist(24,6). Mnt.9: War(0); Ist(24,6). /NeuSchuljahr/ 11.06.2025 || " & _
             "Ruck: Mnt.9: War(32,4); Ist(47). /Neu Prices/ 23.11.2025 || " & _
             "Decl_1: Was(); Is(Das ist SEPA (Declined by Superadmin on 23.11.2025)). || " & _
             "Address: Was(Alte Str 1); Is(Neue Str 2). /Umzug/ 01.01.2025 || "
    
    Dim converted As String
    converted = ConvertHistoryToNewFormat(sample)
    
    Debug.Print "=== ORIGINAL ===" & vbCrLf & sample
    Debug.Print vbCrLf & "=== CONVERTED ===" & vbCrLf & converted
    
    MsgBox "Original:" & vbCrLf & sample & vbCrLf & vbCrLf & _
           "Converted:" & vbCrLf & converted, _
           vbInformation, "Conversion Test"
End Sub

