Attribute VB_Name = "phone_Normalize"
'==========================
'   Module: phone_Normalize
'   Purpose: Normalize phone strings that may be corrupted into scientific notation
'
'   Background:
'   Phone numbers like "01761234567" can be stored in Access/Excel as "1.76E+9"
'   or "1,76E+9" due to auto-conversion. This module provides a normalizer that
'   detects such strings and expands them back to plain digit strings.
'
'   IMPORTANT LIMITATION:
'   If the DB already lost the leading zero (stored as "176..." instead of "0176..."),
'   we CANNOT reliably reconstruct it. This normalizer does NOT guess or prepend "0".
'   Only scientific notation expansion is performed.
'
'   Public API:
'   - NormalizePhoneText(v As Variant) As String
'       Normalizes a phone value, expanding scientific notation if detected.
'
'   - RepairPhoneScientificNotationForYear(year2 As Integer)
'       Manual repair macro for existing corrupted Access records.
'
'   Depends on: valid_YearConfig (for database path resolution)
'==========================

Option Explicit

' ============================================================
' PUBLIC API
' ============================================================

' Normalizes a phone value, expanding scientific notation strings to plain digit strings.
'
' Rules:
' 1) Preserve "+" if present at the start (e.g. "+49176...")
' 2) Remove spaces and common separators (space, -, (, ), /) but NOT "+"
' 3) If the remaining string matches scientific notation (e.g. "1,76E+9", "1.76E9"),
'    expand it to a full integer string via string-based math (no CDbl/Val)
' 4) If already a normal phone string, return unchanged
' 5) If normalization cannot be applied safely, return trimmed original
'
' @param v - The phone value (can be Variant, String, or number)
' @return String - Normalized phone string
Public Function NormalizePhoneText(ByVal v As Variant) As String
    ' Handle null/empty
    If IsNull(v) Or IsEmpty(v) Then
        NormalizePhoneText = ""
        Exit Function
    End If
    
    Dim rawStr As String
    rawStr = Trim$(CStr(v))
    
    If Len(rawStr) = 0 Then
        NormalizePhoneText = ""
        Exit Function
    End If
    
    ' Preserve leading "+" if present
    Dim hasPlus As Boolean
    hasPlus = (Left$(rawStr, 1) = "+")
    
    ' Remove common separators (space, -, (, ), /) but keep + and digits
    Dim cleaned As String
    cleaned = RemoveSeparators(rawStr)
    
    ' Check if it looks like scientific notation
    If LooksLikeScientificNotation(cleaned) Then
        Dim expanded As String
        expanded = ExpandScientificNotation(cleaned)
        
        ' If expansion succeeded, return it
        If Len(expanded) > 0 Then
            NormalizePhoneText = expanded
            Exit Function
        End If
    End If
    
    ' Check if it's already a normal phone string
    ' (starts with 0 or + and contains only digits/+)
    If IsNormalPhoneString(cleaned) Then
        NormalizePhoneText = cleaned
        Exit Function
    End If
    
    ' Cannot normalize safely - return the cleaned string
    NormalizePhoneText = cleaned
End Function

' ============================================================
' ACCESS REPAIR MACRO
' ============================================================

' Repairs scientific-notation phone values in Access tables for a specific year.
' Scans tblKartei, pre_tblKartei, and decl_tblKartei.
' Requires explicit user confirmation before running.
'
' @param year2 - Two-digit year (24, 25, or 26)
Public Sub RepairPhoneScientificNotationForYear(ByVal year2 As Integer)
    On Error GoTo ErrorHandler
    
    ' Validate year
    If Not valid_YearConfig.IsValidYear(year2) Then
        MsgBox "Ungueltiges Jahr: " & year2 & ". Unterstuetzte Jahre: 24, 25, 26.", _
               vbExclamation, "Reparatur - Jahresfehler"
        Exit Sub
    End If
    
    ' Get database path
    Dim dbPath As String
    dbPath = valid_YearConfig.GetDbPathForYear(year2)
    
    If dbPath = "" Then
        MsgBox "Datenbankpfad fuer Jahr " & year2 & " nicht konfiguriert.", _
               vbExclamation, "Reparatur - Pfadfehler"
        Exit Sub
    End If
    
    ' Confirm with user
    Dim confirmMsg As String
    confirmMsg = "ACHTUNG: Diese Funktion repariert Telefonnummern in wissenschaftlicher Notation" & vbCrLf & _
                 "(z.B. '1,76E+9' -> '176000000') in der Datenbank fuer Jahr 20" & Format(year2, "00") & "." & vbCrLf & vbCrLf & _
                 "Tabellen, die gescannt werden:" & vbCrLf & _
                 "  - tblKartei" & vbCrLf & _
                 "  - pre_tblKartei" & vbCrLf & _
                 "  - decl_tblKartei" & vbCrLf & vbCrLf & _
                 "Felder: Value7 (Tel.), Value8 (Handy)" & vbCrLf & vbCrLf & _
                 "HINWEIS: Verlorene fuehrende Nullen koennen NICHT wiederhergestellt werden." & vbCrLf & vbCrLf & _
                 "Moechten Sie fortfahren?"
    
    If MsgBox(confirmMsg, vbQuestion + vbYesNo, "Telefon-Reparatur bestaetigen") <> vbYes Then
        Exit Sub
    End If
    
    Application.ScreenUpdating = False
    
    Dim engine As DAO.DBEngine
    Set engine = New DAO.DBEngine
    
    Dim wsDao As DAO.Workspace
    Set wsDao = engine.Workspaces(0)
    
    Dim db As DAO.Database
    Set db = wsDao.OpenDatabase(dbPath)
    
    ' Repair each table
    Dim countMain As Long, countPre As Long, countDecl As Long
    
    countMain = RepairTablePhoneFields(db, "tblKartei")
    countPre = RepairTablePhoneFields(db, "pre_tblKartei")
    countDecl = RepairTablePhoneFields(db, "decl_tblKartei")
    
    db.Close
    
    Application.ScreenUpdating = True
    
    ' Show summary
    Dim summaryMsg As String
    summaryMsg = "Reparatur abgeschlossen fuer Jahr 20" & Format(year2, "00") & ":" & vbCrLf & vbCrLf & _
                 "tblKartei: " & countMain & " Zeile(n) korrigiert" & vbCrLf & _
                 "pre_tblKartei: " & countPre & " Zeile(n) korrigiert" & vbCrLf & _
                 "decl_tblKartei: " & countDecl & " Zeile(n) korrigiert" & vbCrLf & vbCrLf & _
                 "Gesamt: " & (countMain + countPre + countDecl) & " Zeile(n)"
    
    MsgBox summaryMsg, vbInformation, "Telefon-Reparatur Zusammenfassung"
    Exit Sub
    
ErrorHandler:
    Application.ScreenUpdating = True
    MsgBox "Fehler bei der Reparatur: " & Err.Description, vbCritical, "Reparaturfehler"
End Sub

' Wrapper macros for each year
Public Sub RepairPhoneScientificNotation24()
    RepairPhoneScientificNotationForYear 24
End Sub

Public Sub RepairPhoneScientificNotation25()
    RepairPhoneScientificNotationForYear 25
End Sub

Public Sub RepairPhoneScientificNotation26()
    RepairPhoneScientificNotationForYear 26
End Sub

' ============================================================
' PRIVATE HELPERS - Repair
' ============================================================

' Repairs phone fields (Value7, Value8) in a single table.
' Returns the number of rows that were modified.
Private Function RepairTablePhoneFields(ByVal db As DAO.Database, ByVal tableName As String) As Long
    On Error Resume Next
    
    RepairTablePhoneFields = 0
    
    ' Check if table exists
    Dim tdf As DAO.TableDef
    Set tdf = db.TableDefs(tableName)
    If Err.Number <> 0 Then
        Err.Clear
        Exit Function
    End If
    On Error GoTo 0
    
    Dim rs As DAO.Recordset
    Set rs = db.OpenRecordset("SELECT ID, Value7, Value8 FROM [" & tableName & "]", dbOpenDynaset)
    
    If rs.EOF Then
        rs.Close
        Exit Function
    End If
    
    Dim changedCount As Long
    changedCount = 0
    
    Do While Not rs.EOF
        Dim val7 As String, val8 As String
        Dim norm7 As String, norm8 As String
        Dim changed7 As Boolean, changed8 As Boolean
        
        val7 = NzStr(rs.Fields("Value7").Value)
        val8 = NzStr(rs.Fields("Value8").Value)
        
        ' Check if value contains scientific notation (E+ or E-)
        changed7 = ContainsScientificNotation(val7)
        changed8 = ContainsScientificNotation(val8)
        
        If changed7 Or changed8 Then
            norm7 = NormalizePhoneText(val7)
            norm8 = NormalizePhoneText(val8)
            
            ' Only update if normalization actually changed the value
            If (changed7 And norm7 <> val7) Or (changed8 And norm8 <> val8) Then
                rs.Edit
                If changed7 And norm7 <> val7 Then
                    rs.Fields("Value7").Value = norm7
                End If
                If changed8 And norm8 <> val8 Then
                    rs.Fields("Value8").Value = norm8
                End If
                rs.Update
                changedCount = changedCount + 1
            End If
        End If
        
        rs.MoveNext
    Loop
    
    rs.Close
    RepairTablePhoneFields = changedCount
End Function

' Checks if a string contains scientific notation marker (E+ or E-)
Private Function ContainsScientificNotation(ByVal s As String) As Boolean
    Dim upper As String
    upper = UCase$(s)
    ContainsScientificNotation = (InStr(1, upper, "E+") > 0) Or (InStr(1, upper, "E-") > 0)
End Function

' Null-safe string conversion
Private Function NzStr(ByVal v As Variant) As String
    If IsNull(v) Or IsEmpty(v) Then
        NzStr = ""
    Else
        NzStr = CStr(v)
    End If
End Function

' ============================================================
' PRIVATE HELPERS - Scientific Notation Detection
' ============================================================

' Removes common phone separators (space, -, (, ), /) but keeps + and digits
Private Function RemoveSeparators(ByVal s As String) As String
    Dim result As String
    Dim i As Long
    Dim ch As String
    
    result = ""
    For i = 1 To Len(s)
        ch = Mid$(s, i, 1)
        Select Case ch
            Case " ", "-", "(", ")", "/"
                ' Skip separators
            Case Else
                result = result & ch
        End Select
    Next i
    
    RemoveSeparators = result
End Function

' Checks if a string looks like scientific notation.
' Patterns: 1.76E+9, 1,76E+9, 1.76E9, 1.76E-2, etc.
Private Function LooksLikeScientificNotation(ByVal s As String) As Boolean
    LooksLikeScientificNotation = False
    
    If Len(s) = 0 Then Exit Function
    
    ' Must contain 'E' or 'e'
    Dim upperS As String
    upperS = UCase$(s)
    
    Dim posE As Long
    posE = InStr(1, upperS, "E")
    
    If posE = 0 Then Exit Function
    If posE = 1 Then Exit Function  ' Can't start with E
    If posE = Len(upperS) Then Exit Function  ' Can't end with E
    
    ' Check that there's a number before E
    Dim beforeE As String
    beforeE = Left$(s, posE - 1)
    
    ' Remove leading + if present
    If Left$(beforeE, 1) = "+" Then
        beforeE = Mid$(beforeE, 2)
    End If
    
    ' Replace comma with dot for decimal
    beforeE = Replace(beforeE, ",", ".")
    
    ' Check if beforeE is numeric-like (digits and at most one dot)
    If Not IsNumericLike(beforeE) Then Exit Function
    
    ' Check exponent part
    Dim afterE As String
    afterE = Mid$(upperS, posE + 1)
    
    ' Exponent can start with + or -
    If Left$(afterE, 1) = "+" Or Left$(afterE, 1) = "-" Then
        afterE = Mid$(afterE, 2)
    End If
    
    ' Exponent must be digits only
    If Not IsAllDigits(afterE) Then Exit Function
    If Len(afterE) = 0 Then Exit Function
    
    LooksLikeScientificNotation = True
End Function

' Checks if a string is numeric-like (digits and at most one decimal point)
Private Function IsNumericLike(ByVal s As String) As Boolean
    Dim i As Long
    Dim ch As String
    Dim dotCount As Long
    
    dotCount = 0
    
    For i = 1 To Len(s)
        ch = Mid$(s, i, 1)
        If ch = "." Then
            dotCount = dotCount + 1
            If dotCount > 1 Then
                IsNumericLike = False
                Exit Function
            End If
        ElseIf ch < "0" Or ch > "9" Then
            IsNumericLike = False
            Exit Function
        End If
    Next i
    
    IsNumericLike = (Len(s) > 0)
End Function

' Checks if a string contains only digits
Private Function IsAllDigits(ByVal s As String) As Boolean
    Dim i As Long
    Dim ch As String
    
    If Len(s) = 0 Then
        IsAllDigits = False
        Exit Function
    End If
    
    For i = 1 To Len(s)
        ch = Mid$(s, i, 1)
        If ch < "0" Or ch > "9" Then
            IsAllDigits = False
            Exit Function
        End If
    Next i
    
    IsAllDigits = True
End Function

' ============================================================
' PRIVATE HELPERS - Scientific Notation Expansion
' ============================================================

' Expands a scientific notation string to a plain integer string.
' Uses string-based math to avoid precision loss (no CDbl/Val).
'
' Examples:
'   "1.76E+9"  -> "1760000000"
'   "1,76E+09" -> "1760000000"
'   "1.76E9"   -> "1760000000"
'
' Returns empty string if expansion fails.
Private Function ExpandScientificNotation(ByVal s As String) As String
    ExpandScientificNotation = ""
    
    If Len(s) = 0 Then Exit Function
    
    ' Handle leading "+" for phone numbers
    Dim leadingPlus As Boolean
    Dim workStr As String
    
    leadingPlus = (Left$(s, 1) = "+")
    If leadingPlus Then
        workStr = Mid$(s, 2)
    Else
        workStr = s
    End If
    
    ' Find E position (case-insensitive)
    Dim upperWork As String
    upperWork = UCase$(workStr)
    
    Dim posE As Long
    posE = InStr(1, upperWork, "E")
    
    If posE = 0 Or posE = 1 Then Exit Function
    
    ' Extract mantissa and exponent
    Dim mantissaStr As String
    Dim exponentStr As String
    
    mantissaStr = Left$(workStr, posE - 1)
    exponentStr = Mid$(upperWork, posE + 1)
    
    ' Parse exponent (can have + or -)
    Dim expSign As Integer
    expSign = 1
    
    If Left$(exponentStr, 1) = "+" Then
        exponentStr = Mid$(exponentStr, 2)
    ElseIf Left$(exponentStr, 1) = "-" Then
        expSign = -1
        exponentStr = Mid$(exponentStr, 2)
    End If
    
    If Not IsAllDigits(exponentStr) Or Len(exponentStr) = 0 Then Exit Function
    
    Dim exponent As Long
    exponent = CLng(exponentStr) * expSign
    
    ' For phone numbers, we only handle positive exponents
    ' (scientific notation from large numbers, not small decimals)
    If exponent < 0 Then
        ' Negative exponent would make a decimal - not valid for phone
        Exit Function
    End If
    
    ' Replace comma with dot in mantissa
    mantissaStr = Replace(mantissaStr, ",", ".")
    
    ' Find decimal point position in mantissa
    Dim posDot As Long
    posDot = InStr(1, mantissaStr, ".")
    
    Dim intPart As String
    Dim fracPart As String
    
    If posDot > 0 Then
        intPart = Left$(mantissaStr, posDot - 1)
        fracPart = Mid$(mantissaStr, posDot + 1)
    Else
        intPart = mantissaStr
        fracPart = ""
    End If
    
    ' Remove leading zeros from intPart (but keep at least one digit)
    Do While Len(intPart) > 1 And Left$(intPart, 1) = "0"
        intPart = Mid$(intPart, 2)
    Loop
    
    ' Combine digits: shift decimal point right by exponent positions
    ' Example: 1.76E+9 -> intPart="1", fracPart="76", exponent=9
    ' We need to form: 1 + 76 + (9-2) zeros = 1760000000
    
    Dim allDigits As String
    allDigits = intPart & fracPart
    
    ' Remove any non-digit characters (shouldn't be any, but safety)
    Dim cleanDigits As String
    Dim i As Long
    Dim ch As String
    
    cleanDigits = ""
    For i = 1 To Len(allDigits)
        ch = Mid$(allDigits, i, 1)
        If ch >= "0" And ch <= "9" Then
            cleanDigits = cleanDigits & ch
        End If
    Next i
    
    ' The number of positions to shift is the exponent
    ' Current decimal position: after intPart (position = Len(intPart))
    ' Target decimal position: Len(intPart) + exponent
    '
    ' If exponent > Len(fracPart), we need to add zeros
    ' If exponent <= Len(fracPart), we insert decimal point (but for phones we want integer)
    
    Dim fracLen As Long
    fracLen = Len(fracPart)
    
    If exponent >= fracLen Then
        ' Need to add zeros
        Dim zerosNeeded As Long
        zerosNeeded = exponent - fracLen
        cleanDigits = cleanDigits & String$(zerosNeeded, "0")
    Else
        ' Exponent < fracLen means result would have decimals
        ' For phone numbers, we truncate (or could round, but truncate is safer)
        ' Actually, this shouldn't happen for real phone numbers
        ' Just take the integer part
        cleanDigits = Left$(cleanDigits, Len(intPart) + exponent)
    End If
    
    ' Restore leading + if present
    If leadingPlus Then
        ExpandScientificNotation = "+" & cleanDigits
    Else
        ExpandScientificNotation = cleanDigits
    End If
End Function

' ============================================================
' PRIVATE HELPERS - Phone String Validation
' ============================================================

' Checks if a string is a normal phone string.
' Normal phone string: starts with 0 or + and contains only digits/+
Private Function IsNormalPhoneString(ByVal s As String) As Boolean
    IsNormalPhoneString = False
    
    If Len(s) = 0 Then Exit Function
    
    Dim firstChar As String
    firstChar = Left$(s, 1)
    
    ' Should start with 0 or +
    If firstChar <> "0" And firstChar <> "+" Then Exit Function
    
    ' Rest should be digits only (after potential leading +)
    Dim checkPart As String
    If firstChar = "+" Then
        checkPart = Mid$(s, 2)
    Else
        checkPart = s
    End If
    
    If Not IsAllDigits(checkPart) Then Exit Function
    
    IsNormalPhoneString = True
End Function
