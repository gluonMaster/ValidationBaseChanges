Attribute VB_Name = "NHblank_TeacherSync"
Option Explicit

' Data structure to hold record information for matching
Private Type RecordInfo
    Row As Long
    ID As String
    lastName As String
    firstName As String
    discipline As String
    Teacher As String
End Type

' Main entry point for syncing teacher names from Kartei to Kinder
Public Sub SyncTeacherNamesFromKartei()
    Dim screenUpdateState As Boolean
    Dim calculationState As XlCalculation
    Dim eventsState As Boolean
    
    On Error GoTo ErrorHandler
    
    ' Save current Excel state
    screenUpdateState = Application.ScreenUpdating
    calculationState = Application.Calculation
    eventsState = Application.EnableEvents
    
    ' Optimize performance
    Application.ScreenUpdating = False
    Application.Calculation = xlCalculationManual
    Application.EnableEvents = False
    
    ' Execute synchronization
    If Not ExecuteSync() Then
        ' Error occurred - message already shown in ExecuteSync
        GoTo CleanExit
    End If
    
CleanExit:
    ' Restore Excel state
    Application.ScreenUpdating = screenUpdateState
    Application.Calculation = calculationState
    Application.EnableEvents = eventsState
    Exit Sub
    
ErrorHandler:
    ' Restore Excel state even on error
    Application.ScreenUpdating = screenUpdateState
    Application.Calculation = calculationState
    Application.EnableEvents = eventsState
    
    MsgBox "Fehler beim Ausfuehren des Makros: " & Err.Description, vbCritical, "Fehler"
End Sub

' Execute the synchronization operation
' Returns True if successful
Private Function ExecuteSync() As Boolean
    Dim wbKartei As Workbook
    Dim wbKinder As Workbook
    Dim wsKartei As Worksheet
    Dim wsKinder As Worksheet
    Dim kinderFilterWasActive As Boolean
    Dim karteiFilterWasActive As Boolean
    Dim lastRow As Long
    Dim currentRow As Long
    Dim recordsProcessed As Long
    Dim recordsUpdated As Long
    Dim notFoundRows As String
    Dim response As VbMsgBoxResult
    Dim referenceDate As Variant
    Dim semesterMonth As Integer
    
    ExecuteSync = False
    recordsProcessed = 0
    recordsUpdated = 0
    notFoundRows = ""
    
    ' Check if we're on the correct file (Nachhilfe2425.xlsm)
    Set wbKinder = ActiveWorkbook
    If InStr(1, wbKinder.name, "Nachhilfe", vbTextCompare) = 0 Then
        MsgBox "Bitte fuehren Sie das Makro aus der Datei Nachhilfe2425.xlsm aus.", vbExclamation, "Falsche Datei"
        Exit Function
    End If
    
    ' Get Kinder worksheet
    On Error Resume Next
    Set wsKinder = wbKinder.Worksheets("Kinder")
    On Error GoTo 0
    
    If wsKinder Is Nothing Then
        MsgBox "Worksheet 'Kinder' nicht gefunden!", vbCritical, "Fehler"
        Exit Function
    End If
    
    ' Check if source file (KindElternDaten_25_Admin.xlsm) is open
    On Error Resume Next
    Set wbKartei = Workbooks("KindElternDaten_25_Admin.xlsm")
    On Error GoTo 0
    
    If wbKartei Is Nothing Then
        MsgBox "Quelldatei 'KindElternDaten_25_Admin.xlsm' ist nicht geoeffnet!" & vbCrLf & vbCrLf & _
               "Bitte oeffnen Sie die Datei und versuchen Sie erneut.", vbExclamation, "Datei nicht geoeffnet"
        Exit Function
    End If
    
    ' Get Kartei worksheet from source file
    On Error Resume Next
    Set wsKartei = wbKartei.Worksheets("Kartei")
    On Error GoTo 0
    
    If wsKartei Is Nothing Then
        MsgBox "Worksheet 'Kartei' nicht in Quelldatei gefunden!", vbCritical, "Fehler"
        Exit Function
    End If
    
    ' Get reference date from T2
    On Error Resume Next
    referenceDate = wsKinder.Range("T2").value
    On Error GoTo 0
    
    If Not IsDate(referenceDate) Then
        MsgBox "Ungueltige Referenzdatum in Zelle T2: " & referenceDate, vbExclamation, "Fehler"
        Exit Function
    End If
    
    ' Determine semester from reference date
    semesterMonth = Month(CDate(referenceDate))
    
    ' Ask for confirmation
    response = MsgBox("Moechten Sie die Lehrernamen aus Kartei synchronisieren?" & vbCrLf & vbCrLf & _
                      "Das Makro wird alle leeren Lehrer-Felder (Spalte I) ausfuellen," & vbCrLf & _
                      "basierend auf den Daten aus KindElternDaten_25_Admin.xlsm.", _
                      vbYesNo + vbQuestion, "Lehrernamen synchronisieren?")
    
    If response <> vbYes Then
        MsgBox "Synchronisation abgebrochen.", vbInformation, "Abgebrochen"
        Exit Function
    End If
    
    ' Check and temporarily disable filters on Kinder worksheet
    kinderFilterWasActive = False
    If wsKinder.AutoFilterMode Then
        If wsKinder.FilterMode Then
            kinderFilterWasActive = True
            wsKinder.ShowAllData
        End If
    End If
    
    ' Check and temporarily disable filters on Kartei worksheet
    karteiFilterWasActive = False
    If wsKartei.AutoFilterMode Then
        If wsKartei.FilterMode Then
            karteiFilterWasActive = True
            wsKartei.ShowAllData
        End If
    End If
    
    ' Find last row with data in Kinder (check columns B and C)
    lastRow = wsKinder.Cells(wsKinder.Rows.Count, "B").End(xlUp).Row
    If wsKinder.Cells(wsKinder.Rows.Count, "C").End(xlUp).Row > lastRow Then
        lastRow = wsKinder.Cells(wsKinder.Rows.Count, "C").End(xlUp).Row
    End If
    
    ' Process all records starting from row 5
    For currentRow = 5 To lastRow
        ' Check if row has data (both ID and LastName must be present)
        If Trim(CStr(wsKinder.Cells(currentRow, "B").value)) <> "" And _
           Trim(CStr(wsKinder.Cells(currentRow, "C").value)) <> "" Then
            
            ' Check if Teacher field (column I) is empty
            If Trim(CStr(wsKinder.Cells(currentRow, "I").value)) = "" Then
                recordsProcessed = recordsProcessed + 1
                
                ' Try to find and fill teacher name
                If FillTeacherName(wsKinder, wsKartei, currentRow, semesterMonth) Then
                    recordsUpdated = recordsUpdated + 1
                Else
                    ' Teacher not found - add to list
                    If notFoundRows <> "" Then notFoundRows = notFoundRows & ", "
                    notFoundRows = notFoundRows & currentRow
                End If
            End If
        End If
    Next currentRow
    
    ' Restore filters on Kinder if they were active
    If kinderFilterWasActive Then
        MsgBox "Hinweis: Filter auf Kinder-Sheet wurden entfernt." & vbCrLf & _
               "Bitte wenden Sie Filter manuell erneut an.", vbInformation, "Filter entfernt"
    End If
    
    ' Restore filters on Kartei if they were active
    If karteiFilterWasActive Then
        ' Silently restore - user doesn't need notification for source file
    End If
    
    ' Show results
    Dim resultMsg As String
    resultMsg = "Synchronisation abgeschlossen!" & vbCrLf & vbCrLf
    resultMsg = resultMsg & "Datensaetze mit leerem Lehrer-Feld: " & recordsProcessed & vbCrLf
    resultMsg = resultMsg & "Erfolgreich aktualisiert: " & recordsUpdated & vbCrLf
    
    If notFoundRows <> "" Then
        resultMsg = resultMsg & vbCrLf & "Lehrer nicht gefunden in Zeilen: " & notFoundRows
        MsgBox resultMsg, vbExclamation, "Synchronisation abgeschlossen"
    Else
        MsgBox resultMsg, vbInformation, "Erfolgreich"
    End If
    
    ExecuteSync = True
End Function

' Fill teacher name for a specific row in Kinder worksheet
' Returns True if teacher was found and filled
Private Function FillTeacherName( _
    ByVal wsKinder As Worksheet, _
    ByVal wsKartei As Worksheet, _
    ByVal kinderRow As Long, _
    ByVal semesterMonth As Integer _
) As Boolean
    
    Dim kinderRecord As RecordInfo
    Dim karteiRecord As RecordInfo
    Dim lastKarteiRow As Long
    Dim currentKarteiRow As Long
    Dim disciplineColKartei As String
    Dim teacherColKartei As String
    Dim foundMatch As Boolean
    
    FillTeacherName = False
    foundMatch = False
    
    ' Determine which semester columns to use in Kartei
    If semesterMonth >= 1 And semesterMonth <= 7 Then
        ' First semester
        disciplineColKartei = "J"
        teacherColKartei = "K"
    Else
        ' Second semester
        disciplineColKartei = "O"
        teacherColKartei = "P"
    End If
    
    ' Read record from Kinder
    With kinderRecord
        .Row = kinderRow
        .ID = Trim(CStr(wsKinder.Cells(kinderRow, "B").value))
        .lastName = Trim(CStr(wsKinder.Cells(kinderRow, "C").value))
        .firstName = Trim(CStr(wsKinder.Cells(kinderRow, "D").value))
        .discipline = Trim(CStr(wsKinder.Cells(kinderRow, "E").value))
    End With
    
    ' Validate that we have required data
    If kinderRecord.ID = "" Or kinderRecord.lastName = "" Or kinderRecord.discipline = "" Then
        Exit Function
    End If
    
    ' Find last row in Kartei
    lastKarteiRow = wsKartei.Cells(wsKartei.Rows.Count, "A").End(xlUp).Row
    
    ' Search for matching record in Kartei (starting from row 3)
    For currentKarteiRow = 3 To lastKarteiRow
        ' Read record from Kartei
        With karteiRecord
            .Row = currentKarteiRow
            .ID = Trim(CStr(wsKartei.Cells(currentKarteiRow, "A").value))
            .lastName = ""
            .firstName = ""
            .discipline = Trim(CStr(wsKartei.Cells(currentKarteiRow, disciplineColKartei).value))
            .Teacher = Trim(CStr(wsKartei.Cells(currentKarteiRow, teacherColKartei).value))
        End With
        
        ' Parse name from column D in Kartei
        Dim fullName As String
        fullName = Trim(CStr(wsKartei.Cells(currentKarteiRow, "D").value))
        If fullName <> "" Then
            ParseKarteiName fullName, karteiRecord.lastName, karteiRecord.firstName
        End If
        
        ' Check if this record matches
        If MatchRecords(kinderRecord, karteiRecord) Then
            ' Found match - fill teacher name
            If karteiRecord.Teacher <> "" Then
                wsKinder.Cells(kinderRow, "I").value = karteiRecord.Teacher
                FillTeacherName = True
                Exit Function
            End If
        End If
    Next currentKarteiRow
    
    ' No match found
    FillTeacherName = False
End Function

' Parse name from Kartei format into LastName and FirstName
Private Sub ParseKarteiName(ByVal fullName As String, ByRef lastName As String, ByRef firstName As String)
    Dim cleanedName As String
    Dim words() As String
    Dim i As Integer
    
    ' Clean up separators (replace commas and semicolons with spaces)
    cleanedName = Replace(fullName, ",", " ")
    cleanedName = Replace(cleanedName, ";", " ")
    
    ' Remove multiple spaces
    Do While InStr(cleanedName, "  ") > 0
        cleanedName = Replace(cleanedName, "  ", " ")
    Loop
    
    cleanedName = Trim(cleanedName)
    
    ' Split by spaces
    words = Split(cleanedName, " ")
    
    If UBound(words) < 0 Then
        lastName = ""
        firstName = ""
        Exit Sub
    End If
    
    ' First word is last name
    lastName = words(0)
    
    ' Remaining words are first name
    If UBound(words) >= 1 Then
        firstName = ""
        For i = 1 To UBound(words)
            If i > 1 Then firstName = firstName & " "
            firstName = firstName & words(i)
        Next i
    Else
        firstName = ""
    End If
End Sub

' Check if two records match based on ID, LastName, FirstName, and Discipline
' Returns True if records match
Private Function MatchRecords(ByRef kinderRecord As RecordInfo, ByRef karteiRecord As RecordInfo) As Boolean
    MatchRecords = False
    
    ' Check ID match
    If kinderRecord.ID <> karteiRecord.ID Then
        Exit Function
    End If
    
    ' Check LastName match (case-insensitive)
    If StrComp(kinderRecord.lastName, karteiRecord.lastName, vbTextCompare) <> 0 Then
        Exit Function
    End If
    
    ' Check FirstName match (case-insensitive)
    ' Allow empty first name in comparison
    If kinderRecord.firstName <> "" And karteiRecord.firstName <> "" Then
        If StrComp(kinderRecord.firstName, karteiRecord.firstName, vbTextCompare) <> 0 Then
            Exit Function
        End If
    End If
    
    ' Check Discipline match
    ' Kartei discipline contains Kinder discipline (e.g., "Nachhilfe Mathe" contains "Mathe")
    If kinderRecord.discipline <> "" Then
        If InStr(1, karteiRecord.discipline, kinderRecord.discipline, vbTextCompare) = 0 Then
            Exit Function
        End If
    Else
        Exit Function
    End If
    
    ' All checks passed
    MatchRecords = True
End Function
