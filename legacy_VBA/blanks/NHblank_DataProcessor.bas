Attribute VB_Name = "NHblank_DataProcessor"
Option Explicit

' Process records and create blanks
' Returns the number of blanks created
' activeRecordsFound - output parameter with count of active records found
Public Function ProcessRecords( _
    ByVal processAllRecords As Boolean, _
    ByVal templatePath As String, _
    ByVal targetFolder As String, _
    ByVal currentRowNumber As Long, _
    ByRef activeRecordsFound As Long _
) As Long
    
    Dim wsKinder As Worksheet
    Dim wbTemplate As Workbook
    Dim lastRow As Long
    Dim currentRow As Long
    Dim blanksCreated As Long
    Dim processError As String
    
    blanksCreated = 0
    activeRecordsFound = 0
    
    ' Get Kinder worksheet
    On Error Resume Next
    Set wsKinder = ThisWorkbook.Worksheets("Kinder")
    On Error GoTo 0
    
    If wsKinder Is Nothing Then
        ProcessRecords = 0
        activeRecordsFound = 0
        Exit Function
    End If
    
    ' Open template file once
    On Error GoTo TemplateError
    Set wbTemplate = Workbooks.Open(templatePath, ReadOnly:=False)
    On Error GoTo 0
    
    ' Verify Muster worksheet exists
    Dim wsTemplate As Worksheet
    On Error Resume Next
    Set wsTemplate = wbTemplate.Worksheets("Muster")
    On Error GoTo 0
    
    If wsTemplate Is Nothing Then
        wbTemplate.Close SaveChanges:=False
        NHblank_UI.ShowError "FEHLER: Worksheet 'Muster' nicht in Template-Datei gefunden!" & vbCrLf & _
                             "Template: " & templatePath
        ProcessRecords = 0
        activeRecordsFound = 0
        Exit Function
    End If
    
    If processAllRecords Then
        ' Process all active records
        lastRow = wsKinder.Cells(wsKinder.Rows.Count, "C").End(xlUp).Row
        
        For currentRow = 5 To lastRow
            ' Check if record is active (black font color)
            If IsRecordActive(wsKinder, currentRow) Then
                activeRecordsFound = activeRecordsFound + 1
                If ProcessSingleRecord(wsKinder, wbTemplate, currentRow, targetFolder, processError) Then
                    blanksCreated = blanksCreated + 1
                Else
                    ' Log error for this record
                    Debug.Print "Row " & currentRow & " failed: " & processError
                End If
            End If
        Next currentRow
    Else
        ' Process only current record
        ' Use the row number saved at the start of macro execution
        currentRow = currentRowNumber
        If currentRow > 4 Then
            If IsRecordActive(wsKinder, currentRow) Then
                activeRecordsFound = 1
                If ProcessSingleRecord(wsKinder, wbTemplate, currentRow, targetFolder, processError) Then
                    blanksCreated = blanksCreated + 1
                Else
                    ' Show specific error for single record
                    NHblank_UI.ShowError "Fehler beim Verarbeiten der Zeile " & currentRow & ":" & vbCrLf & processError
                End If
            End If
        End If
    End If
    
    ' Close template workbook without saving
    wbTemplate.Close SaveChanges:=False
    
    ProcessRecords = blanksCreated
    Exit Function
    
TemplateError:
    NHblank_UI.ShowError "FEHLER beim Oeffnen der Template-Datei:" & vbCrLf & _
                         templatePath & vbCrLf & vbCrLf & _
                         "Fehler: " & Err.Description
    ProcessRecords = 0
    activeRecordsFound = 0
End Function

' Debug function to check font color of current cell
Public Sub CheckCurrentCellColor()
    Dim msg As String
    Dim cell As Range
    Dim themeColorStr As String
    
    Set cell = ActiveCell
    
    ' Safely get ThemeColor value
    On Error Resume Next
    themeColorStr = CStr(cell.Font.ThemeColor)
    If Err.Number <> 0 Then
        themeColorStr = "Error reading ThemeColor"
        Err.Clear
    End If
    On Error GoTo 0
    
    msg = "Color Debug for cell " & cell.address & ":" & vbCrLf & vbCrLf
    msg = msg & "Font.ColorIndex: " & cell.Font.ColorIndex & vbCrLf
    msg = msg & "Font.Color: " & cell.Font.Color & vbCrLf
    msg = msg & "Font.ThemeColor: " & themeColorStr & vbCrLf
    msg = msg & "Cell Value: " & cell.value & vbCrLf
    msg = msg & "xlAutomatic value: " & xlAutomatic & vbCrLf & vbCrLf
    msg = msg & "Check results:" & vbCrLf
    msg = msg & "ColorIndex <> 15: " & (cell.Font.ColorIndex <> 15) & vbCrLf
    msg = msg & "ColorIndex <> 16: " & (cell.Font.ColorIndex <> 16) & vbCrLf
    msg = msg & "ColorIndex <> 48: " & (cell.Font.ColorIndex <> 48) & vbCrLf
    msg = msg & "Value <> empty: " & (cell.value <> "") & vbCrLf
    
    MsgBox msg, vbInformation, "Cell Color Info"
End Sub

' Check if record is active (NOT gray font color)
Private Function IsRecordActive(ByVal ws As Worksheet, ByVal rowNum As Long) As Boolean
    Dim cell As Range
    Set cell = ws.Cells(rowNum, "C")
    
    ' Check if cell is not empty
    If cell.value = "" Then
        IsRecordActive = False
        Exit Function
    End If
    
    ' Record is active if ColorIndex is NOT 15 (gray)
    ' Active records have ColorIndex = -4105 (xlAutomatic/black)
    ' Inactive records have ColorIndex = 15 (gray)
    IsRecordActive = (cell.Font.ColorIndex <> 15)
End Function

' Process a single record and create blank
' Returns True if successful
' errorMsg - output parameter with error description if processing fails
Private Function ProcessSingleRecord( _
    ByVal wsKinder As Worksheet, _
    ByVal wbTemplate As Workbook, _
    ByVal rowNum As Long, _
    ByVal targetFolder As String, _
    ByRef errorMsg As String _
) As Boolean
    
    On Error GoTo ErrorHandler
    
    Dim wsTemplate As Worksheet
    Dim lastName As String
    Dim firstName As String
    Dim discipline As String
    Dim teacherName As String
    Dim teacherFolder As String
    Dim dateFrom As Variant
    Dim dateTo As Variant
    Dim referenceDate As Variant
    Dim fileName As String
    Dim fullPath As String
    
    errorMsg = ""
    
    ' Get template worksheet
    Set wsTemplate = wbTemplate.Worksheets("Muster")
    
    ' Read data from Kinder worksheet
    lastName = Trim(wsKinder.Cells(rowNum, "C").value)
    firstName = Trim(wsKinder.Cells(rowNum, "D").value)
    discipline = Trim(wsKinder.Cells(rowNum, "E").value)
    teacherName = Trim(wsKinder.Cells(rowNum, "I").value)
    
    ' Read dates safely (may be empty or invalid)
    On Error Resume Next
    dateFrom = wsKinder.Cells(rowNum, "G").value
    dateTo = wsKinder.Cells(rowNum, "H").value
    referenceDate = wsKinder.Range("T2").value
    On Error GoTo ErrorHandler
    
    ' Check if required data is present
    If lastName = "" Then
        errorMsg = "Nachname (Spalte C) ist leer"
        ProcessSingleRecord = False
        Exit Function
    End If
    
    If firstName = "" Then
        errorMsg = "Vorname (Spalte D) ist leer"
        ProcessSingleRecord = False
        Exit Function
    End If
    
    If discipline = "" Then
        errorMsg = "Disziplin (Spalte E) ist leer"
        ProcessSingleRecord = False
        Exit Function
    End If
    
    ' Validate dates
    If Not IsDate(dateFrom) And Not IsEmpty(dateFrom) Then
        errorMsg = "Ungueltige Datum 'Von' (Spalte G): " & dateFrom
        ProcessSingleRecord = False
        Exit Function
    End If
    
    If Not IsDate(dateTo) And Not IsEmpty(dateTo) Then
        errorMsg = "Ungueltige Datum 'Bis' (Spalte H): " & dateTo
        ProcessSingleRecord = False
        Exit Function
    End If
    
    If Not IsDate(referenceDate) Then
        errorMsg = "Ungueltige Referenzdatum in Zelle T2: " & referenceDate
        ProcessSingleRecord = False
        Exit Function
    End If
    
    ' Fill template cells
    wsTemplate.Range("B1").value = lastName & " " & firstName
    wsTemplate.Range("E1").value = discipline
    wsTemplate.Range("B2").value = NHblank_Utils.FormatDateRange(dateFrom, dateTo)
    wsTemplate.Range("C4").value = NHblank_Utils.GetGermanMonth(Month(CDate(referenceDate)))
    wsTemplate.Range("E4").value = Year(CDate(referenceDate))
    
    ' Determine teacher subfolder
    If teacherName = "" Then
        teacherFolder = "Unsorted"
    Else
        teacherFolder = NHblank_Utils.CreateSafeFolderName(teacherName)
    End If
    
    ' Create teacher subfolder if it doesn't exist
    Dim targetSubfolder As String
    targetSubfolder = targetFolder & "\" & teacherFolder
    If Dir(targetSubfolder, vbDirectory) = "" Then
        MkDir targetSubfolder
    End If
    
    ' Generate file name
    fileName = NHblank_Utils.CreateFileName(lastName, firstName, discipline)
    fullPath = targetSubfolder & "\" & fileName & ".xlsx"
    
    ' Check if file already exists
    If Dir(fullPath) <> "" Then
        errorMsg = "Datei existiert bereits: " & fileName & ".xlsx"
        ProcessSingleRecord = False
        Exit Function
    End If
    
    ' Save template as new file
    wbTemplate.SaveAs fullPath, FileFormat:=xlOpenXMLWorkbook
    
    ProcessSingleRecord = True
    Exit Function
    
ErrorHandler:
    errorMsg = "Fehler #" & Err.Number & ": " & Err.Description
    ProcessSingleRecord = False
End Function

