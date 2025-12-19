Attribute VB_Name = "indBlank_DataProcessor"
Option Explicit

' =============================================================================
' Module: indBlank_DataProcessor
' Purpose: Iterates through visible Kartei rows and generates individual blanks
' =============================================================================

' =============================================================================
' Function: indBlank_ProcessVisibleKarteiRows
' Purpose: Processes all visible rows in Kartei sheet and generates blanks
' Param:   wsKartei - Kartei worksheet to read data from
' Param:   wbTemplate - Template workbook (already opened)
' Param:   wsMuster - Muster worksheet from template
' Param:   targetFolder - Base folder path for output (with trailing backslash)
' Param:   yearValue - User-entered year for template
' Param:   outCreated - Output count of successfully created blanks
' Param:   outSkipped - Output count of skipped rows (not Ind/VSpE)
' Param:   outErrors - Output count of rows with errors
' Returns: True if processing completed (even with row errors), False for fatal errors
' =============================================================================
Public Function indBlank_ProcessVisibleKarteiRows( _
    ByVal wsKartei As Worksheet, _
    ByVal wbTemplate As Workbook, _
    ByVal wsMuster As Worksheet, _
    ByVal targetFolder As String, _
    ByVal yearValue As Long, _
    ByVal monthValue As Long, _
    ByRef outCreated As Long, _
    ByRef outSkipped As Long, _
    ByRef outErrors As Long) As Boolean
    
    Dim disciplineCol As String
    Dim teacherCol As String
    Dim lastRow As Long
    Dim r As Long
    Dim rawDiscipline As String
    Dim parsedDiscipline As String
    Dim childName As String
    Dim teacherName As String
    Dim errMsg As String
    Dim baseFolder As String
    
    On Error GoTo FatalErrorHandler
    
    ' Initialize counters
    outCreated = 0
    outSkipped = 0
    outErrors = 0
    indBlank_ProcessVisibleKarteiRows = False
    
    ' Validate parameters
    If wsKartei Is Nothing Then
        Debug.Print "indBlank_DataProcessor: wsKartei is Nothing"
        Exit Function
    End If
    
    If wbTemplate Is Nothing Then
        Debug.Print "indBlank_DataProcessor: wbTemplate is Nothing"
        Exit Function
    End If
    
    If wsMuster Is Nothing Then
        Debug.Print "indBlank_DataProcessor: wsMuster is Nothing"
        Exit Function
    End If
    
    If Len(targetFolder) = 0 Then
        Debug.Print "indBlank_DataProcessor: targetFolder is empty"
        Exit Function
    End If
    
    baseFolder = targetFolder
    If Right$(baseFolder, 1) <> "\" Then
        baseFolder = baseFolder & "\"
    End If
    
    ' Get semester-based columns for discipline and teacher
    indBlank_GetSemesterColumns disciplineCol, teacherCol, monthValue
    
    ' Find last row with data in column A
    lastRow = wsKartei.Cells(wsKartei.Rows.Count, "A").End(xlUp).Row
    
    ' Check if there's any data
    If lastRow < indBlank_Config.KARTEI_FIRST_DATA_ROW Then
        ' No data rows - not a fatal error, just nothing to process
        indBlank_ProcessVisibleKarteiRows = True
        Exit Function
    End If
    
    ' Process each row
    For r = indBlank_Config.KARTEI_FIRST_DATA_ROW To lastRow
        
        ' Skip hidden rows (covers both filters and manual hiding)
        If wsKartei.Rows(r).Hidden Then
            GoTo NextRow
        End If
        
        ' Read discipline value from current semester column
        rawDiscipline = CStr(wsKartei.Range(disciplineCol & r).Value)
        
        ' Try to parse discipline - skip if not Ind/VSpE
        If Not indBlank_TryParseDiscipline(rawDiscipline, parsedDiscipline) Then
            outSkipped = outSkipped + 1
            GoTo NextRow
        End If
        
        ' Read child name and teacher name
        childName = CStr(wsKartei.Range(indBlank_Config.KARTEI_COL_CHILD & r).Value)
        teacherName = CStr(wsKartei.Range(teacherCol & r).Value)
        
        ' Validate child name is not empty
        If Len(Trim$(childName)) = 0 Then
            outSkipped = outSkipped + 1
            GoTo NextRow
        End If
        
        ' Create blank for this row
        On Error GoTo RowErrorHandler
        
        If CreateSingleBlank(wsMuster, baseFolder, childName, parsedDiscipline, teacherName, yearValue, monthValue, errMsg) Then
            outCreated = outCreated + 1
        Else
            Debug.Print "Row " & r & " error: " & errMsg
            outErrors = outErrors + 1
        End If
        
        On Error GoTo FatalErrorHandler
        
NextRow:
    Next r
    
    indBlank_ProcessVisibleKarteiRows = True
    Exit Function
    
RowErrorHandler:
    ' Handle row-level error - don't stop entire process
    Debug.Print "Row " & r & " exception: " & Err.Description
    outErrors = outErrors + 1
    Resume NextRow
    
FatalErrorHandler:
    ' Fatal error - stop processing
    Debug.Print "Fatal error in indBlank_ProcessVisibleKarteiRows: " & Err.Description
    indBlank_ProcessVisibleKarteiRows = False
End Function

' =============================================================================
' Private Helper: CreateSingleBlank
' Purpose: Creates a single blank document from the Muster template
' Param:   wsMuster - Muster worksheet to copy
' Param:   targetFolder - Base folder path (with trailing backslash)
' Param:   childName - Raw child name from Kartei
' Param:   discipline - Parsed discipline name
' Param:   teacherName - Teacher name for subfolder
' Param:   yearValue - Year for template
' Param:   errMsg - Output error message if failed
' Returns: True if successful, False otherwise
' =============================================================================
Private Function CreateSingleBlank( _
    ByVal wsMuster As Worksheet, _
    ByVal targetFolder As String, _
    ByVal childName As String, _
    ByVal discipline As String, _
    ByVal teacherName As String, _
    ByVal yearValue As Long, _
    ByVal monthValue As Long, _
    ByRef errMsg As String) As Boolean
    
    Dim newWb As Workbook
    Dim newWs As Worksheet
    Dim safeTeacher As String
    Dim safeChild As String
    Dim safeDiscipline As String
    Dim subFolderPath As String
    Dim fileName As String
    Dim fullPath As String
    Dim formattedChildName As String
    
    On Error GoTo ErrorHandler
    
    errMsg = ""
    CreateSingleBlank = False
    
    ' Format child name for display in template
    formattedChildName = indBlank_FormatChildFullName(childName)
    
    ' Create safe names for folder and file
    safeTeacher = Trim$(teacherName)
    If Len(safeTeacher) = 0 Then
        safeTeacher = "Unsorted"
    Else
        safeTeacher = indBlank_CreateSafeFolderName(safeTeacher)
    End If
    
    safeChild = indBlank_CreateSafeFileName(formattedChildName)
    safeDiscipline = indBlank_CreateSafeFileName(discipline)
    
    ' Build subfolder path
    subFolderPath = targetFolder & safeTeacher & "\"
    
    ' Create subfolder if it doesn't exist
    If Dir(subFolderPath, vbDirectory) = "" Then
        MkDir subFolderPath
    End If
    
    ' Build base file name
    fileName = safeChild & "_" & safeDiscipline
    
    ' Get unique file path (add suffix if exists)
    fullPath = GetUniqueFilePath(subFolderPath, fileName, ".xlsx")
    
    ' Copy Muster sheet to new workbook
    wsMuster.Copy
    Set newWb = ActiveWorkbook
    Set newWs = newWb.Sheets(1)
    
    ' Fill template cells
    ' B1 = full child name
    newWs.Range("B1").Value = formattedChildName
    
    ' E1 = discipline
    newWs.Range("E1").Value = discipline
    
    ' C4 = German month name (from user-selected month)
    newWs.Range("C4").Value = indBlank_GetGermanMonth(monthValue)
    
    ' E4 = input year
    newWs.Range("E4").Value = yearValue
    
    ' B2 = do NOT fill (as per requirements)
    
    ' Save as xlsx
    Application.DisplayAlerts = False
    newWb.SaveAs _
        Filename:=fullPath, _
        FileFormat:=xlOpenXMLWorkbook
    Application.DisplayAlerts = True
    
    ' Close the new workbook
    newWb.Close SaveChanges:=False
    Set newWs = Nothing
    Set newWb = Nothing
    
    CreateSingleBlank = True
    Exit Function
    
ErrorHandler:
    errMsg = Err.Description
    
    On Error Resume Next
    Application.DisplayAlerts = True
    If Not newWb Is Nothing Then
        newWb.Close SaveChanges:=False
    End If
    Set newWs = Nothing
    Set newWb = Nothing
    On Error GoTo 0
End Function

' =============================================================================
' Private Helper: GetUniqueFilePath
' Purpose: Returns a unique file path, adding _2, _3, etc. if file exists
' Param:   folderPath - Folder path (with trailing backslash)
' Param:   baseName - Base file name (without extension)
' Param:   extension - File extension (including dot)
' Returns: Unique full file path
' =============================================================================
Private Function GetUniqueFilePath(ByVal folderPath As String, ByVal baseName As String, ByVal extension As String) As String
    Dim fullPath As String
    Dim counter As Long
    
    ' Try base name first
    fullPath = folderPath & baseName & extension
    
    If Dir(fullPath) = "" Then
        GetUniqueFilePath = fullPath
        Exit Function
    End If
    
    ' Add suffix until unique
    counter = 2
    Do
        fullPath = folderPath & baseName & "_" & counter & extension
        If Dir(fullPath) = "" Then
            GetUniqueFilePath = fullPath
            Exit Function
        End If
        counter = counter + 1
        
        ' Safety limit
        If counter > 1000 Then
            ' Fallback with timestamp
            fullPath = folderPath & baseName & "_" & Format$(Now, "hhnnss") & extension
            GetUniqueFilePath = fullPath
            Exit Function
        End If
    Loop
End Function
