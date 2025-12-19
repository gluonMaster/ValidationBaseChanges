Attribute VB_Name = "indBlank_TemplateWriter"
Option Explicit

' =============================================================================
' Module: indBlank_TemplateWriter
' Purpose: Handles template opening, copying, filling, and saving for
'          individual blank (Einzelmeldung) generation
' =============================================================================

' -----------------------------------------------------------------------------
' Module-level variables for template workbook (opened once)
' -----------------------------------------------------------------------------
Private m_TemplateWb As Workbook
Private m_TemplateWs As Worksheet
Private m_IsTemplateOpen As Boolean

' =============================================================================
' Function: indBlank_OpenTemplate
' Purpose: Opens the template workbook once (ReadOnly) and validates Muster sheet
' Param:   errMsg - output error message if failed
' Returns: True if template opened successfully, False otherwise
' =============================================================================
Public Function indBlank_OpenTemplate(ByRef errMsg As String) As Boolean
    On Error GoTo ErrorHandler
    
    errMsg = ""
    indBlank_OpenTemplate = False
    
    ' Check if already open
    If m_IsTemplateOpen Then
        indBlank_OpenTemplate = True
        Exit Function
    End If
    
    ' Check if template file exists
    If Dir(indBlank_Config.TEMPLATE_PATH) = "" Then
        errMsg = "Vorlage nicht gefunden: " & indBlank_Config.TEMPLATE_PATH
        Exit Function
    End If
    
    ' Open template workbook as read-only
    Set m_TemplateWb = Workbooks.Open( _
        Filename:=indBlank_Config.TEMPLATE_PATH, _
        ReadOnly:=True, _
        UpdateLinks:=False)
    
    ' Try to get the Muster sheet
    If Not indBlank_TryGetWorksheet(m_TemplateWb, indBlank_Config.TEMPLATE_SHEET, m_TemplateWs) Then
        errMsg = "Blatt '" & indBlank_Config.TEMPLATE_SHEET & "' nicht in Vorlage gefunden."
        m_TemplateWb.Close SaveChanges:=False
        Set m_TemplateWb = Nothing
        Exit Function
    End If
    
    m_IsTemplateOpen = True
    indBlank_OpenTemplate = True
    Exit Function
    
ErrorHandler:
    errMsg = "Fehler beim Oeffnen der Vorlage: " & Err.Description
    On Error Resume Next
    If Not m_TemplateWb Is Nothing Then
        m_TemplateWb.Close SaveChanges:=False
    End If
    Set m_TemplateWb = Nothing
    Set m_TemplateWs = Nothing
    m_IsTemplateOpen = False
End Function

' =============================================================================
' Function: indBlank_CloseTemplate
' Purpose: Closes the template workbook
' =============================================================================
Public Sub indBlank_CloseTemplate()
    On Error Resume Next
    
    If m_IsTemplateOpen And Not m_TemplateWb Is Nothing Then
        m_TemplateWb.Close SaveChanges:=False
    End If
    
    Set m_TemplateWs = Nothing
    Set m_TemplateWb = Nothing
    m_IsTemplateOpen = False
    
    On Error GoTo 0
End Sub

' =============================================================================
' Function: indBlank_WriteBlank
' Purpose: Creates a single blank document for a child
' Param:   targetFolder - base folder path (with trailing backslash)
' Param:   childName - raw child name from Kartei column D
' Param:   discipline - parsed discipline name (from indBlank_TryParseDiscipline)
' Param:   teacherName - teacher name from Kartei column K/P
' Param:   inputYear - user-entered year for E4
' Param:   errMsg - output error message if failed
' Returns: True if blank created successfully, False otherwise
' =============================================================================
Public Function indBlank_WriteBlank( _
    ByVal targetFolder As String, _
    ByVal childName As String, _
    ByVal discipline As String, _
    ByVal teacherName As String, _
    ByVal inputYear As Long, _
    ByVal inputMonth As Long, _
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
    indBlank_WriteBlank = False
    
    ' Validate template is open
    If Not m_IsTemplateOpen Then
        errMsg = "Vorlage ist nicht geoeffnet."
        Exit Function
    End If
    
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
    ' This creates a new workbook with the copied sheet
    m_TemplateWs.Copy
    Set newWb = ActiveWorkbook
    Set newWs = newWb.Sheets(1)
    
    ' Fill template cells
    ' B1 = full child name
    newWs.Range("B1").Value = formattedChildName
    
    ' E1 = discipline
    newWs.Range("E1").Value = discipline
    
    ' C4 = German month name (from user-selected month)
    newWs.Range("C4").Value = indBlank_GetGermanMonth(inputMonth)
    
    ' E4 = input year
    newWs.Range("E4").Value = inputYear
    
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
    
    indBlank_WriteBlank = True
    Exit Function
    
ErrorHandler:
    errMsg = "Fehler beim Erstellen: " & Err.Description
    
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
' Param:   folderPath - folder path (with trailing backslash)
' Param:   baseName - base file name (without extension)
' Param:   extension - file extension (including dot)
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

' =============================================================================
' Function: indBlank_IsTemplateOpen
' Purpose: Returns whether template is currently open
' =============================================================================
Public Function indBlank_IsTemplateOpen() As Boolean
    indBlank_IsTemplateOpen = m_IsTemplateOpen
End Function
