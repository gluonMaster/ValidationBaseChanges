Attribute VB_Name = "gesh_file_operations"
Function FileExists(filePath As String) As Boolean
    ' Check if file exists at specified path
    FileExists = (Dir(filePath) <> "")
End Function

Function OpenOrGetWorkbook(filePath As String) As Workbook
    ' Open workbook or get reference if already open
    Dim wb As Workbook
    Dim fileName As String
    
    On Error GoTo ErrorHandler
    
    fileName = Mid(filePath, InStrRev(filePath, "\") + 1)
    
    ' Check if workbook is already open
    For Each wb In Application.Workbooks
        If wb.name = fileName Then
            ' Workbook is already open, save it first
            wb.Save
            Set OpenOrGetWorkbook = wb
            Exit Function
        End If
    Next wb
    
    ' Workbook is not open, open it
    Set OpenOrGetWorkbook = Application.Workbooks.Open(filePath)
    Exit Function
    
ErrorHandler:
    MsgBox "Fehler beim Oeffnen der Datei: " & Err.Description, vbCritical, "Dateifehler"
    Set OpenOrGetWorkbook = Nothing
End Function

Function ActivateSheet(wb As Workbook, sheetName As String) As Boolean
    ' Activate specified sheet in the workbook
    Dim ws As Worksheet
    
    On Error GoTo ErrorHandler
    
    Set ws = wb.Worksheets(sheetName)
    ws.Activate
    ActivateSheet = True
    Exit Function
    
ErrorHandler:
    MsgBox "Fehler beim Aktivieren des Blattes " & sheetName & ": " & Err.Description, _
           vbCritical, "Blattfehler"
    ActivateSheet = False
End Function

Function ExecuteMacro(wb As Workbook, macroName As String) As Boolean
    ' Execute specified macro in the workbook
    Dim fullMacroName As String
    
    On Error GoTo ErrorHandler
    
    fullMacroName = wb.name & "!" & macroName
    Application.Run fullMacroName
    ExecuteMacro = True
    Exit Function
    
ErrorHandler:
    MsgBox "Fehler beim Ausfuehren des Makros " & macroName & ": " & Err.Description, _
           vbCritical, "Makrofehler"
    ExecuteMacro = False
End Function

