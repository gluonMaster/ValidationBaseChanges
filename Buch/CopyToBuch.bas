Attribute VB_Name = "CopyToBuch"
' v. 1.1
' automated copying all files
' claer all filters and restore hiden rows and columns
' sorting by ID

Dim WrkBookAllDat As String
Dim WrkBookKindElt1 As String
Dim WrkBookKindElt2 As String
Dim s_Pfad As String

Sub CopyToBuch()
  Dim WrkBookAllDat As String
  Dim WrkBookKindElt As String
  Dim rngDatei As Range
  Dim rngKind As Range
  Dim i As Integer, j As Integer, Letzte As Integer, p1 As Integer
  Dim antwort As Integer
  
'  antwort = MsgBox(prompt:="Haben Sie Kopie von KindElternDatei_23 - 25 gemacht?", Buttons:=vbYesNo)
'  If antwort = vbNo Then
'    Exit Sub
'  End If
  
  s_Pfad = ThisWorkbook.Path

  
  WrkBookAllDat = "KindElternDaten2324.xlsm"
'  WrkBookKindElt1 = "KindElternDaten_23.xlsm"
'  WrkBookKindElt2 = "KindElternDaten_24.xlsm"
    Call CopySpecifiedFiles
  
    Call Ubertragen("2023", "KindElternDaten_23.xlsm")
    Call ImportFromBase
    'Call Ubertragen("2024", "KindElternDaten_24.xlsm")
    'Call Ubertragen("2025", "KindElternDaten_25.xlsm")
    
    MsgBox "The data is fully copied and sorted!", vbInformation
End Sub

Public Sub Ubertragen2023()
    ' Public entry point to import 2023 data from Excel file
    ' Called by ImportData.ImportFromBase_Extended
    ' Initializes s_Pfad and delegates to private Ubertragen
    
    s_Pfad = ThisWorkbook.Path
    Call Ubertragen("2023", "KindElternDaten_23.xlsm")
End Sub

Private Sub Ubertragen(jahrX As String, fileX As String)
    Dim i As Integer, j As Integer
    Dim Letzte As Integer
    Dim ws As Worksheet, wsTarget As Worksheet
    
    Workbooks("KindElternDaten2324.xlsm").Worksheets(jahrX).Activate
    Workbooks("KindElternDaten2324.xlsm").Worksheets(jahrX).Range("A3:AV5000").Clear
    
    Application.EnableEvents = False
    Application.DisplayAlerts = False
    Application.ScreenUpdating = False
    
    On Error GoTo Cleenup
    
    Workbooks.Open s_Pfad & "/" & fileX
    Set ws = Workbooks(fileX).Sheets("Kartei")
    Set wsTarget = Workbooks("KindElternDaten2324.xlsm").Sheets(jahrX)
    
    Call ResetSheetView(wsTarget)
    Call ResetSheetView(ws)
    
    Letzte = ws.Cells(ws.Rows.Count, "A").End(xlUp).row
    
'    For i = 3 To 5000
'      If Workbooks(fileX).Sheets("Kartei").Range("A" & i) = "" Then
'        Letzte = i - 1
'        Exit For
'      End If
'    Next i
    
    ws.Range("A3:AV" & Letzte).Copy
    wsTarget.Range("A3:A3").PasteSpecial Paste:=xlPasteAll
    
    Workbooks(fileX).Close SaveChanges:=False
    
    Letzte = wsTarget.Cells(wsTarget.Rows.Count, "A").End(xlUp).row
    
    Application.CutCopyMode = False
    
    With wsTarget.Sort
        .SortFields.Clear
        .SortFields.Add Key:=wsTarget.Range("B3:B" & Letzte), _
            SortOn:=xlSortOnValues, Order:=xlAscending, DataOption:=xlSortNormal
        .SortFields.Add Key:=wsTarget.Range("D3:D" & Letzte), _
            SortOn:=xlSortOnValues, Order:=xlAscending, DataOption:=xlSortNormal
        .SetRange wsTarget.Range("A3:AZ" & Letzte)
        .Header = xlNo
        .MatchCase = False
        .Orientation = xlTopToBottom
        .SortMethod = xlPinYin
        .Apply
    End With
    
    ' Ensure correct sheet is active before ConvertAndFormatCellsOptimized (uses ActiveSheet)
    wsTarget.Activate
    Call ConvertAndFormatCellsOptimized
    
Cleenup:
    Application.ScreenUpdating = True
    Application.DisplayAlerts = True
    Application.EnableEvents = True
    
'  Workbooks(WrkBookAllDat).Activate
End Sub

Public Sub ResetSheetView(ByVal ws As Worksheet)
    
    On Error Resume Next
        
    ' Verify, if there is filters applied on the sheet
    If ws.AutoFilterMode Then
        If ws.FilterMode Then
            ws.ShowAllData
        End If
    End If
    
    ' Depict all hidden rows
    ws.Rows.Hidden = False
    
    ' Depict all hidden columns
    ws.Columns.Hidden = False
    
    On Error GoTo 0
End Sub

Sub CopySpecifiedFiles()
    Dim sourcePath23, sourcePath24, sourcePath25 As String
    Dim sourcePaths, sourcePath As Variant
    Dim destPath As String
    Dim filesToCopy As Variant
    Dim fileName As Variant
    Dim fso As Object
    
    sourcePath23 = ThisWorkbook.Worksheets("2023").Range("I1").Value & "\"
'    sourcePath24 = "C:\Users\PolinaHerz\Kinder- und Elternzentrum Kolibri e.V\Alla Kremenchutski - Datenbank\2024\"
'    sourcePath25 = "C:\Users\PolinaHerz\Kinder- und Elternzentrum Kolibri e.V\Alla Kremenchutski - Datenbank\2025\"
    
    sourcePaths = Array(sourcePath23)

    destPath = ThisWorkbook.Path & "\"
    
    filesToCopy = Array("KindElternDaten_23.xlsm")
    
    ' Creaation of FileSystemObject
    Set fso = CreateObject("Scripting.FileSystemObject")
    
    ' Verification of sourcePath existance
    For Each sourcePath In sourcePaths
        If Not fso.FolderExists(sourcePath) Then
            MsgBox "SorcePath is absent: " & sourcePath, vbCritical
            Exit Sub
        End If
    Next sourcePath
    
    ' Verification of destination path existance
    If Not fso.FolderExists(destPath) Then
        MsgBox "destPath is absent: " & destPath, vbCritical
        Exit Sub
    End If
    
    Dim i As Integer
    i = 0
    For Each fileName In filesToCopy
        Dim fullSourceFile As String
        sourcePath = sourcePaths(i)
        fullSourceFile = sourcePath & fileName
        
        Dim fullDestFile As String
        fullDestFile = destPath & fileName
        
        ' Verification if the copied file exists
        If fso.FileExists(fullSourceFile) Then
            On Error Resume Next
            fso.CopyFile Source:=fullSourceFile, Destination:=fullDestFile, OverWriteFiles:=True
            If Err.Number <> 0 Then
                MsgBox "I can't copy the file: " & fileName & vbCrLf & "Whyle: " & Err.Description, vbExclamation
                Err.Clear
            End If
            On Error GoTo 0
        Else
            MsgBox "File not found: " & fullSourceFile, vbExclamation
        End If
        
        i = i + 1
    Next fileName
    
    MsgBox "File copying is complete.", vbInformation
End Sub
