Attribute VB_Name = "ExportBackupKartei"
' =============== Module: SaveKartei ===============
Option Explicit

Public Sub BackupKarteiSheet()
    Dim wb As Workbook
    Dim ws As Worksheet
    Dim saveFolderPath As String
    Dim currentWbName As String
    Dim baseName As String
    Dim fileCount As Long
    Dim newFileName As String
    
    ' Turn off updates for performance
    Application.ScreenUpdating = False
    Application.Calculation = xlCalculationManual

    ' Set references
    Set wb = ThisWorkbook
    Set ws = wb.Sheets("Kartei")

    ' Determine the Saves folder path relative to the current workbook
    saveFolderPath = wb.path & Application.PathSeparator & "Saves"

    ' Create the Saves folder if it does not exist
    If Dir(saveFolderPath, vbDirectory) = "" Then
        MkDir saveFolderPath
    End If

    ' Extract the workbook name without extension
    currentWbName = wb.name
    If InStrRev(currentWbName, ".") > 0 Then
        baseName = Left$(currentWbName, InStrRev(currentWbName, ".") - 1)
    Else
        baseName = currentWbName
    End If

    ' Count how many files are in the Saves folder
    fileCount = GetFileCountInFolder(saveFolderPath, "*.xlsx", baseName)

    If fileCount < 50 Then
        ' Determine the next available index
        newFileName = GetNextFileName(saveFolderPath, baseName)
    Else
        ' There are 50 or more files => Overwrite the oldest file
        newFileName = GetOldestFileForOverwrite(saveFolderPath, baseName)
    End If

    ' Actually save the "Kartei" sheet into the new file
    SaveSheetAsXlsx ws, saveFolderPath, newFileName
    
    ' Restore settings
    Application.Calculation = xlCalculationAutomatic
    Application.ScreenUpdating = True
End Sub

Public Sub SaveSheetAsXlsx(ByVal sourceSheet As Worksheet, _
                           ByVal folderPath As String, _
                           ByVal fileName As String)
    Dim tempWorkbook As Workbook
    Dim newFilePath As String
    
    newFilePath = folderPath & Application.PathSeparator & fileName
    
    ' Copy the single sheet to a new workbook
    sourceSheet.Copy
    Set tempWorkbook = ActiveWorkbook
    
    ' Save as XLSX
    Application.DisplayAlerts = False
    tempWorkbook.SaveAs fileName:=newFilePath, FileFormat:=xlOpenXMLWorkbook
    Application.DisplayAlerts = True

    ' Close the temp workbook
    tempWorkbook.Close SaveChanges:=False
End Sub
' =========== End of Module: MdlSaveKartei ============


' ============ Module: MdlUtilityFunctions ============

Public Function GetFileCountInFolder(ByVal folderPath As String, _
                                     ByVal filePattern As String, _
                                     ByVal baseName As String) As Long
    Dim fileName As String
    Dim countFiles As Long
    Dim fullFilePath As String
    
    countFiles = 0
    fileName = Dir(folderPath & Application.PathSeparator & filePattern)
    
    Do While fileName <> ""
        fullFilePath = folderPath & Application.PathSeparator & fileName
        
        ' Check if the file matches the naming convention: "BaseName_#.xlsx"
        If InStr(fileName, baseName & "_") = 1 Then
            countFiles = countFiles + 1
        End If
        
        fileName = Dir()
    Loop
    
    GetFileCountInFolder = countFiles
End Function

Public Function GetNextFileName(ByVal folderPath As String, _
                                ByVal baseName As String) As String
    Dim fileName As String
    Dim dirName As String
    Dim currentIndex As Long
    Dim maxIndex As Long
    Dim fullPath As String

    maxIndex = 0
    dirName = Dir(folderPath & Application.PathSeparator & "*.xlsx")
    
    Do While dirName <> ""
        If InStr(dirName, baseName & "_") = 1 Then
            ' Extract index from file name
            currentIndex = ExtractIndexFromFileName(dirName, baseName)
            If currentIndex > maxIndex Then
                maxIndex = currentIndex
            End If
        End If
        dirName = Dir()
    Loop

    ' Next index is maxIndex + 1
    GetNextFileName = baseName & "_" & (maxIndex + 1) & ".xlsx"
End Function

Public Function GetOldestFileForOverwrite(ByVal folderPath As String, _
                                          ByVal baseName As String) As String
    Dim fileName As String
    Dim oldestFileName As String
    Dim oldestDate As Date
    Dim currentFileDate As Date
    Dim firstFlag As Boolean
    
    firstFlag = True
    fileName = Dir(folderPath & Application.PathSeparator & "*.xlsx")
    
    Do While fileName <> ""
        If InStr(fileName, baseName & "_") = 1 Then
            ' Compare by creation time or last modified time
            currentFileDate = FileDateTime(folderPath & Application.PathSeparator & fileName)
            
            If firstFlag Then
                oldestDate = currentFileDate
                oldestFileName = fileName
                firstFlag = False
            Else
                If currentFileDate < oldestDate Then
                    oldestDate = currentFileDate
                    oldestFileName = fileName
                End If
            End If
        End If
        fileName = Dir()
    Loop
    
    GetOldestFileForOverwrite = oldestFileName
End Function

Private Function ExtractIndexFromFileName(ByVal fName As String, _
                                          ByVal baseName As String) As Long
    Dim underscorePos As Long
    Dim dotPos As Long
    Dim rawIndex As String
    
    ' We expect "baseName_#.xlsx"
    underscorePos = Len(baseName) + 2 ' position after "BaseName_"
    dotPos = InStrRev(fName, ".")
    
    If underscorePos > 0 And dotPos > underscorePos Then
        rawIndex = Mid$(fName, underscorePos, dotPos - underscorePos)
        If IsNumeric(rawIndex) Then
            ExtractIndexFromFileName = CLng(rawIndex)
        Else
            ExtractIndexFromFileName = 0
        End If
    Else
        ExtractIndexFromFileName = 0
    End If
End Function
' =========== End of Module: MdlUtilityFunctions ==========


