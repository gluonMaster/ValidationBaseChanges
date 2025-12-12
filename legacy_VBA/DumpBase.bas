Attribute VB_Name = "DumpBase"
Option Explicit

' This function returns the source path for the file "KindElternDaten_25_front.accdb"
Public Function GetSourceFilePath() As String
    ' Modify this path according to your actual source location
    GetSourceFilePath = ThisWorkbook.Worksheets("Kartei").Range("X1").value & "\Alarm\KindElternDaten_25_front.accdb"
End Function

' This function returns the folder where the file should be copied
Public Function GetDestinationFolder() As String
    ' Modify this path according to your desired destination folder
    GetDestinationFolder = "C:\KindEltern\2025" & "\Saves\Dump\"
End Function

' This function finds the next available suffix for the file "KindElternDaten_24_front_<suffix>.accdb"
Public Function GetNextFileSuffix() As Long
    Dim fName As String
    Dim prefix As String, extension As String
    Dim maxNumber As Long
    Dim numberPart As String
    
    prefix = "KindElternDaten_25_front_"
    extension = ".accdb"
    
    ' Use Dir to retrieve files matching pattern in the destination folder
    fName = Dir(GetDestinationFolder() & prefix & "*" & extension)
    
    ' Loop through all files that match the pattern
    Do While fName <> ""
        ' Extract the suffix from the file name (the numeric part)
        Dim prefixLen As Long, extensionLen As Long
        
        prefixLen = Len(prefix)
        extensionLen = Len(extension)
        
        ' If the file name length is at least enough to contain prefix, suffix, and extension
        If Len(fName) > (prefixLen + extensionLen) Then
            numberPart = Mid(fName, prefixLen + 1, Len(fName) - prefixLen - extensionLen)
            
            ' Check if numberPart is numeric and, if so, compare with maxNumber
            If IsNumeric(numberPart) Then
                If CLng(numberPart) > maxNumber Then
                    maxNumber = CLng(numberPart)
                End If
            End If
        End If
        
        ' Get next matching file
        fName = Dir
    Loop
    
    ' The next available suffix is maxNumber + 1
    GetNextFileSuffix = maxNumber + 1
End Function

' This procedure performs the file copy operation with an incremental suffix
Public Sub CopyKindElternDatenFile()
    Dim fso As Object
    Dim sourcePath As String
    Dim destinationFolder As String
    Dim destinationFileName As String
    Dim nextSuffix As Long
    
    sourcePath = GetSourceFilePath()
    destinationFolder = GetDestinationFolder()
    
    ' Generate next suffix for the file
    nextSuffix = GetNextFileSuffix()
    
    ' Construct the new file name with suffix
    destinationFileName = destinationFolder & "KindElternDaten_25_front_" & nextSuffix & ".accdb"
    
    ' Create a FileSystemObject for file operations
    Set fso = CreateObject("Scripting.FileSystemObject")
    
    ' Copy the file without overwriting existing files
    fso.CopyFile sourcePath, destinationFileName, False
    
    ' Free the FileSystemObject reference
    Set fso = Nothing
    
    'MsgBox "File has been successfully copied to: " & destinationFileName, vbInformation
End Sub

