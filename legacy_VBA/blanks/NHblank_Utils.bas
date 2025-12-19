Attribute VB_Name = "NHblank_Utils"
Option Explicit

' Get German month name without umlauts
Public Function GetGermanMonth(ByVal monthNumber As Integer) As String
    Select Case monthNumber
        Case 1: GetGermanMonth = "Januar"
        Case 2: GetGermanMonth = "Februar"
        Case 3: GetGermanMonth = "Maerz"
        Case 4: GetGermanMonth = "April"
        Case 5: GetGermanMonth = "Mai"
        Case 6: GetGermanMonth = "Juni"
        Case 7: GetGermanMonth = "Juli"
        Case 8: GetGermanMonth = "August"
        Case 9: GetGermanMonth = "September"
        Case 10: GetGermanMonth = "Oktober"
        Case 11: GetGermanMonth = "November"
        Case 12: GetGermanMonth = "Dezember"
        Case Else: GetGermanMonth = ""
    End Select
End Function

' Format date range as "von DD.MM.YYYY bis DD.MM.YYYY"
Public Function FormatDateRange(ByVal dateFrom As Date, ByVal dateTo As Date) As String
    FormatDateRange = "Bewilligungszeitraum von " & Format(dateFrom, "DD.MM.YYYY") & " bis " & Format(dateTo, "DD.MM.YYYY")
End Function

' Create safe file name from last name, first name and discipline
' Replaces spaces with underscores and removes invalid characters
Public Function CreateFileName( _
    ByVal lastName As String, _
    ByVal firstName As String, _
    ByVal discipline As String _
) As String
    
    Dim fileName As String
    
    ' Build file name
    fileName = lastName & "_" & firstName & "_" & discipline
    
    ' Replace spaces with underscores
    fileName = Replace(fileName, " ", "_")
    
    ' Remove invalid file name characters
    fileName = RemoveInvalidChars(fileName)
    
    CreateFileName = fileName
End Function

' Create safe folder name from teacher name
' Replaces spaces with underscores and removes invalid characters
Public Function CreateSafeFolderName(ByVal teacherName As String) As String
    Dim folderName As String
    
    ' Replace spaces with underscores
    folderName = Replace(teacherName, " ", "_")
    
    ' Remove invalid folder name characters
    folderName = RemoveInvalidChars(folderName)
    
    CreateSafeFolderName = folderName
End Function

' Remove characters that are invalid in Windows file names
Private Function RemoveInvalidChars(ByVal fileName As String) As String
    Dim invalidChars As Variant
    Dim i As Long
    Dim result As String
    
    ' List of invalid characters in Windows file names
    invalidChars = Array("/", "\", ":", "*", "?", """", "<", ">", "|")
    
    result = fileName
    
    ' Remove each invalid character
    For i = LBound(invalidChars) To UBound(invalidChars)
        result = Replace(result, invalidChars(i), "")
    Next i
    
    RemoveInvalidChars = result
End Function

' Debug function to show path information
Public Sub ShowPathDebugInfo()
    Dim msg As String
    msg = "Debug Info:" & vbCrLf & vbCrLf
    msg = msg & "ThisWorkbook.Path: " & ThisWorkbook.Path & vbCrLf
    msg = msg & "ThisWorkbook.FullName: " & ThisWorkbook.fullName & vbCrLf
    msg = msg & "Application.ActiveWorkbook.Path: " & Application.ActiveWorkbook.Path & vbCrLf
    msg = msg & "CurDir: " & CurDir & vbCrLf
    msg = msg & "Environ(USERPROFILE): " & Environ("USERPROFILE") & vbCrLf
    
    ' Check if it's a OneDrive URL
    If InStr(1, ThisWorkbook.Path, "https://", vbTextCompare) > 0 Then
        msg = msg & vbCrLf & "DETECTED: OneDrive URL" & vbCrLf
        msg = msg & "GetLocalPath result: " & GetLocalPath(ThisWorkbook.Path) & vbCrLf
    End If
    
    MsgBox msg, vbInformation, "Path Debug Info"
End Sub

' Get local path from OneDrive or SharePoint URL
' Converts cloud paths to local synchronized folder paths
Public Function GetLocalPath(ByVal workbookPath As String) As String
    Dim localPath As String
    
    ' Check if this is a OneDrive or SharePoint URL
    If InStr(1, workbookPath, "https://", vbTextCompare) > 0 Then
        ' Method 1: Try to get local sync path from environment variables
        localPath = GetOneDriveLocalPath(workbookPath)
        If localPath <> "" And Dir(localPath & "\Shablon.xlsx") <> "" Then
            GetLocalPath = localPath
            Exit Function
        End If
        
        ' Method 2: Try to use current workbook's actual file location
        ' Sometimes the file is cached locally even when Path shows URL
        localPath = ExtractLocalPathFromFullName(ThisWorkbook.fullName)
        If localPath <> "" And Dir(localPath & "\Shablon.xlsx") <> "" Then
            GetLocalPath = localPath
            Exit Function
        End If
        
        ' Method 3: Try current directory
        localPath = CurDir
        If Dir(localPath & "\Shablon.xlsx") <> "" Then
            GetLocalPath = localPath
            Exit Function
        End If
        
        ' Method 4: Search in common OneDrive locations
        localPath = SearchInCommonLocations()
        If localPath <> "" Then
            GetLocalPath = localPath
            Exit Function
        End If
        
        ' If all methods fail, return empty string to trigger manual selection
        GetLocalPath = ""
    Else
        ' This is already a local path
        GetLocalPath = workbookPath
    End If
End Function

' Extract local path from FullName (for cached files)
Private Function ExtractLocalPathFromFullName(ByVal fullName As String) As String
    Dim localPath As String
    
    ' If FullName contains local path, extract directory
    If InStr(1, fullName, ":\", vbTextCompare) > 0 Then
        ' This looks like a local path
        Dim lastBackslash As Long
        lastBackslash = InStrRev(fullName, "\")
        If lastBackslash > 0 Then
            localPath = Left(fullName, lastBackslash - 1)
            ExtractLocalPathFromFullName = localPath
        End If
    Else
        ExtractLocalPathFromFullName = ""
    End If
End Function

' Search for Shablon.xlsx in common locations
Private Function SearchInCommonLocations() As String
    Dim userProfile As String
    Dim testPaths As Variant
    Dim i As Long
    
    userProfile = Environ("USERPROFILE")
    
    ' Define search locations
    testPaths = Array( _
        userProfile & "\Desktop", _
        userProfile & "\Documents", _
        userProfile & "\Downloads", _
        "C:\Temp", _
        "C:\Users\Public\Desktop", _
        userProfile & "\OneDrive\Desktop", _
        userProfile & "\OneDrive\Documents" _
    )
    
    ' Search in each location
    For i = LBound(testPaths) To UBound(testPaths)
        If Dir(testPaths(i) & "\Shablon.xlsx") <> "" Then
            SearchInCommonLocations = testPaths(i)
            Exit Function
        End If
    Next i
    
    SearchInCommonLocations = ""
End Function

' Convert OneDrive URL to local synchronized folder path
Private Function GetOneDriveLocalPath(ByVal oneDriveUrl As String) As String
    Dim userProfile As String
    Dim oneDrivePath As String
    Dim tempPath As String
    
    On Error GoTo ErrorHandler
    
    userProfile = Environ("USERPROFILE")
    
    ' Try different OneDrive folder patterns
    ' OneDrive Personal
    oneDrivePath = userProfile & "\OneDrive"
    If Dir(oneDrivePath, vbDirectory) <> "" Then
        If Dir(oneDrivePath & "\Shablon.xlsx") <> "" Then
            GetOneDriveLocalPath = oneDrivePath
            Exit Function
        End If
    End If
    
    ' OneDrive Business - try common patterns
    Dim possiblePaths As Variant
    possiblePaths = Array( _
        userProfile & "\OneDrive - Personal", _
        userProfile & "\OneDrive - Business", _
        userProfile & "\OneDrive for Business" _
    )
    
    Dim i As Long
    For i = LBound(possiblePaths) To UBound(possiblePaths)
        If Dir(possiblePaths(i), vbDirectory) <> "" Then
            If Dir(possiblePaths(i) & "\Shablon.xlsx") <> "" Then
                GetOneDriveLocalPath = possiblePaths(i)
                Exit Function
            End If
        End If
    Next i
    
    ' Try to find any OneDrive folder with Shablon.xlsx
    Dim fso As Object
    Set fso = CreateObject("Scripting.FileSystemObject")
    
    If fso.FolderExists(userProfile) Then
        Dim folder As Object
        Set folder = fso.GetFolder(userProfile)
        
        Dim subFolder As Object
        For Each subFolder In folder.SubFolders
            If InStr(1, subFolder.name, "OneDrive", vbTextCompare) > 0 Then
                If Dir(subFolder.Path & "\Shablon.xlsx") <> "" Then
                    GetOneDriveLocalPath = subFolder.Path
                    Exit Function
                End If
            End If
        Next subFolder
    End If
    
    ' Try current directory where the workbook is actually running from
    ' Sometimes Excel shows OneDrive URL but actually runs from local cache
    Dim currentDir As String
    currentDir = CurDir
    If Dir(currentDir & "\Shablon.xlsx") <> "" Then
        GetOneDriveLocalPath = currentDir
        Exit Function
    End If
    
ErrorHandler:
    GetOneDriveLocalPath = ""
End Function

