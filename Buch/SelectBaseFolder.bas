Attribute VB_Name = "SelectBaseFolder"
Sub SelectFolder()
    Dim folderPath As String
    Dim fD As fileDialog
    
    Set fD = Application.fileDialog(msoFileDialogFolderPicker)
    
    With fD
        .Title = "Chouse the folder with relevant base"
        .InitialFileName = Application.DefaultFilePath ' initial folder
        If .Show = -1 Then ' if user press OK
            folderPath = .SelectedItems(1) ' extract folder path
            ThisWorkbook.ActiveSheet.Range("I1").Value = folderPath
        Else
            folderPath = "" ' if user press cansel path is empty
            ThisWorkbook.ActiveSheet.Range("I1").Value = folderPath
        End If
    End With
    
    ' clear object
    Set fD = Nothing
    
End Sub
