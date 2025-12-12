Attribute VB_Name = "ExportIDAdd"
Public Sub AddIDsToKartei()
    Dim wsKartei As Worksheet
    Dim wsTabelle8 As Worksheet
    Dim lastID As Long
    Dim currentID As Long
    Dim lastRow As Long
    Dim i As Long
    
    ' Disable screen updating and automatic calculations for better performance
    Application.ScreenUpdating = False
    Application.Calculation = xlCalculationManual
    
    ' Set references to the worksheets
    Set wsKartei = ThisWorkbook.Sheets("Kartei")
    Set wsTabelle8 = ThisWorkbook.Sheets("Tabelle8")
    
    ' Get the last used ID from cell B1 on the Tabelle8 sheet
    lastID = wsTabelle8.Range("B1").value
    currentID = lastID
    
    ' Determine the last used row in the range A:AY on the Kartei sheet
    On Error Resume Next
    lastRow = wsKartei.Cells(wsKartei.Rows.count, "A").End(xlUp).row
    On Error GoTo 0
    
    ' Check if there is data to process
    If lastRow < 3 Then
        MsgBox "No data to process."
        GoTo Cleanup
    End If
    
    ' Loop through rows from 3 to the last found row
    For i = 3 To lastRow
        ' Check if there are non-empty cells in the range A:AY of the current row
        If Application.WorksheetFunction.CountA(wsKartei.Range(wsKartei.Cells(i, "A"), wsKartei.Cells(i, "AY"))) > 0 Then
            ' Check if cell AV in the current row is empty
            If IsEmpty(wsKartei.Cells(i, "AV").value) Then
                ' Increment ID by 1
                currentID = currentID + 1
                ' Insert the new ID into cell AV
                wsKartei.Cells(i, "AV").value = currentID
                ' Update the last used ID on the Tabelle8 sheet
                wsTabelle8.Range("B1").value = currentID
            End If
        End If
    Next i
    
    ' Inform the user of successful completion
    'MsgBox "IDs have been successfully added and updated."
    
Cleanup:
    ' Restore Excel settings
    Application.Calculation = xlCalculationAutomatic
    Application.ScreenUpdating = True
End Sub


