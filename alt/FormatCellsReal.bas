Attribute VB_Name = "FormatCellsReal"
Public Sub ConvertAndFormatCellsOptimized()
    Dim ws As Worksheet
    Dim lastRow As Long
    Dim i As Long
    Dim rng As Range, rngAY As Range
    Dim decimalSeparator As String
    
    Application.Calculation = xlCalculationManual
    Application.ScreenUpdating = False

    
    ' Work on the active sheet
    Set ws = ActiveSheet
    
    ' Get the system decimal separator (either "," or ".")
    decimalSeparator = Application.International(xlDecimalSeparator)
    
    ' Find the last used row by checking column A (change if needed)
    lastRow = ws.Cells(ws.Rows.count, 1).End(xlUp).row
    If lastRow < 3 Then
        MsgBox "No data to process from row 3 onwards.", vbInformation
        Exit Sub
    End If
    
    ' --- 1) Convert columns U through AF individually ---
    '    (U=21, V=22, ... AF=32)
    For i = 21 To 32
        Set rng = ws.Range(ws.Cells(3, i), ws.Cells(lastRow, i))
        
        ' Remove extra spaces
        rng.Replace What:=" ", Replacement:="", LookAt:=xlPart
        
        ' Unify decimal separators so Excel can parse them correctly
        If decimalSeparator = "," Then
            ' Replace any '.' with ',' in case the data was typed with a dot
            rng.Replace What:=".", Replacement:=",", LookAt:=xlPart
        Else
            ' Replace any ',' with '.' if system uses dot
            rng.Replace What:=",", Replacement:=".", LookAt:=xlPart
        End If
        
        ' Now do TextToColumns on one column range
        rng.TextToColumns Destination:=rng.Cells(1, 1), DataType:=xlDelimited, _
                          TextQualifier:=xlDoubleQuote, ConsecutiveDelimiter:=False, _
                          Tab:=False, Semicolon:=False, Comma:=False, Space:=False, _
                          Other:=False
        
        ' Finally, set numeric format with two decimals, no leading zeros
        rng.NumberFormat = "0.00"
    Next i
    
    Set rng = ws.Range(ws.Cells(3, 3), ws.Cells(lastRow, 3))
    rng.TextToColumns Destination:=rng.Cells(1, 1), DataType:=xlDelimited, _
                  TextQualifier:=xlDoubleQuote, ConsecutiveDelimiter:=False, _
                  Tab:=False, Semicolon:=False, Comma:=False, Space:=False, _
                  Other:=False
    
    ' Set datum format
    rng.NumberFormat = "DD.MM.YYYY"
     
    Application.Calculation = xlCalculationAutomatic
    Application.ScreenUpdating = True

    
'    MsgBox "Conversion and formatting completed successfully!", vbInformation
End Sub


