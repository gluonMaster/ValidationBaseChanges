Attribute VB_Name = "Notitzen"

' Simple global flag for user input completion
Public UserInputDone As Boolean
Public UserCanceled As Boolean

Public Sub CreateOrClearNotitzenSheet()
    Dim ws As Worksheet
    Dim sheetName As String
    sheetName = "Notitzen"
    
    ' Disable screen updating and automatic calculations for better performance
    Application.ScreenUpdating = False
    Application.Calculation = xlCalculationManual
    
    On Error Resume Next
    Set ws = ThisWorkbook.Worksheets(sheetName)
    On Error GoTo 0
    
    ' If the sheet does not exist, create it
    If ws Is Nothing Then
        Set ws = ThisWorkbook.Worksheets.Add(After:=ThisWorkbook.Worksheets(ThisWorkbook.Worksheets.count))
        ws.name = sheetName
    Else
        ' If the sheet exists, clear its content, formats, and fills
        ws.Cells.Clear
    End If
    
    ' Insert column headers starting from cell B1
    Dim headers As Variant
    headers = Array("ID", "Eltern", "Kind", "Jan", "Feb", "Mrz", "Apr", "Mai", "Jun", "Jul", "Aug", "Sep", "Okt", "Nov", "Dez")
    
    Dim i As Long
    For i = LBound(headers) To UBound(headers)
        With ws.Cells(1, i + 2) ' Start from column B (2)
            .value = headers(i)
            .Font.Bold = True
            .HorizontalAlignment = xlCenter
        End With
    Next i
    
    ' Insert "Notitzen:" into cell Q4
    With ws.Range("Q4")
        .value = "Notitzen:"
        .Font.Bold = True
        .HorizontalAlignment = xlLeft
    End With
    
    ' Set the width of column Q to 40 characters
    ws.Columns("Q").ColumnWidth = 40
    
    ' Set the height of row 5 to accommodate 5 standard rows
    ws.Rows(5).RowHeight = ws.StandardHeight * 2
    
    ' Format cell Q5 with highlighted background and black border
    With ws.Range("Q5")
        .Interior.Color = RGB(255, 255, 0) ' Yellow background for highlighting
        .Borders.LineStyle = xlContinuous
        .Borders.Color = vbBlack
        .Borders.Weight = xlThin
        .value = "" ' Clear any existing content
    End With
    
    ' Additional Requirements
    ' Insert "Initial:" into cell A2 with bold font and left alignment
    With ws.Range("A2")
        .value = "Initial:"
        .Font.Bold = True
        .HorizontalAlignment = xlLeft
    End With
    
    ' Insert "Corrected:" into cell A4 with bold font and left alignment
    With ws.Range("A4")
        .value = "Corrected:"
        .Font.Bold = True
        .HorizontalAlignment = xlLeft
    End With
    
    ' Set column widths
    ws.Columns("A").ColumnWidth = 11
    ws.Columns("B").ColumnWidth = 8
    ws.Columns("C").ColumnWidth = 20
    ws.Columns("D").ColumnWidth = 20
    ws.Columns("E:P").ColumnWidth = 6
    
    ' Apply number format to cells in columns E to O, rows 2 to 4
    With ws.Range("E2:P4")
        .NumberFormat = "0.00"
    End With
    
    ' Create buttons
    Call CreateButtons(ws)
    
    ' Add instruction text
    With ws.Range("Q6")
        .value = "Bitte geben Sie den Grund fuer die Aenderung ein und klicken Sie OK."
        .Font.Size = 10
        .Font.Italic = True
        .WrapText = True
    End With
    
    ws.Activate
    
    ' Re-enable automatic calculations and screen updating
    Application.Calculation = xlCalculationAutomatic
    Application.ScreenUpdating = True
End Sub

Private Sub CreateButtons(ws As Worksheet)
    Dim btn As Shape
    Dim shp As Shape
    
    ' Remove existing buttons
    For Each shp In ws.Shapes
        If shp.name = "btnOKNotitzen" Or shp.name = "btnCancelNotitzen" Then
            shp.Delete
        End If
    Next shp
    
    ' Create OK button
    Set btn = ws.Shapes.AddFormControl(xlButtonControl, 420, 100, 80, 30)
    btn.name = "btnOKNotitzen"
    btn.OnAction = "NotitzenOK"
    btn.TextFrame.Characters.text = "OK"
    
    ' Create Cancel button
    Set btn = ws.Shapes.AddFormControl(xlButtonControl, 420, 140, 80, 30)
    btn.name = "btnCancelNotitzen"
    btn.OnAction = "NotitzenCancel"
    btn.TextFrame.Characters.text = "Abbrechen"
End Sub

Public Sub NotitzenOK()
    Dim ws As Worksheet
    Set ws = ThisWorkbook.Worksheets("Notitzen")
    
    ' Check if comment was entered
    If Trim(ws.Range("Q5").value) = "" Then
        MsgBox "Bitte geben Sie einen Kommentar ein.", vbExclamation, "Kommentar erforderlich"
        ws.Range("Q5").Select
        Exit Sub
    End If
    
    ' Set completion flags
    UserInputDone = True
    UserCanceled = False
End Sub

Public Sub NotitzenCancel()
    Dim result As VbMsgBoxResult
    result = MsgBox("Moechten Sie den Vorgang wirklich abbrechen?", vbYesNo + vbQuestion, "Abbrechen bestaetigen")
    
    If result = vbYes Then
        UserInputDone = True
        UserCanceled = True
    End If
End Sub

Public Sub FillNotitzenSheet(changesArr As Variant, r As Long)
    ' This procedure fills the "Notitzen" sheet based on data from rows on
    ' "Kartei_Original" and "Kartei" sheets and highlights changed months.

    Dim wsNotitzen As Worksheet
    Dim wsKartei As Worksheet
    Dim wsKarteiOriginal As Worksheet
    Dim i As Long
    
    ' Reset flags
    UserInputDone = False
    UserCanceled = False
    
    ' Assign worksheets to variables
    Set wsNotitzen = ThisWorkbook.Sheets("Notitzen")
    Set wsKartei = ThisWorkbook.Sheets("Kartei")
    Set wsKarteiOriginal = ThisWorkbook.Sheets("Kartei_Original")
    
    ' Clear previous comment
    wsNotitzen.Range("Q5").value = ""
    wsNotitzen.Range("Q17").value = ""
    
    wsNotitzen.Range("B2").value = wsKarteiOriginal.Cells(r, 1).value
    wsNotitzen.Range("C2").value = wsKarteiOriginal.Cells(r, 2).value
    wsNotitzen.Range("D2").value = wsKarteiOriginal.Cells(r, 4).value
    wsNotitzen.Range("C3").value = wsKarteiOriginal.Cells(r, 10).value
    wsNotitzen.Range("D3").value = wsKarteiOriginal.Cells(r, 15).value
    ' Copy data from "Kartei_Original" (row r, columns 21-32) to "Notitzen" (row 2, columns E-P)
    wsNotitzen.Range("E2:P2").value = wsKarteiOriginal.Cells(r, 21).Resize(1, 12).value
    
    wsNotitzen.Range("B4").value = wsKartei.Cells(r, 1).value
    wsNotitzen.Range("C4").value = wsKartei.Cells(r, 2).value
    wsNotitzen.Range("D4").value = wsKartei.Cells(r, 4).value
    wsNotitzen.Range("C5").value = wsKartei.Cells(r, 10).value
    wsNotitzen.Range("D5").value = wsKartei.Cells(r, 15).value
    ' Copy data from "Kartei" (row r, columns 21-32) to "Notitzen" (row 4, columns E-P)
    wsNotitzen.Range("E4:P4").value = wsKartei.Cells(r, 21).Resize(1, 12).value
    
    ' Highlight changed months based on changesArr
    For i = 1 To 12
        If changesArr(i) = 1 Then
            ' Color the corresponding cells in rows 2 and 4 with pale pink
            wsNotitzen.Cells(4, 4 + i).Interior.Color = RGB(255, 192, 203)
        Else
            ' Remove any fill if there is no change
            wsNotitzen.Cells(2, 4 + i).Interior.ColorIndex = xlNone
            wsNotitzen.Cells(4, 4 + i).Interior.ColorIndex = xlNone
        End If
    Next i
    
    ' Activate the Notitzen sheet and select the comment cell
    wsNotitzen.Activate
    wsNotitzen.Range("Q5").Select
    
    ' Wait for user input - simple loop
    Do While Not UserInputDone
        DoEvents
    Loop
    
    ' Check if user canceled
    If UserCanceled Then
        Exit Sub
    End If
    
    ' Get the user input and process it like in original code
    Dim userInput As String
    userInput = Trim(wsNotitzen.Range("Q5").value)
    
    If userInput <> "" Then
        ' Remove line breaks and slashes as before
        userInput = Replace(userInput, vbCrLf, " ")
        userInput = Replace(userInput, vbCr, " ")
        userInput = Replace(userInput, vbLf, " ")
        userInput = Replace(userInput, "/", " ")
        
        ' Store cleaned input in Q17 (same as original)
        wsNotitzen.Cells(5, 17).value = userInput
    End If
End Sub
