Attribute VB_Name = "zahl_FileGenerator"
' File generation module for family payments
' Handles creation of Excel output files with proper formatting

Option Explicit

Public Sub GenerateAllReports()
    ' Generate all three report files
    Call GenerateReport("EltKosten.xlsx", "regular")
    Call GenerateReport("EltKostenNH.xlsx", "nachhilfe")
    Call GenerateReport("EltKostenInd.xlsx", "individual")
End Sub

Private Sub GenerateReport(filename As String, reportType As String)
    ' Generate a specific report file
    Dim wb As Workbook
    Dim i As Integer, outputRow As Integer
    Dim hasData As Boolean
    
    ' Create new workbook
    Set wb = Workbooks.Add
    
    ' Save workbook with specified name
    On Error GoTo SaveError
    wb.SaveAs filename:="C:\FamZahlung\" & filename, FileFormat:=xlOpenXMLWorkbook, CreateBackup:=False
    On Error GoTo 0
    
    ' Add headers
    Call AddHeaders(wb.Sheets(1))
    
    outputRow = 2 ' Start from row 2 (row 1 has headers)
    hasData = False
    
    ' Process all families
    For i = 1 To 2000
        If ShouldIncludeFamily(i, reportType) Then
            Call AddFamilyToReport(wb.Sheets(1), i, outputRow, reportType)
            outputRow = outputRow + 1
            hasData = True
        End If
    Next i
    
    If hasData Then
        ' Apply formatting and borders
        Call ApplyReportFormatting(wb.Sheets(1), outputRow - 1, reportType)
    Else
        ' Add a message if no data found
        wb.Sheets(1).Cells(2, 1) = "Keine Daten gefunden"
    End If
    
    ' Save and close workbook
    wb.Save
    wb.Close
    Exit Sub
    
SaveError:
    MsgBox "Fehler beim Speichern der Datei: " & filename & vbCrLf & Err.Description, vbCritical, "Speicherfehler"
    On Error GoTo 0
End Sub

Private Function ShouldIncludeFamily(familyIndex As Integer, reportType As String) As Boolean
    ' Determine if family should be included in specific report
    Select Case LCase(reportType)
        Case "regular"
            ShouldIncludeFamily = zahl_Main.nichtNull(familyIndex)
        Case "nachhilfe"
            ShouldIncludeFamily = zahl_Main.nichtNullN(familyIndex)
        Case "individual"
            ShouldIncludeFamily = zahl_Main.nichtNullI(familyIndex)
        Case Else
            ShouldIncludeFamily = False
    End Select
End Function

Private Sub AddHeaders(ws As Worksheet)
    ' Add column headers to worksheet
    ws.Cells(1, 1) = "Nummer"
    ws.Cells(1, 2) = "Name"
    ws.Cells(1, 3) = "Januar"
    ws.Cells(1, 4) = "Februar"
    ws.Cells(1, 5) = "Maerz"
    ws.Cells(1, 6) = "April"
    ws.Cells(1, 7) = "Mai"
    ws.Cells(1, 8) = "Juni"
    ws.Cells(1, 9) = "Juli"
    ws.Cells(1, 10) = "August"
    ws.Cells(1, 11) = "September"
    ws.Cells(1, 12) = "Oktober"
    ws.Cells(1, 13) = "November"
    ws.Cells(1, 14) = "Dezember"
End Sub

Private Sub AddFamilyToReport(ws As Worksheet, familyIndex As Integer, outputRow As Integer, reportType As String)
    ' Add single family data to report with proper formatting
    Dim monthValues(1 To 12) As Double
    Dim colorArray(1 To 12) As Integer
    Dim i As Integer
    
    ' Set number format for amount columns
    ws.Range("C" & outputRow & ":N" & outputRow).NumberFormat = "0.00"
    
    ' Add family number and name safely
    ws.Cells(outputRow, 1).Value = zahl_Main.num(familyIndex)
    ws.Cells(outputRow, 2).Value = zahl_Main.nam(familyIndex)
    
    ' Get values and colors based on report type
    Call GetFamilyValues(familyIndex, reportType, monthValues(), colorArray())
    
    ' Add monthly values with color formatting
    For i = 1 To 12
        ws.Cells(outputRow, i + 2).Value = monthValues(i)
        
        ' Apply color formatting only for conflicts (yellow)
        Select Case colorArray(i)
            Case 2 ' Yellow for conflicts
                ws.Cells(outputRow, i + 2).Interior.Color = RGB(255, 255, 0) ' Bright yellow
            ' Case 1 - Pink for mixed cases - REMOVED per requirements
        End Select
    Next i
End Sub

Private Sub GetFamilyValues(familyIndex As Integer, reportType As String, ByRef monthValues() As Double, ByRef colorArray() As Integer)
    ' Get monthly values and color information for specific report type
    Select Case LCase(reportType)
        Case "regular"
            monthValues(1) = zahl_Main.sum01(familyIndex): colorArray(1) = zahl_Main.colorRegular(familyIndex, 1)
            monthValues(2) = zahl_Main.sum02(familyIndex): colorArray(2) = zahl_Main.colorRegular(familyIndex, 2)
            monthValues(3) = zahl_Main.sum03(familyIndex): colorArray(3) = zahl_Main.colorRegular(familyIndex, 3)
            monthValues(4) = zahl_Main.sum04(familyIndex): colorArray(4) = zahl_Main.colorRegular(familyIndex, 4)
            monthValues(5) = zahl_Main.sum05(familyIndex): colorArray(5) = zahl_Main.colorRegular(familyIndex, 5)
            monthValues(6) = zahl_Main.sum06(familyIndex): colorArray(6) = zahl_Main.colorRegular(familyIndex, 6)
            monthValues(7) = zahl_Main.sum07(familyIndex): colorArray(7) = zahl_Main.colorRegular(familyIndex, 7)
            monthValues(8) = zahl_Main.sum08(familyIndex): colorArray(8) = zahl_Main.colorRegular(familyIndex, 8)
            monthValues(9) = zahl_Main.sum09(familyIndex): colorArray(9) = zahl_Main.colorRegular(familyIndex, 9)
            monthValues(10) = zahl_Main.sum10(familyIndex): colorArray(10) = zahl_Main.colorRegular(familyIndex, 10)
            monthValues(11) = zahl_Main.sum11(familyIndex): colorArray(11) = zahl_Main.colorRegular(familyIndex, 11)
            monthValues(12) = zahl_Main.sum12(familyIndex): colorArray(12) = zahl_Main.colorRegular(familyIndex, 12)
            
        Case "nachhilfe"
            monthValues(1) = zahl_Main.sum01N(familyIndex): colorArray(1) = zahl_Main.colorNachhilfe(familyIndex, 1)
            monthValues(2) = zahl_Main.sum02N(familyIndex): colorArray(2) = zahl_Main.colorNachhilfe(familyIndex, 2)
            monthValues(3) = zahl_Main.sum03N(familyIndex): colorArray(3) = zahl_Main.colorNachhilfe(familyIndex, 3)
            monthValues(4) = zahl_Main.sum04N(familyIndex): colorArray(4) = zahl_Main.colorNachhilfe(familyIndex, 4)
            monthValues(5) = zahl_Main.sum05N(familyIndex): colorArray(5) = zahl_Main.colorNachhilfe(familyIndex, 5)
            monthValues(6) = zahl_Main.sum06N(familyIndex): colorArray(6) = zahl_Main.colorNachhilfe(familyIndex, 6)
            monthValues(7) = zahl_Main.sum07N(familyIndex): colorArray(7) = zahl_Main.colorNachhilfe(familyIndex, 7)
            monthValues(8) = zahl_Main.sum08N(familyIndex): colorArray(8) = zahl_Main.colorNachhilfe(familyIndex, 8)
            monthValues(9) = zahl_Main.sum09N(familyIndex): colorArray(9) = zahl_Main.colorNachhilfe(familyIndex, 9)
            monthValues(10) = zahl_Main.sum10N(familyIndex): colorArray(10) = zahl_Main.colorNachhilfe(familyIndex, 10)
            monthValues(11) = zahl_Main.sum11N(familyIndex): colorArray(11) = zahl_Main.colorNachhilfe(familyIndex, 11)
            monthValues(12) = zahl_Main.sum12N(familyIndex): colorArray(12) = zahl_Main.colorNachhilfe(familyIndex, 12)
            
        Case "individual"
            monthValues(1) = zahl_Main.sum01I(familyIndex): colorArray(1) = zahl_Main.colorIndividual(familyIndex, 1)
            monthValues(2) = zahl_Main.sum02I(familyIndex): colorArray(2) = zahl_Main.colorIndividual(familyIndex, 2)
            monthValues(3) = zahl_Main.sum03I(familyIndex): colorArray(3) = zahl_Main.colorIndividual(familyIndex, 3)
            monthValues(4) = zahl_Main.sum04I(familyIndex): colorArray(4) = zahl_Main.colorIndividual(familyIndex, 4)
            monthValues(5) = zahl_Main.sum05I(familyIndex): colorArray(5) = zahl_Main.colorIndividual(familyIndex, 5)
            monthValues(6) = zahl_Main.sum06I(familyIndex): colorArray(6) = zahl_Main.colorIndividual(familyIndex, 6)
            monthValues(7) = zahl_Main.sum07I(familyIndex): colorArray(7) = zahl_Main.colorIndividual(familyIndex, 7)
            monthValues(8) = zahl_Main.sum08I(familyIndex): colorArray(8) = zahl_Main.colorIndividual(familyIndex, 8)
            monthValues(9) = zahl_Main.sum09I(familyIndex): colorArray(9) = zahl_Main.colorIndividual(familyIndex, 9)
            monthValues(10) = zahl_Main.sum10I(familyIndex): colorArray(10) = zahl_Main.colorIndividual(familyIndex, 10)
            monthValues(11) = zahl_Main.sum11I(familyIndex): colorArray(11) = zahl_Main.colorIndividual(familyIndex, 11)
            monthValues(12) = zahl_Main.sum12I(familyIndex): colorArray(12) = zahl_Main.colorIndividual(familyIndex, 12)
    End Select
End Sub

Private Sub ApplyReportFormatting(ws As Worksheet, lastRow As Integer, reportType As String)
    ' Apply borders and table formatting to the report
    Dim tableRange As String
    
    If lastRow < 2 Then Exit Sub ' No data to format
    
    tableRange = "A1:N" & lastRow
    
    ' Apply borders
    Call ApplyBorders(ws, tableRange)
    
    ' Create and format table
    Call CreateFormattedTable(ws, tableRange)
End Sub

Private Sub ApplyBorders(ws As Worksheet, rangeAddress As String)
    ' Apply borders to specified range
    With ws.Range(rangeAddress)
        .Borders(xlDiagonalDown).LineStyle = xlNone
        .Borders(xlDiagonalUp).LineStyle = xlNone
        
        With .Borders(xlEdgeLeft)
            .LineStyle = xlContinuous
            .ColorIndex = 0
            .TintAndShade = 0
            .Weight = xlThin
        End With
        
        With .Borders(xlEdgeTop)
            .LineStyle = xlContinuous
            .ColorIndex = 0
            .TintAndShade = 0
            .Weight = xlThin
        End With
        
        With .Borders(xlEdgeBottom)
            .LineStyle = xlContinuous
            .ColorIndex = 0
            .TintAndShade = 0
            .Weight = xlThin
        End With
        
        With .Borders(xlEdgeRight)
            .LineStyle = xlContinuous
            .ColorIndex = 0
            .TintAndShade = 0
            .Weight = xlThin
        End With
        
        With .Borders(xlInsideVertical)
            .LineStyle = xlContinuous
            .ColorIndex = 0
            .TintAndShade = 0
            .Weight = xlThin
        End With
        
        With .Borders(xlInsideHorizontal)
            .LineStyle = xlContinuous
            .ColorIndex = 0
            .TintAndShade = 0
            .Weight = xlThin
        End With
    End With
End Sub

Private Sub CreateFormattedTable(ws As Worksheet, rangeAddress As String)
    ' Create formatted table from range
    Dim tbl As ListObject
    
    On Error GoTo TableError
    
    ' Remove any existing tables in the range
    Dim existingTable As ListObject
    For Each existingTable In ws.ListObjects
        If Not Intersect(existingTable.Range, ws.Range(rangeAddress)) Is Nothing Then
            existingTable.Delete
        End If
    Next existingTable
    
    ' Create new table
    Set tbl = ws.ListObjects.Add(xlSrcRange, ws.Range(rangeAddress), , xlYes)
    tbl.Name = "Tabelle1"
    tbl.TableStyle = "TableStyleMedium2"
    
    Exit Sub
    
TableError:
    ' If table creation fails, continue without table formatting
    On Error GoTo 0
End Sub
