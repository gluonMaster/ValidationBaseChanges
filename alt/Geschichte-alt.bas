Attribute VB_Name = "Geschichte"
'==========================
'   Module: Geschichte (Data file version)
'   Purpose: Generate history report for a SINGLE selected record
'   View-only report with extended field support (new format compatible)
'
'   User interface:
'   - User selects a row on Kartei sheet (any cell in the row)
'   - Runs GeschichteMachen macro
'   - System creates/updates "Geschichte_Einzel" sheet with all history events
'   - Date range can be set in B1/C1 of the history sheet (optional filter)
'
'   Column Structure (A-AB) - same as Superadmin Geschichte:
'   A=FamilyID, B=Parent, C=Child, D=Birthdate, E=Address, F=Phone, G=Mobile, H=Email
'   I=Subject1, J=Price1, K=Subject2, L=Price2, M-X=Months 1-12, Y-AA=Extra1-3
'   AB=Comments
'
'   Uses: History_ParseData.ParseHistory for parsing both legacy and new formats
'==========================

Option Explicit

Private Const SHEET_NAME As String = "Geschichte_Einzel"
Private Const HISTORY_COL As Long = 52       ' AZ - history column
Private Const ID_COL As Long = 48            ' AV - record ID column

' Main entry point: Generate history for the currently selected record on Kartei
' User selects any cell in a row on Kartei, then runs this macro
Sub GeschichteMachen()
    Dim wsKartei As Worksheet
    Dim wsHistory As Worksheet
    Dim currentRow As Long
    Dim recordID As String
    Dim strGeschichte As String
    Dim result As Collection
    Dim startDate As Date
    Dim endDate As Date
    Dim useDateFilter As Boolean
    Dim outputRow As Long
    Dim evt As Object
    Dim eventDate As Date
    Dim i As Long
    
    On Error GoTo Cleanup
    
    ' Validate we're on Kartei sheet
    If ActiveSheet.Name <> "Kartei" Then
        MsgBox "Bitte waehlen Sie zuerst eine Zeile auf dem Blatt 'Kartei' aus.", vbExclamation, "Geschichte"
        Exit Sub
    End If
    
    Set wsKartei = ActiveSheet
    currentRow = ActiveCell.Row
    
    ' Validate row selection (skip header rows)
    If currentRow < 3 Then
        MsgBox "Bitte waehlen Sie eine Datenzeile aus (Zeile 3 oder hoeher).", vbExclamation, "Geschichte"
        Exit Sub
    End If
    
    ' Get record ID and history
    recordID = CStr(wsKartei.Cells(currentRow, ID_COL).Value)
    strGeschichte = wsKartei.Cells(currentRow, HISTORY_COL).Value
    
    If recordID = "" Then
        MsgBox "Die ausgewaehlte Zeile hat keine gueltige ID.", vbExclamation, "Geschichte"
        Exit Sub
    End If
    
    If strGeschichte = "" Then
        MsgBox "Fuer diesen Datensatz (ID: " & recordID & ") wurde keine Aenderungshistorie gefunden.", vbInformation, "Geschichte"
        Exit Sub
    End If
    
    ' Disable screen updating for performance
    Application.ScreenUpdating = False
    Application.Calculation = xlCalculationManual
    
    ' Get or create history sheet
    Set wsHistory = GetOrCreateHistorySheet()
    
    ' Check if date filter should be applied
    useDateFilter = False
    If IsDate(wsHistory.Range("B1").Value) And IsDate(wsHistory.Range("C1").Value) Then
        startDate = CDate(wsHistory.Range("B1").Value)
        endDate = CDate(wsHistory.Range("C1").Value)
        useDateFilter = True
    End If
    
    ' Clear previous content (keep row 1 with date inputs and row 2 with headers)
    wsHistory.Rows("3:" & wsHistory.Rows.Count).Clear
    
    ' Create headers
    Call CreateHistoryHeaders(wsHistory)
    
    ' Display current record info in A1
    wsHistory.Range("D1").Value = "ID: " & recordID
    wsHistory.Range("D1").Font.Bold = True
    
    ' Parse history using new parser
    Set result = History_ParseData.ParseHistory(strGeschichte)
    
    If result.Count = 0 Then
        MsgBox "Die Historie konnte nicht geparst werden oder enthaelt keine gueltigen Ereignisse.", vbInformation, "Geschichte"
        GoTo Cleanup
    End If
    
    outputRow = 3 ' Start output from row 3
    
    ' Process each event
    For i = 1 To result.Count
        Set evt = result(i)
        
        ' Apply date filter if enabled
        If useDateFilter Then
            On Error Resume Next
            eventDate = CDate(evt("ChangeDate"))
            If Err.Number = 0 Then
                If eventDate < startDate Or eventDate > endDate Then
                    Err.Clear
                    GoTo NextEvent
                End If
            Else
                Err.Clear
                GoTo NextEvent
            End If
            On Error GoTo Cleanup
        End If
        
        ' Create history entry
        Call CreateHistoryEntry(wsHistory, wsKartei, outputRow, _
                               evt("IsRuck"), evt("Reason"), CStr(evt("ChangeDate")), _
                               evt("Changes"), currentRow)
        outputRow = outputRow + 3 ' Each entry: War + Ist + Separator
        
NextEvent:
    Next i
    
    ' Apply final formatting
    Call FormatHistorySheet(wsHistory, outputRow - 1)
    
    ' Apply AutoFilter if we have data
    If outputRow > 3 Then
        If wsHistory.AutoFilterMode Then
            wsHistory.AutoFilterMode = False
        End If
        wsHistory.Range("A2:AB" & (outputRow - 1)).AutoFilter
    End If
    
    ' Ensure date format in B1:C1
    wsHistory.Range("B1:C1").NumberFormat = "dd.mm.yyyy"
    
    ' Activate history sheet
    wsHistory.Activate
    
    ' Show summary
    Dim filterInfo As String
    If useDateFilter Then
        filterInfo = vbCrLf & "Zeitraum: " & Format(startDate, "dd.mm.yyyy") & " - " & Format(endDate, "dd.mm.yyyy")
    Else
        filterInfo = vbCrLf & "(Kein Datumsfilter - alle Ereignisse angezeigt)"
    End If
    
    MsgBox "Geschichte fuer ID " & recordID & " erfolgreich generiert." & filterInfo & vbCrLf & _
           "Anzahl Ereignisse: " & ((outputRow - 3) / 3), vbInformation, "Geschichte"
    
Cleanup:
    Application.Calculation = xlCalculationAutomatic
    Application.ScreenUpdating = True
    
    If Err.Number <> 0 Then
        MsgBox "Es ist ein Fehler aufgetreten: " & Err.Description, vbCritical, "Geschichte - Fehler"
    End If
End Sub

' Get or create the history sheet
Private Function GetOrCreateHistorySheet() As Worksheet
    Dim ws As Worksheet
    
    On Error Resume Next
    Set ws = ThisWorkbook.Worksheets(SHEET_NAME)
    On Error GoTo 0
    
    If ws Is Nothing Then
        ' Create new sheet after Kartei
        Dim wsKartei As Worksheet
        On Error Resume Next
        Set wsKartei = ThisWorkbook.Worksheets("Kartei")
        On Error GoTo 0
        
        If wsKartei Is Nothing Then
            Set ws = ThisWorkbook.Worksheets.Add(After:=ThisWorkbook.Worksheets(ThisWorkbook.Worksheets.Count))
        Else
            Set ws = ThisWorkbook.Worksheets.Add(After:=wsKartei)
        End If
        ws.Name = SHEET_NAME
        
        ' Add date range inputs
        ws.Range("A1").Value = "Zeitraum:"
        ws.Range("A1").Font.Bold = True
        ws.Range("B1:C1").NumberFormat = "dd.mm.yyyy"
        ' Leave B1/C1 empty - no filter by default (show all events)
    Else
        ' Sheet exists - preserve date values if valid
        ws.Range("B1:C1").NumberFormat = "dd.mm.yyyy"
    End If
    
    Set GetOrCreateHistorySheet = ws
End Function

' Create headers for the history sheet (columns A-AB)
Private Sub CreateHistoryHeaders(ws As Worksheet)
    Dim headers As Variant
    headers = Array("FamilyID", "Eltern", "Kind", "Geburtsdatum", "Adresse", "Telefon", "Mobil", "Email", _
                    "Fach1", "Preis1", "Fach2", "Preis2", _
                    "Jan", "Feb", "Mrz", "Apr", "Mai", "Jun", "Jul", "Aug", "Sep", "Okt", "Nov", "Dez", _
                    "Extra1", "Extra2", "Extra3", "Kommentar")
    
    Dim i As Long
    For i = LBound(headers) To UBound(headers)
        With ws.Cells(2, i + 1)
            .Value = headers(i)
            .Font.Bold = True
            .HorizontalAlignment = xlCenter
        End With
    Next i
    
    ' Set column widths
    ws.Columns("A").ColumnWidth = 10  ' FamilyID
    ws.Columns("B").ColumnWidth = 20  ' Eltern
    ws.Columns("C").ColumnWidth = 18  ' Kind
    ws.Columns("D").ColumnWidth = 12  ' Geburtsdatum
    ws.Columns("E").ColumnWidth = 25  ' Adresse
    ws.Columns("F").ColumnWidth = 14  ' Telefon
    ws.Columns("G").ColumnWidth = 14  ' Mobil
    ws.Columns("H").ColumnWidth = 22  ' Email
    ws.Columns("I").ColumnWidth = 18  ' Fach1
    ws.Columns("J").ColumnWidth = 8   ' Preis1
    ws.Columns("K").ColumnWidth = 18  ' Fach2
    ws.Columns("L").ColumnWidth = 8   ' Preis2
    ws.Columns("M:X").ColumnWidth = 6 ' Months 1-12
    ws.Columns("Y:AA").ColumnWidth = 15 ' Extra1-3
    ws.Columns("AB").ColumnWidth = 40 ' Kommentar
    
    ' Set text format for non-numeric columns (to prevent auto-conversion)
    ws.Range("A2:L" & ws.Rows.Count).NumberFormat = "@"
    ws.Range("Y2:AB" & ws.Rows.Count).NumberFormat = "@"
End Sub

' Create a single history entry (War/Ist rows + separator)
Private Sub CreateHistoryEntry(wsHistory As Worksheet, wsKartei As Worksheet, _
                               startRow As Long, isRuck As Boolean, _
                               reason As String, changeDate As String, _
                               Changes As Object, karteiRow As Long)
    
    ' Kartei column mapping:
    ' A(1)=FamilyID, B(2)=Parent, D(4)=Child, E(5)=Birthdate, F(6)=Address
    ' G(7)=Phone, H(8)=Mobile, I(9)=Email, J(10)=Subject1, M(13)=Price1
    ' O(15)=Subject2, R(18)=Price2, U-AF(21-32)=Months, AK-AM(37-39)=Extra1-3
    
    Dim rowWar As Long, rowIst As Long, rowSeparator As Long
    Dim i As Long
    Dim decimalSeparator As String
    
    rowWar = startRow
    rowIst = startRow + 1
    rowSeparator = startRow + 2
    
    decimalSeparator = Application.International(xlDecimalSeparator)
    
    ' Create separator row
    wsHistory.Rows(rowSeparator).RowHeight = wsHistory.StandardHeight * 0.25
    wsHistory.Range("A" & rowSeparator & ":AB" & rowSeparator).Interior.Color = RGB(192, 192, 192)
    
    ' Fill current values for both rows (from Kartei)
    For i = 0 To 1
        Dim currentRow As Long
        currentRow = startRow + i
        
        wsHistory.Range("A" & currentRow).Value = CStr(wsKartei.Cells(karteiRow, 1).Value)   ' FamilyID
        wsHistory.Range("B" & currentRow).Value = CStr(wsKartei.Cells(karteiRow, 2).Value)   ' Parent
        wsHistory.Range("C" & currentRow).Value = CStr(wsKartei.Cells(karteiRow, 4).Value)   ' Child
        wsHistory.Range("D" & currentRow).Value = FormatDateAsText(wsKartei.Cells(karteiRow, 5).Value) ' Birthdate
        wsHistory.Range("E" & currentRow).Value = CStr(wsKartei.Cells(karteiRow, 6).Value)   ' Address
        wsHistory.Range("F" & currentRow).Value = FormatAsText(wsKartei.Cells(karteiRow, 7).Value)  ' Phone
        wsHistory.Range("G" & currentRow).Value = FormatAsText(wsKartei.Cells(karteiRow, 8).Value)  ' Mobile
        wsHistory.Range("H" & currentRow).Value = CStr(wsKartei.Cells(karteiRow, 9).Value)   ' Email
        wsHistory.Range("I" & currentRow).Value = CStr(wsKartei.Cells(karteiRow, 10).Value)  ' Subject1
        wsHistory.Range("J" & currentRow).Value = CStr(wsKartei.Cells(karteiRow, 13).Value)  ' Price1
        wsHistory.Range("K" & currentRow).Value = CStr(wsKartei.Cells(karteiRow, 15).Value)  ' Subject2
        wsHistory.Range("L" & currentRow).Value = CStr(wsKartei.Cells(karteiRow, 18).Value)  ' Price2
        
        ' Extra subjects
        wsHistory.Range("Y" & currentRow).Value = CStr(wsKartei.Cells(karteiRow, 37).Value)  ' Extra1
        wsHistory.Range("Z" & currentRow).Value = CStr(wsKartei.Cells(karteiRow, 38).Value)  ' Extra2
        wsHistory.Range("AA" & currentRow).Value = CStr(wsKartei.Cells(karteiRow, 39).Value) ' Extra3
    Next i
    
    ' Fill comment with Ruck highlighting
    If isRuck Then
        wsHistory.Range("AB" & rowIst).Interior.Color = RGB(255, 153, 204) ' Pink for Ruck
    Else
        wsHistory.Range("AB" & rowIst).Interior.Color = RGB(204, 255, 153) ' Green for normal
    End If
    
    ' Add change date to comment
    Dim commentText As String
    commentText = changeDate
    If Len(reason) > 0 Then
        commentText = commentText & " - " & reason
    End If
    wsHistory.Range("AB" & rowIst).Value = commentText
    
    ' Process changes from history
    Dim changeKey As Variant
    Dim colIndex As Long
    
    On Error Resume Next
    
    For Each changeKey In Changes.Keys
        Dim keyStr As String
        keyStr = CStr(changeKey)
        
        Dim warVal As String, istVal As String
        warVal = Changes(changeKey)("War")
        istVal = Changes(changeKey)("Ist")
        
        ' Check if it's a month number (1-12)
        If IsNumeric(keyStr) Then
            Dim monthNum As Long
            monthNum = CLng(keyStr)
            
            If monthNum >= 1 And monthNum <= 12 Then
                colIndex = 12 + monthNum ' M=13 for month 1, X=24 for month 12
                
                ' War value
                If warVal = "" Then
                    wsHistory.Cells(rowWar, colIndex).Value = 0
                Else
                    wsHistory.Cells(rowWar, colIndex).Value = ConvertToNumber(warVal, decimalSeparator)
                End If
                
                ' Ist value
                If istVal = "" Then
                    wsHistory.Cells(rowIst, colIndex).Value = 0
                Else
                    wsHistory.Cells(rowIst, colIndex).Value = ConvertToNumber(istVal, decimalSeparator)
                End If
                
                ' Highlight changed cell
                wsHistory.Cells(rowIst, colIndex).Interior.Color = RGB(255, 192, 203)
            End If
        Else
            ' Non-numeric key - field changes
            Dim targetCol As Long
            targetCol = GetColumnForFieldKey(keyStr)
            
            If targetCol > 0 Then
                ' Set War value
                wsHistory.Cells(rowWar, targetCol).Value = warVal
                ' Set Ist value with highlight
                wsHistory.Cells(rowIst, targetCol).Value = istVal
                wsHistory.Cells(rowIst, targetCol).Interior.Color = RGB(255, 192, 203)
            End If
        End If
    Next changeKey
    
    On Error GoTo 0
    
    ' Format month columns as numbers
    With wsHistory.Range("M" & rowWar & ":X" & rowIst)
        .NumberFormat = "0.00"
    End With
    
    ' Add borders
    With wsHistory.Range("A" & rowWar & ":AB" & rowIst).Borders
        .LineStyle = xlContinuous
        .Color = vbBlack
        .Weight = xlThin
    End With
End Sub

' Map field key to column number
Private Function GetColumnForFieldKey(ByVal keyStr As String) As Long
    Select Case keyStr
        Case "FID"
            GetColumnForFieldKey = 1  ' A
        Case "PAR"
            GetColumnForFieldKey = 2  ' B
        Case "CHD"
            GetColumnForFieldKey = 3  ' C
        Case "DOB"
            GetColumnForFieldKey = 4  ' D
        Case "Address", "ADR"
            GetColumnForFieldKey = 5  ' E
        Case "TEL"
            GetColumnForFieldKey = 6  ' F
        Case "MOB"
            GetColumnForFieldKey = 7  ' G
        Case "EML"
            GetColumnForFieldKey = 8  ' H
        Case "Subject1", "SB1"
            GetColumnForFieldKey = 9  ' I
        Case "PR1"
            GetColumnForFieldKey = 10 ' J
        Case "Subject2", "SB2"
            GetColumnForFieldKey = 11 ' K
        Case "PR2"
            GetColumnForFieldKey = 12 ' L
        Case "EX1"
            GetColumnForFieldKey = 25 ' Y
        Case "EX2"
            GetColumnForFieldKey = 26 ' Z
        Case "EX3"
            GetColumnForFieldKey = 27 ' AA
        Case Else
            GetColumnForFieldKey = 0  ' Unknown field
    End Select
End Function

' Convert string to number with decimal separator handling
Private Function ConvertToNumber(ByVal Value As String, ByVal decSep As String) As Double
    Value = Replace(Value, " ", "")
    
    If decSep = "," Then
        Value = Replace(Value, ".", ",")
    Else
        Value = Replace(Value, ",", ".")
    End If
    
    If IsNumeric(Value) Then
        ConvertToNumber = CDbl(Value)
    Else
        ConvertToNumber = 0
    End If
End Function

' Format value as text to prevent Excel auto-conversion
Private Function FormatAsText(ByVal Value As Variant) As String
    If IsEmpty(Value) Or IsNull(Value) Then
        FormatAsText = ""
    Else
        FormatAsText = "'" & CStr(Value)
    End If
End Function

' Format date as text in dd.mm.yyyy format
Private Function FormatDateAsText(ByVal Value As Variant) As String
    If IsEmpty(Value) Or IsNull(Value) Then
        FormatDateAsText = ""
    ElseIf IsDate(Value) Then
        FormatDateAsText = Format(Value, "dd.mm.yyyy")
    Else
        FormatDateAsText = CStr(Value)
    End If
End Function

' Apply final formatting to history sheet
Private Sub FormatHistorySheet(ws As Worksheet, lastRow As Long)
    If lastRow < 3 Then Exit Sub
    
    ' Format header row
    With ws.Range("A2:AB2").Borders
        .LineStyle = xlContinuous
        .Color = vbBlack
        .Weight = xlMedium
    End With
    ws.Range("A2:AB2").Interior.Color = RGB(220, 220, 220)
End Sub

' ========================================
' Legacy compatibility - keep old sub name working
' ========================================
Sub ConvertAndFormatCells()
    ' This procedure is kept for backward compatibility but is no longer needed
    ' The new CreateHistoryEntry handles all formatting internally
    MsgBox "Diese Funktion ist nicht mehr erforderlich. " & _
           "Die Formatierung erfolgt automatisch.", vbInformation, "Information"
End Sub
