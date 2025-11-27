Attribute VB_Name = "valid_TransposeForm"
'==========================
'   Module: valid_TransposeForm
'   Purpose: Create transposed view of a GrossGeschichte record for easier review
'   Creates temporary sheet "geschichteForm" with vertical layout
'   
'   Column structure on geschichteForm:
'     A = Field Name (header)
'     B = War (old value)
'     C = Ist (new value)
'   
'   Includes "Apply & Close" button to save decision and remove temp sheet
'==========================

Option Explicit

' Constants for the temporary sheet
Private Const TEMP_SHEET_NAME As String = "geschichteForm"

' Store the source row from GrossGeschichte (module-level for button callback)
Private mSourceGrossRow As Long

' Field names array matching GrossGeschichte column structure A-AE
' Note: RecordID (AE) is hidden in GrossGeschichte but shown in transpose form for reference
Private Function GetFieldNames() As Variant
    GetFieldNames = Array("FamilyID", "Parent", "Child", "Birthdate", "Address", "Phone", "Mobile", "Email", _
                          "Subject1", "Price1", "Subject2", "Price2", _
                          "Jan", "Feb", "Mrz", "Apr", "Mai", "Jun", "Jul", "Aug", "Sep", "Okt", "Nov", "Dez", _
                          "Extra1", "Extra2", "Extra3", "Comments", "Decision", "Decline Comment", "RecordID")
End Function

' Main entry point: Show transposed form for current record
Public Sub ShowTransposeForm()
    On Error GoTo ErrorHandler
    
    Application.ScreenUpdating = False
    
    ' Check if we're on GrossGeschichte sheet
    Dim wsGross As Worksheet
    On Error Resume Next
    Set wsGross = ThisWorkbook.Worksheets("GrossGeschichte")
    On Error GoTo ErrorHandler
    
    If wsGross Is Nothing Then
        MsgBox "GrossGeschichte sheet not found. Please generate the history report first.", _
               vbExclamation, "Sheet Not Found"
        GoTo Cleanup
    End If
    
    If ActiveSheet.Name <> "GrossGeschichte" Then
        MsgBox "Please navigate to GrossGeschichte sheet first.", _
               vbExclamation, "Wrong Sheet"
        GoTo Cleanup
    End If
    
    ' Get current row and determine the record block
    Dim currentRow As Long
    currentRow = ActiveCell.Row
    
    ' Validate row is in data area (row 3+)
    If currentRow < 3 Then
        MsgBox "Please select a data row (row 3 or below).", vbExclamation, "Invalid Selection"
        GoTo Cleanup
    End If
    
    ' Determine War and Ist rows for this record block
    ' Structure: Row N = War, Row N+1 = Ist, Row N+2 = Separator
    ' So blocks start at rows 3, 6, 9, 12, ...
    Dim blockStartRow As Long
    Dim rowWar As Long
    Dim rowIst As Long
    
    blockStartRow = GetBlockStartRow(currentRow)
    rowWar = blockStartRow
    rowIst = blockStartRow + 1
    
    ' Store source row for button callback (Ist row contains the decision cells)
    mSourceGrossRow = rowIst
    
    ' Validate we have data in this block
    If Trim(wsGross.Range("A" & rowWar).Value) = "" Then
        MsgBox "No data found in this record block.", vbExclamation, "Empty Record"
        GoTo Cleanup
    End If
    
    ' Create or recreate the temp sheet
    Dim wsForm As Worksheet
    Set wsForm = GetOrCreateTempSheet()
    
    ' Populate the transposed data
    Call PopulateTransposeForm(wsForm, wsGross, rowWar, rowIst)
    
    ' Add the "Apply & Close" button
    Call AddApplyButton(wsForm)
    
    ' Activate the form sheet
    wsForm.Activate
    wsForm.Range("A1").Select
    
    Application.ScreenUpdating = True
    Exit Sub
    
Cleanup:
    Application.ScreenUpdating = True
    Exit Sub
    
ErrorHandler:
    Application.ScreenUpdating = True
    MsgBox "Error creating transpose form: " & Err.Description, vbCritical, "Error"
End Sub

' Calculate the starting row of the block containing the given row
' Blocks are: rows 3-5 (block 1), 6-8 (block 2), 9-11 (block 3), etc.
Private Function GetBlockStartRow(ByVal row As Long) As Long
    ' Each block is 3 rows: War, Ist, Separator
    ' Block 1 starts at row 3
    ' Block N starts at row 3 + (N-1)*3
    
    Dim blockIndex As Long
    blockIndex = (row - 3) \ 3 ' Integer division to get block index (0-based)
    
    GetBlockStartRow = 3 + blockIndex * 3
End Function

' Get or create the temporary sheet
Private Function GetOrCreateTempSheet() As Worksheet
    Dim ws As Worksheet
    
    ' Delete existing temp sheet if exists
    On Error Resume Next
    Set ws = ThisWorkbook.Worksheets(TEMP_SHEET_NAME)
    If Not ws Is Nothing Then
        Application.DisplayAlerts = False
        ws.Delete
        Application.DisplayAlerts = True
    End If
    On Error GoTo 0
    
    ' Create new sheet
    Set ws = ThisWorkbook.Worksheets.Add(After:=ThisWorkbook.Worksheets(ThisWorkbook.Worksheets.Count))
    ws.Name = TEMP_SHEET_NAME
    
    Set GetOrCreateTempSheet = ws
End Function

' Populate the transpose form with data from GrossGeschichte
Private Sub PopulateTransposeForm(ByVal wsForm As Worksheet, ByVal wsGross As Worksheet, _
                                   ByVal rowWar As Long, ByVal rowIst As Long)
    Dim fieldNames As Variant
    fieldNames = GetFieldNames()
    
    ' Set up headers in row 1
    wsForm.Range("A1").Value = "Field"
    wsForm.Range("B1").Value = "War (Old)"
    wsForm.Range("C1").Value = "Ist (New)"
    
    With wsForm.Range("A1:C1")
        .Font.Bold = True
        .Interior.Color = RGB(220, 220, 220)
        .HorizontalAlignment = xlCenter
    End With
    
    ' CRITICAL: Set text format for columns B and C BEFORE writing any values
    ' This prevents Excel from auto-converting phone numbers like 015154358922 to 1.51E+10
    wsForm.Columns("B:C").NumberFormat = "@"
    
    ' Populate field names and values
    Dim i As Long
    Dim outputRow As Long
    outputRow = 2
    
    For i = LBound(fieldNames) To UBound(fieldNames)
        Dim colIndex As Long
        colIndex = i + 1 ' Column A=1, B=2, etc.
        
        ' Field name
        wsForm.Cells(outputRow, 1).Value = fieldNames(i)
        wsForm.Cells(outputRow, 1).Font.Bold = True
        
        ' War value (from War row) - read as text from source
        wsForm.Cells(outputRow, 2).Value = CStr(wsGross.Cells(rowWar, colIndex).Value)
        
        ' Ist value (from Ist row) - read as text from source
        wsForm.Cells(outputRow, 3).Value = CStr(wsGross.Cells(rowIst, colIndex).Value)
        
        ' Highlight changed cells (where War <> Ist)
        If CStr(wsGross.Cells(rowWar, colIndex).Value) <> CStr(wsGross.Cells(rowIst, colIndex).Value) Then
            ' Check if original Ist cell has highlight
            If wsGross.Cells(rowIst, colIndex).Interior.Color = RGB(255, 192, 203) Then
                wsForm.Cells(outputRow, 3).Interior.Color = RGB(255, 192, 203) ' Light pink
            End If
        End If
        
        ' Special formatting for specific fields
        Select Case fieldNames(i)
            Case "Comments"
                ' Copy comment cell color from GrossGeschichte
                wsForm.Cells(outputRow, 3).Interior.Color = wsGross.Cells(rowIst, colIndex).Interior.Color
                
            Case "Decision"
                ' Add dropdown validation
                With wsForm.Cells(outputRow, 3).Validation
                    .Delete
                    .Add Type:=xlValidateList, AlertStyle:=xlValidAlertStop, _
                         Formula1:="Approved,Declined"
                    .IgnoreBlank = True
                    .InCellDropdown = True
                End With
                wsForm.Cells(outputRow, 3).Interior.Color = RGB(255, 255, 204) ' Light yellow
                
            Case "Decline Comment"
                wsForm.Cells(outputRow, 3).Interior.Color = RGB(255, 230, 230) ' Light pink
                
            Case "RecordID"
                ' RecordID is read-only (informational), shown in gray
                wsForm.Cells(outputRow, 2).Interior.Color = RGB(230, 230, 230) ' Light gray
                wsForm.Cells(outputRow, 3).Interior.Color = RGB(230, 230, 230) ' Light gray
                wsForm.Cells(outputRow, 1).Font.Italic = True
        End Select
        
        outputRow = outputRow + 1
    Next i
    
    ' Set column widths
    wsForm.Columns("A").ColumnWidth = 18
    wsForm.Columns("B").ColumnWidth = 30
    wsForm.Columns("C").ColumnWidth = 30
    
    ' Add borders
    With wsForm.Range("A1:C" & (outputRow - 1)).Borders
        .LineStyle = xlContinuous
        .Color = vbBlack
        .Weight = xlThin
    End With
    
    ' Store source row reference in a hidden cell for button callback
    wsForm.Range("E1").Value = mSourceGrossRow
    wsForm.Columns("E").Hidden = True
End Sub

' Add "Apply & Close" button to the form
Private Sub AddApplyButton(ByVal wsForm As Worksheet)
    ' Calculate button position (below the data)
    Dim lastRow As Long
    lastRow = wsForm.Cells(wsForm.Rows.Count, 1).End(xlUp).Row
    
    Dim buttonTop As Double
    buttonTop = wsForm.Cells(lastRow + 2, 1).Top
    
    Dim buttonLeft As Double
    buttonLeft = wsForm.Cells(lastRow + 2, 2).Left
    
    ' Create Forms button (not ActiveX)
    Dim btn As Button
    Set btn = wsForm.Buttons.Add(buttonLeft, buttonTop, 120, 30)
    
    With btn
        .Name = "btnApplyClose"
        .Caption = "Apply && Close"
        .OnAction = "valid_TransposeForm.ApplyDecisionAndClose"
        .Font.Size = 11
        .Font.Bold = True
    End With
    
    ' Add instruction label
    wsForm.Cells(lastRow + 2, 1).Value = "Select Decision:"
    wsForm.Cells(lastRow + 2, 1).Font.Bold = True
End Sub

' Button callback: Apply decision to GrossGeschichte and close temp sheet
Public Sub ApplyDecisionAndClose()
    On Error GoTo ErrorHandler
    
    Application.ScreenUpdating = False
    
    Dim wsForm As Worksheet
    Dim wsGross As Worksheet
    
    Set wsForm = ThisWorkbook.Worksheets(TEMP_SHEET_NAME)
    Set wsGross = ThisWorkbook.Worksheets("GrossGeschichte")
    
    ' Get source row from hidden cell
    Dim sourceRow As Long
    sourceRow = CLng(wsForm.Range("E1").Value)
    
    ' Find Decision and Decline Comment rows in the form
    Dim fieldNames As Variant
    fieldNames = GetFieldNames()
    
    Dim decisionRow As Long
    Dim declineCommentRow As Long
    Dim i As Long
    
    For i = LBound(fieldNames) To UBound(fieldNames)
        If fieldNames(i) = "Decision" Then
            decisionRow = i + 2 ' +2 because data starts at row 2
        ElseIf fieldNames(i) = "Decline Comment" Then
            declineCommentRow = i + 2
        End If
    Next i
    
    ' Get values from form
    Dim decision As String
    Dim declineComment As String
    
    decision = Trim(wsForm.Cells(decisionRow, 3).Value)
    declineComment = Trim(wsForm.Cells(declineCommentRow, 3).Value)
    
    ' Validate decision
    If decision = "" Then
        MsgBox "Please select a Decision (Approved or Declined) before applying.", _
               vbExclamation, "Decision Required"
        Application.ScreenUpdating = True
        Exit Sub
    End If
    
    If UCase(decision) <> "APPROVED" And UCase(decision) <> "DECLINED" Then
        MsgBox "Decision must be 'Approved' or 'Declined'.", _
               vbExclamation, "Invalid Decision"
        Application.ScreenUpdating = True
        Exit Sub
    End If
    
    ' For Declined, check if comment is provided
    If UCase(decision) = "DECLINED" And declineComment = "" Then
        Dim resp As VbMsgBoxResult
        resp = MsgBox("No decline comment provided. Continue anyway?", _
                     vbYesNo + vbQuestion, "Missing Comment")
        If resp = vbNo Then
            Application.ScreenUpdating = True
            Exit Sub
        End If
    End If
    
    ' Apply decision to GrossGeschichte
    ' Column AC (29) = Decision, Column AD (30) = Decline Comment
    wsGross.Cells(sourceRow, 29).Value = decision ' AC
    wsGross.Cells(sourceRow, 30).Value = declineComment ' AD
    
    ' Delete the temp sheet
    Application.DisplayAlerts = False
    wsForm.Delete
    Application.DisplayAlerts = True
    
    ' Activate GrossGeschichte and go to the updated row
    wsGross.Activate
    wsGross.Cells(sourceRow, 29).Select
    
    Application.ScreenUpdating = True
    
    MsgBox "Decision applied successfully." & vbCrLf & vbCrLf & _
           "Decision: " & decision & vbCrLf & _
           IIf(declineComment <> "", "Comment: " & declineComment, ""), _
           vbInformation, "Applied"
    
    Exit Sub
    
ErrorHandler:
    Application.ScreenUpdating = True
    MsgBox "Error applying decision: " & Err.Description, vbCritical, "Error"
End Sub

' Close the transpose form without applying changes
Public Sub CloseTransposeForm()
    On Error Resume Next
    
    Dim wsForm As Worksheet
    Set wsForm = ThisWorkbook.Worksheets(TEMP_SHEET_NAME)
    
    If Not wsForm Is Nothing Then
        ' Return to GrossGeschichte
        Dim wsGross As Worksheet
        Set wsGross = ThisWorkbook.Worksheets("GrossGeschichte")
        
        Application.DisplayAlerts = False
        wsForm.Delete
        Application.DisplayAlerts = True
        
        If Not wsGross Is Nothing Then
            wsGross.Activate
        End If
    End If
End Sub

' Helper function for IIf (not available in all VBA versions)
Private Function IIf(ByVal condition As Boolean, ByVal truePart As Variant, ByVal falsePart As Variant) As Variant
    If condition Then
        IIf = truePart
    Else
        IIf = falsePart
    End If
End Function
