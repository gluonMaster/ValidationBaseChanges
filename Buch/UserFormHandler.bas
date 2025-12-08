Attribute VB_Name = "UserFormHandler"
Option Explicit

' Handle list box click event
Public Sub HandleListBoxClick(listBox As MSForms.listBox, txtBoxEltern As MSForms.TextBox, _
                              txtBoxKonto As MSForms.TextBox, txtBoxAdresse As MSForms.TextBox, _
                              ByRef nameEltern As String, ByRef kontoNr As String, ByRef adresseInfo As String)
    Dim i As Integer, j As Integer
    Dim selectedText As String
  
    selectedText = listBox.Text
    i = InStr(selectedText, ":")
    j = Len(selectedText)
    
    txtBoxEltern.Text = Left(selectedText, (i - 1))
    txtBoxKonto.Text = Right(selectedText, 7)
    txtBoxAdresse.Text = Mid(selectedText, (i + 1), (j - 8 - i))
    
    nameEltern = txtBoxEltern.Text
    kontoNr = txtBoxKonto.Text
    adresseInfo = txtBoxAdresse.Text
End Sub

' Handle text box change event for parent/payer name - DIRECT SEARCH WITHOUT TIMER
Public Sub HandleTextBoxElternChange(listBox As MSForms.listBox, txtBoxEltern As MSForms.TextBox, lastParentRow As Integer)
    Dim searchText As String
    Dim parentData As Variant
    Dim results() As String
    Dim i As Long, resultCount As Long
    
    Application.ScreenUpdating = False
    
    ' Get search text
    searchText = LCase(Trim(txtBoxEltern.Text))
    
    ' Clear the list box
    listBox.Clear
    
    ' If search text is empty or very short, don't search
    If Len(searchText) = 0 Then
        Application.ScreenUpdating = True
        Exit Sub
    End If
    
    ' Read all parent data into memory at once
    parentData = ThisWorkbook.Worksheets("DienstTab").Range("A1:C" & lastParentRow).Value
    
    ' Prepare array for results
    ReDim results(1 To lastParentRow)
    resultCount = 0
    
    ' Find matching items
    For i = 1 To lastParentRow
        If Not IsEmpty(parentData(i, 2)) Then
            If InStr(LCase(parentData(i, 2)), searchText) > 0 Then
                resultCount = resultCount + 1
                results(resultCount) = parentData(i, 2) & ":" & parentData(i, 3) & "  " & parentData(i, 1)
            End If
        End If
    Next i
    
    ' Add results to list box
    For i = 1 To resultCount
        listBox.AddItem results(i)
    Next i
    
    Application.ScreenUpdating = True
End Sub

' Initialize form data
Public Sub InitializeFormData(ByRef pathBriefFuerEMAIL As String, ByRef elternEMAIL As String, _
                             ByRef lastParentRow As Integer)
    Dim j As Integer, jj As Integer, jj1 As Integer
    Dim ws As Worksheet
    Dim currentMonth As Integer
    
    ' Clear ErrorsLog before starting
    DataHandler.ClearErrorsLog
    
    currentMonth = Month(Date)
    pathBriefFuerEMAIL = ""
    elternEMAIL = ""
    
    ' Clear and prepare DienstTab
    lastParentRow = DataHandler.FindLastRow("DienstTab")
    Worksheets("DienstTab").Range("A1:C" & lastParentRow).Clear
    
    ' Process data from yearly sheets
    Set ws = ThisWorkbook.Worksheets("2023")
    DataHandler.ResetSheetView ws
    lastParentRow = DataHandler.ExtractParentInfo("2023")
    
    Set ws = ThisWorkbook.Worksheets("2024")
    DataHandler.ResetSheetView ws
    lastParentRow = DataHandler.ExtractParentInfo("2024")
    
    Set ws = ThisWorkbook.Worksheets("2025")
    DataHandler.ResetSheetView ws
    lastParentRow = DataHandler.ExtractParentInfo("2025")
End Sub

' Standalone procedure to initialize with status updates
Public Sub InitializeFormDataWithStatus()
    Dim frm As Object
    
    On Error Resume Next
    Set frm = UserForms("frmZahlBrief")
    
    On Error GoTo 0
    
    
    If frm Is Nothing Then
        
        InitializeFormData GlobalVariables.pathBriefFuerEMAIL, GlobalVariables.elternEMAIL, _
                          GlobalVariables.elternLetzte
    Else
        
        frm.Caption = "Payment Letter Generation - Processing 2023 data..."
        
        InitializeFormData GlobalVariables.pathBriefFuerEMAIL, GlobalVariables.elternEMAIL, _
                          GlobalVariables.elternLetzte
                                     
        
        frm.Caption = "Payment Letter Generation"
    End If
End Sub
