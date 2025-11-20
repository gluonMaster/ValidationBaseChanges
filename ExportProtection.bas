Attribute VB_Name = "ExportProtection"
'=====================================
'   Code Section 1: modOperatorCheck
'=====================================
Option Explicit

Public Function GetCurrentUserRole() As String
    ' Reads the operator role from Kartei!J1
    Dim ws As Worksheet
    Set ws = ThisWorkbook.Worksheets("Kartei")
    
    Dim roleValue As String
    roleValue = Trim(ws.Range("J1").value)
    If roleValue = "" Then
        roleValue = "Operator"
    End If
    
    GetCurrentUserRole = roleValue
End Function

Public Function IsOperator() As Boolean
    ' Returns True if the current user role is "Operator"
    Dim role As String
    role = Trim(GetCurrentUserRole())
    IsOperator = (UCase(role) = "OPERATOR")
End Function

Public Function IsSepaRow(ByVal ws As Worksheet, ByVal rowIndex As Long) As Boolean
    ' Returns True if column AU (47) contains "SEPA" marker
    Dim sepaValue As String
    sepaValue = Trim(CStr(ws.Cells(rowIndex, 47).value))
    IsSepaRow = (sepaValue = "SEPA")
End Function

Public Function GetDataYear() As Long
    ' Reads the data year from Kartei!O1
    Dim ws As Worksheet
    Set ws = ThisWorkbook.Worksheets("Kartei")
    
    Dim y As Variant
    y = ws.Range("O1").value  ' column 15, row 1
    
    If IsNumeric(y) Then
        GetDataYear = CLng(y)
    Else
        GetDataYear = Year(Date)  ' default to current year if not numeric
    End If
End Function

Public Function IsOperatorAllowedToChange(ByVal dataYear As Long) As Boolean
    ' monthIndex = 1 for January, 2 for February, ... 12 for December
    ' dataYear = year from O1
    ' We compare with the actual current date (Year(Date), Month(Date)).

    Dim nowYear As Long
    nowYear = Year(Date)
    Dim nowMonth As Long
    nowMonth = Month(Date)
    
    ' If nowYear < dataYear => the data year is "future" => operator can change everything
    If nowYear < dataYear Then
        IsOperatorAllowedToChange = True
        Exit Function
    End If
    
    ' If nowYear > dataYear => all months are past => cannot change
    ' If nowYear = dataYear => some months can past => cannot change past months
    If nowYear >= dataYear Then
        IsOperatorAllowedToChange = False
        Exit Function
    End If
    
End Function


'========================================
'   Code Section 2: modPreventPastMonths
'========================================

Public Function ValidateAndFixPastMonths() As Boolean
    ' Returns True if we can proceed with updating the DB,
    ' or False if we must abort due to operator restrictions.

    Dim role As String
    role = GetCurrentUserRole()
    
    If role = "Admin" Then
        ' No restriction
        ValidateAndFixPastMonths = True
        Exit Function
    End If
    
    ' If role=Operator => check dataYear from O1
    Dim y As Long
    y = GetDataYear()
    
    Dim ws As Worksheet
    Set ws = ThisWorkbook.Worksheets("Kartei")
    
    ' We look at range U..AF => columns 21..32
    ' For each row from row 3 to the last row, if there's data in "Kartei",
    ' we determine if the month is allowed or not. If not allowed => clear or revert.
    
    ' But first we check if *all* months are disallowed => then we abort
    Dim nowYear As Long
    nowYear = Year(Date)
    
    If nowYear > y Then
        ' Entire year is in the past => Operator cannot change anything => abort
        MsgBox "You cannot change any data from year " & y & ". Operation canceled.", vbExclamation
        ValidateAndFixPastMonths = False
        Exit Function
    End If
    
    ' If nowYear < y => everything is in the future => no restriction
    If nowYear < y Then
        ValidateAndFixPastMonths = True
        Exit Function
    End If
    
    ' If nowYear = y => only months >= current month are allowed
    ' We check if there's at least one month >= currentMonth in U..AF
    Dim currentM As Long
    currentM = Month(Date)
    
    If currentM > 12 Then currentM = 12 ' just in case, but it won't happen realistically
    
    If currentM > 12 Then
        ' means entire year is ended
        MsgBox "No data can be changed. Operation canceled.", vbExclamation
        ValidateAndFixPastMonths = False
        Exit Function
    End If
    
    ' Actually, let's do a row-level fix
    Call FixPastMonthsInRows(y)
    ValidateAndFixPastMonths = True
End Function

Private Sub FixPastMonthsInRows(ByVal dataYear As Long)
    ' For each row in Kartei, for each column U..AF => decide if it's "past month" => revert or clear
    ' In a real scenario, we might revert to old value, or just do nothing special but skip in final update.
    ' Here we do a sample approach: if month < current month => clear cell.
    
    Application.ScreenUpdating = False
    Application.Calculation = xlCalculationManual

    
    Dim ws As Worksheet
    Dim wsOriginal As Worksheet
    Set ws = ThisWorkbook.Worksheets("Kartei")
    Set wsOriginal = ThisWorkbook.Worksheets("Kartei_Original")
    
    Dim lastRow As Long
    Dim firstRow As Long
    firstRow = 3
    lastRow = ws.Cells(ws.Rows.count, 1).End(xlUp).row
    
    Dim currentM As Long
    currentM = Month(Date)
    
    Dim forbiddenEndCol As Integer
    forbiddenEndCol = 20 + (currentM - 1) 'if month is 7, forbiddenEndCol=26
    Dim forbiddenStartCol As Integer
    forbiddenStartCol = 21
    
    If Not IsOperatorAllowedToChange(dataYear) Then
        ' We simply restore the cell => means "operator not allowed to change"
        If forbiddenStartCol <= forbiddenEndCol Then
            Dim addr As String
            addr = ws.Range(ws.Cells(firstRow, forbiddenStartCol), _
                            ws.Cells(lastRow, forbiddenEndCol)).Address
            ' e.g. "U3:AB250"
            
            ws.Range(addr).value = wsOriginal.Range(addr).value
        End If
    End If
  
    Application.Calculation = xlCalculationAutomatic
    Application.ScreenUpdating = True
End Sub


