Attribute VB_Name = "frmKorKind"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim s_Pfad As String
Dim kindCount As Integer, mont As Integer
Dim letzte As Integer, kindAnzahl As Integer
Dim menge As Integer, nrRow As Integer
Dim EltName(500) As String, EltNr(500) As String
Dim EltKode As String, EName As String, KindName As String
Dim hinz As Integer
Dim Adres As String, Teleph As String, Handy As String, eml As String
Dim valueSW(20) As Double, monatSW(12) As Double

Dim wbMain As Workbook

' Module-level variable to preserve SEPA checkbox state during speich/bearbeit calls
Dim m_SepaStatusPreserved As Boolean

' Normalizes SW value: converts invalid or non-numeric inputs (dates, times, text) to "0"
Private Function NormalizeSummetSWValue(ByVal rawValue As Variant) As String
    Dim s As String
    s = Trim(CStr(rawValue))
    
    If s = "" Then
        NormalizeSummetSWValue = "0"
        Exit Function
    End If
    
    s = Replace(s, ".", ",")
    
    If IsNumeric(s) Then
        NormalizeSummetSWValue = CStr(CDbl(s))
    Else
        NormalizeSummetSWValue = "0"
    End If
End Function

Private Sub OptimizeStart()
    Application.ScreenUpdating = False
    Application.Calculation = xlCalculationManual
    Application.EnableEvents = False
End Sub

Private Sub OptimizeEnd()
    Application.Calculation = xlCalculationAutomatic
    Application.ScreenUpdating = True
    Application.EnableEvents = True
End Sub


Private Sub btnKN1_Click()
  kundig (1)
End Sub
Private Sub btnKN2_Click()
  kundig (2)
End Sub
Private Sub btnKN3_Click()
  kundig (3)
End Sub
Private Sub btnKN4_Click()
  kundig (4)
End Sub
Private Sub btnKN5_Click()
  kundig (5)
End Sub
Private Sub btnKN6_Click()
  kundig (6)
End Sub
Private Sub btnKN7_Click()
  kundig (7)
End Sub
Private Sub btnKN8_Click()
  kundig (8)
End Sub
Private Sub btnKN9_Click()
  kundig (9)
End Sub
Private Sub btnKN10_Click()
  kundig (10)
End Sub
Private Sub btnKN11_Click()
  kundig (11)
End Sub
Private Sub btnKN12_Click()
  kundig (12)
End Sub
Private Sub btnKN13_Click()
  kundig (13)
End Sub
Private Sub btnKN14_Click()
  kundig (14)
End Sub
Private Sub btnKN15_Click()
  kundig (15)
End Sub
Private Sub btnKN16_Click()
  kundig (16)
End Sub
Private Sub btnKN17_Click()
  kundig (17)
End Sub
Private Sub btnKN18_Click()
  kundig (18)
End Sub
Private Sub btnKN19_Click()
  Call kundig(19)
End Sub
Private Sub btnKN20_Click()
  Call kundig(20)
End Sub

Sub kundig(lin As Integer)
    OptimizeStart
  Dim im As Integer
  Worksheets("Kartei").Cells(nrRow + lin, 20) = "KN"
  Worksheets("Kartei").Cells(nrRow + lin, 34).NumberFormat = "00"
  Worksheets("Kartei").Cells(nrRow + lin, 34) = mont  ' 34 - AH
  Worksheets("Kartei").Cells(nrRow + lin, 35) = "kn"  ' 35 - AI
  Worksheets("Kartei").Cells(nrRow + lin, 36).NumberFormat = "dd.mm.yyyy"
  Worksheets("Kartei").Cells(nrRow + lin, 36) = VBA.Date  ' 36 - AJ
  If mont < 8 Then
    '10  -  "J"
    Worksheets("Kartei").Cells(nrRow + lin, 13).NumberFormat = "0.00"
    Worksheets("Kartei").Cells(nrRow + lin, 13) = 0
  Else
    '15  -  "O"
    Worksheets("Kartei").Cells(nrRow + lin, 18).NumberFormat = "0.00"
    Worksheets("Kartei").Cells(nrRow + lin, 18) = 0
  End If
  Worksheets("Kartei").Cells(nrRow + lin, 18).Font.Color = VBA.RGB(255, 0, 0)
  For im = mont To 12
    Worksheets("Kartei").Cells(nrRow + lin, 20 + im).NumberFormat = "0.00"
    Worksheets("Kartei").Cells(nrRow + lin, 20 + im) = 0
  Next im
  Call KorrekturKundig(nrRow + lin)
  Me.Controls("CommandButton" & lin).Visible = False
  OptimizeEnd
End Sub

Private Sub btnKuendig_Click()
    ' Full cancellation for entire child block.
    ' Marks all children in the family block as cancelled using existing kundig() logic.
    ' Does NOT delete rows or move them to another sheet.
    
    OptimizeStart
    
    ' Validate that a family block is selected
    If nrRow = 0 Or kindAnzahl <= 0 Then
        OptimizeEnd
        MsgBox "Bitte waehlen Sie zuerst eine Familie auf dem Blatt 'Kartei' aus.", vbExclamation, "Kuendigung"
        Exit Sub
    End If
    
    Dim lin As Integer
    
    ' Process each child in the block using existing kundig() logic
    ' Note: kundig() handles its own OptimizeStart/End, so we end ours first
    OptimizeEnd
    
    For lin = 1 To kindAnzahl
        ' Call existing kundig procedure for each child row
        Call kundig(lin)
    Next lin
    
    ' Hide all KN/command buttons for the current family block after full cancellation
    ' Check T column for "KN" marker to stay consistent with bearbeit() logic
    Dim i As Integer
    For i = 1 To kindAnzahl
        If Trim$(CStr(Worksheets("Kartei").Cells(nrRow + i, "T").Value)) = "KN" Then
            Me.Controls("CommandButton" & i).Visible = False
            Me.Controls("btnKN" & i).Visible = False
        End If
    Next i
    
End Sub

Private Sub KorrekturKorKind(nrRowX As Integer, skipX As Integer)
    OptimizeStart
  Dim letzteRow As Integer
   Workbooks.Open s_Pfad & "\Korrektur.xlsx"
  letzteRow = FreeRow("Anderung") + skipX
  
  If letzteRow <= 1 Then
    MsgBox "We can no found free rows in list Anderung"
    Exit Sub
  End If
  
  If skipX = 1 Then
     Workbooks("Korrektur.xlsx").Worksheets("Anderung").Cells(letzteRow - 1, "A") = "*"
  End If
  Workbooks("Korrektur.xlsx").Worksheets("Anderung").Cells(letzteRow, "A").NumberFormat = "dd.mm.yyyy"
  Workbooks("Korrektur.xlsx").Worksheets("Anderung").Cells(letzteRow, "A") = Date
  Workbooks("Korrektur.xlsx").Worksheets("Anderung").Cells(letzteRow, "B") = ThisWorkbook.Worksheets("Kartei").Cells(nrRowX, "A")
  Workbooks("Korrektur.xlsx").Worksheets("Anderung").Cells(letzteRow, "C") = ThisWorkbook.Worksheets("Kartei").Cells(nrRowX, "B")
  Workbooks("Korrektur.xlsx").Worksheets("Anderung").Cells(letzteRow, "D") = ThisWorkbook.Worksheets("Kartei").Cells(nrRowX, "D")
  Workbooks("Korrektur.xlsx").Worksheets("Anderung").Cells(letzteRow, "E") = ThisWorkbook.Worksheets("Kartei").Cells(nrRowX, "J")
  Workbooks("Korrektur.xlsx").Worksheets("Anderung").Cells(letzteRow, "F") = ThisWorkbook.Worksheets("Kartei").Cells(nrRowX, "K")
  Workbooks("Korrektur.xlsx").Worksheets("Anderung").Cells(letzteRow, "G") = ThisWorkbook.Worksheets("Kartei").Cells(nrRowX, "L")
  Workbooks("Korrektur.xlsx").Worksheets("Anderung").Cells(letzteRow, "H") = ThisWorkbook.Worksheets("Kartei").Cells(nrRowX, "M")
  Workbooks("Korrektur.xlsx").Worksheets("Anderung").Cells(letzteRow, "I") = ThisWorkbook.Worksheets("Kartei").Cells(nrRowX, "O")
  Workbooks("Korrektur.xlsx").Worksheets("Anderung").Cells(letzteRow, "J") = ThisWorkbook.Worksheets("Kartei").Cells(nrRowX, "P")
  Workbooks("Korrektur.xlsx").Worksheets("Anderung").Cells(letzteRow, "K") = ThisWorkbook.Worksheets("Kartei").Cells(nrRowX, "Q")
  Workbooks("Korrektur.xlsx").Worksheets("Anderung").Cells(letzteRow, "L") = ThisWorkbook.Worksheets("Kartei").Cells(nrRowX, "R")
  Workbooks("Korrektur.xlsx").Worksheets("Anderung").Cells(letzteRow, "M") = ThisWorkbook.Worksheets("Kartei").Cells(nrRowX, 21)
  Workbooks("Korrektur.xlsx").Worksheets("Anderung").Cells(letzteRow, "N") = ThisWorkbook.Worksheets("Kartei").Cells(nrRowX, 22)
  Workbooks("Korrektur.xlsx").Worksheets("Anderung").Cells(letzteRow, "O") = ThisWorkbook.Worksheets("Kartei").Cells(nrRowX, 23)
  Workbooks("Korrektur.xlsx").Worksheets("Anderung").Cells(letzteRow, "P") = ThisWorkbook.Worksheets("Kartei").Cells(nrRowX, 24)
  Workbooks("Korrektur.xlsx").Worksheets("Anderung").Cells(letzteRow, "Q") = ThisWorkbook.Worksheets("Kartei").Cells(nrRowX, 25)
  Workbooks("Korrektur.xlsx").Worksheets("Anderung").Cells(letzteRow, "R") = ThisWorkbook.Worksheets("Kartei").Cells(nrRowX, 26)
  Workbooks("Korrektur.xlsx").Worksheets("Anderung").Cells(letzteRow, "S") = ThisWorkbook.Worksheets("Kartei").Cells(nrRowX, 27)
  Workbooks("Korrektur.xlsx").Worksheets("Anderung").Cells(letzteRow, "T") = ThisWorkbook.Worksheets("Kartei").Cells(nrRowX, 28)
  Workbooks("Korrektur.xlsx").Worksheets("Anderung").Cells(letzteRow, "U") = ThisWorkbook.Worksheets("Kartei").Cells(nrRowX, 29)
  Workbooks("Korrektur.xlsx").Worksheets("Anderung").Cells(letzteRow, "V") = ThisWorkbook.Worksheets("Kartei").Cells(nrRowX, 30)
  Workbooks("Korrektur.xlsx").Worksheets("Anderung").Cells(letzteRow, "W") = ThisWorkbook.Worksheets("Kartei").Cells(nrRowX, 31)
  Workbooks("Korrektur.xlsx").Worksheets("Anderung").Cells(letzteRow, "X") = ThisWorkbook.Worksheets("Kartei").Cells(nrRowX, 32)
  Workbooks("Korrektur.xlsx").Close SaveChanges:=True
    OptimizeEnd
End Sub

Private Sub KorrekturKundig(nrRowX As Integer)
    OptimizeStart
  Dim letzteRow As Integer
  Workbooks.Open s_Pfad & "\Korrektur.xlsx"
  letzteRow = FreeRow("KUNDIGUNG")
  Workbooks("Korrektur.xlsx").Worksheets("KUNDIGUNG").Cells(letzteRow, "A").NumberFormat = "dd.mm.yyyy"
  Workbooks("Korrektur.xlsx").Worksheets("KUNDIGUNG").Cells(letzteRow, "A") = Date
  Workbooks("Korrektur.xlsx").Worksheets("KUNDIGUNG").Cells(letzteRow, "B") = ThisWorkbook.Worksheets("Kartei").Cells(nrRowX, "A")
  Workbooks("Korrektur.xlsx").Worksheets("KUNDIGUNG").Cells(letzteRow, "C") = ThisWorkbook.Worksheets("Kartei").Cells(nrRowX, "B")
  Workbooks("Korrektur.xlsx").Worksheets("KUNDIGUNG").Cells(letzteRow, "D") = ThisWorkbook.Worksheets("Kartei").Cells(nrRowX, "D")
  Workbooks("Korrektur.xlsx").Worksheets("KUNDIGUNG").Cells(letzteRow, "E") = ThisWorkbook.Worksheets("Kartei").Cells(nrRowX, "J")
  Workbooks("Korrektur.xlsx").Worksheets("KUNDIGUNG").Cells(letzteRow, "F") = ThisWorkbook.Worksheets("Kartei").Cells(nrRowX, "K")
  Workbooks("Korrektur.xlsx").Worksheets("KUNDIGUNG").Cells(letzteRow, "G") = ThisWorkbook.Worksheets("Kartei").Cells(nrRowX, "L")
  Workbooks("Korrektur.xlsx").Worksheets("KUNDIGUNG").Cells(letzteRow, "H") = ThisWorkbook.Worksheets("Kartei").Cells(nrRowX, "M")
  Workbooks("Korrektur.xlsx").Worksheets("KUNDIGUNG").Cells(letzteRow, "I") = ThisWorkbook.Worksheets("Kartei").Cells(nrRowX, "O")
  Workbooks("Korrektur.xlsx").Worksheets("KUNDIGUNG").Cells(letzteRow, "J") = ThisWorkbook.Worksheets("Kartei").Cells(nrRowX, "P")
  Workbooks("Korrektur.xlsx").Worksheets("KUNDIGUNG").Cells(letzteRow, "K") = ThisWorkbook.Worksheets("Kartei").Cells(nrRowX, "Q")
  Workbooks("Korrektur.xlsx").Worksheets("KUNDIGUNG").Cells(letzteRow, "L") = ThisWorkbook.Worksheets("Kartei").Cells(nrRowX, "R")
  Workbooks("Korrektur.xlsx").Worksheets("KUNDIGUNG").Cells(letzteRow, "M").NumberFormat = "00"
  Workbooks("Korrektur.xlsx").Worksheets("KUNDIGUNG").Cells(letzteRow, "M") = mont
 Workbooks("Korrektur.xlsx").Close SaveChanges:=True
 OptimizeEnd
End Sub

Private Function FreeRow(TabName As String) As Integer
    Dim i As Integer
    On Error GoTo ErrorHandler
    
    For i = 2 To 2000
        If Workbooks("Korrektur.xlsx").Worksheets(TabName).Cells(i, "A") = "" Then
            FreeRow = i
            Exit Function
        End If
    Next i
    
    ' If empty rows not existed
    FreeRow = 0
    Exit Function
    
ErrorHandler:
    FreeRow = 0
End Function

Private Sub btnNoch_Click()
 Dim j As Integer, jhinz As Integer
 If mont = 0 Then
   MsgBox ("Monat?????")
   Exit Sub
 End If
 
 If nrRow = 0 Then
   MsgBox ("Elterns Name?????")
   Exit Sub
 End If
 hinz = hinz + 1
 If hinz > 20 Then
   MsgBox "20 ist Maximum"
   Exit Sub
 End If
 jhinz = Worksheets("Kartei").Cells(nrRow + hinz - 1, 33)
 Me.Height = 200 + hinz * 60

 Me.Controls("btnKN" & hinz).Visible = True
 Me.Controls("Nm" & (32 + hinz)).Visible = True
 Me.Controls("Nm" & (32 + hinz)).text = Me.Controls("Nm" & (32 + hinz - 1)).text
 Me.Controls("GbD" & (32 + hinz)).Visible = True
 Me.Controls("GbD" & (32 + hinz)).text = Me.Controls("GbD" & (32 + hinz - 1)).text
 Me.Controls("ComboBox" & (32 + hinz)).Visible = True
 Me.Controls("lehr" & (32 + hinz)).Visible = True
 Me.Controls("tg" & (32 + hinz)).Visible = True
 Me.Controls("Preis" & (32 + hinz)).Visible = True
 Me.Controls("Bmk" & (32 + hinz)).Visible = True
 Me.Controls("txtBox_SummetSW" & (32 + hinz)).Visible = True
 Me.Controls("CommandButton" & hinz).Visible = True
 ThisWorkbook.Worksheets("Kartei").Activate
 Rows((nrRow + hinz) & ":" & (nrRow + hinz)).Select
 Selection.Insert Shift:=xlDown, CopyOrigin:=xlFormatFromLeftOrAbove
 Worksheets("Kartei").Cells(nrRow + hinz, 1) = EltKode
 Worksheets("Kartei").Cells(nrRow + hinz, 2) = EName
' Worksheets("Kartei").Cells(nrRow + hinz, 33).NumberFormat = "0"
' Worksheets("Kartei").Cells(nrRow + hinz, "AG") = jhinz + 1
' Worksheets("Kartei").Cells(nrRow + hinz, 34).NumberFormat = "00"
' Worksheets("Kartei").Cells(nrRow + hinz, 34) = mont  ' 34 - AH
' Worksheets("Kartei").Cells(nrRow + hinz, 35) = "na"  ' 35 - AI
' Worksheets("Kartei").Cells(nrRow + hinz, 36).NumberFormat = "dd.mm.yyyy"
' Worksheets("Kartei").Cells(nrRow + hinz, 36) = Date  ' 36 - AJ

 nrRow = nrRow

End Sub

Private Sub btnPause_Click()
 Dim ip As Integer, i_d As Integer
 If mont = 0 Then
   MsgBox ("Monat?????")
   Exit Sub
 End If
 frmPause.Show
 If alle Then
    For i_d = nrRow + 1 To nrRow + kindAnzahl
      Worksheets("Kartei").Cells(i_d, 34).NumberFormat = "00"
      Worksheets("Kartei").Cells(i_d, 34) = mont   ' 34 - AH
      Worksheets("Kartei").Cells(i_d, 35) = "pa"   ' 35 - AI
      Worksheets("Kartei").Cells(i_d, 36).NumberFormat = "dd.mm.yyyy"
      Worksheets("Kartei").Cells(i_d, 36) = Date   ' 36 - AJ
      For ip = nMon1 To nMon2
          Worksheets("Kartei").Cells(i_d, 20 + ip) = 0
      Next ip
      If mont < 8 Then
          Worksheets("Kartei").Cells(i_d, "N").value = Worksheets("Kartei").Cells(i_d, "N").value + "Pause von " + CStr(nMon1) + " bis " + CStr(nMon2)
      Else
          Worksheets("Kartei").Cells(i_d, "S").value = Worksheets("Kartei").Cells(i_d, "S").value + "Pause von " + CStr(nMon1) + " bis " + CStr(nMon2)
      End If
    Next i_d
 Else
    Worksheets("Kartei").Cells(nrRow + nrPos, 34).NumberFormat = "00"
    Worksheets("Kartei").Cells(nrRow + nrPos, 34) = mont  ' 34 - AH
    Worksheets("Kartei").Cells(nrRow + nrPos, 35) = "pa"  ' 35 - AI
    Worksheets("Kartei").Cells(nrRow + nrPos, 36).NumberFormat = "dd.mm.yyyy"
    Worksheets("Kartei").Cells(nrRow + nrPos, 36) = Date  ' 36 - AJ
    For ip = nMon1 To nMon2
      Worksheets("Kartei").Cells(nrRow + nrPos, 20 + ip) = 0
    Next ip
    If mont < 8 Then
      Worksheets("Kartei").Cells(nrRow + nrPos, "N") = Worksheets("Kartei").Cells(nrRow + nrPos, "N") + "Pause von " + CStr(nMon1) + " bis " + CStr(nMon2)
    Else
      Worksheets("Kartei").Cells(nrRow + nrPos, "S") = Worksheets("Kartei").Cells(nrRow + nrPos, "S") + "Pause von " + CStr(nMon1) + " bis " + CStr(nMon2)
    End If
 End If
 MsgBox ("Fertig")
 Unload Me
 Worksheets("Kartei").Activate
End Sub

Private Sub ComboBox31_Change()
  Dim i As Integer, j As Integer
  mont = 0
  mont = ComboBox31.ListIndex + 1
'  If mont = 1 And (Month(Now()) = 12 Or Month(Now()) = 11) Then
'    mont = 27
'  End If
  If mont < 8 Then
    For i = 2 To 500
       If Worksheets("PlusBText").Cells(i, "F") = "" Then
         j = i - 1
         i = 600
      End If
    Next i
    
    For i = 33 To 52
       Me.Controls("ComboBox" & i).RowSource = "[" & ThisWorkbook.name & "]PlusBText!F2:G" & j
    Next i
  Else
    For i = 2 To 500
       If Worksheets("PlusBText").Cells(i, "L") = "" Then
         j = i - 1
         i = 600
      End If
    Next i
    
    For i = 33 To 52
       Me.Controls("ComboBox" & i).RowSource = "[" & ThisWorkbook.name & "]PlusBText!L2:M" & j
    Next i
  End If
End Sub


Sub bearbeit()
 Dim i As Integer, i_t As Integer
 Me.Frame1.Visible = True
 EName = Worksheets("Kartei").Cells(nrRow + 1, 2)
 Adres = Worksheets("Kartei").Cells(nrRow + 1, 6)
 Teleph = Worksheets("Kartei").Cells(nrRow + 1, 7)
 Handy = Worksheets("Kartei").Cells(nrRow + 1, 8)
 eml = Worksheets("Kartei").Cells(nrRow + 1, 9)
 For i = 1 To 20
 If Worksheets("Kartei").Cells(nrRow + i, 1) = EltKode Then
    Me.Height = 200 + i * 60
    Me.Controls("Nm" & (32 + i)).Visible = True
    Me.Controls("Nm" & (32 + i)).text = Worksheets("Kartei").Cells(nrRow + i, 4)
    Me.Controls("GbD" & (32 + i)).Visible = True
    Me.Controls("GbD" & (32 + i)).text = Worksheets("Kartei").Cells(nrRow + i, 5)
     If mont < 8 Then
        Me.Controls("ComboBox" & (32 + i)).Visible = True
        Me.Controls("ComboBox" & (32 + i)).text = Worksheets("Kartei").Cells(nrRow + i, 10)
        Me.Controls("lehr" & (32 + i)).Visible = True
        Me.Controls("lehr" & (32 + i)).text = Worksheets("Kartei").Cells(nrRow + i, 11)
        Me.Controls("tg" & (32 + i)).Visible = True
        Me.Controls("tg" & (32 + i)).text = Worksheets("Kartei").Cells(nrRow + i, 12)
        Me.Controls("Preis" & (32 + i)).Visible = True
        Me.Controls("Preis" & (32 + i)).text = Worksheets("Kartei").Cells(nrRow + i, 13)
        Me.Controls("Bmk" & (32 + i)).Visible = True
        Me.Controls("Bmk" & (32 + i)).text = Worksheets("Kartei").Cells(nrRow + i, 14)
        Me.Controls("txtBox_SummetSW" & (32 + i)).Visible = True
        Me.Controls("txtBox_SummetSW" & (32 + i)).text = NormalizeSummetSWValue(Worksheets("Kartei").Cells(nrRow + i, "AH").value)
        If Worksheets("Kartei").Cells(nrRow + i, "T") <> "KN" Then
           Me.Controls("CommandButton" & i).Visible = True
           Me.Controls("btnKN" & i).Visible = True
        Else
           Me.Controls("CommandButton" & i).Visible = False
           Me.Controls("btnKN" & i).Visible = False
        End If
     Else
        Me.Controls("ComboBox" & (32 + i)).Visible = True
        Me.Controls("ComboBox" & (32 + i)).text = Worksheets("Kartei").Cells(nrRow + i, 15)
        Me.Controls("lehr" & (32 + i)).Visible = True
        Me.Controls("lehr" & (32 + i)).text = Worksheets("Kartei").Cells(nrRow + i, 16)
        Me.Controls("tg" & (32 + i)).Visible = True
        Me.Controls("tg" & (32 + i)).text = Worksheets("Kartei").Cells(nrRow + i, 17)
        Me.Controls("Preis" & (32 + i)).Visible = True
        Me.Controls("Preis" & (32 + i)).text = Worksheets("Kartei").Cells(nrRow + i, 18)
        Me.Controls("Bmk" & (32 + i)).Visible = True
        Me.Controls("Bmk" & (32 + i)).text = Worksheets("Kartei").Cells(nrRow + i, 19)
        
        Me.Controls("txtBox_SummetSW" & (32 + i)).Visible = True
        Me.Controls("txtBox_SummetSW" & (32 + i)).text = NormalizeSummetSWValue(Worksheets("Kartei").Cells(nrRow + i, "AJ").value)
       '  valueSW(i) = CDbl(Me.Controls("txtBox_SummetSW" & (32 + i)).text)
        
        If Worksheets("Kartei").Cells(nrRow + i, "T") <> "KN" Then
           Me.Controls("CommandButton" & i).Visible = True
           Me.Controls("btnKN" & i).Visible = True
        Else
           Me.Controls("CommandButton" & i).Visible = False
           Me.Controls("btnKN" & i).Visible = False
        End If
     End If
     
     If InStr(1, Me.Controls("ComboBox" & 32 + i).text, "Klas") > 0 Then
     If i < 10 Then
        For i_t = 1 To 5
          Me.Controls("cmbPak" & i_t & "0" & i).Visible = True
      ' bis Juli
          If mont < 8 Then
             Me.Controls("cmbPak" & i_t & "0" & i).text = Worksheets("Kartei").Cells(nrRow + i, 36 + i_t)
          Else
          ' ab August
             Me.Controls("cmbPak" & i_t & "0" & i).text = Worksheets("Kartei").Cells(nrRow + i, 41 + i_t)
          End If
        Next i_t
     Else
        For i_t = 1 To 5
          Me.Controls("cmbPak" & i_t & i).Visible = True
          ' bis Juli
          If mont < 8 Then
             Me.Controls("cmbPak" & i_t & i).text = Worksheets("Kartei").Cells(nrRow + i, 36 + i_t)
          Else
          ' ab August
             Me.Controls("cmbPak" & i_t & i).text = Worksheets("Kartei").Cells(nrRow + i, 41 + i_t)
          End If
        Next i_t
     End If
  End If

   Else
     kindAnzahl = i - 1
      hinz = kindAnzahl
     i = 22
   End If
 Next i
 numRw = nrRow
 
    ' Verification of SEPA payment mehode
    Me.lblSEPA.Visible = False ' blend the label
    Me.switchSepa.value = False ' reset SEPA checkbox
    Me.switchSepa.Visible = True ' show SEPA checkbox when child is selected
    Dim sepaPosRow As Long
    Dim foundZahlung As Boolean
    foundZahlung = False
    
    ' Finde the " Zahlung" for the current puple
    For i = nrRow To nrRow + kindAnzahl
        If InStr(Worksheets("Kartei").Cells(i, "D").value, " Zahlung") > 0 Then
            sepaPosRow = i
            foundZahlung = True
            Exit For
        End If
    Next i
    
    ' If we find " Zahlung", verify if "SEPA" is in C
    If foundZahlung Then
        If Worksheets("Kartei").Cells(sepaPosRow, "C").value = "SEPA" Then
            Me.lblSEPA.Visible = True
            Me.switchSepa.value = True ' set SEPA checkbox to checked
        End If
    End If
 
 btnKuendig.Visible = True
 Me.btnPause.Visible = True
End Sub

Sub speich(lin As Integer)
  Dim im As Integer
  Dim strPak As String
  
  ' Preserve SEPA checkbox state before calling bearbeit
  m_SepaStatusPreserved = Me.switchSepa.value
  
  Call KorrekturKorKind(nrRow + lin, 1)

  Worksheets("Kartei").Cells(nrRow + lin, 4) = Me.Controls("Nm" & (lin + 32)).text
  Worksheets("Kartei").Cells(nrRow + lin, 5) = Me.Controls("GbD" & (lin + 32)).text
  Worksheets("Kartei").Cells(nrRow + lin, 6) = Adres
  Worksheets("Kartei").Cells(nrRow + lin, 7) = Teleph
  Worksheets("Kartei").Cells(nrRow + lin, 8) = Handy
  Worksheets("Kartei").Cells(nrRow + lin, 9) = eml
  If mont < 8 Then
    '10  -  "J"
    Worksheets("Kartei").Cells(nrRow + lin, 10) = Me.Controls("ComboBox" & (lin + 32)).text
    Worksheets("Kartei").Cells(nrRow + lin, 11) = Me.Controls("lehr" & (lin + 32)).text
    Worksheets("Kartei").Cells(nrRow + lin, 12) = Me.Controls("tg" & (lin + 32)).text
    Worksheets("Kartei").Cells(nrRow + lin, 13).NumberFormat = "0.00"
        
    If IsNumeric(Replace(Trim(Me.Controls("Preis" & (lin + 32)).text), ".", ",")) Then
        Worksheets("Kartei").Cells(nrRow + lin, 13) = CDbl(Replace(Trim(Me.Controls("Preis" & (lin + 32)).text), ".", ","))
    Else
        Worksheets("Kartei").Cells(nrRow + lin, 13) = 0
    End If
    
    
    'Worksheets("Kartei").Cells(nrRow + lin, 13) = CDbl(Replace(Trim(Me.Controls("Preis" & (lin + 32)).text), ",", "."))
    Worksheets("Kartei").Cells(nrRow + lin, 14) = Me.Controls("Bmk" & (lin + 32)).text
    
'    Call SW_Speichern(lin)
    
    Worksheets("Kartei").Cells(nrRow + lin, "AH").NumberFormat = "0.00"
    Dim swNormAH As String
    swNormAH = NormalizeSummetSWValue(Me.Controls("txtBox_SummetSW" & (lin + 32)).text)
    Me.Controls("txtBox_SummetSW" & (lin + 32)).text = swNormAH
    Worksheets("Kartei").Cells(nrRow + lin, "AH") = CDbl(swNormAH)
    
  Else
    '15  -  "O"
    If mont < 13 Then
        Worksheets("Kartei").Cells(nrRow + lin, 15) = Me.Controls("ComboBox" & (lin + 32)).text
        Worksheets("Kartei").Cells(nrRow + lin, 16) = Me.Controls("lehr" & (lin + 32)).text
        Worksheets("Kartei").Cells(nrRow + lin, 17) = Me.Controls("tg" & (lin + 32)).text
        Worksheets("Kartei").Cells(nrRow + lin, 18).NumberFormat = "0.00"
        
        If IsNumeric(Replace(Trim(Me.Controls("Preis" & (lin + 32)).text), ".", ",")) Then
            Worksheets("Kartei").Cells(nrRow + lin, 18) = CDbl(Replace(Trim(Me.Controls("Preis" & (lin + 32)).text), ".", ","))
        Else
            Worksheets("Kartei").Cells(nrRow + lin, 18) = 0
        End If

'     Worksheets("Kartei").Cells(nrRow + lin, 18) = CDbl(Me.Controls("Preis" & (lin + 32)).text)
     Worksheets("Kartei").Cells(nrRow + lin, 19) = Me.Controls("Bmk" & (lin + 32)).text
     
     Call SW_Speichern(lin)

'     If CDbl(Me.Controls("txtBox_SummetSW" & (lin + 32)).Text) > 4 Then
'       Worksheets("Kartei").Cells(nrRow + lin, "AI") = 0
'     Else
'       Worksheets("Kartei").Cells(nrRow + lin, "AI") = mont
'     End If

     Worksheets("Kartei").Cells(nrRow + lin, "AJ").NumberFormat = "0.00"
     Dim swNormAJ As String
     swNormAJ = NormalizeSummetSWValue(Me.Controls("txtBox_SummetSW" & (lin + 32)).text)
     Me.Controls("txtBox_SummetSW" & (lin + 32)).text = swNormAJ
     Worksheets("Kartei").Cells(nrRow + lin, "AJ") = CDbl(swNormAJ)
   End If
  End If
  
      '--- determine discipline text from the corresponding combobox
    Dim disc As String
    disc = Me.Controls("ComboBox" & (lin + 32)).text
    
    If InStr(1, disc, "Ind.", vbTextCompare) > 0 Or InStr(1, disc, "NH ", vbTextCompare) > 0 Or InStr(1, disc, "Nachhilfe", vbTextCompare) > 0 Or InStr(1, disc, "VSpE", vbTextCompare) > 0 Or InStr(1, disc, "VsPE", vbTextCompare) > 0 Or InStr(1, disc, "VSpE_", vbTextCompare) > 0 Then
      ' Individual lessons: zero out all monthly fees starting from mont
      For im = mont To 12
        Worksheets("Kartei").Cells(nrRow + lin, 20 + im).value = 0
      Next im
    Else
      ' Other lessons: duplicate the price into each month as before
      For im = mont To 12
        Worksheets("Kartei").Cells(nrRow + lin, 20 + im).NumberFormat = "0.00"
        If IsNumeric(Replace(Trim(Me.Controls("Preis" & (lin + 32)).text), ".", ",")) Then
          Worksheets("Kartei").Cells(nrRow + lin, 20 + im).value = _
              CDbl(Replace(Trim(Me.Controls("Preis" & (lin + 32)).text), ".", ","))
        Else
          Worksheets("Kartei").Cells(nrRow + lin, 20 + im).value = 0
        End If
      Next im
    End If
  
  If lin < 10 Then
    strPak = "0" & CStr(lin)
  Else
    strPak = CStr(lin)
  End If
  If Me.Controls("cmbPak1" & strPak).Visible Then
    For im = 1 To 5
'  bis Juli
      If mont < 8 Then
          Worksheets("Kartei").Cells(nrRow + lin, 36 + im) = Me.Controls("cmbPak" & im & strPak).text
      Else
    '  ab  August
          Worksheets("Kartei").Cells(nrRow + lin, 41 + im) = Me.Controls("cmbPak" & im & strPak).text
      End If
    Next im
  End If
  Call KorrekturKorKind(nrRow + lin, 0)

  bearbeit
  
  ' Restore SEPA checkbox state after bearbeit
  Me.switchSepa.value = m_SepaStatusPreserved
  If m_SepaStatusPreserved Then
      Me.lblSEPA.Visible = True
  End If
  
  MsgBox ("Fertig")
  Worksheets("Kartei").Activate

End Sub

Private Sub ComboBox33_Change()
  txtbox_ausf (33)
End Sub
Private Sub ComboBox34_Change()
  txtbox_ausf (34)
End Sub
Private Sub ComboBox35_Change()
  txtbox_ausf (35)
End Sub
Private Sub ComboBox36_Change()
  txtbox_ausf (36)
End Sub
Private Sub ComboBox37_Change()
  txtbox_ausf (37)
End Sub
Private Sub ComboBox38_Change()
  txtbox_ausf (38)
End Sub
Private Sub ComboBox39_Change()
  txtbox_ausf (39)
End Sub
Private Sub ComboBox40_Change()
  txtbox_ausf (40)
End Sub
Private Sub ComboBox41_Change()
  txtbox_ausf (41)
End Sub
Private Sub ComboBox42_Change()
  txtbox_ausf (42)
End Sub
Private Sub ComboBox43_Change()
  txtbox_ausf (43)
End Sub
Private Sub ComboBox44_Change()
  txtbox_ausf (44)
End Sub
Private Sub ComboBox45_Change()
  txtbox_ausf (45)
End Sub
Private Sub ComboBox46_Change()
  txtbox_ausf (46)
End Sub
Private Sub ComboBox47_Change()
  txtbox_ausf (47)
End Sub
Private Sub ComboBox48_Change()
  txtbox_ausf (48)
End Sub
Private Sub ComboBox49_Change()
  txtbox_ausf (49)
End Sub
Private Sub ComboBox50_Change()
  txtbox_ausf (50)
End Sub
Private Sub ComboBox51_Change()
  txtbox_ausf (51)
End Sub
Private Sub ComboBox52_Change()
  txtbox_ausf (52)
End Sub

Sub txtbox_ausf(num As Integer)
  Dim i_t As Integer
  Dim numPak As String
  
  If mont < 8 Then
     Me.Controls("lehr" & num).text = Worksheets("PlusBText").Range("G" & Me.Controls("ComboBox" & num).ListIndex + 2).value
     Me.Controls("tg" & num).text = Worksheets("PlusBText").Range("H" & Me.Controls("ComboBox" & num).ListIndex + 2).value
  Else
     Me.Controls("lehr" & num).text = Worksheets("PlusBText").Range("M" & Me.Controls("ComboBox" & num).ListIndex + 2).value
     Me.Controls("tg" & num).text = Worksheets("PlusBText").Range("N" & Me.Controls("ComboBox" & num).ListIndex + 2).value
  End If
  
  If InStr(1, Me.Controls("ComboBox" & num).text, "Klas") > 0 Then
     For i_t = 1 To 5
       If (num - 32) < 10 Then
          numPak = "0" & (num - 32)
       Else
          numPak = (num - 32)
       End If
       Me.Controls("cmbPak" & i_t & numPak).Visible = True
     Next i_t
  End If
End Sub

Private Sub CommandButton1_Click()
  Call speich(1)
End Sub

Private Sub CommandButton10_Click()
  Call speich(10)
End Sub

Private Sub CommandButton11_Click()
  Call speich(11)
End Sub

Private Sub CommandButton12_Click()
  Call speich(12)
End Sub

Private Sub CommandButton13_Click()
  Call speich(13)
End Sub

Private Sub CommandButton14_Click()
  Call speich(14)
End Sub

Private Sub CommandButton15_Click()
  Call speich(15)
End Sub

Private Sub CommandButton16_Click()
  Call speich(16)
End Sub

Private Sub CommandButton17_Click()
  Call speich(17)
End Sub

Private Sub CommandButton18_Click()
  Call speich(18)
End Sub

Private Sub CommandButton19_Click()
  Call speich(19)
End Sub

Private Sub CommandButton2_Click()
  Call speich(2)
End Sub

Private Sub CommandButton20_Click()
  Call speich(20)
End Sub


Private Sub CommandButton3_Click()
  Call speich(3)
End Sub

Private Sub CommandButton4_Click()
  Call speich(4)
End Sub

Private Sub CommandButton5_Click()
  Call speich(5)
End Sub

Private Sub CommandButton6_Click()
  Call speich(6)
End Sub

Private Sub CommandButton7_Click()
  Call speich(7)
End Sub

Private Sub CommandButton8_Click()
  Call speich(8)
End Sub

Private Sub CommandButton9_Click()
  Call speich(9)
End Sub


Private Sub ListBox1_Click()
  Dim sName As String ', EName As String
  Dim i As Integer
  Dim EltRange As Range

  For i = 1 To 20
    valueSW(i) = 0
  Next i
  
  For i = 1 To 12
    monatSW(i) = 0
  Next i
    
  EltKode = VBA.Strings.Right(frmKorKind.ListBox1.text, 7)
  i = InStr(frmKorKind.ListBox1.text, "   ")
  sName = Left(frmKorKind.ListBox1.text, i - 1)
'    i = ActiveSheet.Range("b1:b1500").Find(sName).row
    
  frmKorKind.txtElt.text = sName
  If mont = 0 Then
     MsgBox ("Monat?????")
     Exit Sub
  End If
  versteken
 KindName = frmKorKind.txtElt.text
''''''''''''''''''''''''''''''''''''''''''
'  EltKode = EltNr(ComboBox32.ListIndex + 1)
'''''''''''''''''''''''''''''''''''''''''''''''
 Set EltRange = ThisWorkbook.Worksheets("DienstTab").Cells.Find _
    (What:=EltKode, SearchFormat:=False)
'  EltKode = Workbooks("KindElternDaten_21.xlsm").Worksheets("DienstTab").Cells(EltRange.row, 1)
  
  Me.Label69.Caption = EltKode
  Worksheets("Kartei").Range("A4:AF10000").AutoFilter Field:=1
  nrRow = Worksheets("Kartei").Range("A1:A3000").Find(EltKode).row
  '************************************
  Me.lblRowNr.Caption = nrRow
  '**********************************
  Call bearbeit

End Sub


Private Sub txtElt_Change()
  Dim lng As Long
  Dim i As Integer
  Dim s_Mappe As String
'  Dim s_Pfad As String
  Application.ScreenUpdating = False
  s_Mappe = ThisWorkbook.name
'  s_Pfad = ThisWorkbook.Path
  
  With frmKorKind
    Me.ListBox1.Clear
    ThisWorkbook.Worksheets("DienstTab").Activate
    i = 0
    For lng = 2 To letzte
      If InStr(VBA.Strings.LCase(Cells(lng, 1).value), VBA.Strings.LCase(Me.txtElt.text)) > 0 Then
        Me.ListBox1.AddItem (Cells(lng, 1).value & "    " & Cells(lng, 2).value & "    " & Cells(lng, 3).value)
        i = i + 1
      End If
    Next lng
  End With
 
  Application.ScreenUpdating = True
 
 If mont = 0 Then
   MsgBox ("Monat?????")
   Exit Sub
 End If
 versteken

End Sub

Private Sub UserForm_Activate()
    OptimizeStart
    
    ' Reset Kartei sheet view to ensure no filters/hidden rows affect form logic
    Dim wsKartei As Worksheet
    Dim wsKarteiOriginal As Worksheet
    
    On Error Resume Next
    Set wsKartei = ThisWorkbook.Worksheets("Kartei")
    Set wsKarteiOriginal = ThisWorkbook.Worksheets("Kartei_Original")
    On Error GoTo 0
    
    If Not wsKartei Is Nothing And Not wsKarteiOriginal Is Nothing Then
        Call ResetSheetView(wsKartei, wsKarteiOriginal)
    End If
    
 Dim i As Integer, j As Integer
 Dim p As Integer, p1 As Integer
 Dim kind As String
 
 kindCount = 1
 Me.Height = 200
 s_Pfad = ThisWorkbook.path
 
' ShowMaximizeButton = True
 ComboBox31.RowSource = "[" & ThisWorkbook.name & "]PlusBText!J2:J13"
 Me.Frame1.Visible = False
'************************************************************************************
 For i = 2 To 300
   If Left(Worksheets("PlusBText").Range("L" & i), 1) = "D" And j = 0 Then
      p = i
      j = 1
   End If
'   If VBA.Strings.Right(Worksheets("PlusBText").Range("L" & i), 1) = "" Then
   If VBA.Strings.Right(Worksheets("PlusBText").Range("F" & i), 1) = "" Then
      p1 = i
      i = 500
   End If
 Next i

 For i = 1 To 5
   For j = 1 To 20

'  bis Juli
       If mont < 8 Then
         If j < 10 Then
            Me.Controls("cmbPak" & i & "0" & j).RowSource = "[" & ThisWorkbook.name & "]PlusBText!F" & p & ":F" & p1
         Else
            Me.Controls("cmbPak" & i & j).RowSource = "[" & ThisWorkbook.name & "]PlusBText!F" & p & ":F" & p1
         End If
       Else
    '   ab August
         If j < 10 Then
            Me.Controls("cmbPak" & i & "0" & j).RowSource = "[" & ThisWorkbook.name & "]PlusBText!L" & p & ":L" & p1
         Else
            Me.Controls("cmbPak" & i & j).RowSource = "[" & ThisWorkbook.name & "]PlusBText!L" & p & ":L" & p1
         End If
       End If
   Next j
 Next i
''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
 Worksheets("Kartei").Range("A4:AF10000").AutoFilter Field:=1

 For i = 3 To 3000
   If Worksheets("Kartei").Range("A" & i) = "" Then
     letzte = i - 1
     i = 10000
   End If
 Next i
 Worksheets("DienstTab").Range("A1:K5000").Clear
 Worksheets("Kartei").Range("A3", "D" & letzte).Copy Destination:=Worksheets("DienstTab").Cells(3, 7).End(xlUp).Offset(1, 0)
''''''   Sortirung auf "D" und "B"  ''''
' Worksheets("DienstTab").Range("A2", "D" & letzte).Sort _
'  Key1:=Worksheets("DienstTab").Range("D2"), Order1:=xlAscending, Header:=xlGuess, _
'   OrderCustom:=1, MatchCase:=False, Orientation:=xlTopToBottom, _
'  Key2:=Worksheets("DienstTab").Range("B2"), Order1:=xlAscending, Header:=xlGuess, _
'   OrderCustom:=1, MatchCase:=False, Orientation:=xlTopToBottom
''''''''''''''''''''''''''''''''''''''''
 j = 2
 For i = 2 To letzte - 1
   If VBA.Strings.Mid(Worksheets("DienstTab").Range("J" & i), 1, 8) <> " Zahlung" Then
     Worksheets("DienstTab").Range("D" & j) = Worksheets("DienstTab").Range("G" & i)
     Worksheets("DienstTab").Range("E" & j) = Worksheets("DienstTab").Range("H" & i)
     Worksheets("DienstTab").Range("F" & j) = Worksheets("DienstTab").Range("J" & i)
     j = j + 1
   End If
 Next i
 
 kind = Worksheets("DienstTab").Range("F2")
 Worksheets("DienstTab").Range("A2") = Worksheets("DienstTab").Range("F2")
 Worksheets("DienstTab").Range("B2") = Worksheets("DienstTab").Range("E2")
 Worksheets("DienstTab").Range("C2") = Worksheets("DienstTab").Range("D2")
 p = 3
 For i = 3 To j
   If Worksheets("DienstTab").Range("F" & i) <> kind Then
      Worksheets("DienstTab").Range("A" & p) = Worksheets("DienstTab").Range("F" & i)
      Worksheets("DienstTab").Range("B" & p) = Worksheets("DienstTab").Range("E" & i)
      Worksheets("DienstTab").Range("C" & p) = Worksheets("DienstTab").Range("D" & i)
      p = p + 1
      kind = Worksheets("DienstTab").Range("F" & i)
   End If
 Next i
 letzte = p
 ''''''   Sortirung auf "C"  ''''
 Worksheets("DienstTab").Range("A2", "C" & letzte).Sort _
  Key1:=Worksheets("DienstTab").Range("A2"), Order1:=xlAscending, Header:=xlGuess, _
  OrderCustom:=1, MatchCase:=False, Orientation:=xlTopToBottom, _
  Key2:=Worksheets("DienstTab").Range("C2"), Order1:=xlAscending, Header:=xlGuess, _
  OrderCustom:=1, MatchCase:=False, Orientation:=xlTopToBottom
''''''''''''''''''''''''''''''''''''''''
 Worksheets("DienstTab").Range("D1:K5000").Clear
 menge = p - 2
 
 nrRow = 0
 hinz = 1
 
    ' Formatting SEPA label and checkbox
    Me.lblSEPA.Font.Bold = True
    Me.lblSEPA.Font.Size = 18
    Me.lblSEPA.Visible = False
    
    ' Initialize SEPA checkbox (hidden until a child is selected)
    Me.switchSepa.value = False
    Me.switchSepa.Visible = False
 
 OptimizeEnd
End Sub

Sub versteken()
 Dim c As Control
 
 For Each c In frmKorKind.Controls
   If Left(c.name, 2) = "Nm" Then
       If CInt(VBA.Strings.Mid(c.name, 3)) > 11 Then
          c.Visible = False
       End If
   End If
   If Left(c.name, 3) = "GbD" Then
       If CInt(VBA.Strings.Mid(c.name, 4)) > 33 Then
          c.Visible = False
       End If
   End If
   If Left(c.name, 8) = "ComboBox" Then
       If CInt(VBA.Strings.Mid(c.name, 9)) > 33 Then
          c.Visible = False
       End If
   End If
   If Left(c.name, 4) = "lehr" Then
       If CInt(VBA.Strings.Mid(c.name, 5)) > 33 Then
          c.Visible = False
       End If
   End If
   If Left(c.name, 2) = "tg" Then
       If CInt(VBA.Strings.Mid(c.name, 3)) > 33 Then
          c.Visible = False
       End If
   End If
   If Left(c.name, 5) = "Preis" Then
       If CInt(VBA.Strings.Mid(c.name, 6)) > 33 Then
          c.Visible = False
       End If
   End If
   If Left(c.name, 3) = "Bmk" Then
       If CInt(VBA.Strings.Mid(c.name, 4)) > 33 Then
          c.Visible = False
       End If
   End If
   If Left(c.name, 5) = "btnKN" Then
       If CInt(VBA.Strings.Mid(c.name, 6)) > 2 Then
          c.Visible = False
       End If
   End If
   If Left(c.name, 13) = "CommandButton" Then
       If CInt(VBA.Strings.Mid(c.name, 14)) > 2 Then
          c.Visible = False
       End If
   End If

 Next c

End Sub

Function numMonat(namm As String)
  If namm = "Januar" Then
    numMonat = 1
  End If
  If namm = "Februar" Then
    numMonat = 2
  End If
  If namm = "Maerz" Then
    numMonat = 3
  End If
  If namm = "April" Then
    numMonat = 4
  End If
  If namm = "Mai" Then
    numMonat = 5
  End If
  If namm = "Juni" Then
    numMonat = 6
  End If
  If namm = "Juli" Then
    numMonat = 7
  End If
  If namm = "August" Then
    numMonat = 8
  End If
  If namm = "September" Then
    numMonat = 9
  End If
  If namm = "Oktober" Then
    numMonat = 10
  End If
  If namm = "November" Then
    nm = 11
  End If
  If namm = "Dezember" Then
    numMonat = 12
  End If
End Function

Sub SW_Speichern(linX As Integer)
 
' Dim j As Integer
' Dim SumAlleSW As Double
'
'    SumAlleSW = 0
'
'    For j = 1 To kindAnzahl
'      If CDbl(Me.Controls("txtBox_SummetSW" & (j + 32)).Text) < 6 Then
'        SumAlleSW = SumAlleSW + CDbl(Me.Controls("txtBox_SummetSW" & (j + 32)).Text)
'      End If
'    Next j
'
'    Call InSW_DatenSpeich(mont, EltKode, SumAlleSW)
'
'    If CDbl(Me.Controls("txtBox_SummetSW" & (linX + 32)).Text) < 6 Then
'        Worksheets("Kartei").Cells(nrRow + lin, "AG") = 0
'    Else
'        Worksheets("Kartei").Cells(nrRow + lin, "AG") = mont
'    End If
End Sub

Private Sub InSW_DatenSpeich(montX As Integer, EltKode As String, SumAlleSWX As Double)
'  Dim iX As Integer, jX As Integer
'  Dim found1 As Range
'
''  SW_Daten_23
'
'    B = MappeOffen(s_Pfad & "\SW_Daten_23.xlsm")
'    If B = False Then
'       Workbooks.Open s_Pfad & "\SW_Daten_23.xlsm"
'    Else
'       Workbooks(s_Pfad & "\SW_Daten_23.xlsm").Activate
'    End If
'
'    Set found1 = ActiveSheet.Range("A1:A1500").Find(EltKode)
'
'    If Not found1 Is Nothing Then
'       iX = ActiveSheet.Range("A1:A1500").Find(EltKode).row
'       For jX = montX To 12
'         ActiveSheet.Cells(iX, jX + 3) = SumAlleSWX
'       Next jX
'    End If
'    ActiveWorkbook.Close SaveChanges:=True

End Sub
Function MappeOffen(s As String) As Boolean
On Error GoTo fehler
    MappeOffen = True
    Windows(s).Activate
    Exit Function
    
fehler:
    MappeOffen = False
End Function

' ============================================================================
' SEPA Status Management
' ============================================================================

Private Sub UserForm_QueryClose(Cancel As Integer, CloseMode As Integer)
    ' Apply SEPA status to all records in the block when form closes
    ' Only process if a child block was selected (nrRow > 0 and kindAnzahl > 0)
    If nrRow > 0 And kindAnzahl > 0 Then
        Call ApplySepaStatus
    End If
    
    ' Navigate to Kartei sheet after closing the form
    On Error Resume Next
    ThisWorkbook.Worksheets("Kartei").Activate
    On Error GoTo 0
End Sub

Private Sub ApplySepaStatus()
    ' Applies or removes SEPA status in column AU (and column C for Zahlung row)
    ' for all records in the current block based on the state of switchSepa checkbox
    
    Dim i As Long
    Dim cellValueT As String
    Dim cellValueD As String
    Dim ws As Worksheet
    
    On Error GoTo Cleanup
    OptimizeStart
    
    Set ws = ThisWorkbook.Worksheets("Kartei")
    
    ' Loop through all records in the block (from nrRow to nrRow + kindAnzahl)
    For i = nrRow To nrRow + kindAnzahl
        ' Get value from column T (case-insensitive check for "KN")
        cellValueT = UCase(Trim(CStr(ws.Cells(i, "T").value)))
        ' Get value from column D (check for " Zahlung")
        cellValueD = CStr(ws.Cells(i, "D").value)
        
        If Me.switchSepa.value = True Then
            ' SEPA checkbox is checked - set "SEPA" in AU for non-KN records
            If cellValueT <> "KN" Then
                ws.Cells(i, "AU").value = "SEPA"
            End If
            
            ' Also set "SEPA" in column C for the " Zahlung" marker row
            If InStr(cellValueD, " Zahlung") > 0 Then
                ws.Cells(i, "C").value = "SEPA"
            End If
        Else
            ' SEPA checkbox is unchecked - clear AU for all records in block
            ws.Cells(i, "AU").value = ""
            
            ' Also clear column C for the " Zahlung" marker row
            If InStr(cellValueD, " Zahlung") > 0 Then
                ws.Cells(i, "C").value = ""
            End If
        End If
    Next i
    
Cleanup:
    OptimizeEnd
    If Err.Number <> 0 Then
        MsgBox "Error applying SEPA status: " & Err.Description, vbExclamation, "SEPA Status Error"
    End If
End Sub

