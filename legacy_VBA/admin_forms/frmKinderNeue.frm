VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} frmKinderNeue 
   Caption         =   "Liste der Neue Kinder und Unterricht"
   ClientHeight    =   12384
   ClientLeft      =   108
   ClientTop       =   456
   ClientWidth     =   18900
   OleObjectBlob   =   "frmKinderNeue.frx":0000
   StartUpPosition =   1  'CenterOwner
End
Attribute VB_Name = "frmKinderNeue"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim s_Pfad As String
Dim nmon As Integer
Dim nameKind(1 To 4) As String, vornmKind(1 To 4) As String, gdKind(1 To 4) As String
Dim mengeKind As Integer
Dim paket As Integer, paketVertrag(1 To 10) As String
Dim NameElt As String, vornameElt As String
Dim Anschrift As String, PLZ As String, Wohnort As String
Dim Telefon As String, Mobil As String, email As String
Dim numExcel As String
Dim vertragTyp As Integer
Dim kindCount As Integer
Dim onlJa As Boolean, onlNein As Boolean
Dim sparkasse As Boolean, Unterricht As Boolean, online As Boolean
Dim fotos As Boolean, videos As Boolean, ton As Boolean, kunst As Boolean, aufsat As Boolean
Dim website As Boolean, socmed As Boolean, raum As Boolean, veranst As Boolean
Dim mitname As Boolean, onename As Boolean

Private Sub cmbName1_Change()
  Me.tbgd1.text = Me.Controls("txtGebDt" & Me.cmbName1.ListIndex + 1).text
End Sub

Private Sub cmbName10_Change()
  Me.tbgd10.text = Me.Controls("txtGebDt" & Me.cmbName10.ListIndex + 1).text
End Sub

Private Sub cmbName11_Change()
  Me.tbgd11.text = Me.Controls("txtGebDt" & Me.cmbName11.ListIndex + 1).text
End Sub

Private Sub cmbName12_Change()
  Me.tbgd12.text = Me.Controls("txtGebDt" & Me.cmbName12.ListIndex + 1).text
End Sub

Private Sub cmbName13_Change()
  Me.tbgd13.text = Me.Controls("txtGebDt" & Me.cmbName13.ListIndex + 1).text
End Sub

Private Sub cmbName14_Change()
  Me.tbgd14.text = Me.Controls("txtGebDt" & Me.cmbName14.ListIndex + 1).text
End Sub

Private Sub cmbName2_Change()
  Me.tbgd2.text = Me.Controls("txtGebDt" & Me.cmbName2.ListIndex + 1).text
End Sub

Private Sub cmbName3_Change()
  Me.tbgd3.text = Me.Controls("txtGebDt" & Me.cmbName3.ListIndex + 1).text
End Sub

Private Sub cmbName4_Change()
  Me.tbgd4.text = Me.Controls("txtGebDt" & Me.cmbName4.ListIndex + 1).text
End Sub

Private Sub cmbName5_Change()
  Me.tbgd5.text = Me.Controls("txtGebDt" & Me.cmbName5.ListIndex + 1).text
End Sub

Private Sub cmbName6_Change()
  Me.tbgd6.text = Me.Controls("txtGebDt" & Me.cmbName6.ListIndex + 1).text
End Sub

Private Sub cmbName7_Change()
  Me.tbgd7.text = Me.Controls("txtGebDt" & Me.cmbName7.ListIndex + 1).text
End Sub

Private Sub cmbName8_Change()
  Me.tbgd8.text = Me.Controls("txtGebDt" & Me.cmbName8.ListIndex + 1).text
End Sub

Private Sub cmbName9_Change()
  Me.tbgd9.text = Me.Controls("txtGebDt" & Me.cmbName9.ListIndex + 1).text
End Sub

Private Sub cmbOKkind_Click()
 Dim i As Integer
 
 If Me.txtFamNm1.text = "" Then
   MsgBox ("Familienname bitte")
   Exit Sub
 End If
 
 For i = 1 To 4
  If Me.Controls("txtFamNm" & i).text <> "" Then
    nameKind(i) = Me.Controls("txtFamNm" & i).text
    vornmKind(i) = Me.Controls("txtVorname" & i).text
    gdKind(i) = Me.Controls("txtGebDt" & i).text
    mengeKind = i
  End If
 Next
 
 For i = 1 To 14
   For j = 1 To mengeKind
     Me.Controls("cmbName" & i).AddItem nameKind(j) & " ; " & vornmKind(j)
   Next j
 Next i
 
 Me.cmbName1.Visible = True
 Me.tbgd1.Visible = True
 Me.ComboBox1.Visible = True
 Me.tblehr1.Visible = True
 Me.tbtag1.Visible = True
 Me.tbpreis1.Visible = True
 Me.tbbem1.Visible = True
 Me.LstPaket1.Visible = True
 Me.TxtBox_SW11.Visible = True
 Me.btnNoch.Visible = True
 Me.btnOK.Visible = True
 
End Sub



Private Sub UserForm_Activate()
   Dim i As Integer, iEnd As Integer
   
   kindCount = 1
   s_Pfad = ThisWorkbook.path
   nmon = numMonat(frmNeueElt.ComboBox31.text)
   numExcel = frmNeueElt.txtNummer.text
   
   NameElt = frmNeueElt.txtFamilName.text
   vornameElt = frmNeueElt.txtVorname.text
   Anschrift = frmNeueElt.txtAnschrift.text
   PLZ = frmNeueElt.txtPLZ.text
   Wohnort = frmNeueElt.txtWohnort.text
   Telefon = frmNeueElt.txtTelefon.text
   Mobil = frmNeueElt.txtMobil.text
   email = frmNeueElt.txtEmail.text
   
   For i = 1 To 4
    paketVertrag(i) = ""
   Next i
   
   For i = 1 To 14
     Me.Controls("cmbName" & i).Visible = False
     Me.Controls("tbgd" & i).Visible = False
     Me.Controls("ComboBox" & i).Visible = False
     Me.Controls("tblehr" & i).Visible = False
     Me.Controls("tbtag" & i).Visible = False
     Me.Controls("tbpreis" & i).Visible = False
     Me.Controls("tbbem" & i).Visible = False
   Next i
  
   If nmon < 8 Then
        iEnd = LetzteNr("PlusBText", 2, "F")
        For i = 1 To 14
          Me.Controls("ComboBox" & i).RowSource = "[" & ThisWorkbook.name & "]PlusBText!F2:F" & iEnd
          Me.Controls("LstPaket" & i).RowSource = "[" & ThisWorkbook.name & "]PlusBText!F2:F" & iEnd
          paket = iEnd
        Next i
   Else
        iEnd = LetzteNr("PlusBText", 2, "L")
        For i = 1 To 14
            Me.Controls("ComboBox" & i).RowSource = "[" & ThisWorkbook.name & "]PlusBText!L2:L" & iEnd
            Me.Controls("LstPaket" & i).RowSource = "[" & ThisWorkbook.name & "]PlusBText!L2:L" & iEnd
            paket = iEnd
        Next i
   End If
   
   vertragTyp = VertfagAuswal(frmNeueElt.obtnJahr.value, frmNeueElt.obtnMonat.value)
'   nachhilfe = frmNeueElt.obtnNH.Value
   onlJa = frmNeueElt.obtnOnlineUnterJa.value
   sparkasse = frmNeueElt.obtnSparkasse
   Unterricht = frmNeueElt.obtnOnlineUnterJa
   fotos = frmNeueElt.cboxFotos.value
   videos = frmNeueElt.cboxVideos.value
   ton = frmNeueElt.cboxTonaufnahmen.value
   kunst = frmNeueElt.cboxKunstWerke.value
   aufsat = frmNeueElt.cboxAufsatze.value
   website = frmNeueElt.cboxInternet.value
   socmed = frmNeueElt.cboxSocMedia.value
   raum = frmNeueElt.cboxRaum
   veranst = frmNeueElt.cboxVeranstalt
   mitname = frmNeueElt.ObtnMitName.value
   onename = frmNeueElt.ObtnOhneName.value
   
   Me.Height = 350
End Sub


Private Sub btnNoch_Click()
 Dim j As Integer
 kindCount = kindCount + 1
 If Me.Controls("cmbName" & (kindCount - 1)).text = "" Then
   MsgBox "'Name' ist leer"
   kindCount = kindCount - 1
   Exit Sub
 End If
 If kindCount > 14 Then
   MsgBox "14 ist Maximum"
   kindCount = 20
 Else
   Me.Controls("cmbName" & kindCount).Visible = True
   Me.Controls("tbgd" & kindCount).Visible = True
   Me.Controls("ComboBox" & kindCount).Visible = True
   Me.Controls("tblehr" & kindCount).Visible = True
   Me.Controls("tbtag" & kindCount).Visible = True
   Me.Controls("tbpreis" & kindCount).Visible = True
   Me.Controls("tbbem" & kindCount).Visible = True
   Me.Controls("LstPaket" & kindCount).Visible = True
'   Me.Controls("TxtBox_SW" & kindCount + 1).Visible = True
   
   Me.Height = 350 + 20 * (kindCount)
   Me.btnNoch.Top = 30 + 20 * (kindCount)
   Me.btnOK.Top = 70 + 20 * (kindCount)
   
   If kindCount > 0 Then
        Me.Controls("LstPaket" & kindCount).Height = 20
   End If
   Me.Controls("LstPaket" & (kindCount + 1)).Height = 20
 End If
End Sub

Private Sub objVis(iVis As Integer)
 Me.Controls(obj1).Visible = False
End Sub

Private Sub btnOK_Click()
  Dim i As Integer, j As Integer, j1 As Integer
  Dim pak1 As Integer, pak2 As Integer
  Dim SumSW_Jahr As Double, SumSW_Monat As Double
  Dim NeueEltName As String
  
  SumSW_Jahr = 0
  SumSW_Monat = 0

  j = 2
  
  letzte = LetzteNr("Kartei", 3, "A")
  
  Worksheets("Kartei").Range("A" & (letzte + 1)) = numExcel
  Worksheets("Kartei").Range("B" & (letzte + 1)) = NameElt & ";" & vornameElt
  If sparkasse Then
    Worksheets("Kartei").Range("C" & (letzte + 1)) = "SEPA"
  End If
  
  Worksheets("Kartei").Range("D" & (letzte + 1)) = " Zahlung"
  Worksheets("Kartei").Range("B" & (letzte + 1) & ":AU" & (letzte + 1)).Interior.ColorIndex = 43
    
  Worksheets("Kartei").Cells((letzte + 1), 34) = nmon  ' 34 - AH
  Worksheets("Kartei").Cells((letzte + 1), 35) = "gn"  ' 35 - AI
  Worksheets("Kartei").Cells((letzte + 1), 36).NumberFormat = "dd.mm.yyyy"
  Worksheets("Kartei").Cells((letzte + 1), 36) = VBA.Date  ' 36 - AJ
    
  For i = 1 To kindCount
    Worksheets("Kartei").Range("A" & (letzte + j)) = numExcel
    Worksheets("Kartei").Range("B" & (letzte + j)) = NameElt & ";" & vornameElt
    Worksheets("Kartei").Range("D" & (letzte + j)) = Me.Controls("cmbName" & i).text
    Worksheets("Kartei").Range("E" & (letzte + j)) = Me.Controls("tbgd" & i).text
    Worksheets("Kartei").Range("F" & (letzte + j)) = Anschrift & ";" & PLZ & ";" & Wohnort
    Worksheets("Kartei").Range("G" & (letzte + j)) = Telefon
    Worksheets("Kartei").Range("H" & (letzte + j)) = Mobil
    Worksheets("Kartei").Range("I" & (letzte + j)) = email
    pak2 = 0
    If InStr(Me.Controls("ComboBox" & i).text, "Klass") Then
      For pak1 = 0 To paket - 2
        If Me.Controls("LstPaket" & i).Selected(pak1) Then
          pak2 = pak2 + 1
          If nmon < 8 Then
              Worksheets("Kartei").Cells(letzte + j, pak2 + 36) = Me.Controls("LstPaket" & i).List(pak1)
          Else
              Worksheets("Kartei").Cells(letzte + j, pak2 + 41) = Me.Controls("LstPaket" & i).List(pak1)
          End If
          paketVertrag(i) = paketVertrag(i) & ", " & Me.Controls("LstPaket" & i).List(pak1)
        End If
      Next
    End If
    
    If nmon < 8 Then
        Worksheets("Kartei").Range("J" & (letzte + j)) = Me.Controls("ComboBox" & i).text
        Worksheets("Kartei").Range("K" & (letzte + j)) = Me.Controls("tblehr" & i).text
        Worksheets("Kartei").Range("L" & (letzte + j)) = Me.Controls("tbtag" & i).text
        Worksheets("Kartei").Range("M" & (letzte + j)).NumberFormat = "0.00"
        Worksheets("Kartei").Range("M" & (letzte + j)) = CDbl(Replace(Trim(Me.Controls("tbpreis" & i).text), ".", ","))
        Worksheets("Kartei").Range("N" & (letzte + j)) = Me.Controls("tbbem" & i).text
        Worksheets("Kartei").Cells((letzte + j), 33) = nmon  ' 33 - AG
        Worksheets("Kartei").Cells((letzte + j), 34).NumberFormat = "0.00"
        If Me.Controls("TxtBox_SW" & (i + 10)) = "" Then
          Worksheets("Kartei").Cells((letzte + j), 34) = 0   ' 34 - AH
        Else
          If CDbl(Me.Controls("TxtBox_SW" & (i + 10)).text) > 6 Then
             Worksheets("Kartei").Cells((letzte + j), "AG") = 0
             SumSW_Jahr = SumSW_Jahr + CDbl(Me.Controls("TxtBox_SW" & (i + 10)).text)

          Else
             Worksheets("Kartei").Cells((letzte + j), "AG") = nmon
             SumSW_Monat = SumSW_Monat + CDbl(Me.Controls("TxtBox_SW" & (i + 10)).text)
             
          End If

          Worksheets("Kartei").Cells((letzte + j), 34) = Me.Controls("TxtBox_SW" & (i + 10))   ' 34 - AH
        End If

    Else
        Worksheets("Kartei").Range("O" & (letzte + j)) = Me.Controls("ComboBox" & i).text
        Worksheets("Kartei").Range("P" & (letzte + j)) = Me.Controls("tblehr" & i).text
        Worksheets("Kartei").Range("Q" & (letzte + j)) = Me.Controls("tbtag" & i).text
        Worksheets("Kartei").Range("R" & (letzte + j)).NumberFormat = "0.00"
        Worksheets("Kartei").Range("R" & (letzte + j)) = CDbl(Me.Controls("tbpreis" & i).text)
        Worksheets("Kartei").Range("S" & (letzte + j)) = Me.Controls("tbbem" & i).text
        Worksheets("Kartei").Cells((letzte + j), 35) = nmon  ' 35 - AI
        Worksheets("Kartei").Cells((letzte + j), 36).NumberFormat = "0.00"
        If Me.Controls("TxtBox_SW" & (i + 10)) = "" Then
          Worksheets("Kartei").Cells((letzte + j), 36) = 0   ' 36 - AJ
        Else
          If CDbl(Me.Controls("TxtBox_SW" & (i + 10)).text) > 6 Then
             Worksheets("Kartei").Cells((letzte + j), "AI") = 0
          Else
             Worksheets("Kartei").Cells((letzte + j), "AI") = nmon
             SumSW_Monat = SumSW_Monat + CDbl(Me.Controls("TxtBox_SW" & (i + 10)).text)
          End If
          Worksheets("Kartei").Cells((letzte + j), 36) = Me.Controls("TxtBox_SW" & (i + 10))   ' 36 - AJ
        End If
    End If
    
    ' Vertrag J - Jahrlich, M - Monatlich
    If vertragTyp = 1 Then
      Worksheets("Kartei").Cells((letzte + j), "T") = "J"
    Else
      If vertragTyp = 2 Then
         Worksheets("Kartei").Cells((letzte + j), "T") = "M"
      Else
         Worksheets("Kartei").Cells((letzte + j), "T") = "N"
      End If
    End If
    
      '--- get discipline text from the corresponding combobox
      Dim disc As String
      disc = Me.Controls("ComboBox" & i).text
    
      If InStr(1, disc, "Ind.", vbTextCompare) > 0 Or InStr(1, disc, "NH ", vbTextCompare) > 0 Or InStr(1, disc, "Nachhilfe", vbTextCompare) > 0 Then
        ' Individual lessons: zero out all monthly fees starting from nmon
        For j1 = nmon To 12
            Worksheets("Kartei").Cells(letzte + j, j1 + 20).NumberFormat = "0.00"
            Worksheets("Kartei").Cells(letzte + j, j1 + 20).value = 0
        Next j1
      Else
        ' Other lessons: duplicate the price into each month as before
        For j1 = nmon To 12
          Worksheets("Kartei").Cells(letzte + j, j1 + 20).NumberFormat = "0.00"
          If IsNumeric(Replace(Trim(Me.Controls("tbpreis" & i).text), ".", ",")) Then
            Worksheets("Kartei").Cells(letzte + j, j1 + 20).value = _
              CDbl(Replace(Trim(Me.Controls("tbpreis" & i).text), ".", ","))
          Else
            Worksheets("Kartei").Cells(letzte + j, j1 + 20).value = 0
          End If
        Next j1
      End If

    Call KorrekturNeueEltern(letzte + j)
    
    j = j + 1
  Next i
  Me.btnNoch.Top = 30
  Me.btnOK.Top = 78
  
  NeueEltName = NameElt & ";" & vornameElt
  
'  Call SW_ElternHinzufuegen(SumSW_Monat, NeueEltName)
  MsgBox "Fertig"
  
  If vertragTyp = 1 Then
      If mengeKind < 4 Then
           Call AusdruckenJ("C:\\Signatur\Unterrichtsvertrag 2022 neu.docx")
      Else
           Call AusdruckenJ("C:\\Signatur\Unterrichtsvertrag 2022 neuG.docx")
      End If
      
      If kindCount < 6 Then
           Call AusdrAnlage("C:\\Signatur\Unterrichtsvertrag_Anlage.docx")
      Else
           Call AusdrAnlage("C:\\Signatur\Unterrichtsvertrag_AnlageG.docx")
      End If
  End If
  
  If vertragTyp = 2 Then
      If mengeKind < 4 And kindCount < 6 Then
         Call AusdruckenM("C:\\Signatur\Unterrichtsvertrag 2022_Monatsvertrag.docx")
      Else
         Call AusdruckenM("C:\\Signatur\Unterrichtsvertrag 2022_MonatsvertragG.docx")
      End If
  End If
  
  
  For i = 1 To 14
'    Me.Controls("tbnm" & i).Text = ""
    Me.Controls("tbgd" & i).text = ""
    Me.Controls("ComboBox" & i).text = ""
    Me.Controls("tblehr" & i).text = ""
    Me.Controls("tbtag" & i).text = ""
    Me.Controls("tbpreis" & i).text = ""
    Me.Controls("tbbem" & i).text = ""
  Next i

  Call UserForm_Activate
End Sub

Private Sub KorrekturNeueEltern(nrRowX As Integer)
  Dim letzteRow As Integer
  Workbooks.Open s_Pfad & "\Korrektur.xlsx"
  letzteRow = FreeRow("NeueEltern")
  Workbooks("Korrektur.xlsx").Worksheets("NeueEltern").Cells(letzteRow, "A").NumberFormat = "dd.mm.yyyy"
  Workbooks("Korrektur.xlsx").Worksheets("NeueEltern").Cells(letzteRow, "A") = Date
  Workbooks("Korrektur.xlsx").Worksheets("NeueEltern").Cells(letzteRow, "B") = ThisWorkbook.Worksheets("Kartei").Cells(nrRowX, "A")
  Workbooks("Korrektur.xlsx").Worksheets("NeueEltern").Cells(letzteRow, "C") = ThisWorkbook.Worksheets("Kartei").Cells(nrRowX, "B")
  Workbooks("Korrektur.xlsx").Worksheets("NeueEltern").Cells(letzteRow, "D") = ThisWorkbook.Worksheets("Kartei").Cells(nrRowX, "D")
  Workbooks("Korrektur.xlsx").Worksheets("NeueEltern").Cells(letzteRow, "E") = ThisWorkbook.Worksheets("Kartei").Cells(nrRowX, "E")
  Workbooks("Korrektur.xlsx").Worksheets("NeueEltern").Cells(letzteRow, "F") = ThisWorkbook.Worksheets("Kartei").Cells(nrRowX, "F")
  Workbooks("Korrektur.xlsx").Worksheets("NeueEltern").Cells(letzteRow, "G") = ThisWorkbook.Worksheets("Kartei").Cells(nrRowX, "G")
  Workbooks("Korrektur.xlsx").Worksheets("NeueEltern").Cells(letzteRow, "H") = ThisWorkbook.Worksheets("Kartei").Cells(nrRowX, "H")
  Workbooks("Korrektur.xlsx").Worksheets("NeueEltern").Cells(letzteRow, "I") = ThisWorkbook.Worksheets("Kartei").Cells(nrRowX, "I")
  Workbooks("Korrektur.xlsx").Worksheets("NeueEltern").Cells(letzteRow, "J") = ThisWorkbook.Worksheets("Kartei").Cells(nrRowX, "J")
  Workbooks("Korrektur.xlsx").Worksheets("NeueEltern").Cells(letzteRow, "K") = ThisWorkbook.Worksheets("Kartei").Cells(nrRowX, "K")
  Workbooks("Korrektur.xlsx").Worksheets("NeueEltern").Cells(letzteRow, "L") = ThisWorkbook.Worksheets("Kartei").Cells(nrRowX, "L")
  Workbooks("Korrektur.xlsx").Worksheets("NeueEltern").Cells(letzteRow, "M") = ThisWorkbook.Worksheets("Kartei").Cells(nrRowX, "M")
  Workbooks("Korrektur.xlsx").Worksheets("NeueEltern").Cells(letzteRow, "N") = ThisWorkbook.Worksheets("Kartei").Cells(nrRowX, "O")
  Workbooks("Korrektur.xlsx").Worksheets("NeueEltern").Cells(letzteRow, "O") = ThisWorkbook.Worksheets("Kartei").Cells(nrRowX, "P")
  Workbooks("Korrektur.xlsx").Worksheets("NeueEltern").Cells(letzteRow, "P") = ThisWorkbook.Worksheets("Kartei").Cells(nrRowX, "Q")
  Workbooks("Korrektur.xlsx").Worksheets("NeueEltern").Cells(letzteRow, "Q") = ThisWorkbook.Worksheets("Kartei").Cells(nrRowX, "R")
  Workbooks("Korrektur.xlsx").Worksheets("NeueEltern").Cells(letzteRow, "R") = ThisWorkbook.Worksheets("Kartei").Cells(nrRowX, 21)
  Workbooks("Korrektur.xlsx").Worksheets("NeueEltern").Cells(letzteRow, "S") = ThisWorkbook.Worksheets("Kartei").Cells(nrRowX, 22)
  Workbooks("Korrektur.xlsx").Worksheets("NeueEltern").Cells(letzteRow, "T") = ThisWorkbook.Worksheets("Kartei").Cells(nrRowX, 23)
  Workbooks("Korrektur.xlsx").Worksheets("NeueEltern").Cells(letzteRow, 21) = ThisWorkbook.Worksheets("Kartei").Cells(nrRowX, 24)
  Workbooks("Korrektur.xlsx").Worksheets("NeueEltern").Cells(letzteRow, 22) = ThisWorkbook.Worksheets("Kartei").Cells(nrRowX, 25)
  Workbooks("Korrektur.xlsx").Worksheets("NeueEltern").Cells(letzteRow, 23) = ThisWorkbook.Worksheets("Kartei").Cells(nrRowX, 26)
  Workbooks("Korrektur.xlsx").Worksheets("NeueEltern").Cells(letzteRow, 24) = ThisWorkbook.Worksheets("Kartei").Cells(nrRowX, 27)
  Workbooks("Korrektur.xlsx").Worksheets("NeueEltern").Cells(letzteRow, 25) = ThisWorkbook.Worksheets("Kartei").Cells(nrRowX, 28)
  Workbooks("Korrektur.xlsx").Worksheets("NeueEltern").Cells(letzteRow, 26) = ThisWorkbook.Worksheets("Kartei").Cells(nrRowX, 29)
  Workbooks("Korrektur.xlsx").Worksheets("NeueEltern").Cells(letzteRow, 27) = ThisWorkbook.Worksheets("Kartei").Cells(nrRowX, 30)
  Workbooks("Korrektur.xlsx").Worksheets("NeueEltern").Cells(letzteRow, 28) = ThisWorkbook.Worksheets("Kartei").Cells(nrRowX, 31)
  Workbooks("Korrektur.xlsx").Worksheets("NeueEltern").Cells(letzteRow, 29) = ThisWorkbook.Worksheets("Kartei").Cells(nrRowX, 32)
  Workbooks("Korrektur.xlsx").Worksheets("NeueEltern").Cells(letzteRow, 30) = ThisWorkbook.Worksheets("Kartei").Cells(nrRowX, "T")
  Workbooks("Korrektur.xlsx").Close SaveChanges:=True
End Sub

Private Function FreeRow(TabName As String)
  Dim i As Integer
  For i = 2 To 5000
    If Workbooks("Korrektur.xlsx").Worksheets(TabName).Cells(i, "A") = "" Then
      FreeRow = i
      i = 6000
    End If
  Next i
End Function


'********************************************************
'      SW_Hinzufuegen
'********************************************************
Private Sub SW_ElternHinzufuegen(sumSW As Double, ElternName As String)
  Dim i As Integer, j As Integer
  Dim letz As Integer
  
'  j = 2
''  s_Pfad = ThisWorkbook.Path
'  B = MappeOffen(s_Pfad & "\SW_Daten_23.xlsm")
'  If B = False Then
'       Workbooks.Open s_Pfad & "\SW_Daten_23.xlsm"
'  Else
'       Workbooks(s_Pfad & "\SW_Daten_23.xlsm").Activate
'  End If
'
'  For i = 3 To 5000
'    If ActiveSheet.Cells(i, 1) = "" Then
'      letz = i
'      i = 5001
'    End If
'  Next i
'
'  ActiveSheet.Range("A" & letz) = numExcel
'  ActiveSheet.Range("B" & letz) = ElternName
'  ActiveSheet.Range("C" & letz) = " Zahlung"
'
'  For i = nmon + 3 To 15
'    ActiveSheet.Cells(letz, i).NumberFormat = "0.00"
'    ActiveSheet.Cells(letz, i) = sumSW
'  Next i
'
'  Call SortSW(letz)
'  ActiveWorkbook.Close SaveChanges:=True
End Sub
Sub SortSW(letz As Integer)
' Dim i As Integer, Letzte As Integer
'
' Worksheets("SW_Liste").Sort.SortFields.Clear
' Worksheets("SW_Liste").Sort.SortFields.Add Key:=Range("B3:B" & letz), _
'   SortOn:=xlSortOnValues, Order:=xlAscending, DataOption:=xlSortNormal
' With Worksheets("SW_Liste").Sort
'    .SetRange Range("A2:O" & letz)
'    .Header = xlYes
'    .MatchCase = False
'    .Orientation = xlTopToBottom
'    .SortMethod = xlPinYin
'    .Apply
' End With
 
End Sub

Private Sub btnOKelt_Click()
  frmKinderNeue.Show
End Sub

Private Sub ComboBox1_Change()
  txtbox_ausf (1)
End Sub
Private Sub ComboBox2_Change()
  txtbox_ausf (2)
End Sub
Private Sub ComboBox3_Change()
  txtbox_ausf (3)
End Sub
Private Sub ComboBox4_Change()
  txtbox_ausf (4)
End Sub
Private Sub ComboBox5_Change()
  txtbox_ausf (5)
End Sub
Private Sub ComboBox6_Change()
  txtbox_ausf (6)
End Sub
Private Sub ComboBox7_Change()
  txtbox_ausf (7)
End Sub
Private Sub ComboBox8_Change()
  txtbox_ausf (8)
End Sub
Private Sub ComboBox9_Change()
  txtbox_ausf (9)
End Sub
Private Sub ComboBox10_Change()
  txtbox_ausf (10)
End Sub
Private Sub ComboBox11_Change()
  txtbox_ausf (11)
End Sub
Private Sub ComboBox12_Change()
  txtbox_ausf (12)
End Sub
Private Sub ComboBox13_Change()
  txtbox_ausf (13)
End Sub
Private Sub ComboBox14_Change()
  txtbox_ausf (14)
End Sub



Sub txtbox_ausf(num As Integer)
  If nmon < 8 Then
     Me.Controls("tblehr" & num).text = Worksheets("PlusBText").Range("G" & Me.Controls("ComboBox" & num).ListIndex + 2).value
     Me.Controls("tbtag" & num).text = Worksheets("PlusBText").Range("H" & Me.Controls("ComboBox" & num).ListIndex + 2).value
  Else
     Me.Controls("tblehr" & num).text = Worksheets("PlusBText").Range("M" & Me.Controls("ComboBox" & num).ListIndex + 2).value
     Me.Controls("tbtag" & num).text = Worksheets("PlusBText").Range("N" & Me.Controls("ComboBox" & num).ListIndex + 2).value
  End If
  
  If InStr(Me.Controls("ComboBox" & num).text, "Klass") Then
     Me.Controls("LstPaket" & num).Visible = True
     Me.Controls("LstPaket" & num).Height = 80
  End If

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
    numMonat = 11
  End If
  If namm = "Dezember" Then
    numMonat = 12
  End If
End Function

Function MappeOffen(s As String) As Boolean
On Error GoTo fehler
    MappeOffen = True
    Windows(s).Activate
    Exit Function
    
fehler:
    MappeOffen = False
End Function

Function LetzteNr(TabName As String, ErsteRow As Integer, Spalte As String)
  Dim i As Integer
  
  For i = ErsteRow To 10000
    If Worksheets(TabName).Cells(i, Spalte) = "" Then
       LetzteNr = i - 1
       i = 10001
    End If
  Next i
End Function

Private Sub AusdruckenJ(pfadX As String)
'    Dim DokumNm As String
    Dim wda As Word.Application
    Dim wdd As Word.Document
    
    Set wdd = GetObject(pfadX)
    Set wda = wdd.Parent
    wda.Visible = True
    
    numExcel = Replace(numExcel, ". ", "")

    wdd.Bookmarks("VertNum").Select
    wda.Selection.TypeText (" " & numExcel)

    wdd.Bookmarks("FamNm").Select
    wda.Selection.TypeText ("  " & NameElt)

    wdd.Bookmarks("Vorname").Select
    wda.Selection.TypeText ("  " & vornameElt)
    
    wdd.Bookmarks("Anschrift").Select
    wda.Selection.TypeText ("  " & Anschrift)
    
    wdd.Bookmarks("PLZ").Select
    wda.Selection.TypeText ("  " & PLZ)
    
    wdd.Bookmarks("Wohnort").Select
    wda.Selection.TypeText ("  " & Wohnort)
    
    wdd.Bookmarks("Telefon").Select
    wda.Selection.TypeText ("  " & Telefon)
    
    wdd.Bookmarks("Mobil").Select
    wda.Selection.TypeText ("  " & Mobil)
    
    wdd.Bookmarks("EMail").Select
    wda.Selection.TypeText (email)

    wdd.Bookmarks("FamNmKind1").Select
    wda.Selection.TypeText ("  " & txtFamNm1)
    
    If txtFamNm2 <> "" Then
        wdd.Bookmarks("FamNmKind2").Select
        wda.Selection.TypeText ("  " & txtFamNm2)
    End If

    If txtFamNm3 <> "" Then
        wdd.Bookmarks("FamNmKind3").Select
        wda.Selection.TypeText ("  " & txtFamNm3)
    End If
    
    wdd.Bookmarks("VornameKind1").Select
    wda.Selection.TypeText ("  " & txtVorname1)
    
    If txtVorname2 <> "" Then
        wdd.Bookmarks("VornameKind2").Select
        wda.Selection.TypeText ("  " & txtVorname2)
    End If
    
    If txtVorname3 <> "" Then
        wdd.Bookmarks("VornameKind3").Select
        wda.Selection.TypeText ("  " & txtVorname3)
    End If
   
    wdd.Bookmarks("GebDtKind1").Select
    wda.Selection.TypeText ("  " & txtGebDt1)

    If txtGebDt2 <> "" Then
        wdd.Bookmarks("GebDtKind2").Select
        wda.Selection.TypeText ("  " & txtGebDt2)
    End If

    If txtGebDt3 <> "" Then
        wdd.Bookmarks("GebDtKind3").Select
        wda.Selection.TypeText ("  " & txtGebDt3)
    End If
 '*****************************************************
    If onlJa Then
        wdd.Bookmarks("OnLineJa").Select
        wda.Selection.TypeText ("+")
        wdd.Bookmarks("OnLineNein").Select
        wda.Selection.TypeText ("-")
   
    Else
        wdd.Bookmarks("OnLineJa").Select
        wda.Selection.TypeText ("-")
        wdd.Bookmarks("OnLineNein").Select
        wda.Selection.TypeText ("+")
   
    End If
'
    If sparkasse Then
        wdd.Bookmarks("BankSEPA").Select
        wda.Selection.TypeText ("+")
        wdd.Bookmarks("BankVolksbank").Select
        wda.Selection.TypeText ("-")
    Else
        wdd.Bookmarks("BankSEPA").Select
        wda.Selection.TypeText ("-")
        wdd.Bookmarks("BankVolksbank").Select
        wda.Selection.TypeText ("+")
    End If

   
    If fotos Then
        wdd.Bookmarks("Fotos").Select
        wda.Selection.TypeText ("+")
    Else
        wdd.Bookmarks("Fotos").Select
        wda.Selection.TypeText ("-")
    End If
    
    If videos Then
        wdd.Bookmarks("Videos").Select
        wda.Selection.TypeText ("+")
    Else
        wdd.Bookmarks("Videos").Select
        wda.Selection.TypeText ("-")
    End If

    If ton Then
        wdd.Bookmarks("Ton").Select
        wda.Selection.TypeText ("+")
    Else
        wdd.Bookmarks("Ton").Select
        wda.Selection.TypeText ("-")
    End If
    
    If kunst Then
        wdd.Bookmarks("kWerke").Select
        wda.Selection.TypeText ("+")
    Else
        wdd.Bookmarks("kWerke").Select
        wda.Selection.TypeText ("-")
    End If
    
    If aufsat Then
        wdd.Bookmarks("Aufsetze").Select
        wda.Selection.TypeText ("+")
    Else
        wdd.Bookmarks("Aufsetze").Select
        wda.Selection.TypeText ("-")
    End If

    If website Then
        wdd.Bookmarks("Internet").Select
        wda.Selection.TypeText ("+")
    Else
        wdd.Bookmarks("Internet").Select
        wda.Selection.TypeText ("-")
    End If

    If socmed Then
        wdd.Bookmarks("SocMedia").Select
        wda.Selection.TypeText ("+")
    Else
        wdd.Bookmarks("SocMedia").Select
        wda.Selection.TypeText ("-")
    End If

    If raum Then
        wdd.Bookmarks("Raum").Select
        wda.Selection.TypeText ("+")
    Else
        wdd.Bookmarks("Raum").Select
        wda.Selection.TypeText ("-")
    End If

    If veranst Then
        wdd.Bookmarks("Veranstalt").Select
        wda.Selection.TypeText ("+")
    Else
        wdd.Bookmarks("Veranstalt").Select
        wda.Selection.TypeText ("-")
    End If
    
    If mitname Then
        wdd.Bookmarks("MitName").Select
        wda.Selection.TypeText ("+")
        wdd.Bookmarks("OneName").Select
        wda.Selection.TypeText ("-")
   Else
        wdd.Bookmarks("MitName").Select
        wda.Selection.TypeText ("-")
        wdd.Bookmarks("OneName").Select
        wda.Selection.TypeText ("+")
    End If
    
   
    NameElt = Replace(NameElt, "ö", "o")
    NameElt = Replace(NameElt, "ü", "u")
    NameElt = Replace(NameElt, "ä", "a")
    

    wdd.SaveAs2 "C:\\Vertrage" & "\VertJ_1" & Right(numExcel, 4) & "_" & Left(NameElt, 6) & Left(vornameElt, 6) & ".pdf", _
    FileFormat:=wdExportFormatPDF
    
    wdd.Close False

 End Sub



Private Sub AusdrAnlage(pfadX As String)
    Dim DokumName As String
    Dim i As Integer
    Dim KinderName As String
    Dim wda As Word.Application
    Dim wdd As Word.Document
    
'    Set wdd = GetObject("C:\\Signatur\Unterrichtsvertrag_Anlage.docx")pfadX
    Set wdd = GetObject(pfadX)
    Set wda = wdd.Parent
    wda.Visible = True
    
    wdd.Bookmarks("VertNum").Select
    wda.Selection.TypeText (" " & numExcel)

    wdd.Bookmarks("ElternName").Select
    wda.Selection.TypeText (" " & NameElt & ", " & vornameElt)

    For i = 1 To kindCount
        wdd.Bookmarks("Datum" & i).Select
        wda.Selection.TypeText (CStr(VBA.Date))
        
        KinderName = Me.Controls("cmbName" & i).text
        KinderName = Replace(KinderName, ";", "_")
        wdd.Bookmarks("KindName" & i).Select
        wda.Selection.TypeText (KinderName)
        
        wdd.Bookmarks("Kurs" & i).Select
        If paketVertrag(i) <> "" Then
           wda.Selection.TypeText (Me.Controls("ComboBox" & i).text & paketVertrag(i))
        Else
           wda.Selection.TypeText (Me.Controls("ComboBox" & i).text)
        End If
        
        wdd.Bookmarks("Price" & i).Select
        wda.Selection.TypeText (Me.Controls("tbpreis" & i).text)
        
        wdd.Bookmarks("Bemerk" & i).Select
        wda.Selection.TypeText (Me.Controls("tbbem" & i).text)
    Next i
    
    DokumName = "\VertJ_1" & Right(numExcel, 4) & "_" & Left(NameElt, 6) & Left(vornameElt, 6) & "_Anlage.pdf"
    
    wdd.SaveAs2 "C:\\Vertrage" & DokumName, FileFormat:=wdExportFormatPDF
    
    wdd.Close False
   
 End Sub

Function VertfagAuswal(ValueJ As Boolean, ValueM As Boolean)
  If ValueJ Then
    VertfagAuswal = 1
  Else
    If ValueM Then
       VertfagAuswal = 2
    Else
       VertfagAuswal = 3
    End If
  End If
End Function

Private Sub AusdruckenM(pfadX As String)
    Dim DokumNm As String
    Dim i As Integer
    Dim KinderName As String
    Dim wda As Word.Application
    Dim wdd As Word.Document
    
'    Dim KinderName As String
    
    Set wdd = GetObject(pfadX)
    Set wda = wdd.Parent
    wda.Visible = True
    
    numExcel = Replace(numExcel, ". ", "")

    wdd.Bookmarks("VertNum").Select
    wda.Selection.TypeText (" " & numExcel)

    wdd.Bookmarks("FamNm").Select
    wda.Selection.TypeText (NameElt)

    wdd.Bookmarks("Vorname").Select
    wda.Selection.TypeText (vornameElt)
    
    wdd.Bookmarks("Anschrift").Select
    wda.Selection.TypeText (Anschrift)
    
    wdd.Bookmarks("PLZ").Select
    wda.Selection.TypeText (PLZ)
    
    wdd.Bookmarks("Wohnort").Select
    wda.Selection.TypeText (Wohnort)
    
    wdd.Bookmarks("Telefon").Select
    wda.Selection.TypeText (Telefon)
    
    wdd.Bookmarks("Mobil").Select
    wda.Selection.TypeText (Mobil)
    
    wdd.Bookmarks("EMail").Select
    wda.Selection.TypeText (email)

    wdd.Bookmarks("FamNmKind1").Select
    wda.Selection.TypeText (txtFamNm1)
    
    If txtFamNm2 <> "" Then
        wdd.Bookmarks("FamNmKind2").Select
        wda.Selection.TypeText (txtFamNm2)
    End If

    If txtFamNm3 <> "" Then
        wdd.Bookmarks("FamNmKind3").Select
        wda.Selection.TypeText (txtFamNm3)
    End If
    
    If txtFamNm4 <> "" Then
        wdd.Bookmarks("FamNmKind4").Select
        wda.Selection.TypeText (txtFamNm4)
    End If
    
    wdd.Bookmarks("VornameKind1").Select
    wda.Selection.TypeText (txtVorname1)
    
    If txtVorname2 <> "" Then
        wdd.Bookmarks("VornameKind2").Select
        wda.Selection.TypeText (txtVorname2)
    End If
    
    If txtVorname3 <> "" Then
        wdd.Bookmarks("VornameKind3").Select
        wda.Selection.TypeText (txtVorname3)
    End If
   
    If txtVorname4 <> "" Then
        wdd.Bookmarks("VornameKind4").Select
        wda.Selection.TypeText (txtVorname4)
    End If
    
    wdd.Bookmarks("GebDtKind1").Select
    wda.Selection.TypeText (txtGebDt1)

    If txtGebDt2 <> "" Then
        wdd.Bookmarks("GebDtKind2").Select
        wda.Selection.TypeText (txtGebDt2)
    End If

    If txtGebDt3 <> "" Then
        wdd.Bookmarks("GebDtKind3").Select
        wda.Selection.TypeText (txtGebDt3)
    End If
    
    If txtGebDt4 <> "" Then
        wdd.Bookmarks("GebDtKind4").Select
        wda.Selection.TypeText (txtGebDt4)
    End If
 '*****************************************************
    For i = 1 To kindCount
'        wdd.Bookmarks("Datum" & i).Select
'        wda.Selection.TypeText (CStr(VBA.Date))
        
        KinderName = Me.Controls("cmbName" & i).text
        KinderName = Replace(KinderName, ";", "_")
        wdd.Bookmarks("KindName" & i).Select
        wda.Selection.TypeText (KinderName)
        
        wdd.Bookmarks("Kurs" & i).Select
        If paketVertrag(i) <> "" Then
           wda.Selection.TypeText (Me.Controls("ComboBox" & i).text & paketVertrag(i))
        Else
           wda.Selection.TypeText (Me.Controls("ComboBox" & i).text)
        End If
        
        wdd.Bookmarks("Price" & i).Select
        wda.Selection.TypeText (Me.Controls("tbpreis" & i).text)
        
    Next i
 '************************************************************
 
 
 '************************************************************
    If onlJa Then
        wdd.Bookmarks("OnLineJa").Select
        wda.Selection.TypeText ("+")
        wdd.Bookmarks("OnLineNein").Select
        wda.Selection.TypeText ("-")
   
    Else
        wdd.Bookmarks("OnLineJa").Select
        wda.Selection.TypeText ("-")
        wdd.Bookmarks("OnLineNein").Select
        wda.Selection.TypeText ("+")
   
    End If
'
    If sparkasse Then
        wdd.Bookmarks("BankSEPA").Select
        wda.Selection.TypeText ("+")
        wdd.Bookmarks("BankSparkasse").Select
        wda.Selection.TypeText ("-")
    Else
        wdd.Bookmarks("BankSEPA").Select
        wda.Selection.TypeText ("-")
        wdd.Bookmarks("BankSparkasse").Select
        wda.Selection.TypeText ("+")
    End If

    
    If fotos Then
        wdd.Bookmarks("Fotos").Select
        wda.Selection.TypeText ("+")
    Else
        wdd.Bookmarks("Fotos").Select
        wda.Selection.TypeText ("-")
    End If
    
    If videos Then
        wdd.Bookmarks("Videos").Select
        wda.Selection.TypeText ("+")
    Else
        wdd.Bookmarks("Videos").Select
        wda.Selection.TypeText ("-")
    End If

    If ton Then
        wdd.Bookmarks("Ton").Select
        wda.Selection.TypeText ("+")
    Else
        wdd.Bookmarks("Ton").Select
        wda.Selection.TypeText ("-")
    End If
    
    If kunst Then
        wdd.Bookmarks("kWerke").Select
        wda.Selection.TypeText ("+")
    Else
        wdd.Bookmarks("kWerke").Select
        wda.Selection.TypeText ("-")
    End If
    
    If aufsat Then
        wdd.Bookmarks("Aufsetze").Select
        wda.Selection.TypeText ("+")
    Else
        wdd.Bookmarks("Aufsetze").Select
        wda.Selection.TypeText ("-")
    End If

    If website Then
        wdd.Bookmarks("Internet").Select
        wda.Selection.TypeText ("+")
    Else
        wdd.Bookmarks("Internet").Select
        wda.Selection.TypeText ("-")
    End If

    If socmed Then
        wdd.Bookmarks("SocMedia").Select
        wda.Selection.TypeText ("+")
    Else
        wdd.Bookmarks("SocMedia").Select
        wda.Selection.TypeText ("-")
    End If

    If raum Then
        wdd.Bookmarks("Raum").Select
        wda.Selection.TypeText ("+")
    Else
        wdd.Bookmarks("Raum").Select
        wda.Selection.TypeText ("-")
    End If

    If veranst Then
        wdd.Bookmarks("Veranstalt").Select
        wda.Selection.TypeText ("+")
    Else
        wdd.Bookmarks("Veranstalt").Select
        wda.Selection.TypeText ("-")
    End If
    
    If mitname Then
        wdd.Bookmarks("MitName").Select
        wda.Selection.TypeText ("+")
        wdd.Bookmarks("OneName").Select
        wda.Selection.TypeText ("-")
   Else
        wdd.Bookmarks("MitName").Select
        wda.Selection.TypeText ("-")
        wdd.Bookmarks("OneName").Select
        wda.Selection.TypeText ("+")
    End If
    
   
    NameElt = Replace(NameElt, "ö", "o")
    NameElt = Replace(NameElt, "ü", "u")
    NameElt = Replace(NameElt, "ä", "a")
    
'    DokumNm = "\VertJ_1" & Right(numExcel, 4) & "_" & Left(nameElt, 6) & Left(vornameElt, 6) & ".docx"
'    wdd.SaveAs ThisWorkbook.Path & "\Vertrage" & DokumNm

    wdd.SaveAs2 "C:\\Vertrage" & "\VertM_1" & Right(numExcel, 4) & "_" & Left(NameElt, 6) & Left(vornameElt, 6) & ".pdf", _
    FileFormat:=wdExportFormatPDF
    
    wdd.Close False

 End Sub

