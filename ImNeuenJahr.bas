Attribute VB_Name = "ImNeuenJahr"
'************************* ImNeuenJahr ******************************
Dim arrKonto(1000 To 2500) As Integer
Sub ImNeuenJahr()
  Dim i As Integer, j As Integer, p1 As Integer
  Dim tabsName As String
  
  
   For i = 1000 To 2500
     arrKonto(i) = 1
   Next i
   
   tabsName = "DienstTab"
   i = 3
   While Worksheets(tabsName).Cells(i, "A") <> ""
      If Worksheets(tabsName).Cells(i, "D") <> " Zahlung" Then
        Worksheets(tabsName).Cells(i, "J") = Worksheets(tabsName).Cells(i, "O")
        Worksheets(tabsName).Cells(i, "O") = ""
        Worksheets(tabsName).Cells(i, "K") = Worksheets(tabsName).Cells(i, "P")
        Worksheets(tabsName).Cells(i, "P") = ""
        Worksheets(tabsName).Cells(i, "L") = Worksheets(tabsName).Cells(i, "Q")
        Worksheets(tabsName).Cells(i, "Q") = ""
'        Worksheets(tabsName).Cells(i, "M") = Worksheets(tabsName).Cells(i, "AP")
        Worksheets(tabsName).Cells(i, "M") = Worksheets(tabsName).Cells(i, "R")
        For j = 21 To 27
'          Worksheets(tabsName).Cells(i, j) = Worksheets(tabsName).Cells(i, "AP")
          Worksheets(tabsName).Cells(i, j) = Worksheets(tabsName).Cells(i, "R")
        Next j
        Worksheets(tabsName).Cells(i, "R") = ""
        
        For j = 28 To 32
          Worksheets(tabsName).Cells(i, j) = 0
        Next j
        Worksheets(tabsName).Cells(i, "AK") = ""
        Worksheets(tabsName).Cells(i, "AL") = ""
        Worksheets(tabsName).Cells(i, "AM") = ""
        Worksheets(tabsName).Cells(i, "AN") = ""
        Worksheets(tabsName).Cells(i, "AO") = ""
        Worksheets(tabsName).Cells(i, "AK") = Worksheets(tabsName).Cells(i, "AP")
        Worksheets(tabsName).Cells(i, "AP") = ""
        Worksheets(tabsName).Cells(i, "AL") = Worksheets(tabsName).Cells(i, "AQ")
        Worksheets(tabsName).Cells(i, "AQ") = ""
        Worksheets(tabsName).Cells(i, "AM") = Worksheets(tabsName).Cells(i, "AR")
        Worksheets(tabsName).Cells(i, "AR") = ""
        Worksheets(tabsName).Cells(i, "AN") = Worksheets(tabsName).Cells(i, "AS")
        Worksheets(tabsName).Cells(i, "AS") = ""
        Worksheets(tabsName).Cells(i, "AO") = Worksheets(tabsName).Cells(i, "AT")
        Worksheets(tabsName).Cells(i, "AT") = ""

      End If
      
      i = i + 1
   Wend
'   KN_Zuweisen (i - 1)
   Call DateiUebertragen("DienstTab", "Kartei", "<>KN", i - 1)

End Sub

Sub KN_Zuweisen(LetzteX As Integer)
 Dim i As Integer, p1 As Integer
 
 For i = 3 To LetzteX
   If Worksheets("DienstTab").Range("D" & i) = " Zahlung" Then
      p1 = CInt(VBA.Right(Worksheets("DienstTab").Range("A" & i), 4))
      If arrKonto(p1) = 1 Then
         Worksheets("DienstTab").Range("T" & i) = "KN"
      End If
    End If
  Next i
 
End Sub

Sub DateiUebertragen(WorkShQuelle As String, WorkShZiel As String, FiltrName As String, LetzteX As Integer)
  Dim rng As Range
  
  Worksheets(WorkShQuelle).Range("A3:AV" & LetzteX).AutoFilter Field:=20, Criteria1:=FiltrName _
        , Operator:=xlAnd
  Set rng = Worksheets(WorkShQuelle).AutoFilter.Range
  rng.Offset(1, 0).Resize(rng.rows.count - 1).Copy Destination:=Worksheets(WorkShZiel).Range("A3")
  Worksheets(WorkShQuelle).Range("A2:AZ" & LetzteX).AutoFilter Field:=20
 
End Sub
'************************* ImNeuenJahr ******************************


