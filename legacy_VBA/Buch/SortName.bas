Attribute VB_Name = "SortName"
Sub SortNameZ()
Attribute SortNameZ.VB_ProcData.VB_Invoke_Func = " \n14"
 Dim i As Integer, Letzte As Integer
 For i = 3 To 3000
   If Worksheets("Kartei").Range("A" & i) = "" Then
     Letzte = i - 1
     i = 10000
   End If
 Next i
 
  Range("A3:AU" & Letzte).RowHeight = 16
 
 Worksheets("Kartei").Sort.SortFields.Clear
 Worksheets("Kartei").Sort.SortFields.Add Key:=Range("B3:B" & Letzte), _
   SortOn:=xlSortOnValues, Order:=xlAscending, DataOption:=xlSortNormal
 Worksheets("Kartei").Sort.SortFields.Add Key:=Range("D3:D" & Letzte), _
   SortOn:=xlSortOnValues, Order:=xlAscending, DataOption:=xlSortNormal
 With Worksheets("Kartei").Sort
    .SetRange Range("A2:AV" & Letzte)
    .Header = xlYes
    .MatchCase = False
    .Orientation = xlTopToBottom
    .SortMethod = xlPinYin
    .Apply
 End With
 
End Sub

Sub SortNummer()
 Dim i As Integer, Letzte As Integer
 For i = 3 To 3000
   If Worksheets("Kartei").Range("A" & i) = "" Then
     Letzte = i - 1
     i = 10000
   End If
 Next i
 
  Range("A3:AU" & Letzte).RowHeight = 16
 
 Worksheets("Kartei").Sort.SortFields.Clear
 Worksheets("Kartei").Sort.SortFields.Add Key:=Range("A3:A" & Letzte), _
   SortOn:=xlSortOnValues, Order:=xlAscending, DataOption:=xlSortNormal
 Worksheets("Kartei").Sort.SortFields.Add Key:=Range("D3:D" & Letzte), _
   SortOn:=xlSortOnValues, Order:=xlAscending, DataOption:=xlSortNormal
 With Worksheets("Kartei").Sort
    .SetRange Range("A2:AV" & Letzte)
    .Header = xlYes
    .MatchCase = False
    .Orientation = xlTopToBottom
    .SortMethod = xlPinYin
    .Apply
 End With

End Sub
