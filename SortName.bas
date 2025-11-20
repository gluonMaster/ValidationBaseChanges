Attribute VB_Name = "SortName"
Public Sub SortNameZ()
Attribute SortNameZ.VB_ProcData.VB_Invoke_Func = " \n14"
 Dim i As Integer, letzte As Integer
 For i = 3 To 3000
   If Worksheets("Kartei").Range("A" & i) = "" Then
     letzte = i - 1
     i = 10000
   End If
 Next i
 
  Range("A3:AZ" & letzte).RowHeight = 16
 
 Worksheets("Kartei").Sort.SortFields.Clear
 Worksheets("Kartei").Sort.SortFields.Add key:=Range("B3:B" & letzte), _
   SortOn:=xlSortOnValues, Order:=xlAscending, DataOption:=xlSortNormal
 Worksheets("Kartei").Sort.SortFields.Add key:=Range("A3:A" & letzte), _
   SortOn:=xlSortOnValues, Order:=xlAscending, DataOption:=xlSortNormal
 Worksheets("Kartei").Sort.SortFields.Add key:=Range("D3:D" & letzte), _
   SortOn:=xlSortOnValues, Order:=xlAscending, DataOption:=xlSortNormal
 With Worksheets("Kartei").Sort
    .SetRange Range("A2:AZ" & letzte)
    .Header = xlYes
    .MatchCase = False
    .Orientation = xlTopToBottom
    .SortMethod = xlPinYin
    .Apply
 End With
 
End Sub

Sub SortNummer()
 Dim i As Integer, letzte As Integer
 For i = 3 To 3000
   If Worksheets("Kartei").Range("A" & i) = "" Then
     letzte = i - 1
     i = 10000
   End If
 Next i
 
  Range("A3:AU" & letzte).RowHeight = 16
 
 Worksheets("Kartei").Sort.SortFields.Clear
 Worksheets("Kartei").Sort.SortFields.Add key:=Range("A3:A" & letzte), _
   SortOn:=xlSortOnValues, Order:=xlAscending, DataOption:=xlSortNormal
 Worksheets("Kartei").Sort.SortFields.Add key:=Range("D3:D" & letzte), _
   SortOn:=xlSortOnValues, Order:=xlAscending, DataOption:=xlSortNormal
 With Worksheets("Kartei").Sort
    .SetRange Range("A2:AZ" & letzte)
    .Header = xlYes
    .MatchCase = False
    .Orientation = xlTopToBottom
    .SortMethod = xlPinYin
    .Apply
 End With

End Sub
