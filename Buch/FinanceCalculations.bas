Attribute VB_Name = "FinanceCalculations"
Option Explicit

' Calculate and display tariff summaries by month
Public Function SumTarifMonat(TabNameX As String, vol As Boolean, Xmonat As Integer, nameEltern As String, kontoNr As String, ByRef RowFurTabelle As Integer) As Double
    '  U - January,  AB - August
    '  AC - September, AF - December
    '  J - Gruppe I, O - Gruppe II
    '  Xmonat - last month of school year

    Dim i As Integer, j As Integer
    Dim sumTarifM As Double
    Dim letzteMonat As Integer, rowTotal As Integer
    Dim sumTarifJahr As Double
    Dim lastRow As Integer
    Dim is2023 As Boolean
  
    sumTarifJahr = 0
    is2023 = (TabNameX = "2023")
  
    If vol Then
        RowFurTabelle = 5
        letzteMonat = 12
    Else
        RowFurTabelle = RowFurTabelle + 3
    
        ' Determine last month based on whether TabNameX is the current year
        ' For current year: use current month; for past years: use full year (12)
        If CInt(TabNameX) = Year(Date) Then
            letzteMonat = Month(Date)
        Else
            letzteMonat = 12
        End If
    End If
  
    Worksheets("Kosten").Cells(RowFurTabelle, 2) = TabNameX
    RowFurTabelle = RowFurTabelle + 1
    Worksheets("Kosten").Cells(RowFurTabelle, 2) = nameEltern
    RowFurTabelle = RowFurTabelle + 2
  
    rowTotal = RowFurTabelle
 
    ' Set month headers based on year
    If is2023 Then
        ' 2023: Original layout - months 1-8 in columns D-K, "Unterricht" in L, months 9-12 in M-P
        For i = 1 To Xmonat
            Worksheets("Kosten").Cells(RowFurTabelle, i + 3) = GetMonthName(i)
            FormattingHelper.ApplyBorders RowFurTabelle
        Next i
        
        Worksheets("Kosten").Cells(RowFurTabelle, i + 4) = "Unterricht"
        
        For i = Xmonat + 1 To 12
            Worksheets("Kosten").Cells(RowFurTabelle, i + 4) = GetMonthName(i)
            FormattingHelper.ApplyBorders RowFurTabelle
        Next i
    Else
        ' 2024/2025: New layout - months 1-7 in columns D-J, "Unterricht" in K, month 8 in L, months 9-12 in M-P
        For i = 1 To Xmonat - 1  ' Months 1-7 (Jan-Jul)
            Worksheets("Kosten").Cells(RowFurTabelle, i + 3) = GetMonthName(i)
            FormattingHelper.ApplyBorders RowFurTabelle
        Next i
        
        ' Put "Unterricht" in column K (position 11)
        Worksheets("Kosten").Cells(RowFurTabelle, 11) = "Unterricht"
        
        ' Put August in column L (position 12)
        Worksheets("Kosten").Cells(RowFurTabelle, 12) = GetMonthName(Xmonat)
        FormattingHelper.ApplyBorders RowFurTabelle
        
        ' Put remaining months (Sep-Dec) in columns M-P
        For i = Xmonat + 1 To 12
            Worksheets("Kosten").Cells(RowFurTabelle, i + 4) = GetMonthName(i)
            FormattingHelper.ApplyBorders RowFurTabelle
        Next i
    End If
    
    FormattingHelper.ApplyBorders RowFurTabelle

    RowFurTabelle = RowFurTabelle + 1

    lastRow = DataHandler.FindLastRow("DienstTab")
    For i = 2 To lastRow
        Worksheets("Kosten").Cells(RowFurTabelle, 2) = Worksheets("DienstTab").Cells(i, 4)
        
        ' Set data based on year layout
        If is2023 Then
            ' 2023: Original data layout
            Worksheets("Kosten").Cells(RowFurTabelle, 3) = Worksheets("DienstTab").Cells(i, "J")
            
            If letzteMonat > Xmonat Then
                For j = 1 To Xmonat
                    Worksheets("Kosten").Cells(RowFurTabelle, 3 + j) = Worksheets("DienstTab").Cells(i, j + 20)
                    Worksheets("Kosten").Cells(RowFurTabelle, 3 + j).NumberFormat = "0.00"
                Next j
                
                Worksheets("Kosten").Cells(RowFurTabelle, 3 + j) = Worksheets("DienstTab").Cells(i, "O")
                
                For j = Xmonat + 1 To letzteMonat
                    Worksheets("Kosten").Cells(RowFurTabelle, 4 + j) = Worksheets("DienstTab").Cells(i, j + 20)
                    Worksheets("Kosten").Cells(RowFurTabelle, 4 + j).NumberFormat = "0.00"
                Next j
            Else
                For j = 1 To letzteMonat
                    Worksheets("Kosten").Cells(RowFurTabelle, 3 + j) = Worksheets("DienstTab").Cells(i, j + 20)
                    Worksheets("Kosten").Cells(RowFurTabelle, 3 + j).NumberFormat = "0.00"
                Next j
            End If
        Else
            ' 2024/2025: New data layout
            ' Put group info from column O in column K
            Worksheets("Kosten").Cells(RowFurTabelle, 11) = Worksheets("DienstTab").Cells(i, "O")
            
            ' Put group info from column J in column C
            Worksheets("Kosten").Cells(RowFurTabelle, 3) = Worksheets("DienstTab").Cells(i, "J")
            
            If letzteMonat > Xmonat Then
                ' Months 1-7 in columns D-J
                For j = 1 To Xmonat - 1
                    Worksheets("Kosten").Cells(RowFurTabelle, 3 + j) = Worksheets("DienstTab").Cells(i, j + 20)
                    Worksheets("Kosten").Cells(RowFurTabelle, 3 + j).NumberFormat = "0.00"
                Next j
                
                ' Month 8 (August) in column L
                Worksheets("Kosten").Cells(RowFurTabelle, 12) = Worksheets("DienstTab").Cells(i, Xmonat + 20)
                Worksheets("Kosten").Cells(RowFurTabelle, 12).NumberFormat = "0.00"
                
                ' Months 9-12 in columns M-P
                For j = Xmonat + 1 To letzteMonat
                    Worksheets("Kosten").Cells(RowFurTabelle, 4 + j) = Worksheets("DienstTab").Cells(i, j + 20)
                    Worksheets("Kosten").Cells(RowFurTabelle, 4 + j).NumberFormat = "0.00"
                Next j
            Else
                ' Handle case where last month <= Xmonat
                For j = 1 To Xmonat - 1
                    Worksheets("Kosten").Cells(RowFurTabelle, 3 + j) = Worksheets("DienstTab").Cells(i, j + 20)
                    Worksheets("Kosten").Cells(RowFurTabelle, 3 + j).NumberFormat = "0.00"
                Next j
                
                If letzteMonat = Xmonat Then
                    ' Include August if it's the last month
                    Worksheets("Kosten").Cells(RowFurTabelle, 12) = Worksheets("DienstTab").Cells(i, Xmonat + 20)
                    Worksheets("Kosten").Cells(RowFurTabelle, 12).NumberFormat = "0.00"
                End If
            End If
        End If
        
        RowFurTabelle = RowFurTabelle + 1
    Next i
  
    Cells(RowFurTabelle + 1, 3) = "Summe"
    
    ' Calculate sums based on layout
    If is2023 Then
        ' 2023: Original sum calculation
        For j = 1 To Xmonat
            Cells(RowFurTabelle + 1, 3 + j).NumberFormat = "0.00"
            Cells(RowFurTabelle + 1, 3 + j).FormulaR1C1 = "=SUM(R[-1]C:R[-" & lastRow & "]C)"
        Next j
        
        For j = Xmonat + 1 To 12
            Cells(RowFurTabelle + 1, 4 + j).NumberFormat = "0.00"
            Cells(RowFurTabelle + 1, 4 + j).FormulaR1C1 = "=SUM(R[-1]C:R[-" & lastRow & "]C)"
        Next j
    Else
        ' 2024/2025: New sum calculation
        ' Sum for months 1-7 in columns D-J
        For j = 1 To Xmonat - 1
            Cells(RowFurTabelle + 1, 3 + j).NumberFormat = "0.00"
            Cells(RowFurTabelle + 1, 3 + j).FormulaR1C1 = "=SUM(R[-1]C:R[-" & lastRow & "]C)"
        Next j
        
        ' Sum for August in column L
        Cells(RowFurTabelle + 1, 12).NumberFormat = "0.00"
        Cells(RowFurTabelle + 1, 12).FormulaR1C1 = "=SUM(R[-1]C:R[-" & lastRow & "]C)"
        
        ' Sum for months 9-12 in columns M-P
        For j = Xmonat + 1 To 12
            Cells(RowFurTabelle + 1, 4 + j).NumberFormat = "0.00"
            Cells(RowFurTabelle + 1, 4 + j).FormulaR1C1 = "=SUM(R[-1]C:R[-" & lastRow & "]C)"
        Next j
    End If
  
    For j = 4 To 16
        Cells(RowFurTabelle + 1, j).NumberFormat = "0.00"
        sumTarifJahr = sumTarifJahr + Cells(RowFurTabelle + 1, j)
    Next j
  
    RowFurTabelle = RowFurTabelle + 1
    Cells(RowFurTabelle + 1, 3) = "Bezahlt"
  
    For j = 4 To 16
        Cells(RowFurTabelle + 1, j).NumberFormat = "0.00"
    Next j

    Cells(RowFurTabelle + 1, 17).NumberFormat = "0.00"
    Cells(RowFurTabelle + 1, 17).FormulaR1C1 = "=SUM(RC[-13]:RC[-1])"
    Cells(RowFurTabelle, 19).NumberFormat = "0.00"
  
    Cells(RowFurTabelle, 17).FormulaR1C1 = "=SUM(RC[-13]:RC[-1])"
  
    RowFurTabelle = RowFurTabelle + 2
    Cells(RowFurTabelle, 3) = "Differenz"
    Cells(RowFurTabelle, 19).NumberFormat = "0.00"
    Cells(RowFurTabelle, 19).FormulaR1C1 = "=R[-2]C[-2]-R[-1]C[-2]"

    SumTarifMonat = sumTarifJahr
End Function

' Get month name abbreviation
Public Function GetMonthName(monthNumber As Integer) As String
    Dim monthNames As Variant
    monthNames = Array("Jan.", "Feb.", "Mrz.", "Apr.", "Mai.", "Jun.", _
                       "Jul.", "Aug.", "Spt.", "Okt.", "Nov.", "Dez.")
    
    If monthNumber >= 1 And monthNumber <= 12 Then
        GetMonthName = monthNames(monthNumber - 1)
    Else
        GetMonthName = "Unknown"
    End If
End Function

