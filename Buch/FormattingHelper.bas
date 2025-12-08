Attribute VB_Name = "FormattingHelper"
Option Explicit

' Apply borders to a specified row in the Kosten worksheet
Public Sub ApplyBorders(rowNumber As Integer)
    With Worksheets("Kosten")
        .Range(.Cells(rowNumber, 4), .Cells(rowNumber, 16)).Select
        
        With Selection.Borders(xlDiagonalDown)
            .LineStyle = xlNone
        End With
        
        With Selection.Borders(xlDiagonalUp)
            .LineStyle = xlNone
        End With
        
        With Selection.Borders(xlEdgeLeft)
            .LineStyle = xlContinuous
            .colorIndex = 0
            .TintAndShade = 0
            .Weight = xlThin
        End With
        
        With Selection.Borders(xlEdgeTop)
            .LineStyle = xlContinuous
            .colorIndex = 0
            .TintAndShade = 0
            .Weight = xlThin
        End With
        
        With Selection.Borders(xlEdgeBottom)
            .LineStyle = xlContinuous
            .colorIndex = 0
            .TintAndShade = 0
            .Weight = xlThin
        End With
        
        With Selection.Borders(xlEdgeRight)
            .LineStyle = xlContinuous
            .colorIndex = 0
            .TintAndShade = 0
            .Weight = xlThin
        End With
        
        With Selection.Borders(xlInsideVertical)
            .LineStyle = xlContinuous
            .colorIndex = 0
            .TintAndShade = 0
            .Weight = xlThin
        End With
        
        With Selection.Borders(xlInsideHorizontal)
            .LineStyle = xlContinuous
            .colorIndex = 0
            .TintAndShade = 0
            .Weight = xlThin
        End With
    End With
End Sub

' Set cell formatting for financial data
Public Sub FormatFinancialCell(cellRef As Range, Optional decimalPlaces As Integer = 2)
    With cellRef
        .NumberFormat = "0." & String(decimalPlaces, "0")
    End With
End Sub

' Format a cell as a header
Public Sub FormatHeaderCell(cellRef As Range, Optional bgColorIndex As Integer = -1)
    With cellRef
        .Font.Bold = True
        If bgColorIndex > 0 Then
            .Interior.colorIndex = bgColorIndex
        End If
    End With
End Sub

' Apply color highlighting to cells based on condition
Public Sub ApplyConditionalFormatting(rangeRef As Range, condition As String, colorIndex As Integer)
    rangeRef.FormatConditions.Add Type:=xlExpression, Formula1:=condition
    rangeRef.FormatConditions(rangeRef.FormatConditions.Count).Interior.colorIndex = colorIndex
End Sub
