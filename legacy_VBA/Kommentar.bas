Attribute VB_Name = "Kommentar"
Option Explicit

' Function that builds the string to be inserted into column 52
Public Function BuildCommentString(ByVal targetCell As Range) As String
    Dim currentDate As String
    currentDate = Format(Date, "dd.mm.yyyy")
    Dim textVar As String

    ' Example input
    textVar = targetCell.Comment.text

    ' Remove line breaks (vbCr, vbLf, and vbCrLf)
    textVar = Replace(textVar, vbCrLf, "") ' Remove combined carriage return and line feed
    textVar = Replace(textVar, vbCr, "")   ' Remove carriage return
    textVar = Replace(textVar, vbLf, "")   ' Remove line feed

    ' Remove slashes
    textVar = Replace(textVar, "/", " ")
    textVar = Replace(textVar, "-", " ")

    
    ' Build the required string:
    ' "Mnt.col-20: War(); Ist(X). /comentar/ DD || "
    Dim nRow As Long
    nRow = targetCell.Column
    BuildCommentString = "Mnt." & (nRow - 20) & ": War(); " & _
                        "Ist(" & targetCell.value & "). /" & _
                         textVar & "/ " & currentDate & " || "
End Function

' Main procedure that processes the comments in columns U:AF
Public Sub ProcessComments()
    Dim ws As Worksheet
    Dim C As Range
    Dim newText As String
    
    Set ws = ActiveSheet
    
    ' Loop through each cell in the range of columns U:AF
    For Each C In ws.Range("U:AF")
        ' Check if the cell has a comment
        If Not C.Comment Is Nothing Then
            
            ' Build the string to be inserted
            newText = BuildCommentString(C)
            
            ' Add (append) the new text to cell (same row, column 52)
            ws.Cells(C.row, 52).value = ws.Cells(C.row, 52).value & newText
            
            ' Remove the comment from the cell
            C.Comment.Delete
        End If
    Next C
End Sub




