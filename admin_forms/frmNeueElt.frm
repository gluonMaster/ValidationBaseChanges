VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} frmNeueElt 
   Caption         =   "Neue Eltern-Kinder"
   ClientHeight    =   13728
   ClientLeft      =   48
   ClientTop       =   372
   ClientWidth     =   21768
   OleObjectBlob   =   "frmNeueElt.frx":0000
   StartUpPosition =   1  'CenterOwner
End
Attribute VB_Name = "frmNeueElt"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False




Dim nmon As Integer
Dim kindCount As Integer
Dim letzte As Integer, paket As Integer
Dim neuNr As String
Dim EltName(500) As String, EltNr(500) As String
Dim EltKode As String, EName As String
Dim hinz As Integer
Dim Adres As String, Teleph As String, Handy As String, eml As String
Dim B As Boolean
Dim s_Pfad As String
' ng - neue ganz   AI


' Checks if given FamilyID does not exist yet in Kartei column A
Private Function IsFamilyIdUnique(ByVal familyId As String) As Boolean
    Dim ws As Worksheet
    Dim lastRow As Long
    Dim i As Long
    Dim cellValue As String
    
    Set ws = ThisWorkbook.Worksheets("Kartei")
    lastRow = ws.Cells(ws.Rows.Count, 1).End(xlUp).Row
    
    If lastRow < 3 Then
        IsFamilyIdUnique = True
        Exit Function
    End If
    
    For i = 3 To lastRow
        cellValue = Trim(CStr(ws.Cells(i, 1).Value))
        If StrComp(cellValue, Trim(familyId), vbTextCompare) = 0 Then
            IsFamilyIdUnique = False
            Exit Function
        End If
    Next i
    
    IsFamilyIdUnique = True
End Function

Private Sub btnOKelt_Click()
  Dim famId As String
  
  If Me.ComboBox31.text = "" Then
    MsgBox "Monat??????"
    Exit Sub
  End If
  
  famId = Trim(Me.txtNummer.text)
  If famId <> "" Then
    If Not IsFamilyIdUnique(famId) Then
      MsgBox "Diese FamilyID existiert bereits im Blatt Kartei. Bitte Formular schliessen und neue Nummer verwenden.", _
             vbExclamation, "FamilyID bereits vorhanden"
      Exit Sub
    End If
  End If
  
  If Me.obtnNH.value Then
    frmKinderNeueNH.Show
  Else
    frmKinderNeue.Show
  End If
End Sub


Private Sub obtnJahr_Click()
  If Me.obtnJahr.value Then
     Me.framKindFotoVideo.Visible = True
     Me.FramAufnamen.Visible = True
     Me.framOnlineUnter.Visible = True
  Else
     Me.framKindFotoVideo.Visible = False
     Me.FramAufnamen.Visible = False
     Me.framOnlineUnter.Visible = False
  End If
End Sub

Private Sub obtnMonat_Click()
  If Me.obtnMonat.value Then
     Me.framKindFotoVideo.Visible = True
     Me.FramAufnamen.Visible = True
     Me.framOnlineUnter.Visible = True
  Else
     Me.framKindFotoVideo.Visible = False
     Me.FramAufnamen.Visible = False
     Me.framOnlineUnter.Visible = False
  End If
  
End Sub

Private Sub obtnNH_Click()
  If Me.obtnNH.value Then
     Me.framKindFotoVideo.Visible = False
     Me.FramAufnamen.Visible = False
     Me.framOnlineUnter.Visible = False
  Else
     Me.framKindFotoVideo.Visible = True
     Me.FramAufnamen.Visible = True
     Me.framOnlineUnter.Visible = True
  End If

End Sub

Private Sub UserForm_Activate()
 Dim i As Integer, j As Integer
 Dim p As Integer
 
 s_Pfad = ThisWorkbook.path
 kindCount = 1
' Me.Height = 300
 ComboBox31.RowSource = "[" & ThisWorkbook.name & "]PlusBText!J2:J13"
 For i = 3 To 3000
   If Worksheets("Kartei").Range("A" & i) = "" Then
     letzte = i - 1
     i = 10000
   End If
 Next i
' Worksheets("DienstTab").Range("A1:B5000").Clear
' Worksheets("Kartei").Range("A3", "A" & Letzte).Copy Destination:=Worksheets("DienstTab").Cells(3, 1).End(xlUp).Offset(1, 0)
' Worksheets("DienstTab").Range("A2", "A" & Letzte).Sort Key1:=Worksheets("DienstTab").Range("A2"), Order1:=xlAscending, Header:=xlGuess, _
'     OrderCustom:=1, MatchCase:=False, Orientation:=xlTopToBottom
'
' For i = 2 To Letzte - 1
'   If VBA.Mid(Worksheets("DienstTab").Range("A" & i), 1, 2) = "" Then
'     Exit For
'   End If
' Next i
 neuNr = VBA.Mid(Worksheets("LetzteNummer").Range("A1"), 4, 4)
 p = VBA.CInt(neuNr) + 1
 neuNr = "1. " & p
 Me.txtNummer.text = neuNr
 Worksheets("LetzteNummer").Range("A1") = neuNr
 ActiveWorkbook.Save
End Sub





