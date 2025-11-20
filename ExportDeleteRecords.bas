Attribute VB_Name = "ExportDeleteRecords"
Option Explicit

Public Sub DeleteSelectedRecordsFromKartei()
    ' This procedure deletes selected rows from the "Kartei" sheet
    ' and also removes matching records from the Access database.
    ' It now confirms the deletion with the user and handles partial row selections properly.
    
    Dim ws As Worksheet
    Set ws = ThisWorkbook.Worksheets("Kartei")
    
'    ws.Unprotect password:="1212"
'    ws.Cells.Locked = False
    
    Dim rngSelection As Range
    Dim rngRows As Range
    Dim singleArea As Range
    
    ' Intersect the current selection with the used range to avoid unnecessary references
    On Error Resume Next
    Set rngSelection = Intersect(Selection, ws.UsedRange)
    On Error GoTo 0
    
    If rngSelection Is Nothing Then
        MsgBox "No valid selection in the 'Kartei' sheet.", vbExclamation
        Exit Sub
    End If
    
    ' Ask for confirmation before deletion
    If MsgBox("Wichtig: Bevor Sie Datensätze loschen, synchronisieren Sie die Datenbank (Taste Zu Base)." _
        & vbCrLf & "Sind Sie sicher, dass Sie die markierten Datensatze loschen wollen?" & vbCrLf & _
              "Diese Aktion kann nicht ruckgangig gemacht werden.", _
               vbYesNo + vbQuestion, "Confirm Deletion") = vbNo Then
        Exit Sub
    End If
    
    ' Convert partial row selections to entire rows
    For Each singleArea In rngSelection.Areas
        If rngRows Is Nothing Then
            Set rngRows = singleArea.EntireRow
        Else
            Set rngRows = Union(rngRows, singleArea.EntireRow)
        End If
    Next singleArea
    
    If rngRows Is Nothing Then
        MsgBox "No valid rows could be determined for deletion.", vbExclamation
        Exit Sub
    End If
    
    Dim colIDs As Collection
    Set colIDs = New Collection
    
    ' Collect IDs from the 48th column for each row
    Dim rowRange As Range
    For Each rowRange In rngRows.Rows
        Dim rowID As Variant
        rowID = rowRange.Cells(1, 48).value  ' the 48th column in "Kartei"
        
        If IsNumeric(rowID) Then
            If CLng(rowID) > 0 Then
                colIDs.Add CLng(rowID)
            End If
        End If
    Next rowRange
    
    ' Delete from Access by collected IDs
    If colIDs.count > 0 Then
        DeleteRecordsFromAccess colIDs
    End If
    
    ' Now remove these entire rows from the worksheet
    ' Sort rows top-to-bottom (using a Range.Sort or something) is not necessary
    ' because we can delete the entire range in one go if desired.
    rngRows.Delete
    
    ' Rebuild Kartei_Original
    RebuildKarteiOriginal
'    ' Call the initialization for protection
'    Call InitializeWorkbookPubl
    
    MsgBox "Selected records have been deleted from both the 'Kartei' sheet and the database.", vbInformation
End Sub

Public Sub DeleteRecordsFromAccess(ByVal colIDs As Collection)
    ' This procedure removes rows from the Access table "tblKartei" by a collection of IDs.
    ' It uses a DAO.Recordset approach similar to the existing write routines.
    
    Dim engine As DAO.DBEngine
    Set engine = New DAO.DBEngine
    
    Dim wsDao As DAO.Workspace
    Set wsDao = engine.Workspaces(0)
    
    Dim dbPath As String
    dbPath = ThisWorkbook.Worksheets("Kartei").Range("X1").value & "\Alarm\KindElternDaten_25_front.accdb"
    
    Dim db As DAO.Database
    Set db = wsDao.OpenDatabase(dbPath)
    
    Dim rs As DAO.Recordset
    Set rs = db.OpenRecordset("tblKartei", dbOpenDynaset, dbSeeChanges)
    
    wsDao.BeginTrans
    
    Dim varID As Variant
    For Each varID In colIDs
        rs.FindFirst "ID=" & CLng(varID)
        If Not rs.NoMatch Then
            rs.Delete
        End If
    Next varID
    
    wsDao.CommitTrans
    
    rs.Close
    db.Close
End Sub




