Attribute VB_Name = "Export_DeclinedHelpers"
'==========================
'   Declined Records Helper Functions
'   Utility functions for comparing, copying, and managing declined records
'==========================
Option Explicit

' ========================================
' Dictionary Building Functions
' ========================================

Public Sub BuildOverviewDictionary(ByVal ws As Worksheet, ByVal dict As Object)
    ' Builds dictionary of DeclinedOverview data: key=ID, value=array of values
    ' Array structure: (0)=ID, (1-20)=key fields and history
    
    Dim lastRow As Long
    lastRow = ws.Cells(ws.Rows.Count, 1).End(xlUp).Row
    
    If lastRow < 2 Then Exit Sub
    
    Dim r As Long
    For r = 2 To lastRow
        Dim idVal As Variant
        idVal = ws.Cells(r, 1).value
        
        If Not IsEmpty(idVal) And IsNumeric(idVal) Then
            Dim idLong As Long
            idLong = CLng(idVal)
            
            ' Create array matching DeclinedOverview structure
            Dim recData(0 To 20) As Variant
            recData(0) = idLong
            recData(1) = NzStr(ws.Cells(r, 2).value)  ' A (Value1)
            recData(2) = NzStr(ws.Cells(r, 3).value)  ' B (Value2)
            recData(3) = NzStr(ws.Cells(r, 4).value)  ' C (Value3)
            recData(4) = NzStr(ws.Cells(r, 5).value)  ' F (Value6)
            recData(5) = NzStr(ws.Cells(r, 6).value)  ' J (Value10)
            recData(6) = NzStr(ws.Cells(r, 7).value)  ' O (Value15)
            
            ' Months 1-12 (columns H-S = Value21-32)
            Dim m As Integer
            For m = 1 To 12
                recData(6 + m) = NzStr(ws.Cells(r, 7 + m).value)
            Next m
            
            recData(19) = NzStr(ws.Cells(r, 21).value)  ' AZ (Value52)
            recData(20) = 0  ' DeclCount (not used in comparison)
            
            dict(idLong) = recData
        End If
    Next r
End Sub

Public Sub BuildDeclDictionary(ByVal db As DAO.Database, ByVal dict As Object)
    ' Builds dictionary of decl_tblKartei data: key=ID, value=array of values
    
    On Error Resume Next
    Dim rs As DAO.Recordset
    Set rs = db.OpenRecordset("SELECT * FROM decl_tblKartei", dbOpenSnapshot)
    
    If Err.Number <> 0 Then
        On Error GoTo 0
        Exit Sub
    End If
    On Error GoTo 0
    
    If rs.EOF Then
        rs.Close
        Exit Sub
    End If
    
    Do While Not rs.EOF
        Dim recData(0 To 20) As Variant
        
        recData(0) = rs.Fields("ID").value
        recData(1) = NzStr(rs.Fields("Value1").value)
        recData(2) = NzStr(rs.Fields("Value2").value)
        recData(3) = NzStr(rs.Fields("Value3").value)
        recData(4) = NzStr(rs.Fields("Value6").value)
        recData(5) = NzStr(rs.Fields("Value10").value)
        recData(6) = NzStr(rs.Fields("Value15").value)
        
        Dim m As Integer
        For m = 1 To 12
            recData(6 + m) = NzStr(rs.Fields("Value" & (20 + m)).value)
        Next m
        
        recData(19) = NzStr(rs.Fields("Value52").value)
        recData(20) = 0
        
        dict(CLng(recData(0))) = recData
        rs.MoveNext
    Loop
    
    rs.Close
End Sub

' ========================================
' Comparison Functions
' ========================================

Public Function RecordsAreDifferent(ByVal rec1 As Variant, ByVal rec2 As Variant) As Boolean
    ' Compares two record arrays (from overview or decl dict)
    ' Returns True if any key field differs (excluding history AZ and DeclCount)
    
    Dim i As Integer
    For i = 1 To 18  ' Compare A, B, C, F, J, O, Months 1-12
        If CStr(rec1(i)) <> CStr(rec2(i)) Then
            RecordsAreDifferent = True
            Exit Function
        End If
    Next i
    
    RecordsAreDifferent = False
End Function

' ========================================
' Data Copy Functions
' ========================================

Public Sub CopyOverviewDataToKartei(ByVal ws As Worksheet, _
                                    ByVal karteiRow As Long, _
                                    ByVal overviewData As Variant)
    ' Copies data from DeclinedOverview array to Kartei sheet row
    ' Maps overview positions to correct Kartei columns
    
    ' overviewData structure:
    ' (0)=ID, (1)=A, (2)=B, (3)=C, (4)=F, (5)=J, (6)=O, (7-18)=Months, (19)=AZ
    
    ws.Cells(karteiRow, 1).value = overviewData(1)   ' A = Value1
    ws.Cells(karteiRow, 2).value = overviewData(2)   ' B = Value2
    ws.Cells(karteiRow, 3).value = overviewData(3)   ' C = Value3
    ws.Cells(karteiRow, 6).value = overviewData(4)   ' F = Value6 (Address)
    ws.Cells(karteiRow, 10).value = overviewData(5)  ' J = Value10 (Subject1)
    ws.Cells(karteiRow, 15).value = overviewData(6)  ' O = Value15 (Subject2)
    
    ' Months U-AF (columns 21-32 = Value21-32)
    Dim m As Integer
    For m = 1 To 12
        ws.Cells(karteiRow, 20 + m).value = overviewData(6 + m)
    Next m
    
    ' History AZ (column 52 = Value52)
    ws.Cells(karteiRow, 52).value = overviewData(19)
End Sub

' ========================================
' Table Management Functions
' ========================================

Public Sub EnsurePreTableExists(ByVal db As DAO.Database)
    ' Ensures pre_tblKartei exists, creates if not
    
    On Error Resume Next
    Dim tbl As DAO.TableDef
    For Each tbl In db.TableDefs
        If tbl.Name = "pre_tblKartei" Then
            On Error GoTo 0
            Exit Sub
        End If
    Next tbl
    On Error GoTo 0
    
    ' Create table
    Call CreatePreTable(db)
End Sub

Private Sub CreatePreTable(ByVal db As DAO.Database)
    ' Creates pre_tblKartei with same structure as tblKartei
    ' ID is regular Long (not AutoNumber) to preserve original IDs
    
    Dim tbl As DAO.TableDef
    Set tbl = db.CreateTableDef("pre_tblKartei")
    
    ' Add ID field (regular Long, not AutoNumber)
    Dim fld As DAO.Field
    Set fld = tbl.CreateField("ID", dbLong)
    tbl.Fields.Append fld
    
    ' Add Value1..Value51 fields (Text)
    Dim i As Long
    For i = 1 To 51
        Set fld = tbl.CreateField("Value" & i, dbText, 255)
        fld.AllowZeroLength = True  ' Allow empty strings
        tbl.Fields.Append fld
    Next i
    
    ' Add Value52 (Memo, for history)
    Set fld = tbl.CreateField("Value52", dbMemo)
    tbl.Fields.Append fld
    
    ' Add format fields
    For i = 1 To 51
        Set fld = tbl.CreateField("InteriorColor" & i, dbLong)
        tbl.Fields.Append fld
    Next i
    
    Set fld = tbl.CreateField("FontColor3", dbLong)
    tbl.Fields.Append fld
    
    Set fld = tbl.CreateField("FontColor18", dbLong)
    tbl.Fields.Append fld
    
    db.TableDefs.Append tbl
End Sub

' ========================================
' Utility Functions
' ========================================

Public Function NzStr(ByVal value As Variant) As String
    ' Converts null/empty to empty string
    If IsNull(value) Or IsEmpty(value) Then
        NzStr = ""
    Else
        NzStr = CStr(value)
    End If
End Function
