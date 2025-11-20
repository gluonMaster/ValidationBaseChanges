Attribute VB_Name = "ExportDataWork"
'==========================
'   Code Section: modAccessData
'==========================

Option Explicit

Public Function ReadAccessIntoDictionary_ID() As Scripting.Dictionary

    Dim dict As New Scripting.Dictionary
    Dim engine As DAO.DBEngine
    Set engine = New DAO.DBEngine
    
    Dim wsDao As DAO.Workspace
    Set wsDao = engine.Workspaces(0)
    
    Dim dbPath As String
    dbPath = ThisWorkbook.Worksheets("Kartei").Range("X1").value & "\Alarm\KindElternDaten_25_front.accdb"
    'dbPath = "d:\Work\Konst\2025\Kolibri\Alla\Release\KindElternDaten_24_front.accdb"
    
    Dim db As DAO.Database
    Set db = wsDao.OpenDatabase(dbPath)
    
    Dim rs As DAO.Recordset
    Dim sqlStr As String
    
    ' We'll just do SELECT * and assume the fields are "ID, Value1..Value51"
    ' then we'll map them carefully into arrRow(1,1..51).
    sqlStr = "SELECT * FROM tblKartei ORDER BY ID"
    
    Set rs = db.OpenRecordset(sqlStr, dbOpenDynaset)
    If rs.EOF Then
        rs.Close
        db.Close
        Set ReadAccessIntoDictionary_ID = dict
        Exit Function
    End If
    
    Do While Not rs.EOF
        Dim arrRow As Variant
        ReDim arrRow(1 To 1, 1 To 51)
        
        Dim c As Long
        ' We assume "Value1" => arrRow(1,1), "Value2" => arrRow(1,2), ... "Value51" => arrRow(1,51),
        ' and "ID" => arrRow(1,48).
        
        ' We'll fill each array column from the DB field "ValueX".
        ' But we do a small loop for columns 1..51, skipping ID logic.
        For c = 1 To 51
            If c = 48 Then
                ' c=48 => ID
                arrRow(1, 48) = NzToEmpty(rs.Fields("ID").value)
            Else
                ' "Value1..Value51" in the DB
                Dim fieldName As String
                fieldName = "Value" & c
                arrRow(1, c) = NzToEmpty(rs.Fields(fieldName).value)
            End If
        Next c
        
        Dim strID As String
        strID = CStr(arrRow(1, 48))  ' Convert ID to string as dictionary key
        
        If Not dict.exists(strID) Then
            dict.Add strID, arrRow
        End If
        
        rs.MoveNext
    Loop
    
    rs.Close
    db.Close
    
    Set ReadAccessIntoDictionary_ID = dict
End Function

Public Sub WriteDictionaryChangesToAccess_Recordset(ByVal dictLocal As Scripting.Dictionary, _
                                                    ByVal dictLocalFormats As Scripting.Dictionary, _
                                                    ByVal changedIDs As Collection)
    ' Simplest fix - just replace FindFirst with direct SELECT
    ' Uses same structure as original code but avoids FindFirst completely
    
    Dim engine As DAO.DBEngine
    Set engine = New DAO.DBEngine
    
    Dim wsDao As DAO.Workspace
    Set wsDao = engine.Workspaces(0)
    
    Dim dbPath As String
    dbPath = ThisWorkbook.Worksheets("Kartei").Range("X1").value & "\Alarm\KindElternDaten_25_front.accdb"
    
    Dim db As DAO.Database
    Set db = wsDao.OpenDatabase(dbPath)
    
    wsDao.BeginTrans
    
    Dim varID As Variant
    For Each varID In changedIDs
        If dictLocal.exists(varID) Then
            Dim arrRow As Variant
            Dim arrFormats As Variant
            arrRow = dictLocal(varID)
            arrFormats = dictLocalFormats(varID)
            
            Dim strID As String
            strID = CStr(arrRow(1, 48))
            
            If IsNumeric(strID) And Val(strID) > 0 Then
                Dim targetID As Long
                targetID = CLng(strID)
                
                ' CHANGED: Use SELECT instead of FindFirst
                Dim rsCheck As DAO.Recordset
                Set rsCheck = db.OpenRecordset("SELECT * FROM tblKartei WHERE ID = " & targetID, dbOpenDynaset)
                
                If rsCheck.EOF Then
                    ' Record not found - create new using separate recordset
                    rsCheck.Close
                    
                    Dim rsNew As DAO.Recordset
                    Set rsNew = db.OpenRecordset("tblKartei", dbOpenDynaset)
                    rsNew.AddNew
                    FillRecordFromArray rsNew, arrRow, arrFormats
                    rsNew.Update
                    rsNew.Close
                Else
                    ' Record found - update it directly
                    rsCheck.Edit
                    FillRecordFromArray rsCheck, arrRow, arrFormats
                    rsCheck.Update
                    rsCheck.Close
                End If
            Else
                ' ID is empty or invalid - create new record
                Dim rsEmpty As DAO.Recordset
                Set rsEmpty = db.OpenRecordset("tblKartei", dbOpenDynaset)
                rsEmpty.AddNew
                FillRecordFromArray rsEmpty, arrRow, arrFormats
                rsEmpty.Update
                rsEmpty.Close
            End If
        End If
    Next varID
    
    wsDao.CommitTrans
    db.Close
End Sub

Private Sub FillRecordFromArray(ByVal rs As DAO.Recordset, ByVal arrRow As Variant, ByVal arrFormats As Variant)
    ' Assign recordset fields from arrRow(1,1..51).
    ' We skip ID field because it's AutoNumber.
    Dim c As Long
    Application.Calculation = xlCalculationManual
    Application.ScreenUpdating = False

    For c = 1 To 51
        Dim fieldName As String
        fieldName = "Value" & c
        If Not IsError(arrRow(1, c)) Then
            rs.Fields(fieldName).value = arrRow(1, c)
        Else
            rs.Fields(fieldName).value = ""
        End If
        
        fieldName = "InteriorColor" & c
        If IsNull(arrFormats(1, c)) Or IsEmpty(arrFormats(1, c)) Then
            rs.Fields(fieldName).value = 0
        Else
            rs.Fields(fieldName).value = arrFormats(1, c)
        End If
    Next c
    
    If Not IsError(arrRow(1, c)) Then
        rs.Fields("Value52").value = arrRow(1, 52)
    Else
        rs.Fields("Value52").value = ""
    End If
    
    If IsNull(arrFormats(1, 52)) Or IsEmpty(arrFormats(1, 52)) Then
        rs.Fields("FontColor3").value = 0
    Else
        rs.Fields("FontColor3").value = arrFormats(1, 52)
    End If
    
    If IsNull(arrFormats(1, 53)) Or IsEmpty(arrFormats(1, 53)) Then
        rs.Fields("FontColor18").value = 0
    Else
        rs.Fields("FontColor18").value = arrFormats(1, 53)
    End If
    
    Application.Calculation = xlCalculationAutomatic
    Application.ScreenUpdating = True
    
    ' If you do want to store ID in Value48, adapt the logic accordingly.
End Sub

Private Function NzToEmpty(ByVal v As Variant) As Variant
    If IsNull(v) Then
        NzToEmpty = ""
    Else
        NzToEmpty = v
    End If
End Function


