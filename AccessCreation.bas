Attribute VB_Name = "AccessCreation"
Option Explicit

Public Sub CreateAccessDatabase(ByVal dbPath As String)
    ' Creates a new .accdb database using DAO if it doesn't exist.
    ' Requires reference "Microsoft DAO x.x Object Library".
    
    If Dir(dbPath) <> "" Then
        ' If the file already exists, optionally delete or skip.
        ' Kill dbPath
    End If
    
    Dim engine As DAO.DBEngine
    Set engine = New DAO.DBEngine
    
    Dim ws As DAO.Workspace
    Set ws = engine.Workspaces(0)
    
    Dim db As DAO.Database
    Set db = ws.CreateDatabase(dbPath, dbLangGeneral)
    
    db.Close
    Set db = Nothing
    
    Set ws = Nothing
    Set engine = Nothing
End Sub

Public Sub CreateTableKartei(ByVal dbPath As String)
    ' Creates tblKartei with:
    '   ID (AutoNumber, Primary Key)
    '   For each col 1..51 => ValueX (Text), InteriorColorX (Long)
    '   For col 3 (C) => FontColor3 (Long)
    '   For col 18 (R) => FontColor18 (Long)
    
    Dim engine As DAO.DBEngine
    Set engine = New DAO.DBEngine
    
    Dim ws As DAO.Workspace
    Set ws = engine.Workspaces(0)
    
    Dim db As DAO.Database
    Set db = ws.OpenDatabase(dbPath)
    
    ' Create TableDef
    Dim tdf As DAO.TableDef
    Set tdf = db.CreateTableDef("tblKartei")
    
    ' 1) Add an AutoNumber field ID
    Dim fldID As DAO.Field
    Set fldID = tdf.CreateField("ID", dbLong)
    fldID.Attributes = dbAutoIncrField  ' AutoNumber
    tdf.Fields.Append fldID
    
    ' 2) For each column 1..51, add ValueX (Text(255)) and InteriorColorX (Long)
    Dim i As Long
    For i = 1 To 51
        Dim fldValue As DAO.Field
        Set fldValue = tdf.CreateField("Value" & i, dbText, 255)
        fldValue.Required = False
        fldValue.AllowZeroLength = True
        tdf.Fields.Append fldValue
        
        Dim fldInterior As DAO.Field
        Set fldInterior = tdf.CreateField("InteriorColor" & i, dbLong)
        fldInterior.Required = False

        tdf.Fields.Append fldInterior
    Next i
    
    ' 3) Add ONLY FontColor3, FontColor18
    Dim fldFont3 As DAO.Field
    Set fldFont3 = tdf.CreateField("FontColor3", dbLong)
    fldFont3.Required = False
    tdf.Fields.Append fldFont3
    
    Dim fldFont18 As DAO.Field
    Set fldFont18 = tdf.CreateField("FontColor18", dbLong)
    fldFont18.Required = False
    tdf.Fields.Append fldFont18
    
    Set fldValue = tdf.CreateField("Value" & 52, dbMemo)
    fldValue.Required = False
    fldValue.AllowZeroLength = True
    tdf.Fields.Append fldValue
    
    ' Append table
    db.TableDefs.Append tdf
    
    ' Create a primary key index on ID
    db.Execute "CREATE INDEX PK_ID ON tblKartei(ID) WITH PRIMARY"
    
    db.Close
    Set tdf = Nothing
    Set db = Nothing
    Set ws = Nothing
    Set engine = Nothing
End Sub


' ****************************************
' **    Section 2                       **
' ****************************************

Public Sub ExportKarteiToAccess(ByVal dbPath As String)
    ' Reads the Kartei sheet from row 3 down, columns 1..52 (A..AZ).
    ' For each row:
    '   - builds a SQL INSERT with all ValueX, InteriorColorX
    '   - if col=3 or col=18 => also sets FontColor3 or FontColor18
    ' Uses a single transaction for performance.

    Dim ws As Worksheet
    Set ws = ThisWorkbook.Worksheets("Kartei")
    
    Dim lastRow As Long
    lastRow = ws.Cells(ws.Rows.count, 1).End(xlUp).row
    If lastRow < 3 Then
        MsgBox "No data rows in Kartei!", vbInformation
        Exit Sub
    End If
    
    ' Set up DAO
    Dim engine As DAO.DBEngine
    Set engine = New DAO.DBEngine
    
    Dim wsDao As DAO.Workspace
    Set wsDao = engine.Workspaces(0)
    
    Dim db As DAO.Database
    Set db = wsDao.OpenDatabase(dbPath)
    
    ' Begin a single transaction to speed up insertion
    wsDao.BeginTrans
    
    Dim r As Long
    For r = 3 To lastRow
        
        ' We'll build an INSERT statement for each row
        Dim sqlInsert As String
        sqlInsert = "INSERT INTO tblKartei ("
        
        ' We won't insert 'ID' because it's AutoNumber
        ' So we list all the fields we want to fill in:
        
        Dim fieldList As String
        Dim valueList As String
        
        fieldList = ""
        valueList = ""
        
        Dim c As Long
        For c = 1 To 51
            ' fieldList for Value c, InteriorColor c
            fieldList = fieldList & "Value" & c & ",InteriorColor" & c & ","
            
            ' read the cell
            Dim cellVal As Variant
            cellVal = ws.Cells(r, c).value
            
            Dim sVal As String
            ' If we want text, we must escape quotes if any
            sVal = Replace(CStr(cellVal), "'", "''")
            
            ' Also read the interior color
            Dim iColor As Long
            iColor = ws.Cells(r, c).Interior.Color
            
            ' We'll put them in the SQL
            valueList = valueList & "'" & sVal & "'" & "," & iColor & ","
        Next c
        
        ' Now handle FontColor3, FontColor18
        fieldList = fieldList & "FontColor3,FontColor18,Value52"
        
        ' read FontColor for col 3
        Dim fc3 As Long
        fc3 = ws.Cells(r, 3).Font.Color
        
        ' read FontColor for col 18
        Dim fc18 As Long
        fc18 = ws.Cells(r, 18).Font.Color
        
        cellVal = ws.Cells(r, 52).value
        sVal = Replace(CStr(cellVal), "'", "''")
        
        valueList = valueList & fc3 & "," & fc18 & ",'" & sVal & "'"
        
        ' Build final SQL
        sqlInsert = sqlInsert & fieldList & ") VALUES (" & valueList & ")"
        
        db.Execute sqlInsert, dbFailOnError
        
    Next r
    
    ' Commit once, after all inserts
    wsDao.CommitTrans
    
    db.Close
    Set db = Nothing
    Set wsDao = Nothing
    Set engine = Nothing
    
    MsgBox "Exported " & (lastRow - 2) & " rows to tblKartei in " & dbPath, vbInformation
End Sub

' ****************************************
' **    Section 3                       **
' ****************************************

Public Sub CreateAndExportKarteiAccess()
    Dim dbPath As String
    dbPath = ThisWorkbook.path & "\KindElternDaten_25_front.accdb"
    
    ' 1) Create new .accdb (if needed)
    CreateAccessDatabase dbPath
    
    ' 2) Create table structure
    CreateTableKartei dbPath
    
    ' 3) Export Kartei from Excel
    ExportKarteiToAccess dbPath
End Sub

Public Sub ExportKarteiToAccessManual()
    Dim dbPath As String
    dbPath = ThisWorkbook.path & "\Alarm\KindElternDaten_25_front.accdb"
    
    ' Export Kartei from Excel
    ExportKarteiToAccess dbPath
End Sub
