Attribute VB_Name = "nj_NewYearAccess"
Option Explicit

' ============================================================================
' Module: nj_NewYearAccess
' Purpose: Creates Access database for the new year using existing schema.
'          Uses functions from AccessCreation module to create database,
'          table structure (tblKartei), and export data from Kartei sheet.
' ============================================================================

Public Sub nj_CreateAccessDbForYear(ByVal targetYear As Long, ByVal baseFolderPath As String)
    ' Creates a new Access database for the specified year.
    ' Database is created in [baseFolderPath]\Alarm\ subfolder.
    ' Uses existing AccessCreation procedures to:
    '   1. Create the .accdb file
    '   2. Create tblKartei table structure
    '   3. Export Kartei data (preserves ID from AV column)
    '
    ' Parameters:
    '   targetYear: Year number (e.g., 26 for 2026)
    '   baseFolderPath: Base folder path (same as where .xlsm was saved)
    
    Dim dbPath As String
    Dim alarmFolder As String
    
    ' Build database path
    dbPath = nj_GetDbPathForYear(targetYear, baseFolderPath)
    
    ' Ensure Alarm subfolder exists
    alarmFolder = baseFolderPath & "\Alarm"
    nj_EnsureFolderExists alarmFolder
    
    ' Create database file
    CreateAccessDatabase dbPath
    
    ' Create tblKartei table with standard schema
    CreateTableKartei dbPath
    
    ' Export data from Kartei sheet (ID taken from AV, columns 1-52)
    ExportKarteiToAccess dbPath
End Sub

Private Function nj_GetDbPathForYear(ByVal targetYear As Long, ByVal baseFolderPath As String) As String
    ' Builds the full database path for the specified year.
    ' Format: [baseFolderPath]\Alarm\KindElternDaten_[year]_front.accdb
    
    nj_GetDbPathForYear = baseFolderPath & "\Alarm\KindElternDaten_" & targetYear & "_front.accdb"
End Function

Private Sub nj_EnsureFolderExists(ByVal folderPath As String)
    ' Creates the folder if it does not exist.
    ' Uses FileSystemObject for reliable folder creation.
    
    Dim fso As Object
    Set fso = CreateObject("Scripting.FileSystemObject")
    
    If Not fso.FolderExists(folderPath) Then
        fso.CreateFolder folderPath
    End If
    
    Set fso = Nothing
End Sub
