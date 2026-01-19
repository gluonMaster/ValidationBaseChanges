Attribute VB_Name = "valid_YearConfig"
'==========================
'   Module: valid_YearConfig
'   Purpose: Multi-year database configuration for Superadmin
'
'   Manages database paths for years 2024, 2025, 2026:
'   - Cloud root folder (contains year subfolders 2024/2025/2026)
'   - Per-year overrides if a specific year lives in a different location
'   - DB file naming: KindElternDaten_<YY>_front.accdb
'   - Path structure: <CloudRoot>\<Year>\Alarm\KindElternDaten_<YY>_front.accdb
'
'   Configuration stored in hidden worksheet "DBConfig":
'     A1: CloudRootPath (main cloud folder containing year subfolders)
'     B1: Year24Override (optional, empty = use CloudRoot)
'     C1: Year25Override (optional, empty = use CloudRoot)
'     D1: Year26Override (optional, empty = use CloudRoot)
'
'   Sheet naming convention:
'     - KarteiYY (e.g. Kartei24, Kartei25, Kartei26)
'     - grossGeschichteYY
'     - GeschichteYY
'     - GeschichteYY_Alle
'==========================

Option Explicit

' Configuration sheet name (hidden)
Private Const CONFIG_SHEET_NAME As String = "DBConfig"

' Database subfolder and filename pattern
Private Const DB_SUBFOLDER As String = "Alarm"
Private Const DB_FILENAME_PATTERN As String = "KindElternDaten_<YY>_front.accdb"

' Supported years
Private Const YEAR_24 As Integer = 24
Private Const YEAR_25 As Integer = 25
Private Const YEAR_26 As Integer = 26

' Config sheet cell addresses
Private Const CELL_CLOUD_ROOT As String = "A1"
Private Const CELL_OVERRIDE_24 As String = "B1"
Private Const CELL_OVERRIDE_25 As String = "C1"
Private Const CELL_OVERRIDE_26 As String = "D1"

' ============================================================
' PUBLIC API - Year Configuration
' ============================================================

' Returns array of supported 2-digit years (24, 25, 26)
Public Function GetSupportedYears() As Variant
    GetSupportedYears = Array(YEAR_24, YEAR_25, YEAR_26)
End Function

' Returns the full 4-digit year from 2-digit year
Public Function GetFullYear(ByVal year2 As Integer) As Integer
    GetFullYear = 2000 + year2
End Function

' Returns the 2-digit year suffix (e.g., 24 from 2024)
Public Function GetYear2Digit(ByVal fullYear As Integer) As Integer
    GetYear2Digit = fullYear Mod 100
End Function

' ============================================================
' PUBLIC API - Database Path Resolution
' ============================================================

' Get the validated database path for a specific year.
' If the path is not configured or invalid, prompts the user.
' Returns empty string if user cancels or configuration fails.
'
' @param year2 - Two-digit year (24, 25, or 26)
' @return String - Full path to the database file, or empty string on failure
Public Function GetDbPathForYear(ByVal year2 As Integer) As String
    On Error GoTo ErrorHandler
    
    ' Validate year parameter
    If Not IsValidYear(year2) Then
        MsgBox "Ungueltiges Jahr: " & year2 & ". Unterstuetzte Jahre: 24, 25, 26.", _
               vbExclamation, "Jahresfehler"
        GetDbPathForYear = ""
        Exit Function
    End If
    
    ' Try to build the database path
    Dim dbPath As String
    dbPath = BuildDbPathForYear(year2)
    
    ' Check if database file exists
    If Not FileExists(dbPath) Then
        ' Path is not valid - prompt user to configure
        Dim userPath As String
        userPath = PromptForYearDbPath(year2, dbPath)
        
        If userPath = "" Then
            ' User cancelled
            GetDbPathForYear = ""
            Exit Function
        End If
        
        dbPath = userPath
    End If
    
    GetDbPathForYear = dbPath
    Exit Function
    
ErrorHandler:
    MsgBox "Fehler beim Abrufen des Datenbankpfads fuer Jahr " & year2 & ": " & Err.Description, _
           vbCritical, "Pfadfehler"
    GetDbPathForYear = ""
End Function

' Get database path for the current year (2026)
' Convenience wrapper for common use case
Public Function GetDbPathForCurrentYear() As String
    GetDbPathForCurrentYear = GetDbPathForYear(YEAR_26)
End Function

' Get database path for year 25 (backward compatibility)
' This function is used by existing code that calls GetValidatedDatabasePath
Public Function GetDbPathForYear25() As String
    GetDbPathForYear25 = GetDbPathForYear(YEAR_25)
End Function

' Get the expected database path for a year based on current configuration.
' - No validation, no file existence checks, no prompts.
' - Returns empty string if not configured.
Public Function GetExpectedDbPathForYear(ByVal year2 As Integer) As String
    If Not IsValidYear(year2) Then
        GetExpectedDbPathForYear = ""
        Exit Function
    End If
    
    GetExpectedDbPathForYear = BuildDbPathForYear(year2)
End Function

' Non-interactive attempt to resolve a DB path for a year.
' Returns the expected path and sets status:
'   - "ok"            : file exists
'   - "missing-file"  : config exists, but file not found at expected path
'   - "not-configured": config missing / empty
'   - "invalid-year"  : year not supported
'   - "error"         : unexpected error
Public Function TryGetDbPathForYearNoPrompt(ByVal year2 As Integer, ByRef status As String) As String
    On Error GoTo ErrorHandler
    
    status = ""
    
    If Not IsValidYear(year2) Then
        status = "invalid-year"
        TryGetDbPathForYearNoPrompt = ""
        Exit Function
    End If
    
    Dim expectedPath As String
    expectedPath = BuildDbPathForYear(year2)
    
    If Len(Trim(expectedPath)) = 0 Then
        status = "not-configured"
        TryGetDbPathForYearNoPrompt = ""
        Exit Function
    End If
    
    If FileExists(expectedPath) Then
        status = "ok"
    Else
        status = "missing-file"
    End If
    
    TryGetDbPathForYearNoPrompt = expectedPath
    Exit Function
    
ErrorHandler:
    status = "error"
    TryGetDbPathForYearNoPrompt = ""
End Function

' Read the configured year override folder (no prompts, no creation)
Public Function GetYearOverrideFromConfig(ByVal year2 As Integer) As String
    On Error Resume Next
    
    Dim wsConfig As Worksheet
    Set wsConfig = GetConfigSheet()
    
    If wsConfig Is Nothing Then
        GetYearOverrideFromConfig = ""
        Exit Function
    End If
    
    GetYearOverrideFromConfig = GetYearOverride(wsConfig, year2)
    On Error GoTo 0
End Function

' Build the database path for a year without validation
' Returns the expected path based on current configuration
Private Function BuildDbPathForYear(ByVal year2 As Integer) As String
    Dim basePath As String
    basePath = GetYearBasePath(year2)
    
    If basePath = "" Then
        BuildDbPathForYear = ""
        Exit Function
    End If
    
    Dim dbFilename As String
    dbFilename = Replace(DB_FILENAME_PATTERN, "<YY>", Format(year2, "00"))
    
    BuildDbPathForYear = basePath & "\" & DB_SUBFOLDER & "\" & dbFilename
End Function

' Get the base path for a year (CloudRoot\Year or year-specific override)
Private Function GetYearBasePath(ByVal year2 As Integer) As String
    Dim wsConfig As Worksheet
    Set wsConfig = GetConfigSheet()
    
    If wsConfig Is Nothing Then
        ' Config sheet doesn't exist yet - return empty to trigger setup
        GetYearBasePath = ""
        Exit Function
    End If
    
    ' Check for year-specific override first
    Dim overridePath As String
    overridePath = GetYearOverride(wsConfig, year2)
    
    If overridePath <> "" Then
        GetYearBasePath = overridePath
        Exit Function
    End If
    
    ' Use cloud root + year folder
    Dim cloudRoot As String
    cloudRoot = Trim(CStr(wsConfig.Range(CELL_CLOUD_ROOT).Value))
    
    If cloudRoot = "" Then
        GetYearBasePath = ""
        Exit Function
    End If
    
    ' Append year folder (e.g., 2024, 2025, 2026)
    GetYearBasePath = cloudRoot & "\" & CStr(GetFullYear(year2))
End Function

' Get the year-specific override path (if configured)
Private Function GetYearOverride(ByVal wsConfig As Worksheet, ByVal year2 As Integer) As String
    On Error Resume Next
    
    Dim cellAddr As String
    Select Case year2
        Case YEAR_24: cellAddr = CELL_OVERRIDE_24
        Case YEAR_25: cellAddr = CELL_OVERRIDE_25
        Case YEAR_26: cellAddr = CELL_OVERRIDE_26
        Case Else: GetYearOverride = "": Exit Function
    End Select
    
    GetYearOverride = Trim(CStr(wsConfig.Range(cellAddr).Value))
    On Error GoTo 0
End Function

' ============================================================
' PUBLIC API - Configuration Setup
' ============================================================

' User-friendly entry point for database configuration
' Shows guided setup and reports final status
Public Sub ConfigureAllDatabases()
    Dim success As Boolean
    success = EnsureAllDbPathsConfigured()
    
    If success Then
        MsgBox "Alle Datenbankpfade sind korrekt konfiguriert.", vbInformation, "Konfiguration OK"
    End If
End Sub

' Ensure all database paths are configured for all supported years.
' Prompts user if any path is missing or invalid.
' Returns True if all paths are valid, False otherwise.
Public Function EnsureAllDbPathsConfigured() As Boolean
    On Error GoTo ErrorHandler
    
    Dim years As Variant
    years = GetSupportedYears()
    
    Dim allValid As Boolean
    allValid = True
    
    Dim missingYears As String
    missingYears = ""
    
    Dim i As Long
    For i = LBound(years) To UBound(years)
        Dim year2 As Integer
        year2 = CInt(years(i))
        
        Dim dbPath As String
        dbPath = BuildDbPathForYear(year2)
        
        If Not FileExists(dbPath) Then
            If missingYears <> "" Then missingYears = missingYears & ", "
            missingYears = missingYears & "20" & Format(year2, "00")
            allValid = False
        End If
    Next i
    
    If Not allValid Then
        ' Show guided setup dialog
        Dim result As VbMsgBoxResult
        result = MsgBox("Datenbankpfade fuer folgende Jahre nicht gefunden: " & missingYears & vbCrLf & vbCrLf & _
                       "Moechten Sie den Cloud-Stammordner konfigurieren?" & vbCrLf & _
                       "(Der Ordner sollte Unterordner 2024, 2025, 2026 enthalten)", _
                       vbYesNo + vbQuestion, "Datenbankkonfiguration erforderlich")
        
        If result = vbYes Then
            allValid = RunGuidedDbSetup()
        Else
            allValid = False
        End If
    End If
    
    EnsureAllDbPathsConfigured = allValid
    Exit Function
    
ErrorHandler:
    MsgBox "Fehler bei der Datenbankkonfiguration: " & Err.Description, vbCritical, "Konfigurationsfehler"
    EnsureAllDbPathsConfigured = False
End Function

' Run guided setup to configure all database paths
' Returns True if all paths are valid after setup
Private Function RunGuidedDbSetup() As Boolean
    On Error GoTo ErrorHandler
    
    ' Step 1: Ask user for cloud root folder
    Dim cloudRoot As String
    cloudRoot = BrowseForFolder("Waehlen Sie den Cloud-Stammordner (enthaelt 2024, 2025, 2026)")
    
    If cloudRoot = "" Then
        MsgBox "Kein Ordner ausgewaehlt. Konfiguration abgebrochen.", vbInformation, "Abgebrochen"
        RunGuidedDbSetup = False
        Exit Function
    End If
    
    ' Save cloud root to config
    SaveCloudRoot cloudRoot
    
    ' Step 2: Validate all years and prompt for overrides if needed
    Dim years As Variant
    years = GetSupportedYears()
    
    Dim allValid As Boolean
    allValid = True
    
    Dim i As Long
    For i = LBound(years) To UBound(years)
        Dim year2 As Integer
        year2 = CInt(years(i))
        
        Dim dbPath As String
        dbPath = BuildDbPathForYear(year2)
        
        If Not FileExists(dbPath) Then
            ' This year's DB not found in cloud root - ask for override
            Dim overridePath As String
            overridePath = PromptForYearOverride(year2, dbPath)
            
            If overridePath = "" Then
                ' User cancelled or override also failed
                allValid = False
            End If
        End If
    Next i
    
    If allValid Then
        MsgBox "Alle Datenbankpfade erfolgreich konfiguriert!", vbInformation, "Konfiguration abgeschlossen"
    End If
    
    RunGuidedDbSetup = allValid
    Exit Function
    
ErrorHandler:
    MsgBox "Fehler bei der Einrichtung: " & Err.Description, vbCritical, "Einrichtungsfehler"
    RunGuidedDbSetup = False
End Function

' Prompt user to select a database path for a specific year
Private Function PromptForYearDbPath(ByVal year2 As Integer, ByVal currentPath As String) As String
    PromptForYearDbPath = ""
    
    Dim dbFilename As String
    dbFilename = Replace(DB_FILENAME_PATTERN, "<YY>", Format(year2, "00"))
    
    ' Show message explaining the problem
    Dim msgResult As VbMsgBoxResult
    msgResult = MsgBox("Datenbankdatei fuer Jahr 20" & Format(year2, "00") & " nicht gefunden:" & vbCrLf & vbCrLf & _
                       currentPath & vbCrLf & vbCrLf & _
                       "Moechten Sie den Ordner auswaehlen?" & vbCrLf & _
                       "(Der Ordner sollte '" & DB_SUBFOLDER & "\" & dbFilename & "' enthalten)", _
                       vbYesNo + vbQuestion, "Datenbank nicht gefunden - Jahr 20" & Format(year2, "00"))
    
    If msgResult = vbNo Then
        Exit Function
    End If
    
    ' Show folder picker dialog
    Dim folderPath As String
    folderPath = BrowseForFolder("Waehlen Sie den Jahresordner fuer 20" & Format(year2, "00") & " (uebergeordneter Ordner von " & DB_SUBFOLDER & ")")
    
    If folderPath = "" Then
        MsgBox "Kein Ordner ausgewaehlt. Vorgang abgebrochen.", vbInformation, "Abgebrochen"
        Exit Function
    End If
    
    ' Build and validate the new path
    Dim newDbPath As String
    newDbPath = folderPath & "\" & DB_SUBFOLDER & "\" & dbFilename
    
    If Not FileExists(newDbPath) Then
        MsgBox "Datenbankdatei immer noch nicht gefunden unter:" & vbCrLf & vbCrLf & _
               newDbPath & vbCrLf & vbCrLf & _
               "Bitte stellen Sie sicher, dass der ausgewaehlte Ordner '" & DB_SUBFOLDER & "\" & dbFilename & "' enthaelt.", _
               vbExclamation, "Datenbank nicht gefunden"
        Exit Function
    End If
    
    ' Save the selected path as year override
    SaveYearOverride year2, folderPath
    
    MsgBox "Datenbankpfad fuer Jahr 20" & Format(year2, "00") & " erfolgreich aktualisiert!" & vbCrLf & vbCrLf & _
           "Neuer Pfad: " & newDbPath, vbInformation, "Pfad aktualisiert"
    
    PromptForYearDbPath = newDbPath
End Function

' Prompt user for a year-specific override path
Private Function PromptForYearOverride(ByVal year2 As Integer, ByVal expectedPath As String) As String
    PromptForYearOverride = ""
    
    Dim dbFilename As String
    dbFilename = Replace(DB_FILENAME_PATTERN, "<YY>", Format(year2, "00"))
    
    Dim msgResult As VbMsgBoxResult
    msgResult = MsgBox("Datenbank fuer Jahr 20" & Format(year2, "00") & " nicht im Cloud-Stammordner gefunden:" & vbCrLf & vbCrLf & _
                       expectedPath & vbCrLf & vbCrLf & _
                       "Moechten Sie einen alternativen Ordner fuer dieses Jahr auswaehlen?", _
                       vbYesNo + vbQuestion, "Jahr 20" & Format(year2, "00") & " nicht gefunden")
    
    If msgResult = vbNo Then
        Exit Function
    End If
    
    Dim folderPath As String
    folderPath = BrowseForFolder("Waehlen Sie den Jahresordner fuer 20" & Format(year2, "00"))
    
    If folderPath = "" Then
        Exit Function
    End If
    
    ' Validate
    Dim newDbPath As String
    newDbPath = folderPath & "\" & DB_SUBFOLDER & "\" & dbFilename
    
    If Not FileExists(newDbPath) Then
        MsgBox "Datenbankdatei nicht gefunden unter:" & vbCrLf & newDbPath, vbExclamation, "Nicht gefunden"
        Exit Function
    End If
    
    ' Save override
    SaveYearOverride year2, folderPath
    
    PromptForYearOverride = newDbPath
End Function

' ============================================================
' PUBLIC API - Sheet Management
' ============================================================

' Ensure all required sheets exist for a given year.
' Creates missing sheets with proper headers.
'
' @param year2 - Two-digit year (24, 25, or 26)
Public Sub EnsureYearSheetsExist(ByVal year2 As Integer)
    On Error GoTo ErrorHandler
    
    If Not IsValidYear(year2) Then
        MsgBox "Ungueltiges Jahr: " & year2, vbExclamation, "Jahresfehler"
        Exit Sub
    End If
    
    Dim suffix As String
    suffix = Format(year2, "00")
    
    ' Ensure each required sheet exists
    EnsureSheetExists "Kartei" & suffix, "Kartei"
    EnsureSheetExists "grossGeschichte" & suffix, "grossGeschichte"
    EnsureSheetExists "Geschichte" & suffix, "Geschichte"
    EnsureSheetExists "Geschichte" & suffix & "_Alle", "Geschichte"
    
    Exit Sub
    
ErrorHandler:
    MsgBox "Fehler beim Erstellen der Blaetter fuer Jahr " & year2 & ": " & Err.Description, _
           vbCritical, "Blattfehler"
End Sub

' Ensure a specific sheet exists, creating it with headers if needed
Private Sub EnsureSheetExists(ByVal sheetName As String, ByVal templateName As String)
    On Error Resume Next
    
    Dim ws As Worksheet
    Set ws = ThisWorkbook.Worksheets(sheetName)
    
    If ws Is Nothing Then
        ' Sheet doesn't exist - create it
        Set ws = ThisWorkbook.Worksheets.Add(After:=ThisWorkbook.Worksheets(ThisWorkbook.Worksheets.Count))
        ws.Name = sheetName
        
        ' Copy headers from template sheet if it exists
        Dim wsTemplate As Worksheet
        Set wsTemplate = ThisWorkbook.Worksheets(templateName)
        
        If Not wsTemplate Is Nothing Then
            ' Copy header rows (typically rows 1-2)
            wsTemplate.Rows("1:2").Copy ws.Rows("1:2")
        End If
    End If
    
    On Error GoTo 0
End Sub

' ============================================================
' PUBLIC API - Wrapper Macros for UI (Ribbon/Buttons)
' ============================================================

' Select cloud root folder (main configuration)
Public Sub SelectDbRootFolder()
    Dim cloudRoot As String
    cloudRoot = BrowseForFolder("Waehlen Sie den Cloud-Stammordner (enthaelt 2024, 2025, 2026)")
    
    If cloudRoot = "" Then
        MsgBox "Kein Ordner ausgewaehlt.", vbInformation, "Abgebrochen"
        Exit Sub
    End If
    
    SaveCloudRoot cloudRoot
    
    ' Validate all years
    Dim validationResult As String
    validationResult = ValidateAllDbPaths()
    
    MsgBox "Cloud-Stammordner gesetzt auf:" & vbCrLf & vbCrLf & cloudRoot & vbCrLf & vbCrLf & _
           "Validierung:" & vbCrLf & validationResult, vbInformation, "Konfiguration aktualisiert"
End Sub

' Select database folder for year 24
Public Sub SelectDbFolder24()
    SelectDbFolderForYear YEAR_24
End Sub

' Select database folder for year 25
Public Sub SelectDbFolder25()
    SelectDbFolderForYear YEAR_25
End Sub

' Select database folder for year 26
Public Sub SelectDbFolder26()
    SelectDbFolderForYear YEAR_26
End Sub

' Select database folder for a specific year
Private Sub SelectDbFolderForYear(ByVal year2 As Integer)
    Dim dbFilename As String
    dbFilename = Replace(DB_FILENAME_PATTERN, "<YY>", Format(year2, "00"))
    
    Dim folderPath As String
    folderPath = BrowseForFolder("Waehlen Sie den Jahresordner fuer 20" & Format(year2, "00") & _
                                 " (uebergeordneter Ordner von " & DB_SUBFOLDER & ")")
    
    If folderPath = "" Then
        MsgBox "Kein Ordner ausgewaehlt.", vbInformation, "Abgebrochen"
        Exit Sub
    End If
    
    ' Validate
    Dim dbPath As String
    dbPath = folderPath & "\" & DB_SUBFOLDER & "\" & dbFilename
    
    If Not FileExists(dbPath) Then
        MsgBox "Datenbankdatei nicht gefunden unter:" & vbCrLf & vbCrLf & dbPath, _
               vbExclamation, "Nicht gefunden"
        Exit Sub
    End If
    
    ' Save as override
    SaveYearOverride year2, folderPath
    
    MsgBox "Datenbankpfad fuer Jahr 20" & Format(year2, "00") & " gesetzt auf:" & vbCrLf & vbCrLf & dbPath, _
           vbInformation, "Pfad aktualisiert"
End Sub

' Validate all database paths and return status string
Private Function ValidateAllDbPaths() As String
    Dim result As String
    result = ""
    
    Dim years As Variant
    years = GetSupportedYears()
    
    Dim i As Long
    For i = LBound(years) To UBound(years)
        Dim year2 As Integer
        year2 = CInt(years(i))
        
        Dim dbPath As String
        dbPath = BuildDbPathForYear(year2)
        
        Dim status As String
        If FileExists(dbPath) Then
            status = "OK"
        Else
            status = "NICHT GEFUNDEN"
        End If
        
        result = result & "20" & Format(year2, "00") & ": " & status & vbCrLf
    Next i
    
    ValidateAllDbPaths = result
End Function

' ============================================================
' CONFIG SHEET MANAGEMENT
' ============================================================

' Get the configuration worksheet (hidden), or Nothing if it doesn't exist
Private Function GetConfigSheet() As Worksheet
    On Error Resume Next
    Set GetConfigSheet = ThisWorkbook.Worksheets(CONFIG_SHEET_NAME)
    On Error GoTo 0
End Function

' Get or create the configuration worksheet (hidden)
Private Function GetOrCreateConfigSheet() As Worksheet
    On Error Resume Next
    
    Dim ws As Worksheet
    Set ws = ThisWorkbook.Worksheets(CONFIG_SHEET_NAME)
    
    If ws Is Nothing Then
        ' Create the config sheet
        Set ws = ThisWorkbook.Worksheets.Add(After:=ThisWorkbook.Worksheets(ThisWorkbook.Worksheets.Count))
        ws.Name = CONFIG_SHEET_NAME
        ws.Visible = xlSheetVeryHidden
        
        ' Add labels for clarity (optional, can be removed)
        ws.Range("A2").Value = "CloudRoot"
        ws.Range("B2").Value = "Override24"
        ws.Range("C2").Value = "Override25"
        ws.Range("D2").Value = "Override26"
    End If
    
    Set GetOrCreateConfigSheet = ws
    On Error GoTo 0
End Function

' Save cloud root path to config sheet
Private Sub SaveCloudRoot(ByVal cloudRoot As String)
    Dim wsConfig As Worksheet
    Set wsConfig = GetOrCreateConfigSheet()
    wsConfig.Range(CELL_CLOUD_ROOT).Value = cloudRoot
End Sub

' Get cloud root path from config sheet
Public Function GetCloudRoot() As String
    On Error Resume Next
    
    Dim wsConfig As Worksheet
    Set wsConfig = GetConfigSheet()
    
    If wsConfig Is Nothing Then
        GetCloudRoot = ""
        Exit Function
    End If
    
    GetCloudRoot = Trim(CStr(wsConfig.Range(CELL_CLOUD_ROOT).Value))
    On Error GoTo 0
End Function

' Save year-specific override path
Private Sub SaveYearOverride(ByVal year2 As Integer, ByVal folderPath As String)
    Dim wsConfig As Worksheet
    Set wsConfig = GetOrCreateConfigSheet()
    
    Dim cellAddr As String
    Select Case year2
        Case YEAR_24: cellAddr = CELL_OVERRIDE_24
        Case YEAR_25: cellAddr = CELL_OVERRIDE_25
        Case YEAR_26: cellAddr = CELL_OVERRIDE_26
        Case Else: Exit Sub
    End Select
    
    wsConfig.Range(cellAddr).Value = folderPath
End Sub

' ============================================================
' UTILITY FUNCTIONS
' ============================================================

' Check if year is valid (24, 25, or 26)
' Made Public for use by valid_Dashboard and other modules
Public Function IsValidYear(ByVal year2 As Integer) As Boolean
    IsValidYear = (year2 = YEAR_24 Or year2 = YEAR_25 Or year2 = YEAR_26)
End Function

' Check if file exists
Private Function FileExists(ByVal filePath As String) As Boolean
    On Error Resume Next
    If Len(Trim(filePath)) = 0 Then
        FileExists = False
        Exit Function
    End If
    FileExists = (Dir(filePath) <> "")
    On Error GoTo 0
End Function

' Browse for folder using FileDialog
Private Function BrowseForFolder(ByVal dialogTitle As String) As String
    BrowseForFolder = ""
    
    Dim fd As FileDialog
    Set fd = Application.FileDialog(msoFileDialogFolderPicker)
    
    With fd
        .Title = dialogTitle
        .InitialFileName = ThisWorkbook.Path & "\"
        .AllowMultiSelect = False
        
        If .Show = -1 Then
            BrowseForFolder = .SelectedItems(1)
        End If
    End With
    
    Set fd = Nothing
End Function

' ============================================================
' BACKWARD COMPATIBILITY - Migration from old single-year config
' ============================================================

' Migrate from old Kartei!X1 storage to new DBConfig sheet
' Call this once when upgrading from single-year to multi-year
Public Sub MigrateFromLegacyConfig()
    On Error Resume Next
    
    ' Try to read old path from Kartei!X1
    Dim wsKartei As Worksheet
    Set wsKartei = ThisWorkbook.Worksheets("Kartei")
    
    If wsKartei Is Nothing Then Exit Sub
    
    Dim oldPath As String
    oldPath = Trim(CStr(wsKartei.Range("X1").Value))
    
    If oldPath = "" Then Exit Sub
    
    ' The old path pointed to the year folder (e.g., C:\...\2025)
    ' We need to extract the cloud root (parent of 2024/2025/2026)
    
    ' Check if this looks like a year-specific path
    Dim cloudRoot As String
    If Right(oldPath, 4) = "2024" Or Right(oldPath, 4) = "2025" Or Right(oldPath, 4) = "2026" Then
        ' Extract parent folder as cloud root
        cloudRoot = Left(oldPath, Len(oldPath) - 5) ' Remove "\2024" or "\2025" or "\2026"
    Else
        ' Assume old path is the cloud root
        cloudRoot = oldPath
    End If
    
    ' Check if config sheet already has a cloud root
    Dim existingRoot As String
    existingRoot = GetCloudRoot()
    
    If existingRoot = "" Then
        ' Migrate to new config
        SaveCloudRoot cloudRoot
        
        ' Optionally: Set year 25 override to the old path if it was year-specific
        If Right(oldPath, 4) = "2025" Then
            SaveYearOverride YEAR_25, oldPath
        End If
        
        MsgBox "Konfiguration von Kartei!X1 migriert." & vbCrLf & _
               "Cloud-Stammordner: " & cloudRoot, vbInformation, "Migration abgeschlossen"
    End If
    
    On Error GoTo 0
End Sub

' ============================================================
' SHEET NAME HELPERS
' ============================================================

' Get Kartei sheet name for a year
Public Function GetKarteiSheetName(ByVal year2 As Integer) As String
    GetKarteiSheetName = "Kartei" & Format(year2, "00")
End Function

' Get grossGeschichte sheet name for a year
Public Function GetGrossGeschichteSheetName(ByVal year2 As Integer) As String
    GetGrossGeschichteSheetName = "grossGeschichte" & Format(year2, "00")
End Function

' Get Geschichte (single-record) sheet name for a year
Public Function GetGeschichteSheetName(ByVal year2 As Integer) As String
    GetGeschichteSheetName = "Geschichte" & Format(year2, "00")
End Function

' Get Geschichte_Alle (all-records) sheet name for a year
Public Function GetGeschichteAlleSheetName(ByVal year2 As Integer) As String
    GetGeschichteAlleSheetName = "Geschichte" & Format(year2, "00") & "_Alle"
End Function

' Extract year from sheet name (returns 0 if not a year-suffixed sheet)
Public Function GetYearFromSheetName(ByVal sheetName As String) As Integer
    On Error Resume Next
    
    GetYearFromSheetName = 0
    
    ' Check for pattern like "Kartei25" or "grossGeschichte24_Alle"
    Dim suffix As String
    
    ' Try to extract 2 digits before "_Alle" or at end
    If InStr(sheetName, "_Alle") > 0 Then
        suffix = Mid(sheetName, InStr(sheetName, "_Alle") - 2, 2)
    Else
        suffix = Right(sheetName, 2)
    End If
    
    If IsNumeric(suffix) Then
        Dim year2 As Integer
        year2 = CInt(suffix)
        If IsValidYear(year2) Then
            GetYearFromSheetName = year2
        End If
    End If
    
    On Error GoTo 0
End Function
