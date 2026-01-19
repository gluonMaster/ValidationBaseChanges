Attribute VB_Name = "valid_DatabasePath"
'==========================
'   Module: valid_DatabasePath
'   Purpose: Centralized database path management for Superadmin
'
'   MULTI-YEAR SUPPORT (2024, 2025, 2026):
'   This module now delegates to valid_YearConfig for multi-year path resolution.
'   The old single-year functions are preserved for backward compatibility
'   but internally route through the new year-aware configuration.
'
'   For new code, prefer using valid_YearConfig directly:
'     - valid_YearConfig.GetDbPathForYear(year2) - Get DB path for specific year
'     - valid_YearConfig.EnsureAllDbPathsConfigured() - Validate all years
'     - valid_YearConfig.SelectDbRootFolder - Configure cloud root
'
'   Legacy storage (Kartei!X1) is migrated to DBConfig sheet on first use.
'==========================

Option Explicit

' Default year for backward compatibility (existing code expects year 25)
Private Const DEFAULT_YEAR As Integer = 25

' One-time legacy migration guard (must be declared before procedures)
Private m_LegacyMigrated As Boolean

' ============================================================
' BACKWARD COMPATIBILITY API
' ============================================================

' Get validated database path (defaults to year 25 for backward compatibility)
' New code should use valid_YearConfig.GetDbPathForYear(year2) instead.
Public Function GetValidatedDatabasePath() As String
    ' Ensure legacy config is migrated
    MigrateLegacyConfigIfNeeded
    
    ' Delegate to year config module
    GetValidatedDatabasePath = valid_YearConfig.GetDbPathForYear(DEFAULT_YEAR)
End Function

' Get validated database path for a specific year
' This is the preferred function for new multi-year code.
'
' @param year2 - Two-digit year (24, 25, or 26)
' @return String - Full path to the database file, or empty string on failure
Public Function GetValidatedDatabasePathForYear(ByVal year2 As Integer) As String
    ' Ensure legacy config is migrated
    MigrateLegacyConfigIfNeeded
    
    ' Delegate to year config module
    GetValidatedDatabasePathForYear = valid_YearConfig.GetDbPathForYear(year2)
End Function

' Public procedure to manually select database folder (legacy - year 25)
' For multi-year setup, use SelectDbRootFolder or SelectDbFolderXX instead.
Public Sub SelectDatabaseFolder()
    ' Migrate legacy config if needed
    MigrateLegacyConfigIfNeeded
    
    ' Use year-specific selection for year 25
    valid_YearConfig.SelectDbFolder25
End Sub

' ============================================================
' MULTI-YEAR API (Thin wrappers - delegate to valid_YearConfig)
' ============================================================

' Select the cloud root folder (contains 2024, 2025, 2026 subfolders)
' This is the recommended way to configure all database paths at once.
Public Sub SelectDbRootFolder()
    valid_YearConfig.SelectDbRootFolder
End Sub

' Select database folder for year 24
Public Sub SelectDbFolder24()
    valid_YearConfig.SelectDbFolder24
End Sub

' Select database folder for year 25
Public Sub SelectDbFolder25()
    valid_YearConfig.SelectDbFolder25
End Sub

' Select database folder for year 26
Public Sub SelectDbFolder26()
    valid_YearConfig.SelectDbFolder26
End Sub

' Ensure all database paths are configured for all supported years
' Returns True if all paths are valid, False otherwise
Public Function EnsureAllDbPathsConfigured() As Boolean
    MigrateLegacyConfigIfNeeded
    EnsureAllDbPathsConfigured = valid_YearConfig.EnsureAllDbPathsConfigured()
End Function

' Get array of supported years
Public Function GetSupportedYears() As Variant
    GetSupportedYears = valid_YearConfig.GetSupportedYears()
End Function

' ============================================================
' LEGACY MIGRATION
' ============================================================

' Migrate from old Kartei!X1 storage to new DBConfig sheet (one-time)
Private Sub MigrateLegacyConfigIfNeeded()
    If m_LegacyMigrated Then Exit Sub
    
    ' Check if cloud root is already configured
    If valid_YearConfig.GetCloudRoot() <> "" Then
        m_LegacyMigrated = True
        Exit Sub
    End If
    
    ' Try to migrate from Kartei!X1
    On Error Resume Next
    
    Dim wsKartei As Worksheet
    Set wsKartei = ThisWorkbook.Worksheets("Kartei")
    
    If wsKartei Is Nothing Then
        m_LegacyMigrated = True
        Exit Sub
    End If
    
    Dim oldPath As String
    oldPath = Trim(CStr(wsKartei.Range("X1").Value))
    
    If oldPath <> "" Then
        ' Run the migration
        valid_YearConfig.MigrateFromLegacyConfig
    End If
    
    m_LegacyMigrated = True
    On Error GoTo 0
End Sub
