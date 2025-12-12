Attribute VB_Name = "BulkComment"
'==========================
'   Module: BulkComment
'   Purpose: Bulk comment tool for Admin to apply the same comment to all changed records
'   Usage: Call EnableBulkComment before running CompareAndSyncKartei to skip individual Notitzen dialogs
'
'   Workflow:
'   1. Admin makes multiple changes (e.g., indexation of charges)
'   2. Admin calls EnableBulkCommentMode (or runs it via button/shortcut)
'   3. Admin enters a common comment once
'   4. Admin runs CompareAndSyncKartei (Save or manual)
'   5. All changed records receive the same comment without individual prompts
'   6. Mode automatically resets after sync completes
'==========================

Option Explicit

' ========================================
' Global State Variables
' ========================================

' Flag indicating bulk comment mode is active
Public BulkCommentModeActive As Boolean

' The comment text to use for all changes
Public BulkCommentText As String

' Counter for records processed in bulk mode
Public BulkCommentRecordCount As Long

' ========================================
' Public API
' ========================================

' Enable bulk comment mode - prompts user for comment and activates mode
Public Sub EnableBulkCommentMode()
    ' Reset state
    BulkCommentModeActive = False
    BulkCommentText = ""
    BulkCommentRecordCount = 0
    
    ' Show information about what this mode does
    Dim infoResult As VbMsgBoxResult
    infoResult = MsgBox("Der Massen-Kommentar-Modus ermoeglicht es, denselben Kommentar auf alle " & _
                       "geaenderten Datensaetze bei der naechsten Synchronisation anzuwenden." & vbCrLf & vbCrLf & _
                       "Dies ist nuetzlich, wenn Sie aehnliche Aenderungen an vielen Datensaetzen " & _
                       "vorgenommen haben (z.B. jaehrliche Indexierung der Gebuehren)." & vbCrLf & vbCrLf & _
                       "Nach Eingabe Ihres Kommentars fuehren Sie 'Speichern' oder 'Sync' aus, " & _
                       "um ihn auf alle Aenderungen anzuwenden." & vbCrLf & vbCrLf & _
                       "Moechten Sie den Massen-Kommentar-Modus aktivieren?", _
                       vbYesNo + vbQuestion, "Massen-Kommentar-Modus aktivieren")
    
    If infoResult <> vbYes Then
        Exit Sub
    End If
    
    ' Prompt for comment
    Dim userComment As String
    userComment = InputBox("Geben Sie den Kommentar ein, der auf ALLE geaenderten Datensaetze angewendet wird:" & vbCrLf & vbCrLf & _
                          "Beispiel: 'Indexierung 2025 - 3% Erhoehung'" & vbCrLf & vbCrLf & _
                          "(Dieser Kommentar wird fuer alle Datensaetze ohne individuelle Abfrage verwendet)", _
                          "Massen-Kommentar", "")
    
    ' Check if user provided a comment
    If Trim(userComment) = "" Then
        Dim confirmEmpty As VbMsgBoxResult
        confirmEmpty = MsgBox("Kein Kommentar eingegeben. Der Massen-Kommentar-Modus wird nicht aktiviert." & vbCrLf & vbCrLf & _
                             "Moechten Sie es erneut versuchen?", _
                             vbYesNo + vbQuestion, "Kein Kommentar eingegeben")
        
        If confirmEmpty = vbYes Then
            ' Recursive call to try again
            Call EnableBulkCommentMode
        End If
        Exit Sub
    End If
    
    ' Sanitize comment (remove problematic characters)
    userComment = SanitizeBulkComment(userComment)
    
    ' Activate mode
    BulkCommentModeActive = True
    BulkCommentText = userComment
    BulkCommentRecordCount = 0
    
    MsgBox "Massen-Kommentar-Modus AKTIVIERT" & vbCrLf & vbCrLf & _
           "Kommentar: """ & userComment & """" & vbCrLf & vbCrLf & _
           "Fuehren Sie jetzt 'Speichern' oder 'Sync' aus, um diesen Kommentar auf alle geaenderten Datensaetze anzuwenden." & vbCrLf & vbCrLf & _
           "Der Modus wird nach Abschluss der Synchronisation automatisch deaktiviert.", _
           vbInformation, "Massen-Kommentar-Modus aktiv"
End Sub

' Disable bulk comment mode manually (if user changes their mind)
Public Sub DisableBulkCommentMode()
    If Not BulkCommentModeActive Then
        MsgBox "Der Massen-Kommentar-Modus ist derzeit nicht aktiv.", vbInformation, "Modus-Status"
        Exit Sub
    End If
    
    Dim confirmResult As VbMsgBoxResult
    confirmResult = MsgBox("Sind Sie sicher, dass Sie den Massen-Kommentar-Modus deaktivieren moechten?" & vbCrLf & vbCrLf & _
                          "Sie werden waehrend der Synchronisation nach individuellen Kommentaren gefragt.", _
                          vbYesNo + vbQuestion, "Massen-Kommentar-Modus deaktivieren")
    
    If confirmResult = vbYes Then
        Call ResetBulkCommentMode
        MsgBox "Der Massen-Kommentar-Modus wurde deaktiviert.", vbInformation, "Modus deaktiviert"
    End If
End Sub

' Check current status of bulk comment mode
Public Sub ShowBulkCommentStatus()
    Dim statusMsg As String
    
    If BulkCommentModeActive Then
        statusMsg = "Massen-Kommentar-Modus: AKTIV" & vbCrLf & vbCrLf & _
                   "Kommentar: """ & BulkCommentText & """" & vbCrLf & vbCrLf & _
                   "Bisher verarbeitete Datensaetze: " & BulkCommentRecordCount
    Else
        statusMsg = "Massen-Kommentar-Modus: INAKTIV" & vbCrLf & vbCrLf & _
                   "Fuehren Sie 'EnableBulkCommentMode' aus, um zu aktivieren."
    End If
    
    MsgBox statusMsg, vbInformation, "Massen-Kommentar-Modus Status"
End Sub

' ========================================
' Internal Functions (called from other modules)
' ========================================

' Check if bulk comment mode is active
Public Function IsBulkCommentModeActive() As Boolean
    IsBulkCommentModeActive = BulkCommentModeActive
End Function

' Get the bulk comment text
Public Function GetBulkCommentText() As String
    GetBulkCommentText = BulkCommentText
End Function

' Increment the record counter (called after each record is processed)
Public Sub IncrementBulkCommentCounter()
    BulkCommentRecordCount = BulkCommentRecordCount + 1
End Sub

' Reset bulk comment mode (called after sync completes)
Public Sub ResetBulkCommentMode()
    Dim wasActive As Boolean
    Dim recordsProcessed As Long
    
    wasActive = BulkCommentModeActive
    recordsProcessed = BulkCommentRecordCount
    
    BulkCommentModeActive = False
    BulkCommentText = ""
    BulkCommentRecordCount = 0
    
    ' Only show message if mode was actually active and records were processed
    If wasActive And recordsProcessed > 0 Then
        MsgBox "Massen-Kommentar-Modus abgeschlossen." & vbCrLf & vbCrLf & _
               "Mit Massen-Kommentar verarbeitete Datensaetze: " & recordsProcessed, _
               vbInformation, "Massen-Kommentar abgeschlossen"
    End If
End Sub

' Show summary after sync (called from CompareAndSyncKartei)
Public Sub ShowBulkCommentSummary()
    If BulkCommentRecordCount > 0 Then
        ' Summary will be shown by ResetBulkCommentMode
    End If
End Sub

' ========================================
' Helper Functions
' ========================================

' Sanitize comment text for safe inclusion in history
Private Function SanitizeBulkComment(ByVal comment As String) As String
    Dim result As String
    result = comment
    
    ' Remove line breaks
    result = Replace(result, vbCrLf, " ")
    result = Replace(result, vbCr, " ")
    result = Replace(result, vbLf, " ")
    
    ' Remove characters that could break history format
    result = Replace(result, "/", " ")
    result = Replace(result, "||", "|")
    result = Replace(result, "->", "-")
    result = Replace(result, "/@", "/")
    result = Replace(result, "@/", "/")
    
    ' Trim whitespace
    result = Trim(result)
    
    ' Limit length to prevent excessively long history entries
    If Len(result) > 500 Then
        result = Left(result, 500) & "..."
    End If
    
    SanitizeBulkComment = result
End Function
