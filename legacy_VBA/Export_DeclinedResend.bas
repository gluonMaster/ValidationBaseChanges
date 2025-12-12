Attribute VB_Name = "Export_DeclinedResend"
'==========================
'   Module: Export_DeclinedResend
'   Purpose: Admin tools to resend DECLINED records for re-verification
'==========================
Option Explicit

' ============================================================================
' PUBLIC MACROS
' ============================================================================

Public Sub Export_ResendAllDeclined()
    ' Resends all DECLINED records from Kartei sheet for re-verification.
    ' Collects all rows with BA="DECLINED" and delegates to ExportSyncKartei.
    
    ' Verify active sheet is Kartei
    If ActiveSheet.Name <> "Kartei" Then
        MsgBox "Bitte oeffnen Sie zuerst das Blatt 'Kartei'.", vbExclamation, "Abgelehnte wieder senden"
        Exit Sub
    End If
    
    Dim ws As Worksheet
    Set ws = ThisWorkbook.Worksheets("Kartei")
    
    ' Collect all declined IDs
    Dim ids As Collection
    Set ids = Export_CollectDeclinedIDs_All(ws)
    
    ' Delegate to internal handler
    Call Export_ResendDeclinedInternal(ws, ids, True)
End Sub

Public Sub Export_ResendSelectedDeclined()
    ' Resends only selected DECLINED records from Kartei sheet for re-verification.
    ' Collects IDs from selected rows with BA="DECLINED" and delegates to ExportSyncKartei.
    
    ' Verify active sheet is Kartei
    If ActiveSheet.Name <> "Kartei" Then
        MsgBox "Bitte oeffnen Sie zuerst das Blatt 'Kartei'.", vbExclamation, "Abgelehnte wieder senden"
        Exit Sub
    End If
    
    Dim ws As Worksheet
    Set ws = ThisWorkbook.Worksheets("Kartei")
    
    ' Collect declined IDs from selection
    Dim ids As Collection
    Set ids = Export_CollectDeclinedIDs_Selected(ws)
    
    ' Delegate to internal handler
    Call Export_ResendDeclinedInternal(ws, ids, False)
End Sub

' ============================================================================
' PRIVATE HELPERS
' ============================================================================

Private Function Export_CollectDeclinedIDs_All(ws As Worksheet) As Collection
    ' Collects all IDs from rows with BA="DECLINED" on the given worksheet.
    ' Returns a Collection of ID strings (no duplicates).
    
    Const STATUS_COL As Long = 53  ' BA column
    Const ID_COL As Long = 48      ' AV column
    
    Dim result As New Collection
    Dim dictSeen As New Scripting.Dictionary
    
    Dim lastRow As Long
    lastRow = ws.Cells(ws.Rows.Count, 1).End(xlUp).Row
    
    Dim r As Long
    Dim status As String
    Dim idVal As Variant
    Dim idKey As String
    
    For r = 3 To lastRow
        status = Trim$(CStr(ws.Cells(r, STATUS_COL).Value))
        idVal = ws.Cells(r, ID_COL).Value
        
        If status = "DECLINED" Then
            If IsNumeric(idVal) Then
                If CLng(idVal) > 0 Then
                    idKey = CStr(CLng(idVal))
                    
                    ' Avoid duplicates
                    If Not dictSeen.Exists(idKey) Then
                        dictSeen.Add idKey, True
                        result.Add idKey
                    End If
                End If
            End If
        End If
    Next r
    
    Set Export_CollectDeclinedIDs_All = result
End Function

Private Function Export_CollectDeclinedIDs_Selected(ws As Worksheet) As Collection
    ' Collects IDs from selected rows with BA="DECLINED" on the given worksheet.
    ' Iterates through all areas in the selection.
    ' Returns a Collection of ID strings (no duplicates).
    
    Const STATUS_COL As Long = 53  ' BA column
    Const ID_COL As Long = 48      ' AV column
    
    Dim result As New Collection
    Dim dictSeen As New Scripting.Dictionary
    
    Dim area As Range
    Dim rowIndex As Long
    Dim status As String
    Dim idVal As Variant
    Dim idKey As String
    
    For Each area In Selection.Areas
        For rowIndex = area.Row To area.Row + area.Rows.Count - 1
            ' Skip header rows
            If rowIndex >= 3 Then
                status = Trim$(CStr(ws.Cells(rowIndex, STATUS_COL).Value))
                idVal = ws.Cells(rowIndex, ID_COL).Value
                
                If status = "DECLINED" Then
                    If IsNumeric(idVal) Then
                        If CLng(idVal) > 0 Then
                            idKey = CStr(CLng(idVal))
                            
                            ' Avoid duplicates
                            If Not dictSeen.Exists(idKey) Then
                                dictSeen.Add idKey, True
                                result.Add idKey
                            End If
                        End If
                    End If
                End If
            End If
        Next rowIndex
    Next area
    
    Set Export_CollectDeclinedIDs_Selected = result
End Function

Private Sub Export_ResendDeclinedInternal(ws As Worksheet, ids As Collection, ByVal isAll As Boolean)
    ' Internal handler for resending declined records.
    ' Validates collection, prompts for confirmation, and delegates to ExportSyncKartei.
    
    ' Check if any IDs were collected
    If ids Is Nothing Then
        If isAll Then
            MsgBox "Keine abgelehnten Datensaetze fuer Wiederholung gefunden.", vbInformation, "Abgelehnte wieder senden"
        Else
            MsgBox "Keine geeigneten abgelehnten Zeilen in der Auswahl gefunden.", vbInformation, "Abgelehnte wieder senden"
        End If
        Exit Sub
    End If
    
    If ids.Count = 0 Then
        If isAll Then
            MsgBox "Keine abgelehnten Datensaetze fuer Wiederholung gefunden.", vbInformation, "Abgelehnte wieder senden"
        Else
            MsgBox "Keine geeigneten abgelehnten Zeilen in der Auswahl gefunden.", vbInformation, "Abgelehnte wieder senden"
        End If
        Exit Sub
    End If
    
    ' Confirmation prompt
    Dim confirmText As String
    Dim response As VbMsgBoxResult
    
    If isAll Then
        confirmText = "Moechten Sie alle " & ids.Count & " abgelehnten Datensaetze zur erneuten Pruefung an den Superadmin senden?"
    Else
        confirmText = "Moechten Sie die " & ids.Count & " ausgewaehlten abgelehnten Datensaetze zur erneuten Pruefung an den Superadmin senden?"
    End If
    
    response = MsgBox(confirmText, vbYesNo + vbQuestion, "Abgelehnte wieder senden")
    
    If response <> vbYes Then
        Exit Sub
    End If
    
    ' Delegate to ExportSyncKartei helper
    On Error GoTo ErrorHandler
    Call ExportSyncKartei.Export_ResendDeclinedRecordsByIDs(ids)
    On Error GoTo 0
    
    ' Success message
    MsgBox "Ausgewaehlte Datensaetze wurden zur erneuten Pruefung gesendet.", vbInformation, "Abgelehnte wieder senden"
    Exit Sub
    
ErrorHandler:
    MsgBox "Fehler beim Senden der Datensaetze: " & Err.Description, vbCritical, "Abgelehnte wieder senden"
End Sub
