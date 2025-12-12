Attribute VB_Name = "ExportMessagePlanung"
Option Explicit

Public nextTime As Date
Public messageDisplayed As Boolean
Public messageText As String

' Starts the hourly message scheduling.
Public Sub StartHourlyMessage()
    messageText = "Sie arbeiten seit mehr als 60 Minuten ohne Synchronisierung mit der Datenbank. Es wird dringend empfohlen, die Synchronisierung zu starten (Taste 'Zu Base')."
    ScheduleNextMessage
End Sub

' Stops the scheduled hourly message.
Public Sub StopHourlyMessage()
    On Error Resume Next
    Application.OnTime EarliestTime:=nextTime, Procedure:="ShowHourlyMessage", Schedule:=False
    On Error GoTo 0
End Sub

' Schedules the next message display in one hour.
Public Sub ScheduleNextMessage()
    nextTime = Now + TimeValue("01:00:00")
    Application.OnTime EarliestTime:=nextTime, Procedure:="ShowHourlyMessage", Schedule:=True
End Sub

' Displays the hourly message if no other user forms are open and no message is currently shown.
Public Sub ShowHourlyMessage()
    ' Schedule the next message before attempting to show the current one.
    ScheduleNextMessage
    
    ' If any userform is open, then exit to avoid conflicts.
    If UserForms.count > 0 Then Exit Sub
    
    ' Check if a message is already being displayed.
    If messageDisplayed Then Exit Sub
    
    messageDisplayed = True
    
    ' Show the message in a modeless user form.
    frmHourlyMessage.Show vbModeless
End Sub
