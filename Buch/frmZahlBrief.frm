VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} frmZahlBrief 
   Caption         =   "Zahlung f�r Tabelle und Brief"
   ClientHeight    =   4584
   ClientLeft      =   48
   ClientTop       =   372
   ClientWidth     =   6960
   OleObjectBlob   =   "frmZahlBrief.frx":0000
   StartUpPosition =   1  'CenterOwner
End
Attribute VB_Name = "frmZahlBrief"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
' v. 2.2
' Process button click event - All years (2023-2026)
Private Sub btnOK_Click()
    On Error GoTo Cleanup
    Application.ScreenUpdating = False
    Application.Calculation = xlCalculationManual

    ' Clear report area and build report for all years (2023-2026)
    kosten_Report.Kosten_ClearReportArea
    kosten_Report.Kosten_BuildReport_AllYears
    
Cleanup:
    Application.Calculation = xlCalculationAutomatic
    Application.ScreenUpdating = True
    
    If Err.Number <> 0 Then
        MsgBox "Fehler aufgetreten: " & Err.Description, vbCritical, "Fehler"
    End If
End Sub

' Process button click event - Recent years only (2024-2026)
Private Sub btnLast3Years_Click()
    On Error GoTo Cleanup
    Application.ScreenUpdating = False
    Application.Calculation = xlCalculationManual

    ' Clear report area and build report for recent years (2024-2026)
    kosten_Report.Kosten_ClearReportArea
    kosten_Report.Kosten_BuildReport_RecentYears
    
Cleanup:
    Application.Calculation = xlCalculationAutomatic
    Application.ScreenUpdating = True
    
    If Err.Number <> 0 Then
        MsgBox "Fehler aufgetreten: " & Err.Description, vbCritical, "Fehler"
    End If
End Sub

' List box click event
Private Sub ListBox1_Click()
    UserFormHandler.HandleListBoxClick Me.ListBox1, Me.txtBox_Eltern, Me.txtBox_Konto, _
                                     Me.txtBox_Adresse, GlobalVariables.nameEltern, _
                                     GlobalVariables.Konto, GlobalVariables.Adresse
End Sub

' Text box change event
Private Sub txtBox_Eltern_Change()
    UserFormHandler.HandleTextBoxElternChange Me.ListBox1, Me.txtBox_Eltern, GlobalVariables.elternLetzte
End Sub

' Form initialization event
Private Sub UserForm_Activate()
    ' Show status indicator
    Me.Caption = "Payment Letter Generation - Loading data..."
    
    ' Start loading data with a slight delay to allow form to show
    Application.OnTime Now + TimeValue("00:00:01"), "InitializeFormDataWithStatus"
End Sub



