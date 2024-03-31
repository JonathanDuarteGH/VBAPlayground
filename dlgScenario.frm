VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} dlgScenario 
   Caption         =   "Save Scenario"
   ClientHeight    =   1800
   ClientLeft      =   45
   ClientTop       =   345
   ClientWidth     =   4380
   OleObjectBlob   =   "dlgScenario.frx":0000
   StartUpPosition =   1  'CenterOwner
End
Attribute VB_Name = "dlgScenario"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False


Option Explicit

Private Sub cmdCancel_Click()
    Me.Hide
End Sub

Private Sub cmdOK_Click()
    ' user clicks OK in scenario dialog
    Dim i As Integer
    If Range("solver_adj").Cells.Count > 32 Then
       DisplayMessage "solver_msg_25", 1830, 0
       Exit Sub
    End If
    If Me.editScenario.Text = "" Then
       DisplayMessage "solver_msg_24a", 1830, 0
       Exit Sub
    End If
    On Error GoTo nameconflict
    ActiveSheet.Scenarios.Add Name:=Me.editScenario.Text, ChangingCells:=Range("solver_adj"), Locked:=False
    On Error GoTo 0
    Me.Hide
    Exit Sub
nameconflict:
    DisplayMessage "solver_msg_24b", 1830, 0
    On Error GoTo 0
End Sub

Private Sub UserForm_Activate()
    Me.editScenario.Text = ""
End Sub

Private Sub UserForm_Initialize()
    Me.editScenario.Text = ""

    Dim stzFont As String
    Dim cctl As control

    Me.Caption = GlobalX4Mess.Range("solv_dlg7_title").Text
    Me.lblScenario.Caption = GlobalX4Mess.Range("solv_dlg7_scen").Text
    Me.lblScenario.Accelerator = GlobalX4Mess.Range("solv_dlg7_acc1").Text
    Me.cmdOK.Caption = GlobalX4Mess.Range("solv_dlg7_ok").Text
    Me.cmdOK.Accelerator = GlobalX4Mess.Range("solv_dlg7_acc2").Text
    Me.cmdCancel.Caption = GlobalX4Mess.Range("solv_dlg7_cancel").Text
    Me.cmdCancel.Accelerator = GlobalX4Mess.Range("solv_dlg7_acc3").Text
    
    Call fnUpdateDialogFonts(Me)
End Sub
