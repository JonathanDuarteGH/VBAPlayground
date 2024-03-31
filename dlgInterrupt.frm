VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} dlgInterrupt 
   Caption         =   "Show Trial Solution"
   ClientHeight    =   1755
   ClientLeft      =   45
   ClientTop       =   345
   ClientWidth     =   5700
   OleObjectBlob   =   "dlgInterrupt.frx":0000
   StartUpPosition =   1  'CenterOwner
End
Attribute VB_Name = "dlgInterrupt"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Option Explicit
Private Sub cmdContinue_Click()
    GlobalContinuation = 0
    Unload Me
End Sub

Private Sub cmdScenario_Click()
    Me.Hide
    dlgScenario.Show
    Me.Show
End Sub

Private Sub cmdStop_Click()
    GlobalContinuation = 1
    Unload Me
End Sub

Private Sub UserForm_Initialize()
    Dim stzFont As String
    Dim cctl As control

    Me.Caption = GlobalX4Mess.Range("solv_dlg3_title").Text
    Me.cmdContinue.Caption = GlobalX4Mess.Range("solv_dlg3_cont").Text
    Me.cmdContinue.Accelerator = GlobalX4Mess.Range("solv_dlg3_acc1").Text
    Me.cmdStop.Caption = GlobalX4Mess.Range("solv_dlg3_stop").Text
    Me.cmdStop.Accelerator = GlobalX4Mess.Range("solv_dlg3_acc2").Text
    Me.cmdScenario.Caption = GlobalX4Mess.Range("solv_dlg3_scen").Text
    Me.cmdScenario.Accelerator = GlobalX4Mess.Range("solv_dlg3_acc3").Text
    
    Call fnUpdateDialogFonts(Me)
End Sub
