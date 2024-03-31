VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} dlgMerge 
   Caption         =   "Load Model"
   ClientHeight    =   1605
   ClientLeft      =   45
   ClientTop       =   345
   ClientWidth     =   4830
   OleObjectBlob   =   "dlgMerge.frx":0000
   StartUpPosition =   1  'CenterOwner
End
Attribute VB_Name = "dlgMerge"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Option Explicit

Private Sub cmdCancel_Click()
    GlobalMerger = 0
    Me.Hide
End Sub

Private Sub cmdMerge_Click()
    GlobalMerger = 1
    Me.Hide
End Sub

Private Sub cmdReplace_Click()
    GlobalMerger = 2
    Me.Hide
End Sub


Private Sub UserForm_Initialize()
    Dim stzFont As String
    Dim cctl As control
    
    Me.Caption = GlobalX4Mess.Range("solv_dlg5_title").Text
    Me.lblMessage.Caption = GlobalX4Mess.Range("solv_dlg5_quest").Text
    Me.cmdReplace.Caption = GlobalX4Mess.Range("solv_dlg5_replace").Text
    Me.cmdReplace.Accelerator = GlobalX4Mess.Range("solv_dlg5_acc1").Text
    Me.cmdMerge.Caption = GlobalX4Mess.Range("solv_dlg5_merge").Text
    Me.cmdMerge.Accelerator = GlobalX4Mess.Range("solv_dlg5_acc2").Text
    Me.cmdCancel.Caption = GlobalX4Mess.Range("solv_dlg5_cancel").Text
    Me.cmdCancel.Accelerator = GlobalX4Mess.Range("solv_dlg5_acc3").Text
    
    Call fnUpdateDialogFonts(Me)
End Sub

