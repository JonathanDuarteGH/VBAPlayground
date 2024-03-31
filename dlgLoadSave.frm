VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} dlgLoadSave 
   Caption         =   "Load/Save Model"
   ClientHeight    =   2280
   ClientLeft      =   45
   ClientTop       =   345
   ClientWidth     =   4800
   OleObjectBlob   =   "dlgLoadSave.frx":0000
   StartUpPosition =   1  'CenterOwner
   WhatsThisButton =   -1  'True
   WhatsThisHelp   =   -1  'True
End
Attribute VB_Name = "dlgLoadSave"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False



Option Explicit

Private Sub cmdCancel_Click()
    Me.Hide
End Sub

Private Sub cmdLoad_Click()
    Dim vbacode As String
    GlobalMerger = 2
    With dlgSolverParameters
        If .listConstraints.ListCount <> 0 Or .refVariables.Text <> "" Then
           GlobalMerger = 0
           dlgMerge.Show
        End If
    End With
    Select Case GlobalMerger
       Case 1
          Loadmod
          vbacode = GlobalX4Mess.Range("loadfunc").Text & " " & GlobalX4Mess.Range("loadarg1").Text & _
                ":=" & Chr(34) & Me.refArea.Text & Chr(34) & ", " & GlobalX4Mess.Range("loadarg2").Text & _
                "=True"
          Application.RecordMacro basiccode:=vbacode
       Case 2
          Reset_all (False)
          Loadmod
          vbacode = GlobalX4Mess.Range("loadfunc").Text & " " & GlobalX4Mess.Range("loadarg1").Text & ":=" & Chr(34) & Me.refArea.Text & Chr(34)
          Application.RecordMacro basiccode:=vbacode
    End Select
    GlobalMerger = 2
    Me.Hide
End Sub

Private Sub cmdSave_Click()
    If Not SaveMod Then
       Exit Sub
    End If
    Me.Hide
End Sub

Private Sub UserForm_Activate()
    Dim cctl As control
    Dim stzFont As String

    Me.lblArea2.Caption = GlobalX4Mess.Range("solv_dlg4_select1").Text & _
        (4 + dlgSolverParameters.listConstraints.ListCount)
  
    Me.refArea.Text = Selection.Address(ReferenceStyle:=GlobalR1C1)
    Me.refArea.SetFocus
    
    Call fnUpdateDialogFonts(Me)
End Sub

Private Sub UserForm_Initialize()
    Me.refArea.Text = ""
    Me.Caption = GlobalX4Mess.Range("solv_dlg4_title").Text
    Me.lblArea.Caption = GlobalX4Mess.Range("solv_dlg4_select").Text
    Me.lblArea.Accelerator = GlobalX4Mess.Range("solv_dlg4_acc1").Text
    Me.cmdLoad.Caption = GlobalX4Mess.Range("solv_dlg4_load").Text
    Me.cmdLoad.Accelerator = GlobalX4Mess.Range("solv_dlg4_acc2").Text
    Me.cmdSave.Caption = GlobalX4Mess.Range("solv_dlg4_save").Text
    Me.cmdSave.Accelerator = GlobalX4Mess.Range("solv_dlg4_acc3").Text
    Me.cmdCancel.Caption = GlobalX4Mess.Range("solv_dlg4_cancel").Text
    Me.cmdCancel.Accelerator = GlobalX4Mess.Range("solv_dlg4_acc4").Text
End Sub


