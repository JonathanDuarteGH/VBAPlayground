VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} dlgOptions 
   Caption         =   "Options"
   ClientHeight    =   7725
   ClientLeft      =   45
   ClientTop       =   345
   ClientWidth     =   6525
   HelpContextID   =   1838
   OleObjectBlob   =   "dlgOptions.frx":0000
   StartUpPosition =   1  'CenterOwner
   WhatsThisButton =   -1  'True
   WhatsThisHelp   =   -1  'True
End
Attribute VB_Name = "dlgOptions"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False



Option Explicit

Private Sub chkBounds_Click()
    Me.chkBoundsEvol = Me.chkBounds
End Sub

Private Sub chkBoundsEvol_Click()
    Me.chkBounds = Me.chkBoundsEvol
End Sub

Private Sub cmdCancel_Click()
    GetOptionSettings
    Me.Hide
End Sub

Private Sub cmdOK_Click()
    If Standard_ok Then
       SaveOptions
       Me.Hide
    End If
End Sub

Private Sub editConvergence_Change()
    Me.editConvEvol.Text = Me.editConvergence.Text
End Sub

Private Sub editConvEvol_Change()
    Me.editConvergence.Text = Me.editConvEvol.Text
End Sub

Private Sub editPopEvol_Change()
     Me.editPopSizeGRG.Text = Me.editPopEvol.Text
End Sub

Private Sub editPopSizeGRG_Change()
    Me.editPopEvol.Text = Me.editPopSizeGRG.Text
End Sub

Private Sub editSeedEvol_Change()
    Me.editSeedGRG.Text = Me.editSeedEvol.Text
End Sub

Private Sub editSeedGRG_Change()
    Me.editSeedEvol.Text = Me.editSeedGRG.Text
End Sub

Private Sub UserForm_Activate()
    Dim cctl As control
    
    Me.multiOptions.Value = 0

    Call fnUpdateDialogFonts(Me)
End Sub


Private Sub UserForm_Initialize()
    On Error Resume Next
    Me.editConvEvol.Text = ""
    Me.editConvergence.Text = ""
    Me.editIterations.Text = ""
    Me.editMaxSolutions.Text = ""
    Me.editMaxSubproblems.Text = ""
    Me.editMipGap.Text = ""
    Me.editMutate.Text = ""
    Me.editPopEvol.Text = ""
    Me.editPopSizeGRG.Text = ""
    Me.editPrecision.Text = ""
    Me.editSeedEvol.Text = ""
    Me.editSeedGRG.Text = ""
    Me.editTime.Text = ""
    Me.editTimeLimit.Text = ""
    On Error GoTo 0

    Me.Caption = GlobalX4Mess.Range("solv_dlg6_title").Text
    Me.multiOptions.Pages(0).Caption = GlobalX4Mess.Range("solv_dlg6_methods").Text
    Me.lblPrecision.Caption = GlobalX4Mess.Range("solv_dlg6_prec").Text
    Me.lblPrecision.Accelerator = GlobalX4Mess.Range("solv_dlg6_acc1").Text
    Me.chkScaling.Caption = GlobalX4Mess.Range("solv_dlg6_scale").Text
    Me.chkScaling.Accelerator = GlobalX4Mess.Range("solv_dlg6_acc2").Text
    Me.chkIterations.Caption = GlobalX4Mess.Range("solv_dlg6_iter").Text
    Me.chkIterations.Accelerator = GlobalX4Mess.Range("solv_dlg6_acc3").Text
    Me.frmInteger.Caption = GlobalX4Mess.Range("solv_dlg6_relax").Text
    Me.chkRelax.Caption = GlobalX4Mess.Range("solv_dlg6_ignore").Text
    Me.chkRelax.Accelerator = GlobalX4Mess.Range("solv_dlg6_acc4").Text
    Me.lblMipGap.Caption = GlobalX4Mess.Range("solv_dlg6_mipgap").Text
    Me.lblMipGap.Accelerator = GlobalX4Mess.Range("solv_dlg6_acc5").Text
    Me.frmEvol.Caption = GlobalX4Mess.Range("solv_dlg6_limits").Text
    Me.lblTime.Caption = GlobalX4Mess.Range("solv_dlg6_secs").Text
    Me.lblTime.Accelerator = GlobalX4Mess.Range("solv_dlg6_acc6").Text
    Me.lblIter.Caption = GlobalX4Mess.Range("solv_dlg6_iters").Text
    Me.lblIter.Accelerator = GlobalX4Mess.Range("solv_dlg6_acc7").Text
    Me.lblEvoLimits.Caption = GlobalX4Mess.Range("solv_dlg6_evol").Text
    Me.lblMaxSubProblems.Caption = GlobalX4Mess.Range("solv_dlg6_subs").Text
    Me.lblMaxSubProblems.Accelerator = GlobalX4Mess.Range("solv_dlg6_acc8").Text
    Me.lblMaxSolutions.Caption = GlobalX4Mess.Range("solv_dlg6_sols").Text
    Me.lblMaxSolutions.Accelerator = GlobalX4Mess.Range("solv_dlg6_acc9").Text
    Me.multiOptions.Pages(1).Caption = GlobalX4Mess.Range("solv_dlg6_grg").Text
    Me.lblConvergence.Caption = GlobalX4Mess.Range("solv_dlg6_conv").Text
    Me.lblConvergence.Accelerator = GlobalX4Mess.Range("solv_dlg6_acc10").Text
    Me.frameDerivs.Caption = GlobalX4Mess.Range("solv_dlg6_deriv").Text
    Me.radioForward.Caption = GlobalX4Mess.Range("solv_dlg6_fwd").Text
    Me.radioForward.Accelerator = GlobalX4Mess.Range("solv_dlg6_acc11").Text
    Me.radioCentral.Caption = GlobalX4Mess.Range("solv_dlg6_central").Text
    Me.radioCentral.Accelerator = GlobalX4Mess.Range("solv_dlg6_acc12").Text
    Me.frmMulti.Caption = GlobalX4Mess.Range("solv_dlg6_multi").Text
    Me.chkMultiStart.Caption = GlobalX4Mess.Range("solv_dlg6_usemult").Text
    Me.chkMultiStart.Accelerator = GlobalX4Mess.Range("solv_dlg6_acc13").Text
    Me.lblPopSizeGRG.Caption = GlobalX4Mess.Range("solv_dlg6_pop").Text
    Me.lblPopSizeGRG.Accelerator = GlobalX4Mess.Range("solv_dlg6_acc14").Text
    Me.lblSeedGRG.Caption = GlobalX4Mess.Range("solv_dlg6_seed").Text
    Me.lblSeedGRG.Accelerator = GlobalX4Mess.Range("solv_dlg6_acc15").Text
    Me.chkBounds.Caption = GlobalX4Mess.Range("solv_dlg6_reqbounds").Text
    Me.chkBounds.Accelerator = GlobalX4Mess.Range("solv_dlg6_acc16").Text
    Me.multiOptions.Pages(2).Caption = GlobalX4Mess.Range("solv_dlg6_evolu").Text
    Me.lblConvEvol.Caption = GlobalX4Mess.Range("solv_dlg6_evoconv").Text
    Me.lblConvEvol.Accelerator = GlobalX4Mess.Range("solv_dlg6_acc17").Text
    Me.lblMutation.Caption = GlobalX4Mess.Range("solv_dlg6_muta").Text
    Me.lblMutation.Accelerator = GlobalX4Mess.Range("solv_dlg6_acc18").Text
    Me.lblPopulationEvol.Caption = GlobalX4Mess.Range("solv_dlg6_popsize").Text
    Me.lblPopulationEvol.Accelerator = GlobalX4Mess.Range("solv_dlg6_acc19").Text
    Me.lblRandomSeed.Caption = GlobalX4Mess.Range("solv_dlg6_evoseed").Text
    Me.lblRandomSeed.Accelerator = GlobalX4Mess.Range("solv_dlg6_acc20").Text
    Me.lblTimeLimit.Caption = GlobalX4Mess.Range("solv_dlg6_maxtime").Text
    Me.lblTimeLimit.Accelerator = GlobalX4Mess.Range("solv_dlg6_acc21").Text
    Me.chkBoundsEvol.Caption = GlobalX4Mess.Range("solv_dlg6_evobounds").Text
    Me.chkBoundsEvol.Accelerator = GlobalX4Mess.Range("solv_dlg6_acc22").Text
    Me.cmdOK.Caption = GlobalX4Mess.Range("solv_dlg6_ok").Text
    Me.cmdOK.Accelerator = GlobalX4Mess.Range("solv_dlg6_acc23").Text
    Me.cmdCancel.Caption = GlobalX4Mess.Range("solv_dlg6_cancel").Text
    Me.cmdCancel.Accelerator = GlobalX4Mess.Range("solv_dlg6_acc24").Text
End Sub

Private Sub UserForm_QueryClose(Cancel As Integer, CloseMode As Integer)
    GetOptionSettings
End Sub


