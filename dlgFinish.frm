VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} dlgFinish 
   Caption         =   "Solver Results"
   ClientHeight    =   5295
   ClientLeft      =   45
   ClientTop       =   345
   ClientWidth     =   8655.001
   OleObjectBlob   =   "dlgFinish.frx":0000
   StartUpPosition =   1  'CenterOwner
End
Attribute VB_Name = "dlgFinish"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False


Option Explicit
Private Const frameMargin As Integer = 6
Dim bIntegers As Boolean
Dim bWarning As Boolean
Dim strWarningTitle As String
Dim strWarning As String

Private Sub ResizeLblHlp(bWarningVisible As Boolean)
    With Me
        .imgWarning.Visible = bWarningVisible
        If bWarningVisible Then
           .lblHelpTitle.Left = frameMargin + .imgWarning.Width
           .lblHelpLP.Left = frameMargin + .imgWarning.Width
        Else
           .lblHelpTitle.Left = frameMargin
           .lblHelpLP.Left = frameMargin
        End If
        .lblHelpTitle.Width = .frameHelpLP.Width - .lblHelpTitle.Left - frameMargin
        .lblHelpLP.Width = .frameHelpLP.Width - .lblHelpLP.Left - frameMargin
    End With
End Sub

Private Sub LayoutRadio(bRelaxVisible As Boolean)
    Dim radioTotalHeight As Integer, frameKeepGap As Integer

    With Me
        .radioRelax.Visible = bRelaxVisible

        frameKeepGap = .frameKeep.Height - 2 * frameMargin

        radioTotalHeight = .radioKeep.Height + .radioRestore.Height
        If bRelaxVisible Then radioTotalHeight = radioTotalHeight + .radioRelax.Height

        .radioKeep.Top = frameMargin + _
                        .radioKeep.Height / 2 * _
                        (frameKeepGap - radioTotalHeight) / _
                        radioTotalHeight

        .radioRestore.Top = frameMargin - .radioRestore.Height / 2 + _
                        (.radioKeep.Height + .radioRestore.Height / 2) * _
                        frameKeepGap / radioTotalHeight

        If bRelaxVisible Then
            .radioRelax.Top = frameMargin - .radioRelax.Height / 2 + _
                            (radioTotalHeight - .radioRelax.Height / 2) * _
                            frameKeepGap / radioTotalHeight
        End If
    End With
End Sub

Private Sub cmdCancel_Click()
    Action = "Cancel"
    If SolverCls Is Nothing Then
        Set SolverCls = New SolverCalls
    End If
    Call SolverCls.Solve(2)
    Call SolverCls.Solve(1)
    Unhook
    Me.Hide
End Sub

Private Sub cmdOK_Click()
    Dim DoManualRestore As Boolean, c, bRelax As Boolean
    Action = "Cancel"
    DoManualRestore = False
    bRelax = False
    If SolverCls Is Nothing Then
        Set SolverCls = New SolverCalls
    End If
    With Me
        If .radioRelax Then
            bRelax = True
            dlgOptions.chkRelax = True
            SaveOptions
            Me.Hide
            Call SolverCls.Solve(2)
            Call SolverCls.Solve(1)
            GlobalAnswer = SolverCls.Solve(0)
            Me.Show
            dlgOptions.chkRelax = False
            SaveOptions
            Exit Sub
        End If
        .Hide
        If .listReports.Enabled Then
            GlobalOutline = .chkOutline
            If ThisWorkbook.DialogSheets("Solver_dialog").DropDowns("EngineList").ListIndex <> 3 Then
                If GlobalAnswer = 5 Then
                    If .listReports.Selected(0) Then
                        Do_Infeas (1)
                        DoManualRestore = True
                    End If
                    If .listReports.ListCount > 1 Then
                        If .listReports.Selected(1) Then
                            Do_Infeas (2)
                            DoManualRestore = True
                        End If
                    End If
                ElseIf GlobalAnswer = 7 And .listReports.Selected(0) Then
                    Do_Linear
                Else
                    If .listReports.Selected(0) Then
                        Do_Answer
                    End If
                    If .listReports.ListCount > 1 Then
                        If .listReports.Selected(1) And ((Not bIntegers) Or dlgOptions.chkRelax) Then
                            Do_Sensitivity
                        End If
                    End If
                    If .listReports.ListCount > 2 Then
                        If .listReports.Selected(2) And ((Not bIntegers) Or dlgOptions.chkRelax) Then
                            Do_Limits
                            DoManualRestore = True
                        End If
                    End If
                End If
            Else
                If .listReports.Selected(0) Then
                    Do_Answer
                End If
                If .listReports.Selected(1) Then
                    Do_Pop
                End If
            End If
        End If
        If SolverCls Is Nothing Then
            Set SolverCls = New SolverCalls
        End If
        If DoManualRestore = False Then
            If .radioKeep Then
                Call SolverCls.Solve(1)
            Else
                Call SolverCls.Solve(2)
                Call SolverCls.Solve(1)
            End If
        Else
            Call SolverCls.Solve(1)
            If Not .radioKeep Then
                Dim i As Integer
                i = 1
                For Each c In Range("solver_adj").Cells
                    c.Value = GlobalOldVars(i)
                    i = i + 1
                Next
            End If
        End If
        Unhook
        If .chkReturn Then
            If bRelax Then
                dlgOptions.chkRelax = False
                SaveOptions
            End If
           Action = ""
        End If
    End With
End Sub

Private Sub cmdScenario_Click()
    Me.Hide
    dlgScenario.Show
    Me.Show
End Sub

Private Sub lblResult_Click()
    ResizeLblHlp bWarning
    Me.lblHelpTitle.Caption = strWarningTitle
    Me.lblHelpLP.Caption = strWarning
End Sub

Private Sub listReports_Enter()
    ResizeLblHlp False
    Me.lblHelpTitle.Caption = GlobalX4Mess.Range("solver_hlp_reports1").Text
    Me.lblHelpLP.Caption = GlobalX4Mess.Range("solver_hlp_reports1a").Text
End Sub

Private Sub radioKeep_Click()
    ResizeLblHlp False
    Me.lblHelpTitle.Caption = GlobalX4Mess.Range("solver_hlp_keep1").Text
    Me.lblHelpLP.Caption = GlobalX4Mess.Range("solver_hlp_keep1a").Text
End Sub

Private Sub radioRelax_Click()
    ResizeLblHlp False
    Me.lblHelpTitle.Caption = GlobalX4Mess.Range("solver_hlp_without1").Text
    Me.lblHelpLP.Caption = GlobalX4Mess.Range("solver_hlp_without1a").Text
End Sub

Private Sub radioRestore_Click()
    ResizeLblHlp False
    Me.lblHelpTitle.Caption = GlobalX4Mess.Range("solver_hlp_restore1").Text
    Me.lblHelpLP.Caption = GlobalX4Mess.Range("solver_hlp_restore1a").Text
End Sub

Private Sub UserForm_Activate()
 ' prepare and display finish dialog
    Dim Msg As String, dmy As Integer, i As Integer
    Dim c, tmp As Integer, tmp2 As Boolean

    With Me
        .listReports.Clear
        .listReports.Enabled = True
        LayoutRadio False
        
        bIntegers = False
        bWarning = False
        For i = 1 To CDbl(Mid(CStr(ActiveSheet.Names("solver_num")), 2))
            dmy = CDbl(Mid(CStr(ActiveSheet.Names("solver_rel" & Trim(Str(i)))), 2))
            If dmy > 3 Then
                bIntegers = True
                Exit For
            End If
        Next
        If CDbl(Mid(CStr(ActiveSheet.Names("solver_eng")), 2)) <> 3 Then
            Select Case GlobalAnswer
            Case 0, 1, 2, 14
                If (Not bIntegers) Or dlgOptions.chkRelax Then
                    .listReports.List = _
                        Array(GlobalX4Mess.Range("solver_msg_26").Text, _
                            GlobalX4Mess.Range("solver_msg_27").Text, _
                            GlobalX4Mess.Range("solver_msg_28").Text)
                Else
                    .listReports.List = Array(GlobalX4Mess.Range("solver_msg_26").Text)
                End If
            Case 5
                If (Not bIntegers) Or dlgOptions.chkRelax Then
                    .listReports.List = Array(GlobalX4Mess.Range("solver_msg_28b").Text, _
                                       GlobalX4Mess.Range("solver_msg_28c").Text)
          
                Else
                    .listReports.Clear
                    .listReports.Enabled = False
                    If Not dlgOptions.chkRelax Then
                        ' show third radiobutton
                        LayoutRadio True
                    End If
                End If
            Case 7
                .listReports.List = Array(GlobalX4Mess.Range("solver_msg_28a").Text)
            Case 3, 6, 10, 15, 16, 17
                .listReports.List = Array(GlobalX4Mess.Range("solver_msg_26").Text)
            Case Else
                .listReports.Clear
                .listReports.Enabled = False
            End Select
        Else
            If GlobalAnswer <= 17 Then
                .listReports.List = _
                 Array(GlobalX4Mess.Range("solver_msg_26").Text, _
                       GlobalX4Mess.Range("solver_msg_28d").Text)
            Else
                .listReports.Clear
                .listReports.Enabled = False
            End If
        End If
        .radioKeep = True
        On Error GoTo 0
        Select Case GlobalAnswer
        Case 0
            Msg = GlobalX4Mess.Range("solver_msg_29").Text
            bWarning = False
            .lblHelpLP.Caption = ""
        Case 1
            Msg = GlobalX4Mess.Range("solver_msg_30").Text
            bWarning = False
        Case 2
            Msg = GlobalX4Mess.Range("solver_msg_31").Text
            bWarning = False
        Case 3
            Msg = GlobalX4Mess.Range("solver_msg_32").Text
            bWarning = False
        Case 4
            Msg = GlobalX4Mess.Range("solver_msg_33").Text
            bWarning = True
        Case 5
            Msg = GlobalX4Mess.Range("solver_msg_34").Text
            bWarning = True
        Case 6
            Msg = GlobalX4Mess.Range("solver_msg_35").Text
            bWarning = False
        Case 7
            Msg = GlobalX4Mess.Range("solver_msg_37").Text
            bWarning = True
        Case 8
            Msg = GlobalX4Mess.Range("solver_msg_38").Text
            bWarning = True
        Case 9
            Msg = GlobalX4Mess.Range("solver_msg_39").Text
            bWarning = True
        Case 10
            Msg = GlobalX4Mess.Range("solver_msg_40").Text
            bWarning = False
        Case 11
            Msg = GlobalX4Mess.Range("solver_msg_41").Text
            bWarning = True
        Case 12
            Msg = GlobalX4Mess.Range("solver_msg_42").Text
            bWarning = True
        Case 13
            Msg = GlobalX4Mess.Range("solver_msg_43").Text
            bWarning = True
        Case 14
            Msg = GlobalX4Mess.Range("solver_msg_29a").Text
            bWarning = False
        Case 15
            If CDbl(Mid(CStr(ActiveSheet.Names("solver_eng")), 2)) = 3 Then
                Msg = GlobalX4Mess.Range("solver_stop3").Text
            Else
                Msg = GlobalX4Mess.Range("solver_stop2").Text
            End If
            bWarning = False
        Case 16
            If CDbl(Mid(CStr(ActiveSheet.Names("solver_eng")), 2)) = 3 Then
                Msg = GlobalX4Mess.Range("solver_stop4").Text
            Else
                Msg = GlobalX4Mess.Range("solver_stop1").Text
            End If
            bWarning = False
        Case 17
            Msg = GlobalX4Mess.Range("solver_result17").Text
            bWarning = False
        Case 18
            Msg = GlobalX4Mess.Range("solver_result18").Text
            bWarning = True
        Case 19
            Msg = GlobalX4Mess.Range("solver_result19").Text
            bWarning = True
        Case 20
            Msg = GlobalX4Mess.Range("solver_result20").Text
            bWarning = True
        Case Else
            Msg = GlobalX4Mess.Range("solver_msg_44").Text
            bWarning = True
        End Select
        ResizeLblHlp bWarning
        strWarningTitle = Msg
        If GlobalAnswer <> 15 And GlobalAnswer <> 16 Then
           strWarning = GlobalX4Mess.Range("solver_exp" & Trim(Str(GlobalAnswer))).Text
        Else
            If CDbl(Mid(CStr(ActiveSheet.Names("solver_eng")), 2)) = 3 Then
               strWarning = GlobalX4Mess.Range("solver_exp" & Trim(Str(GlobalAnswer))).Text
            Else
               strWarning = GlobalX4Mess.Range("solver_exp" & Trim(Str(GlobalAnswer)) & "a").Text
            End If
        End If
        .lblHelpTitle.Caption = strWarningTitle
        .lblHelpLP.Caption = strWarning
        .lblResult = Msg
        Beep
    End With
End Sub

Private Sub UserForm_Initialize()
    Me.chkOutline = False
    Me.chkReturn = False
    Me.Caption = GlobalX4Mess.Range("solv_dlg2_title").Text
    Me.radioKeep.Caption = GlobalX4Mess.Range("solv_dlg2_keep").Text
    Me.radioKeep.Accelerator = GlobalX4Mess.Range("solv_dlg2_acc1").Text
    Me.radioRestore.Caption = GlobalX4Mess.Range("solv_dlg2_restore").Text
    Me.radioRestore.Accelerator = GlobalX4Mess.Range("solv_dlg2_acc2").Text
    Me.radioRelax.Caption = GlobalX4Mess.Range("solv_dlg2_relax").Text
    Me.radioRelax.Accelerator = GlobalX4Mess.Range("solv_dlg2_acc3").Text
    Me.chkReturn.Caption = GlobalX4Mess.Range("solv_dlg2_return").Text
    Me.chkReturn.Accelerator = GlobalX4Mess.Range("solv_dlg2_acc4").Text
    Me.lblReports.Caption = GlobalX4Mess.Range("solv_dlg2_reports").Text
    Me.lblReports.Accelerator = GlobalX4Mess.Range("solv_dlg2_acc5").Text
    Me.chkOutline.Caption = GlobalX4Mess.Range("solv_dlg2_outline").Text
    Me.chkOutline.Accelerator = GlobalX4Mess.Range("solv_dlg2_acc6").Text
    Me.cmdOK.Caption = GlobalX4Mess.Range("solv_dlg2_ok").Text
    Me.cmdOK.Accelerator = GlobalX4Mess.Range("solv_dlg2_acc7").Text
    Me.cmdCancel.Caption = GlobalX4Mess.Range("solv_dlg2_cancel").Text
    Me.cmdCancel.Accelerator = GlobalX4Mess.Range("solv_dlg2_acc8").Text
    Me.cmdScenario.Caption = GlobalX4Mess.Range("solv_dlg2_scen").Text
    Me.cmdScenario.Accelerator = GlobalX4Mess.Range("solv_dlg2_acc9").Text

    Call fnUpdateDialogFonts(Me)
End Sub

Private Sub UserForm_QueryClose(Cancel As Integer, CloseMode As Integer)
   Call cmdCancel_Click
End Sub


