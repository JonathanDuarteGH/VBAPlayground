VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} dlgSolverParameters 
   Caption         =   "Solver Parameters"
   ClientHeight    =   8340.001
   ClientLeft      =   45
   ClientTop       =   345
   ClientWidth     =   7470
   OleObjectBlob   =   "dlgSolverParameters.frx":0000
   StartUpPosition =   1  'CenterOwner
   WhatsThisButton =   -1  'True
   WhatsThisHelp   =   -1  'True
End
Attribute VB_Name = "dlgSolverParameters"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Option Explicit
' Windows API calls used to make the frmControl form resizable.
'The offset of a window's style
Private Const GWL_STYLE As Long = (-16)

'Style to add a sizable frame
Private Const WS_THICKFRAME As Long = &H40000

#If VBA7 Then

'Find the form's window handle
Private Declare PtrSafe Function FindWindow Lib "user32" _
    Alias "FindWindowA" ( _
    ByVal lpClassName As String, _
    ByVal lpWindowName As String) As LongPtr

'Get the form's window style
Private Declare PtrSafe Function GetWindowLong Lib "user32" _
    Alias "GetWindowLongA" ( _
    ByVal hwnd As LongPtr, _
    ByVal nIndex As Long) As Long

'Set the form's window style
Private Declare PtrSafe Function SetWindowLong Lib "user32" _
    Alias "SetWindowLongA" ( _
    ByVal hwnd As LongPtr, _
    ByVal nIndex As Long, _
    ByVal dwNewLong As Long) As Long

Private Declare PtrSafe Function SetWindowPos Lib "user32" _
   (ByVal hwnd As LongPtr, ByVal hWndInsAfter As LongPtr, ByVal x As Long, _
    ByVal y As Long, ByVal cx As Long, ByVal cy As Long, _
    ByVal wFlags As Long) As Long

#Else

'Find the form's window handle
Private Declare Function FindWindow Lib "user32" _
    Alias "FindWindowA" ( _
    ByVal lpClassName As String, _
    ByVal lpWindowName As String) As Long

'Get the form's window style
Private Declare Function GetWindowLong Lib "user32" _
    Alias "GetWindowLongA" ( _
    ByVal hwnd As Long, _
    ByVal nIndex As Long) As Long

'Set the form's window style
Private Declare Function SetWindowLong Lib "user32" _
    Alias "SetWindowLongA" ( _
    ByVal hwnd As Long, _
    ByVal nIndex As Long, _
    ByVal dwNewLong As Long) As Long

Private Declare Function SetWindowPos Lib "user32" _
   (ByVal hwnd As Long, ByVal hWndInsAfter As Long, ByVal x As Long, _
    ByVal y As Long, ByVal cx As Long, ByVal cy As Long, _
    ByVal wFlags As Long) As Long

#End If

Private iHeight As Integer
Private iWidth As Integer

Private Sub cmdAdd_Click()
    'validation of objective
    dlgSolverParameters.refObj.Text = GetName(dlgSolverParameters.refObj.Text)
    dlgSolverParameters.refVariables.Text = GetName(dlgSolverParameters.refVariables.Text)
    If dlgSolverParameters.refObj.Text <> "" Then
       On Error Resume Next
       If Range(Stylecheck(dlgSolverParameters.refObj.Text)).Count <> 1 Or Range(Stylecheck(dlgSolverParameters.refObj.Text)).Worksheet.Name <> ActiveSheet.Name Then
          DisplayMessage "solver_msg_7", 1830, 0
          dlgSolverParameters.refObj.SetFocus
          On Error GoTo 0
          Exit Sub
       End If
    End If
    DefineObjVars
    Me.Hide
    GlobalChange = False
    dlgConstraint.Show
    Me.Show
End Sub

Private Sub cmdChange_Click()
    'validation of objective
    dlgSolverParameters.refObj.Text = GetName(dlgSolverParameters.refObj.Text)
    dlgSolverParameters.refVariables.Text = GetName(dlgSolverParameters.refVariables.Text)
    If dlgSolverParameters.refObj.Text <> "" Then
       On Error Resume Next
       If Range(Stylecheck(dlgSolverParameters.refObj.Text)).Count <> 1 Or Range(Stylecheck(dlgSolverParameters.refObj.Text)).Worksheet.Name <> ActiveSheet.Name Then
          DisplayMessage "solver_msg_7", 1830, 0
          dlgSolverParameters.refObj.SetFocus
          On Error GoTo 0
          Exit Sub
       End If
    End If
    GlobalInd = Me.listConstraints.ListIndex
    If GlobalInd < 0 Then
       DisplayMessage "solver_msg_25a", 1830, 0
       Exit Sub
    End If
    DefineObjVars
    Me.Hide
    GlobalChange = True
    dlgConstraint.Show
    Me.Show
End Sub

Private Sub cmdClose_Click()
    'validation of objective
    dlgSolverParameters.refObj.Text = GetName(dlgSolverParameters.refObj.Text)
    dlgSolverParameters.refVariables.Text = GetName(dlgSolverParameters.refVariables.Text)
    If dlgSolverParameters.refObj.Text <> "" Then
       On Error Resume Next
       If Range(Stylecheck(dlgSolverParameters.refObj.Text)).Count <> 1 Or Range(Stylecheck(dlgSolverParameters.refObj.Text)).Worksheet.Name <> ActiveSheet.Name Then
          DisplayMessage "solver_msg_7", 1830, 0
          dlgSolverParameters.refObj.SetFocus
          On Error GoTo 0
          Exit Sub
       End If
    End If
    Me.cmdOptions.SetFocus
    DefineModel
    'SaveOptions
    Me.Hide
End Sub

Private Sub cmdDelete_Click()
    Dim place As Integer, the_index As Integer
    Dim theleft As String, theright As String
    Dim constraint As String
    
     ' delete a constraint
    With Me.listConstraints
        If .ListIndex < 0 Then
           DisplayMessage "solver_msg_25b", 1830, 0
           Exit Sub
        End If
        constraint = .Text
        If InStr(constraint, " <= ") > 0 Then
           place = InStr(constraint, " <= ")
           the_index = 1
        ElseIf InStr(constraint, " >= ") > 0 Then
           place = InStr(constraint, " >= ")
           the_index = 3
        Else
           place = InStr(constraint, " = ")
          the_index = 2
        End If
        theleft = Trim(Left(constraint, place))
        theright = Trim(Right(constraint, Len(constraint) - place - 2))
        If theright = GlobalX4Mess.Range("solver_msg_int").Formula Then
          the_index = 4
          GlobalVBHelp = GlobalX4Mess.Range("Delfunc").Text & " " & GlobalX4Mess.Range("addarg1").Text & _
                    ":=range(" & Chr(34) & Stylecheck(theleft) & Chr(34) & "), " & _
                    GlobalX4Mess.Range("addarg2").Text & ":=" & the_index
        ElseIf theright = GlobalX4Mess.Range("solver_msg_bin").Formula Then
          the_index = 5
          GlobalVBHelp = GlobalX4Mess.Range("Delfunc").Text & " " & GlobalX4Mess.Range("addarg1").Text & _
                    ":=range(" & Chr(34) & theleft & Chr(34) & "), " & _
                    GlobalX4Mess.Range("addarg2").Text & ":=" & the_index
        ElseIf theright = GlobalX4Mess.Range("solver_msg_dif").Formula Then
          the_index = 6
          GlobalVBHelp = GlobalX4Mess.Range("Delfunc").Text & " " & GlobalX4Mess.Range("addarg1").Text & _
                    ":=range(" & Chr(34) & theleft & Chr(34) & "), " & _
                    GlobalX4Mess.Range("addarg2").Text & ":=" & the_index
        Else
          On Error Resume Next
          GlobalVBHelp = GlobalX4Mess.Range("Delfunc").Text & " " & GlobalX4Mess.Range("addarg1").Text & _
                    ":=range(" & Chr(34) & theleft & Chr(34) & "), " & _
                    GlobalX4Mess.Range("addarg2").Text & ":=" & the_index & ", " & _
                    GlobalX4Mess.Range("addarg3").Text & ":=" & Chr(34) & theright & Chr(34)
          On Error GoTo 0
        End If
        Application.RecordMacro GlobalVBHelp
        .RemoveItem .ListIndex
    End With
End Sub

Private Sub cmdLoadSave_Click()
'validation of objective
    dlgSolverParameters.refObj.Text = GetName(dlgSolverParameters.refObj.Text)
    dlgSolverParameters.refVariables.Text = GetName(dlgSolverParameters.refVariables.Text)
    If dlgSolverParameters.refObj.Text <> "" Then
       On Error Resume Next
       If Range(Stylecheck(dlgSolverParameters.refObj.Text)).Count <> 1 Or Range(Stylecheck(dlgSolverParameters.refObj.Text)).Worksheet.Name <> ActiveSheet.Name Then
          DisplayMessage "solver_msg_7", 1830, 0
          dlgSolverParameters.refObj.SetFocus
          On Error GoTo 0
          Exit Sub
       End If
    End If
    Me.Hide
    dlgLoadSave.Show
    Me.Show
End Sub

Private Sub cmdOptions_Click()
    'validation of objective
    dlgSolverParameters.refObj.Text = GetName(dlgSolverParameters.refObj.Text)
    dlgSolverParameters.refVariables.Text = GetName(dlgSolverParameters.refVariables.Text)
    If dlgSolverParameters.refObj.Text <> "" Then
       On Error Resume Next
       If Range(Stylecheck(dlgSolverParameters.refObj.Text)).Count <> 1 Or Range(Stylecheck(dlgSolverParameters.refObj.Text)).Worksheet.Name <> ActiveSheet.Name Then
          DisplayMessage "solver_msg_7", 1830, 0
          dlgSolverParameters.refObj.SetFocus
          On Error GoTo 0
          Exit Sub
       End If
    End If
    Me.Hide
    dlgOptions.Show
    Me.Show
End Sub

Private Sub cmdReset_Click()
    Reset_all (True)
End Sub

Private Sub cmdSolve_Click()
    Dim savestatusbar As Boolean
    Dim savecalc As Integer, i As Integer, c
    
    'validation of objective
    dlgSolverParameters.refObj.Text = GetName(dlgSolverParameters.refObj.Text)
    dlgSolverParameters.refVariables.Text = GetName(dlgSolverParameters.refVariables.Text)
    If dlgSolverParameters.refObj.Text <> "" Then
       On Error Resume Next
       If Range(Stylecheck(dlgSolverParameters.refObj.Text)).Count <> 1 Or Range(Stylecheck(dlgSolverParameters.refObj.Text)).Worksheet.Name <> ActiveSheet.Name Then
          DisplayMessage "solver_msg_7", 1830, 0
          dlgSolverParameters.refObj.SetFocus
          On Error GoTo 0
          Exit Sub
       End If
    End If
    Me.Hide
    DefineModel
    SaveOptions
    If SolverCls Is Nothing Then
        Set SolverCls = New SolverCalls
    End If
    SolverCls.UDF = ""
    With Me
        If Problem_ok() Then
            Application.ScreenUpdating = False
            savestatusbar = Application.DisplayStatusBar
            Application.DisplayStatusBar = True
            savecalc = Application.Calculation
            Application.Calculation = xlCalculationManual
           
            On Error Resume Next
            If .refObj.Text = "" Then
                GlobalOldObj = CVErr(xlErrNA)
            Else
                GlobalOldObj = Range(ActiveSheet.Names("solver_opt").Name).Value
                GlobalOldObjFormat = Range(ActiveSheet.Names("solver_opt").Name).NumberFormat
            End If
            ReDim GlobalOldVars(Range("solver_adj").Cells.Count)
            ReDim GlobalOldVarFormats(Range("solver_adj").Cells.Count)
            i = 1
            For Each c In Range("solver_adj").Cells
                GlobalOldVars(i) = c.Value
                GlobalOldVarFormats(i) = c.NumberFormat
                i = i + 1
            Next
            
            GlobalAnswer = SolverCls.Solve(0)
            Application.RecordMacro basiccode:=GlobalX4Mess.Range("solvefunc").Text
            Application.DisplayStatusBar = savestatusbar
            Application.StatusBar = False
            Application.ScreenUpdating = True
            If Application.Calculation <> savecalc Then
                Application.Calculation = savecalc
            End If
            dlgFinish.Show
        Else
           Me.Show
        End If
    End With
End Sub

Private Sub radioMax_Change()
    If Me.radioMin Or Me.radioMax Then
       Me.editValueOf.Enabled = False
    Else
        Me.editValueOf.Enabled = True
    End If
End Sub

Private Sub radioMin_Click()
    If Me.radioMin Or Me.radioMax Then
       Me.editValueOf.Enabled = False
    Else
        Me.editValueOf.Enabled = True
    End If
End Sub

Private Sub radioValueOf_Click()
    If Me.radioMin Or Me.radioMax Then
       Me.editValueOf.Enabled = False
    Else
        Me.editValueOf.Enabled = True
        Me.editValueOf.SetFocus
    End If
End Sub

Private Sub UserForm_Activate()
    Dim cctl As control

    If globalHooked = False Then
        Hook
        globalHooked = True
    End If
    
    Call fnUpdateDialogFonts(Me)
    
    Me.refObj.SetFocus
End Sub

Private Sub UserForm_Initialize()
    On Error Resume Next
    Me.editValueOf.Text = ""
    Me.listConstraints.Clear
    Me.refObj.Text = ""
    Me.refVariables.Text = ""
    On Error GoTo 0


#If VBA7 Then
    Dim mhWndForm As LongPtr
#Else
    Dim mhWndForm As Long
#End If

    ' Get the handle for the form (for Excel 2000 or later).
    mhWndForm = FindWindow("ThunderDFrame", Me.Caption)
    
    'Save handle to the form.
    gHW = mhWndForm

    'Begin subclassing.
    'Hook
    
    Me.comboEngines.List = _
           Array(GlobalX4Mess.Range("solver_grg_eng").Text, _
           GlobalX4Mess.Range("solver_lp_eng").Text, _
           GlobalX4Mess.Range("solver_crs_eng").Text)
    Me.comboEngines.ListIndex = 0
   
    Dim iStyle As Long

    Me.lblHelpTitle.Caption = GlobalX4Mess.Range("solver_hlp_main5").Text
    Me.lblHelp.Caption = GlobalX4Mess.Range("solver_hlp_main5a").Text
    
    Me.Caption = GlobalX4Mess.Range("solv_dlg8_title").Text
    Me.lblObjective.Caption = GlobalX4Mess.Range("solv_dlg8_obj").Text
    Me.lblObjective.Accelerator = GlobalX4Mess.Range("solv_dlg8_acc1").Text
    Me.lblTo.Caption = GlobalX4Mess.Range("solv_dlg8_to").Text
    Me.radioMax.Caption = GlobalX4Mess.Range("solv_dlg8_max").Text
    Me.radioMax.Accelerator = GlobalX4Mess.Range("solv_dlg8_acc2").Text
    Me.radioMin.Caption = GlobalX4Mess.Range("solv_dlg8_min").Text
    Me.radioMin.Accelerator = GlobalX4Mess.Range("solv_dlg8_acc3").Text
    Me.radioValueOf.Caption = GlobalX4Mess.Range("solv_dlg8_val").Text
    Me.radioValueOf.Accelerator = GlobalX4Mess.Range("solv_dlg8_acc4").Text
    Me.lblVariables.Caption = GlobalX4Mess.Range("solv_dlg8_vars").Text
    Me.lblVariables.Accelerator = GlobalX4Mess.Range("solv_dlg8_acc5").Text
    Me.lblConstraints.Caption = GlobalX4Mess.Range("solv_dlg8_cons").Text
    Me.lblConstraints.Accelerator = GlobalX4Mess.Range("solv_dlg8_acc6").Text
    Me.chkAssumeNonNeg.Caption = GlobalX4Mess.Range("solv_dlg8_nonneg").Text
    Me.chkAssumeNonNeg.Accelerator = GlobalX4Mess.Range("solv_dlg8_acc7").Text
    Me.lblMethod.Caption = GlobalX4Mess.Range("solv_dlg8_method").Text
    Me.lblMethod.Accelerator = GlobalX4Mess.Range("solv_dlg8_acc8").Text
    Me.cmdAdd.Caption = GlobalX4Mess.Range("solv_dlg8_add").Text
    Me.cmdAdd.Accelerator = GlobalX4Mess.Range("solv_dlg8_acc9").Text
    Me.cmdChange.Caption = GlobalX4Mess.Range("solv_dlg8_change").Text
    Me.cmdChange.Accelerator = GlobalX4Mess.Range("solv_dlg8_acc10").Text
    Me.cmdDelete.Caption = GlobalX4Mess.Range("solv_dlg8_delete").Text
    Me.cmdDelete.Accelerator = GlobalX4Mess.Range("solv_dlg8_acc11").Text
    Me.cmdReset.Caption = GlobalX4Mess.Range("solv_dlg8_reset").Text
    Me.cmdReset.Accelerator = GlobalX4Mess.Range("solv_dlg8_acc12").Text
    Me.cmdLoadSave.Caption = GlobalX4Mess.Range("solv_dlg8_load").Text
    Me.cmdLoadSave.Accelerator = GlobalX4Mess.Range("solv_dlg8_acc13").Text
    Me.cmdOptions.Caption = GlobalX4Mess.Range("solv_dlg8_options").Text
    Me.cmdOptions.Accelerator = GlobalX4Mess.Range("solv_dlg8_acc14").Text
    Me.cmdSolve.Caption = GlobalX4Mess.Range("solv_dlg8_solve").Text
    Me.cmdSolve.Accelerator = GlobalX4Mess.Range("solv_dlg8_acc15").Text
    Me.cmdClose.Caption = GlobalX4Mess.Range("solv_dlg8_close").Text
    Me.cmdClose.Accelerator = GlobalX4Mess.Range("solv_dlg8_acc16").Text
    
    'Make the form resizable
    iStyle = GetWindowLong(mhWndForm, GWL_STYLE)
    iStyle = iStyle Or WS_THICKFRAME
    SetWindowLong mhWndForm, GWL_STYLE, iStyle
    SetWindowPos mhWndForm, 0, 0, 0, 0, 0, 39
    iHeight = Me.Height
    iWidth = Me.Width
End Sub

Private Sub UserForm_QueryClose(Cancel As Integer, CloseMode As Integer)
    Unhook
End Sub

Private Sub UserForm_Resize()
    Dim iHeightIncrease As Integer
    Dim iWidthIncrease As Integer
    If iHeight = 0 Or iWidth = 0 Then Exit Sub

    iHeightIncrease = (Me.Height - iHeight)
    If iHeightIncrease < 0 Then iHeightIncrease = 0
    iWidthIncrease = (Me.Width - iWidth)
    If iWidthIncrease < 0 Then iWidthIncrease = 0
    Me.MainFrame.Width = 354 + iWidthIncrease
    Me.MainFrame.Height = 372 + iHeightIncrease
    Me.cmdAdd.Left = 246 + iWidthIncrease
    Me.cmdChange.Left = 246 + iWidthIncrease
    Me.cmdDelete.Left = 246 + iWidthIncrease
    Me.cmdOptions.Left = 246 + iWidthIncrease
    Me.cmdReset.Left = 246 + iWidthIncrease
    Me.cmdLoadSave.Left = 246 + iWidthIncrease
    Me.cmdSolve.Left = 150 + iWidthIncrease
    Me.cmdClose.Left = 252 + iWidthIncrease
    Me.cmdSolve.Top = 384 + iHeightIncrease
    Me.cmdClose.Top = 384 + iHeightIncrease
    
    Me.refObj.Width = 240 + iWidthIncrease
    Me.editValueOf.Width = 156 + iWidthIncrease
    Me.refVariables.Width = 330 + iWidthIncrease
    Me.listConstraints.Width = 222 + iWidthIncrease

    Me.cmdReset.Top = 192 + iHeightIncrease
    Me.cmdLoadSave.Top = 215 + iHeightIncrease
    Me.cmdOptions.Top = 270 + iHeightIncrease
    Me.listConstraints.Height = 126.8 + iHeightIncrease
    Me.chkAssumeNonNeg.Top = 248 + iHeightIncrease
    Me.lblMethod.Top = 276 + iHeightIncrease
    Me.comboEngines.Top = 274 + iHeightIncrease
    Me.frameHelp.Top = 296 + iHeightIncrease
    Me.frameHelp.Width = 330 + iWidthIncrease
    Me.lblHelp.Width = 318 + iWidthIncrease
    Me.lblHelpTitle.Width = 318 + iWidthIncrease
  
End Sub



