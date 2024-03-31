VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} dlgConstraint 
   Caption         =   "Add Constraint"
   ClientHeight    =   1605
   ClientLeft      =   45
   ClientTop       =   345
   ClientWidth     =   5835
   OleObjectBlob   =   "dlgConstraint.frx":0000
   StartUpPosition =   1  'CenterOwner
   WhatsThisButton =   -1  'True
   WhatsThisHelp   =   -1  'True
End
Attribute VB_Name = "dlgConstraint"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False



Option Explicit

Private Sub AddConstraint()
    Dim newentry As String, i As Integer
    Dim temp As String, j As Integer
    With Me
    temp = ""
    On Error Resume Next
    If Range(Stylecheck(.refRHS.Text)).Worksheet.Name <> ActiveSheet.Name Then
       If Not IsNumeric(.refRHS.Text) And .comboRelation.ListIndex < 3 Then
          .refRHS.Text = Application.ConvertFormula(.refRHS.Text, Application.ReferenceStyle, Application.ReferenceStyle, xlAbsolute)
          temp = .refRHS.Text
         .refRHS.Text = EvaluateRHS(.refRHS.Text)
       End If
    End If
    On Error GoTo 0
    If .comboRelation.ListIndex < 3 Then
       newentry = .refLHS.Text & " " & .comboRelation.Text & " " & .refRHS.Text
    ElseIf .comboRelation.ListIndex = 3 Then
       newentry = .refLHS.Text & " = " & GlobalX4Mess.Range("solver_msg_int").Formula
    ElseIf .comboRelation.ListIndex = 4 Then
       newentry = .refLHS.Text & " = " & GlobalX4Mess.Range("solver_msg_bin").Formula
    Else
       newentry = .refLHS.Text & " = " & GlobalX4Mess.Range("solver_msg_dif").Formula
    End If
    End With
    If GlobalChange Then
       Application.RecordMacro GlobalVBHelp
       dlgSolverParameters.listConstraints.RemoveItem GlobalInd
    End If
  
    i = 0
    If dlgSolverParameters.listConstraints.ListCount > 0 Then
       Do While newentry >= dlgSolverParameters.listConstraints.List(i)
          i = i + 1
          If i = dlgSolverParameters.listConstraints.ListCount Then
             Exit Do
          End If
       Loop
    End If
    dlgSolverParameters.listConstraints.AddItem newentry, i
    
    If GlobalChange Then
       dlgSolverParameters.listConstraints.ListIndex = GlobalInd
    End If
    With Me
       If .comboRelation.ListIndex > 2 Then
          Application.RecordMacro basiccode:=GlobalX4Mess.Range("addfunc").Text & _
                " " & GlobalX4Mess.Range("addarg1").Text & ":=" & Chr(34) & _
                Range(Stylecheck(.refLHS.Text)).Address(, , GlobalR1C1) & Chr(34) & _
                ", " & GlobalX4Mess.Range("addarg2").Text & ":=" & .comboRelation.ListIndex + 1
       Else
          Application.RecordMacro basiccode:=GlobalX4Mess.Range("addfunc").Text & _
                " " & GlobalX4Mess.Range("addarg1").Text & ":=" & Chr(34) & _
                Range(Stylecheck(.refLHS.Text)).Address(, , GlobalR1C1) & Chr(34) & _
                ", " & GlobalX4Mess.Range("addarg2").Text & ":=" & .comboRelation.ListIndex + 1 & _
                ", " & GlobalX4Mess.Range("addarg3").Text & ":=" & Chr(34) & .refRHS.Text & Chr(34)
       End If
    End With
End Sub


Private Sub cmdAdd_Click()
    Me.refLHS.Text = GetName(Me.refLHS.Text)
    Me.refRHS.Text = GetName(Me.refRHS.Text)
    If Not Validconstraint Then
       Exit Sub
    End If
    AddConstraint
    With Me
        .refLHS.Text = ""
        .refRHS.Text = ""
        .comboRelation.ListIndex = 0
        .refLHS.SetFocus
    End With
End Sub

Private Sub cmdCancel_Click()
    Me.Hide
End Sub

Private Sub cmdOK_Click()
    Me.refLHS.Text = GetName(Me.refLHS.Text)
    Me.refRHS.Text = GetName(Me.refRHS.Text)
    If Not Validconstraint Then
       Exit Sub
    End If
    AddConstraint
    Me.Hide
End Sub

Private Sub comboRelation_Change()
    ' user changes middle of constraint
    With Me
       Select Case .comboRelation.ListIndex
          Case 0, 1, 2
            If .refRHS.Text = GlobalX4Mess.Range("solver_msg_int").Formula Or _
               .refRHS.Text = GlobalX4Mess.Range("solver_msg_bin").Formula Or _
               .refRHS.Text = GlobalX4Mess.Range("solver_msg_dif").Formula Then
               .refRHS.Text = ""
            End If
            .refRHS.SetFocus
          Case 3
            .refRHS.Text = GlobalX4Mess.Range("solver_msg_int").Formula
          Case 4
            .refRHS.Text = GlobalX4Mess.Range("solver_msg_bin").Formula
          Case 5
            .refRHS.Text = GlobalX4Mess.Range("solver_msg_dif").Formula
        End Select
    End With
End Sub

Private Sub UserForm_Activate()
    Dim place As Integer, the_index As Integer, constraint As String
    Dim stzFont As String
    Dim cctl As control
    
    If GlobalChange Then
        ' change a constraint
        With dlgSolverParameters.listConstraints
            If .ListIndex < 0 Then
                DisplayMessage "solver_msg_25a", 1830, 0
                Me.Hide
                Exit Sub
            End If
            constraint = .Text
        End With
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
        With Me
            .refLHS.Text = Trim(Left(constraint, place))
            .refRHS.Text = Trim(Right(constraint, Len(constraint) - place - 2))
            If .refRHS.Text = GlobalX4Mess.Range("solver_msg_int").Formula Then
                the_index = 4
                GlobalVBHelp = "SolverDelete CellRef:=range(" & Chr(34) & Stylecheck(.refLHS.Text) & Chr(34) & "), Relation:=" & the_index
            ElseIf .refRHS.Text = GlobalX4Mess.Range("solver_msg_bin").Formula Then
                the_index = 5
                GlobalVBHelp = "SolverDelete CellRef:=range(" & Chr(34) & Stylecheck(.refLHS.Text) & Chr(34) & "), Relation:=" & the_index
            ElseIf .refRHS.Text = GlobalX4Mess.Range("solver_msg_dif").Formula Then
                the_index = 6
                GlobalVBHelp = "SolverDelete CellRef:=range(" & Chr(34) & Stylecheck(.refLHS.Text) & Chr(34) & "), Relation:=" & the_index
            Else
                On Error Resume Next
                GlobalVBHelp = "SolverDelete CellRef:=range(" & Chr(34) & Stylecheck(.refLHS.Text) & Chr(34) & "), Relation:=" & the_index & ", FormulaText:=" & Chr(34) & .refRHS.Text & Chr(34)
                On Error GoTo 0
            End If
            .comboRelation.ListIndex = the_index - 1
            .Caption = GlobalX4Mess.Range("solver_msg_44b").Text
        End With
    Else
        With Me
            .refLHS.Text = ""
            .refRHS.Text = ""
            .comboRelation.ListIndex = 0
            .Caption = GlobalX4Mess.Range("solver_msg_44a").Text
        End With
    End If
    
    Call fnUpdateDialogFonts(Me)

    Me.refLHS.SetFocus
End Sub

Private Sub UserForm_Initialize()
    Me.refLHS.Text = ""
    Me.refRHS.Text = ""
    Me.comboRelation.List = Array("<=", "=", ">=", GlobalX4Mess.Range("solver_int").Text, _
            GlobalX4Mess.Range("solver_bin").Text, GlobalX4Mess.Range("solver_dif").Text)
    Me.lblLHS.Caption = GlobalX4Mess.Range("solv_dlg1_lhs").Text
    Me.lblLHS.Accelerator = GlobalX4Mess.Range("solv_dlg1_acc1").Text
    Me.lblRHS.Caption = GlobalX4Mess.Range("solv_dlg1_rhs").Text
    Me.lblRHS.Accelerator = GlobalX4Mess.Range("solv_dlg1_acc2").Text
    Me.cmdOK.Caption = GlobalX4Mess.Range("solv_dlg1_ok").Text
    Me.cmdOK.Accelerator = GlobalX4Mess.Range("solv_dlg1_acc3").Text
    Me.cmdAdd.Caption = GlobalX4Mess.Range("solv_dlg1_add").Text
    Me.cmdAdd.Accelerator = GlobalX4Mess.Range("solv_dlg1_acc4").Text
    Me.cmdCancel.Caption = GlobalX4Mess.Range("solv_dlg1_cancel").Text
    Me.cmdCancel.Accelerator = GlobalX4Mess.Range("solv_dlg1_acc5").Text
End Sub

Private Function Validconstraint()
    ' validates constraint
    ' check lhs
    Dim dummy As Integer, tmp As String
    tmp = dlgSolverParameters.refVariables.Text
    With Me
         On Error Resume Next
         If Range(Stylecheck(.refLHS.Text)).Worksheet.Name <> GlobalSheetName Or _
                Range(Stylecheck(.refLHS.Text)).Areas.Count > 1 Or .refLHS.Text = "" Then
            DisplayMessage "solver_msg_20", 1830, 0
            .refLHS.SetFocus
            Validconstraint = False
            On Error GoTo 0
            Exit Function
         End If
         ' check rhs
         If .comboRelation.ListIndex = 3 Then
            If .refRHS.Text = GlobalX4Mess.Range("solver_msg_int").Formula Or _
               .refRHS.Text = "=" & GlobalX4Mess.Range("solver_msg_int").Formula Then
               ' check if lhs is ok
               If tmp <> "" Then
                  If Range("solver_adj").Cells.Count <> Union(Range("solver_adj"), Range(Stylecheck(.refLHS.Text))).Cells.Count Then
                     DisplayMessage "solver_msg_23", 1830, 0
                     .refLHS.SetFocus
                     Validconstraint = False
                     On Error GoTo 0
                     Exit Function
                  End If
               Else
                  DisplayMessage "solver_msg_23", 1830, 0
                  .refLHS.SetFocus
                  Validconstraint = False
                  On Error GoTo 0
                  Exit Function
               End If
            Else
               DisplayMessage "solver_msg_21", 1830, 0
              .refRHS.SetFocus
               Validconstraint = False
               On Error GoTo 0
               Exit Function
            End If
            Validconstraint = True
            On Error GoTo 0
            Exit Function
         End If
         If .comboRelation.ListIndex = 4 Then
            If .refRHS.Text = GlobalX4Mess.Range("solver_msg_bin").Formula Or _
               .refRHS.Text = "=" & GlobalX4Mess.Range("solver_msg_bin").Formula Then
               ' check if lhs is ok
               If tmp <> "" Then
                  If Range("solver_adj").Cells.Count <> Union(Range("solver_adj"), Range(Stylecheck(.refLHS.Text))).Cells.Count Then
                     DisplayMessage "solver_msg_24", 1830, 0
                     .refLHS.SetFocus
                     Validconstraint = False
                     On Error GoTo 0
                     Exit Function
                  End If
               Else
                  DisplayMessage "solver_msg_24", 1830, 0
                  .refLHS.SetFocus
                  Validconstraint = False
                  On Error GoTo 0
                  Exit Function
               End If
            Else
               DisplayMessage "solver_msg_21", 1830, 0
               .refLHS.SetFocus
               Validconstraint = False
               On Error GoTo 0
               Exit Function
            End If
            Validconstraint = True
            On Error GoTo 0
            Exit Function
         End If
         If .comboRelation.ListIndex = 5 Then
            If .refRHS.Text = GlobalX4Mess.Range("solver_msg_dif").Formula Or _
               .refRHS.Text = "=" & GlobalX4Mess.Range("solver_msg_dif").Formula Then
               ' check if lhs is ok
               If tmp <> "" Then
                  If Range("solver_adj").Cells.Count <> Union(Range("solver_adj"), Range(Stylecheck(.refLHS.Text))).Cells.Count Then
                     DisplayMessage "solver_msg_23a", 1830, 0
                     .refLHS.SetFocus
                     Validconstraint = False
                     On Error GoTo 0
                     Exit Function
                  End If
               Else
                  DisplayMessage "solver_msg_23a", 1830, 0
                  .refLHS.SetFocus
                  Validconstraint = False
                  On Error GoTo 0
                  Exit Function
               End If
            Else
               DisplayMessage "solver_msg_21", 1830, 0
               .refLHS.SetFocus
               Validconstraint = False
               On Error GoTo 0
               Exit Function
            End If
            Validconstraint = True
            On Error GoTo 0
            Exit Function
         End If
         ' "normal constraint"
         If Not IsNumeric(.refRHS.Text) Then
            ' is it a range?
            Err = 0
            On Error Resume Next
            dummy = Range(Stylecheck(.refRHS.Text)).Areas.Count
            If Err = 0 Then
               'it's a range!
               If dummy > 1 Then ' Or Range(.refRHS.Text).Worksheet.Name <> GlobalSheetname Then
                  DisplayMessage "solver_msg_21", 1830, 0
                  .refRHS.SetFocus
                  Validconstraint = False
                  On Error GoTo 0
                  Exit Function
               End If
               If Range(Stylecheck(.refRHS.Text)).Cells.Count > 1 And _
                    Range(Stylecheck(.refRHS.Text)).Cells.Count <> Range(Stylecheck(.refLHS.Text)).Cells.Count Then
                  DisplayMessage "solver_msg_22", 1830, 0
                  .refRHS.SetFocus
                  Validconstraint = False
                  On Error GoTo 0
                  Exit Function
               End If
            Else
              ' must be formula or bogus entry
               If Not IsNumeric(Application.Evaluate(.refRHS.Text)) Then
                  ' bogus!
                  DisplayMessage "solver_msg_21", 1830, 0
                  .refRHS.SetFocus
                  Validconstraint = False
                  On Error GoTo 0
                  Exit Function
               End If
            End If
            On Error GoTo 0
         End If
     End With
     Validconstraint = True
End Function




