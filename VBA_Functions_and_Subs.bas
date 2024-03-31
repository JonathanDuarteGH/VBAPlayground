Attribute VB_Name = "VBA_Functions_and_Subs"
Option Explicit

' ANALYSIS TOOLPAK  -  Excel AddIn
' The following function declarations provide interface between VBA and ATP XLL.

Private Const c_sAddinFolder As String = "Analysis"
Private Const c_sXllName As String = "ANALYS32.XLL" 'Must be UPPER-CASE for auto_close
Private s_sXllFullName As String

Dim FunctionIDs(37, 0 To 1)

Private Function GetMacroRegId(FuncText As String) As String
    Dim i As Variant
    For i = LBound(FunctionIDs) To UBound(FunctionIDs)
        If (LCase(FunctionIDs(i, 0)) = LCase(FuncText)) Then
            If (Not (IsError(FunctionIDs(i, 1)))) Then
                GetMacroRegId = FunctionIDs(i, 1)
                Exit Function
            End If
        End If
    Next i
End Function

'Procedures
 
Sub Anova1(inprng As Variant, Optional outrng As Variant, Optional grouped As Variant, Optional labels As Variant, Optional alpha As Variant)
Attribute Anova1.VB_Description = "Performs single-factor analysis of variance"
Attribute Anova1.VB_HelpID = 3457
Attribute Anova1.VB_ProcData.VB_Invoke_Func = " \n3"
    Dim xAnova1 As Variant
    xAnova1 = Application.Run(GetMacroRegId("fnAnova1"), inprng, outrng, grouped, labels, alpha)
End Sub

Sub Anova1Q(Optional inprng As Variant, Optional outrng As Variant, Optional grouped As Variant, Optional labels As Variant, Optional alpha As Variant)
Attribute Anova1Q.VB_Description = "Performs single-factor analysis of variance"
Attribute Anova1Q.VB_HelpID = 3458
Attribute Anova1Q.VB_ProcData.VB_Invoke_Func = " \n3"
    Dim xAnova1Q As Variant
    xAnova1Q = Application.Run(GetMacroRegId("fnAnova1Q"), inprng, outrng, grouped, labels, alpha)
End Sub

Sub Anova2(inprng As Variant, Optional outrng As Variant, Optional sample_rows As Variant, Optional alpha As Variant)
Attribute Anova2.VB_Description = "Performs two-factor analysis of variance with replication"
Attribute Anova2.VB_HelpID = 3459
Attribute Anova2.VB_ProcData.VB_Invoke_Func = " \n3"
    Dim xAnova2 As Variant
    xAnova2 = Application.Run(GetMacroRegId("fnAnova2"), inprng, outrng, sample_rows, alpha)
End Sub

Sub Anova2Q(Optional inprng As Variant, Optional outrng As Variant, Optional sample_rows As Variant, Optional alpha As Variant)
Attribute Anova2Q.VB_Description = "Performs two-factor analysis of variance with replication"
Attribute Anova2Q.VB_HelpID = 3460
Attribute Anova2Q.VB_ProcData.VB_Invoke_Func = " \n3"
    Dim xAnova2Q As Variant
    xAnova2Q = Application.Run(GetMacroRegId("fnAnova2Q"), inprng, outrng, sample_rows, alpha)
End Sub

Sub Anova3(inprng As Variant, Optional outrng As Variant, Optional labels As Variant, Optional alpha As Variant)
Attribute Anova3.VB_Description = "Performs two-factor analysis of variance without replication"
Attribute Anova3.VB_HelpID = 3461
Attribute Anova3.VB_ProcData.VB_Invoke_Func = " \n3"
    Dim xAnova3 As Variant
    xAnova3 = Application.Run(GetMacroRegId("fnAnova3"), inprng, outrng, labels, alpha)
End Sub

Sub Anova3Q(Optional inprng As Variant, Optional outrng As Variant, Optional labels As Variant, Optional alpha As Variant)
Attribute Anova3Q.VB_Description = "Performs two-factor analysis of variance without replication"
Attribute Anova3Q.VB_HelpID = 3462
Attribute Anova3Q.VB_ProcData.VB_Invoke_Func = " \n3"
    Dim xAnova3Q As Variant
    xAnova3Q = Application.Run(GetMacroRegId("fnAnova3Q"), inprng, outrng, labels, alpha)
End Sub

Sub Descr(inprng As Variant, Optional outrng As Variant, Optional grouped As Variant, Optional labels As Variant, Optional summary As Variant, Optional ds_large As Variant, Optional ds_small As Variant, Optional confid As Variant)
Attribute Descr.VB_Description = "Generates descriptive statistics for data in the input range"
Attribute Descr.VB_HelpID = 3463
Attribute Descr.VB_ProcData.VB_Invoke_Func = " \n3"
    Dim xDescr As Variant
    xDescr = Application.Run(GetMacroRegId("fnDescr"), inprng, outrng, grouped, labels, summary, ds_large, ds_small, confid)
End Sub

Sub DescrQ(Optional inprng As Variant, Optional outrng As Variant, Optional grouped As Variant, Optional labels As Variant, Optional summary As Variant, Optional ds_large As Variant, Optional ds_small As Variant, Optional confid As Variant)
Attribute DescrQ.VB_Description = "Generates descriptive statistics for data in the input range"
Attribute DescrQ.VB_HelpID = 3464
Attribute DescrQ.VB_ProcData.VB_Invoke_Func = " \n3"
    Dim xDescrQ As Variant
    xDescrQ = Application.Run(GetMacroRegId("fnDescrQ"), inprng, outrng, grouped, labels, summary, ds_large, ds_small, confid)
End Sub

Sub Expon(inprng As Variant, Optional outrng As Variant, Optional damp As Variant, Optional stderrs As Variant, Optional chart As Variant, Optional labels As Variant)
Attribute Expon.VB_Description = "Predicts a value based on the forecast for the prior period, adjusted for the error in that prior period"
Attribute Expon.VB_HelpID = 3465
Attribute Expon.VB_ProcData.VB_Invoke_Func = " \n3"
    Dim xExpon As Variant
    xExpon = Application.Run(GetMacroRegId("fnExpon"), inprng, outrng, damp, stderrs, chart, labels)
End Sub

Sub ExponQ(Optional inprng As Variant, Optional outrng As Variant, Optional damp As Variant, Optional stderrs As Variant, Optional chart As Variant, Optional labels As Variant)
Attribute ExponQ.VB_Description = "Predicts a value based on the forecast for the prior period, adjusted for the error in that prior period"
Attribute ExponQ.VB_HelpID = 3466
Attribute ExponQ.VB_ProcData.VB_Invoke_Func = " \n3"
    Dim xExponQ As Variant
    xExponQ = Application.Run(GetMacroRegId("fnExponQ"), inprng, outrng, damp, stderrs, chart, labels)
End Sub

Sub Fourier(inprng As Variant, Optional outrng As Variant, Optional inverse As Variant, Optional labels As Variant)
Attribute Fourier.VB_Description = "Performs a Fast Fourier Transform"
Attribute Fourier.VB_HelpID = 3467
Attribute Fourier.VB_ProcData.VB_Invoke_Func = " \n15"
    Dim xFourier As Variant
    xFourier = Application.Run(GetMacroRegId("fnFourier"), inprng, outrng, inverse, labels)
End Sub

Sub FourierQ(Optional inprng As Variant, Optional outrng As Variant, Optional inverse As Variant, Optional labels As Variant)
Attribute FourierQ.VB_Description = "Performs a Fast Fourier Transform"
Attribute FourierQ.VB_HelpID = 3468
Attribute FourierQ.VB_ProcData.VB_Invoke_Func = " \n15"
    Dim xFourierQ As Variant
    xFourierQ = Application.Run(GetMacroRegId("fnFourierQ"), inprng, outrng, inverse, labels)
End Sub

Sub Ftestv(inprng1 As Variant, inprng2 As Variant, Optional outrng As Variant, Optional labels As Variant, Optional alpha As Variant)
Attribute Ftestv.VB_Description = "Performs a two-sample F-test"
Attribute Ftestv.VB_HelpID = 3469
Attribute Ftestv.VB_ProcData.VB_Invoke_Func = " \n3"
    Dim xFtestv As Variant
    xFtestv = Application.Run(GetMacroRegId("fnFtestV"), inprng1, inprng2, outrng, labels, alpha)
End Sub

Sub FtestvQ(Optional inprng1 As Variant, Optional inprng2 As Variant, Optional outrng As Variant, Optional labels As Variant, Optional alpha As Variant)
Attribute FtestvQ.VB_Description = "Performs a two-sample F-test"
Attribute FtestvQ.VB_HelpID = 3470
Attribute FtestvQ.VB_ProcData.VB_Invoke_Func = " \n3"
    Dim xFtestvQ As Variant
    xFtestvQ = Application.Run(GetMacroRegId("fnFtestVQ"), inprng1, inprng2, outrng, labels, alpha)
End Sub

Sub Histogram(inprng As Variant, Optional outrng As Variant, Optional binrng As Variant, Optional pareto As Variant, Optional chartc As Variant, Optional chart As Variant, Optional labels As Variant)
Attribute Histogram.VB_Description = "Calculates individual and cumulative percentages for a range of data and a corresponding range of data bins"
Attribute Histogram.VB_HelpID = 3471
Attribute Histogram.VB_ProcData.VB_Invoke_Func = " \n3"
    Dim xHistogram As Variant
    xHistogram = Application.Run(GetMacroRegId("fnHistogram"), inprng, outrng, binrng, pareto, chartc, chart, labels)
End Sub

Sub HistogramQ(Optional inprng As Variant, Optional outrng As Variant, Optional binrng As Variant, Optional pareto As Variant, Optional chartc As Variant, Optional chart As Variant, Optional labels As Variant)
Attribute HistogramQ.VB_Description = "Calculates individual and cumulative percentages for a range of data and a corresponding range of data bins"
Attribute HistogramQ.VB_HelpID = 3472
Attribute HistogramQ.VB_ProcData.VB_Invoke_Func = " \n3"
    Dim xHistogramQ As Variant
    xHistogramQ = Application.Run(GetMacroRegId("fnHistogramQ"), inprng, outrng, binrng, pareto, chartc, chart, labels)
End Sub

Sub Mcorrel(inprng As Variant, Optional outrng As Variant, Optional grouped As Variant, Optional labels As Variant)
Attribute Mcorrel.VB_Description = "Returns a correlation matrix that measures the correlation between two or more data sets"
Attribute Mcorrel.VB_HelpID = 3473
Attribute Mcorrel.VB_ProcData.VB_Invoke_Func = " \n3"
    Dim xMcorrel As Variant
    xMcorrel = Application.Run(GetMacroRegId("fnMCorrel"), inprng, outrng, grouped, labels)
End Sub

Sub McorrelQ(Optional inprng As Variant, Optional outrng As Variant, Optional grouped As Variant, Optional labels As Variant)
Attribute McorrelQ.VB_Description = "Returns a correlation matrix that measures the correlation between two or more data sets"
Attribute McorrelQ.VB_HelpID = 3474
Attribute McorrelQ.VB_ProcData.VB_Invoke_Func = " \n3"
    Dim xMcorrelQ As Variant
    xMcorrelQ = Application.Run(GetMacroRegId("fnMCorrelQ"), inprng, outrng, grouped, labels)
End Sub

Sub Mcovar(inprng As Variant, Optional outrng As Variant, Optional grouped As Variant, Optional labels As Variant)
Attribute Mcovar.VB_Description = "Returns a covariance matrix that measures the covariance between two or more data sets"
Attribute Mcovar.VB_HelpID = 3475
Attribute Mcovar.VB_ProcData.VB_Invoke_Func = " \n3"
    Dim xMcovar As Variant
    xMcovar = Application.Run(GetMacroRegId("fnMCovar"), inprng, outrng, grouped, labels)
End Sub

Sub McovarQ(Optional inprng As Variant, Optional outrng As Variant, Optional grouped As Variant, Optional labels As Variant)
Attribute McovarQ.VB_Description = "Returns a covariance matrix that measures the covariance between two or more data sets"
Attribute McovarQ.VB_HelpID = 3476
Attribute McovarQ.VB_ProcData.VB_Invoke_Func = " \n3"
    Dim xMcovarQ As Variant
    xMcovarQ = Application.Run(GetMacroRegId("fnMCovarQ"), inprng, outrng, grouped, labels)
End Sub

Sub Moveavg(inprng As Variant, Optional outrng As Variant, Optional interval As Variant, Optional stderrs As Variant, Optional chart As Variant, Optional labels As Variant)
Attribute Moveavg.VB_Description = "Projects values in a forecast period, based on the average value of the variable over a specific number of preceding periods"
Attribute Moveavg.VB_HelpID = 3477
Attribute Moveavg.VB_ProcData.VB_Invoke_Func = " \n3"
    Dim xMoveavg As Variant
    xMoveavg = Application.Run(GetMacroRegId("fnMoveAvg"), inprng, outrng, interval, stderrs, chart, labels)
End Sub

Sub MoveavgQ(Optional inprng As Variant, Optional outrng As Variant, Optional interval As Variant, Optional stderrs As Variant, Optional chart As Variant, Optional labels As Variant)
Attribute MoveavgQ.VB_Description = "Projects values in a forecast period, based on the average value of the variable over a specific number of preceding periods"
Attribute MoveavgQ.VB_HelpID = 3478
Attribute MoveavgQ.VB_ProcData.VB_Invoke_Func = " \n3"
    Dim xMoveavgQ As Variant
    xMoveavgQ = Application.Run(GetMacroRegId("fnMoveAvgQ"), inprng, outrng, interval, stderrs, chart, labels)
End Sub

Sub Pttestm(inprng1 As Variant, inprng2 As Variant, Optional outrng As Variant, Optional labels As Variant, Optional alpha As Variant, Optional difference As Variant)
Attribute Pttestm.VB_Description = "Performs a paired two-sample Students t-Test"
Attribute Pttestm.VB_HelpID = 3479
Attribute Pttestm.VB_ProcData.VB_Invoke_Func = " \n3"
    Dim xPttestm As Variant
    xPttestm = Application.Run(GetMacroRegId("fnTtestM"), inprng1, inprng2, outrng, labels, alpha, difference)
End Sub

Sub PttestmQ(Optional inprng1 As Variant, Optional inprng2 As Variant, Optional outrng As Variant, Optional labels As Variant, Optional alpha As Variant, Optional difference As Variant)
Attribute PttestmQ.VB_Description = "Performs a paired two-sample Students t-Test"
Attribute PttestmQ.VB_HelpID = 3480
Attribute PttestmQ.VB_ProcData.VB_Invoke_Func = " \n3"
    Dim xPttestmQ As Variant
    xPttestmQ = Application.Run(GetMacroRegId("fnTtestMQ"), inprng1, inprng2, outrng, labels, alpha, difference)
End Sub

Sub Pttestv(inprng1 As Variant, inprng2 As Variant, Optional outrng As Variant, Optional labels As Variant, Optional alpha As Variant, Optional difference As Variant)
Attribute Pttestv.VB_Description = "Performs a two-sample Student t-Test, assuming unequal variances"
Attribute Pttestv.VB_HelpID = 3481
Attribute Pttestv.VB_ProcData.VB_Invoke_Func = " \n3"
    Dim xPttestv As Variant
    xPttestv = Application.Run(GetMacroRegId("fnTtestUeq"), inprng1, inprng2, outrng, labels, alpha, difference)
End Sub

Sub PttestvQ(Optional inprng1 As Variant, Optional inprng2 As Variant, Optional outrng As Variant, Optional labels As Variant, Optional alpha As Variant, Optional difference As Variant)
Attribute PttestvQ.VB_Description = "Performs a two-sample Student t-Test, assuming unequal variances"
Attribute PttestvQ.VB_HelpID = 3482
Attribute PttestvQ.VB_ProcData.VB_Invoke_Func = " \n3"
    Dim xPttestvQ As Variant
    xPttestvQ = Application.Run(GetMacroRegId("fnTtestUeqQ"), inprng1, inprng2, outrng, labels, alpha, difference)
End Sub

Sub Ttestm(inprng1 As Variant, inprng2 As Variant, Optional outrng As Variant, Optional labels As Variant, Optional alpha As Variant, Optional difference As Variant)
Attribute Ttestm.VB_Description = "Performs a two-sample Student t-Test for means, assuming equal variances"
Attribute Ttestm.VB_HelpID = 3491
Attribute Ttestm.VB_ProcData.VB_Invoke_Func = " \n3"
    Dim xTtestm As Variant
    xTtestm = Application.Run(GetMacroRegId("fnTtestEq"), inprng1, inprng2, outrng, labels, alpha, difference)
End Sub

Sub TtestmQ(Optional inprng1 As Variant, Optional inprng2 As Variant, Optional outrng As Variant, Optional labels As Variant, Optional alpha As Variant, Optional difference As Variant)
Attribute TtestmQ.VB_Description = "Performs a two-sample Student t-Test for means, assuming equal variances"
Attribute TtestmQ.VB_HelpID = 3492
Attribute TtestmQ.VB_ProcData.VB_Invoke_Func = " \n3"
    Dim xTtestmQ As Variant
    xTtestmQ = Application.Run(GetMacroRegId("fnTtestEqQ"), inprng1, inprng2, outrng, labels, alpha, difference)
End Sub

Sub zTestm(inprng1 As Variant, inprng2 As Variant, Optional outrng As Variant, Optional labels As Variant, Optional alpha As Variant, Optional difference As Variant, Optional var1 As Variant, Optional var2 As Variant)
Attribute zTestm.VB_Description = "Performs a two-sample z-test for means, assuming the two samples have known variances"
Attribute zTestm.VB_HelpID = 3493
Attribute zTestm.VB_ProcData.VB_Invoke_Func = " \n3"
    Dim xzTestm As Variant
    xzTestm = Application.Run(GetMacroRegId("fnZtestM"), inprng1, inprng2, outrng, labels, alpha, difference, var1, var2)
End Sub

Sub zTestmQ(Optional inprng1 As Variant, Optional inprng2 As Variant, Optional outrng As Variant, Optional labels As Variant, Optional alpha As Variant, Optional difference As Variant, Optional var1 As Variant, Optional var2 As Variant)
Attribute zTestmQ.VB_Description = "Performs a two-sample z-test for means, assuming the two samples have known variances"
Attribute zTestmQ.VB_HelpID = 3494
Attribute zTestmQ.VB_ProcData.VB_Invoke_Func = " \n3"
    Dim xzTestmQ As Variant
    xzTestmQ = Application.Run(GetMacroRegId("fnZtestMQ"), inprng1, inprng2, outrng, labels, alpha, difference, var1, var2)
End Sub

Sub Random(Optional outrng As Variant, Optional variables As Variant, Optional points As Variant, Optional distribution As Variant, Optional seed As Variant, Optional randarg1 As Variant, Optional randarg2 As Variant, Optional randarg3 As Variant, Optional randarg4 As Variant, Optional randarg5 As Variant)
Attribute Random.VB_Description = "Fills a range with independent random or patterned numbers drawn from one of several distributions"
Attribute Random.VB_HelpID = 3483
Attribute Random.VB_ProcData.VB_Invoke_Func = " \n3"
    Dim xRandom As Variant
    xRandom = Application.Run(GetMacroRegId("fnRandom"), outrng, variables, points, distribution, seed, randarg1, randarg2, randarg3, randarg4, randarg5)
End Sub

Sub RandomQ(Optional outrng As Variant, Optional variables As Variant, Optional points As Variant, Optional distribution As Variant, Optional seed As Variant, Optional randarg1 As Variant, Optional randarg2 As Variant, Optional randarg3 As Variant, Optional randarg4 As Variant, Optional randarg5 As Variant)
Attribute RandomQ.VB_Description = "Fills a range with independent random or patterned numbers drawn from one of several distributions"
Attribute RandomQ.VB_HelpID = 3484
Attribute RandomQ.VB_ProcData.VB_Invoke_Func = " \n3"
    Dim xRandomQ As Variant
    xRandomQ = Application.Run(GetMacroRegId("fnRandomQ"), outrng, variables, points, distribution, seed, randarg1, randarg2, randarg3, randarg4, randarg5)
End Sub

Sub RankPerc(inprng As Variant, Optional outrng As Variant, Optional grouped As Variant, Optional labels As Variant)
Attribute RankPerc.VB_Description = "Returns a table that contains the ordinal and percent rank of each value in a data set"
Attribute RankPerc.VB_HelpID = 3485
Attribute RankPerc.VB_ProcData.VB_Invoke_Func = " \n3"
    Dim xRankPerc As Variant
    xRankPerc = Application.Run(GetMacroRegId("fnRankPerc"), inprng, outrng, grouped, labels)
End Sub

Sub RankPercQ(Optional inprng As Variant, Optional outrng As Variant, Optional grouped As Variant, Optional labels As Variant)
Attribute RankPercQ.VB_Description = "Returns a table that contains the ordinal and percent rank of each value in a data set"
Attribute RankPercQ.VB_HelpID = 3486
Attribute RankPercQ.VB_ProcData.VB_Invoke_Func = " \n3"
    Dim xRankPercQ As Variant
    xRankPercQ = Application.Run(GetMacroRegId("fnRankPercQ"), inprng, outrng, grouped, labels)
End Sub

Sub Regress(inpyrng As Variant, Optional inpxrng As Variant, Optional constant As Variant, Optional labels As Variant, Optional confid As Variant, Optional soutrng As Variant, Optional residuals As Variant, Optional sresiduals As Variant, Optional rplots As Variant, Optional lplots As Variant, Optional routrng As Variant, Optional nplots As Variant, Optional poutrng As Variant)
Attribute Regress.VB_Description = "Perform multiple linear regression analysis"
Attribute Regress.VB_HelpID = 3487
Attribute Regress.VB_ProcData.VB_Invoke_Func = " \n3"
    Dim xRegress As Variant
    xRegress = Application.Run(GetMacroRegId("fnRegress"), inpyrng, inpxrng, constant, labels, confid, soutrng, residuals, sresiduals, rplots, lplots, routrng, nplots, poutrng)
End Sub

Sub RegressQ(Optional inpyrng As Variant, Optional inpxrng As Variant, Optional constant As Variant, Optional labels As Variant, Optional confid As Variant, Optional soutrng As Variant, Optional residuals As Variant, Optional sresiduals As Variant, Optional rplots As Variant, Optional lplots As Variant, Optional routrng As Variant, Optional nplots As Variant, Optional poutrng As Variant)
Attribute RegressQ.VB_Description = "Perform multiple linear regression analysis"
Attribute RegressQ.VB_HelpID = 3488
Attribute RegressQ.VB_ProcData.VB_Invoke_Func = " \n3"
    Dim xRegressQ As Variant
    xRegressQ = Application.Run(GetMacroRegId("fnRegressQ"), inpyrng, inpxrng, constant, labels, confid, soutrng, residuals, sresiduals, rplots, lplots, routrng, nplots, poutrng)
End Sub

Sub Sample(inprng As Variant, Optional outrng As Variant, Optional method As Variant, Optional rate As Variant, Optional labels As Variant)
Attribute Sample.VB_Description = "Samples data"
Attribute Sample.VB_HelpID = 3489
Attribute Sample.VB_ProcData.VB_Invoke_Func = " \n3"
    Dim xSample As Variant
    xSample = Application.Run(GetMacroRegId("fnSample"), inprng, outrng, method, rate, labels)
End Sub

Sub SampleQ(Optional inprng As Variant, Optional outrng As Variant, Optional method As Variant, Optional rate As Variant, Optional labels As Variant)
Attribute SampleQ.VB_Description = "Samples data"
Attribute SampleQ.VB_HelpID = 3490
Attribute SampleQ.VB_ProcData.VB_Invoke_Func = " \n3"
    Dim xSampleQ As Variant
    xSampleQ = Application.Run(GetMacroRegId("fnSampleQ"), inprng, outrng, method, rate, labels)
End Sub

' Setup & Registering functions

Sub auto_open()
    Application.EnableCancelKey = xlDisabled
    SetupFunctionIDs
    PickPlatform
    VerifyOpen
    RegisterFunctionIDs
End Sub

' O12:624902 - unregister analys32.xll if it's not installed so that funcres.xlam
'              closes and the UI is removed
Sub auto_close()
    Dim fATPInstalled As Boolean
    
    fATPInstalled = False
    Dim ai As Variant
    For Each ai In Application.AddIns
        If UCase(ai.Name) = c_sXllName Then
            fATPInstalled = ai.Installed
            Exit For
        End If
    Next ai
    
    If Not fATPInstalled Then
        Dim sQuote As String
        sQuote = """"
        Application.ExecuteExcel4Macro ("UNREGISTER(" & sQuote & c_sXllName & sQuote & ")")
    End If
    
End Sub

Private Sub VerifyOpen()
    s_sXllFullName = ""
    Dim sPathSep As String
    sPathSep = Application.PathSeparator
    s_sXllFullName = Application.LibraryPath & sPathSep & c_sAddinFolder & sPathSep & c_sXllName
    Dim theArray As Variant
    theArray = Application.RegisteredFunctions
    If Not (IsNull(theArray)) Then
        Dim i As Variant
        For i = LBound(theArray) To UBound(theArray)
            If (StrComp(theArray(i, 1), s_sXllFullName, vbTextCompare) = 0) Then
                Exit Sub
            End If
        Next i
    End If
    
    ThisWorkbook.Sheets("REG").Activate
    Dim XLLFound As Boolean
    XLLFound = Application.RegisterXLL(s_sXllFullName)
    If (XLLFound) Then
        Exit Sub
    End If

    MsgBox (ThisWorkbook.Sheets("Loc Table").Range("B12").Value)
    ThisWorkbook.Close (False)
End Sub

Private Sub PickPlatform()
    ThisWorkbook.Sheets("REG").Activate
    Range("C3").Select
End Sub

Private Sub RegisterFunctionIDs()
    If (s_sXllFullName = "") Then
        Exit Sub 'VerifyOpen failed
    End If
    Dim Quote As String
    Quote = String(1, 34)
    Dim i As Variant
    For i = LBound(FunctionIDs) To UBound(FunctionIDs)
        Dim StrCall
        StrCall = "REGISTER.ID(" & Quote & Replace(s_sXllFullName, Quote, Quote & Quote) & Quote & "," & Quote & FunctionIDs(i, 0) & Quote & ")"
        FunctionIDs(i, 1) = ExecuteExcel4Macro(StrCall)
    Next i
End Sub

Private Sub SetupFunctionIDs()
    FunctionIDs(0, 0) = "fnAnova1"
    FunctionIDs(1, 0) = "fnAnova2"
    FunctionIDs(2, 0) = "fnAnova3"
    FunctionIDs(3, 0) = "fnMCorrel"
    FunctionIDs(4, 0) = "fnMCovar"
    FunctionIDs(5, 0) = "fnDescr"
    FunctionIDs(6, 0) = "fnExpon"
    FunctionIDs(7, 0) = "fnFourier"
    FunctionIDs(8, 0) = "fnFtestV"
    FunctionIDs(9, 0) = "fnHistogram"
    FunctionIDs(10, 0) = "fnMoveAvg"
    FunctionIDs(11, 0) = "fnRandom"
    FunctionIDs(12, 0) = "fnRankPerc"
    FunctionIDs(13, 0) = "fnRegress"
    FunctionIDs(14, 0) = "fnSample"
    FunctionIDs(15, 0) = "fnTtestM"
    FunctionIDs(16, 0) = "fnTtestUeq"
    FunctionIDs(17, 0) = "fnTtestEq"
    FunctionIDs(18, 0) = "fnZtestM"
    FunctionIDs(19, 0) = "fnAnova1Q"
    FunctionIDs(20, 0) = "fnAnova2Q"
    FunctionIDs(21, 0) = "fnAnova3Q"
    FunctionIDs(22, 0) = "fnMCorrelQ"
    FunctionIDs(23, 0) = "fnMCovarQ"
    FunctionIDs(24, 0) = "fnDescrQ"
    FunctionIDs(25, 0) = "fnExponQ"
    FunctionIDs(26, 0) = "fnFourierQ"
    FunctionIDs(27, 0) = "fnFtestVQ"
    FunctionIDs(28, 0) = "fnHistogramQ"
    FunctionIDs(29, 0) = "fnMoveAvgQ"
    FunctionIDs(30, 0) = "fnRandomQ"
    FunctionIDs(31, 0) = "fnRankPercQ"
    FunctionIDs(32, 0) = "fnRegressQ"
    FunctionIDs(33, 0) = "fnSampleQ"
    FunctionIDs(34, 0) = "fnTtestMQ"
    FunctionIDs(35, 0) = "fnTtestUeqQ"
    FunctionIDs(36, 0) = "fnTtestEqQ"
    FunctionIDs(37, 0) = "fnZtestMQ"
End Sub

