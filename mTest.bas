Attribute VB_Name = "mTest"
Option Explicit

Private bRegressionTest As Boolean

Private Function ErrSrc(ByVal s As String) As String
    ErrSrc = "mTest." & s
End Function

Private Function RegressionTestInfo() As String
' ----------------------------------------------------
' Adds s to the Err.Description as an additional info.
' ----------------------------------------------------
    RegressionTestInfo = err.Description
    If Not bRegressionTest Then Exit Function
    
    If InStr(RegressionTestInfo, CONCAT) <> 0 _
    Then RegressionTestInfo = RegressionTestInfo & vbLf & vbLf & "Please notice that  this is a  r e g r e s s i o n  t e s t ! Click any but the ""Terminate"" button to continue with the test in case another one follows." _
    Else RegressionTestInfo = RegressionTestInfo & CONCAT & "Please notice that  this is a  r e g r e s s i o n  t e s t !  Click any but the ""Terminate"" button to continue with the test in case another one follows."

End Function

Public Sub Regression_Test()
' -----------------------------------------------------------------------------
' 1. This regression test requires the Conditional Compile Argument "Test = 1"
'    which provides additional buttons to continue with the next test after a
'    procedure which tests an error condition
' 2. The BoP/EoP statements in this regression test procedure produce one
'    execution trace at the end of this regression test provided the
'    Conditional Compile Argument "ExecTrace = 1". Attention must be paid on
'    the execution time however because it will include the time used by the
'    user action when an error message is displayed!
' 3. The Conditional Compile Argument "Debugging = 1" allows to identify the
'    code line which causes the error through an extra "Resume error code line"
'    button displayed with the error message and processed when clicked as
'    "Stop: Resume" when the button is clicked.
' ------------------------------------------------------------------------------
    
    On Error GoTo eh
    Const PROC = "Regression_Test"
    bRegressionTest = True
    
    mErH.BoP ErrSrc(PROC)
'    Test_1_BoP_EoP
    Test_2_Application_Error
    Test_3_VB_Runtime_Error
#If ExecTrace = 0 Then
    Test_6_Execution_Trace
#End If

xt: mErH.EoP ErrSrc(PROC)
    bRegressionTest = False
    Exit Sub
    
eh: mErH.ErrMsg errnumber:=err.Number, errsource:=ErrSrc(PROC), errdscrptn:=err.Description, errline:=Erl
End Sub

Public Sub Test_1_BoP_EoP()
' ---------------------------------------------------
' White-box- and regression-test procedure obligatory
' to be performed after any code modification.
' ---------------------------------------------------
    
    Const PROC = "Test_1_BoP_EoP"
    mErH.BoP ErrSrc(PROC)
    Test_1_BoP_EoP_TestProc_1a_missing_BoP
    
xt: mErH.EoP ErrSrc(PROC)
    Exit Sub

eh: If mErH.ErrMsg(err.Number, ErrSrc(PROC), err.Description, Erl) = ResumeError Then Stop: Resume
End Sub

Private Sub Test_1_BoP_EoP_TestProc_1a_missing_BoP()
' -----------------------------------------------------------
' The error handler is trying its very best not to confuse
' with unpaired BoP/EoP code lines. However, it depends at
' which level this is the case.
' -----------------------------------------------------------

    Const PROC = "Test_1_BoP_EoP_TestProc_1a_missing_BoP"
    
'    mErH.BoP ErrSrc(PROC)
    Test_1_BoP_EoP_TestProc_1b_paired_BoP_EoP
    Test_1_BoP_EoP_TestProc_1d_missing_EoP
    mErH.EoP ErrSrc(PROC)
    
    Exit Sub

eh:
    If mErH.ErrMsg(err.Number, ErrSrc(PROC), err.Description, Erl) = ResumeError Then Stop: Resume
End Sub

Private Sub Test_1_BoP_EoP_TestProc_1b_paired_BoP_EoP()
    
    Const PROC = "Test_1_BoP_EoP_TestProc_1b_paired_BoP_EoP"
    On Error GoTo eh
    
    mErH.BoP ErrSrc(PROC)
    Test_1_BoP_EoP_TestProc_1c_missing_EoC
    mErH.EoP ErrSrc(PROC)
    Exit Sub
    
eh:
    If mErH.ErrMsg(err.Number, ErrSrc(PROC), err.Description, Erl) = ResumeError Then Stop: Resume
End Sub

Private Sub Test_1_BoP_EoP_TestProc_1c_missing_EoC()
    
    Const PROC = "Test_1_BoP_EoP_TestProc_1c_missing_EoC"
    On Error GoTo eh
    
    mErH.BoP ErrSrc(PROC)
    BoC ErrSrc(PROC) & " trace of some code lines" ' missing EoC statement

xt: mErH.EoP ErrSrc(PROC)
    Exit Sub
    
eh: If mErH.ErrMsg(err.Number, ErrSrc(PROC), err.Description, Erl) = ResumeError Then Stop: Resume
End Sub

Private Sub Test_1_BoP_EoP_TestProc_1e_BoC_EoC()
    
    Const PROC = "Test_1_BoP_EoP_TestProc_1e_BoC_EoC"
    On Error GoTo eh
    
    mErH.BoP ErrSrc(PROC)
        
    Dim i As Long: Dim j As Long: j = 10000000
    BoC PROC & " code trace empty loop 1 to " & j
    For i = 1 To j
    Next i
    EoC PROC & " code trace empty loop 1 to " & j ' !!! the string must match with the BoC statement !!!
    
    mErH.EoP ErrSrc(PROC)
    Exit Sub

eh:
    If mErH.ErrMsg(err.Number, ErrSrc(PROC), err.Description, Erl) = ResumeError Then Stop: Resume
End Sub

Private Sub Test_1_BoP_EoP_TestProc_1d_missing_EoP()

    Const PROC = "Test_1_BoP_EoP_TestProc_1d_missing_EoP"
    On Error GoTo eh
    
    mErH.BoP ErrSrc(PROC)
    Test_1_BoP_EoP_TestProc_1e_BoC_EoC
    
    Exit Sub

eh:
    If mErH.ErrMsg(err.Number, ErrSrc(PROC), err.Description, Erl) = ResumeError Then Stop: Resume
End Sub

Public Sub Test_2_Application_Error()
' -----------------------------------------------------------
' This test procedure obligatory after any code modification.
' The option to continue with the next test procedure (in
' case this one runs within a regression test) is only
' displayed when the Conditional Compile Argument Test = 1
' The display of an execution trace along with this test
' requires a Conditional Compile Argument ExecTrace = 1.
' ------------------------------------------------------
    
    Const PROC = "Test_2_Application_Error"
    On Error GoTo eh
    
'    mTrc.DisplayedInfo = Detailed
    mErH.BoP ErrSrc(PROC)
    Test_2_Application_Error_TestProc_2a

xt: mErH.EoP ErrSrc(PROC)
    Exit Sub

eh: Select Case mErH.ErrMsg(err.Number, ErrSrc(PROC), err.Description, Erl)
        Case ResumeError: Stop: Resume
    End Select
End Sub

Private Sub Test_2_Application_Error_TestProc_2a()

    Const PROC = "Test_2_Application_Error_TestProc_2a"
    On Error GoTo eh
    
    mErH.BoP ErrSrc(PROC)
    Test_2_Application_Error_TestProc_2b
    mErH.EoP ErrSrc(PROC)
    Exit Sub

eh:
    If mErH.ErrMsg(err.Number, ErrSrc(PROC), err.Description, Erl) = ResumeError Then Stop: Resume
End Sub

Private Sub Test_2_Application_Error_TestProc_2b()
    
    Const PROC = "Test_2_Application_Error_TestProc_2b"
    On Error GoTo eh
    
    mErH.BoP ErrSrc(PROC)
    Test_2_Application_Error_TestProc_2c
    mErH.EoP ErrSrc(PROC)
    Exit Sub

eh:
    If mErH.ErrMsg(err.Number, ErrSrc(PROC), err.Description, Erl) = ResumeError Then Stop: Resume
End Sub

Private Sub Test_2_Application_Error_TestProc_2c()
' ------------------------------------------------
' Note: The line number is added just for test to
' demonstrate how it effects the error message.
' ------------------------------------------------
    
    On Error GoTo eh
    Const PROC = "Test_2_Application_Error_TestProc_2c"
    Dim sErrDscrptn As String

    mErH.BoP ErrSrc(PROC)
181 err.Raise AppErr(1), ErrSrc(PROC), _
        "This is a programmed i.e. an ""Application Error""!" & CONCAT & _
        "The function AppErr() has been used to turn the positive into a negative number by adding the VB constant 'vbObjectError' to assure an error number which does not conflict with a VB Runtime error. " & _
        "The ErrMsg identified the negative number as an ""Application Error"" and converted it back to the orginal positive number by means of the AppErr() function." & vbLf & _
        vbLf & _
        "Also note that this information is part of the raised error message but concatenated with two vertical bars indicating that it is an additional information regarding this error."

xt: mErH.EoP ErrSrc(PROC)
    Exit Sub

eh: sErrDscrptn = RegressionTestInfo
    Select Case mErH.ErrMsg(err.Number, ErrSrc(PROC), sErrDscrptn, Erl)
        Case ResumeError:       Stop: Resume
        Case ResumeNext:        Resume Next
        Case ExitAndContinue:   GoTo xt
    End Select
End Sub

Public Sub Test_3_VB_Runtime_Error()
' -----------------------------------------------
' - With Conditional Compile Argument BopEop = 0:
'   Display of the error with the error path only
' - With Conditional Compile Argument BopEop = 1:
'   Display of the error with the error path plus
'   Display of a full execution trace
'
' Requires:
' - Conditional Compile Argument ExecTrace = 1.
' -----------------------------------------------
    
    Const PROC = "Test_3_VB_Runtime_Error"
    On Error GoTo eh
    
    mTrc.DisplayedInfo = Detailed
    mErH.BoP ErrSrc(PROC)
    Test_3_VB_Runtime_Error_TestProc_3a

xt: mErH.EoP ErrSrc(PROC)
    Exit Sub

eh: Select Case mErH.ErrMsg(err.Number, ErrSrc(PROC), err.Description, Erl)
        Case ResumeError: Stop: Resume
        Case ResumeNext: Resume Next
        Case ExitAndContinue: GoTo xt
    End Select
End Sub

Private Sub Test_3_VB_Runtime_Error_TestProc_3a()

    Const PROC = "Test_3_VB_Runtime_Error_TestProc_3a"
    On Error GoTo eh

    mErH.BoP ErrSrc(PROC)
    Test_3_VB_Runtime_Error_TestProc_3b
    mErH.EoP ErrSrc(PROC)
    Exit Sub

eh:
    If mErH.ErrMsg(err.Number, ErrSrc(PROC), err.Description, Erl) = ResumeError Then Stop: Resume
End Sub

Private Sub Test_3_VB_Runtime_Error_TestProc_3b()
    
    Const PROC = "Test_3_VB_Runtime_Error_TestProc_3b"
    On Error GoTo eh

    mErH.BoP ErrSrc(PROC)
    Test_3_VB_Runtime_Error_TestProc_3c
    mErH.EoP ErrSrc(PROC)
    Exit Sub

eh:
    If mErH.ErrMsg(err.Number, ErrSrc(PROC), err.Description, Erl) = ResumeError Then Stop: Resume
End Sub

Private Sub Test_3_VB_Runtime_Error_TestProc_3c()

    Const PROC = "Test_3_VB_Runtime_Error_TestProc_3c"
    On Error GoTo eh
    
    mErH.BoP ErrSrc(PROC)
    Test_3_VB_Runtime_Error_TestProc_3d
    mErH.EoP ErrSrc(PROC)
    Exit Sub

eh:
    If mErH.ErrMsg(err.Number, ErrSrc(PROC), err.Description, Erl) = ResumeError Then Stop: Resume
End Sub

Private Sub Test_3_VB_Runtime_Error_TestProc_3d()
' ------------------------------------------------
' Note: The error line intentionally has no line
' number to demonstrate how it effects the error
' message.
' ------------------------------------------------
    
    On Error GoTo eh
    Const PROC = "Test_3_VB_Runtime_Error_TestProc_3d"
    Dim sErrDscrptn As String
    mErH.BoP ErrSrc(PROC)
    Dim l As Long
    l = 7 / 0

xt: mErH.EoP ErrSrc(PROC)
    Exit Sub

eh: sErrDscrptn = RegressionTestInfo
    Select Case mErH.ErrMsg(err.Number, ErrSrc(PROC), sErrDscrptn, Erl)
        Case ResumeError:       Stop: Resume
        Case ResumeNext:        Resume Next
        Case ExitAndContinue:   GoTo xt
    End Select
End Sub

Public Sub Test_4_DebugAndTest_with_ErrMsg()
' -----------------------------------------
' This test the Conditional Compile
' Argument DebugAndTest = 1 is required.
' -----------------------------------------
    On Error GoTo eh
    Const PROC = "Test_4_DebugAndTest_with_ErrMsg"
      
    mErH.BoP ErrSrc(PROC)
    Test_4_DebugAndTest_with_ErrMsg_TestProc_5a
    mErH.EoP ErrSrc(PROC)
    Exit Sub

eh:
    If mErH.ErrMsg(err.Number, ErrSrc(PROC), err.Description, Erl) = ResumeError Then Stop: Resume
End Sub

Private Sub Test_4_DebugAndTest_with_ErrMsg_TestProc_5a()

    Const PROC = "Test_5_DebugAndTest_with_ErrMsg_TestProc_5a"
    On Error GoTo eh
       
    mErH.BoP ErrSrc(PROC)
15  Debug.Print ThisWorkbook.Named
    mErH.EoP ErrSrc(PROC)
    Exit Sub
    
eh:
    Select Case mErH.ErrMsg(errnumber:=err.Number, errsource:=ErrSrc(PROC), errdscrptn:=err.Description, errline:=Erl)
        Case ResumeError: Stop: Resume ' Continue with F8 to end up at the code line which caused the error
    End Select
End Sub

Public Sub Test_5_No_Exit_Statement()
' -----------------------------------
' Exit statement missing
' -----------------------------------

    Const PROC = "Test_6_No_Exit_Statement"
    On Error GoTo eh
    
eh:
    If mErH.ErrMsg(err.Number, ErrSrc(PROC), err.Description, Erl) = ResumeError Then Stop: Resume
End Sub

Public Sub Test_6_Execution_Trace()
' ------------------------------------------------------
' White-box- and regression-test procedure obligatory
' to be performed after any code modification.
' Display of an execution trace along with this test
' requires a conditional compile argument ExecTrace = 1.
' ------------------------------------------------------
    
    Const PROC = "Test_6_Execution_Trace"
    On Error GoTo eh
'    mTrc.DisplayedInfo = Compact
    mTrc.DisplayedInfo = Detailed
    
    mTrc.BoP ErrSrc(PROC)
    Test_6_Execution_Trace_TestProc_6a
    mTrc.EoP ErrSrc(PROC)
    Exit Sub

eh:
    If mErH.ErrMsg(err.Number, ErrSrc(PROC), err.Description, Erl) = ResumeError Then Stop: Resume
End Sub

Private Sub Test_6_Execution_Trace_TestProc_6a()

    On Error GoTo eh
    Const PROC = "Test_6_Execution_Trace_TestProc_6a"
    
    mTrc.BoP ErrSrc(PROC)
    mTrc.BoC ErrSrc(PROC) & " call of 6b and 6c"
    Test_6_Execution_Trace_TestProc_6b
    Test_6_Execution_Trace_TestProc_6c
    mTrc.EoC ErrSrc(PROC) & " call of 6b and 6c"
    mTrc.EoP ErrSrc(PROC)
    Exit Sub

eh:
    If mErH.ErrMsg(err.Number, ErrSrc(PROC), err.Description, Erl) = ResumeError Then Stop: Resume
End Sub

Private Sub Test_6_Execution_Trace_TestProc_6b()
    
    Const PROC = "Test_6_Execution_Trace_TestProc_6b"
    On Error GoTo eh

    mTrc.BoP ErrSrc(PROC)
    
    Dim i As Long
    Dim s As String
    For i = 1 To 10000
        s = Application.Path ' to produce some execution time
    Next i
    
    mTrc.EoP ErrSrc(PROC)
    Exit Sub

eh:
    If mErH.ErrMsg(err.Number, ErrSrc(PROC), err.Description, Erl) = ResumeError Then Stop: Resume
End Sub

Private Sub Test_6_Execution_Trace_TestProc_6c()
    
    Const PROC = "Test_6_Execution_Trace_TestProc_6c"
    On Error GoTo eh

    mTrc.BoP ErrSrc(PROC)

xt: mTrc.EoP ErrSrc(PROC)
    Exit Sub

eh: mErH.ErrMsg err.Number, ErrSrc(PROC), err.Description, Erl
End Sub

