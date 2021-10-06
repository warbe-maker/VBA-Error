Attribute VB_Name = "mTest"
Option Explicit

Private bRegressionTest As Boolean

Private Function AppErr(ByVal app_err_no As Long) As Long
' ------------------------------------------------------------------------------
' Ensures that a programmed (i.e. an application) error numbers never conflicts
' with the number of a VB runtime error. Thr function returns a given positive
' number (app_err_no) with the vbObjectError added - which turns it into a
' negative value. When the provided number is negative it returns the original
' positive "application" error number e.g. for being used with an error message.
' ------------------------------------------------------------------------------
    AppErr = IIf(app_err_no < 0, app_err_no - vbObjectError, vbObjectError - app_err_no)
End Function

Private Function ErrSrc(ByVal s As String) As String
    ErrSrc = "mTest." & s
End Function

Private Function RegressionTestInfo() As String
' ----------------------------------------------------
' Adds s to the Err.Description as an additional info.
' ----------------------------------------------------
    RegressionTestInfo = Err.Description
    If Not bRegressionTest Then Exit Function
    
    If InStr(RegressionTestInfo, CONCAT) <> 0 _
    Then RegressionTestInfo = RegressionTestInfo & vbLf & vbLf & "Please notice that  this is a  r e g r e s s i o n  t e s t ! Click any but the ""Terminate"" button to continue with the test in case another one follows." _
    Else RegressionTestInfo = RegressionTestInfo & CONCAT & "Please notice that  this is a  r e g r e s s i o n  t e s t !  Click any but the ""Terminate"" button to continue with the test in case another one follows."

End Function

Public Sub Regression_Test()
' -----------------------------------------------------------------------------
' 1. This regression test requires the Conditional Compile Argument "Test = 1"
'    to run un-attended
' 2. The BoP/EoP statements in this regression test procedure produce one final
'    execution trace provided the Conditional Compile Argument "ExecTrace = 1".
' 3. Error conditions tested provide the asserted error number which bypasses
'    the display of the error message - which is documented in the execution
'    trace however. By avoiding a user action required when the error is
'    displayed allows a fully automated regression test.
' 4. In case any tests fails: The Conditional Compile Argument "Debugging = 1"
'    allows to identify the code line which causes the error through an extra
'    "Debug: Resume error code line" button displayed with the error message
'    and processed when clicked as "Stop: Resume" when the button is clicked.
' ------------------------------------------------------------------------------
    Const PROC = "Regression_Test"
    
    On Error GoTo eh
    
    bRegressionTest = True
    mTrc.DisplayedInfo = Compact
    
'    mErH.BoTP ErrSrc(PROC), AppErr(1), 11
    mErH.BoP ErrSrc(PROC)
    Test_1_Application_Error
    Test_2_VB_Runtime_Error

    Debug.Assert RecentErrors(1) = AppErr(1)
    Debug.Assert RecentErrors(2) = 11
    
xt: mErH.EoP ErrSrc(PROC)
    bRegressionTest = False
    Exit Sub
    
eh: mErH.ErrMsg err_source:=ErrSrc(PROC)
End Sub

Public Sub Test_1_Application_Error()
' -----------------------------------------------------------
' This test procedure obligatory after any code modification.
' The option to continue with the next test procedure (in
' case this one runs within a regression test) is only
' displayed when the Conditional Compile Argument Test = 1
' The display of an execution trace along with this test
' requires a Conditional Compile Argument ExecTrace = 1.
' ------------------------------------------------------
    Const PROC = "Test_1_Application_Error"
    
    On Error GoTo eh
    
    mErH.BoP ErrSrc(PROC)
    
    Test_1_Application_Error_TestProc_2a

    Debug.Assert mErH.MostRecentError = AppErr(1)
    
xt: mErH.EoP ErrSrc(PROC)
    Exit Sub

eh: Select Case mErH.ErrMsg(err_source:=ErrSrc(PROC))
        Case DebugOptResumeErrorLine: Stop: Resume
    End Select
End Sub

Private Sub Test_1_Application_Error_TestProc_2a()
    Const PROC = "Test_1_Application_Error_TestProc_2a"
    
    On Error GoTo eh
    
    mErH.BoP ErrSrc(PROC)
    Test_1_Application_Error_TestProc_2b
    mErH.EoP ErrSrc(PROC)
    Exit Sub

eh:
    If mErH.ErrMsg(err_source:=ErrSrc(PROC)) = DebugOptResumeErrorLine Then Stop: Resume
End Sub

Private Sub Test_1_Application_Error_TestProc_2b()
    Const PROC = "Test_1_Application_Error_TestProc_2b"
    
    On Error GoTo eh
    
    mErH.BoP ErrSrc(PROC)
    Test_1_Application_Error_TestProc_2c
    mErH.EoP ErrSrc(PROC)
    Exit Sub

eh:
    If mErH.ErrMsg(err_source:=ErrSrc(PROC)) = DebugOptResumeErrorLine Then Stop: Resume
End Sub

Private Sub Test_1_Application_Error_TestProc_2c()
' ------------------------------------------------
' Note: The line number is added just for test to
' demonstrate how it effects the error message.
' ------------------------------------------------
    Const PROC = "Test_1_Application_Error_TestProc_2c"
    
    On Error GoTo eh

    mErH.BoP ErrSrc(PROC)
181 Err.Raise AppErr(1), ErrSrc(PROC), _
        "This is a programmed i.e. an ""Application Error""!" & CONCAT & _
        "The AppErr service has been used to turn the positive into a negative number by adding " & _
        "the VB constant 'vbObjectError' to assure the error number is not confused with a VB Runtime error. " & _
        "The ErrMsg service used the AppErr service to identify the number as an 'Application Error' " & _
        "and turn the negative number back into the orginal positive number." & vbLf & _
        vbLf & _
        "By the way: Note that all the above information had been provided with the err.Description " & _
        "by concatenating it with two vertical bars indicating that it as this additional information."

xt: mErH.EoP ErrSrc(PROC)
    Exit Sub

eh: Select Case mErH.ErrMsg(err_source:=ErrSrc(PROC), err_dscrptn:=RegressionTestInfo)
        Case DebugOptResumeErrorLine:       Stop: Resume
        Case DebugOptResumeNext:        Resume Next
        Case DebugOptCleanExit:   GoTo xt
    End Select
End Sub

Public Sub Test_2_VB_Runtime_Error()
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
    Const PROC = "Test_2_VB_Runtime_Error"
    
    On Error GoTo eh
    
    mErH.BoP ErrSrc(PROC)
    Test_2_VB_Runtime_Error_TestProc_3a

    Debug.Assert mErH.MostRecentError = 11

xt: mErH.EoP ErrSrc(PROC)
    Exit Sub

eh: mErH.ErrMsg err_source:=ErrSrc(PROC)
    Select Case mErH.ErrReply
        Case DebugOptResumeErrorLine: Stop: Resume
        Case DebugOptResumeNext: Resume Next
        Case DebugOptCleanExit: GoTo xt
    End Select
End Sub

Private Sub Test_2_VB_Runtime_Error_TestProc_3a()
    Const PROC = "Test_2_VB_Runtime_Error_TestProc_3a"
    
    On Error GoTo eh

    mErH.BoP ErrSrc(PROC)
    Test_2_VB_Runtime_Error_TestProc_3b
    mErH.EoP ErrSrc(PROC)
    Exit Sub

eh:
    If mErH.ErrMsg(err_source:=ErrSrc(PROC)) = DebugOptResumeErrorLine Then Stop: Resume
End Sub

Private Sub Test_2_VB_Runtime_Error_TestProc_3b()
    Const PROC = "Test_2_VB_Runtime_Error_TestProc_3b"
    
    On Error GoTo eh

    mErH.BoP ErrSrc(PROC)
    Test_2_VB_Runtime_Error_TestProc_3c
    mErH.EoP ErrSrc(PROC)
    Exit Sub

eh:
    If mErH.ErrMsg(err_source:=ErrSrc(PROC)) = DebugOptResumeErrorLine Then Stop: Resume
End Sub

Private Sub Test_2_VB_Runtime_Error_TestProc_3c()
    Const PROC = "Test_2_VB_Runtime_Error_TestProc_3c"
    
    On Error GoTo eh
    
    mErH.BoP ErrSrc(PROC)
    Test_2_VB_Runtime_Error_TestProc_3d test_arg1:="Test string", test_arg2:=20.5
    mErH.EoP ErrSrc(PROC)
    Exit Sub

eh:
    If mErH.ErrMsg(err_source:=ErrSrc(PROC)) = DebugOptResumeErrorLine Then Stop: Resume
End Sub

Private Sub Test_2_VB_Runtime_Error_TestProc_3d( _
      ByVal test_arg1 As String, _
      ByVal test_arg2 As Currency)
' ------------------------------------------------
' Note: The error line intentionally has no line
' number to demonstrate how it effects the error
' message.
' ------------------------------------------------
    Const PROC = "Test_2_VB_Runtime_Error_TestProc_3d"
    
    On Error GoTo eh
    
    mErH.BoP ErrSrc(PROC), "test_arg1 = ", test_arg1, "test_arg2 = ", test_arg2
    Dim l As Long
    l = 7 / 0

xt: mErH.EoP ErrSrc(PROC)
    Exit Sub

eh: Select Case mErH.ErrMsg(err_source:=ErrSrc(PROC), err_dscrptn:=RegressionTestInfo)
        Case DebugOptResumeErrorLine:       Stop: Resume
        Case DebugOptResumeNext:            Resume Next
        Case DebugOptCleanExit:             GoTo xt
    End Select
End Sub

Public Sub Test_4_DebugAndTest_with_ErrMsg()
' -----------------------------------------
' This test the Conditional Compile
' Argument DebugAndTest = 1 is required.
' -----------------------------------------
    Const PROC = "Test_4_DebugAndTest_with_ErrMsg"
    
    On Error GoTo eh
      
    mErH.BoP ErrSrc(PROC)
    Test_4_DebugAndTest_with_ErrMsg_TestProc_5a
    mErH.EoP ErrSrc(PROC)
    Exit Sub

eh:
    If mErH.ErrMsg(err_source:=ErrSrc(PROC)) = DebugOptResumeErrorLine Then Stop: Resume
End Sub

Private Sub Test_4_DebugAndTest_with_ErrMsg_TestProc_5a()
    Const PROC = "Test_5_DebugAndTest_with_ErrMsg_TestProc_5a"
    
    On Error GoTo eh
       
    mErH.BoP ErrSrc(PROC)
15  Debug.Print ThisWorkbook.Named
    mErH.EoP ErrSrc(PROC)
    Exit Sub
    
eh:
    Select Case mErH.ErrMsg(err_source:=ErrSrc(PROC))
        Case DebugOptResumeErrorLine: Stop: Resume ' Continue with F8 to end up at the code line which caused the error
    End Select
End Sub

Public Sub Test_5_No_Exit_Statement()
' -----------------------------------
' Exit statement missing
' -----------------------------------
    Const PROC = "Test_6_No_Exit_Statement"
    
    On Error GoTo eh
    
eh:
    If mErH.ErrMsg(err_source:=ErrSrc(PROC)) = DebugOptResumeErrorLine Then Stop: Resume
End Sub

Public Sub Test_6_VB_Runtime_Error_Pass_on()
' ----------------------------------------------------------------------
' About this test (with the Conditional Compile Argument Debugging = 1):
'
' The tester presses the 'Pass on' button in order to finally end up with
' the error message is displayed again at the 'Entry.Procedure' now with
' the full path to the error displayed - which is not available with the
' initial error message because only the 'Entry-Procedure' is known but
' none of the sub-procedures. = 1.
' ----------------------------------------------------------------------
    Const PROC = "Test_6_VB_Runtime_Error_Pass_on"
    
    On Error GoTo eh
    
    mErH.BoP ErrSrc(PROC)
    Test_6_VB_Runtime_Error_TestProc_3a

    Debug.Assert mErH.MostRecentError = 11

xt: mErH.EoP ErrSrc(PROC)
    Exit Sub

eh: mErH.ErrMsg ErrSrc(PROC)
    Select Case mErH.ErrReply
        Case DebugOptResumeErrorLine:   Stop: Resume
        Case DebugOptResumeNext:        Resume Next
        Case DebugOptCleanExit:         GoTo xt
    End Select
End Sub

Private Sub Test_6_VB_Runtime_Error_TestProc_3a()
    Const PROC = "Test_6_VB_Runtime_Error_TestProc_3a"
    
    On Error GoTo eh
    Test_6_VB_Runtime_Error_TestProc_3b

xt: Exit Sub

eh: Select Case mErH.ErrMsg(err_source:=ErrSrc(PROC), err_dscrptn:=RegressionTestInfo)
        Case DebugOptResumeErrorLine:   Stop: Resume
        Case DebugOptResumeNext:        Resume Next
        Case DebugOptCleanExit:         GoTo xt
    End Select
End Sub

Private Sub Test_6_VB_Runtime_Error_TestProc_3b()
    Const PROC = "Test_6_VB_Runtime_Error_TestProc_3b"
    
    On Error GoTo eh
    Test_6_VB_Runtime_Error_TestProc_3c

xt: Exit Sub

eh: Select Case mErH.ErrMsg(err_source:=ErrSrc(PROC), err_dscrptn:=RegressionTestInfo)
        Case DebugOptResumeErrorLine:   Stop: Resume
        Case DebugOptResumeNext:        Resume Next
        Case DebugOptCleanExit:         GoTo xt
    End Select
End Sub

Private Sub Test_6_VB_Runtime_Error_TestProc_3c()
    Const PROC = "Test_6_VB_Runtime_Error_TestProc_3c"
    
    On Error GoTo eh
    Test_6_VB_Runtime_Error_TestProc_3d test_arg1:="Test string", test_arg2:=20.5

xt: Exit Sub

eh: Select Case mErH.ErrMsg(err_source:=ErrSrc(PROC), err_dscrptn:=RegressionTestInfo)
        Case DebugOptResumeErrorLine:   Stop: Resume
        Case DebugOptResumeNext:        Resume Next
        Case DebugOptCleanExit:         GoTo xt
    End Select
End Sub

Private Sub Test_6_VB_Runtime_Error_TestProc_3d( _
      ByVal test_arg1 As String, _
      ByVal test_arg2 As Currency)
' ------------------------------------------------
' Note: The error line intentionally has no line
' number to demonstrate how it effects the error
' message.
' ------------------------------------------------
    Const PROC = "Test_6_VB_Runtime_Error_TestProc_3d"
    
    On Error GoTo eh
    Dim l As Long
    l = 7 / 0

xt: Exit Sub

eh: Select Case mErH.ErrMsg(err_source:=ErrSrc(PROC), err_dscrptn:=RegressionTestInfo)
        Case DebugOptResumeErrorLine:   Stop: Resume
        Case DebugOptResumeNext:        Resume Next
        Case DebugOptCleanExit:         GoTo xt
    End Select
End Sub
