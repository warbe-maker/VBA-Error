Attribute VB_Name = "mErHTest"
Option Explicit
' ----------------------------------------------------------------------------
' Standard Module mErHTest:
' =========================
'
' Uses (for the test only): mBasic, fMsg/mMsg, mTrc
'
' W. Rauschenberger Berlin, June 2023
' See: "https://github.com/warbe-maker/VBA-Error"
' ----------------------------------------------------------------------------
#If XcTrc_clsTrc = 1 Then
    Public Trc As New clsTrc
#End If

Private Function AppErr(ByVal app_err_no As Long) As Long
' ------------------------------------------------------------------------------
' Ensures that a programmed (i.e. an application) error numbers never conflicts
' with the number of a VB runtime error. Thr function returns a given positive
' number (app_err_no) with the vbObjectError added - which turns it into a
' negative value. When the provided number is negative it returns the original
' positive "application" error number e.g. for being used with an error message.
' ------------------------------------------------------------------------------
    If app_err_no >= 0 Then AppErr = app_err_no + vbObjectError Else AppErr = Abs(app_err_no - vbObjectError)
End Function

Private Function ErrSrc(ByVal s As String) As String
    ErrSrc = "mErHTest." & s
End Function

Public Sub Test_0_Regression()
' -----------------------------------------------------------------------------
' 1. This regression test requires the Cond. Comp. Arg.:
'    `Debugging = 1 : ErHComp = 1 : MsgComp = 1`
'    and `XcTrc_clsTrc = 0 : XcTrc_mTrc = 1`
'    or  `XcTrc_clsTrc = 1 : XcTrc_mTrc = 0`
'    I.e. it can use both variants of the Execution Trace component.
'
' 2. The BoP/EoP statements in this regression test procedure produce one final
'    execution trace.
'
' 3. Explicitly tested error conditions are bypassed by mErH.Regression = True
'    and the error asserted by mErH.Asserted ....
'
' 4. In case any tests fails the Debugging option supports 'resume error line'
' ------------------------------------------------------------------------------
    Const PROC = "Test_0_Regression"
    
    On Error GoTo eh
    
    '~~ Initializations (must be done prior the first BoP!)
#If XcTrc_mTrc = 1 Then
    mTrc.FileName = "RegressionTest.ExecTrace.log"
    mTrc.Title = "Regression Test mErH"
    mTrc.NewFile
#ElseIf XcTrc_clsTrc = 1 Then
    Set Trc = New clsTrc
    With Trc
        .FileName = "RegressionTest.ExecTrace.log"
        .Title = "Regression Test mErH"
        .NewFile
    End With
#End If
    
    mErH.Regression = True ' to bypass Asserted errors
      
    mBasic.BoP ErrSrc(PROC)
    Test_1_Application_Error
    Test_2_VB_Runtime_Error
    
xt: mBasic.EoP ErrSrc(PROC)
    mErH.Regression = False
    
#If XcTc_mTrc = 1 Then
    mTrc.Dsply
#ElseIf XcTrc_clsTrc = 1 Then
    Trc.Dsply
#End If
    Exit Sub
    
eh: Select Case mBasic.ErrMsg(ErrSrc(PROC))
        Case vbResume:  Stop: Resume
        Case Else:      GoTo xt
    End Select
End Sub

Public Sub Test_1_Application_Error()
' -----------------------------------------------------------
' This test procedure obligatory after any code modification.
' The option to continue with the next test procedure (in
' case this one runs within a regression test) is only
' displayed when the Cond. Comp. Arg. Test = 1
' The display of an execution trace along with this test
' requires a Cond. Comp. Arg. ExecTrace = 1.
' ------------------------------------------------------
    Const PROC = "Test_1_Application_Error"
    
    On Error GoTo eh
    mBasic.BoP ErrSrc(PROC)
    
    mErH.Asserted AppErr(1)
    Test_1_Application_Error_TestProc_2a
  
xt: mBasic.EoP ErrSrc(PROC)
    Exit Sub

eh: Select Case mBasic.ErrMsg(ErrSrc(PROC))
        Case vbResume:  Stop: Resume
        Case Else:      GoTo xt
    End Select
End Sub

Private Sub Test_1_Application_Error_TestProc_2a()
    Const PROC = "Test_1_Application_Error_TestProc_2a"
    
    On Error GoTo eh
    
    mBasic.BoP ErrSrc(PROC)
    Test_1_Application_Error_TestProc_2b
    
xt: mBasic.EoP ErrSrc(PROC)
    Exit Sub

eh: Select Case mBasic.ErrMsg(ErrSrc(PROC))
        Case vbResume:  Stop: Resume
        Case Else:      GoTo xt
    End Select
End Sub

Private Sub Test_1_Application_Error_TestProc_2b()
    Const PROC = "Test_1_Application_Error_TestProc_2b"
    
    On Error GoTo eh
    
    mBasic.BoP ErrSrc(PROC)
    Test_1_Application_Error_TestProc_2c
    
xt: mBasic.EoP ErrSrc(PROC)
    Exit Sub

eh: Select Case mBasic.ErrMsg(ErrSrc(PROC))
        Case vbResume:  Stop: Resume
        Case Else:      GoTo xt
    End Select
End Sub

Private Sub Test_1_Application_Error_TestProc_2c()
' ------------------------------------------------
' Note: The line number is added just for test to
' demonstrate how it effects the error message.
' ------------------------------------------------
    Const PROC = "Test_1_Application_Error_TestProc_2c"
    
    On Error GoTo eh

    mBasic.BoP ErrSrc(PROC)
    mErH.Asserted AppErr(1) ' has only an effect when mErh.Regression = True
    
181 Err.Raise AppErr(1), ErrSrc(PROC), _
        "This is a programmed i.e. an ""Application Error""!" & CONCAT & _
        "The AppErr service has been used to turn the positive into a negative number by adding " & _
        "the VB constant 'vbObjectError' to assure the error number is not confused with a VB Runtime error. " & _
        "The ErrMsg service used the AppErr service to identify the number as an 'Application Error' " & _
        "and turn the negative number back into the orginal positive number." & vbLf & _
        vbLf & _
        "By the way: Note that all the above information had been provided with the err.Description " & _
        "by concatenating it with two vertical bars indicating that it as this additional information."

xt: mBasic.EoP ErrSrc(PROC)
    Exit Sub

eh: Select Case mBasic.ErrMsg(ErrSrc(PROC))
        Case vbResume:  Stop: Resume
        Case Else:      GoTo xt
    End Select
End Sub

Public Sub Test_2_VB_Runtime_Error()
' -------------------------------------------------------------------------------
' With Cond. Comp. Arg.:
' - MsgComp = 0 : ErHComp = 0  Display of the error with VBA.MsgBox
' - MsgComp = 1 : ErHComp = 0  Display of the error without error path
' - ErHComp = 1                Display of the error with the path to the error
' - ExecTrace = 1              Display of the test execution trace
' -------------------------------------------------------------------------------
    Const PROC = "Test_2_VB_Runtime_Error"
    
    On Error GoTo eh
    
    mBasic.BoP ErrSrc(PROC)
    mErH.Asserted 11
    Test_2_VB_Runtime_Error_TestProc_3a

xt: mBasic.EoP ErrSrc(PROC)
    Exit Sub

eh: Select Case mBasic.ErrMsg(ErrSrc(PROC))
        Case vbResume:  Stop: Resume
        Case Else:      GoTo xt
    End Select
End Sub

Private Sub Test_2_VB_Runtime_Error_TestProc_3a()
    Const PROC = "Test_2_VB_Runtime_Error_TestProc_3a"
    
    On Error GoTo eh

    mBasic.BoP ErrSrc(PROC)
    Test_2_VB_Runtime_Error_TestProc_3b
    
xt: mBasic.EoP ErrSrc(PROC)
    Exit Sub

eh: Select Case mBasic.ErrMsg(ErrSrc(PROC))
        Case vbResume:  Stop: Resume
        Case Else:      GoTo xt
    End Select
End Sub

Private Sub Test_2_VB_Runtime_Error_TestProc_3b()
    Const PROC = "Test_2_VB_Runtime_Error_TestProc_3b"
    
    On Error GoTo eh

    mBasic.BoP ErrSrc(PROC)
    Test_2_VB_Runtime_Error_TestProc_3c
    
xt: mBasic.EoP ErrSrc(PROC)
    Exit Sub

eh: Select Case mBasic.ErrMsg(ErrSrc(PROC))
        Case vbResume:  Stop: Resume
        Case Else:      GoTo xt
    End Select
End Sub

Private Sub Test_2_VB_Runtime_Error_TestProc_3c()
    Const PROC = "Test_2_VB_Runtime_Error_TestProc_3c"
    
    On Error GoTo eh
    
    mBasic.BoP ErrSrc(PROC)
    Test_2_VB_Runtime_Error_TestProc_3d "Test string", 20.5
    
xt: mBasic.EoP ErrSrc(PROC)
    Exit Sub

eh: Select Case mBasic.ErrMsg(ErrSrc(PROC))
        Case vbResume:  Stop: Resume
        Case Else:      GoTo xt
    End Select
End Sub

Private Sub Test_2_VB_Runtime_Error_TestProc_3d( _
      ByVal test_arg1 As String, _
      ByVal test_arg2 As Currency)
' ----------------------------------------------------------------------------
' Test and demonstrate an error message without knowing the line of error
' (which is common) but still displaying a path to the error plus a debugging
' option button - provided the Cond. Comp. Arg. 'Debuggging = 1'.
' ----------------------------------------------------------------------------
    Const PROC = "Test_2_VB_Runtime_Error_TestProc_3d"
    
    On Error GoTo eh
    
    mBasic.BoP ErrSrc(PROC), "test_arg1=" & test_arg1 & ", test_arg2=" & test_arg2
    Dim l As Long
    l = 7 / 0

xt: mBasic.EoP ErrSrc(PROC)
    Exit Sub

eh: Select Case mBasic.ErrMsg(ErrSrc(PROC))
        Case vbResume:  Stop: Resume
        Case Else:      GoTo xt
    End Select
End Sub

Public Sub Test_4_DebugAndTest_with_ErrMsg()
' -----------------------------------------
' This test the Conditional Compile
' Argument DebugAndTest = 1 is required.
' -----------------------------------------
    Const PROC = "Test_4_DebugAndTest_with_ErrMsg"
    
    On Error GoTo eh
      
    mBasic.BoP ErrSrc(PROC)
    Test_4_DebugAndTest_with_ErrMsg_TestProc_5a
    
xt: mBasic.EoP ErrSrc(PROC)
    Exit Sub

eh: Select Case mBasic.ErrMsg(ErrSrc(PROC))
        Case vbResume:  Stop: Resume
        Case Else:      GoTo xt
    End Select
End Sub

Private Sub Test_4_DebugAndTest_with_ErrMsg_TestProc_5a()
    Const PROC = "Test_5_DebugAndTest_with_ErrMsg_TestProc_5a"
    
    On Error GoTo eh
       
    mBasic.BoP ErrSrc(PROC)
15  Debug.Print ThisWorkbook.Named
    
xt: mBasic.EoP ErrSrc(PROC)
    Exit Sub
    
eh: Select Case mBasic.ErrMsg(ErrSrc(PROC))
        Case vbResume:  Stop: Resume
        Case Else:      GoTo xt
    End Select
End Sub

Public Sub Test_5_No_Exit_Statement()
' -----------------------------------
' Exit statement missing
' -----------------------------------
    Const PROC = "Test_6_No_Exit_Statement"
    
    On Error GoTo eh
xt: Exit Sub

eh: Select Case mBasic.ErrMsg(ErrSrc(PROC))
        Case vbResume:  Stop: Resume
        Case Else:      GoTo xt
    End Select
End Sub

Public Sub Test_6_VB_Runtime_Error_Pass_on()
' ----------------------------------------------------------------------
' About this test (with the Cond. Comp. Arg. Debugging = 1):
'
' The tester presses the 'Pass on' button in order to finally end up with
' the error message is displayed again at the 'Entry.Procedure' now with
' the full path to the error displayed - which is not available with the
' initial error message because only the 'Entry-Procedure' is known but
' none of the sub-procedures. = 1.
' ----------------------------------------------------------------------
    Const PROC = "Test_6_VB_Runtime_Error_Pass_on"
    
    On Error GoTo eh
    
    mBasic.BoP ErrSrc(PROC)
    mErH.Asserted AppErr(1)
    Test_6_VB_Runtime_Error_TestProc_3a

xt: mBasic.EoP ErrSrc(PROC)
    Exit Sub

eh: Select Case mBasic.ErrMsg(ErrSrc(PROC))
        Case vbResume:  Stop: Resume
        Case Else:      GoTo xt
    End Select
End Sub

Private Sub Test_6_VB_Runtime_Error_TestProc_3a()
    Const PROC = "Test_6_VB_Runtime_Error_TestProc_3a"
    
    On Error GoTo eh
    Test_6_VB_Runtime_Error_TestProc_3b

xt: Exit Sub

eh: Select Case mBasic.ErrMsg(ErrSrc(PROC))
        Case vbResume:  Stop: Resume
        Case Else:      GoTo xt
    End Select
End Sub

Private Sub Test_6_VB_Runtime_Error_TestProc_3b()
    Const PROC = "Test_6_VB_Runtime_Error_TestProc_3b"
    
    On Error GoTo eh
    Test_6_VB_Runtime_Error_TestProc_3c

xt: Exit Sub

eh: Select Case mBasic.ErrMsg(ErrSrc(PROC))
        Case vbResume:  Stop: Resume
        Case Else:      GoTo xt
    End Select
End Sub

Private Sub Test_6_VB_Runtime_Error_TestProc_3c()
    Const PROC = "Test_6_VB_Runtime_Error_TestProc_3c"
    
    On Error GoTo eh
    Test_6_VB_Runtime_Error_TestProc_3d test_arg1:="Test string", test_arg2:=20.5

xt: Exit Sub

eh: Select Case mBasic.ErrMsg(ErrSrc(PROC))
        Case vbResume:  Stop: Resume
        Case Else:      GoTo xt
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

eh: Select Case mBasic.ErrMsg(ErrSrc(PROC))
        Case vbResume:  Stop: Resume
        Case Else:      GoTo xt
    End Select
End Sub

Public Sub Test_7_VB_Runtime_Error_EntryProc_known()
' -------------------------------------------------------------------------------
' Test of a couple of nested procedures where only the 'Entry Procedure' has
' a BoP statement. The effect is that there is no 'path to the error' available
' except when the error is passed on to ther 'Entry Procedure' which is only the
' case when the debugging option is turned off (Cond. Comp. Arg.
' 'Debugging = 0'.
' -------------------------------------------------------------------------------
    Const PROC = "Test_7_VB_Runtime_Error_EntryProc_known"
    
    On Error GoTo eh
    
    mBasic.BoP ErrSrc(PROC)
    Test_7_VB_Runtime_Error_TestProc_3a

    Debug.Assert mErH.MostRecentError = 11

xt: mBasic.EoP ErrSrc(PROC)
    Exit Sub

eh: Select Case mBasic.ErrMsg(ErrSrc(PROC))
        Case vbResume:  Stop: Resume
        Case Else:      GoTo xt
    End Select
End Sub

Private Sub Test_7_VB_Runtime_Error_TestProc_3a()
    Const PROC = "Test_7_VB_Runtime_Error_TestProc_3a"
    
    On Error GoTo eh

    Test_7_VB_Runtime_Error_TestProc_3b
    
xt: Exit Sub

eh: Select Case mBasic.ErrMsg(ErrSrc(PROC))
        Case vbResume:  Stop: Resume
        Case Else:      GoTo xt
    End Select
End Sub

Private Sub Test_7_VB_Runtime_Error_TestProc_3b()
    Const PROC = "Test_7_VB_Runtime_Error_TestProc_3b"
    
    On Error GoTo eh

    Test_7_VB_Runtime_Error_TestProc_3c
    
xt: Exit Sub

eh: Select Case mBasic.ErrMsg(ErrSrc(PROC))
        Case vbResume:  Stop: Resume
        Case Else:      GoTo xt
    End Select
End Sub

Private Sub Test_7_VB_Runtime_Error_TestProc_3c()
    Const PROC = "Test_7_VB_Runtime_Error_TestProc_3c"
    
    On Error GoTo eh
    
    Test_7_VB_Runtime_Error_TestProc_3d test_arg1:="Test string", test_arg2:=20.5
    
xt: Exit Sub

eh: Select Case mBasic.ErrMsg(ErrSrc(PROC))
        Case vbResume:  Stop: Resume
        Case Else:      GoTo xt
    End Select
End Sub

Private Sub Test_7_VB_Runtime_Error_TestProc_3d( _
      ByVal test_arg1 As String, _
      ByVal test_arg2 As Currency)
' ------------------------------------------------
' Note: The error line intentionally has no line
' number to demonstrate how it effects the error
' message.
' ------------------------------------------------
    Const PROC = "Test_7_VB_Runtime_Error_TestProc_3d"
    
    On Error GoTo eh
    
    Dim l As Long
    l = 7 / 0

xt: Exit Sub

eh: Select Case mBasic.ErrMsg(ErrSrc(PROC))
        Case vbResume:  Stop: Resume
        Case Else:      GoTo xt
    End Select
End Sub

Public Sub Test_8_VB_Runtime_Error_EntryProc_unknown()
' -------------------------------------------------------------------------------
' Test of a couple of nested procedures where only the 'Entry Procedure' has
' a BoP statement. The effect is that there is no 'path to the error' available
' except when the error is passed on to ther 'Entry Procedure' which is only the
' case when the debugging option is turned off (Cond. Comp. Arg.
' 'Debugging = 0'.
' -------------------------------------------------------------------------------
    Const PROC = "Test_8_VB_Runtime_Error_EntryProc_unknown"
    
    On Error GoTo eh
    
    Test_8_VB_Runtime_Error_TestProc_3a

    Debug.Assert mErH.MostRecentError = 11

xt: Exit Sub

eh: Select Case mBasic.ErrMsg(ErrSrc(PROC))
        Case vbResume:  Stop: Resume
        Case Else:      GoTo xt
    End Select
End Sub

Private Sub Test_8_VB_Runtime_Error_TestProc_3a()
    Const PROC = "Test_8_VB_Runtime_Error_TestProc_3a"
    
    On Error GoTo eh

    Test_8_VB_Runtime_Error_TestProc_3b
    
xt: Exit Sub

eh: Select Case mBasic.ErrMsg(ErrSrc(PROC))
        Case vbResume:  Stop: Resume
        Case Else:      GoTo xt
    End Select
End Sub

Private Sub Test_8_VB_Runtime_Error_TestProc_3b()
    Const PROC = "Test_8_VB_Runtime_Error_TestProc_3b"
    
    On Error GoTo eh

    Test_8_VB_Runtime_Error_TestProc_3c
    
xt: Exit Sub

eh: Select Case mBasic.ErrMsg(ErrSrc(PROC))
        Case vbResume:  Stop: Resume
        Case Else:      GoTo xt
    End Select
End Sub

Private Sub Test_8_VB_Runtime_Error_TestProc_3c()
    Const PROC = "Test_8_VB_Runtime_Error_TestProc_3c"
    
    On Error GoTo eh
    
    Test_8_VB_Runtime_Error_TestProc_3d test_arg1:="Test string", test_arg2:=20.5
    
xt: Exit Sub

eh: Select Case mBasic.ErrMsg(ErrSrc(PROC))
        Case vbResume:  Stop: Resume
        Case Else:      GoTo xt
    End Select
End Sub

Private Sub Test_8_VB_Runtime_Error_TestProc_3d( _
      ByVal test_arg1 As String, _
      ByVal test_arg2 As Currency)
' ------------------------------------------------
' Note: The error line intentionally has no line
' number to demonstrate how it effects the error
' message.
' ------------------------------------------------
    Const PROC = "Test_8_VB_Runtime_Error_TestProc_3d"
    
    On Error GoTo eh
    
    Dim l As Long
    l = 7 / 0

xt: Exit Sub

eh: Select Case mBasic.ErrMsg(ErrSrc(PROC))
        Case vbResume:  Stop: Resume
        Case Else:      GoTo xt
    End Select
End Sub

