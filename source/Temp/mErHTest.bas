Attribute VB_Name = "mErHTest"
Option Explicit
' ----------------------------------------------------------------------------
' Standard Module mErHTest
'
' Uses the following procedures for keeping the use of the Common VBA Error
' Services, the Common VBA Message Service, and the Common VBA Execution
' Trace Service optional:
' - AppErr
' - BoP
' - EoP
' - ErrMsg
'
' See: https://warbe-maker.github.io/vba/common/2022/02/15/Personal-and-public-Common-Components.html#conditional-compile-arguments
' ----------------------------------------------------------------------------

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

Private Sub BoP(ByVal b_proc As String, ParamArray b_arguments() As Variant)
' ------------------------------------------------------------------------------
' Common 'Begin of Procedure' interface for the 'Common VBA Error Services' and
' the 'Common VBA Execution Trace Service' (only in case the first one is not
' installed/activated). The services, when installed, are activated by the
' | Cond. Comp. Arg.        | Installed component |
' |-------------------------|---------------------|
' | XcTrc_mTrc = 1          | mTrc                |
' | XcTrc_clsTrc = 1        | clsTrc              |
' | ErHComp = 1             | mErH                |
' I.e. both components are independant from each other!
' Note: This procedure is obligatory for any VB-Component using either the
'       the 'Common VBA Error Services' and/or the 'Common VBA Execution Trace
'       Service'.
' ------------------------------------------------------------------------------
    Dim s As String
    If UBound(b_arguments) >= 0 Then s = Join(b_arguments, ",")

#If ErHComp = 1 Then
    '~~ The error handling also hands over to the mTrc/clsTrc component when
    '~~ either of the two is installed.
    mErH.BoP b_proc, s
#ElseIf XcTrc_clsTrc = 1 Then
    '~~ mErH is not installed but the mTrc is
    Trc.BoP b_proc, s
#ElseIf XcTrc_mTrc = 1 Then
    '~~ mErH neither mTrc is installed but clsTrc is
    mTrc.BoP b_proc, s
#End If

End Sub

Private Sub EoP(ByVal e_proc As String, Optional ByVal e_inf As String = vbNullString)
' ------------------------------------------------------------------------------
' Common 'End of Procedure' interface for the 'Common VBA Error Services' and
' the 'Common VBA Execution Trace Service' (only in case the first one is not
' installed/activated).
' Note 1: The services, when installed, are activated by the
'         | Cond. Comp. Arg.        | Installed component |
'         |-------------------------|---------------------|
'         | XcTrc_mTrc = 1          | mTrc                |
'         | XcTrc_clsTrc = 1        | clsTrc              |
'         | ErHComp = 1             | mErH                |
'         I.e. both components are independant from each other!
' Note 2: This procedure is obligatory for any VB-Component using either the
'         the 'Common VBA Error Services' and/or the 'Common VBA Execution
'         Trace Service'.
' ------------------------------------------------------------------------------
#If ErHComp = 1 Then
    '~~ The error handling also hands over to the mTrc component when 'ExecTrace = 1'
    '~~ so the Else is only for the case the mTrc is installed but the merH is not.
    mErH.EoP e_proc
#ElseIf XcTrc_clsTrc = 1 Then
    Trc.EoP e_proc, e_inf
#ElseIf XcTrc_mTrc = 1 Then
    mTrc.EoP e_proc, e_inf
#End If

End Sub

Private Function ErrMsg(ByVal err_source As String, _
               Optional ByVal err_no As Long = 0, _
               Optional ByVal err_dscrptn As String = vbNullString, _
               Optional ByVal err_line As Long = 0) As Variant
' ------------------------------------------------------------------------------
' This is the 'Universal Error Message Display' function used by (i.e. copied
' into) all (my) modules when procedures do have an error handling - which is
' the case for the most of them. This universal function already includes:
' - a 'Debugging Option' activated by the Cond. Comp. Arg.
'   'Debugging = 1'
' - an optional additional 'About the error' section displayed when the error
'   description has an extra information concatenated by two vertical bars (||).
'
' The function passes on the display:
' - to the ErrMsg function of the mErH module provided this module is installed
'   and this is indicated by the Cond. Comp. Arg. 'ErHComp = 1'
' - to the ErrMsg function of the mMsg module provided this module is installed
'   and this is indicated by the Cond. Comp. Arg. 'MsgComp = 1').
' Only when none of the two Common Components is installed the error is
' displayed by means of the VBA.MsgBox. The latter is just a fall-back option
' because the display of the error message misses some valuable features.
'
' Usage: Example with the Cond. Comp. Arg. 'Debugging = 1'
'
'        Private/Public <procedure-name>
'            Const PROC = "<procedure-name>"
'
'            On Error Goto eh
'            ....
'        xt: Exit Sub/Function/Property
'
'        eh: Select Case ErrMsg(ErrSrc(PROC))
'               Case vbResume:  Stop: Resume
'               Case Else:      GoTo xt
'            End Select
'        End Sub/Function/Property
'
'        The above may appear a lot of code lines but will be a godsend in case
'        of an error!
'
' Uses:  - For programmed application errors (Err.Raise AppErr(n), ....) the
'          function AppErr will be used which turns the positive number into a
'          negative one. The error message will regard a negative error number
'          as an 'Application Error' and will use AppErr to turn it back for
'          the message into its original positive number. Together with the
'          ErrSrc there will be no need to maintain numerous different error
'          numbers for a VB-Project.
'        - The caller provides the source of the error through the module
'          specific function ErrSrc(PROC) which adds the module name to the
'          procedure name.
'
' W. Rauschenberger Berlin, Nov 2021
' ------------------------------------------------------------------------------
#If ErHComp = 1 Then
    '~~ When the Common VBA Error Handling Component (ErH) is installed/used by in the VB-Project
    '~~ which also includes the installation of the mMsg component for the display of the error message.
    ErrMsg = mErH.ErrMsg(err_source:=err_source, err_number:=err_no, err_dscrptn:=err_dscrptn, err_line:=err_line)
    GoTo xt
#ElseIf MsgComp = 1 Then
    ErrMsg = mMsg.ErrMsg(err_source:=err_source)
    GoTo xt
#End If

    '~~ -------------------------------------------------------------------
    '~~ Neither the Common mMsg not the Commen mErH Component is installed.
    '~~ The error message is prepared for the VBA.MsgBox
    '~~ -------------------------------------------------------------------
    Dim ErrBttns    As Variant
    Dim ErrAtLine   As String
    Dim ErrDesc     As String
    Dim ErrLine     As Long
    Dim ErrNo       As Long
    Dim ErrSrc      As String
    Dim ErrText     As String
    Dim ErrTitle    As String
    Dim ErrType     As String
    Dim ErrAbout    As String
        
    '~~ Obtain error information from the Err object for any argument not provided
    If err_no = 0 Then err_no = Err.Number
    If err_line = 0 Then ErrLine = Erl
    If err_source = vbNullString Then err_source = Err.source
    If err_dscrptn = vbNullString Then err_dscrptn = Err.Description
    If err_dscrptn = vbNullString Then err_dscrptn = "--- No error description available ---"
    
    If InStr(err_dscrptn, "||") <> 0 Then
        ErrDesc = Split(err_dscrptn, "||")(0)
        ErrAbout = Split(err_dscrptn, "||")(1)
    Else
        ErrDesc = err_dscrptn
    End If
    
    '~~ Determine the type of error
    Select Case err_no
        Case Is < 0
            ErrNo = AppErr(err_no)
            ErrType = "Application Error "
        Case Else
            ErrNo = err_no
            If (InStr(1, err_dscrptn, "DAO") <> 0 _
            Or InStr(1, err_dscrptn, "ODBC Teradata Driver") <> 0 _
            Or InStr(1, err_dscrptn, "ODBC") <> 0 _
            Or InStr(1, err_dscrptn, "Oracle") <> 0) _
            Then ErrType = "Database Error " _
            Else ErrType = "VB Runtime Error "
    End Select
    
    If err_source <> vbNullString Then ErrSrc = " in: """ & err_source & """"   ' assemble ErrSrc from available information"
    If err_line <> 0 Then ErrAtLine = " at line " & err_line                    ' assemble ErrAtLine from available information
    ErrTitle = Replace(ErrType & ErrNo & ErrSrc & ErrAtLine, "  ", " ")         ' assemble ErrTitle from available information
       
    ErrText = "Error: " & vbLf & _
              ErrDesc & vbLf & vbLf & _
              "Source: " & vbLf & _
              err_source & ErrAtLine
    If ErrAbout <> vbNullString _
    Then ErrText = ErrText & vbLf & vbLf & _
                  "About: " & vbLf & _
                  ErrAbout
    
#If Debugging Then
    ErrBttns = vbYesNoCancel
    ErrText = ErrText & vbLf & vbLf & _
              "Debugging:" & vbLf & _
              "Yes    = Resume error line" & vbLf & _
              "No     = Resume Next (skip error line)" & vbLf & _
              "Cancel = Terminate"
#Else
    ErrBttns = vbCritical
#End If
    
    ErrMsg = MsgBox(Title:=ErrTitle _
                  , Prompt:=ErrText _
                  , Buttons:=ErrBttns)
xt: Exit Function

End Function

Private Function ErrSrc(ByVal s As String) As String
    ErrSrc = "mErHTest." & s
End Function

Private Function RegressionInfo() As String
' ----------------------------------------------------
' Adds s to the Err.Description as an additional info.
' ----------------------------------------------------
    RegressionInfo = Err.Description
    If Not mErH.Regression Then Exit Function
    
    If InStr(RegressionInfo, CONCAT) <> 0 _
    Then RegressionInfo = RegressionInfo & vbLf & vbLf & "Please notice that  this is a  r e g r e s s i o n  t e s t ! Click any but the ""Terminate"" button to continue with the test in case another one follows." _
    Else RegressionInfo = RegressionInfo & CONCAT & "Please notice that  this is a  r e g r e s s i o n  t e s t !  Click any but the ""Terminate"" button to continue with the test in case another one follows."

End Function

Private Sub RegressionKeepLog()
    Dim sFile As String

#If ExecTrace = 1 Then
#If MsgComp = 1 Or ErHComp = 1 Then
    '~~ avoid the error message when the Cond. Comp. Arg. 'MsgComp = 0'!
    mTrc.Dsply
#End If
    '~~ Keep the regression test result
    With New FileSystemObject
        sFile = .GetParentFolderName(mTrc.LogFile) & "\RegressionTest.log"
        If .FileExists(sFile) Then .DeleteFile (sFile)
        .GetFile(mTrc.LogFile).Name = "RegressionTest.log"
    End With
    mTrc.Terminate
#End If

End Sub

Public Sub Test_0_Regression()
' -----------------------------------------------------------------------------
' 1. This regression test requires the Cond. Comp. Arg. "Test = 1"
'    to run un-attended
' 2. The BoP/EoP statements in this regression test procedure produce one final
'    execution trace provided the Cond. Comp. Arg. "ExecTrace = 1".
' 3. Error conditions tested provide the asserted error number which bypasses
'    the display of the error message - which is documented in the execution
'    trace however. By avoiding a user action required when the error is
'    displayed allows a fully automated regression test.
' 4. In case any tests fails: The Cond. Comp. Arg. "Debugging = 1"
'    allows to identify the code line which causes the error through an extra
'    "Debug: Resume error code line" button displayed with the error message
'    and processed when clicked as "Stop: Resume" when the button is clicked.
' ------------------------------------------------------------------------------
    Const PROC = "Test_0_Regression"
    
    On Error GoTo eh
    
    '~~ Initialization of a new Trace Log File for this Regression test
    '~~ ! must be done prior the first BoP !
    
    mTrc.LogFileFullName = Replace(ThisWorkbook.FullName, ThisWorkbook.Name, "Regression Test.log")
    mTrc.LogTitle = "Regression Test module mErH"
    
    mErH.Regression = True
      
    BoP ErrSrc(PROC)
    Test_1_Application_Error
    Test_2_VB_Runtime_Error
    
xt: EoP ErrSrc(PROC)
    mErH.Regression = False
    mTrc.Dsply
    Exit Sub
    
eh: Select Case ErrMsg(err_source:=ErrSrc(PROC))
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
    BoP ErrSrc(PROC)
    
    mErH.Asserted AppErr(1)
    Test_1_Application_Error_TestProc_2a
  
xt: EoP ErrSrc(PROC)
    Exit Sub

eh: Select Case ErrMsg(ErrSrc(PROC))
        Case vbResume:  Stop: Resume
        Case Else:      GoTo xt
    End Select
End Sub

Private Sub Test_1_Application_Error_TestProc_2a()
    Const PROC = "Test_1_Application_Error_TestProc_2a"
    
    On Error GoTo eh
    
    BoP ErrSrc(PROC)
    Test_1_Application_Error_TestProc_2b
    
xt: EoP ErrSrc(PROC)
    Exit Sub

eh: Select Case ErrMsg(ErrSrc(PROC))
        Case vbResume:  Stop: Resume
        Case Else:      GoTo xt
    End Select
End Sub

Private Sub Test_1_Application_Error_TestProc_2b()
    Const PROC = "Test_1_Application_Error_TestProc_2b"
    
    On Error GoTo eh
    
    BoP ErrSrc(PROC)
    Test_1_Application_Error_TestProc_2c
    
xt: EoP ErrSrc(PROC)
    Exit Sub

eh: Select Case ErrMsg(err_source:=ErrSrc(PROC))
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

    BoP ErrSrc(PROC)
    mErH.Asserted AppErr(1)
    
181 Err.Raise AppErr(1), ErrSrc(PROC), _
        "This is a programmed i.e. an ""Application Error""!" & CONCAT & _
        "The AppErr service has been used to turn the positive into a negative number by adding " & _
        "the VB constant 'vbObjectError' to assure the error number is not confused with a VB Runtime error. " & _
        "The ErrMsg service used the AppErr service to identify the number as an 'Application Error' " & _
        "and turn the negative number back into the orginal positive number." & vbLf & _
        vbLf & _
        "By the way: Note that all the above information had been provided with the err.Description " & _
        "by concatenating it with two vertical bars indicating that it as this additional information."

xt: EoP ErrSrc(PROC)
    Exit Sub

eh: Select Case ErrMsg(err_source:=ErrSrc(PROC))
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
    
    BoP ErrSrc(PROC)
    mErH.Asserted 11
    Test_2_VB_Runtime_Error_TestProc_3a

xt: EoP ErrSrc(PROC)
    Exit Sub

eh: Select Case ErrMsg(err_source:=ErrSrc(PROC))
        Case vbResume:  Stop: Resume
        Case Else:      GoTo xt
    End Select
End Sub

Private Sub Test_2_VB_Runtime_Error_TestProc_3a()
    Const PROC = "Test_2_VB_Runtime_Error_TestProc_3a"
    
    On Error GoTo eh

    BoP ErrSrc(PROC)
    Test_2_VB_Runtime_Error_TestProc_3b
    
xt: EoP ErrSrc(PROC)
    Exit Sub

eh: Select Case ErrMsg(err_source:=ErrSrc(PROC))
        Case vbResume:  Stop: Resume
        Case Else:      GoTo xt
    End Select
End Sub

Private Sub Test_2_VB_Runtime_Error_TestProc_3b()
    Const PROC = "Test_2_VB_Runtime_Error_TestProc_3b"
    
    On Error GoTo eh

    BoP ErrSrc(PROC)
    Test_2_VB_Runtime_Error_TestProc_3c
    
xt: EoP ErrSrc(PROC)
    Exit Sub

eh: Select Case ErrMsg(err_source:=ErrSrc(PROC))
        Case vbResume:  Stop: Resume
        Case Else:      GoTo xt
    End Select
End Sub

Private Sub Test_2_VB_Runtime_Error_TestProc_3c()
    Const PROC = "Test_2_VB_Runtime_Error_TestProc_3c"
    
    On Error GoTo eh
    
    BoP ErrSrc(PROC)
    Test_2_VB_Runtime_Error_TestProc_3d "Test string", 20.5
    
xt: EoP ErrSrc(PROC)
    Exit Sub

eh: Select Case ErrMsg(err_source:=ErrSrc(PROC))
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
    
    BoP ErrSrc(PROC), "test_arg1=", test_arg1, "test_arg2=", test_arg2
    Dim l As Long
    l = 7 / 0

xt: EoP ErrSrc(PROC)
    Exit Sub

eh: Select Case ErrMsg(err_source:=ErrSrc(PROC))
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
      
    BoP ErrSrc(PROC)
    Test_4_DebugAndTest_with_ErrMsg_TestProc_5a
    
xt: EoP ErrSrc(PROC)
    Exit Sub

eh: Select Case ErrMsg(err_source:=ErrSrc(PROC))
        Case vbResume:  Stop: Resume
        Case Else:      GoTo xt
    End Select
End Sub

Private Sub Test_4_DebugAndTest_with_ErrMsg_TestProc_5a()
    Const PROC = "Test_5_DebugAndTest_with_ErrMsg_TestProc_5a"
    
    On Error GoTo eh
       
    BoP ErrSrc(PROC)
15  Debug.Print ThisWorkbook.Named
    
xt: EoP ErrSrc(PROC)
    Exit Sub
    
eh: Select Case ErrMsg(err_source:=ErrSrc(PROC))
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

eh: Select Case ErrMsg(err_source:=ErrSrc(PROC))
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
    
    BoP ErrSrc(PROC)
    mErH.Asserted AppErr(1)
    Test_6_VB_Runtime_Error_TestProc_3a

xt: EoP ErrSrc(PROC)
    Exit Sub

eh: Select Case ErrMsg(err_source:=ErrSrc(PROC))
        Case vbResume:  Stop: Resume
        Case Else:      GoTo xt
    End Select
End Sub

Private Sub Test_6_VB_Runtime_Error_TestProc_3a()
    Const PROC = "Test_6_VB_Runtime_Error_TestProc_3a"
    
    On Error GoTo eh
    Test_6_VB_Runtime_Error_TestProc_3b

xt: Exit Sub

eh: Select Case ErrMsg(err_source:=ErrSrc(PROC))
        Case vbResume:  Stop: Resume
        Case Else:      GoTo xt
    End Select
End Sub

Private Sub Test_6_VB_Runtime_Error_TestProc_3b()
    Const PROC = "Test_6_VB_Runtime_Error_TestProc_3b"
    
    On Error GoTo eh
    Test_6_VB_Runtime_Error_TestProc_3c

xt: Exit Sub

eh: Select Case ErrMsg(err_source:=ErrSrc(PROC))
        Case vbResume:  Stop: Resume
        Case Else:      GoTo xt
    End Select
End Sub

Private Sub Test_6_VB_Runtime_Error_TestProc_3c()
    Const PROC = "Test_6_VB_Runtime_Error_TestProc_3c"
    
    On Error GoTo eh
    Test_6_VB_Runtime_Error_TestProc_3d test_arg1:="Test string", test_arg2:=20.5

xt: Exit Sub

eh: Select Case ErrMsg(err_source:=ErrSrc(PROC))
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

eh: Select Case ErrMsg(err_source:=ErrSrc(PROC))
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
    
    BoP ErrSrc(PROC)
    Test_7_VB_Runtime_Error_TestProc_3a

    Debug.Assert mErH.MostRecentError = 11

xt: EoP ErrSrc(PROC)
    Exit Sub

eh: Select Case ErrMsg(err_source:=ErrSrc(PROC))
        Case vbResume:  Stop: Resume
        Case Else:      GoTo xt
    End Select
End Sub

Private Sub Test_7_VB_Runtime_Error_TestProc_3a()
    Const PROC = "Test_7_VB_Runtime_Error_TestProc_3a"
    
    On Error GoTo eh

    Test_7_VB_Runtime_Error_TestProc_3b
    
xt: Exit Sub

eh: Select Case ErrMsg(err_source:=ErrSrc(PROC))
        Case vbResume:  Stop: Resume
        Case Else:      GoTo xt
    End Select
End Sub

Private Sub Test_7_VB_Runtime_Error_TestProc_3b()
    Const PROC = "Test_7_VB_Runtime_Error_TestProc_3b"
    
    On Error GoTo eh

    Test_7_VB_Runtime_Error_TestProc_3c
    
xt: Exit Sub

eh: Select Case ErrMsg(err_source:=ErrSrc(PROC))
        Case vbResume:  Stop: Resume
        Case Else:      GoTo xt
    End Select
End Sub

Private Sub Test_7_VB_Runtime_Error_TestProc_3c()
    Const PROC = "Test_7_VB_Runtime_Error_TestProc_3c"
    
    On Error GoTo eh
    
    Test_7_VB_Runtime_Error_TestProc_3d test_arg1:="Test string", test_arg2:=20.5
    
xt: Exit Sub

eh: Select Case ErrMsg(ErrSrc(PROC))
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

eh: Select Case ErrMsg(ErrSrc(PROC))
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

eh: Select Case ErrMsg(err_source:=ErrSrc(PROC))
        Case vbResume:  Stop: Resume
        Case Else:      GoTo xt
    End Select
End Sub

Private Sub Test_8_VB_Runtime_Error_TestProc_3a()
    Const PROC = "Test_8_VB_Runtime_Error_TestProc_3a"
    
    On Error GoTo eh

    Test_8_VB_Runtime_Error_TestProc_3b
    
xt: Exit Sub

eh: Select Case ErrMsg(err_source:=ErrSrc(PROC))
        Case vbResume:  Stop: Resume
        Case Else:      GoTo xt
    End Select
End Sub

Private Sub Test_8_VB_Runtime_Error_TestProc_3b()
    Const PROC = "Test_8_VB_Runtime_Error_TestProc_3b"
    
    On Error GoTo eh

    Test_8_VB_Runtime_Error_TestProc_3c
    
xt: Exit Sub

eh: Select Case ErrMsg(err_source:=ErrSrc(PROC))
        Case vbResume:  Stop: Resume
        Case Else:      GoTo xt
    End Select
End Sub

Private Sub Test_8_VB_Runtime_Error_TestProc_3c()
    Const PROC = "Test_8_VB_Runtime_Error_TestProc_3c"
    
    On Error GoTo eh
    
    Test_8_VB_Runtime_Error_TestProc_3d test_arg1:="Test string", test_arg2:=20.5
    
xt: Exit Sub

eh: Select Case ErrMsg(ErrSrc(PROC))
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

eh: Select Case ErrMsg(ErrSrc(PROC))
        Case vbResume:  Stop: Resume
        Case Else:      GoTo xt
    End Select
End Sub

