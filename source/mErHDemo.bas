Attribute VB_Name = "mErHDemo"
Option Explicit
' ----------------------------------------------------------------------------
' Standard Module mErHDemo: Demonstrations around the Common VBA Error
' ========================= Services including examples without.
'
' Uses the following procedures for keeping the use of the Common VBA Error
' Services, the Common VBA Message Service, and the Common VBA Execution
' Trace Service optionsl:
' - BoP
' - EoP
' - ErrMsg, AppErr
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

Private Sub BoP(ByVal b_proc As String, _
       Optional ByVal b_args As String = vbNullString)
' ------------------------------------------------------------------------------
' Common 'Begin of Procedure' interface serving the 'Common VBA Error Services'
' and - if not installed/activated the 'Common VBA Execution Trace Service'.
' Obligatory copy Private for any VB-Component using the service but not having
' the mBasic common component installed.
' ------------------------------------------------------------------------------
#If ErHComp = 1 Then          ' serves the mTrc/clsTrc when installed and active
    mErH.BoP b_proc, b_args
#ElseIf XcTrc_clsTrc = 1 Then ' when only clsTrc is installed and active
    Trc.BoP b_proc, b_args
#ElseIf XcTrc_mTrc = 1 Then   ' when only mTrc is installed and activate
    mTrc.BoP b_proc, b_args
#End If
End Sub

Public Sub Demo_Application_Error()
' ------------------------------------------------------------------------------
' This test procedure is obligatory after any code modification. The option to
' continue with the next test procedure (in case this one runs within a
' regression test) is only displayed when the Cond. Comp. Arg.
' 'Test = 1'. The display of an execution trace log along with this test
' requires a Cond. Comp. Arg. 'XcTrc_mTrc = 1' or
' 'XcTrc_clsTrc = 1 depending on which one is installed/to be used.
' ------------------------------------------------------------------------------
    Const PROC = "Demo_Application_Error"
    
    On Error GoTo eh
    
    BoP ErrSrc(PROC)
    Demo_Application_Error_a

xt: EoP ErrSrc(PROC)
    Exit Sub

eh: Select Case ErrMsg(err_source:=ErrSrc(PROC))
        Case vbResume:  Stop: Resume
        Case Else:      GoTo xt
    End Select
End Sub

Private Sub Demo_Application_Error_a()
    Const PROC = "Demo_Application_Error_a"
    
    On Error GoTo eh
    
    BoP ErrSrc(PROC)
    Demo_Application_Error_b
    
xt: EoP ErrSrc(PROC)
    Exit Sub

eh: Select Case ErrMsg(ErrSrc(PROC))
        Case vbResume:  Stop: Resume
        Case Else:      GoTo xt
    End Select
End Sub

Private Sub Demo_Application_Error_b()
    Const PROC = "Demo_Application_Error_b"
    
    On Error GoTo eh
    
    BoP ErrSrc(PROC)
    Demo_Application_Error_c
    
xt: EoP ErrSrc(PROC)
    Exit Sub

eh: Select Case ErrMsg(ErrSrc(PROC))
        Case vbResume:  Stop: Resume
        Case Else:      GoTo xt
    End Select
End Sub

Private Sub Demo_Application_Error_c()
' ------------------------------------------------
' Note: The line number is added just for test to
' demonstrate how it effects the error message.
' ------------------------------------------------
    Const PROC = "Demo_Application_Error_c"
    
    On Error GoTo eh

    BoP ErrSrc(PROC)
181 Err.Raise AppErr(1), ErrSrc(PROC), _
        "This is a programmed i.e. an ""Application Error""!" & CONCAT & _
        "The function AppErr() has been used to turn the positive into a negative number by adding the VB constant 'vbObjectError' to assure an error number which does not conflict with a VB Runtime error. " & _
        "The ErrMsg identified the negative number as an ""Application Error"" and converted it back to the orginal positive number by means of the AppErr() function." & vbLf & _
        vbLf & _
        "Also note that this information is part of the raised error message but concatenated with two vertical bars indicating that it is an additional information regarding this error."

xt: EoP ErrSrc(PROC)
    Exit Sub

eh: Select Case ErrMsg(ErrSrc(PROC))
        Case vbResume:  Stop: Resume
        Case Else:      GoTo xt
    End Select
End Sub

Private Sub Demo_ErH_BetterThanNothing()
' ----------------------------------------------------------------------------
' See: https://github.com/warbe-maker/VBA-Error#exploring-the-matter
' ----------------------------------------------------------------------------
    Demo_ErH_BetterThanNothing_a 10, 0
    MsgBox "Error ignored and thus continued!"
End Sub

Private Sub Demo_ErH_BetterThanNothing_a(ByVal d_a As Long, _
                                         ByVal d_b As Long)
' ----------------------------------------------------------------------------
' See: https://github.com/warbe-maker/VBA-Error#exploring-the-matter
' ----------------------------------------------------------------------------
    On Error GoTo eh
    
    Debug.Assert d_a / d_b

xt: Exit Sub

eh: Select Case MsgBox(Title:="An error occoured!" _
                    , Prompt:="Error " & Err.Number & ": " & Err.Description & vbLf & vbLf & _
                              "Retry  = Proceed to the error line option" & vbLf & _
                              "Ignore = Proceed to the end of the error causing procdure." & vbLf & _
                              "Abort  = No action" _
                    , Buttons:=vbAbortRetryIgnore)
        Case vbRetry:   Stop: Resume
        Case vbIgnore
        Case vbAbort: GoTo xt
    End Select
End Sub

Public Sub Demo_01_NoErrorHandling()
    Dim l As Long
    l = Demo_01_NoErrorHandling_a(10, 0)
End Sub

Private Function Demo_01_NoErrorHandling_a(ByVal op1 As Variant, _
                                           ByVal op2 As Variant) As Variant
' ----------------------------------------------------------------------------
' - Error message:              Mere VBA only
'   - Error source:             No
'   - Application error number: Not supported
'   - Error line:               No, even when one is provided/available
'   - Info about error:         Not supported
'   - Path to the error:        No, because a call stack is not maintained
' - Variant value assertion:    No
' - Execution Trace:            No
' - Debugging/Test choice:      No
' ----------------------------------------------------------------------------
    Demo_01_NoErrorHandling_a = op1 / op2
End Function

Public Sub Demo_02_Elaborated()
' ------------------------------------------------------------------------------
'
' ------------------------------------------------------------------------------
    Const PROC = "Demo_02_Elaborated"
    On Error GoTo eh
    
    BoP ErrSrc(PROC)
    Demo_02_Elaborated_a
    EoP ErrSrc(PROC)

xt: Exit Sub

eh: Select Case ErrMsg(ErrSrc(PROC))
        Case vbResume:  Stop: Resume
        Case Else:      GoTo xt
    End Select
End Sub

Public Sub Demo_02_Elaborated_a()
' ------------------------------------------------------------------------------
'
' ------------------------------------------------------------------------------
    Const PROC = "Demo_02_Elaborated_a"
    On Error GoTo eh
    
    BoP ErrSrc(PROC)
    Demo_02_Elaborated_b 10, 0
    EoP ErrSrc(PROC)

xt: Exit Sub

eh: Select Case ErrMsg(ErrSrc(PROC))
        Case vbResume:  Stop: Resume
        Case Else:      GoTo xt
    End Select
End Sub

Private Function Demo_02_Elaborated_b(ByVal op1 As Variant, _
                                       ByVal op2 As Variant) As Variant
' ------------------------------------------------------------------------------
'
' ------------------------------------------------------------------------------
    Const PROC = "Demo_02_Elaborated_a"    ' error source

    On Error GoTo eh
    BoP ErrSrc(PROC)
46  Demo_02_Elaborated_b = op1 / op2
    EoP ErrSrc(PROC)

xt: Exit Function

eh: Select Case ErrMsg(ErrSrc(PROC))
        Case vbResume:  Stop: Resume
        Case Else:      GoTo xt
    End Select
End Function

Public Sub Demo_Execution_Trace()
' ------------------------------------------------------------------------------
' White-box- and regression-test procedure obligatory to be performed after any
' code modification. The display of an execution trace log along with this test
' requires a conditional compile argument ExecTrace = 1.
' ------------------------------------------------------------------------------
    Const PROC = "Demo_Execution_Trace"
    
    On Error GoTo eh
    
#If XcTrc_mTrc = 1 Then
    mTrc.FileFullName = Replace(ThisWorkbook.FullName, ThisWorkbook.Name, "DemoExecTrace.log")
    mTrc.Title = "Demo of an Execution Trace (Cond. Comp. Arg. 'ExecTraceMymTrc = 1'"
#ElseIf XcTrc_clsTrc Then
    Trc.FileFullName = Replace(ThisWorkbook.FullName, ThisWorkbook.Name, "DemoExecTrace.log")
    Trc.Title = "Demo of an Execution Trace (Cond. Comp. Arg. 'XcTrc_clsTrc = 1'"
#End If

    BoP ErrSrc(PROC)
    Demo_Execution_Trace_a
    
xt: EoP ErrSrc(PROC)
    mTrc.Dsply
    Exit Sub

eh: Select Case ErrMsg(ErrSrc(PROC))
        Case vbResume:  Stop: Resume
        Case Else:      GoTo xt
    End Select
End Sub

Private Sub Demo_Execution_Trace_a()
    Const PROC = "Demo_Execution_Trace_a"
    
    On Error GoTo eh
    
    BoP ErrSrc(PROC)
    Demo_Execution_Trace_b
    
xt: EoP ErrSrc(PROC)
    Exit Sub

eh: Select Case ErrMsg(ErrSrc(PROC))
        Case vbResume:  Stop: Resume
        Case Else:      GoTo xt
    End Select
End Sub

Private Sub Demo_Execution_Trace_b()
    Const PROC = "Demo_Execution_Trace_b"
    
    On Error GoTo eh
    
    BoP ErrSrc(PROC)
    
    Demo_Execution_Trace_c
    
    Dim i As Long: Dim j As Long: j = 10000000
    mTrc.BoC PROC & " empty loop 1 to " & j
    For i = 1 To j
    Next i
    mTrc.EoC PROC & " empty loop 1 to " & j ' !!! the string must match with the BoC statement !!!
    
xt: EoP ErrSrc(PROC)
    Exit Sub

eh: Select Case ErrMsg(ErrSrc(PROC))
        Case vbResume:  Stop: Resume
        Case Else:      GoTo xt
    End Select
End Sub

Private Sub Demo_Execution_Trace_c()
    Const PROC = "Demo_Execution_Trace_c"
    
    On Error GoTo eh

    BoP ErrSrc(PROC)

xt: EoP ErrSrc(PROC)
    Exit Sub

eh: Select Case ErrMsg(ErrSrc(PROC))
        Case vbResume:  Stop: Resume
        Case Else:      GoTo xt
    End Select
End Sub

Public Sub Demo_No_Exit_Statement()
' -----------------------------------
' Exit statement missing
' -----------------------------------
    Const PROC = "Demo_No_Exit_Statement"
    
    On Error GoTo eh
    
eh: Select Case ErrMsg(ErrSrc(PROC))
        Case vbResume:  Stop: Resume
    End Select
End Sub

Public Sub Demo_VB_Runtime_Error()
' ------------------------------------------------------------------------------
' - With Cond. Comp. Arg. BopEop = 0:
'   Display of the error with the error path only
' - With Cond. Comp. Arg. BopEop = 1:
'   Display of the error with the error path plus
'   Display of a full execution trace
'
' Requires:
' - Cond. Comp. Arg. ExecTrace = 1.
' ------------------------------------------------------------------------------
    Const PROC = "Demo_VB_Runtime_Error"
    
    On Error GoTo eh
    
    BoP ErrSrc(PROC)
    Demo_VB_Runtime_Error_a

xt: EoP ErrSrc(PROC)
    Exit Sub

eh: Select Case ErrMsg(ErrSrc(PROC))
        Case vbResume:  Stop: Resume
        Case Else:      GoTo xt
    End Select
End Sub

Private Sub Demo_VB_Runtime_Error_a()
    Const PROC = "Demo_VB_Runtime_Error_a"
    
    On Error GoTo eh

    BoP ErrSrc(PROC)
    Demo_VB_Runtime_Error_b
    
xt: EoP ErrSrc(PROC)
    Exit Sub

eh: Select Case ErrMsg(ErrSrc(PROC))
        Case vbResume:  Stop: Resume
        Case Else:      GoTo xt
    End Select
End Sub

Private Sub Demo_VB_Runtime_Error_b()
    Const PROC = "Demo_VB_Runtime_Error_b"
    
    On Error GoTo eh

    BoP ErrSrc(PROC)
    Demo_VB_Runtime_Error_c
    
xt: EoP ErrSrc(PROC)
    Exit Sub

eh: Select Case ErrMsg(ErrSrc(PROC))
        Case vbResume:  Stop: Resume
        Case Else:      GoTo xt
    End Select
End Sub

Private Sub Demo_VB_Runtime_Error_c()
    Const PROC = "Demo_VB_Runtime_Error_c"
    
    On Error GoTo eh
    
    BoP ErrSrc(PROC)
    Demo_VB_Runtime_Error_d
    
xt: EoP ErrSrc(PROC)
    Exit Sub

eh: Select Case ErrMsg(ErrSrc(PROC))
        Case vbResume:  Stop: Resume
        Case Else:      GoTo xt
    End Select
End Sub

Private Sub Demo_VB_Runtime_Error_d()
' ------------------------------------------------
' Note: The error line intentionally has no line
' number to demonstrate how it effects the error
' message.
' ------------------------------------------------
    Const PROC = "Demo_VB_Runtime_Error_d"
    
    On Error GoTo eh

    BoP ErrSrc(PROC)
    Dim l As Long
    l = 7 / 0

xt: EoP ErrSrc(PROC)
    Exit Sub

eh: Select Case ErrMsg(ErrSrc(PROC))
        Case vbResume:  Stop: Resume
        Case Else:      GoTo xt
    End Select
End Sub

Private Sub EoP(ByVal e_proc As String, _
       Optional ByVal e_args As String = vbNullString)
' ------------------------------------------------------------------------------
' Common 'Begin of Procedure' interface serving the 'Common VBA Error Services'
' and - if not installed/activated the 'Common VBA Execution Trace Service'.
' Obligatory copy Private for any VB-Component using the service but not having
' the mBasic common component installed.
' ------------------------------------------------------------------------------
#If ErHComp = 1 Then          ' serves the mTrc/clsTrc when installed and active
    mErH.EoP e_proc, e_args
#ElseIf XcTrc_clsTrc = 1 Then ' when only clsTrc is installed and active
    Trc.EoP e_proc, e_args
#ElseIf XcTrc_mTrc = 1 Then   ' when only mTrc is installed and activate
    mTrc.EoP e_proc, e_args
#End If
End Sub

Private Function ErrMsg(ByVal err_source As String, _
               Optional ByVal err_no As Long = 0, _
               Optional ByVal err_dscrptn As String = vbNullString, _
               Optional ByVal err_line As Long = 0) As Variant
' ------------------------------------------------------------------------------
' Universal error message display service. Obligatory copy Private for any
' VB-Component using the common error service but not having the mBasic common
' component installed.
' Displays: - a debugging option button when the Cond. Comp. Arg. 'Debugging = 1'
'           - an optional additional "About:" section when the err_dscrptn has
'             an additional string concatenated by two vertical bars (||)
'           - the error message by means of the Common VBA Message Service
'             (fMsg/mMsg) when installed and active (Cond. Comp. Arg.
'             `MsgComp = 1`)
'
' Uses: AppErr  For programmed application errors (Err.Raise AppErr(n), ....)
'               to turn them into a negative and in the error message back into
'               its origin positive number.
'
' W. Rauschenberger Berlin, June 2023
' See: https://github.com/warbe-maker/VBA-Error
' ------------------------------------------------------------------------------
#If ErHComp = 1 Then
    '~~ When Common VBA Error Services (mErH) is availabel in the VB-Project
    '~~ (which includes the mMsg component) the mErh.ErrMsg service is invoked.
    ErrMsg = mErH.ErrMsg(err_source, err_no, err_dscrptn, err_line): GoTo xt
    GoTo xt
#ElseIf MsgComp = 1 Then
    '~~ When (only) the Common Message Service (mMsg, fMsg) is available in the
    '~~ VB-Project, mMsg.ErrMsg is invoked for the display of the error message.
    ErrMsg = mMsg.ErrMsg(err_source, err_no, err_dscrptn, err_line): GoTo xt
    GoTo xt
#End If
    '~~ When neither of the Common Component is available in the VB-Project
    '~~ the error message is displayed by means of the VBA.MsgBox
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
    
    '~~ Consider extra information is provided with the error description
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
            If err_dscrptn Like "*DAO*" _
            Or err_dscrptn Like "*ODBC*" _
            Or err_dscrptn Like "*Oracle*" _
            Then ErrType = "Database Error " _
            Else ErrType = "VB Runtime Error "
    End Select
    
    If err_source <> vbNullString Then ErrSrc = " in: """ & err_source & """"   ' assemble ErrSrc from available information"
    If err_line <> 0 Then ErrAtLine = " at line " & err_line                    ' assemble ErrAtLine from available information
    ErrTitle = Replace(ErrType & ErrNo & ErrSrc & ErrAtLine, "  ", " ")         ' assemble ErrTitle from available information
       
    ErrText = "Error: " & vbLf & ErrDesc & vbLf & vbLf & "Source: " & vbLf & err_source & ErrAtLine
    If ErrAbout <> vbNullString Then ErrText = ErrText & vbLf & vbLf & "About: " & vbLf & ErrAbout
    
#If Debugging = 1 Then
    ErrBttns = vbYesNo
    ErrText = ErrText & vbLf & vbLf & "Debugging:" & vbLf & "Yes    = Resume Error Line" & vbLf & "No     = Terminate"
#Else
    ErrBttns = vbCritical
#End If
    ErrMsg = MsgBox(Title:=ErrTitle, Prompt:=ErrText, Buttons:=ErrBttns)
xt:
End Function

Private Function ErrSrc(ByVal s As String) As String
' ---------------------------------------------------
' Prefix procedure name (s) by this module's name.
' Attention: The characters > and < must not be used!
' ---------------------------------------------------
    ErrSrc = "mErHDemo." & s
End Function

