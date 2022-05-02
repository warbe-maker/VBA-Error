Attribute VB_Name = "mErHDemo"
Option Explicit
' ----------------------------------------------------------------------------
' Standard Module mErHDemo
' Demonstrations around the Common VBA Error Services including examples
' without.
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

Public Sub BoP(ByVal b_proc As String, ParamArray b_arguments() As Variant)
' ------------------------------------------------------------------------------
' (B)egin-(o)f-(P)rocedure named (b_proc). Procedure to be copied as Private
' into any module potentially either using the Common VBA Error Service and/or
' the Common VBA Execution Trace Service. Has no effect when Conditional Compile
' Arguments are 0 or not set at all.
' ------------------------------------------------------------------------------
    Dim s As String: If UBound(b_arguments) >= 0 Then s = Join(b_arguments, ",")
#If ErHComp = 1 Then
    mErH.BoP b_proc, s
#ElseIf ExecTrace = 1 Then
    mTrc.BoP b_proc, s
#End If
End Sub

Public Sub Demo_Application_Error()
' -----------------------------------------------------------
' This test procedure obligatory after any code modification.
' The option to continue with the next test procedure (in
' case this one runs within a regression test) is only
' displayed when the Conditional Compile Argument Test = 1
' The display of an execution trace along with this test
' requires a Conditional Compile Argument ExecTrace = 1.
' ------------------------------------------------------
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

Public Sub Demo_VB_Runtime_Error()
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

Public Sub Demo_Execution_Trace()
' ------------------------------------------------------
' White-box- and regression-test procedure obligatory
' to be performed after any code modification.
' Display of an execution trace along with this test
' requires a conditional compile argument ExecTrace = 1.
' ------------------------------------------------------
    Const PROC = "Demo_Execution_Trace"
    
    On Error GoTo eh
    mTrc.LogFile = Replace(ThisWorkbook.FullName, ThisWorkbook.Name, "DemoExecTrace.log")
    mTrc.LogTitle = "Demo of an Execution Trace (Cond. Comp. Arg. 'ExecTrace = 1'"
    
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

Public Sub Demo_ErH_NoErrorHandling()
    Dim l As Long
    l = Demo_ErH_NoErrorHandling_a(10, 0)
End Sub

Private Function Demo_ErH_NoErrorHandling_a( _
           ByVal op1 As Variant, _
           ByVal op2 As Variant) As Variant
' ------------------------------------------------------------------
' - Error message:              Mere VBA only
'   - Error source:             No
'   - Application error number: Not supported
'   - Error line:               No, even when one is provided/available
'   - Info about error:         Not supported
'   - Path to the error:        No, because a call stack is not maintained
' - Variant value assertion:    No
' - Execution Trace:            No
' - Debugging/Test choice:      No
' ------------------------------------------------------------------
    Demo_ErH_NoErrorHandling_a = op1 / op2
End Function

Public Sub EoP(ByVal e_proc As String, _
      Optional ByVal e_inf As String = vbNullString)
' ------------------------------------------------------------------------------
' (E)nd-(o)f-(P)rocedure named (e_proc). Procedure to be copied as Private Sub
' into any module potentially either using the Common VBA Error Service and/or
' the Common VBA Execution Trace Service. Has no effect when Conditional Compile
' Arguments are 0 or not set at all.
' ------------------------------------------------------------------------------
#If ErHComp = 1 Then
    mErH.EoP e_proc
#ElseIf ExecTrace = 1 Then
    mTrc.EoP e_proc, e_inf
#End If
End Sub

Private Function Demo_ErH_BetterThanNothing_a(ByVal op1 As Variant, _
                                              ByVal op2 As Variant) As Variant
' ----------------------------------------------------------------------
' - Error message:              Yes, well or even better formated
'   - Error source:             Yes
'   - Application error number: Supported by the function AppErr() but not used in this demo
'   - Error line:               Yes, if one is provided/available
'   - Info about error:         Possible, when attached to the error description
'                               by means of tw vertical bars (||)
'   - Path to the error:        No, because a call stack is not maintained
' - Variant value assertion:    No
' - Execution Trace:            No
' - Debug/Test choice:          No
' ---------------------------------------------------------------------
    Const PROC = "Demo_ErH_BetterThanNothing"    ' error source

    On Error GoTo eh
    
46  Demo_ErH_BetterThanNothing_a = op1 / op2
    Exit Function

eh:
    MsgBox Prompt:="Error description" & vbLf & _
                    Err.Description, _
           Buttons:=vbOKOnly, _
           Title:="VB Runtime error " & Err.Number & " in " & ErrSrc(PROC) & IIf(Erl <> 0, " at line " & Erl, "")
End Function

Public Sub Demo_ErH_BetterThanNothing()
    Demo_ErH_BetterThanNothing_a 10, 0
End Sub

Private Sub Demo_ErH_Elaborated_a()
' ----------------------------------------------------------
' - Error message:                    Yes (global common module)
'   - Error source:                   Yes
'   - Programmed error number:        Yes (1,2,... per procedure)
'   - Error line:                     Yes (if available)
'   - Info about error:               Yes (optionally concatenated to the error message with '||')
'   - Path to the error (call stack): Yes
' - Execution Trace:                  Yes (with Conditional Compile Argument 'ExecTrace = !'
' - Debug/Test choice:                Yes (with Conditional Compile Argument 'DebugAndTest= 1'
' -----------------------------------------------------------------------
Const PROC As String = "Demo_ErH_Elaborated_a"

    On Error GoTo eh
    BoP ErrSrc(PROC)    ' Push procedure on call stack
    
    Demo_ErH_Elaborated_b 10, 0

xt: EoP ErrSrc(PROC)    ' Pull procedure from call stack
    Exit Sub

eh: ErrMsg err_source:=ErrSrc(PROC)
End Sub

Private Function Demo_ErH_Elaborated_b(ByVal op1 As Variant, _
                                      ByVal op2 As Variant) As Variant
' -------------------------------------------------------------------------
' - Error message:                    Yes (global common module)
'   - Error source:                   Yes
'   - Programmed error number:        Yes, the function AppErr() ensures non VB conflicting
'                                          application error numbers 1 to n per procedure
'   - Error line:                     Yes (if available)
'   - Info on error:                  Yes (optionally concatenated to the error message with '||'
'   - Path to the error (call stack): Yes
' - Variant value assertion:          Yes
' - Execution Trace:                  Yes (with Conditional Compile Argument 'ExecTrace = !'
' - Debug/Test choice:                Yes (with Conditional Compile Argument 'DebugAndTest = 1'
' ---------------------------------------------------------------------------------------
Const PROC  As String = "Demo_ErH_Elaborated_b"

    On Error GoTo eh
    BoP ErrSrc(PROC)    ' Push procedure on call stack
    
    If Not IsNumeric(op1) _
    Then Err.Raise AppErr(1), ErrSrc(PROC), "The parameter (op1) is not numeric!"
    If Not IsNumeric(op2) _
    Then Err.Raise AppErr(2), ErrSrc(PROC), "The parameter (op2) is not numeric!"

163 If op2 = 0 Then Err.Raise AppErr(3), ErrSrc(PROC), "The parameter (op2) is 0 which would cause a 'Division by zero' error!" & CONCAT & _
                                                "This error has been detected by a programed assertion of correct values provided for the function call." & vbLf & _
                                                "(this extra information is part of the error message but split by means of two vertical bars, which is only possible by programed (Err.Raise) error message "
    Demo_ErH_Elaborated_b = op1 / op2

xt: EoP ErrSrc(PROC)    ' Pull procedure from call stack
    Exit Function

eh: ErrMsg err_source:=ErrSrc(PROC)
End Function

Public Sub Demo_ErH_Elaborated()
' - Error message:              Yes, well or even better formated
'   - Error source:             Yes
'   - Application error number: Supported by the function AppErr() but not used in this demo
'   - Error line:               Yes, if one is provided/available
'   - Info about error:         Possible, when attached to the error description
'                               by means of tw vertical bars (||)
'   - Path to the error:        Yes, because a call stack is maintained from the
'                               'Entry-Procedure' down to the error causing procedure
'                               and the error is passed on to the calling procedure by
'                               the common ErrMsg procedure
' - Variant value assertion:    Yes, with programmed "Application" error numbers supported
'                               by the function AppErr()
' - Execution Trace:            Yes, by the use of the common ErrMsg procedure which
'                               automatically displays it in the immediate window when the
'                               'Entry-Procedure' is reached.
' - Debug/Test choice:          Yes, demonstrated
' -----------------------------------------------------------------------------------------
Const PROC  As String = "Demo_ErH_Elaborated"
    
    On Error GoTo eh
    BoP ErrSrc(PROC)
    Demo_ErH_Elaborated
    EoP ErrSrc(PROC)

xt: Exit Sub

eh: ErrMsg err_source:=ErrSrc(PROC)
End Sub

Private Function Demo_ErH_Reasonable_a(ByVal op1 As Variant, _
                                          ByVal op2 As Variant) As Variant
' ------------------------------------------------------------------------
' - Error message:              Yes, well or even better formated
'   - Error source:             Yes
'   - Application error number: Supported by the function AppErr() but not used in this demo
'   - Error line:               Yes, if one is provided/available
'   - Info about error:         Possible, when attached to the error description
'                               by means of tw vertical bars (||)
'   - Path to the error:        No, because a call stack is not maintained from the
'                               'Entry-Procedure' down to the error causing procedure
' - Variant value assertion:    No
' - Execution Trace:            No
' - Debug/Test choice:          No
' ---------------------------------------------------------------------
    Const PROC = "Demo_ErH_Reasonable_a"    ' error source

    On Error GoTo eh
    BoP ErrSrc(PROC)
46  Demo_ErH_Reasonable_a = op1 / op2
    EoP ErrSrc(PROC)
    Exit Function

eh:
    ErrMsg err_source:=ErrSrc(PROC), err_dscrptn:=Err.Description & "||Line number manually added for demonstration."
End Function

Public Sub Demo_ErH_Reasonable()
    Demo_ErH_Reasonable_a 10, 0
End Sub

Private Function ErrMsg(ByVal err_source As String, _
               Optional ByVal err_no As Long = 0, _
               Optional ByVal err_dscrptn As String = vbNullString, _
               Optional ByVal err_line As Long = 0) As Variant
' ------------------------------------------------------------------------------
' Universal error message display service. See:
' https://warbe-maker.github.io/vba/common/2022/02/15/Personal-and-public-Common-Components.html
'
' - Displays a debugging option button when the Conditional Compile Argument
'   'Debugging = 1'
' - Displays an optional additional "About the error:" section when a string is
'   concatenated with the error message by two vertical bars (||)
' - Invokes ErrMsg when the Conditional Compile Argument ErHComp = !
' - Invokes mMsg.ErrMsg when the Conditional Compile Argument MsgComp = ! (and
'   the mErH module is not installed / MsgComp not set)
' - Displays the error message by means of VBA.MsgBox when neither of the two
'   components is installed
'
' Uses:
' - AppErr For programmed application errors (Err.Raise AppErr(n), ....) to
'          turn them into negative and in the error message back into a
'          positive number.
' - ErrSrc To provide an unambiguous procedure name by prefixing is with the
'          module name.
'
' See:
' https://github.com/warbe-maker/Common-VBA-Error-Services
'
' W. Rauschenberger Berlin, Feb 2022
' ------------------------------------------------------------------------------' ------------------------------------------------------------------------------
#If ErHComp = 1 Then
    '~~ When Common VBA Error Services (mErH) is availabel in the VB-Project
    '~~ (which includes the mMsg component) the mErh.ErrMsg service is invoked.
    ErrMsg = ErrMsg(err_source, err_no, err_dscrptn, err_line): GoTo xt
#ElseIf MsgComp = 1 Then
    '~~ When (only) the Common Message Service (mMsg, fMsg) is available in the
    '~~ VB-Project, mMsg.ErrMsg is invoked for the display of the error message.
    ErrMsg = mMsg.ErrMsg(err_source, err_no, err_dscrptn, err_line): GoTo xt
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
    
#If Debugging Then
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

