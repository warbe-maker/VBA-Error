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

Private Sub BoP(ByVal b_proc As String, ParamArray b_arguments() As Variant)
' ------------------------------------------------------------------------------
' Common 'Begin of Procedure' interface for the 'Common VBA Error Services' and
' the 'Common VBA Execution Trace Service' (only in case the first one is not
' installed/activated). The services, when installed, are activated by the
' | Cond. Comp. Arg.             | Installed component |
' |------------------------------|---------------------|
' | XcTrc_mTrc = 1          | mTrc                |
' | XcTrc_clsTrc = 1        | clsTrc              |
' | ErHComp = 1                  | mErH                |
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
' ------------------------------------------------------------------------------
' White-box- and regression-test procedure obligatory to be performed after any
' code modification. The display of an execution trace log along with this test
' requires a conditional compile argument ExecTrace = 1.
' ------------------------------------------------------------------------------
    Const PROC = "Demo_Execution_Trace"
    
    On Error GoTo eh
    
#If XcTrc_mTrc = 1 Then
    mTrc.LogFileFullName = Replace(ThisWorkbook.FullName, ThisWorkbook.Name, "DemoExecTrace.log")
    mTrc.LogTitle = "Demo of an Execution Trace (Cond. Comp. Arg. 'ExecTraceMymTrc = 1'"
#ElseIf XcTrc_clsTrc Then
    Trc.LogFileFullName = Replace(ThisWorkbook.FullName, ThisWorkbook.Name, "DemoExecTrace.log")
    Trc.LogTitle = "Demo of an Execution Trace (Cond. Comp. Arg. 'XcTrc_clsTrc = 1'"
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

Private Sub EoP(ByVal e_proc As String, Optional ByVal e_inf As String = vbNullString)
' ------------------------------------------------------------------------------
' Common 'End of Procedure' interface for the 'Common VBA Error Services' and
' the 'Common VBA Execution Trace Service' (only in case the first one is not
' installed/activated).
' Note 1: The services, when installed, are activated by the
'         | Cond. Comp. Arg.             | Installed component |
'         |------------------------------|---------------------|
'         | XcTrc_mTrc = 1          | mTrc                |
'         | XcTrc_clsTrc = 1        | clsTrc              |
'         | ErHComp = 1                  | mErH                |
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
' ------------------------------------------------------------------------------
' Error message by Common VBA Error Services module mErH providing:
' - Error source
' - Programmed error number (1,2, 3, ... per procedure)
' - Error line              (if available)
' - Info about error        (optionally when concatenated to the error message
'                            with '||')
' - Path to the error       (provided either by a call stack or by the back
'                            trace)
' - Execution Trace         depending on whether the Common VBA Execution Trace
'                           Service by the Standard Module mTrc (Cond. Comp.
'                           Arg. 'XcTrc_mTrc = !') or by the Classs Module
'                           clsTrc (Cond. Comp. Arg. 'XcTrc_clsTrc) is
'                           installed
' - Debug/Test choice       (Cond. Comp. Arg. 'DebugAndTest= 1')
' ------------------------------------------------------------------------------
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
' ------------------------------------------------------------------------------
' Error message by Common VBA Error Services module mErH providing:
' - Error source
' - Programmed error number (1,2, 3, ... per procedure)
' - Error line              (if available)
' - Info about error        (optionally when concatenated to the error message
'                            with '||')
' - Path to the error       (provided either by a call stack or by the back
'                            trace)
' - Execution Trace         depending on whether the Common VBA Execution Trace
'                           Service by the Standard Module mTrc (Cond. Comp.
'                           Arg. 'XcTrc_mTrc = !') or by the Classs Module
'                           clsTrc (Cond. Comp. Arg. 'XcTrc_clsTrc) is
'                           installed
' - Debug/Test choice       (Cond. Comp. Arg. 'DebugAndTest= 1')
' ------------------------------------------------------------------------------
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
' ------------------------------------------------------------------------------
' Error message by Common VBA Error Services module mErH providing:
' - Error source
' - Programmed error number (1,2, 3, ... per procedure)
' - Error line              (if available)
' - Info about error        (optionally when concatenated to the error message
'                            with '||')
' - Path to the error       (provided either by a call stack or by the back
'                            trace)
' - Execution Trace         depending on whether the Common VBA Execution Trace
'                           Service by the Standard Module mTrc (Cond. Comp.
'                           Arg. 'XcTrc_mTrc = !') or by the Classs Module
'                           clsTrc (Cond. Comp. Arg. 'XcTrc_clsTrc) is
'                           installed
' - Debug/Test choice       (Cond. Comp. Arg. 'DebugAndTest= 1')
' ------------------------------------------------------------------------------
Const PROC  As String = "Demo_ErH_Elaborated"
    
    On Error GoTo eh
    BoP ErrSrc(PROC)
    Demo_ErH_Elaborated_a
    EoP ErrSrc(PROC)

xt: Exit Sub

eh: ErrMsg err_source:=ErrSrc(PROC)
End Sub

Private Function Demo_ErH_Reasonable_a(ByVal op1 As Variant, _
                                       ByVal op2 As Variant) As Variant
' ------------------------------------------------------------------------------
' Error message by Common VBA Error Services module mErH providing:
' - Error source
' - Application Error Number by AppErr procedure (1,2, 3, ... per procedure)
' - Error line               if available
' - Info about error         optionally when concatenated to the error message
'                            with '||'
' - Path to the error        not available
' - Execution Trace          not supported
' - Debug/Test choice        not supported
' ------------------------------------------------------------------------------
    Const PROC = "Demo_ErH_Reasonable_a"    ' error source

    On Error GoTo eh
    BoP ErrSrc(PROC)
46  Demo_ErH_Reasonable_b = op1 / op2
    EoP ErrSrc(PROC)
    Exit Function

eh:
    ErrMsg err_source:=ErrSrc(PROC), err_dscrptn:=Err.Description & "||Line number manually added for demonstration."
End Function

Private Property Let Demo_ErH_Reasonable_b(ByVal v As Variant)
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
    EoP ErrSrc(PROC)
    Exit Property

eh:
    ErrMsg err_source:=ErrSrc(PROC), err_dscrptn:=Err.Description & "||Line number manually added for demonstration."
End Property

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
' Services:
' - Debugging option button when the Cond. Comp. Arg.
'   'Debugging = 1'
' - Displays an optional additional "About the error:" section when a string is
'   concatenated with the error message by two vertical bars (||)
' - Invokes ErrMsg when the Cond. Comp. Arg. ErHComp = !
' - Invokes mMsg.ErrMsg when the Cond. Comp. Arg. MsgComp = ! (and
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

