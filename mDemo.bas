Attribute VB_Name = "mDemo"
Option Explicit

Public Sub ErrorHandling_None_Demo()
    Dim l As Long
    l = ErrorHandling_None(10, 0)
End Sub

Private Function ErrorHandling_None(ByVal op1 As Variant, _
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
    ErrorHandling_None = op1 / op2
End Function

Public Sub ErrorHandling_BetterThanNothing_Demo()
    ErrorHandling_BetterThanNothing 10, 0
End Sub

Private Function ErrorHandling_BetterThanNothing(ByVal op1 As Variant, _
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
Const PROC = "ErrorHandling_BetterThanNothing"    ' error source

    On Error GoTo on_error
46  ErrorHandling_BetterThanNothing = op1 / op2
    Exit Function

on_error:
    mErrHndlr.ErrMsg Err.Number, ErrSrc(PROC), Err.Description & "||Line number manually added for demonstration.", Erl
End Function

Public Sub ErrorHandling_Reasonable_Demo()
    ErrorHandling_Reasonable 10, 0
End Sub

Private Function ErrorHandling_Reasonable(ByVal op1 As Variant, _
                                          ByVal op2 As Variant) As Variant
' ------------------------------------------------------------------------
' - Error message:              Yes, well or even better formated
'   - Error source:             Yes
'   - Application error number: Supported by the function AppErr() but not used in this demo
'   - Error line:               Yes, if one is provided/available
'   - Info about error:         Possible, when attached to the error description
'                               by means of tw vertical bars (||)
'   - Path to the error:        No, because a call stack is not maintained from the
'                               entry procedure down to the error causing procedure
' - Variant value assertion:    No
' - Execution Trace:            No
' - Debug/Test choice:          No
' ---------------------------------------------------------------------
Const PROC = "ErrorHandling_Reasonable"    ' error source

    On Error GoTo on_error
    BoP ErrSrc(PROC)
46  ErrorHandling_Reasonable = op1 / op2
    EoP ErrSrc(PROC)
    Exit Function

on_error:
    mErrHndlr.ErrHndlr Err.Number, ErrSrc(PROC), Err.Description & "||Line number manually added for demonstration.", Erl
End Function

Public Sub ErrorHandling_Eleborated_Demo()
' - Error message:              Yes, well or even better formated
'   - Error source:             Yes
'   - Application error number: Supported by the function AppErr() but not used in this demo
'   - Error line:               Yes, if one is provided/available
'   - Info about error:         Possible, when attached to the error description
'                               by means of tw vertical bars (||)
'   - Path to the error:        Yes, because a call stack is maintained from the
'                               entry procedure down to the error causing procedure
'                               and the error is passed on to the calling procedure by
'                               the common ErrHndlr procedure
' - Variant value assertion:    Yes, with programmed "Application" error numbers supported
'                               by the function AppErr()
' - Execution Trace:            Yes, by the use of the common ErrHndlr procedure which
'                               automatically displays it in the immediate window when the
'                               entry procedure is reached.
' - Debug/Test choice:          Yes, demonstrated
' -----------------------------------------------------------------------------------------
Const PROC  As String = "ErrorHandling_Eleborated_Demo"
    
    On Error GoTo on_error
    BoP ErrSrc(PROC)
    ErrorHandling_Elaborated1
    EoP ErrSrc(PROC)

on_error:
    mErrHndlr.ErrHndlr Err.Number, ErrSrc(PROC), Err.Description, Erl
End Sub

Private Sub ErrorHandling_Elaborated1()
' ----------------------------------------------------------
' - Error message:                    Yes (global common module)
'   - Error source:                   Yes
'   - Programmed error number:       Yes (1,2,... per procedure)
'   - Error line:                     Yes (if available)
'   - Info about error:               Yes (optionally concatenated to the error message with '||')
'   - Path to the error (call stack): Yes
' - Execution Trace:                  Yes (with Conditional Compile Argument 'ExecTrace = !'
' - Debug/Test choice:                Yes (with Conditional Compile Argument 'DebugAndTest= 1'
' -----------------------------------------------------------------------
Const PROC  As String = "ErrorHandling_Elaborated1"

    On Error GoTo on_error
    BoP ErrSrc(PROC)    ' Push procedure on call stack
    
    ErrorHandling_Elaborated2 10, 0

exit_proc:
    EoP ErrSrc(PROC)    ' Pull procedure from call stack
    Exit Sub

on_error:
    mErrHndlr.ErrHndlr Err.Number, ErrSrc(PROC), Err.Description, Erl
End Sub

Private Function ErrorHandling_Elaborated2(ByVal op1 As Variant, _
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
Const PROC  As String = "ErrorHandling_Elaborated2"

    On Error GoTo on_error
    BoP ErrSrc(PROC)    ' Push procedure on call stack
    
    If Not IsNumeric(op1) Then Err.Raise AppErr(1), ErrSrc(PROC), "The parameter (op1) is not numeric!"
    If Not IsNumeric(op2) Then Err.Raise AppErr(2), ErrSrc(PROC), "The parameter (op2) is not numeric!"
163 If op2 = 0 Then Err.Raise AppErr(3), ErrSrc(PROC), "The parameter (op2) is 0 which would cause a 'Division by zero' error!" & DCONCAT & _
                                                "This error has been detected by a programed assertion of correct values provided for the function call." & vbLf & _
                                                "(this extra information is part of the error message but split by means of two vertical bars, which is only possible by programed (Err.Raise) error message "
    ErrorHandling_Elaborated2 = op1 / op2

exit_proc:
    EoP ErrSrc(PROC)    ' Pull procedure from call stack
    Exit Function

on_error:
    mErrHndlr.ErrHndlr Err.Number, ErrSrc(PROC), Err.Description, Erl
End Function

Public Sub Demo_2_Application_Error()
' -----------------------------------------------------------
' This test procedure obligatory after any code modification.
' The option to continue with the next test procedure (in
' case this one runs within a regression test) is only
' displayed when the Conditional Compile Argument Test = 1
' The display of an execution trace along with this test
' requires a Conditional Compile Argument ExecTrace = 1.
' ------------------------------------------------------
    
    Const PROC = "Demo_2_Application_Error"
    On Error GoTo on_error
    
    BoP ErrSrc(PROC)
    Demo_2_Application_Error_DemoProc_2a

exit_proc:
    EoP ErrSrc(PROC)
    Exit Sub

on_error:
    Select Case mErrHndlr.ErrHndlr(Err.Number, ErrSrc(PROC), Err.Description, Erl)
        Case ResumeError: Stop: Resume
    End Select
End Sub

Private Sub Demo_2_Application_Error_DemoProc_2a()

    Const PROC = "Demo_2_Application_Error_DemoProc_2a"
    On Error GoTo on_error
    
    BoP ErrSrc(PROC)
    Demo_2_Application_Error_DemoProc_2b
    EoP ErrSrc(PROC)
    Exit Sub

on_error:
    If mErrHndlr.ErrHndlr(Err.Number, ErrSrc(PROC), Err.Description, Erl) = ResumeError Then Stop: Resume
End Sub

Private Sub Demo_2_Application_Error_DemoProc_2b()
    
    Const PROC = "Demo_2_Application_Error_DemoProc_2b"
    On Error GoTo on_error
    
    BoP ErrSrc(PROC)
    Demo_2_Application_Error_DemoProc_2c
    EoP ErrSrc(PROC)
    Exit Sub

on_error:
    If mErrHndlr.ErrHndlr(Err.Number, ErrSrc(PROC), Err.Description, Erl) = ResumeError Then Stop: Resume
End Sub

Private Sub Demo_2_Application_Error_DemoProc_2c()
' ------------------------------------------------
' Note: The line number is added just for test to
' demonstrate how it effects the error message.
' ------------------------------------------------
    
    Const PROC = "Demo_2_Application_Error_DemoProc_2c"
    On Error GoTo on_error

    BoP ErrSrc(PROC)
181 Err.Raise AppErr(1), ErrSrc(PROC), _
        "This is a programmed i.e. an ""Application Error""!" & DCONCAT & _
        "The function AppErr() has been used to turn the positive into a negative number by adding the VB constant 'vbObjectError' to assure an error number which does not conflict with a VB Runtime error. " & _
        "The ErrHndlr identified the negative number as an ""Application Error"" and converted it back to the orginal positive number by means of the AppErr() function." & vbLf & _
        vbLf & _
        "Also note that this information is part of the raised error message but concatenated with two vertical bars indicating that it is an additional information regarding this error."

exit_proc:
    EoP ErrSrc(PROC)
    Exit Sub

on_error:
    Select Case mErrHndlr.ErrHndlr(Err.Number, ErrSrc(PROC), Err.Description, Erl)
        Case ResumeError:       Stop: Resume
        Case ResumeNext:        Resume Next
        Case ExitAndContinue:   GoTo exit_proc
    End Select
End Sub

Public Sub Demo_3_VB_Runtime_Error()
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
    
    Const PROC = "Demo_3_VB_Runtime_Error"
    On Error GoTo on_error
    
    BoP ErrSrc(PROC)
    Demo_3_VB_Runtime_Error_DemoProc_3a
    EoP ErrSrc(PROC)

exit_proc:
    Exit Sub

on_error:
    Select Case mErrHndlr.ErrHndlr(Err.Number, ErrSrc(PROC), Err.Description, Erl)
        Case ResumeError: Stop: Resume
        Case ResumeNext: Resume Next
        Case ExitAndContinue: GoTo exit_proc
    End Select
End Sub

Private Sub Demo_3_VB_Runtime_Error_DemoProc_3a()

    Const PROC = "Demo_3_VB_Runtime_Error_DemoProc_3a"
    On Error GoTo on_error

    BoP ErrSrc(PROC)
    Demo_3_VB_Runtime_Error_DemoProc_3b
    EoP ErrSrc(PROC)
    Exit Sub

on_error:
    If mErrHndlr.ErrHndlr(Err.Number, ErrSrc(PROC), Err.Description, Erl) = ResumeError Then Stop: Resume
End Sub

Private Sub Demo_3_VB_Runtime_Error_DemoProc_3b()
    
    Const PROC = "Demo_3_VB_Runtime_Error_DemoProc_3b"
    On Error GoTo on_error

    BoP ErrSrc(PROC)
    Demo_3_VB_Runtime_Error_DemoProc_3c
    EoP ErrSrc(PROC)
    Exit Sub

on_error:
    If mErrHndlr.ErrHndlr(Err.Number, ErrSrc(PROC), Err.Description, Erl) = ResumeError Then Stop: Resume
End Sub

Private Sub Demo_3_VB_Runtime_Error_DemoProc_3c()

    Const PROC = "Demo_3_VB_Runtime_Error_DemoProc_3c"
    On Error GoTo on_error
    
    BoP ErrSrc(PROC)
    Demo_3_VB_Runtime_Error_DemoProc_3d
    EoP ErrSrc(PROC)
    Exit Sub

on_error:
    If mErrHndlr.ErrHndlr(Err.Number, ErrSrc(PROC), Err.Description, Erl) = ResumeError Then Stop: Resume
End Sub

Private Sub Demo_3_VB_Runtime_Error_DemoProc_3d()
' ------------------------------------------------
' Note: The error line intentionally has no line
' number to demonstrate how it effects the error
' message.
' ------------------------------------------------
    
    Const PROC = "Demo_3_VB_Runtime_Error_DemoProc_3d"
    On Error GoTo on_error

    BoP ErrSrc(PROC)
    Dim l As Long
    l = 7 / 0

exit_proc:
    EoP ErrSrc(PROC)
    Exit Sub

on_error:
    Select Case mErrHndlr.ErrHndlr(Err.Number, ErrSrc(PROC), Err.Description, Erl)
        Case ResumeError:       Stop: Resume
        Case ResumeNext:        Resume Next
        Case ExitAndContinue:   GoTo exit_proc
    End Select
End Sub

Public Sub Demo_4_DebugAndDemo_with_ErrHndlr()
' -----------------------------------------
' This test the Conditional Compile
' Argument DebugAndTest = 1 is required.
' -----------------------------------------
    On Error GoTo on_error
    Const PROC = "Demo_4_DebugAndDemo_with_ErrHndlr"
      
    BoP ErrSrc(PROC)
    Demo_4_DebugAndDemo_with_ErrHndlr_DemoProc_5a
    EoP ErrSrc(PROC)
    Exit Sub

on_error:
    If mErrHndlr.ErrHndlr(Err.Number, ErrSrc(PROC), Err.Description, Erl) = ResumeError Then Stop: Resume
End Sub

Private Sub Demo_4_DebugAndDemo_with_ErrHndlr_DemoProc_5a()

    Const PROC = "Demo_5_DebugAndDemo_with_ErrHndlr_DemoProc_5a"
    On Error GoTo on_error
       
    BoP ErrSrc(PROC)
15  Debug.Print ThisWorkbook.Named
    EoP ErrSrc(PROC)
    Exit Sub
    
on_error:
    Select Case mErrHndlr.ErrHndlr(errnumber:=Err.Number, errsource:=ErrSrc(PROC), errdscrptn:=Err.Description, errline:=Erl)
        Case ResumeError: Stop: Resume ' Continue with F8 to end up at the code line which caused the error
    End Select
End Sub

Public Sub Demo_5_No_Exit_Statement()
' -----------------------------------
' Exit statement missing
' -----------------------------------

    Const PROC = "Demo_6_No_Exit_Statement"
    On Error GoTo on_error
    
on_error:
    If mErrHndlr.ErrHndlr(Err.Number, ErrSrc(PROC), Err.Description, Erl) = ResumeError Then Stop: Resume
End Sub

Public Sub Demo_6_Execution_Trace()
' ------------------------------------------------------
' White-box- and regression-test procedure obligatory
' to be performed after any code modification.
' Display of an execution trace along with this test
' requires a conditional compile argument ExecTrace = 1.
' ------------------------------------------------------
    
    Const PROC = "Demo_6_Execution_Trace"
    On Error GoTo on_error
    
    BoP ErrSrc(PROC)
    Demo_6_Execution_Trace_DemoProc_6a
    EoP ErrSrc(PROC)
    Exit Sub

on_error:
    If mErrHndlr.ErrHndlr(Err.Number, ErrSrc(PROC), Err.Description, Erl) = ResumeError Then Stop: Resume
End Sub

Private Sub Demo_6_Execution_Trace_DemoProc_6a()

    Const PROC = "Demo_6_Execution_Trace_DemoProc_6a"
    On Error GoTo on_error
    
    BoP ErrSrc(PROC)
    Demo_6_Execution_Trace_DemoProc_6b
    EoP ErrSrc(PROC)
    Exit Sub

on_error:
    If mErrHndlr.ErrHndlr(Err.Number, ErrSrc(PROC), Err.Description, Erl) = ResumeError Then Stop: Resume
End Sub

Private Sub Demo_6_Execution_Trace_DemoProc_6b()
    
    Const PROC = "Demo_6_Execution_Trace_DemoProc_6b"
    On Error GoTo on_error
    
    BoP ErrSrc(PROC)
    
    Demo_6_Execution_Trace_DemoProc_6c
    
    Dim i As Long: Dim j As Long: j = 10000000
    BoT PROC & " empty loop 1 to " & j
    For i = 1 To j
    Next i
    EoT PROC & " empty loop 1 to " & j ' !!! the string must match with the BoT statement !!!
    
    EoP ErrSrc(PROC)
    Exit Sub

on_error:
    If mErrHndlr.ErrHndlr(Err.Number, ErrSrc(PROC), Err.Description, Erl) = ResumeError Then Stop: Resume
End Sub

Private Sub Demo_6_Execution_Trace_DemoProc_6c()
    
    Const PROC = "Demo_6_Execution_Trace_DemoProc_6c"
    On Error GoTo on_error

    BoP ErrSrc(PROC)
    EoP ErrSrc(PROC)

exit_proc:
    Exit Sub

on_error:
    mErrHndlr.ErrHndlr Err.Number, ErrSrc(PROC), Err.Description, Erl
End Sub

Private Sub Demo_7_Free_Button_Display()

    On Error GoTo on_error
    Const PROC = "Demo_7_Free_Button_Display"

    Err.Raise AppErr(1), ErrSrc(PROC), "Display of a free defined button in addition to the usual Ok button (resumes the error when clicked)"
    Exit Sub

on_error:
    Select Case mErrHndlr.ErrHndlr(Err.Number, ErrSrc(PROC), Err.Description, Erl, buttons:=vbOKOnly & "," & vbLf & ",My button")
        Case "My button": Resume
    End Select
End Sub

Private Function ErrSrc(ByVal s As String) As String
' ---------------------------------------------------
' Prefix procedure name (s) by this module's name.
' Attention: The characters > and < must not be used!
' ---------------------------------------------------
    ErrSrc = "mDemo." & s
End Function
