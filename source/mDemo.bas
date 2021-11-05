Attribute VB_Name = "mDemo"
Option Explicit

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

Public Sub Demo_NoErrorHandling()
    Dim l As Long
    l = Demo_NoErrorHandling1(10, 0)
End Sub

Private Function Demo_NoErrorHandling1( _
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
    Demo_NoErrorHandling1 = op1 / op2
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

    On Error GoTo eh
    
46  ErrorHandling_BetterThanNothing = op1 / op2
    Exit Function

eh:
    MsgBox Prompt:="Error description" & vbLf & _
                    Err.Description, _
           Buttons:=vbOKOnly, _
           Title:="VB Runtime error " & Err.Number & " in " & ErrSrc(PROC) & IIf(Erl <> 0, " at line " & Erl, "")
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
'                               'Entry-Procedure' down to the error causing procedure
' - Variant value assertion:    No
' - Execution Trace:            No
' - Debug/Test choice:          No
' ---------------------------------------------------------------------
    Const PROC = "ErrorHandling_Reasonable"    ' error source

    On Error GoTo eh
    mErH.BoP ErrSrc(PROC)
46  ErrorHandling_Reasonable = op1 / op2
    mErH.EoP ErrSrc(PROC)
    Exit Function

eh:
    mErH.ErrMsg err_source:=ErrSrc(PROC), err_dscrptn:=Err.Description & "||Line number manually added for demonstration."
End Function

Public Sub ErrorHandling_Eleborated_Demo()
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
Const PROC  As String = "ErrorHandling_Eleborated_Demo"
    
    On Error GoTo eh
    mErH.BoP ErrSrc(PROC)
    ErrorHandling_Elaborated1
    mErH.EoP ErrSrc(PROC)

eh:
    mErH.ErrMsg err_source:=ErrSrc(PROC)
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
Const PROC As String = "ErrorHandling_Elaborated1"

    On Error GoTo eh
    mErH.BoP ErrSrc(PROC)    ' Push procedure on call stack
    
    ErrorHandling_Elaborated2 10, 0

xt: mErH.EoP ErrSrc(PROC)    ' Pull procedure from call stack
    Exit Sub

eh: mErH.ErrMsg err_source:=ErrSrc(PROC)
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

    On Error GoTo eh
    mErH.BoP ErrSrc(PROC)    ' Push procedure on call stack
    
    If Not IsNumeric(op1) Then Err.Raise AppErr(1), ErrSrc(PROC), "The parameter (op1) is not numeric!"
    If Not IsNumeric(op2) Then Err.Raise AppErr(2), ErrSrc(PROC), "The parameter (op2) is not numeric!"
163 If op2 = 0 Then Err.Raise AppErr(3), ErrSrc(PROC), "The parameter (op2) is 0 which would cause a 'Division by zero' error!" & CONCAT & _
                                                "This error has been detected by a programed assertion of correct values provided for the function call." & vbLf & _
                                                "(this extra information is part of the error message but split by means of two vertical bars, which is only possible by programed (Err.Raise) error message "
    ErrorHandling_Elaborated2 = op1 / op2

xt: mErH.EoP ErrSrc(PROC)    ' Pull procedure from call stack
    Exit Function

eh: mErH.ErrMsg err_source:=ErrSrc(PROC)
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
    
    On Error GoTo eh
    
    mErH.BoP ErrSrc(PROC)
    Demo_2_Application_Error_DemoProc_2a

xt: mErH.EoP ErrSrc(PROC)
    Exit Sub

eh: Select Case mErH.ErrMsg(err_source:=ErrSrc(PROC))
        Case DebugOptResumeErrorLine: Stop: Resume
    End Select
End Sub

Private Sub Demo_2_Application_Error_DemoProc_2a()
    Const PROC = "Demo_2_Application_Error_DemoProc_2a"
    
    On Error GoTo eh
    
    mErH.BoP ErrSrc(PROC)
    Demo_2_Application_Error_DemoProc_2b
    mErH.EoP ErrSrc(PROC)
    Exit Sub

eh:
    If mErH.ErrMsg(err_source:=ErrSrc(PROC)) = DebugOptResumeErrorLine Then Stop: Resume
End Sub

Private Sub Demo_2_Application_Error_DemoProc_2b()
    Const PROC = "Demo_2_Application_Error_DemoProc_2b"
    
    On Error GoTo eh
    
    mErH.BoP ErrSrc(PROC)
    Demo_2_Application_Error_DemoProc_2c
    mErH.EoP ErrSrc(PROC)
    Exit Sub

eh:
    If mErH.ErrMsg(err_source:=ErrSrc(PROC)) = DebugOptResumeErrorLine Then Stop: Resume
End Sub

Private Sub Demo_2_Application_Error_DemoProc_2c()
' ------------------------------------------------
' Note: The line number is added just for test to
' demonstrate how it effects the error message.
' ------------------------------------------------
    Const PROC = "Demo_2_Application_Error_DemoProc_2c"
    
    On Error GoTo eh

    mErH.BoP ErrSrc(PROC)
181 Err.Raise AppErr(1), ErrSrc(PROC), _
        "This is a programmed i.e. an ""Application Error""!" & CONCAT & _
        "The function AppErr() has been used to turn the positive into a negative number by adding the VB constant 'vbObjectError' to assure an error number which does not conflict with a VB Runtime error. " & _
        "The ErrMsg identified the negative number as an ""Application Error"" and converted it back to the orginal positive number by means of the AppErr() function." & vbLf & _
        vbLf & _
        "Also note that this information is part of the raised error message but concatenated with two vertical bars indicating that it is an additional information regarding this error."

xt: mErH.EoP ErrSrc(PROC)
    Exit Sub

eh: Select Case mErH.ErrMsg(err_source:=ErrSrc(PROC))
        Case DebugOptResumeErrorLine:       Stop: Resume
        Case DebugOptResumeNext:        Resume Next
        Case DebugOptCleanExit:   GoTo xt
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
    
    On Error GoTo eh
    
    mErH.BoP ErrSrc(PROC)
    Demo_3_VB_Runtime_Error_DemoProc_3a

xt: mErH.EoP ErrSrc(PROC)
    Exit Sub

eh: Select Case mErH.ErrMsg(err_source:=ErrSrc(PROC))
        Case DebugOptResumeErrorLine: Stop: Resume
        Case DebugOptResumeNext: Resume Next
        Case DebugOptCleanExit: GoTo xt
    End Select
End Sub

Private Sub Demo_3_VB_Runtime_Error_DemoProc_3a()
    Const PROC = "Demo_3_VB_Runtime_Error_DemoProc_3a"
    
    On Error GoTo eh

    mErH.BoP ErrSrc(PROC)
    Demo_3_VB_Runtime_Error_DemoProc_3b
    mErH.EoP ErrSrc(PROC)
    Exit Sub

eh:
    If mErH.ErrMsg(err_source:=ErrSrc(PROC)) = DebugOptResumeErrorLine Then Stop: Resume
End Sub

Private Sub Demo_3_VB_Runtime_Error_DemoProc_3b()
    Const PROC = "Demo_3_VB_Runtime_Error_DemoProc_3b"
    
    On Error GoTo eh

    mErH.BoP ErrSrc(PROC)
    Demo_3_VB_Runtime_Error_DemoProc_3c
    mErH.EoP ErrSrc(PROC)
    Exit Sub

eh:
    If mErH.ErrMsg(err_source:=ErrSrc(PROC)) = DebugOptResumeErrorLine Then Stop: Resume
End Sub

Private Sub Demo_3_VB_Runtime_Error_DemoProc_3c()
    Const PROC = "Demo_3_VB_Runtime_Error_DemoProc_3c"
    
    On Error GoTo eh
    
    mErH.BoP ErrSrc(PROC)
    Demo_3_VB_Runtime_Error_DemoProc_3d
    mErH.EoP ErrSrc(PROC)
    Exit Sub

eh:
    If mErH.ErrMsg(err_source:=ErrSrc(PROC)) = DebugOptResumeErrorLine Then Stop: Resume
End Sub

Private Sub Demo_3_VB_Runtime_Error_DemoProc_3d()
' ------------------------------------------------
' Note: The error line intentionally has no line
' number to demonstrate how it effects the error
' message.
' ------------------------------------------------
    Const PROC = "Demo_3_VB_Runtime_Error_DemoProc_3d"
    
    On Error GoTo eh

    mErH.BoP ErrSrc(PROC)
    Dim l As Long
    l = 7 / 0

xt: mErH.EoP ErrSrc(PROC)
    Exit Sub

eh: Select Case mErH.ErrMsg(err_source:=ErrSrc(PROC))
        Case DebugOptResumeErrorLine:       Stop: Resume
        Case DebugOptResumeNext:        Resume Next
        Case DebugOptCleanExit:   GoTo xt
    End Select
End Sub

Public Sub Demo_4_With_Debugging_Support()
' ----------------------------------------------
' Attention! This test requires the
' Conditional Compile Argument "Debugging = 1" !
' ----------------------------------------------
    Const PROC = "Demo_4_With_Debugging_Support"
    
    On Error GoTo eh
      
    mErH.BoP ErrSrc(PROC)
    Demo_4_With_Debugging_Support_DemoProc_5a
    mErH.EoP ErrSrc(PROC)
    Exit Sub

eh:
    If mErH.ErrMsg(err_source:=ErrSrc(PROC)) = DebugOptResumeErrorLine Then Stop: Resume
End Sub

Private Sub Demo_4_With_Debugging_Support_DemoProc_5a()
    Const PROC = "Demo_5_With_Debugging_Support_DemoProc_5a"
    
    On Error GoTo eh
       
    mErH.BoP ErrSrc(PROC)
376 Debug.Print ThisWorkbook.Named
    mErH.EoP ErrSrc(PROC)
    Exit Sub
    
eh:
    Select Case mErH.ErrMsg(err_source:=ErrSrc(PROC))
        Case DebugOptResumeErrorLine: Stop: Resume ' Continue with F8 to end up at the code line which caused the error
    End Select
End Sub

Public Sub Demo_5_No_Exit_Statement()
' -----------------------------------
' Exit statement missing
' -----------------------------------
    Const PROC = "Demo_6_No_Exit_Statement"
    
    On Error GoTo eh
    
eh:
    If mErH.ErrMsg(err_source:=ErrSrc(PROC)) = DebugOptResumeErrorLine Then Stop: Resume
End Sub

Public Sub Demo_6_Execution_Trace()
' ------------------------------------------------------
' White-box- and regression-test procedure obligatory
' to be performed after any code modification.
' Display of an execution trace along with this test
' requires a conditional compile argument ExecTrace = 1.
' ------------------------------------------------------
    Const PROC = "Demo_6_Execution_Trace"
    
    On Error GoTo eh
    
    mErH.BoP ErrSrc(PROC)
    Demo_6_Execution_Trace_DemoProc_6a
    mErH.EoP ErrSrc(PROC)
    Exit Sub

eh:
    If mErH.ErrMsg(err_source:=ErrSrc(PROC)) = DebugOptResumeErrorLine Then Stop: Resume
End Sub

Private Sub Demo_6_Execution_Trace_DemoProc_6a()
    Const PROC = "Demo_6_Execution_Trace_DemoProc_6a"
    
    On Error GoTo eh
    
    mErH.BoP ErrSrc(PROC)
    Demo_6_Execution_Trace_DemoProc_6b
    mErH.EoP ErrSrc(PROC)
    Exit Sub

eh:
    If mErH.ErrMsg(err_source:=ErrSrc(PROC)) = DebugOptResumeErrorLine Then Stop: Resume
End Sub

Private Sub Demo_6_Execution_Trace_DemoProc_6b()
    Const PROC = "Demo_6_Execution_Trace_DemoProc_6b"
    
    On Error GoTo eh
    
    mErH.BoP ErrSrc(PROC)
    
    Demo_6_Execution_Trace_DemoProc_6c
    
    Dim i As Long: Dim j As Long: j = 10000000
    mTrc.BoC PROC & " empty loop 1 to " & j
    For i = 1 To j
    Next i
    mTrc.EoC PROC & " empty loop 1 to " & j ' !!! the string must match with the BoC statement !!!
    
    mErH.EoP ErrSrc(PROC)
    Exit Sub

eh:
    If mErH.ErrMsg(err_source:=ErrSrc(PROC)) = DebugOptResumeErrorLine Then Stop: Resume
End Sub

Private Sub Demo_6_Execution_Trace_DemoProc_6c()
    Const PROC = "Demo_6_Execution_Trace_DemoProc_6c"
    
    On Error GoTo eh

    mErH.BoP ErrSrc(PROC)

xt: mErH.EoP ErrSrc(PROC)
    Exit Sub

eh: mErH.ErrMsg err_source:=ErrSrc(PROC)
End Sub

Private Sub Demo_7_Free_Button_Display()
    Const PROC = "Demo_7_Free_Button_Display"

    On Error GoTo eh

    Err.Raise AppErr(1), ErrSrc(PROC), "Display of a free defined button in addition to the usual Ok button (resumes the error when clicked)"
    Exit Sub

eh:
    Select Case mErH.ErrMsg(err_source:=ErrSrc(PROC), err_buttons:=vbOKOnly & "," & vbLf & ",My button")
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
