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
' - Debugging choice:           No
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
' - Debugging choice:           No
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
' - Debugging choice:           No
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
' - Debugging choice:           Yes, demonstrated
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
' - Debugging choice:                 Yes (with Conditional Compile Argument 'Debugging = 1'
' -----------------------------------------------------------------------
Const PROC  As String = "ErrorHandling_Elaborated1"

    On Error GoTo on_error
    BoP ErrSrc(PROC)    ' Push procedure on call stack
    
    ErrorHandling_Elaborated2 10, 0

exit_proc:
    EoP ErrSrc(PROC)    ' Pull procedure from call stack
    Exit Sub

on_error:
#If Debugging Then
    Debug.Print Err.Description: Stop: Resume    ' Resumes the statement which caused the error, turned into a comment to continue
#End If
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
' - Debugging choice:                 Yes (with Conditional Compile Argument 'Debugging = 1'
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
#If Debugging Then
    Debug.Print Err.Description: Stop: Resume    ' Resumes the statement which caused the error, turned into a comment to continue
#End If
    mErrHndlr.ErrHndlr Err.Number, ErrSrc(PROC), Err.Description, Erl
End Function

Private Function ErrSrc(ByVal s As String) As String
' --------------------------------------------------
' Prefix procedure name (s) by this module's name.
' --------------------------------------------------
    ErrSrc = ThisWorkbook.Name & ">mDemo" & ">" & s
End Function
