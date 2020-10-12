Attribute VB_Name = "mTest"
Option Explicit

Private Function ErrSrc(ByVal s As String) As String
    ErrSrc = "mTest." & s
End Function

Public Sub Test_1_Unpaired_BoP_EoP()
' ---------------------------------------------------
' White-box- and regression-test procedure obligatory
' to be performed after any code modification.
' ---------------------------------------------------
    
    Const PROC = "Test_1_Unpaired_BoP_EoP"
    
    Test_1_Unpaired_BoP_EoP_TestProc_1a
    
exit_proc:
    EoP ErrSrc(PROC) ' unpaired code line! BoP is missing
    Exit Sub

on_error:
    mErrHndlr.ErrHndlr errnumber:=Err.Number, errsource:=ErrSrc(PROC), errdscrptn:=Err.Description, errline:=Erl
End Sub

Private Sub Test_1_Unpaired_BoP_EoP_TestProc_1a()
' -----------------------------------------------
' The error handler is trying its very best not
' to confuse with unpaired BoP/EoP code lines.
' However, it depends at which level this is the
' case.
' -----------------------------------------------

    Const PROC = "Test_1_Unpaired_BoP_EoP_TestProc_1a"
    
    Test_1_Unpaired_BoP_EoP_TestProc_1b ' missing End of Procedure statement
    
    BoP ErrSrc(PROC)
    
    Test_1_Unpaired_BoP_EoP_TestProc_1d   ' missing Begin of Procedure statement
    Test_1_Unpaired_BoP_EoP_TestProc_1e   ' missing Begin of Procedure statement
    
    EoP ErrSrc(PROC)
    Exit Sub

on_error:
    mErrHndlr.ErrHndlr errnumber:=Err.Number, errsource:=ErrSrc(PROC), errdscrptn:=Err.Description, errline:=Erl
End Sub

Private Sub Test_1_Unpaired_BoP_EoP_TestProc_1b()
    
    Const PROC = "Test_1_Unpaired_BoP_EoP_TestProc_1b"
    On Error GoTo on_error
    
    BoP ErrSrc(PROC)
    Test_1_Unpaired_BoP_EoP_TestProc_1c
    EoP ErrSrc(PROC)
    Exit Sub
    
on_error:
    mErrHndlr.ErrHndlr errnumber:=Err.Number, errsource:=ErrSrc(PROC), errdscrptn:=Err.Description, errline:=Erl
End Sub

Private Sub Test_1_Unpaired_BoP_EoP_TestProc_1c()
    
    Const PROC = "Test_1_Unpaired_BoP_EoP_TestProc_1c"
    On Error GoTo on_error
    
    BoP ErrSrc(PROC)
    EoP ErrSrc(PROC)
    Exit Sub
    
on_error:
    mErrHndlr.ErrHndlr errnumber:=Err.Number, errsource:=ErrSrc(PROC), errdscrptn:=Err.Description, errline:=Erl
End Sub

Private Sub Test_1_Unpaired_BoP_EoP_TestProc_1d()

    Const PROC = "Test_1_Unpaired_BoP_EoP_TestProc_1d"
    On Error GoTo on_error
    
    BoP ErrSrc(PROC) & " (missing EoP)"
    Exit Sub

on_error:
    mErrHndlr.ErrHndlr errnumber:=Err.Number, errsource:=ErrSrc(PROC), errdscrptn:=Err.Description, errline:=Erl
End Sub

Private Sub Test_1_Unpaired_BoP_EoP_TestProc_1e()
' -----------------------------------------------
' BoP missing
' -----------------------------------------------

    Const PROC = "Test_1_Unpaired_BoP_EoP_TestProc_1e"
    On Error GoTo on_error

exit_proc:
    EoP ErrSrc(PROC) & " (missing BoP)"
    Exit Sub

on_error:
    mErrHndlr.ErrHndlr errnumber:=Err.Number, errsource:=ErrSrc(PROC), errdscrptn:=Err.Description, errline:=Erl
End Sub

Public Sub Test_2_Application_Error()
' ------------------------------------------------------
' White-box- and regression-test procedure obligatory
' to be performed after any code modification.
' Display of an execution trace along with this test
' requires a conditional compile argument ExecTrace = 1.
' ------------------------------------------------------
    
    Const PROC = "Test_2_Application_Error"
    On Error GoTo on_error
    
    BoP ErrSrc(PROC)
    Test_2_Application_Error_TestProc_2a
    EoP ErrSrc(PROC)
    Exit Sub

on_error:
    mErrHndlr.ErrHndlr errnumber:=Err.Number, errsource:=ErrSrc(PROC), errdscrptn:=Err.Description, errline:=Erl
End Sub

Private Sub Test_2_Application_Error_TestProc_2a()

    Const PROC = "Test_2_Application_Error_TestProc_2a"
    On Error GoTo on_error
    
    BoP ErrSrc(PROC)
    Test_2_Application_Error_TestProc_2b
    EoP ErrSrc(PROC)
    Exit Sub

on_error:
    mErrHndlr.ErrHndlr errnumber:=Err.Number, errsource:=ErrSrc(PROC), errdscrptn:=Err.Description, errline:=Erl
End Sub

Private Sub Test_2_Application_Error_TestProc_2b()
    
    Const PROC = "Test_2_Application_Error_TestProc_2b"
    On Error GoTo on_error
    
    BoP ErrSrc(PROC)
    Test_2_Application_Error_TestProc_2c
    EoP ErrSrc(PROC)
    Exit Sub

on_error:
    mErrHndlr.ErrHndlr errnumber:=Err.Number, errsource:=ErrSrc(PROC), errdscrptn:=Err.Description, errline:=Erl
End Sub

Private Sub Test_2_Application_Error_TestProc_2c()
    
    Const PROC = "Test_2_Application_Error_TestProc_2c"
    On Error GoTo on_error

    BoP ErrSrc(PROC)
199 Err.Raise AppErr(1), ErrSrc(PROC), _
        "This is a programmed error!" & DCONCAT & _
        "Attention: This is not a VBA error!" & vbLf & _
        "The function AppErr() adds the 'vbObjectError' variable to assure non-conflicting programmed (Application) error numbers " & _
        "which allows programmed error numbers 1 to n for each individual procedure." & vbLf & _
        "Furthermore the error line has been added manually to the code to demonstrate its use when available." & vbLf & _
        "This additional info had been concatenated with the error description by two vertical bars indicating it as additional message."
    EoP ErrSrc(PROC)
    Exit Sub

on_error:
#If Debugging Then
    ' Debug.Print Err.Description: Stop: Resume
#End If
    mErrHndlr.ErrHndlr errnumber:=Err.Number, errsource:=ErrSrc(PROC), errdscrptn:=Err.Description, errline:=Erl
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
    On Error GoTo on_error
    
    BoP ErrSrc(PROC)
    Test_3_VB_Runtime_Error_TestProc_3a
    EoP ErrSrc(PROC)
    Exit Sub

on_error:
    mErrHndlr.ErrHndlr errnumber:=Err.Number, errsource:=ErrSrc(PROC), errdscrptn:=Err.Description, errline:=Erl
End Sub

Private Sub Test_3_VB_Runtime_Error_TestProc_3a()

    Const PROC = "Test_3_VB_Runtime_Error_TestProc_3a"
    On Error GoTo on_error

    BoP ErrSrc(PROC)
    Test_3_VB_Runtime_Error_TestProc_3b
    EoP ErrSrc(PROC)
    Exit Sub

on_error:
    mErrHndlr.ErrHndlr errnumber:=Err.Number, errsource:=ErrSrc(PROC), errdscrptn:=Err.Description, errline:=Erl
End Sub

Private Sub Test_3_VB_Runtime_Error_TestProc_3b()
    
    Const PROC = "Test_3_VB_Runtime_Error_TestProc_3b"
    On Error GoTo on_error

    BoP ErrSrc(PROC)
    Test_3_VB_Runtime_Error_TestProc_3c
    EoP ErrSrc(PROC)
    Exit Sub

on_error:
    mErrHndlr.ErrHndlr errnumber:=Err.Number, errsource:=ErrSrc(PROC), errdscrptn:=Err.Description, errline:=Erl
End Sub

Private Sub Test_3_VB_Runtime_Error_TestProc_3c()

    Const PROC = "Test_3_VB_Runtime_Error_TestProc_3c"
    On Error GoTo on_error
    
    BoP ErrSrc(PROC)
    Test_3_VB_Runtime_Error_TestProc_3d
    EoP ErrSrc(PROC)
    Exit Sub

on_error:
#If Debugging Then
    Debug.Print Err.Description: Stop: Resume
#End If
    mErrHndlr.ErrHndlr errnumber:=Err.Number, errsource:=ErrSrc(PROC), errdscrptn:=Err.Description, errline:=Erl
End Sub

Private Sub Test_3_VB_Runtime_Error_TestProc_3d()
    
    Const PROC = "Test_3_VB_Runtime_Error_TestProc_3d"
    On Error GoTo on_error

    BoP ErrSrc(PROC)
    Dim l As Long
423 l = 7 / 0
    EoP ErrSrc(PROC)
    Exit Sub

on_error:
    mErrHndlr.ErrHndlr errnumber:=Err.Number, errsource:=ErrSrc(PROC), errdscrptn:=Err.Description, errline:=Erl
End Sub

Public Sub Test_4_Debugging_with_ErrHndlr()

    On Error GoTo on_error
    Const PROC = "Test_4_Debugging_with_ErrHndlr"
      
    BoP ErrSrc(PROC)
    Test_4_Debugging_with_ErrHndlr_TestProc_5a
    EoP ErrSrc(PROC)
    Exit Sub

on_error:
    mErrHndlr.ErrHndlr errnumber:=Err.Number, errsource:=ErrSrc(PROC), errdscrptn:=Err.Description, errline:=Erl
End Sub

Private Sub Test_4_Debugging_with_ErrHndlr_TestProc_5a()

    Const PROC = "Test_5_Debugging_with_ErrHndlr_TestProc_5a"
    On Error GoTo on_error
       
    BoP ErrSrc(PROC)
15  Debug.Print ThisWorkbook.Named
    EoP ErrSrc(PROC)
    Exit Sub
    
on_error:
    Select Case mErrHndlr.ErrHndlr(errnumber:=Err.Number, errsource:=ErrSrc(PROC), errdscrptn:=Err.Description, errline:=Erl)
        Case ResumeButton: Stop: Resume ' Continue with F8 to end up at the code line which caused the error
    End Select
End Sub

Public Sub Test_5_No_Exit_Statement()
' -----------------------------------
' Exit statement missing
' -----------------------------------

    Const PROC = "Test_6_No_Exit_Statement"
    On Error GoTo on_error
    
on_error:
    mErrHndlr.ErrHndlr errnumber:=Err.Number, errsource:=ErrSrc(PROC), errdscrptn:=Err.Description, errline:=Erl
End Sub

