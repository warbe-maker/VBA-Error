Attribute VB_Name = "mAnyModule"
Option Explicit
' -------------------------------------------------------------------------------
' Sample module when using mErrHndlr
' -------------------------------------------------------------------------------
Const MODNAME = "mAnyModule" ' Module name for error handling and execution trace

Private Sub AnyProc()
' -------------------------------------------------------------------------------
' Sample procedure using mErrHndlr
' -------------------------------------------------------------------------------
Const PROC As String = "AnyProc" ' This procedure's name for error handling and execution trace

    On Error GoTo on_error
    BoP ErrSrc(PROC) ' Begin of Procedure (push stack and begin of execution trace)

    ' any code

exit_proc:
    ' any "finally" code
    EoP ErrSrc(PROC) ' End of Procedure (pop stack and end of execution trace)
    Exit Sub

on_error:
'~~ The conditional compile argument Debugging=1 allows to resume the statement which raised the error
'~~ or continue with any desired code line (highlite the desired code line and Ctrl+F9)
#If Debugging Then
    Debug.Print Err.Description: Stop: Resume
#End If
    '~~ Pass on the error to the calling procedure and finally
    '~~ display the error message when the entry procedure is reached
    mErrHndlr.ErrHndlr Err.Number, ErrSrc(PROC), Err.Description, Erl
                                                     
End Sub

Private Function ErrSrc(ByVal sProc As String) As String
    ErrSrc = Split(ThisWorkbook.Name, ".")(0) & "." & MODNAME & "." & sProc
End Function
