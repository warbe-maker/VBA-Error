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
    BoP errsrc(PROC) ' Begin of Procedure (push stack and begin of execution trace)

    ' any code

exit_proc:
    ' any "finally" code
    EoP errsrc(PROC) ' End of Procedure (pop stack and end of execution trace)
    Exit Sub

on_error:
    mErrHndlr.ErrHndlr Err.Number, errsrc(PROC), Err.Description, Erl
End Sub

Private Function errsrc(ByVal sProc As String) As String
    errsrc = Split(ThisWorkbook.Name, ".")(0) & "." & MODNAME & "." & sProc
End Function
