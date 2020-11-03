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

    On Error GoTo eh
    mErH.BoP ErrSrc(PROC) ' Begin of Procedure (push stack and begin of execution trace)

    ' any code

xt: ' any "finally" code
    mErH.EoP ErrSrc(PROC) ' End of Procedure (pop stack and end of execution trace)
    Exit Sub

eh: mErH.ErrMsg err.Number, ErrSrc(PROC), err.Description, Erl
End Sub

Private Function ErrSrc(ByVal sProc As String) As String
    ErrSrc = Split(ThisWorkbook.Name, ".")(0) & "." & MODNAME & "." & sProc
End Function
