Attribute VB_Name = "mAnyModule"
Option Explicit
' -------------------------------------------------------------------------------
' Sample module when using mErrHndlr
' -------------------------------------------------------------------------------
Const MODNAME = "mAnyModule" ' Module name for error handling and execution trace

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

eh: mErH.ErrMsg err_source:=ErrSrc(PROC)
End Sub

Private Function ErrSrc(ByVal sProc As String) As String
    ErrSrc = Split(ThisWorkbook.Name, ".")(0) & "." & MODNAME & "." & sProc
End Function
