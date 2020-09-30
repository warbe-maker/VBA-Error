Attribute VB_Name = "mErrHndlr"
Option Explicit
#Const AlternateMsgBox = 0  ' 1 = Error displayed by means of the Alternative MsgBox fMsg
                            ' 0 = Error displayed by means of the VBA MsgBox
' -----------------------------------------------------------------------------------------------
' Standard  Module mErrHndlr: Global error handling for any VBA Project.
'
' Methods: - AppErr   Converts a positive number into a negative error number ensuring it not
'                     conflicts with a VB error. A negative error number is turned back into the
'                     original positive Application  Error Number.
'          - ErrHndlr Either passes on the error to the caller or when the entry procedure is
'                     reached, displays the error with a complete path from the entry procedure
'                     to the procedure with the error.
'          - BoP      Maintains the call stack at the Begin of a Procedure (optional when using
'                     this common error handler)
'          - EoP      Maintains the call stack at the End of a Procedure, triggers the display of
'                     the Execution Trace when the entry procedure is finished and the
'                     Conditional Compile Argument ExecTrace = 1
'          - BoT      Begin of Trace. In contrast to BoP this is for any group of code lines
'                     within a procedure
'          - EoT      End of trace corresponding with the BoT.
'          - ErrMsg   Displays the error message in a proper formated manner
'                     The local Conditional Compile Constant 'ErrMsg = "Custom" enforces the use
'                     of the Alternative VBA MsgBox which provideds a significantly better
'                     readability. The method may be used with or without a call stack.
'
' Usage:   Private/Public Sub/Function any()
'             Const PROC = "any"  ' procedure's name as error source
'
'             On Error GoTo on_error
'             BoP ErrSrc(PROC)   ' puts the procedure on the call stack
'
'             ' <any code>
'
'          exit_proc:
'                               ' <any "finally" code like re-protecting an unprotected sheet for instance>
'                               EoP ErrSrc(PROC)   ' takes the procedure off from the call stack
'                               Exit Sub/Function
'
'          on_error:
'          #If Debugging = 1 Then
'             Debug.Print Err.Description: Stop: Resume    ' allows to exactly locate the line where the error occurs
'          #Else
'             mErrHndlr.ErrHndlr Err.Number, ErrSrc(PROC), Err.Description, Erl
'          #End If
'
' Note: When the call stack is not maintained the ErrHndlr will display the message
'       immediately with the procedure the error occours. When the call stack is
'       maintained, the error message will display the call path to the error beginning
'       with the first (entry) procedure in which the call stack is maintained all the
'       call sequence down to the procedure where the error occoured.
'
' Uses: fMsg (only when the Conditional Compile Argument AlternateMsgBox = 1)
'
' Requires: Reference to "Microsoft Scripting Runtime"
'
' W. Rauschenberger, Berlin January 2020, https://github.com/warbe-maker/Common-VBA-Error-Handler
' -----------------------------------------------------------------------------------------------
' ~~ Begin of Declarations for withdrawing the title bar ------------------------------------
'Private Declare PtrSafe Function GetForegroundWindow Lib "User32.dll" () As LongPtr
'Private Declare PtrSafe Function GetWindowLong Lib "User32.dll" _
'                          Alias "GetWindowLongPtrA" (ByVal hwnd As LongPtr, _
'                                                     ByVal nIndex As Long) As LongPtr
'Private Declare PtrSafe Function SetWindowLong Lib "User32.dll" _
'                          Alias "SetWindowLongPtrA" (ByVal hwnd As LongPtr, _
'                                                     ByVal nIndex As Long, _
'                                                     ByVal dwNewLong As LongPtr) As LongPtr
'Private Declare PtrSafe Function DrawMenuBar Lib "User32.dll" (ByVal hwnd As LongPtr) As Long
'Private Const GWL_STYLE  As Long = (-16)
'Private Const WS_CAPTION As Long = &HC00000
' ~~ End of Declarations for withdrawing the title bar --------------------------------------

Public Enum StartupPosition         ' ---------------------------
    Manual = 0                      ' Used to position the
    CenterOwner = 1                 ' final setup message form
    CenterScreen = 2                ' horizontally and vertically
    WindowsDefault = 3              ' centered on the screen
End Enum                            ' ---------------------------

Public Type tSection                ' ------------------
       sLabel As String             ' Structure of the
       sText As String              ' UserForm's
       bMonspaced As Boolean        ' message area which
End Type                            ' consists of

Public Type tMessage                ' three message
       section(1 To 3) As tSection  ' sections
End Type

Public CallStack    As clsCallStack
Public dicTrace     As Dictionary       ' Procedure execution trancing records
Private cllErrPath  As Collection

Public Function AppErr(ByVal lNo As Long) As Long
' -----------------------------------------------------------------
' Used with Err.Raise AppErr().
' When lNo is > 0 it is considered an Application error number and
' vbObjectErrro is added to it to turn it into a negative number
' in order not to confuse with a VB runtime error. When lNo is
' negative it is considered an Application error and vbObjectError
' is added to convert it back into its origin positive number.
' ------------------------------------------------------------------
    IIf lNo < 0, AppErr = lNo - vbObjectError, AppErr = vbObjectError + lNo
End Function

' Keep record of the Begin of a Procedure by maintaining a call stack.
' --------------------------------------------------------------------
Public Sub BoP(ByVal sErrSource As String)
    If CallStack Is Nothing Then
        Set CallStack = New clsCallStack
    ElseIf CallStack.StackIsEmpty Then
        Set CallStack = Nothing
        Set CallStack = New clsCallStack
    End If
    CallStack.TraceBegin sErrSource   ' implicite start of the procedure's exec trace
    CallStack.StackPush sErrSource

End Sub

Public Sub BoT(ByVal s As String)
' ---------------------------------------
' Explicit execution trace start for (s).
' ---------------------------------------
#If ExecTrace Then
    CallStack.TraceBegin s
#End If
End Sub

Private Sub DsplyTrace()
' ------------------------------------------------------------
' Displays the execution trace when the entry procedure has
' been reached.
' Note: The call stack is primarily used to detect whether or
'       not there was an initial entry procedure. It is not
'       used to maintain the error path which is done in any
'       case along with the process of passing on the error
'       to the calling procedure.
' ------------------------------------------------------------
    If CallStack Is Nothing Then
        Set CallStack = New clsCallStack
    End If
    CallStack.TraceDsply
    Set CallStack = Nothing
End Sub

Public Sub EoP(ByVal sErrSource As String)
' -------------------------------------------
' End of Procedure. Maintains the call stack.
' -------------------------------------------
    If Not CallStack Is Nothing Then
        CallStack.TraceEnd sErrSource
        CallStack.StackPop sErrSource

        If CallStack.StackIsEmpty Then
            If CallStack.ErrorPath = vbNullString Then
                Set CallStack = Nothing
            End If
        End If
    End If
End Sub

Public Sub EoT(ByVal s As String)
' -------------------------------------
' Explicit execution trace end for (s).
' -------------------------------------
    CallStack.TraceEnd s
End Sub

Public Sub ErrHndlr(ByVal errnumber As Long, _
                    ByVal errsource As String, _
                    ByVal errdscrptn As String, _
                    ByVal errline As String)
' -----------------------------------------------
' When the caller (errsource) is the entry
' procedure the error is displayed with the path
' to the error. Otherwise the error is raised
' again to pass it on to the calling procedure.
' The .ErrorPath string is maintained with all
' the way up to the calling procedure.
' -----------------------------------------------
Const PROC      As String = "ErrHndlr"
Static sLine    As String   ' provided error line (if any) for the the finally displayed message
   
   
    If errnumber = 0 Then
        MsgBox "Apparently an ""Exit ..."" statement before the error handling is missing! The error handling has been aproached with a 0 error number!", vbExclamation, _
               "Problem deteced with " & ErrSrc(PROC)
        Exit Sub
    End If
    
    If CallStack Is Nothing Then Set CallStack = New clsCallStack
    If errline <> 0 Then sLine = errline
    
    With CallStack
        If .ErrorSource = vbNullString Then
            '~~ When the ErrorSource property is still empty, this indicates that the
            '~~ error handler is executed the first time This is the error raising procedure. Backtracking to the entry procedure is due
            Set cllErrPath = Nothing: Set cllErrPath = New Collection
            .ErrorSource = errsource
            .SourceErrorNo = errnumber
            .ErrorNumber = errnumber
            .ErrorDescription = errdscrptn
            .ErrorPath = .ErrorPath & errsource & " (" & ErrorDetails(errnumber, errline) & ")" & vbLf
            .TraceError errsource & ": " & ErrorDetails(errnumber, errline) & " """ & errdscrptn & """"
        ElseIf .ErrorNumber <> errnumber Then
            '~~ The error number had changed during the process
            '~~ of passing the error on to the entry procedure
            .ErrorPath = .ErrorPath & errsource & " (" & ErrorDetails(errnumber, errline) & ")" & vbLf
            .TraceError errsource & ": " & ErrorDetails(errnumber, errline) & " """ & errdscrptn & """"
            .ErrorNumber = errnumber
        Else
            '~~ This is the error handling called during the "backtracing" process,
            '~~ i.e. the process when the error is passed on up to the entry procedure
            .ErrorPath = .ErrorPath & errsource & vbLf
        End If
        
        If .EntryProc <> errsource Then ' And Not .ErrorPath <> vbNullString Then
            '~~ This is the call of the error handling for the error causing procedure or
            '~~ any of the procedures up to the entry procedure which has yet not been reached.
            '~~ The "backtrace" error path is maintained ....
            cllErrPath.Add errsource
            '~~ ... and the error is passed on to the calling procedure.
            Err.Raise errnumber, errsource, errdscrptn
        
        ElseIf .EntryProc = errsource Then
            '~~ The entry procedure has been reached
            '~~ The "backtrace" error path is maintained ....
            cllErrPath.Add errsource
            '~~ .. and the error is displayed
            ErrMsg .SourceErrorNo, .ErrorSource, .ErrorDescription, sLine
            
#If ExecTrace Then
            '~~ Display of the full execution trace which had been maintained by
            '~~ the BoP and EoP and the BoT and EoT statements executed
            DsplyTrace
#End If
        End If
    End With

End Sub

Public Sub ErrMsg(ByVal errnumber As Long, _
                  ByVal errsource As String, _
                  ByVal errdscrptn As String, _
                  ByVal errline As String)
' -----------------------------------------------------------------------------------------
' Displays the error message either by means of VBA MsgBox or, when the Conditional Compile
' Argument AlternateMsgBox = 1 by means of the Alternate VBA MsgBox (fMsg). In any case the
' path to the error may be displayed, provided the entry procedure has BoP/EoP code lines.
'
' W. Rauschenberger, Berlin, Sept 2020
' -----------------------------------------------------------------------------------------
    Dim sErrMsg     As String
    Dim sErrPath    As String
    
#If AlternateMsgBox Then
    '~~ Display the error message by means of the Common UserForm fMsg
    With fMsg
        .MsgTitle = ErrMsgErrType(errnumber, errsource) & " in " & errsource & ErrMsgErrLine(errline)
        .MsgLabel(1) = "Error Message/Description:":    .MsgText(1) = ErrMsgErrDscrptn(errdscrptn)
        .MsgLabel(2) = "Error path (call stack):":      .MsgText(2) = ErrMsgErrPath(errline):   .MsgMonoSpaced(2) = True
        .MsgLabel(3) = "Info:":                         .MsgText(3) = ErrMsgInfo(errdscrptn)
        .MsgButtons = vbOKOnly
        .Setup
        .Show
    End With
#Else
    '~~ Display the error message by means of the VBA MsgBox
    sErrMsg = "Description: " & vbLf & ErrMsgErrDscrptn(errdscrptn) & vbLf & vbLf & _
              "Source:" & vbLf & errsource & ErrMsgErrLine(errline)
    sErrPath = ErrMsgErrPath(errline)
    If sErrPath <> vbNullString _
    Then sErrMsg = sErrMsg & vbLf & vbLf & _
                   "Path:" & vbLf & sErrPath
    If ErrMsgInfo(errdscrptn) <> vbNullString _
    Then sErrMsg = sErrMsg & vbLf & vbLf & _
                   "Info:" & vbLf & ErrMsgInfo(errdscrptn)
    MsgBox sErrMsg, vbCritical, ErrMsgErrType(errnumber, errsource) & " in " & errsource & ErrMsgErrLine(errline)
#End If
End Sub

Private Function ErrMsgErrDscrptn(ByVal s As String) As String
' -------------------------------------------------------------------
' Return the string before a "||" in the error description. May only
' be the case when the error has been raised by means of err.Raise
' which means when it is an "Application Error".
' -------------------------------------------------------------------
    If InStr(s, DCONCAT) <> 0 _
    Then ErrMsgErrDscrptn = Split(s, DCONCAT)(0) _
    Else ErrMsgErrDscrptn = s
End Function

Private Function ErrMsgErrLine(ByVal errline As Long) As String
    If errline <> 0 _
    Then ErrMsgErrLine = " (at line " & errline & ")" _
    Else ErrMsgErrLine = vbNullString
End Function

Private Function ErrMsgErrPath(ByVal errline As Long) As String
' ------------------------------------------------------------------------------
' Path from the "Entry Procedure" - i.e. the first procedure in the call stack
' with an BoP/EoP code line - all the way down to the procedure in which the
' error occoured. When the call stack had not been maintained the path is empty.
' ------------------------------------------------------------------------------
    Dim i, iIndent As Long
    
    If Not CallStack Is Nothing Then
        If Not CallStack.ErrorPath = vbNullString Then
            CallStack.TraceEndTime = Now()
            CallStack.StackUnwind
        End If
    End If
    
    For i = cllErrPath.Count To 1 Step -1
        If i = cllErrPath.Count Then
            ErrMsgErrPath = cllErrPath(i) & vbLf
        ElseIf i = 1 Then
            ErrMsgErrPath = ErrMsgErrPath & mBasic.Space((iIndent) * 2) & "|_" & cllErrPath(i) & ErrMsgErrLine(errline)
        Else
            ErrMsgErrPath = ErrMsgErrPath & mBasic.Space((iIndent) * 2) & "|_" & cllErrPath(i) & vbLf
        End If
        iIndent = iIndent + 1
    Next i

End Function

Private Function ErrMsgInfo(ByVal s As String) As String
' -------------------------------------------------------------------
' Return the string after a "||" in the error description. May only
' be the case when the error has been raised by means of err.Raise
' which means when it is an "Application Error".
' -------------------------------------------------------------------
    If InStr(s, DCONCAT) <> 0 _
    Then ErrMsgInfo = Split(s, DCONCAT)(1) _
    Else ErrMsgInfo = vbNullString
End Function

Private Function ErrorDetails(ByVal errnumber As Long, _
                              ByVal sErrLine As String) As String
' -----------------------------------------------------------------
' Returns kind of error, error number, and error line if available.
' -----------------------------------------------------------------
Dim s As String
    If errnumber < 0 Then
        s = "App error " & AppErr(errnumber)
    Else
        s = "VB error " & errnumber
    End If
    If sErrLine <> 0 Then
        s = s & " at line " & sErrLine
    End If
    ErrorDetails = s
End Function

Private Function ErrMsgErrType( _
        ByVal errnumber As Long, _
        ByVal errsource As String) As String
' ------------------------------------------
' Return the kind of error considering the
' Err.Source and the error number.
' ------------------------------------------

   If InStr(1, Err.Source, "DAO") <> 0 _
   Or InStr(1, Err.Source, "ODBC Teradata Driver") <> 0 _
   Or InStr(1, Err.Source, "ODBC") <> 0 _
   Or InStr(1, Err.Source, "Oracle") <> 0 Then
      ErrMsgErrType = "Database Error"
   Else
      If errnumber > 0 _
      Then ErrMsgErrType = "VB Runtime Error" _
      Else ErrMsgErrType = "Application Error"
   End If
   
End Function

Private Function ErrSrc(ByVal sProc As String) As String
    ErrSrc = ThisWorkbook.Name & ">mErrHndlr" & ">" & sProc
End Function
