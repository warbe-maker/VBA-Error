Attribute VB_Name = "mErrHndlr"
Option Explicit
#Const AlternateMsgBox = 1  ' 1 = Error displayed by means of the Alternative MsgBox fMsg
                            ' 0 = Error displayed by means of the VBA MsgBox
' -----------------------------------------------------------------------------------------------
' Standard  Module mErrHndlr: Global error handling for any VBA Project.
'
' Methods: - AppErr   Converts a positive number into a negative error number ensuring it not
'                     conflicts with a VB Runtime Error. A negative error number is turned back into the
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

Private dctStack            As Dictionary
Private sErrHndlrEntryProc  As String
Private lSourceErrorNo      As Long
Private sErrorSource        As String
Private sErrorDescription   As String
Private sErrorPath          As String

Public Property Get ResumeButton() As String: ResumeButton = "Resume error" & vbLf & "code line":   End Property

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

Public Sub BoP(ByVal sErrSource As String)
' --------------------------------------------------------------------
' Keep record of the Begin of a Procedure by maintaining a call stack.
' --------------------------------------------------------------------
    
    StackPush sErrSource
    
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
#If ExecTrace Then
    If CallStack Is Nothing Then
        Set CallStack = New clsCallStack
    End If
    CallStack.TraceDsply
    Set CallStack = Nothing
#End If
End Sub

Public Sub EoP(ByVal sErrSource As String)
' -------------------------------------------
' End of Procedure. Maintains the call stack.
' -------------------------------------------
    
'    StackPop sErrSource
    
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

Public Function ErrHndlr(ByVal errnumber As Long, _
                         ByVal errsource As String, _
                         ByVal errdscrptn As String, _
                         ByVal errline As Long, _
                Optional ByVal buttons As Variant = vbOKOnly) As Variant
' ----------------------------------------------------------------------
' When the buttons argument specifies more than one button the error
'      message is immediately displayed and the users choice is returned
' else when the caller (errsource) is the "Entry Procedure" the error
'      is displayed with the path to the error
' else the error is passed on to the "Entry Procedure" whereby the
' .ErrorPath string is assebled.
' ----------------------------------------------------------------------
    
    Const PROC      As String = "ErrHndlr"
    Static sLine    As String   ' provided error line (if any) for the the finally displayed message
   
    If errnumber = 0 Then
        MsgBox "The error handling has been called with an error number = 0 !" & vbLf & vbLf & _
               "This indicates that in procedure" & vbLf & _
               ">>>>> " & errsource & " <<<<<" & vbLf & _
               "an ""Exit ..."" statement before the call of the error handling is missing!" _
               , vbExclamation, _
               "Exit ... statement missing in " & errsource & "!"
        Exit Function
    End If
#If Debugging Then
    buttons = buttons & "," & ResumeButton
#End If

    If CallStack Is Nothing Then Set CallStack = New clsCallStack
    If cllErrPath Is Nothing Then Set cllErrPath = New Collection
    If errline <> 0 Then sLine = errline Else sLine = "0"
    
    With CallStack
        If ErrPathIsEmpty Then
            '~~ When there's yet no error path collected this indicates that the
            '~~ error handler is executed the first time This is the error raising procedure. Backtracking to the entry procedure is due
            ErrPathAdd errsource & " (" & ErrorDetails(errnumber, errline) & ")"
            .TraceError errsource & ": " & ErrorDetails(errnumber, errline) & " """ & errdscrptn & """"
            lSourceErrorNo = errnumber
            sErrorSource = errsource
            sErrorDescription = errdscrptn
        ElseIf .ErrorNumber <> errnumber Then
            '~~ The error number had changed during the process of passing it on to the entry procedure
'            ErrHndlrErrPathAdd errsource & " (" & ErrorDetails(errnumber, errline) & ")"
            .TraceError errsource & ": " & ErrorDetails(errnumber, errline) & " """ & errdscrptn & """"
            .ErrorNumber = errnumber
        Else
            '~~ This is the error handling called during the "backtracing" process,
            '~~ i.e. the process when the error is passed on up to the entry procedure
'            ErrHndlrErrPathAdd errsource
        End If
        
        '~~ When the user has no choice for the user to press any button but the only one displayed
        '~~ and the Entry Procedure is known but yet not reached the path back up to the Entry Procedure
        '~~ is maintained and the error is passed on to the caller
        If ErrHndlrNumberOfButtons(buttons) = 1 _
        And sErrHndlrEntryProc <> vbNullString _
        And .EntryProc <> errsource Then
            ErrPathAdd errsource
            Err.Raise errnumber, errsource, errdscrptn
        End If
        
        '~~ When more than one button is displayed for the user to choose one
        '~~ or the Entry Procedure is unknown or has been reached
        '~~ the error is displayed
        If ErrHndlrNumberOfButtons(buttons) > 1 _
        Or .EntryProc = errsource _
        Or .EntryProc = vbNullString Then
            ErrPathAdd errsource
            ErrHndlr = ErrMsg(errnumber:=lSourceErrorNo, errsource:=sErrorSource, errdscrptn:=sErrorDescription, errline:=errline, errpath:=ErrPathErrMsg, buttons:=buttons)
            ErrPathErase
            StackErase
            Set cllErrPath = Nothing
        End If
        
        '~~ Each time a known Entry Procedure is reached the execution trace
        '~~ maintained by the BoP and EoP and the BoT and EoT statements is displayed
        If .EntryProc = errsource _
        Or .EntryProc = vbNullString Then
            DsplyTrace
        End If
            
    End With

End Function

Private Sub ErrHndlrErrPathAdd(ByVal s As String)
' -----------------------------------------------
' Adds s to the collection of procedures provided
' the procedure has not already been eadded.
' -----------------------------------------------

    If cllErrPath Is Nothing Then Set cllErrPath = New Collection
    If cllErrPath.Count = 0 _
    Then cllErrPath.Add s _
    Else If InStr(1, cllErrPath(cllErrPath.Count), s & " ") = 0 Then cllErrPath.Add s

End Sub

Private Function ErrHndlrNumberOfButtons(ByVal buttons As Variant) As Long
    ErrHndlrNumberOfButtons = UBound(Split(buttons, ",")) + 1
    If ErrHndlrNumberOfButtons = 2 Then
        If Split(buttons, ",")(1) = vbNullString Then
            ErrHndlrNumberOfButtons = 1
        End If
    End If
End Function

Public Function ErrMsg( _
                ByVal errnumber As Long, _
                ByVal errsource As String, _
                ByVal errdscrptn As String, _
                ByVal errline As Long, _
       Optional ByVal errpath As String = vbNullString, _
       Optional ByVal buttons As Variant = vbOKOnly) As Variant
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
        .MsgTitle = ErrMsgErrType(errnumber, errsource) & " in " & errsource
        .MsgLabel(1) = "Error Message/Description:":    .MsgText(1) = ErrMsgErrDscrptn(errdscrptn)
        .MsgLabel(2) = "Error path (call stack):":      .MsgText(2) = ErrPathErrMsg:   .MsgMonoSpaced(2) = True
        .MsgLabel(3) = "Info:":                         .MsgText(3) = ErrMsgInfo(errdscrptn)
        .MsgButtons = buttons
        .Setup
        .Show
        If ErrHndlrNumberOfButtons(buttons) = 1 Then
            ErrMsg = buttons ' a single reply buttons return value cannot be obtained since the form is unloaded with its click
        Else
            ErrMsg = .ReplyValue ' when more than one button is displayed the form is unloadhen the return value is obtained
        End If
    End With
#Else
    '~~ Display the error message by means of the VBA MsgBox
    sErrMsg = "Description: " & vbLf & ErrMsgErrDscrptn(errdscrptn) & vbLf & vbLf & _
              "Source:" & vbLf & errsource & ErrMsgErrLine(errline)
    sErrPath = ErrMsgErrPath
    If sErrPath <> vbNullString _
    Then sErrMsg = sErrMsg & vbLf & vbLf & _
                   "Path:" & vbLf & errpath
    If ErrMsgInfo(errdscrptn) <> vbNullString _
    Then sErrMsg = sErrMsg & vbLf & vbLf & _
                   "Info:" & vbLf & ErrMsgInfo(errdscrptn)
    ErrMsg = MsgBox(Prompt:=sErrMsg, buttons:=buttons, Title:=ErrMsgErrType(errnumber, errsource) & " in " & errsource & ErrMsgErrLine(errline))
#End If
End Function

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

Private Function ErrMsgErrPath() As String
' ---------------------------------------------------------------
' Returns errpath as indented list.
' ---------------------------------------------------------------
    Dim i       As Long
    Dim lIndent As Long: lIndent = 0
    
    For i = cllErrPath.Count To 1 Step -1
        If i = cllErrPath.Count Then
            ErrMsgErrPath = cllErrPath(i)
        Else
            lIndent = lIndent + 1
            ErrMsgErrPath = ErrMsgErrPath & vbLf & Space((lIndent - 1) * 3) & "'- " & cllErrPath(i)
        End If
    Next i

End Function

Private Function ErrMsgErrType(ByVal errnumber As Long, _
                               ByVal errsource As String) As String
' -------------------------------------------------------------------------
' Return the kind of error considering the Err.Source and the error number.
' -------------------------------------------------------------------------

   If InStr(1, Err.Source, "DAO") <> 0 _
   Or InStr(1, Err.Source, "ODBC Teradata Driver") <> 0 _
   Or InStr(1, Err.Source, "ODBC") <> 0 _
   Or InStr(1, Err.Source, "Oracle") <> 0 Then
      ErrMsgErrType = "Database Error " & errnumber
   Else
      If errnumber > 0 _
      Then ErrMsgErrType = "VB Runtime Error " & errnumber _
      Else ErrMsgErrType = "Application Error " & AppErr(errnumber)
   End If
   
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
    
    If errnumber < 0 _
    Then ErrorDetails = "Application error " & AppErr(errnumber) _
    Else ErrorDetails = "VB Runtime Error " & errnumber
    If sErrLine <> 0 Then ErrorDetails = ErrorDetails & " at line " & sErrLine

End Function

Private Function ErrSrc(ByVal sProc As String) As String
    ErrSrc = ThisWorkbook.Name & ">mErrHndlr" & ">" & sProc
End Function

Public Function Space(ByVal l As Long) As String
' --------------------------------------------------
' Unifies the VB differences SPACE$ and Space$ which
' lead to code diferences where there aren't any.
' --------------------------------------------------
    Space = VBA.Space$(l)
End Function

Private Sub StackErase()
    If Not dctStack Is Nothing Then dctStack.RemoveAll
End Sub

Private Sub StackInit()
    If dctStack Is Nothing Then Set dctStack = New Dictionary Else dctStack.RemoveAll
End Sub

Private Function StackIsEmpty() As Boolean
    StackIsEmpty = dctStack Is Nothing
    If Not StackIsEmpty Then StackIsEmpty = dctStack.Count = 0
End Function

Private Sub StackPop(ByVal s As String)

    Const PROC = "ErrHandlrStackPop"

    If StackIsEmpty _
    Then Err.Raise AppErr(1), ErrSrc(PROC), "No item '" & s & "' on the stack!"
    If dctStack.Items()(dctStack.Count - 1) <> s _
    Then Err.Raise AppErr(1), ErrSrc(PROC), "No item '" & s & "' on the stack!"

    '~~ Remove item s from stack
    dctStack.Remove dctStack.Count

End Sub

Private Sub StackPush(ByVal s As String)

    If dctStack Is Nothing Then Set dctStack = New Dictionary
    If dctStack.Count = 0 Then
        sErrHndlrEntryProc = s ' First pushed = bottom item = entry procedure
    End If
    dctStack.Add dctStack.Count + 1, s

End Sub

Private Sub ErrPathAdd(ByVal s As String)
    If ErrPathIsEmpty Then
        sErrorPath = s
    Else
        If InStr(sErrorPath, s & " ") = 0 Then
            sErrorPath = s & vbLf & sErrorPath
        End If
    End If
End Sub

Private Function ErrPathIsEmpty() As Boolean
   ErrPathIsEmpty = sErrorPath = vbNullString
End Function

Private Sub ErrPathErase()
    sErrorPath = vbNullString
End Sub

Private Function ErrPathErrMsg() As String

    Dim i    As Long: i = 0
    Dim s    As String:  s = sErrorPath
   
    While InStr(s, vbLf) <> 0
        s = Replace(s, vbLf, "@@@@@" & Space(i * 2) & "|_", 1, 1)
        i = i + 1
    Wend
    ErrPathErrMsg = Replace(s, "@@@@@", vbLf)
    
End Function
