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

' Used with Err.Raise AppErr() to convert a positive application error number
' into a negative number to avoid any conflict with a VB error. Used when the
' error is displayed to turn the negative number back into the original
' positive application number.
' The function ensures that a programmed (application) error numbers never
' conflicts with VB error numbers by adding vbObjectError which turns it
' into a negative value. In return, translates a negative error number
' back into an Application error number. The latter is the reason why this
' function must never be used with a true VB error number.
' ------------------------------------------------------------------------
Public Function AppErr(ByVal lNo As Long) As Long
    If lNo < 0 Then
        AppErr = lNo - vbObjectError
    Else
        AppErr = vbObjectError + lNo
    End If
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

Public Sub ErrHndlr(ByVal lErrNo As Long, _
                    ByVal sErrSource As String, _
                    ByVal sErrText As String, _
                    ByVal sErrLine As String)
' -----------------------------------------------
' When the caller (sErrSource) is the entry
' procedure the error is displayed with the path
' to the error. Otherwise the error is raised
' again to pass it on to the calling procedure.
' The .ErrorPath string is maintained with all
' the way up to the calling procedure.
' -----------------------------------------------
Const PROC      As String = "ErrHndlr"
Static sLine    As String   ' provided error line (if any) for the the finally displayed message
   
   
    If lErrNo = 0 Then
        MsgBox "Apparently an ""Exit ..."" statement before the error handling is missing! The error handling has been aproached with a 0 error number!", vbExclamation, _
               "Problem deteced with " & ErrSrc(PROC)
        Exit Sub
    End If
    
    If CallStack Is Nothing Then Set CallStack = New clsCallStack
    If sErrLine <> 0 Then sLine = sErrLine
    
    With CallStack
        If .ErrorSource = vbNullString Then
            '~~ When the ErrorSource property is still empty, this indicates that the
            '~~ error handler is executed the first time This is the error raising procedure. Backtracking to the entry procedure is due
            Set cllErrPath = Nothing: Set cllErrPath = New Collection
            .ErrorSource = sErrSource
            .SourceErrorNo = lErrNo
            .ErrorNumber = lErrNo
            .ErrorDescription = sErrText
            .ErrorPath = .ErrorPath & sErrSource & " (" & ErrorDetails(lErrNo, sErrLine) & ")" & vbLf
            .TraceError sErrSource & ": " & ErrorDetails(lErrNo, sErrLine) & " """ & sErrText & """"
        ElseIf .ErrorNumber <> lErrNo Then
            '~~ The error number had changed during the process
            '~~ of passing the error on to the entry procedure
            .ErrorPath = .ErrorPath & sErrSource & " (" & ErrorDetails(lErrNo, sErrLine) & ")" & vbLf
            .TraceError sErrSource & ": " & ErrorDetails(lErrNo, sErrLine) & " """ & sErrText & """"
            .ErrorNumber = lErrNo
        Else
            '~~ This is the error handling called during the "backtracing" process,
            '~~ i.e. the process when the error is passed on up to the entry procedure
            .ErrorPath = .ErrorPath & sErrSource & vbLf
        End If
        
        If .EntryProc <> sErrSource Then ' And Not .ErrorPath <> vbNullString Then
            '~~ This is the call of the error handling for the error causing procedure or
            '~~ any of the procedures up to the entry procedure which has yet not been reached.
            '~~ The "backtrace" error path is maintained ....
            cllErrPath.Add sErrSource
            '~~ ... and the error is passed on to the calling procedure.
            Err.Raise lErrNo, sErrSource, sErrText
        
        ElseIf .EntryProc = sErrSource Then
            '~~ The entry procedure has been reached
            '~~ The "backtrace" error path is maintained ....
            cllErrPath.Add sErrSource
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

Public Sub ErrMsg(ByVal lErrNo As Long, _
                  ByVal sErrSrc As String, _
                  ByVal sErrDesc As String, _
                  ByVal sErrLine As String)
' -----------------------------------------------------------------------------------------
' Displays the error message either by means of VBA MsgBox or, when the Conditional Compile
' Argument AlternateMsgBox = 1 by means of the Alternate VBA MsgBox (fMsg). In any case the
' path to the error may be displayed, provided the entry procedure has BoP/EoP code lines.
'
' W. Rauschenberger, Berlin, Sept 2020
' -----------------------------------------------------------------------------------------
Dim sErrMsg     As String
Dim sTitle      As String
Dim sErrPath    As String
Dim sIndicate   As String
Dim i           As Long
Dim sErrText    As String
Dim sErrInfo    As String
Dim iIndent     As Long

    '~~ Additional info about the error line in case one had been provided
    If sErrLine = vbNullString Or sErrLine = "0" Then
        sIndicate = vbNullString
    Else
        sIndicate = " (at line " & sErrLine & ")"
    End If
    sTitle = sTitle & sIndicate
        
    '~~ Path from the entry procedure (the first which uses BoP/EoP)
    '~~ all the way down to the procedure in which the error occoured.
    '~~ When the call stack had not been maintained the path is empty.
    If Not CallStack Is Nothing Then
        If Not CallStack.ErrorPath = vbNullString Then
            CallStack.TraceEndTime = Now()
            CallStack.StackUnwind
        End If
    End If
    
    For i = cllErrPath.Count To 1 Step -1
        If i = cllErrPath.Count Then
            sErrPath = cllErrPath(i) & vbLf
        ElseIf i = 1 Then
            sErrPath = sErrPath & mBasic.Space((iIndent) * 2) & "|_" & cllErrPath(i) & sIndicate
        Else
            sErrPath = sErrPath & mBasic.Space((iIndent) * 2) & "|_" & cllErrPath(i) & vbLf
        End If
        iIndent = iIndent + 1
    Next i
    '~~ Prepare the Title with the error number and the procedure which caused the error
    Select Case lErrNo
        Case Is > 0:    sTitle = "VBA Error " & lErrNo
        Case Is < 0:    sTitle = "Application Error " & AppErr(lErrNo)
    End Select
    sTitle = sTitle & " in:  " & sErrSrc & sIndicate
         
    '~~ Consider the error description may include an additional information about the error
    '~~ possible only when the error is raised by Err.Raise
    If InStr(sErrDesc, DCONCAT) <> 0 Then
        sErrText = Split(sErrDesc, DCONCAT)(0)
        sErrInfo = Split(sErrDesc, DCONCAT)(1)
    Else
        sErrText = sErrDesc
        sErrInfo = vbNullString
    End If
                       
#If AlternateMsgBox Then
    '~~ Display the error message by means of the Common UserForm fMsg
    ErrMsgAlternate _
       errnumber:=lErrNo, _
       errsource:=sErrSrc, _
       errdescription:=sErrText, _
       errline:="", _
       errtitle:=sTitle, _
       errpath:=sErrPath, _
       errinfo:=sErrInfo
#Else
    '~~ Assemble error message to be displayed by MsgBox
    ErrMsgBox _
       errnumber:=lErrNo, _
       errsource:=sErrSrc, _
       errdescription:=sErrDesc, _
       errline:=sErrLine, _
       errpath:=sErrPath
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

Private Function ErrorDetails(ByVal lErrNo As Long, _
                              ByVal sErrLine As String) As String
' -----------------------------------------------------------------
' Returns kind of error, error number, and error line if available.
' -----------------------------------------------------------------
Dim s As String
    If lErrNo < 0 Then
        s = "App error " & AppErr(lErrNo)
    Else
        s = "VB error " & lErrNo
    End If
    If sErrLine <> 0 Then
        s = s & " at line " & sErrLine
    End If
    ErrorDetails = s
End Function

#If AlternateMsgBox Then
' Elaborated error message using fMsg which supports the display of
' up to 3 message sections, optionally monospaced (here used for the
' error path) and each optionally with a label (here used to specify
' the message sections).
' Note: The error title is automatically assembled.
' -------------------------------------------------------------------
Public Sub ErrMsgAlternate(Optional ByVal errnumber As Long = 0, _
                  Optional ByVal errsource As String = vbNullString, _
                  Optional ByVal errdescription As String = vbNullString, _
                  Optional ByVal errline As String = vbNullString, _
                  Optional ByVal errtitle As String = vbNullString, _
                  Optional ByVal errpath As String = vbNullString, _
                  Optional ByVal errinfo As String = vbNullString)

    Const PROC      As String = "ErrMsgAlternate"
    Dim sIndicate   As String
    Dim sErrText    As String   ' May be a first part of the errdescription
    
    If errnumber = 0 _
    Then MsgBox "Apparently there is no exit statement line above the error handling! Error number is 0!", vbCritical, "Application error in " & ErrSrc(PROC) & "!"

    '~~ Error line info in case one had been provided - additionally integrated in the assembled error title
    If errline = vbNullString Or errline = "0" Then
        sIndicate = vbNullString
    Else
        sIndicate = " (at line " & errline & ")"
    End If

    If errtitle = vbNullString Then
        '~~ When no title is provided, one is assembled by the provided info
        errtitle = errtitle & sIndicate
        '~~ Distinguish between VBA and Application error
        Select Case errnumber
            Case Is > 0:    errtitle = "VBA Error " & errnumber
            Case Is < 0:    errtitle = "Application Error " & AppErr(errnumber)
        End Select
        errtitle = errtitle & " in:  " & errsource & sIndicate
    End If

    If errinfo = vbNullString Then
        '~~ When no error information is provided one may be within the error description
        '~~ which is only possible with an application error raised by Err.Raise
        If InStr(errdescription, "||") <> 0 Then
            sErrText = Split(errdescription, "||")(0)
            errinfo = Split(errdescription, "||")(1)
        Else
            sErrText = errdescription
            errinfo = vbNullString
        End If
    Else
        sErrText = errdescription
    End If

    '~~ Display error message by UserForm fErrMsg
    With fMsg
        .MsgTitle = errtitle
        .MsgLabel(1) = "Error Message/Description:"
        .MsgText(1) = sErrText
        If errpath <> vbNullString Then
            .MsgLabel(2) = "Error path (call stack):"
            .MsgText(2) = errpath
            .MsgMonoSpaced(2) = True
        End If
        If errinfo <> vbNullString Then
            .MsgLabel(3) = "Info:"
            .MsgText(3) = errinfo
        End If
        .MsgButtons = vbOKOnly
        
        '~~ Setup prior activating/displaying the message form is essential!
        '~~ To aviod flickering, the whole setup process must be done before the form is displayed.
        '~~ This  m u s t  be the method called after passing the arguments and before .show
        .Setup
        .Show
    End With

End Sub

#Else
Public Sub ErrMsgBox(ByVal errnumber As Long, _
                     ByVal errsource As String, _
                     ByVal errdescription As String, _
                     ByVal errline As String, _
            Optional ByVal errpath As String = vbNullString)
' ----------------------------------------------------------
' Error message displayed by means of the VBA MsgBox.
' ----------------------------------------------------------
    Const PROC          As String = "ErrMsgBox"
    Dim sMsg            As String
    Dim sMsgTitle       As String
    Dim sDescription    As String
    Dim sInfo           As String

    If errnumber = 0 _
    Then MsgBox "Exit statement before error handling missing! Error number is 0!", vbCritical, "Application error in " & ErrSrc(PROC) & "!"

    '~~ Prepare Title
    If errnumber < 0 Then
        sMsgTitle = "Application Error " & AppErr(errnumber)
    Else
        sMsgTitle = "VB Error " & errnumber
    End If
    sMsgTitle = sMsgTitle & " in " & errsource
    If errline <> 0 Then sMsgTitle = sMsgTitle & " (at line " & errline & ")"

    '~~ Prepare message
    If InStr(errdescription, "||") <> 0 Then
        '~~ Split error description/message and info
        sDescription = Split(errdescription, "||")(0)
        sInfo = Split(errdescription, "||")(1)
    Else
        sDescription = errdescription
    End If
    sMsg = "Description: " & vbLf & sDescription & vbLf & vbLf & _
           "Source:" & vbLf & errsource
    If errline <> 0 Then sMsg = sMsg & " (at line " & errline & ")"
    If errpath <> vbNullString Then
        sMsg = sMsg & vbLf & vbLf & _
               "Path:" & vbLf & errpath
    End If
    If sInfo <> vbNullString Then
        sMsg = sMsg & vbLf & vbLf & _
               "Info:" & vbLf & sInfo
    End If
    MsgBox sMsg, vbCritical, sMsgTitle

End Sub
#End If

Private Function ErrSrc(ByVal sProc As String) As String
    ErrSrc = ThisWorkbook.Name & ">mErrHndlr" & ">" & sProc
End Function
