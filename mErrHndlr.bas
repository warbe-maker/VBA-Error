Attribute VB_Name = "mErrHndlr"
Option Explicit
Option Private Module
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
'                     The local Conditional Compile Argument "AlternativeMsgBox = 1" enforces the use
'                     of the Alternative VBA MsgBox which provideds an improved readability.
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
'             mErrHndlr.ErrHndlr Err.Number, ErrSrc(PROC), Err.Description, Erl
'           End ....
'
' Note: When the call stack is not maintained the ErrHndlr will display the message
'       immediately with the procedure the error occours. When the call stack is
'       maintained, the error message will display the call path to the error beginning
'       with the first (entry) procedure in which the call stack is maintained all the
'       call sequence down to the procedure where the error occoured.
'
' Uses: fMsg (only when the Conditional Compile Argument AlternativeMsgBox = 1)
'
' Requires: Reference to "Microsoft Scripting Runtime"
'
' W. Rauschenberger, Berlin January 2020, https://github.com/warbe-maker/Common-VBA-Error-Handler
' -----------------------------------------------------------------------------------------------
' --- Begin of declaration for the Execution Tracing
Private Declare PtrSafe Function getFrequency Lib "kernel32" _
Alias "QueryPerformanceFrequency" (cyFrequency As Currency) As Long
Private Declare PtrSafe Function getTickCount Lib "kernel32" _
Alias "QueryPerformanceCounter" (cyTickCount As Currency) As Long

Private Const TRACE_BEGIN_ID        As String = ">"                   ' Begin procedure or code trace indicator
Private Const TRACE_END_ID          As String = "<"                   ' End procedure or code trace indicator
Private Const EXEC_TRACE_APP_ERR    As String = "Application error "
Private Const EXEC_TRACE_VB_ERR     As String = "VB Runtime error "
Private Const EXEC_TRACE_DB_ERROR   As String = "Database error "
Private Const TRACE_COMMENT         As String = " !!! "
Public Const CONCAT                 As String = "||"

Private cyFrequency         As Currency     ' Execution Trace Frequency (initialized with init)
Private cyTicks             As Currency     ' Execution Trace Ticks counter
Private iTraceItem          As Long         ' Execution Trace Call counter to unify key
Private lPrecisionDecimals  As Long         ' Execution Trace Default Precision (6=0,000000)
Private iSec                As Integer      ' Execution Trace digits left from decimal point
Private iDec                As Integer      ' Execution Trace decimal digits right from decimal point
Private sFormat             As String       ' Execution Trace tracking time presentation format
Private cyOverhead          As Currency     ' Execution Trace time accumulated by caused by the time tracking itself
Private dtTraceBeginTime    As Date         ' Execution Trace start time
' --- End of declaration for the Execution Tracing

Public Enum StartupPosition         ' ---------------------------
    Manual = 0                      ' Used to position the
    CenterOwner = 1                 ' final setup message form
    CenterScreen = 2                ' horizontally and vertically
    WindowsDefault = 3              ' centered on the screen
End Enum                            ' ---------------------------

Private Enum enErrorType
    VBRuntimeError
    ApplicationError
    DatabaseError
End Enum

Public Type tSection                ' ------------------
       sLabel As String             ' Structure of the
       sText As String              ' UserForm's
       bMonspaced As Boolean        ' message area which
End Type                            ' consists of
Public Type tMessage                ' three message
       section(1 To 3) As tSection  ' sections
End Type                            ' -------------------

Private dicTrace            As Dictionary       ' Procedure execution trancing records
Private cllErrPath          As Collection
Private cllErrorPath        As Collection   ' managed by ErrPath... procedures exclusively
Private dctStack            As Dictionary
Private sErrHndlrEntryProc  As String
Private lSubsequErrNo       As Long ' a number possibly different from lInitialErrNo when it changes when passed on to the Entry Procedure
Private lInitialErrLine     As Long
Private lInitialErrNo       As Long
Private sInitialErrSource   As String
Private sInitialErrDscrptn  As String

Private Property Get TRACE_PROC_BEGIN_ID() As String:   TRACE_PROC_BEGIN_ID = TRACE_BEGIN_ID & TRACE_BEGIN_ID & " ":                        End Property
Private Property Get TRACE_PROC_END_ID() As String:     TRACE_PROC_END_ID = TRACE_END_ID & TRACE_END_ID & " ":                              End Property
Private Property Get TRACE_CODE_BEGIN_ID() As String:   TRACE_CODE_BEGIN_ID = TRACE_BEGIN_ID & " ":                                         End Property
Private Property Get TRACE_CODE_END_ID() As String:     TRACE_CODE_END_ID = TRACE_END_ID & " ":                                             End Property
Private Property Get INCOMPLETE_TRACE() As String:      INCOMPLETE_TRACE = TRACE_COMMENT & "Incomplete trace" & TRACE_COMMENT:              End Property

' Test button, displayed with Conditional Compile Argument Test = 1
Public Property Get ExitAndContinue() As String:        ExitAndContinue = "Exit procedure" & vbLf & "and continue" & vbLf & "with next":    End Property

' Debugging button, displayed with Conditional Compile Argument Debugging = 1
Public Property Get ResumeError() As String:            ResumeError = "Resume" & vbLf & "error code line":                                  End Property

' Test button, displayed with Conditional Compile Argument Test = 1
Public Property Get ResumeNext() As String:             ResumeNext = "Continue with code line" & vbLf & "following the error line":         End Property

' Default error message button
Public Property Get ErrMsgDefaultButton() As String:    ErrMsgDefaultButton = "Terminate execution":                                                  End Property

Private Property Get StackEntryProc() As String
    If Not StackIsEmpty _
    Then StackEntryProc = dctStack.Items()(0) _
    Else StackEntryProc = vbNullString
End Property

Public Function AppErr(ByVal errno As Long) As Long
' -----------------------------------------------------------------
' Used with Err.Raise AppErr(<l>).
' When the error number <l> is > 0 it is considered an "Application
' Error Number and vbObjectErrror is added to it into a negative
' number in order not to confuse with a VB runtime error.
' When the error number <l> is negative it is considered an
' Application Error and vbObjectError is added to convert it back
' into its origin positive number.
' ------------------------------------------------------------------
    If errno < 0 Then
        AppErr = errno - vbObjectError
    Else
        AppErr = vbObjectError + errno
    End If
End Function

Public Sub BoP(ByVal s As String)
    StackPush s
End Sub

Public Sub BoT(ByVal s As String)
' -------------------------------
' Begin of Trace for code lines.
' -------------------------------
#If ExecTrace Then
    TrcBegin s, TRACE_CODE_BEGIN_ID
#End If
End Sub

Public Sub EoP(ByVal s As String)
' --------------------------------
' Trace and stack End of Procedure
' --------------------------------

    StackPop s
    
    If StackIsEmpty Or s = sErrHndlrEntryProc Then
        ErrHndlrDisplayTrace
    End If
 
End Sub

Public Sub EoT(ByVal s As String)
' -------------------------------
' End of Trace for code lines.
' -------------------------------
#If ExecTrace Then
    TrcEnd s, TRACE_CODE_END_ID
#End If
End Sub

            
Public Function ErrHndlr(ByVal errnumber As Long, _
                         ByVal errsource As String, _
                         ByVal errdscrptn As String, _
                         ByVal errline As Long, _
                Optional ByVal buttons As Variant = vbNullString) As Variant
' ----------------------------------------------------------------------
' When the buttons argument specifies more than one button the error
'      message is immediately displayed and the users choice is returned
' else when the caller (errsource) is the "Entry Procedure" the error
'      is displayed with the path to the error
' else the error is passed on to the "Entry Procedure" whereby the
' .ErrorPath string is assebled.
' ----------------------------------------------------------------------
    
    Static sLine    As String   ' provided error line (if any) for the the finally displayed message
    Dim sTrace      As String
    
    If ErrHndlrFailed(errnumber, errsource, buttons) Then Exit Function
    If cllErrPath Is Nothing Then Set cllErrPath = New Collection
    If errline <> 0 Then sLine = errline Else sLine = "0"
    ErrHndlrManageButtons buttons

    If sInitialErrSource = vbNullString Then
        '~~ This is the initial/first execution of the error handler within the error raising procedure.
        TrcError errsource & TRACE_COMMENT & ErrorDetails(errnumber, errsource, errline) & TRACE_COMMENT
        StackPop errsource, trace:=False
        lInitialErrLine = errline
        lInitialErrNo = errnumber
        sInitialErrSource = errsource
        sInitialErrDscrptn = errdscrptn
    ElseIf errnumber <> lInitialErrNo _
        And errnumber <> lSubsequErrNo _
        And errsource <> sInitialErrSource Then
        '~~ In the rare case when the error number had changed during the process of passing it back up to the entry procedure
        lSubsequErrNo = errnumber
        TrcError errsource & TRACE_COMMENT & ErrorDetails(lSubsequErrNo, errsource, errline) & " """ & errdscrptn & """"
    End If
    
    '~~ When the user has no choice to press any button but the only one displayed button
    '~~ and the Entry Procedure is known but yet not reached the path back up to the Entry Procedure
    '~~ is maintained and the error is passed on to the caller
    If ErrorButtons(buttons) = 1 _
    And sErrHndlrEntryProc <> vbNullString _
    And StackEntryProc <> errsource Then
        ErrPathAdd errsource
        StackPop errsource
#If ExecTrace Then
        If errsource <> sInitialErrSource Then TrcEnd errsource, TRACE_PROC_END_ID
#End If
        Err.Raise errnumber, errsource, errdscrptn
    End If
    
    If ErrorButtons(buttons) > 1 _
    Or StackEntryProc = errsource _
    Or StackEntryProc = vbNullString Then
        ' More than one button is displayed with the error message for the user to choose one
        ' or the Entry Procedure is unknown
        ' or has been reached
        If Not ErrPathIsEmpty Then ErrPathAdd errsource
        StackPop errsource
'#If ExecTrace Then
'        TrcEnd errsource, TRACE_PROC_END_ID
'#End If
        '~~ Display the error message
        ErrHndlr = ErrMsg(errnumber:=lInitialErrNo, errsource:=sInitialErrSource, errdscrptn:=sInitialErrDscrptn, errline:=lInitialErrLine, buttons:=buttons)
        Select Case ErrHndlr
            Case ResumeError, ResumeNext, ExitAndContinue
            Case Else: ErrPathErase
        End Select
        lInitialErrNo = 0
        sInitialErrSource = vbNullString
        sInitialErrDscrptn = vbNullString

    End If
    
    '~~ Each time a known Entry Procedure is reached the execution trace
    '~~ maintained by the BoP and EoP and the BoT and EoT statements is displayed
    If StackEntryProc = errsource _
    Or StackEntryProc = vbNullString Then
        ErrHndlrDisplayTrace
        Select Case ErrHndlr
            Case ResumeError, ResumeNext, ExitAndContinue
            Case vbOK
            Case Else: StackErase
        End Select
    End If
            
End Function

Private Sub ErrHndlrDisplayTrace()
' --------------------------------------
' Display the trace in the fMsg UserForm
' --------------------------------------
#If ExecTrace Then
        With fMsg
            .MaxFormWidthPrcntgOfScreenSize = 95
            .MsgTitle = "Execution Trace, displayed because the Conditional Compile Argument ""ExecTrace = 1""!"
            .MsgText(1) = TrcDsply:   .MsgMonoSpaced(1) = True
            .Setup
            .Show
        End With
#End If

End Sub
Private Sub ErrHndlrManageButtons(ByRef buttons As Variant)

    If buttons = vbNullString _
    Then buttons = ErrMsgDefaultButton _
    Else ErrHndlrAddButtons ErrMsgDefaultButton, buttons ' add the default button before the buttons specified
    
'~~ Special features are only available with the Alternative VBA MsgBox
#If Debugging Or Test Then
    ErrHndlrAddButtons buttons, vbLf ' buttons in new row
#End If
#If Debugging Then
    ErrHndlrAddButtons buttons, ResumeError
#End If
#If Test Then
     ErrHndlrAddButtons buttons, ResumeNext
     ErrHndlrAddButtons buttons, ExitAndContinue
#End If

End Sub
Private Function ErrHndlrFailed( _
        ByVal errnumber As Long, _
        ByVal errsource As String, _
        ByVal buttons As Variant) As Boolean
' ------------------------------------------
'
' ------------------------------------------

    If errnumber = 0 Then
        MsgBox "The error handling has been called with an error number = 0 !" & vbLf & vbLf & _
               "This indicates that in procedure" & vbLf & _
               ">>>>> " & errsource & " <<<<<" & vbLf & _
               "an ""Exit ..."" statement before the call of the error handling is missing!" _
               , vbExclamation, _
               "Exit ... statement missing in " & errsource & "!"
                ErrHndlrFailed = True
        Exit Function
    End If
    
    If IsNumeric(buttons) Then
        '~~ When buttons is a numeric value, only the VBA MsgBox values for the button argument are supported
        Select Case buttons
            Case vbOKOnly, vbOKCancel, vbYesNo, vbRetryCancel, vbYesNoCancel, vbAbortRetryIgnore
            Case Else
                MsgBox "When the buttons argument is a numeric value Only the valid VBA MsgBox vaulues are supported. " & _
                       "For valid values please refer to:" & vbLf & _
                       "https://docs.microsoft.com/en-us/office/vba/Language/Reference/User-Interface-Help/msgbox-function" _
                       , vbOKOnly, "Only the valid VBA MsgBox vaulues are supported!"
                ErrHndlrFailed = True
                Exit Function
        End Select
    End If

End Function

Private Sub ErrHndlrAddButtons(ByRef v1 As Variant, _
                               ByRef v2 As Variant)
' ---------------------------------------------------
' Returns v1 followed by v2 whereby both may be a
' buttons argument which means  a string, a
' Dictionary or a Collection. When v1 is a Dictionary
' or Collection v2 must be a string or long and vice
' versa.
' ---------------------------------------------------
    
    Dim dct As New Dictionary
    Dim cll As New Collection
    Dim v   As Variant
    
    Select Case TypeName(v1)
        Case "Dictionary"
            Select Case TypeName(v2)
                Case "String", "Long": v1.Add v2, v2
                Case Else ' Not added !
            End Select
        Case "Collection"
            Select Case TypeName(v2)
                Case "String", "Long": v1.Add v2
                Case Else ' Not added !
            End Select
        Case "String", "Long"
            Select Case TypeName(v2)
                Case "String"
                    v1 = v1 & "," & v2
                Case "Dictionary"
                    dct.Add v1, v1
                    For Each v In v2
                        dct.Add v, v
                    Next v
                    Set v2 = dct
                Case "Collection"
                    cll.Add v1
                    For Each v In v2
                        cll.Add v
                    Next v
                    Set v2 = cll
            End Select
    End Select
    
End Sub

Public Function ErrMsg( _
                ByVal errnumber As Long, _
                ByVal errsource As String, _
                ByVal errdscrptn As String, _
                ByVal errline As Long, _
       Optional ByVal buttons As Variant = vbOKOnly) As Variant
' -------------------------------------------------------------
' Displays the error message either by means of VBA MsgBox or,
' when the Conditional Compile Argument AlternativeMsgBox = 1 by
' means of the Alternative VBA MsgBox (UserForm fMsg). In any
' case the path to the error may be displayed, provided the
' entry procedure has BoP/EoP code lines.
'
' W. Rauschenberger, Berlin, Sept 2020
' -------------------------------------------------------------
    
    Dim sErrPath    As String:      sErrPath = ErrPathErrMsg
    Dim sTitle      As String
    
    Select Case ErrorType(errnumber, errsource)
        Case VBRuntimeError, DatabaseError:         sTitle = ErrorTypeString(ErrorType(errnumber, errsource)) & errnumber & " in " & errsource & ErrorLine(errline)
        Case ApplicationError:                      sTitle = ErrorTypeString(ErrorType(errnumber, errsource)) & AppErr(errnumber) & " in " & errsource & ErrorLine(errline)
    End Select
    
    '~~ Display the error message by means of the Common UserForm fMsg
    With fMsg
        .MsgTitle = sTitle
        .MsgLabel(1) = "Error description:":            .MsgText(1) = ErrorDescription(sInitialErrDscrptn)
        If Not ErrPathIsEmpty Then
            .MsgLabel(2) = "Error path (call stack):":  .MsgText(2) = sErrPath:   .MsgMonoSpaced(2) = True
        Else
            .MsgLabel(2) = "Error source:":             .MsgText(2) = sInitialErrSource & ErrorLine(lInitialErrLine)
        End If
        .MsgLabel(3) = "Info:":                         .MsgText(3) = ErrorInfo(errdscrptn)
        .MsgButtons = buttons
        .Setup
        .Show
        If ErrorButtons(buttons) = 1 Then
            ErrMsg = buttons ' a single reply buttons return value cannot be obtained since the form is unloaded with its click
        Else
            ErrMsg = .ReplyValue ' when more than one button is displayed the form is unloadhen the return value is obtained
        End If
    End With

End Function

Private Function ErrorButtons( _
                 ByVal buttons As Variant) As Long
' ------------------------------------------------
' Returns the number of specified buttons.
' ------------------------------------------------
    Dim v As Variant
    
    For Each v In Split(buttons, ",")
        If IsNumeric(v) Then
            Select Case v
                Case vbOKOnly:                              ErrorButtons = ErrorButtons + 1
                Case vbOKCancel, vbYesNo, vbRetryCancel:    ErrorButtons = ErrorButtons + 2
                Case vbAbortRetryIgnore, vbYesNoCancel:     ErrorButtons = ErrorButtons + 3
            End Select
        Else
            Select Case v
                Case vbNullString, vbLf, vbCr, vbCrLf
                Case Else:  ErrorButtons = ErrorButtons + 1
            End Select
        End If
    Next v

End Function

Private Function ErrorDescription(ByVal s As String) As String
' ------------------------------------------------------------
' Returns the string which follows a "||" in the error
' description which indicates an additional information
' regarding the error.
' ------------------------------------------------------------
    If InStr(s, CONCAT) <> 0 _
    Then ErrorDescription = Split(s, CONCAT)(0) _
    Else ErrorDescription = s
End Function

Private Function ErrorDetails( _
                 ByVal errnumber As Long, _
                 ByVal errsource As String, _
                 ByVal sErrLine As String) As String
' --------------------------------------------------
' Returns the kind of error, the error number, and
' the error line (if available) as string.
' --------------------------------------------------
    
    Select Case ErrorType(errnumber, errsource)
        Case ApplicationError:              ErrorDetails = ErrorTypeString(ErrorType(errnumber, errsource)) & AppErr(errnumber)
        Case DatabaseError, VBRuntimeError: ErrorDetails = ErrorTypeString(ErrorType(errnumber, errsource)) & errnumber
    End Select
        
    If sErrLine <> 0 Then ErrorDetails = ErrorDetails & ErrorLine(lInitialErrLine)

End Function

Private Function ErrorInfo(ByVal s As String) As String
' -----------------------------------------------------
' Returns the string which follows a "||" in the error
' description which indicates an additional information
' regarding the error.
' -----------------------------------------------------
    If InStr(s, CONCAT) <> 0 _
    Then ErrorInfo = Split(s, CONCAT)(1) _
    Else ErrorInfo = vbNullString
End Function

Private Function ErrorLine( _
                 ByVal errline As Long) As String
' -----------------------------------------------
' Returns a complete errol line message.
' -----------------------------------------------
    If errline <> 0 _
    Then ErrorLine = " (at line " & errline & ")" _
    Else ErrorLine = vbNullString
End Function

Private Function ErrorTypeString(ByVal errtype As enErrorType) As String
    Select Case errtype
        Case ApplicationError:  ErrorTypeString = EXEC_TRACE_APP_ERR
        Case DatabaseError:     ErrorTypeString = EXEC_TRACE_DB_ERROR
        Case VBRuntimeError:    ErrorTypeString = EXEC_TRACE_VB_ERR
    End Select
End Function

Private Function ErrorType( _
                 ByVal errnumber As Long, _
                 ByVal errsource As String) As enErrorType
' --------------------------------------------------------
' Return the kind of error considering the error source
' (errsource) and the error number (errnumber).
' --------------------------------------------------------

   If InStr(1, errsource, "DAO") <> 0 _
   Or InStr(1, errsource, "ODBC Teradata Driver") <> 0 _
   Or InStr(1, errsource, "ODBC") <> 0 _
   Or InStr(1, errsource, "Oracle") <> 0 Then
      ErrorType = DatabaseError
   Else
      If errnumber > 0 _
      Then ErrorType = VBRuntimeError _
      Else ErrorType = ApplicationError
   End If
   
End Function

Private Sub ErrPathAdd(ByVal s As String)
    
    If cllErrorPath Is Nothing Then Set cllErrorPath = New Collection _

    If Not ErrPathItemExists(s) Then
        Debug.Print s & " added to path"
        cllErrorPath.Add s ' avoid duplicate recording of the same procedure/item
    End If
End Sub

Private Function ErrPathItemExists(ByVal s As String) As Boolean

    Dim v As Variant
    
    For Each v In cllErrorPath
        If InStr(v & " ", s & " ") <> 0 Then
            ErrPathItemExists = True
            Exit Function
        End If
    Next v
    
End Function

Private Sub ErrPathErase()
    Set cllErrorPath = Nothing
End Sub

Private Function ErrPathErrMsg() As String
' ----------------------------------------
' Returns the error path for being
' displayed in the error message.
' ----------------------------------------
    
    Dim i   As Long:    i = 0
    Dim j   As Long:    j = 0
    Dim s   As String
    
    ErrPathErrMsg = vbNullString
    If Not ErrPathIsEmpty Then
        '~~ When the error path is not empty and not only contains the error source procedure
        For i = cllErrorPath.Count To 1 Step -1
            s = cllErrorPath.Item(i)
            If i = cllErrorPath.Count _
            Then ErrPathErrMsg = s _
            Else ErrPathErrMsg = ErrPathErrMsg & vbLf & Space(j * 2) & "|_" & s
            j = j + 1
        Next i
    End If
    ErrPathErrMsg = ErrPathErrMsg & vbLf & Space(j * 2) & "|_" & sInitialErrSource & " " & ErrorDetails(lInitialErrNo, sInitialErrSource, lInitialErrLine)

End Function

Private Function ErrPathIsEmpty() As Boolean
    ErrPathIsEmpty = cllErrorPath Is Nothing
    If Not ErrPathIsEmpty Then ErrPathIsEmpty = cllErrorPath.Count = 0
End Function

Private Function ErrSrc(ByVal sProc As String) As String
    ErrSrc = ThisWorkbook.Name & ">mErrHndlr" & ">" & sProc
End Function

Private Function Replicate(ByVal s As String, _
                        ByVal ir As Long) As String
' -------------------------------------------------
' Returns the string (s) repeated (ir) times.
' -------------------------------------------------
    Dim i   As Long
    
    For i = 1 To ir
        Replicate = Replicate & s
    Next i
    
End Function

Public Function Space(ByVal l As Long) As String
' --------------------------------------------------
' Unifies the VB differences SPACE$ and Space$ which
' lead to code diferences where there aren't any.
' --------------------------------------------------
    Space = VBA.Space$(l)
End Function

Private Function StackBottom() As String
    If Not StackIsEmpty Then StackBottom = dctStack.Items()(0)
End Function

Private Sub StackErase()
    If Not dctStack Is Nothing Then dctStack.RemoveAll
End Sub

Private Function StackIsEmpty() As Boolean
    StackIsEmpty = dctStack Is Nothing
    If Not StackIsEmpty Then StackIsEmpty = dctStack.Count = 0
End Function

Private Function StackPop( _
        Optional ByVal s As String = vbNullString, _
        Optional ByVal trace As Boolean = True) As String
' -------------------------------------------------------
' Returns the item removed from the top of the stack.
' When s is provided and is not on the top of the stack
' an error is raised.
' -------------------------------------------------------
    
    On Error GoTo on_error
    Const PROC = "StackPop"
    Dim sPop    As String
    
    If s = StackTop Then
        StackPop = s                    ' Return the poped item
        dctStack.Remove dctStack.Count  ' Remove item s from stack
#If ExecTrace Then
        If trace Then
            TrcEnd s, TRACE_PROC_END_ID
        End If
#End If
    ElseIf s = vbNullString And Not StackIsEmpty Then
        dctStack.Remove dctStack.Count  ' Unwind! Remove item s from stack
#If ExecTrace Then
        If trace Then
            TrcEnd s, TRACE_PROC_END_ID
        End If
#End If
    End If
    
exit_proc:
    Exit Function

on_error:
    MsgBox Err.Description, vbOKOnly, "Error in " & ErrSrc(PROC)
End Function

Private Sub StackPush(ByVal s As String)

    If dctStack Is Nothing Then Set dctStack = New Dictionary
    If dctStack.Count = 0 Then
        sErrHndlrEntryProc = s ' First pushed = bottom item = entry procedure
    End If
    dctStack.Add dctStack.Count + 1, s
#If ExecTrace Then
    TrcBegin s, TRACE_PROC_BEGIN_ID    ' start of the procedure's execution trace
#End If

End Sub

Private Function StackTop() As String
    If Not StackIsEmpty Then StackTop = dctStack.Items()(dctStack.Count - 1)
End Function

Private Function StackUnwind() As String
' --------------------------------------
' Completes the execution trace for
' items still on the stack.
' --------------------------------------
    Dim s As String
    Do Until StackIsEmpty
        s = StackPop
        If s <> vbNullString Then
            TrcEnd s, TrcBegEndId(s)
        End If
    Loop
End Function

Private Function TrcBegEndId(ByVal s As String) As String
    TrcBegEndId = Split(s, " ")(0) & " "
End Function

Public Sub TrcBegin(ByVal s As String, _
                    ByVal id As String)
' --------------------------------------
' Keep a record of the current tick
' count at the begin of the execution of
' the procedure (s) identified as
' procedure or code trace begin (id).
' ---------------------------------------
    Dim cy      As Currency
    
    getTickCount cy
    TrcInit
    
    getTickCount cyTicks
    TrcAdd id & s, cy
    cyOverhead = cyOverhead + (cyTicks - cy)
    
    '~~ Reset a possibly error raised procedure
    sInitialErrSource = vbNullString

End Sub

Private Function TrcBeginLine( _
                 ByVal cyInitial As Currency, _
                 ByVal iTt As Long, _
                 ByVal sIndent As String, _
                 ByVal iIndent As Long, _
                 ByVal sProcName As String, _
                 ByVal sMsg As String) As String
' ----------------------------------------------
' Return a trace begin line for being displayed.
' ----------------------------------------------
    TrcBeginLine = TrcSecs(cyInitial, dicTrace.Items(iTt)) _
                 & "    " & sIndent _
                 & " " & Replicate("|  ", iIndent) _
                 & sProcName & sMsg
End Function

Private Function TrcBeginTicks(ByVal s As String, _
                               ByVal i As Single) As Currency
' -----------------------------------------------------------
' Returns the number of ticks recorded with the begin item
' corresponding with the end item (s) by searching the trace
' Dictionary back up starting with the index (i) -1 (= index
' of the end time (s)).
' Returns 0 when no start item coud be found.
' To avoid multiple identifications of the begin item it is
' set to vbNullString with the return of the number of begin ticks.
' ----------------------------------------------------------
    
    Dim j       As Single
    Dim sItem   As String
    Dim sKey    As String

    TrcBeginTicks = 0
    s = Replace(s, TRACE_END_ID, TRACE_BEGIN_ID)  ' turn the end item into a begin item string
    For j = i - 1 To 0 Step -1
        sKey = Split(TrcUnstripItemNo(dicTrace.Keys(j)), TRACE_COMMENT)(0)
        sItem = Split(s, TRACE_COMMENT)(0)
        If sItem = sKey Then
            If dicTrace.Items(j) <> vbNullString Then
                '~~ Return the begin ticks and replace the value by vbNullString
                '~~ to avoid multiple recognition of the same start item
                TrcBeginTicks = dicTrace.Items(j)
                dicTrace.Items(j) = vbNullString
                Exit For
            End If
        End If
    Next j
    
End Function

Public Function TrcDsply() As String
' -------------------------------------------------
' Returns the execution trace a displayable string.
' -------------------------------------------------
#If ExecTrace Then
    
    On Error GoTo on_error
    Const PROC = "TrcDsply"        ' This procedure's name for the error handling and execution tracking
    Const ELAPSED = "Elapsed"
    
    Dim cyStrt      As Currency ' ticks count at start
    Dim cyEnd       As Currency ' ticks count at end
    Dim cyElapsed   As Currency ' elapsed ticks since start
    Dim cyInitial   As Currency ' initial ticks count (at first traced proc)
    Dim iTt         As Single   ' index for dictionary dicTrace
    Dim sProcName   As String   ' tracked procedure/vba code
    Dim iIndent     As Single   ' indentation nesting level
    Dim sIndent     As String   ' Indentation string defined by the precision
    Dim sMsg        As String
    Dim dbl         As Double
    Dim i           As Long
    Dim sTrace      As String
    Dim sTraceLine  As String
    Dim sInfo       As String
       
    StackUnwind ' end procedures still on the stack
    
    If dicTrace Is Nothing Then Exit Function   ' When the contional compile argument where not
    If dicTrace.Count = 0 Then Exit Function    ' ExecTrace = 1 there will be no execution trace result

    cyElapsed = 0
    
    If lPrecisionDecimals = 0 Then lPrecisionDecimals = 6
    iDec = lPrecisionDecimals
    cyStrt = dicTrace.Items(0)
    For i = dicTrace.Count - 1 To 0 Step -1
        cyEnd = dicTrace.Items(i)
        If cyEnd <> 0 Then Exit For
    Next i
    
    If cyFrequency = 0 Then getFrequency cyFrequency
    dbl = (cyEnd - cyStrt) / cyFrequency
    Select Case dbl
        Case Is >= 100000:  iSec = 6
        Case Is >= 10000:   iSec = 5
        Case Is >= 1000:    iSec = 4
        Case Is >= 100:     iSec = 3
        Case Is >= 10:      iSec = 2
        Case Else:          iSec = 1
    End Select
    sFormat = String$(iSec - 1, "0") & "0." & String$(iDec, "0") & " "
    sIndent = Space$(Len(sFormat))
    iIndent = -1
    
    '~~ Header
    sTrace = ELAPSED & VBA.Space$(Len(sIndent) - Len(ELAPSED) + 1) & "Exec secs" & " >> Begin execution trace " & Format(dtTraceBeginTime, "hh:mm:ss") & " (exec time in seconds)"
    
    '~~ Exec trace lines
    For iTt = 0 To dicTrace.Count - 1
        sProcName = dicTrace.Keys(iTt)
        If TrcIsEndItem(sProcName, sInfo) Then
            '~~ Trace End Line
            cyEnd = dicTrace.Items(iTt)
            cyStrt = TrcBeginTicks(sProcName, iTt)   ' item is set to vbNullString to avoid multiple recognition
            If cyStrt = 0 Then
                '~~ The corresponding BoP/BoT entry for a EoP/EoT entry couldn't be found within the trace
                iIndent = iIndent + 1
                sTraceLine = Space$((Len(sFormat) * 2) + 1) & "    " & Replicate("|  ", iIndent) & sProcName
                If InStr(sTraceLine, TRACE_PROC_END_ID) <> 0 _
                Then sTraceLine = sTraceLine & INCOMPLETE_TRACE & "The corresponding BoP statement for a EoP statement is missing." _
                Else sTraceLine = sTraceLine & INCOMPLETE_TRACE & "The corresponding BoT statement for a EoT statement is missing."
                sTrace = sTrace & vbLf & sTraceLine
                iIndent = iIndent - 1
            Else
                '~~ End line
                sTraceLine = TrcEndLine(cyInitial, cyEnd, cyStrt, iIndent, sProcName)
                sTrace = sTrace & vbLf & sTraceLine
                iIndent = iIndent - 1
            End If
        ElseIf TrcIsBegItem(sProcName) Then
            '~~ Trace Begin Line
            iIndent = iIndent + 1
            If iTt = 0 Then cyInitial = dicTrace.Items(iTt)
            sMsg = TrcEndItemMissing(sProcName)
            
            sTraceLine = TrcBeginLine(cyInitial, iTt, sIndent, iIndent, sProcName, sMsg)
            sTrace = sTrace & vbLf & sTraceLine
            If sMsg <> vbNullString Then iIndent = iIndent - 1
        
        End If
        sInfo = vbNullString
    Next iTt
    
    dicTrace.RemoveAll
    '~~ Footer
    sTraceLine = Space$((Len(sFormat) * 2) + 2) & "<< End execution trace " & Format(Now(), "hh:mm:ss") & " (only " & Format(TrcSecs(0, cyOverhead), "0.000000") & " seconds exec time were caused by the executuion trace itself)"
    sTrace = sTrace & vbLf & sTraceLine
    
exit_proc:
    TrcDsply = sTrace
    Exit Function
    
on_error:
    MsgBox Err.Description, vbOKOnly, "Error in " & ErrSrc(PROC)
#End If
End Function

Private Sub TrcAdd(ByVal s As String, _
                   ByVal cy As Currency)
                   
    iTraceItem = iTraceItem + 1
    dicTrace.Add iTraceItem & s, cy
'    Debug.Print "Added to Trace: '" & iTraceItem & s
    
End Sub

Private Function TrcItem(ByVal s As String) As String
' ---------------------------------------------------
' Returns the item (i.e. the traced ErrSrc()) element
' within the trace entry.
' Precondition: The ErrSrc() must not contain spaces
' ---------------------------------------------------
    TrcItem = Split(s, " ")(1)
End Function

Private Sub TrcEnd(ByVal s As String, _
                   ByVal id As String)
' -----------------------------------------
' End of Trace. Keeps a record of the ticks
' count for the execution trace of the
' group of code lines named (s).
' -----------------------------------------
#If ExecTrace Then
    
    On Error GoTo on_error
    Const PROC = "TrcEnd"
    Dim cy      As Currency
    
    getTickCount cyTicks
    cy = cyTicks
    TrcAdd id & s, cyTicks
    getTickCount cyTicks
    cyOverhead = cyOverhead + (cyTicks - cy)

exit_proc:
    Exit Sub
    
on_error:
    MsgBox Err.Description, vbOKOnly, "Error in " & ErrSrc(PROC)
#End If
End Sub

Private Function TrcEndItemMissing(ByVal s As String) As String
' -------------------------------------------------------------------
' Returns a message string when a corresponding end item is missing.
' -------------------------------------------------------------------
    Dim i       As Long
    Dim sKey    As String
    Dim sInfo   As String

    TrcEndItemMissing = "missing"
    s = Replace(s, TRACE_BEGIN_ID, TRACE_END_ID)  ' turn the end item into a begin item string
    For i = 0 To dicTrace.Count - 1
        sKey = dicTrace.Keys(i)
        If TrcIsEndItem(sKey, sInfo) Then
            If TrcItem(sKey) = TrcItem(s) Then
                TrcEndItemMissing = vbNullString
                GoTo exit_proc
            End If
        End If
    Next i
    
exit_proc:
    If TrcEndItemMissing <> vbNullString Then
        If Split(s, " ")(0) & " " = TRACE_PROC_END_ID _
        Then TrcEndItemMissing = INCOMPLETE_TRACE & "The corresponding EoP statement for a BoP statement is missing." _
        Else TrcEndItemMissing = INCOMPLETE_TRACE & "The corresponding EoT statement for a BoT statement is missing."
    End If
End Function

Private Function TrcEndLine( _
                 ByVal cyInitial As Currency, _
                 ByVal cyEnd As Currency, _
                 ByVal cyStrt As Currency, _
                 ByVal iIndent As Long, _
                 ByVal sProcName As String) As String
' ---------------------------------------------------
' Assemble a Trace End Line
' ---------------------------------------------------
    
    TrcEndLine = TrcSecs(cyInitial, cyEnd) & " " & _
                 TrcSecs(cyStrt, cyEnd) & "    " & _
                 Replicate("|  ", iIndent) & _
                 sProcName

End Function

Public Sub TrcError(ByVal s As String)
' --------------------------------------
' Keep record of the error (s) raised
' during the execution of any procedure.
' --------------------------------------
#If ExecTrace Then
    Dim cy As Currency

    getTickCount cy
    TrcInit
    
    getTickCount cyTicks
    '~~ Add the error indication line to the trace by ignoring any additional error information
    '~~ optionally attached by two vertical bars
    TrcAdd TRACE_PROC_END_ID & s, cyTicks
    getTickCount cyTicks
    cyOverhead = cyOverhead + (cyTicks - cy)
#End If
End Sub

Private Sub TrcInit()
    If Not dicTrace Is Nothing Then
        If dicTrace.Count = 0 Then
            dtTraceBeginTime = Now()
            iTraceItem = 0
            cyOverhead = 0
        End If
    Else
        Set dicTrace = New Dictionary
        dtTraceBeginTime = Now()
        iTraceItem = 0
        cyOverhead = 0
    End If

End Sub

Private Function TrcIsBegItem(ByRef s As String) As Boolean
' ---------------------------------------------------------
' Returns TRUE if s is an execution trace begin item.
' Returns s with the call counter unstripped.
' ---------------------------------------------------------
Dim i As Single
    TrcIsBegItem = False
    i = InStr(1, s, TRACE_BEGIN_ID)
    If i <> 0 Then
        TrcIsBegItem = True
        s = TrcUnstripItemNo(s)
    End If
End Function

Private Function TrcIsEndItem( _
                 ByRef s As String, _
        Optional ByRef sRight As String) As Boolean
' -------------------------------------------------
' Returns TRUE if s is an execution trace end item.
' Returns s with the item counter unstripped. Any
' additional info is returne in sRight.
' -------------------------------------------------
    
    Dim sIndicator  As String
    
    s = TrcUnstripItemNo(s)
    sIndicator = Split(s)(0)
    Select Case Split(s)(0)
        Case Trim(TRACE_PROC_END_ID), Trim(TRACE_CODE_END_ID)
            TrcIsEndItem = True
        Case Else
            TrcIsEndItem = False
    End Select
        
    If InStr(s, TRACE_COMMENT) <> 0 Then
        sRight = TRACE_COMMENT & Split(s, TRACE_COMMENT)(1)
    Else
        sRight = vbNullString
    End If
    
End Function

Private Function TrcSecs( _
                 ByVal cyStrt As Currency, _
                 ByVal cyEnd As Currency) As String
' --------------------------------------------------
' Returns the difference between cyStrt and cyEnd as
' formatted seconds string (decimal = nanoseconds).
' --------------------------------------------------
    Dim dbl As Double

    dbl = (cyEnd - cyStrt) / cyFrequency
    TrcSecs = Format(dbl, sFormat)

End Function

Private Function TrcUnstripItemNo( _
                 ByVal s As String) As String
    Dim i As Long

    i = 1
    While IsNumeric(Mid(s, i, 1))
        i = i + 1
    Wend
    s = Right(s, Len(s) - (i - 1))
    TrcUnstripItemNo = s
    
End Function

