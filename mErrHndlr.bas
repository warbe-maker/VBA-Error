Attribute VB_Name = "mErrHndlr"
Option Explicit
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
'                     The local Conditional Compile Argument "AlternateMsgBox = 1" enforces the use
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
' Uses: fMsg (only when the Conditional Compile Argument AlternateMsgBox = 1)
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

Const EXEC_TRACE_BEGIN_ID   As String = ">> "       ' Begin item indicator
Const EXEC_TRACE_END_ID     As String = "<< "       ' End item indicator
Const EXEC_TRACE_APP_ERR    As String = "Application error "
Const EXEC_TRACE_VB_ERR     As String = "VB Runtime error "

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

Public Type tSection                ' ------------------
       sLabel As String             ' Structure of the
       sText As String              ' UserForm's
       bMonspaced As Boolean        ' message area which
End Type                            ' consists of

Public Type tMessage                ' three message
       section(1 To 3) As tSection  ' sections
End Type

Public dicTrace             As Dictionary       ' Procedure execution trancing records
Private cllErrPath          As Collection

Private dctStack            As Dictionary
Private sErrHndlrEntryProc  As String
Private lSourceErrorNo      As Long
Private lErrorNumber        As Long ' a number possibly different from lSourceErrorNo when it changes when passed on to the Entry Procedure
Private sErrorSource        As String
Private sErrorDescription   As String
Private sErrorPath          As String

' Test button, displayed with Conditional Compile Argument Test = 1
Public Property Get ExitAndContinue() As String:    ExitAndContinue = "Exit procedure" & vbLf & "and continue" & vbLf & "with next":    End Property

' Debugging button, displayed with Conditional Compile Argument Debugging = 1
Public Property Get ResumeError() As String:        ResumeError = "Resume" & vbLf & "error code line":                                  End Property

' Test button, displayed with Conditional Compile Argument Test = 1
Public Property Get ResumeNext() As String:         ResumeNext = "Continue with code line" & vbLf & "following the error line":         End Property

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
' ----------------------------------
' Trace and stack Begin of Procedure
' ----------------------------------
    
#If ExecTrace Then
    TrcBegin s   ' start of the procedure's execution trace
#End If
    StackPush s
'    Debug.Print "BoP: '" & s & "', Stack Top: '" & StackTop & "', Stack Bottom: '" & StackBottom & "'"
    
End Sub

Public Sub BoT(ByVal s As String)
' -------------------------------
' Begin of Trace for code lines.
' -------------------------------
#If ExecTrace Then
    CallStack.TraceBegin s
#End If
End Sub

Public Sub EoP(ByVal s As String)
' --------------------------------
' Trace and stack End of Procedure
' --------------------------------
    
#If ExecTrace Then
    TrcEnd s
#End If
    StackPop s
    
'    Debug.Print "EoP: '" & s & "', Stack Top: '" & StackTop & "', Stack Bottom: '" & StackBottom & "'"
    If StackIsEmpty Or s = sErrHndlrEntryProc Then
        TrcDsply
    End If
    
End Sub

Public Sub EoT(ByVal s As String)
' -------------------------------
' End of Trace for code lines.
' -------------------------------
    TrcEnd s
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
#Const AlternateMsgBox = 1  ' 1 = Error displayed by means of the Alternative MsgBox fMsg
                            ' 0 = Error displayed by means of the VBA MsgBox
    
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

'    If CallStack Is Nothing Then Set CallStack = New clsCallStack
    If cllErrPath Is Nothing Then Set cllErrPath = New Collection
    If errline <> 0 Then sLine = errline Else sLine = "0"
    
    If ErrPathIsEmpty Then
        '~~ When there's yet no error path collected this indicates that the
        '~~ error handler is executed the first time This is the error raising procedure. Backtracking to the entry procedure is due
        ErrPathAdd errsource & " (" & ErrorDetails(errnumber, errline) & ")"
        TrcError errsource & " !!! " & ErrorDetails(errnumber, errline) & " """ & errdscrptn & """"
        lSourceErrorNo = errnumber
        sErrorSource = errsource
        sErrorDescription = errdscrptn
    ElseIf lErrorNumber <> errnumber Then
        '~~ The error number had changed during the process of passing it on to the entry procedure
'            ErrHndlrErrPathAdd errsource & " (" & ErrorDetails(errnumber, errline) & ")"
        TrcError errsource & ": " & ErrorDetails(errnumber, errline) & " """ & errdscrptn & """"
        lErrorNumber = errnumber
    Else
        '~~ This is the error handling called during the "backtracing" process,
        '~~ i.e. the process when the error is passed on up to the entry procedure
'            ErrHndlrErrPathAdd errsource
    End If
    
    '~~ When the user has no choice to press any button but the only one displayed button
    '~~ and the Entry Procedure is known but yet not reached the path back up to the Entry Procedure
    '~~ is maintained and the error is passed on to the caller
    If ErrorButtons(buttons) = 1 _
    And sErrHndlrEntryProc <> vbNullString _
    And StackEntryProc <> errsource Then
        ErrPathAdd errsource
        StackPop errsource
        Err.Raise errnumber, errsource, errdscrptn
    End If
    
    '~~ When more than one button is displayed for the user to choose one
    '~~ or the Entry Procedure is unknown or has been reached
    '~~ the error is displayed
    If ErrorButtons(buttons) > 1 _
    Or StackEntryProc = errsource _
    Or StackEntryProc = vbNullString Then
        ErrPathAdd errsource
        StackPop errsource
        ErrHndlr = ErrMsg(errnumber:=lSourceErrorNo, errsource:=sErrorSource, errdscrptn:=sErrorDescription, errline:=errline, errpath:=ErrPathErrMsg, buttons:=buttons)
        Select Case ErrHndlr
            Case ResumeError, ResumeNext, ExitAndContinue
            Case Else: ErrPathErase
        End Select
    End If
    
    '~~ Each time a known Entry Procedure is reached the execution trace
    '~~ maintained by the BoP and EoP and the BoT and EoT statements is displayed
    If StackEntryProc = errsource _
    Or StackEntryProc = vbNullString Then
#If ExecTrace Then
        TrcDsply
#End If
        Select Case ErrHndlr
            Case ResumeError, ResumeNext, ExitAndContinue
            Case vbOK
            Case Else: StackErase
        End Select
    End If
            
End Function

Private Sub ErrHndlrAddButtons( _
            ByRef buttons As Variant, _
            ByVal s As String)
    buttons = buttons & "," & s
End Sub

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
    
    Dim sErrMsg, sErrPath As String
    
#If AlternateMsgBox Then
    '~~ Display the error message by means of the Common UserForm fMsg
    With fMsg
        .MsgTitle = ErrorType(errnumber, errsource) & " in " & errsource
        .MsgLabel(1) = "Error Message/Description:":    .MsgText(1) = ErrorDescription(errdscrptn)
        .MsgLabel(2) = "Error path (call stack):":      .MsgText(2) = ErrPathErrMsg:   .MsgMonoSpaced(2) = True
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
#Else
    '~~ Display the error message by means of the VBA MsgBox
    sErrMsg = "Description: " & vbLf & ErrorDescription(errdscrptn) & vbLf & vbLf & _
              "Source:" & vbLf & errsource & ErrorLine(errline)
    sErrPath = ErrMsgErrPath
    If sErrPath <> vbNullString _
    Then sErrMsg = sErrMsg & vbLf & vbLf & _
                   "Path:" & vbLf & errpath
    If ErrorInfo(errdscrptn) <> vbNullString _
    Then sErrMsg = sErrMsg & vbLf & vbLf & _
                   "Info:" & vbLf & ErrorInfo(errdscrptn)
    ErrMsg = MsgBox(Prompt:=sErrMsg, buttons:=buttons, Title:=ErrMsgErrType(errnumber, errsource) & " in " & errsource & ErrorLine(errline))
#End If
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
    If InStr(s, DCONCAT) <> 0 _
    Then ErrorDescription = Split(s, DCONCAT)(0) _
    Else ErrorDescription = s
End Function

Private Function ErrorDetails( _
                 ByVal errnumber As Long, _
                 ByVal sErrLine As String) As String
' -----------------------------------------------------------------
' Returns kind of error, error number, and error line if available.
' -----------------------------------------------------------------
    
    If errnumber < 0 _
    Then ErrorDetails = EXEC_TRACE_APP_ERR & AppErr(errnumber) _
    Else ErrorDetails = EXEC_TRACE_VB_ERR & errnumber
    If sErrLine <> 0 Then ErrorDetails = ErrorDetails & " at line " & sErrLine

End Function

Private Function ErrorInfo(ByVal s As String) As String
' -----------------------------------------------------
' Returns the string which follows a "||" in the error
' description which indicates an additional information
' regarding the error.
' -----------------------------------------------------
    If InStr(s, DCONCAT) <> 0 _
    Then ErrorInfo = Split(s, DCONCAT)(1) _
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

Private Function ErrorType( _
                 ByVal errnumber As Long, _
                 ByVal errsource As String) As String
' ---------------------------------------------------
' Return the kind of error considering the error
' source (errsource) and the error number (errnumber).
' ----------------------------------------------------

   If InStr(1, errsource, "DAO") <> 0 _
   Or InStr(1, errsource, "ODBC Teradata Driver") <> 0 _
   Or InStr(1, errsource, "ODBC") <> 0 _
   Or InStr(1, errsource, "Oracle") <> 0 Then
      ErrorType = "Database Error " & errnumber
   Else
      If errnumber > 0 _
      Then ErrorType = "VB Runtime Error " & errnumber _
      Else ErrorType = "Application Error " & AppErr(errnumber)
   End If
   
End Function

Private Sub ErrPathAdd(ByVal s As String)
    
    If ErrPathIsEmpty Then
        sErrorPath = s
    Else
        If InStr(sErrorPath, s & " ") = 0 Then
            sErrorPath = s & vbLf & sErrorPath
        End If
    End If
End Sub

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

Private Function ErrPathIsEmpty() As Boolean
   ErrPathIsEmpty = sErrorPath = vbNullString
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

Private Sub StackInit()
    If dctStack Is Nothing Then Set dctStack = New Dictionary Else dctStack.RemoveAll
End Sub

Private Function StackIsEmpty() As Boolean
    StackIsEmpty = dctStack Is Nothing
    If Not StackIsEmpty Then StackIsEmpty = dctStack.Count = 0
End Function

Private Function StackPop( _
        Optional ByVal s As String = vbNullString) As String
' ----------------------------------------------------------
' Returns the item removed from the top of the stack.
' When s is provided and is not on the top of the stack
' an error is raised.
' ----------------------------------------------------------
    
    On Error GoTo on_error
    Const PROC = "StackPop"

    If Not StackIsEmpty Then
        If s <> vbNullString And s = dctStack.Items()(dctStack.Count - 1) Then
            StackPop = dctStack.Items()(dctStack.Count - 1) ' Return the poped item
            dctStack.Remove dctStack.Count                  ' Remove item s from stack
        ElseIf s = vbNullString Then
            dctStack.Remove dctStack.Count                  ' Unwind! Remove item s from stack
        End If
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

End Sub

Private Function StackTop() As String
    If Not StackIsEmpty Then StackTop = dctStack.Items()(dctStack.Count - 1)
End Function

Private Function StackUnwind() As String
' --------------------------------------
' Completes the execution trace for
' items still on the stack.
' --------------------------------------

    Do Until StackIsEmpty
        TrcEnd StackPop
    Loop
     
End Function

Public Sub TrcBegin(ByVal s As String)
' ------------------------------------------
' Keep a record (tick count) of the begin of
' the execution of any group of vba code
' lines described/named (s).
' ------------------------------------------
    Dim cy      As Currency
    
    getTickCount cy
    TrcInit
    
    getTickCount cyTicks
    TrcAdd EXEC_TRACE_BEGIN_ID & s, cyTicks
    getTickCount cyTicks
    cyOverhead = cyOverhead + (cyTicks - cy)
    
    '~~ Reset a possibly error raised procedure
    sErrorSource = vbNullString

End Sub

Private Function TrcBeginLine(ByVal cyInitial As Currency, _
                                ByVal iTt As Long, _
                                ByVal sIndent As String, _
                                ByVal iIndent As Long, _
                                ByVal sProcName As String, _
                                ByVal sMsg As String) As String
' -------------------------------------------------------------
'
' -------------------------------------------------------------
    TrcBeginLine = TrcSecs(cyInitial, dicTrace.Items(iTt)) & _
                     "    " & _
                     sIndent & _
                     " " & _
                     Replicate("|  ", iIndent) & _
                     sProcName & _
                     sMsg
End Function

Private Function TrcBeginTicks(ByVal s As String, _
                          ByVal i As Single) As Currency
' -------------------------------------------------------
' Returns the number of ticks recorded with the begin
' item corresponding with the end item (s) by searching
' the trae Dictionary back up starting with the index (i)
' -1 (= index of the end time (s)).
' Returns 0 when no start item coud be found.
' To avoid multiple identifications of the begin item
' it is set to vbNullString with the return of the number
' of begin ticks.
' -------------------------------------------------------
    
    Dim j           As Single
    Dim sItem, sKey As String

    TrcBeginTicks = 0
    s = Replace(s, EXEC_TRACE_END_ID, EXEC_TRACE_BEGIN_ID)  ' turn the end item into a begin item string
    For j = i - 1 To 0 Step -1
        sKey = dicTrace.Keys(j)
        sItem = Right(sKey, Len(sKey) - InStr(1, sKey, EXEC_TRACE_BEGIN_ID) + 1)
        If sItem = s Then
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

Public Function TrcDsply(Optional ByVal bDebugPrint As Boolean = True) As String
' --------------------------------------------------------------------------------
' Displays the precision time tracking result with the execution
' time in seconds with each vba code lines end tracking line.
' --------------------------------------------------------------------------------
#If ExecTrace Then
    
    On Error GoTo on_error
    Const PROC = "TrcDsply"        ' This procedure's name for the error handling and execution tracking
    Const ELAPSED = "Elapsed"
    Const EXEC_SECS = "Exec secs"
    
    Dim cyStrt      As Currency ' ticks count at start
    Dim cyEnd       As Currency ' ticks count at end
    Dim cyElapsed   As Currency ' elapsed ticks since start
    Dim cyInitial   As Currency ' initial ticks count (at first traced proc)
    Dim iTt         As Single   ' index for dictionary dicTrace
    Dim sProcName   As String   ' tracked procedure/vba code
    Dim iIndent     As Single   ' indentation nesting level
    Dim sIndent     As String   ' Indentation string defined by the precision
    Dim cll         As Collection
    Dim sMsg        As String
    Dim dbl         As Double
    Dim i           As Long
    Dim sTrace      As String
    Dim sTraceLine  As String
    Dim sInfo       As String
       
    StackUnwind ' end procedures still on the stack
    
    If dicTrace Is Nothing Then Exit Function   ' When the contional compile argument where not
    If dicTrace.Count = 0 Then Exit Function    ' ExecTrace = 1 there will be no execution trace result

    Set cll = TrcItemsUnPaired ' collect all unpaired trace entries
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
    If dbl >= 1000 Then
        iSec = 4
    ElseIf dbl >= 100 Then
        iSec = 3
    ElseIf dbl >= 10 Then
        iSec = 2
    Else
        iSec = 1
    End If
    
    sFormat = String$(iSec - 1, "0") & "0." & String$(iDec, "0") & " "
    sIndent = Space$(Len(sFormat))
    iIndent = -1
    '~~ Header
    
    sTraceLine = ELAPSED & VBA.Space$(Len(sIndent) - Len(ELAPSED) + 1) & EXEC_SECS & " >> Begin execution trace " & Format(dtTraceBeginTime, "hh:mm:ss") & " (exec time in seconds)"
    If bDebugPrint Then Debug.Print sTraceLine Else sTrace = sTrace & sTraceLine
    
    For iTt = 0 To dicTrace.Count - 1
        sProcName = dicTrace.Keys(iTt)
        
        If TrcIsEndItem(sProcName, sInfo) Then
            '~~ Trace End Line
            cyEnd = dicTrace.Items(iTt)
            cyStrt = TrcBeginTicks(sProcName, iTt)   ' item is set to vbNullString to avoid multiple recognition
            If cyStrt = 0 Then
                '~~ BoP/BoT code line missing
                iIndent = iIndent + 1
                
                sTraceLine = Space$((Len(sFormat) * 2) + 1) & "    " & Replicate("|  ", iIndent) & sProcName & sInfo & " !!! the corresponding BoP/BoT code line missing !!! (trace result incomplete)"
                If bDebugPrint Then Debug.Print sTraceLine Else sTrace = sTrace & sTraceLine
                
                iIndent = iIndent - 1
            Else
                '~~ End line
                sTraceLine = TrcEndLine(cyInitial, cyEnd, cyStrt, iIndent, sProcName & sInfo)
                If bDebugPrint Then Debug.Print sTraceLine Else sTrace = sTrace & sTraceLine
                iIndent = iIndent - 1
            End If
        ElseIf TrcIsBegItem(sProcName) Then
            '~~ Begin Trace Line
            iIndent = iIndent + 1
            If iTt = 0 Then cyInitial = dicTrace.Items(iTt)
            sMsg = TrcEndItemMissing(sProcName)
            
            sTraceLine = TrcBeginLine(cyInitial, iTt, sIndent, iIndent, sProcName, sMsg)
            If bDebugPrint Then Debug.Print sTraceLine Else sTrace = sTrace & sTraceLine
            
            If sMsg <> vbNullString Then iIndent = iIndent - 1
        
        End If
    Next iTt
    
    dicTrace.RemoveAll
    sTraceLine = Space$((Len(sFormat) * 2) + 2) & "<< End execution trace " & Format(Now(), "hh:mm:ss") & " (only " & Format(TrcSecs(0, cyOverhead), "0.000000") & " seconds exec time were caused by the executuion trace itself)"
    If bDebugPrint Then Debug.Print sTraceLine Else sTrace = sTrace & sTraceLine

    sTraceLine = Space$((Len(sFormat) * 2) + 2) & "The Conditional Compile Argument 'ExecTrace = 0' will turn off the trace and its display." & vbLf
    If bDebugPrint Then Debug.Print sTraceLine Else sTrace = sTrace & sTraceLine
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

Private Sub TrcEnd(ByVal s As String)
' -----------------------------------------
' End of Trace. Keeps a record of the ticks
' count for the execution trace of the
' group of code lines named (s).
' -----------------------------------------
#If ExecTrace Then
    Const PROC = "TrcEnd"
    Dim cy      As Currency

    On Error GoTo on_error
    
    getTickCount cyTicks
    cy = cyTicks
    TrcAdd EXEC_TRACE_END_ID & s, cyTicks
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
    Dim i, j    As Long
    Dim sKey    As String
    Dim sItem   As String
    Dim sInfo   As String

    TrcEndItemMissing = " !!! the corresponding EoP/EoT code line missing !!! (trace result incomplete)"
    s = Replace(s, EXEC_TRACE_BEGIN_ID, EXEC_TRACE_END_ID)  ' turn the end item into a begin item string
    For i = 0 To dicTrace.Count - 1
        sKey = dicTrace.Keys(i)
        If TrcIsEndItem(sKey, sInfo) Then
            If sKey = s Then
                TrcEndItemMissing = vbNullString
                GoTo exit_proc
            End If
        End If
    Next i
    
exit_proc:
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
    TrcAdd Split(EXEC_TRACE_END_ID & s, DCONCAT)(0), cyTicks
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
    i = InStr(1, s, EXEC_TRACE_BEGIN_ID)
    If i <> 0 Then
        TrcIsBegItem = True
        s = Right(s, Len(s) - i + 1)
    End If
End Function

Private Function TrcIsEndItem( _
                 ByRef s As String, _
                 ByRef sRight As String) As Boolean
' -------------------------------------------------
' Returns TRUE if s is an execution trace end item.
' Returns s with the item counter unstripped. Any
' additional info is returne in sRight.
' -------------------------------------------------
    Dim i As Single

    TrcIsEndItem = False
    i = InStr(1, s, EXEC_TRACE_END_ID)
    If i <> 0 Then
        TrcIsEndItem = True
        
        If InStr(s, " !!!") <> 0 _
        Then sRight = " !!!" & Split(s, " !!!")(1) _
        Else sRight = vbNullString
        
        s = Split(s, " !!!")(0)
        s = Right(s, Len(s) - i + 1)
    End If
    
End Function

Public Function TrcIsErrorItem(ByRef s As String) As Boolean
' ----------------------------------------------------------
' Returns TRUE if s is an execution trace begin item.
' Returns s with the call counter unstripped.
' ----------------------------------------------------------
    
    TrcIsErrorItem = False
    If InStr(1, s, EXEC_TRACE_APP_ERR) <> 0 _
    Or InStr(s, EXEC_TRACE_VB_ERR) <> 0 Then
        TrcIsErrorItem = True
        s = UnstripNumber(s) ' Unstrip item number
    End If
    
End Function

Private Function TrcItemsUnPaired() As Collection
' -----------------------------------------------
'
' -----------------------------------------------
    Dim cll     As Collection
    Dim i, j, k As Long
    Dim s1, s2  As String
    Dim b       As Boolean

    Set cll = New Collection
    
    '~~ 1. paired end items
    For i = 0 To dicTrace.Count - 1
        b = False
        s1 = dicTrace.Keys(i)
        k = InStr(1, s1, EXEC_TRACE_BEGIN_ID)
        If k <> 0 Then
            s1 = Right(s1, Len(s1) - k + 1) ' unstrip leading digits
            '~~ Search for associated end item
            For j = 0 To dicTrace.Count - 1
                s2 = dicTrace.Keys(j)
                k = InStr(1, s2, EXEC_TRACE_END_ID)
                If k <> 0 Then
                    s2 = Right(s2, Len(s2) - k + 1) ' unstrip leading digits
                    s2 = Replace(s2, EXEC_TRACE_END_ID, EXEC_TRACE_BEGIN_ID)
                    If s2 = s1 Then
                        b = True
                        Exit For
                    End If
                End If
            Next j
            If Not b Then
                cll.Add s1
            End If
        End If
    Next i
    Set TrcItemsUnPaired = cll

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

Private Function UnstripNumber( _
                 ByVal s As String) As String
    Dim i As Long

    i = 1
    While IsNumeric(Mid(s, i, 1))
        i = i + 1
    Wend
    s = Right(s, Len(s) - (i - 1))
    UnstripNumber = s
    
End Function

