VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "clsCallStack"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit
' -----------------------------------------------------------------------------
' Class Module clsCallStack
'
' Methods:  - StackPush Obligatory for each procedure with Error handling at
'                       procedure start. Traces the procedure's execution start
'                       when theconditional compile argument Trace = 1.
'           - StackPop  Obligatory for each procedure with Error handling at
'                       procedure end. Traces the procedure's execution end
'                       when the conditional compile argument Trace = 1
'           - DsplayTt  Displays the procedure execution tracing and precision
'                       time tracking result.
'
' Uses:     - Class Module clsCallStackItem
' Requires: - Reference to "Microsoft Scripting Runtime"
' Note: In order not to interfere with the common error handling this module
'       uses a local ErrHndlr
'
' W. Rauschenberger, Berlin Oct 2019
' -----------------------------------------------------------------------------
Const MODNAME       As String = "clsCallStack"
Private TopItem     As clsCallStackItem
' --- Begin of declaration for the Procedure Execution Tracing and PrecisionDecimals Time Tracking
Private Declare PtrSafe Function getFrequency Lib "kernel32" _
Alias "QueryPerformanceFrequency" (cyFrequency As Currency) As Long

Private Declare PtrSafe Function getTickCount Lib "kernel32" _
Alias "QueryPerformanceCounter" (cyTickCount As Currency) As Long

Const EXEC_TRACE_BEGIN_ID   As String = ">> "       ' Begin item indicator
Const EXEC_TRACE_END_ID     As String = "<< "       ' End item indicator
Const EXEC_TRACE_APP_ERR    As String = "App err"
Const EXEC_TRACE_VB_ERR     As String = "VB err"

Private cyFrequency         As Currency     ' Execution Trace Frequency (initialized with init)
Private cyTicks             As Currency     ' Execution Trace Ticks counter
Private iCall               As Long         ' Execution Trace Call counter to unify key
Private lPrecisionDecimals  As Long         ' Execution Trace Default Precision (6=0,000000)
Private iSec                As Integer      ' Execution Trace digits left from decimal point
Private iDec                As Integer      ' Execution Trace decimal digits right from decimal point
Private sFormat             As String       ' Execution Trace tracking time presentation format
Private cyOverhead          As Currency     ' Execution Trace time accumulated by caused by the time tracking itself
Private sErrorSource        As String       ' Execution Trace procedure which initially raised an error
Private lSourceErrorNo      As Long         ' Execution Trace Error Number raised by the ErrorSource
Private lErrorNumber        As Long         ' Execution Trace Error Number all the way up to the entry procedure
Private dtTraceBeginTime    As Date         ' Execution Trace start time
Private dtTraceEndTime      As Date         ' Execution Trace end time
Private sErrorDescription   As String       ' The error description of the initial error raising procedure
' --- End of declaration for the Procedure Execution Tracing and PrecisionDecimals Time Tracking ------
Private sEntryProc          As String       ' Entry Procedure (the first with a call stack maintenance statement)
Private sErrorPath          As String       ' Path from ther error raising procedure up to the entry procedure

Public Property Let PrecisionDecimals(ByVal l As Long):     lPrecisionDecimals = l:                 End Property
Public Property Get EntryProc() As String:                  EntryProc = sEntryProc:                 End Property
Public Property Let ErrorSource(ByVal s As String):         sErrorSource = s:                       End Property
Public Property Get ErrorSource() As String:                ErrorSource = sErrorSource:             End Property
Public Property Let ErrorNumber(ByVal l As Long):           lErrorNumber = l:                       End Property
Public Property Get ErrorNumber() As Long:                  ErrorNumber = lErrorNumber:             End Property
Public Property Let SourceErrorNo(ByVal l As Long):         lSourceErrorNo = l:                     End Property
Public Property Get SourceErrorNo() As Long:                SourceErrorNo = lSourceErrorNo:         End Property
Public Property Let TraceBeginTime(ByVal dt As Date):       dtTraceBeginTime = dt:                  End Property
Public Property Let TraceEndTime(ByVal dt As Date):         dtTraceEndTime = dt:                    End Property
Public Property Let ErrorPath(ByVal s As String):           sErrorPath = s:                         End Property
Public Property Get ErrorPath() As String:                  ErrorPath = sErrorPath:                 End Property
Public Property Let ErrorDescription(ByVal s As String):    sErrorDescription = s:                  End Property
Public Property Get ErrorDescription() As String:           ErrorDescription = sErrorDescription:   End Property

Private Sub TraceInit()
    dtTraceBeginTime = Now()
    iCall = 0
    cyOverhead = 0
End Sub

Public Property Get StackIsEmpty() As Boolean
     StackIsEmpty = (TopItem Is Nothing)
End Property

Public Function StackPop(Optional ByVal sErrSource As String = vbNullString) As String
' ------------------------------------------------------------------------------------
' Takes off the last pushed procedure from the stack. When no sErrSource is provided
' this is an "unwind pop".
' ------------------------------------------------------------------------------------
Dim sTopValue   As String
    
    '~~ Take off last pushed procedure
    If Not StackIsEmpty Then
        sTopValue = TopItem.Value
        '~~ When the to be poped procedure is not the last pushed one
        '~~ an error must have happened. The pop is thus postponed to the Unwind
        If (sErrSource <> vbNullString And sTopValue = sErrSource) _
        Or sErrSource = vbNullString Then
            sTopValue = TopItem.Value
            StackPop = sTopValue
            Set TopItem = TopItem.NextItem
        End If
     
        If Not sErrSource = vbNullString Then
'            TraceEnd sErrSource '~~ Implicite end the exec trace
            If sErrSource = EntryProc Then
                ' Whenever the entry procedure's execution ends the exec trace is displayed
#If ExecTrace Then
                TraceDsply
#Else
            Set dicTrace = Nothing
#End If
            End If
        End If
    End If
          
End Function

Public Sub StackPush(ByVal sErrSource As String)
' ----------------------------------------------
' StackPush/put the procedure (sErrSource) on
' top of the stack and begin exec trace when
' the conditional compile argument Trace=1.
' ---------------------------------------------
Dim NewTop As clsCallStackItem

    If StackIsEmpty Then
        sEntryProc = sErrSource ' First pushed = bottom item = entry procedure
    End If
    '~~ StackPush sErrSource on top of the CallStack
    Set NewTop = New clsCallStackItem
    With NewTop
        .Value = sErrSource
        Set .NextItem = TopItem
    End With
    Set TopItem = NewTop
        
#If ExecTrace Then
'    TraceBegin sErrSource   ' implicite start of the procedure's exec trace
#End If

End Sub

Public Function StackUnwind() As String
Dim sProc As String

     Do Until StackIsEmpty
        sProc = StackPop
        Me.TraceEnd sProc
     Loop
     
End Function

Public Function StackTop() As String
    StackTop = TopItem.Value
End Function

Public Sub TraceBegin(ByVal s As String)
' ------------------------------------------
' Keep a record (tick count) of the begin of
' the execution of any group of vba code
' lines described/named (s).
' ------------------------------------------
Dim cy As Currency

    getTickCount cy
    If Not dicTrace Is Nothing Then
        If dicTrace.Count = 0 Then TraceInit
    Else
        Set dicTrace = New Dictionary
        TraceInit
    End If
    
    getTickCount cyTicks
    iCall = iCall + 1
    dicTrace.Add iCall & EXEC_TRACE_BEGIN_ID & s, cyTicks
    getTickCount cyTicks
    cyOverhead = cyOverhead + (cyTicks - cy)
    
    '~~ Reset a possibly error raised procedure
    sErrorSource = vbNullString

End Sub

Public Sub TraceError(ByVal s As String)
' --------------------------------------
' Keep record of the error (s) raised
' during the execution of any procedure.
' --------------------------------------
#If ExecTrace Then
Dim cy      As Currency
Dim sErr    As String

    getTickCount cy
    If Not dicTrace Is Nothing Then
        If dicTrace.Count = 0 Then TraceInit
    Else
        Set dicTrace = New Dictionary
        TraceInit
    End If
    
    getTickCount cyTicks
    iCall = iCall + 1
    '~~ Add the error indication line to the trace by ignoring any additional error information
    '~~ optionally attached by two vertical bars
    dicTrace.Add iCall & Split(s, DCONCAT)(0), cyTicks
    getTickCount cyTicks
    cyOverhead = cyOverhead + (cyTicks - cy)
#End If
End Sub

Public Sub TraceEnd(ByVal s As String)
' -----------------------------------------
' End of Trace. Keeps a record of the ticks
' count for the execution trace of the
' group of code lines named (s).
' -----------------------------------------
#If ExecTrace Then
Const PROC  As String = "TraceEnd"
Dim cy      As Currency

    On Error GoTo on_error
    
    getTickCount cyTicks
    cy = cyTicks
    iCall = iCall + 1
    dicTrace.Add iCall & EXEC_TRACE_END_ID & s, cyTicks
    getTickCount cyTicks
    cyOverhead = cyOverhead + (cyTicks - cy)

exit_proc:
    Exit Sub
    
on_error:
    ErrHndlr Err.Number, ErrSrc(PROC), Err.Description, Erl
#End If
End Sub

Private Function TraceBeginLine(ByVal cyInitial As Currency, _
                                ByVal iTt As Long, _
                                ByVal sIndent As String, _
                                ByVal iIndent As Long, _
                                ByVal sProcName As String, _
                                ByVal sMsg As String) As String
' -------------------------------------------------------------
'
' -------------------------------------------------------------
    TraceBeginLine = TraceSecs(cyInitial, dicTrace.Items(iTt)) & _
                     "    " & _
                     sIndent & _
                     " " & _
                     Replicate("|  ", iIndent) & _
                     sProcName & _
                     sMsg
End Function

Private Function TraceEndLine(ByVal cyInitial As Currency, _
                              ByVal cyEnd As Currency, _
                              ByVal cyStrt As Currency, _
                              ByVal iIndent As Long, _
                              ByVal sProcName As String) As String
' ----------------------------------------------------------------
' Assemble a Trace End Line
' ----------------------------------------------------------------
    TraceEndLine = TraceSecs(cyInitial, cyEnd) & _
                   " " & _
                   TraceSecs(cyStrt, cyEnd) & _
                   "    " & _
                   Replicate("|  ", iIndent) & _
                   sProcName
End Function

Public Function TraceDsply(Optional ByVal bDebugPrint As Boolean = True) As String
' --------------------------------------------------------------------------------
' Displays the precision time tracking result with the execution
' time in seconds with each vba code lines end tracking line.
' --------------------------------------------------------------------------------
#If ExecTrace Then
Const PROC      As String = "TraceDsply"        ' This procedure's name for the error handling and execution tracking
Const sElapsed  As String = "Elapsed"
Const sExecSec  As String = "Exec secs"
Dim cyStrt      As Currency ' ticks count at start
Dim cyEnd       As Currency ' ticks count at end
Dim cyElapsed   As Currency ' elapsed ticks since start
Dim cyInitial   As Currency ' initial ticks count (at first traced proc)
Dim iTt         As Single   ' index for dictionary dicTrace
Dim sProcName   As String   ' tracked procedure/vba code
Dim iIndent     As Single   ' indentation nesting level
Dim sIndent     As String   ' Indentation string defined by the precision
Dim cll         As Collection
Dim v           As Variant
Dim sMsg        As String
Dim lSec        As Long
Dim dbl         As Double
Dim i           As Long
Dim sTrace      As String
Dim sTraceLine  As String

    On Error GoTo on_error
    
    '~~ When the trace result is to be displayed there may bee still procedures on the call stack not
    '~~ poped due to an error. All these remaing procedures are "unwind" in order to close their tracing.
    CallStack.StackUnwind   ' Any remaining items
    
    '~~ When the contional compile argument where not Trace=1 there will be no execution trace result
    If dicTrace Is Nothing Then Exit Function
    If dicTrace.Count = 0 Then Exit Function

    Set cll = TraceItemsUnPaired ' collect all unpaired trace entries
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
    
    sTraceLine = sElapsed & VBA.Space$(Len(sIndent) - Len(sElapsed) + 1) & sExecSec & " >> Begin execution trace " & Format(dtTraceBeginTime, "hh:mm:ss") & " (exec time in seconds)"
    If bDebugPrint Then Debug.Print sTraceLine Else sTrace = sTrace & sTraceLine
    
    For iTt = 0 To dicTrace.Count - 1
        sProcName = dicTrace.Keys(iTt)
        If TraceIsBegItem(sProcName) Then
            '~~ Begin Trace Line
            iIndent = iIndent + 1
            If iTt = 0 Then cyInitial = dicTrace.Items(iTt)
            sMsg = TraceEndItemMissing(sProcName)
            
            sTraceLine = TraceBeginLine(cyInitial, iTt, sIndent, iIndent, sProcName, sMsg)
            If bDebugPrint Then Debug.Print sTraceLine Else sTrace = sTrace & sTraceLine
            
            If sMsg <> vbNullString Then iIndent = iIndent - 1
        ElseIf TraceIsErrorItem(sProcName) Then
            '~~ Error Line
            sTraceLine = Space$((Len(sFormat) * 2) + 1) & "    " & Replicate("|  ", iIndent) & VBA.Space$(Len(EXEC_TRACE_END_ID)) & sProcName
            If bDebugPrint Then Debug.Print sTraceLine Else sTrace = sTrace & sTraceLine
        
        ElseIf TraceIsEndItem(sProcName) Then
            '~~ Trace End Line
            cyEnd = dicTrace.Items(iTt)
            cyStrt = TraceBeginTicks(sProcName, iTt)   ' item is set to vbNullString to avoid multiple recognition
            If cyStrt = 0 Then
                '~~ BoP/BoT code line missing
                iIndent = iIndent + 1
                
                sTraceLine = Space$((Len(sFormat) * 2) + 1) & "    " & Replicate("|  ", iIndent) & sProcName & " !!! the corresponding BoP/BoT code line missing !!! (trace result incomplete)"
                If bDebugPrint Then Debug.Print sTraceLine Else sTrace = sTrace & sTraceLine
                
                iIndent = iIndent - 1
            Else
                '~~ End line
                sTraceLine = TraceEndLine(cyInitial, cyEnd, cyStrt, iIndent, sProcName)
                If bDebugPrint Then Debug.Print sTraceLine Else sTrace = sTrace & sTraceLine
                iIndent = iIndent - 1
            End If
        End If
    Next iTt
    
    dicTrace.RemoveAll
    sTraceLine = Space$((Len(sFormat) * 2) + 2) & "<< End execution trace " & Format(Now(), "hh:mm:ss") & " (only " & Format(TraceSecs(0, cyOverhead), "0.000000") & " seconds exec time were caused by the executuion trace itself)"
    If bDebugPrint Then Debug.Print sTraceLine Else sTrace = sTrace & sTraceLine

    sTraceLine = Space$((Len(sFormat) * 2) + 2) & "The Conditional Compile Argument 'ExecTrace = 0' will turn off the trace and its display." & vbLf
    If bDebugPrint Then Debug.Print sTraceLine Else sTrace = sTrace & sTraceLine
    TraceDsply = sTrace
    Exit Function
    
on_error:
    ErrHndlr Err.Number, ErrSrc(PROC), Err.Description, Erl
#End If
End Function

Private Function TraceSecs(ByVal cyStrt As Currency, _
                     ByVal cyEnd As Currency) As String
' --------------------------------------------------------
' Returns the difference between cyStrt and cyEnd as
' formatted seconds string (decimal = nanoseconds).
' --------------------------------------------------------
Dim dbl     As Double

    dbl = (cyEnd - cyStrt) / cyFrequency
    TraceSecs = Format(dbl, sFormat)

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

Public Function TraceIsErrorItem(ByRef s As String) As Boolean
' ---------------------------------------------------------------
' Returns TRUE if s is an execution trace begin item.
' Returns s with the call counter unstripped.
' ---------------------------------------------------------------
Dim i As Single
    TraceIsErrorItem = False
    i = InStr(1, s, EXEC_TRACE_APP_ERR)
    If i = 0 Then
        i = InStr(s, EXEC_TRACE_VB_ERR)
    Else
        TraceIsErrorItem = True
        '~~ Unstrip number
        s = UnstripNumber(s)
    End If
End Function

Private Function UnstripNumber(ByVal s As String) As String
Dim i As Long

    i = 1
    While IsNumeric(Mid(s, i, 1))
        i = i + 1
    Wend
    s = Right(s, Len(s) - (i - 1))
    UnstripNumber = s
    
End Function
Private Function TraceIsBegItem(ByRef s As String) As Boolean
' ---------------------------------------------------------------
' Returns TRUE if s is an execution trace begin item.
' Returns s with the call counter unstripped.
' ---------------------------------------------------------------
Dim i As Single
    TraceIsBegItem = False
    i = InStr(1, s, EXEC_TRACE_BEGIN_ID)
    If i <> 0 Then
        TraceIsBegItem = True
        s = Right(s, Len(s) - i + 1)
    End If
End Function

Private Function TraceEndItemMissing(ByVal s As String) As String
' -------------------------------------------------------------------
' Returns a message string when a corresponding end item is missing.
' -------------------------------------------------------------------
Dim i       As Long
Dim sKey    As String
Dim sItem   As String
Dim j       As Long

    TraceEndItemMissing = " !!! the corresponding EoP/EoT code line missing !!! (trace result incomplete)"
    s = Replace(s, EXEC_TRACE_BEGIN_ID, EXEC_TRACE_END_ID)  ' turn the end item into a begin item string
    For i = 0 To dicTrace.Count - 1
        sKey = dicTrace.Keys(i)
        j = InStr(sKey, EXEC_TRACE_END_ID)
        If j <> 0 Then ' end item
            sItem = Right(sKey, Len(sKey) - j + 1)
            If sItem = s Then
                TraceEndItemMissing = vbNullString
                GoTo exit_proc
            End If
        End If
    Next i
    
exit_proc:
End Function

Private Function TraceIsEndItem(ByRef s As String) As Boolean
' ---------------------------------------------------------------
' Returns TRUE if s is an execution trace end item.
' Returns s with the call counter unstripped.
' ---------------------------------------------------------------
Dim i As Single

    TraceIsEndItem = False
    i = InStr(1, s, EXEC_TRACE_END_ID)
    If i <> 0 Then
        TraceIsEndItem = True
        s = Right(s, Len(s) - i + 1)
    End If
    
End Function

Private Function TraceBeginTicks(ByVal s As String, _
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
Dim j       As Single
Dim sItem   As String
Dim sKey    As String

    TraceBeginTicks = 0
    s = Replace(s, EXEC_TRACE_END_ID, EXEC_TRACE_BEGIN_ID)  ' turn the end item into a begin item string
    For j = i - 1 To 0 Step -1
        sKey = dicTrace.Keys(j)
        sItem = Right(sKey, Len(sKey) - InStr(1, sKey, EXEC_TRACE_BEGIN_ID) + 1)
        If sItem = s Then
            If dicTrace.Items(j) <> vbNullString Then
                '~~ Return the begin ticks and replace the value by vbNullString
                '~~ to avoid multiple recognition of the same start item
                TraceBeginTicks = dicTrace.Items(j)
                dicTrace.Items(j) = vbNullString
                Exit For
            End If
        End If
    Next j
    
End Function

Private Function TraceItemsUnPaired() As Collection
' -----------------------------------------------------
'
' -----------------------------------------------------
Dim cll As Collection
Dim i   As Long
Dim j   As Long
Dim s1  As String
Dim s2  As String
Dim k   As Long
Dim b   As Boolean

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
    Set TraceItemsUnPaired = cll

End Function

Private Function ErrSrc(ByVal sProc As String) As String
    ErrSrc = "fMsgFrm" & "." & sProc
End Function

Private Sub ErrHndlr(ByVal lErrNo As Long, _
                    ByVal sErrSource As String, _
                    ByVal sErrText As String, _
                    ByVal sErrLine As String)
' -----------------------------------------------
' This is the private error handler of the
' common error handler which cannot make use of
' itself.
' -----------------------------------------------
Dim sMsg    As String
    sMsg = "Error: " & sErrText & vbLf & vbLf & _
           "in: " & sErrSource
    If sErrLine <> 0 Then sMsg = sMsg & " (at line " & sErrLine & ")"
    MsgBox sMsg, vbCritical, "VBA Error " & lErrNo
End Sub
