Attribute VB_Name = "mTrace"
Option Explicit
' --------------------------------------------------------------------------------
'
' Structure of a collected trace entry:
' -------------------------------------
' |   Entry Item   | Origin, Transformation |   Type    |
' | ---------------| -----------------------| --------- |
' | EntryNumber    | = Count + 1,           | Long      |
' |                | before add Index       |           |
' | TickCount      | Collected              | Currency  |
' | ElapsedSeconds | Computed               | Long      |
' | ExecSeconds    | Computed               | Long      |
' | CallLevel      | = Number if items on   | Long      |
' |                | the stack -1 after push|           |
' |                | or item on stack before|           |
' |                | push                   |           |
' | Indentation    | Computed based on      | String    |
' |                | CallLevel              |           |
' | DirectiveId    | Collected              | String    |
' | Procedure      | Collected              | String    |
' | ExecErrorInfo  | Collected              | String    |
' | TraceErrorInfo | Computed               | String    |

' --- Begin of declaration for the Execution Tracing
Private Declare PtrSafe Function getFrequency Lib "kernel32" _
Alias "QueryPerformanceFrequency" (cyFrequency As Currency) As Long
Private Declare PtrSafe Function getTickCount Lib "kernel32" _
Alias "QueryPerformanceCounter" (cyTickCount As Currency) As Long

Private Const TRACE_BEGIN_ID    As String = ">"                   ' Begin procedure or code trace indicator
Private Const TRACE_END_ID      As String = "<"                   ' End procedure or code trace indicator
Private Const TRACE_COMMENT     As String = " !!! "


Private dicTrace            As Dictionary   ' For the collection of execution trance entries/lines
Private cllTrace            As Collection
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

Private Property Get INCOMPLETE_TRACE() As String:      INCOMPLETE_TRACE = TRACE_COMMENT & "Incomplete trace" & TRACE_COMMENT:              End Property

Private Property Get TRACE_CODE_BEGIN_ID() As String:   TRACE_CODE_BEGIN_ID = TRACE_BEGIN_ID & " ":                                         End Property

Private Property Get TRACE_CODE_END_ID() As String:     TRACE_CODE_END_ID = TRACE_END_ID & " ":                                             End Property

Private Property Get TRACE_PROC_BEGIN_ID() As String:   TRACE_PROC_BEGIN_ID = TRACE_BEGIN_ID & TRACE_BEGIN_ID & " ":                        End Property

Private Property Get TRACE_PROC_END_ID() As String:     TRACE_PROC_END_ID = TRACE_END_ID & TRACE_END_ID & " ":                              End Property

Private Function ErrSrc(ByVal sProc As String) As String
    ErrSrc = "mTrace." & sProc
End Function

Private Function Replicate( _
                 ByVal s As String, _
                 ByVal r As Long) As String
' -------------------------------------------
' Returns the string (s) repeated (r) times.
' -------------------------------------------
    Dim i   As Long
    
    For i = 1 To r
        Replicate = Replicate & s
    Next i
    
End Function

Private Sub TrcAdd(ByVal s As String, _
                   ByVal cy As Currency)
                   
    iTraceItem = iTraceItem + 1
    dicTrace.Add iTraceItem & s, cy
'    Debug.Print "Added to Trace: '" & iTraceItem & s
    
End Sub

Private Function TrcBegEndId(ByVal s As String) As String
    TrcBegEndId = Split(s, " ")(0) & " "
End Function

Private Sub TrcBegin(ByVal s As String, _
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

Public Sub TrcCodeBegin(ByVal s As String)
    TrcBegin s, TRACE_CODE_BEGIN_ID
End Sub

Public Sub TrcCodeEnd(ByVal s As String)
    TrcEnd s, TRACE_CODE_END_ID
End Sub

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
           
    If dicTrace Is Nothing Then Exit Function   ' When the contional compile argument where not
    If dicTrace.Count = 0 Then Exit Function    ' ExecTrace = 1 there will be no execution trace result

    TrcEndStack ' end procedures still on the stack
    
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

Public Sub TrcEnd(ByVal s As String, _
         Optional ByVal id As String = vbNullString, _
         Optional ByVal errinfo As String = vbNullString)
' --------------------------------------------------------
' End of Trace. Keeps a record of the ticks count for the
' execution trace of the group of code lines named (s).
' --------------------------------------------------------
#If ExecTrace Then
    
    On Error GoTo on_error
    Const PROC = "TrcEnd"
    Dim cy As Currency
        
    getTickCount cyTicks
    cy = cyTicks
    If errinfo <> vbNullString Then
        errinfo = TRACE_COMMENT & errinfo & TRACE_COMMENT
    End If
    TrcAdd id & s & errinfo, cyTicks
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

Private Sub TrcEndStack()
' --------------------------------------
' Completes the execution trace for
' items still on the stack.
' --------------------------------------
    Dim s As String
    Do Until StackIsEmpty
        s = StackPop
        If s <> vbNullString Then
            TrcEnd s
        End If
    Loop
End Sub

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

Private Function TrcItem(ByVal s As String) As String
' ---------------------------------------------------
' Returns the item (i.e. the traced ErrSrc()) element
' within the trace entry.
' Precondition: The ErrSrc() must not contain spaces
' ---------------------------------------------------
    TrcItem = Split(s, " ")(1)
End Function

Public Sub TrcProcBegin(ByVal s As String)
    TrcBegin s, TRACE_PROC_BEGIN_ID
End Sub

Public Sub TrcProcEnd(ByVal s As String, _
             Optional ByVal errinfo As String = vbNullString)
            
    TrcEnd s, TRACE_PROC_END_ID, errinfo
End Sub

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

Private Sub CllTrcAdd( _
      ByVal v1 As Variant, _
      ByVal v2 As Variant, _
      ByVal v3 As Variant, _
      ByVal v4 As Variant, _
      ByVal v5 As Variant, _
      ByVal v6 As Variant, _
      ByVal v7 As Variant, _
      ByVal v8 As Variant, _
      ByVal v9 As Variant)
            
   cllTrace.Add CllTrcEntry(v1, v2, v3, v4, v5, v6, v7, v8, v9)
   
End Sub

Private Function CllTrcEntry( _
           ByVal v1 As Variant, _
           ByVal v2 As Variant, _
           ByVal v3 As Variant, _
           ByVal v4 As Variant, _
           ByVal v5 As Variant, _
           ByVal v6 As Variant, _
           ByVal v7 As Variant, _
           ByVal v8 As Variant, _
           ByVal v9 As Variant) As Collection
           
    Dim cll As New Collection
    
    cll.Add v1
    cll.Add v2
    cll.Add v3
    cll.Add v4
    cll.Add v5
    cll.Add v6
    cll.Add v7
    cll.Add v8
    cll.Add v9
    Set CllTrcEntry = cll
End Function

Private Property Get TrcEntryNo(Optional ByVal entry As Collection) As Long:        TrcEntryNo = entry(1):      End Property
Private Property Get TrcTickCount(Optional ByVal entry As Collection) As Currency:  TrcTickCount = entry(2):    End Property
'Private Property Get TrcElapsedSeconds | Computed               | Long      |
'Private Property Get TrcExecSeconds    | Computed               | Long      |
'Private Property Get TrcCallLevel      | = Number if items on   | Long      |
'Private Property Get Trc               | the stack -1 after push|           |
'Private Property Get Trc               | or item on stack before|           |
'Private Property Get Trc               | push                   |           |
'Private Property Get TrcIndentation    | Computed based on      | String    |
'Private Property Get Trc               | CallLevel              |           |
'Private Property Get TrcDirectiveId    | Collected              | String    |
'Private Property Get TrcProcedure      | Collected              | String    |
'Private Property Get TrcExecErrorInfo  | Collected              | String    |
'Private Property Get TrcTraceErrorInfo | Computed               | String    |

Private Sub CllTrcDsply()
    Dim v As Variant
    For Each v In cllTrace
'        If TrcBeginEntry(v) Then
'            If Not TrcHasEndEntry(v) Then
'                TrcAddTraceErrInfo v, "end entry missing"
'                ' no indentation! i'll never go left
'            Else
'               ' Indentation is due
'            End If
''        ElseIf DirectiveId(v)
'
'        End If
    Next v
    
 End Sub
 
 Private Function TrcBeginEntry(v As Collection) As Boolean
    Select Case TrcDirectiveId(v)
        Case TRACE_PROC_BEGIN_ID, TRACE_CODE_BEGIN_ID
            TrcBeginEntry = True
    End Select
End Function

Private Function TrcDirectiveId(ByVal cll As Collection) As String
    TrcDirectiveId = cll(7)
End Function

