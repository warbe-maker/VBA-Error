Attribute VB_Name = "mTrace"
Option Explicit
' --------------------------------------------------------------------------------
'
' Date structure of collected begin/end trace entries:
' +--------------------+------------------------+-----------+-----+
' | Entry Item         | Origin, Transformation |   Type    | Key |
' +--------------------+------------------------+-----------+-----+
' | TrcEntryNo         | Computed (Count + 1)   | Long      |  1  | not used!
' | TrcEntryTicksSys   | Collected              | Currency  |  2  |
' | TrcEntryTicksElpsd | Computed               | Currency  |     |
' | TrcEntryTicksGross | Computed               | Currency  |     |
' | TrcEntryTicksOvrhd | Collected              | Currency  |     |
' | TrcEntryTicksNet   | Computed               | Currency  |     |
' | TrcEntrySecsElpsd  | Computed               | Long      |     |
' | TrcEntrySecsGross  | Computed               | Long      |     |
' | TrcEntrySecsOvrhd  | Computed               | Long      |     |
' | TrcEntrySecsNet    | Computed               | Long      |     |
' |                    | gross - overhad        |           |     |
' | TrcEntryCallLvl    | = Number if items on   | Long      |     |
' |                    | the stack -1 after push|           |     |
' |                    | or item on stack before|           |     |
' |                    | push                   |           |     |
' | TrcEntryCallLvl    |                        |           |     |
' | TrcEntryDirective  | Collected              | String    |     |
' | TrcEntryItem       | Collected              | String    |     |
' | TrcCollecExecError | Collected              | String    |     |
' | TrcEntryError      | Computed               | String    |     |

Public Enum enTraceDisplay
    Detailed = 1
    Compact = 2
End Enum

Private Declare PtrSafe Function getFrequency Lib "kernel32" _
Alias "QueryPerformanceFrequency" (cyFrequency As Currency) As Long
Private Declare PtrSafe Function getTickCount Lib "kernel32" _
Alias "QueryPerformanceCounter" (cyTickCount As Currency) As Long

Private Const BEGIN_ID      As String = ">"                   ' Begin procedure or code trace indicator
Private Const END_ID        As String = "<"                   ' End procedure or code trace indicator
Private Const COMMENT       As String = " !!! "

Private cllTrace            As Collection   ' Collection of begin and end trace entries
Private cyFrequency         As Currency     ' Execution Trace Frequency (initialized with init)
Private lPrecision          As Long         ' Execution Trace TrcPrecision, defaults to 6 (0,000000)
Private cyTicksOvrhdTotal   As Currency     ' Execution Trace time accumulated by caused by the time tracking itself
Private cyTicksOvrhdPcntg   As Currency     ' Percentage of the above
Private dtTraceBegin        As Date         ' Initialized at start of execution trace
Private iTraceLevel         As Long         ' Increased with each begin entry and decreased with each end entry
Private sFrmtSecsElpsd      As String       ' Seconds Format
Private sFrmtSecsGross      As String       ' Seconds Format
Private sFrmtSecsOvrhdTotal As String       ' Seconds Format
Private sFrmtSecsNet        As String       ' Seconds Format
Private lTraceDisplay       As Long         ' Detailed or Compact, defaults to Compact
Private cllLastEntry        As Collection   '
Private cyTicksOvrhdLast    As Currency
Private cyTicksOvrhd        As Currency

Private Property Get CODE_BEGIN() As String:                                                CODE_BEGIN = BEGIN_ID:                  End Property

Private Property Get CODE_END() As String:                                                  CODE_END = END_ID:                      End Property

Private Property Get CurrentTicks() As Currency:                                            getTickCount CurrentTicks:              End Property

Private Property Get FrmtTicksElpsd() As String:                                          FrmtTicksElpsd = "000000.0000":       End Property

Private Property Get FrmtTicksGross() As String:                                          FrmtTicksGross = "000000.0000":       End Property

Private Property Get FrmtTicksNet() As String:                                            FrmtTicksNet = "000000.0000":         End Property

Private Property Get FrmtTicksOvrhd() As String:                                          FrmtTicksOvrhd = "0.0000":            End Property

Private Property Get FrmtTicksSys() As String:                                            FrmtTicksSys = "000000000.0000":      End Property

Private Property Get Frequency() As Currency
    If cyFrequency = 0 Then getFrequency cyFrequency
    Frequency = cyFrequency
End Property

Private Property Get FrmtSecsElpsd() As String
    If sFrmtSecsElpsd = vbNullString _
    Then sFrmtSecsElpsd = String$(ValueLength(TrcEntrySecsElpsd(cllTrace(cllTrace.Count))), "0") & "." & String$(TrcPrecision, "0")
    FrmtSecsElpsd = sFrmtSecsElpsd
End Property

Private Property Get FrmtSecsGross() As String
    If sFrmtSecsGross = vbNullString _
    Then sFrmtSecsGross = String$(ValueLength(TrcEntrySecsGross(cllTrace(cllTrace.Count))), "0") & "." & String$(TrcPrecision, "0")
    FrmtSecsGross = sFrmtSecsGross
End Property

Private Property Get FrmtSecsNet() As String
    If sFrmtSecsNet = vbNullString _
    Then sFrmtSecsNet = String$(ValueLength(TrcEntrySecsNet(cllTrace(cllTrace.Count))), "0") & "." & String$(TrcPrecision, "0")
    FrmtSecsNet = sFrmtSecsNet
End Property

Private Property Get FrmtSecsOvrhdTotal() As String
    If sFrmtSecsOvrhdTotal = vbNullString _
    Then sFrmtSecsOvrhdTotal = String$(ValueLength(TrcEntrySecsOvrhdTotal(cllTrace(cllTrace.Count))), "0") & "." & String$(TrcPrecision, "0")
    FrmtSecsOvrhdTotal = sFrmtSecsOvrhdTotal
End Property

Private Property Get PROC_BEGIN() As String:                                                PROC_BEGIN = Repeat(BEGIN_ID, 2):           End Property

Private Property Get PROC_END() As String:                                                  PROC_END = Repeat(END_ID, 2):               End Property

Private Property Get TraceDisplay() As enTraceDisplay:                                      TraceDisplay = lTraceDisplay:           End Property

Public Property Let TraceDisplay(ByVal l As enTraceDisplay):                                lTraceDisplay = l:                      End Property

Private Property Get TrcCollecExecError(Optional ByRef entry As Collection) As String
    On Error Resume Next
    TrcCollecExecError = entry("TCEE")
    If err.Number <> 0 Then TrcCollecExecError = vbNullString
End Property

Private Property Let TrcCollecExecError(Optional ByRef entry As Collection, ByRef s As String):     entry.Add s, "CEE":            End Property

Private Property Get TrcDsplyLineIndentation(Optional ByRef entry As Collection) As String
    TrcDsplyLineIndentation = Repeat("|  ", TrcEntryCallLvl(entry))
End Property

Private Property Get TrcEntryCallLvl(Optional ByRef entry As Collection) As Long:                   TrcEntryCallLvl = entry("ECL"):       End Property

Private Property Let TrcEntryCallLvl(Optional ByRef entry As Collection, ByRef l As Long):          entry.Add l, "ECL":                   End Property

Private Property Get TrcEntryDirective(Optional ByRef entry As Collection) As String:               TrcEntryDirective = entry("ED"):     End Property

Private Property Let TrcEntryDirective(Optional ByRef entry As Collection, ByRef s As String):      entry.Add s, "ED":                   End Property

Private Property Get TrcEntryError(Optional ByRef entry As Collection) As String
    On Error Resume Next ' in case this has never been collected
    TrcEntryError = entry("EE")
    If err.Number <> 0 Then TrcEntryError = vbNullString
End Property

Private Property Let TrcEntryError(Optional ByRef entry As Collection, ByRef s As String):          entry.Add s, "EE":                  End Property

Private Property Get TrcEntryItem(Optional ByRef entry As Collection) As String:                    TrcEntryItem = entry("EI"):          End Property

Private Property Let TrcEntryItem(Optional ByRef entry As Collection, ByRef s As String):           entry.Add s, "EI":                   End Property

Private Property Get TrcEntryNo(Optional ByRef entry As Collection) As Long:                        TrcEntryNo = entry("EN"):            End Property

Private Property Let TrcEntryNo(Optional ByRef entry As Collection, ByRef l As Long):               entry.Add l, "EN":                   End Property

Private Property Get TrcEntrySecsElpsd(Optional ByRef entry As Collection) As Currency
    On Error Resume Next
    TrcEntrySecsElpsd = entry("ESE")
    If err.Number <> 0 Then TrcEntrySecsElpsd = Space$(Len(FrmtSecsElpsd))
End Property

Private Property Let TrcEntrySecsElpsd(Optional ByRef entry As Collection, ByRef cy As Currency):      entry.Add cy, "ESE":                   End Property

Private Property Get TrcEntrySecsGross(Optional ByRef entry As Collection) As Currency
    On Error Resume Next ' in case no value exists (the case for each begin entry)
    TrcEntrySecsGross = entry("ESG")
    If err.Number <> 0 Then TrcEntrySecsGross = Space$(Len(FrmtSecsElpsd))
End Property

Private Property Let TrcEntrySecsGross(Optional ByRef entry As Collection, ByRef cy As Currency):        entry.Add cy, "ESG":                   End Property

Private Property Get TrcEntrySecsNet(Optional ByRef entry As Collection) As Currency
    On Error Resume Next
    TrcEntrySecsNet = entry("ESN")
    If err.Number <> 0 Then TrcEntrySecsNet = Space$(Len(FrmtSecsElpsd))
End Property

Private Property Let TrcEntrySecsNet(Optional ByRef entry As Collection, ByRef cy As Currency):     entry.Add cy, "ESN":                 End Property

Private Property Get TrcEntrySecsOvrhd(Optional ByRef entry As Collection) As Currency
    On Error Resume Next
    TrcEntrySecsOvrhd = entry("ESO")
    If err.Number <> 0 Then TrcEntrySecsOvrhd = Space$(Len(FrmtSecsElpsd))
End Property

Private Property Let TrcEntrySecsOvrhd(Optional ByRef entry As Collection, ByRef cy As Currency):      entry.Add cy, "ESO":                 End Property

Private Property Get TrcEntrySecsOvrhdTotal(Optional ByRef entry As Collection) As Currency
    On Error Resume Next
    TrcEntrySecsOvrhdTotal = entry("ESOT")
    If err.Number <> 0 Then TrcEntrySecsOvrhdTotal = Space$(Len(FrmtSecsElpsd))
End Property

Private Property Let TrcEntrySecsOvrhdTotal(Optional ByRef entry As Collection, ByRef cy As Currency):      entry.Add cy, "ESOT":                 End Property

Private Property Get TrcEntryTicksElpsd(Optional ByRef entry As Collection) As Currency:            TrcEntryTicksElpsd = entry("ETE"):   End Property

Private Property Let TrcEntryTicksElpsd(Optional ByRef entry As Collection, ByRef cy As Currency):  entry.Add cy, "ETE":                 End Property

Private Property Get TrcEntryTicksGross(Optional ByRef entry As Collection) As Currency
    On Error Resume Next
    TrcEntryTicksGross = entry("ETG")
    If err.Number <> 0 Then TrcEntryTicksGross = 0
End Property

Private Property Let TrcEntryTicksGross(Optional ByRef entry As Collection, ByRef cy As Currency):   entry.Add cy, "ETG":                 End Property

Private Property Get TrcEntryTicksNet(Optional ByRef entry As Collection) As Currency
    On Error Resume Next
    TrcEntryTicksNet = entry("ETN")
    If err.Number <> 0 Then TrcEntryTicksNet = 0
End Property

Private Property Let TrcEntryTicksNet(Optional ByRef entry As Collection, ByRef cy As Currency):    entry.Add cy, "ETN":                 End Property

Private Property Get TrcEntryTicksOvrhd(Optional ByRef entry As Collection) As Currency
    On Error Resume Next
    TrcEntryTicksOvrhd = entry("ETO")
    If err.Number <> 0 Then TrcEntryTicksOvrhd = 0
End Property

Private Property Let TrcEntryTicksOvrhd(Optional ByRef entry As Collection, ByRef cy As Currency):  entry.Add cy, "ETO":                 End Property

Private Property Get TrcEntryTicksOvrhdTotal(Optional ByRef entry As Collection) As Currency
    On Error Resume Next
    TrcEntryTicksOvrhdTotal = entry("ETOT")
    If err.Number <> 0 Then TrcEntryTicksOvrhdTotal = 0
End Property

Private Property Let TrcEntryTicksOvrhdTotal(Optional ByRef entry As Collection, ByRef cy As Currency):  entry.Add cy, "ETOT":                 End Property

Private Property Get TrcEntryTicksSys(Optional ByRef entry As Collection) As Currency:              TrcEntryTicksSys = entry("ETS"):      End Property

Private Property Let TrcEntryTicksSys(Optional ByRef entry As Collection, ByRef cy As Currency):    entry.Add cy, "ETS":                  End Property

Private Property Get TrcPrecision() As Long:                                                        TrcPrecision = lPrecision:          End Property

Public Property Let TrcPrecision(ByVal l As Long):                                                  lPrecision = l:                     End Property

Private Function ErrSrc(ByVal sProc As String) As String
    ErrSrc = "mTrace." & sProc
End Function

Private Function ExecSecs( _
                 ByVal beginticks As Currency, _
                 ByVal endticks As Currency) As String
' ----------------------------------------------------
' Returns the difference between begin- and endticks
' as formatted seconds string (decimal = nanoseconds).
' ----------------------------------------------------
    Dim cy As Currency
    If endticks - beginticks > 0 Then cy = (endticks - beginticks) / cyFrequency Else cy = 0
    ExecSecs = Format(cy, FrmtSecsElpsd)
End Function

Public Function Hca(ByVal s1 As String, _
                     ByVal s2 As String) As String
' ------------------------------------------------
' Returns the string (s1) centered above s2.
' ------------------------------------------------
    
    If Len(s1) > Len(s2) Then
        Hca = s1
    Else
        Hca = s1 & Space$(Int((Len(s2) - Len(s1)) / 2))
        Hca = Space$(Len(s2) - Len(Hca)) & Hca
    End If
    
End Function

Public Sub Initialize()
    If Not cllTrace Is Nothing Then
        If cllTrace.Count <> 0 Then Exit Sub ' already initialized
    Else
        Set cllTrace = New Collection
        Set cllLastEntry = Nothing
    End If
    dtTraceBegin = Now()
    cyTicksOvrhdTotal = 0
    iTraceLevel = 0
    If lPrecision = 0 Then lPrecision = 6 ' default when another valus has not been set by the caller
    
End Sub

Private Function Repeat( _
                 ByVal s As String, _
                 ByVal r As Long) As String
' -------------------------------------------
' Returns the string (s) repeated (r) times.
' -------------------------------------------
    Dim i   As Long
    
    For i = 1 To r
        Repeat = Repeat & s
    Next i
    
End Function

Private Sub SecsElapsed()
' ----------------------------------------------------------
' Calculate the elapsed seconds based on the elapsed ticks.
' ----------------------------------------------------------
    Dim v               As Variant
    Dim cll             As Collection
    
    For Each v In cllTrace
        Set cll = v
        TrcEntrySecsElpsd(cll) = Format((TrcEntryTicksElpsd(cll) / Frequency), FrmtSecsElpsd)
    Next v
    Set cll = Nothing

End Sub

Private Sub SecsExecuted()
' ----------------------------------------------------------
' Calculate the executed seconds gross and net based on the
' correponding ticks.
' ----------------------------------------------------------
    Dim v               As Variant
    Dim cll             As Collection
    
    For Each v In cllTrace
        Set cll = v
        If Not TrcEntryIsBegin(cll) Then
            TrcEntrySecsGross(cll) = TrcEntryTicksGross(cll) / Frequency
            TrcEntrySecsOvrhdTotal(cll) = TrcEntryTicksOvrhdTotal(cll) / Frequency
            TrcEntrySecsNet(cll) = TrcEntryTicksNet(cll) / Frequency
        End If
    Next v
    Set cll = Nothing

End Sub

Private Sub TicksElapsed()
' ---------------------------------------------
' Calculate the elapsed seconds for each entry.
' ---------------------------------------------
    Dim v               As Variant
    Dim cll             As Collection
    Dim cyTicksBegin    As Currency
    Dim cyTicksElapsed  As Currency
    
    For Each v In cllTrace
        Set cll = v
        If cyTicksBegin = 0 Then cyTicksBegin = TrcEntryTicksSys(cll)
        cyTicksElapsed = TrcEntryTicksSys(cll) - cyTicksBegin
        TrcEntryTicksElpsd(cll) = cyTicksElapsed
    Next v
    Set cll = Nothing

End Sub

Private Sub TicksExecNet()
' ---------------------------------------------
' Calculate the net executed ticks by deducting
' the overhad ticks from the exec ticks.
' ---------------------------------------------
    Dim v               As Variant
    Dim cll             As Collection
    Dim cyTicksBegin    As Currency
    Dim cyTicksElapsed  As Currency
    
    For Each v In cllTrace
        Set cll = v
        TrcEntryTicksNet(cll) = TrcEntryTicksGross(cll) - TrcEntryTicksOvrhd(cll)
    Next v
    Set cll = Nothing

End Sub

Public Sub TrcCodeBegin(ByVal s As String): TrcCollectBegin s, CODE_BEGIN: End Sub

Public Sub TrcCodeEnd(ByVal s As String):   TrcCollectEnd s, CODE_END:     End Sub

Private Sub TrcCollectAdd( _
            ByVal itm As String, _
            ByVal ticks As Currency, _
            ByVal dir As String, _
            ByVal lvl As Long, _
   Optional ByVal err As String = vbNullString)
   
    If Not cllLastEntry Is Nothing Then
        TrcEntryTicksOvrhd(cllLastEntry) = cyTicksOvrhdLast
    End If
    
    Set cllLastEntry = TrcEntry(no:=cllTrace.Count + 1, ticks:=ticks, dir:=dir, itm:=itm, lvl:=lvl, err:=err)
    cyTicksOvrhdLast = cyTicksOvrhd
    
    cllTrace.Add cllLastEntry
    
End Sub

Private Sub TrcCollectBegin( _
            ByVal itm As String, _
            ByVal dir As String)
' --------------------------------
' Collect a trace begin entry with
' the current ticks count for the
' procedure or code (item).
' --------------------------------
    Dim cy As Currency: cy = CurrentTicks
    iTraceLevel = iTraceLevel + 1
    TrcCollectAdd itm:=itm, ticks:=cy, dir:=Trim(dir), lvl:=iTraceLevel
    cyTicksOvrhd = CurrentTicks - cy
    cyTicksOvrhdTotal = cyTicksOvrhdTotal + cyTicksOvrhd
    
End Sub

Private Sub TrcCollectEnd( _
            ByVal s As String, _
   Optional ByVal dir As String = vbNullString, _
   Optional ByVal errinf As String = vbNullString)
' ------------------------------------------------
' Collect an end trace entry with the current
' ticks count for the procedure or code (item).
' ------------------------------------------------
    
    On Error GoTo on_error
    Const PROC = "TrcCollectEnd"
    Dim cy As Currency: cy = CurrentTicks
    
    If errinf <> vbNullString Then
        errinf = COMMENT & errinf & COMMENT
    End If
    
    TrcCollectAdd itm:=s, ticks:=cy, dir:=Trim(dir), lvl:=iTraceLevel, err:=errinf
    iTraceLevel = iTraceLevel - 1
    cyTicksOvrhdTotal = cyTicksOvrhdTotal + (CurrentTicks - cy)

exit_proc:
    Exit Sub
    
on_error:
    MsgBox err.Description, vbOKOnly, "Error in " & ErrSrc(PROC)
End Sub

Public Sub TrcDsply()
    
    On Error GoTo on_error
    Const PROC = "TrcDsply"
    Dim dct         As Dictionary
    Dim v           As Variant
    Dim entry       As Collection
    Dim sTrace      As String
    Dim lLenHeader  As Long
    
    TrcEntryTicksOvrhd(cllLastEntry) = cyTicksOvrhdLast

    If TraceDisplay = 0 Then TraceDisplay = Compact
    TrcDsplyEndStack ' end procedures still on the stack

    If Not TrcEntryAllConsistent(dct) Then
        For Each v In dct
            sTrace = sTrace & dct(v) & v & vbLf
        Next v
        With fMsg
            .MsgTitle = "Inconsistent begin/end trace code lines!"
            .MsgLabel(1) = "The following incositencies had been detected which made the display of the execution trace result useless:"
            .MsgText(1) = sTrace:   .MsgMonoSpaced(1) = True
            .Setup
            .Show
        End With
        Set dct = Nothing
        GoTo exit_proc
    End If
        
    TicksExecNet
    SecsExecuted ' Calculate gross and net execution seconds
    
    sTrace = TrcDsplyHeader(lLenHeader)
    For Each v In cllTrace
        Set entry = v
        sTrace = sTrace & vbLf & TrcDsplyLine(entry)
    Next v
    sTrace = sTrace & vbLf & TrcDsplyFooter(lLenHeader)
    With fMsg
        .MaxFormWidthPrcntgOfScreenSize = 95
        .MsgTitle = "Execution Trace, displayed because the Conditional Compile Argument ""ExecTrace = 1""!"
        .MsgText(1) = sTrace:   .MsgMonoSpaced(1) = True
        .MsgLabel(2) = "About overhead, precision, etc.:": .MsgText(2) = TrcDsplyAbout
        .Setup
        .Show
    End With
    
    
exit_proc:
    Set cllTrace = Nothing
    Exit Sub
    
on_error:
    MsgBox err.Description, vbOKOnly, "Error in " & ErrSrc(PROC)
    Stop: Resume
    Set cllTrace = Nothing
End Sub

Private Function TrcDsplyAbout() As String
    TrcDsplyAbout = "The trace itself caused an overhead (performance loss respectively) of " _
      & Format(ExecSecs(0, cyTicksOvrhdTotal), FrmtSecsElpsd) & "(" & cyTicksOvrhdPcntg * 100 & "%)" _
      & " seconds. For a best possible execution time precision the average overhead of " _
      & Format((cyTicksOvrhdTotal / (cllTrace.Count / 2)) / cyFrequency, FrmtSecsElpsd) & " secs " _
      & "had been deducted from each of the " & cllTrace.Count / 2 & " traced item's execution time." _
      & vbLf _
      & "The displayed precision of the trace defaults to 0,000000 (6 decimals) which may " _
      & "be changed via the property ""TrcPrecision""." _
      & vbLf _
      & "The displayed execution time varies from execution to execution and can only be estimated " _
      & "as an average of many executions." _
      & vbLf _
      & "When an error had been displayed the traced execution time includes the time " _
      & "of the user reaction and thus does not provide a meaningful result."

End Function

Private Sub TrcDsplyEndStack()
' ---------------------------------
' Completes the execution trace for
' items still on the stack.
' ---------------------------------
    Dim s As String
    Do Until mErrHndlr.StackIsEmpty
        s = mErrHndlr.StackPop
        If s <> vbNullString Then
            TrcCollectEnd s
        End If
    Loop
End Sub

Private Function TrcDsplyFooter(ByVal lLenHeaderData As Long)
    TrcDsplyFooter = _
        Space$(lLenHeaderData) _
      & PROC_END _
      & " End execution trace " _
      & Format(Now(), "hh:mm:ss")
End Function

Private Function TrcDsplyHeader( _
                 ByRef lLenHeaderData As Long) As String
' ------------------------------------------------------
' Compact:  Header1 = Elapsed Exec >> Start
' Detailed: Header1 =    Ticks   |   Secs
'           Header2 = xxx xxxx xx xxx xxx xx
'           Header3 " --- ---- -- --- --- -- >> Start
' ------------------------------------------------------
    On Error GoTo eh
    Const PROC = "TrcDsplyHeader"
    Dim sIndent         As String: sIndent = Space$(Len(FrmtSecsElpsd))
    Dim sHeader1        As String
    Dim sHeader2        As String
    Dim sHeader2Ticks   As String
    Dim sHeader2Secs    As String
    Dim sHeader3        As String
    Dim sHeaderData1    As String
    Dim sHeaderData2    As String
    Dim sHeaderTrace    As String
    
    sHeaderTrace = _
        PROC_BEGIN _
      & " Begin execution trace " _
      & Format(dtTraceBegin, "hh:mm:ss") _
      & " (exec time in seconds)"

    Select Case TraceDisplay
        Case Compact
            TrcDsplyHeader = _
                "Elapsed" _
              & Space$(Len(sIndent) - Len("Elapsed") + 1) _
              & "Exec secs " _
              & sHeaderTrace _
              & vbLf
        
        Case Detailed
            sHeader2Ticks = _
                Hca("System", FrmtTicksSys) & " " _
              & Hca("Elapsed", FrmtTicksElpsd) & " " _
              & Hca("Gross", FrmtTicksGross) & " " _
              & Hca("Ovrhd", FrmtTicksOvrhd) & " " _
              & Hca("Net", FrmtTicksNet)
            sHeader2Secs = _
                Hca("Elapsed", FrmtSecsElpsd) & " " _
              & Hca("Gross", FrmtSecsElpsd) & " " _
              & Hca("Ovrhd", FrmtSecsElpsd) & " " _
              & Hca("Net", FrmtSecsElpsd)
            sHeader2 = sHeader2Ticks & " " & sHeader2Secs & " "
            lLenHeaderData = Len(sHeader2)
            
            sHeader3 = _
                Repeat$("-", Len(FrmtTicksSys)) & " " _
              & Repeat$("-", Len(FrmtTicksElpsd)) & " " _
              & Repeat$("-", Len(FrmtTicksGross)) & " " _
              & Repeat$("-", Len(FrmtTicksOvrhd)) & " " _
              & Repeat$("-", Len(FrmtTicksNet)) & " " _
              & Repeat$("-", Len(FrmtSecsElpsd)) & " " _
              & Repeat$("-", Len(FrmtSecsElpsd)) & " " _
              & Repeat$("-", Len(FrmtSecsElpsd)) & " " _
              & Repeat$("-", Len(FrmtSecsElpsd)) & " " _
              & sHeaderTrace
              
            sHeader1 = _
                Hca("Ticks", sHeader2Ticks) & "|" _
              & Hca("Seconds", sHeader2Secs)
            
            TrcDsplyHeader = _
                sHeader1 & vbLf _
              & sHeader2 & vbLf _
              & sHeader3
            
    End Select
    Exit Function
eh:
    MsgBox err.Description, vbOKOnly, "Error in " & ErrSrc(PROC)
    Stop: Resume
    Set cllTrace = Nothing
End Function

Private Function TrcDsplyLine(ByVal entry As Collection) As String
' -------------------------------------------------------------
' Returns a trace line for being displayed.
' -------------------------------------------------------------
    On Error GoTo eh
    Const PROC = "TrcDsplyLine"
    
    Select Case TraceDisplay
        Case Compact
            TrcDsplyLine = _
                TrcDsplyValue(entry, TrcEntrySecsElpsd(entry), FrmtSecsElpsd) _
              & " " _
              & TrcDsplyValue(entry, TrcEntrySecsGross(entry), FrmtSecsGross) _
              & " " _
              & TrcDsplyLineIndentation(entry) _
              & TrcEntryDirective(entry) _
              & " " _
              & TrcEntryItem(entry) _
              & " " _
              & TrcEntryError(entry)
        Case Detailed
            TrcDsplyLine = _
                  Format(TrcEntryTicksSys(entry), FrmtTicksSys) _
                & " " _
                & Format(TrcEntryTicksElpsd(entry), FrmtTicksElpsd) _
                & " " _
                & IIf(TrcEntryTicksGross(entry) > 0, Format(TrcEntryTicksGross(entry), FrmtTicksGross), Space$(Len(FrmtTicksGross))) _
                & " " _
                & IIf(TrcEntryTicksOvrhdTotal(entry) > 0, Format(TrcEntryTicksOvrhdTotal(entry), FrmtTicksOvrhd), Space$(Len(FrmtTicksOvrhd))) _
                & " " _
                & IIf(TrcEntryTicksNet(entry) > 0, Format(TrcEntryTicksNet(entry), FrmtTicksNet), Space$(Len(FrmtTicksNet))) _
                & " " _
                & TrcDsplyValue(entry, TrcEntrySecsElpsd(entry), FrmtSecsElpsd) _
                & " " _
                & TrcDsplyValue(entry, TrcEntrySecsGross(entry), FrmtSecsGross) _
                & " " _
                & TrcDsplyValue(entry, TrcEntrySecsOvrhdTotal(entry), FrmtSecsOvrhdTotal) _
                & " " _
                & TrcDsplyValue(entry, TrcEntrySecsNet(entry), FrmtSecsNet) _
                & " " _
                & TrcDsplyLineIndentation(entry) _
                & TrcEntryDirective(entry) _
                & " " _
                & TrcEntryItem(entry) _
                & " " _
                & TrcEntryError(entry)
    End Select
    Exit Function
eh:
    MsgBox err.Description, vbOKOnly, "Error in " & ErrSrc(PROC)
    Stop: Resume
    Set cllTrace = Nothing
End Function

Private Function TrcDsplyValue(ByVal entry As Collection, _
    ByVal value As Variant, _
    ByVal frmt As String) As String
    If TrcEntryIsBegin(entry) _
    Then TrcDsplyValue = Space$(Len(frmt)) _
    Else TrcDsplyValue = IIf(value > 0, Format(value, frmt), Space$(Len(frmt)))
End Function

Private Function TrcEntry( _
      ByVal no As Long, _
      ByVal ticks As Currency, _
      ByVal dir As String, _
      ByVal itm As String, _
      ByVal lvl As Long, _
      ByVal err As String) As Collection
      
    Dim entry As New Collection
    
    TrcEntryNo(entry) = no
    TrcEntryTicksSys(entry) = ticks
    TrcEntryDirective(entry) = dir
    TrcEntryItem(entry) = itm
    TrcEntryCallLvl(entry) = lvl
    TrcEntryError(entry) = err
    Set TrcEntry = entry
End Function

Private Function TrcEntryAllConsistent(ByRef dct As Dictionary) As Boolean
' ------------------------------------------------------------------------
' Returns TRUE when for each begin entry there is a corresponding end
' entry and vice versa. Else the function returns FALSE and the items
' without a corresponding counterpart are returned as Dictionary (dct).
' The consistency check is based on the calculated execution ticks as the
' difference between the elapsed begin and end ticks.
' ------------------------------------------------------------------------
    On Error GoTo eh
    Const PROC = "TrcEntryAllConsistent"
    Dim v                   As Variant
    Dim cllEndEntry         As Collection
    Dim cllBeginEntry       As Collection
    Dim bConsistent         As Boolean
    Dim cyBeginTicks        As Currency
    Dim cllEntry            As Collection
    Dim i                   As Long
    Dim j                   As Long
    Dim sComment            As String
    Dim cyAvrgOvrhdTicks    As Currency:    cyAvrgOvrhdTicks = cyTicksOvrhdTotal / (cllTrace.Count / 2)
    Dim cyTicksTotal        As Currency:    cyTicksTotal = TrcEntryTicksSys(cllTrace(cllTrace.Count)) - TrcEntryTicksSys(cllTrace(1))
    Dim cyTicksExec         As Currency
    Dim cyTicksNet          As Currency
    Dim cyTicks             As Currency
    
    If dct Is Nothing Then Set dct = New Dictionary
    
    cyTicksOvrhdPcntg = cyTicksOvrhdTotal / cyTicksTotal
    
    TicksElapsed ' Calculates the ticks elapsed since trace start
    
    '~~ Check for missing corresponding end entries while calculating the execution time for each end entry.
    For i = 1 To TrcEntryLastBegin
        If TrcEntryIsBegin(cllTrace(i), cllBeginEntry) Then
            bConsistent = False
            For j = i + 1 To cllTrace.Count
                If TrcEntryIsEnd(cllTrace(j), cllEndEntry) Then
                    If TrcEntryItem(cllBeginEntry) = TrcEntryItem(cllEndEntry) Then
                        If TrcEntryCallLvl(cllBeginEntry) = TrcEntryCallLvl(cllEndEntry) Then
                            '~~ Calculate the executesd seconds for the end entry
                            TrcEntryTicksGross(cllEndEntry) = TrcEntryTicksElpsd(cllEndEntry) - TrcEntryTicksElpsd(cllBeginEntry)
                            TrcEntryTicksOvrhdTotal(cllEndEntry) = TrcEntryTicksOvrhd(cllEndEntry) + TrcEntryTicksOvrhd(cllBeginEntry)
                            GoTo next_begin_entry
                        End If
                    End If
                End If
            Next j
            '~~ No corresponding end entry found
            Select Case TrcEntryDirective(cllBeginEntry)
                Case PROC_BEGIN: sComment = "No corresponding End of Procedure (EoP) code line in:    "
                Case CODE_BEGIN: sComment = "No corresponding End of CodeTrace (EoC) code line in:    "
            End Select
            If Not dct.Exists(TrcEntryItem(cllBeginEntry)) Then dct.Add TrcEntryItem(cllBeginEntry), sComment
        End If

next_begin_entry:
    Next i
    
    '~~ Check for missing corresponding begin entries (if the end entry has no executed ticks)
    For Each v In cllTrace
        If TrcEntryIsEnd(v, cllEndEntry) Then
            If TrcEntryTicksGross(cllEndEntry) = 0 Then
                Debug.Print TrcEntryItem(cllBeginEntry)
                '~~ No corresponding begin entry found
                Select Case TrcEntryDirective(cllEndEntry)
                    Case PROC_END: sComment = "No corresponding Begin of Procedure (BoP) code line in:  "
                    Case CODE_END: sComment = "No corresponding Begin of CodeTrace (BoC) code line in:  "
                End Select
                If Not dct.Exists(TrcEntryItem(cllBeginEntry)) Then dct.Add TrcEntryItem(cllBeginEntry), sComment
            End If
        End If
    Next v
    
    TrcEntryAllConsistent = dct.Count = 0
    Exit Function

eh:
    MsgBox err.Description, vbOKOnly, "Error in " & ErrSrc(PROC)
    Stop: Resume
End Function

 Private Function TrcEntryIsBegin( _
                  ByVal v As Collection, _
         Optional ByRef cll As Collection = Nothing) As Boolean
' -------------------------------------------------------------
' Returns TRUE and v as cll when the entry is a begin entry,
' else FALSE and cll = Nothing
' ---------------------------------------------------
    If InStr(TrcEntryDirective(v), BEGIN_ID) <> 0 Then
        TrcEntryIsBegin = True
        Set cll = v
    End If
End Function

Private Function TrcEntryIsEnd( _
                 ByVal v As Collection, _
                 ByRef cll As Collection) As Boolean
' --------------------------------------------------
' Returns TRUE and v as cll when the entry is an end
' entry, else FALSE and cll = Nothing
' --------------------------------------------------
    If InStr(TrcEntryDirective(v), END_ID) <> 0 Then
        TrcEntryIsEnd = True
        Set cll = v
    End If
End Function

Private Function TrcEntryLastBegin() As Long
    
    Dim i As Long
    
    For i = cllTrace.Count To 1 Step -1
        If TrcEntryIsBegin(cllTrace(i)) Then
            TrcEntryLastBegin = i
            Exit Function
        End If
    Next i
    
End Function

Public Sub TrcProcBegin(ByVal s As String)
    TrcCollectBegin s, PROC_BEGIN
End Sub

Public Sub TrcProcEnd(ByVal s As String, _
             Optional ByVal errinfo As String = vbNullString)
    TrcCollectEnd s, PROC_END, errinfo
    If errinfo <> vbNullString Then
        Debug.Print errinfo
    End If
End Sub

Private Function ValueLength(ByVal v As Variant) As Long

    Select Case v
        Case Is >= 100000000:   ValueLength = 9
        Case Is >= 10000000:    ValueLength = 8
        Case Is >= 1000000:     ValueLength = 7
        Case Is >= 100000:      ValueLength = 6
        Case Is >= 10000:       ValueLength = 5
        Case Is >= 1000:        ValueLength = 4
        Case Is >= 100:         ValueLength = 3
        Case Is >= 10:          ValueLength = 2
        Case Else:              ValueLength = 1
    End Select

End Function

