Attribute VB_Name = "mTrace"
Option Explicit
' --------------------------------------------------------------------------------
'
' Structure of a collected trace entry:
' ---------------------------------------------------------
' | Entry Item         | Origin, Transformation |   Type    |
' | -------------------| -----------------------| --------- |
' | TrcEntryNo         | = Count + 1,           | Long      |
' |                    | before add Index       |           |
' | Ticks              | Collected              | Currency  |
' | TrcEntryElpsdSecs  | Computed               | Long      |
' | TrcEntryExcSecs    | Computed               | Long      |
' | TrcEntryCallLvl    | = Number if items on   | Long      |
' |                    | the stack -1 after push|           |
' |                    | or item on stack before|           |
' |                    | push                   |           |
' | TrcEntryCallLvl    |                        |           |
' | TrcEntryDirective  | Collected              | String    |
' | TrcEntryItem       | Collected              | String    |
' | TrcCollecExecError | Collected              | String    |
' | TrcEntryError      | Computed               | String    |

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
Private cyOverheadTicks     As Currency     ' Execution Trace time accumulated by caused by the time tracking itself
Private cyOverheadTicksPcntg As Currency    ' Percentage of the above
Private dtTraceBegin        As Date         ' Initialized at start of execution trace
Private iTraceLevel         As Long         ' Increased with each begin entry and decreased with each end entry
Private sSecsFormat         As String       ' Format for elapsed and executed seconds

Private Property Get CODE_BEGIN() As String:                                                    CODE_BEGIN = BEGIN_ID:          End Property

Private Property Get CODE_END() As String:                                                      CODE_END = END_ID:              End Property

Private Property Get CurrentTicks() As Currency:                                                getTickCount CurrentTicks:      End Property

Private Property Get ExecSecsFrmt() As String
    Dim i As Long
    
    If sSecsFormat = vbNullString Then
        Select Case (TrcEntryTicks(cllTrace(cllTrace.Count)) - TrcEntryTicks(cllTrace(1))) / Frequency
            Case Is >= 100000:  i = 6
            Case Is >= 10000:   i = 5
            Case Is >= 1000:    i = 4
            Case Is >= 100:     i = 3
            Case Is >= 10:      i = 2
            Case Else:          i = 1
        End Select
        sSecsFormat = String$(i - 1, "0") & "0." & String$(TrcPrecision, "0") & " "
    End If
    ExecSecsFrmt = sSecsFormat
    
End Property

Private Property Get Frequency() As Currency
    If cyFrequency = 0 Then getFrequency cyFrequency
    Frequency = cyFrequency
End Property

Private Property Get PROC_BEGIN() As String:                                                PROC_BEGIN = Repeat(BEGIN_ID, 2):   End Property

Private Property Get PROC_END() As String:                                                  PROC_END = Repeat(END_ID, 2):       End Property

Private Property Get TrcCollecExecError(Optional ByRef entry As Collection) As String
    On Error Resume Next
    TrcCollecExecError = entry("9")
    If err.Number <> 0 Then TrcCollecExecError = vbNullString
End Property

Private Property Let TrcCollecExecError(Optional ByRef entry As Collection, ByRef s As String): entry.Add s, "9":               End Property

Private Property Get TrcDsplyLineIndentation(Optional ByRef entry As Collection) As String
    TrcDsplyLineIndentation = Repeat("|  ", TrcEntryCallLvl(entry))
End Property

Private Property Get TrcEntryCallLvl(Optional ByRef entry As Collection) As Long:               TrcEntryCallLvl = entry("5"):   End Property

Private Property Let TrcEntryCallLvl(Optional ByRef entry As Collection, ByRef l As Long):      entry.Add l, "5":               End Property

Private Property Get TrcEntryDirective(Optional ByRef entry As Collection) As String:           TrcEntryDirective = entry("7"): End Property

Private Property Let TrcEntryDirective(Optional ByRef entry As Collection, ByRef s As String):  entry.Add s, "7":               End Property

Private Property Get TrcEntryElpsdSecs(Optional ByRef entry As Collection) As String:           TrcEntryElpsdSecs = entry("3"): End Property

Private Property Let TrcEntryElpsdSecs(Optional ByRef entry As Collection, ByRef s As String):  entry.Add s, "3":               End Property

Private Property Get TrcEntryError(Optional ByRef entry As Collection) As String
    On Error Resume Next ' in case this has never been collected
    TrcEntryError = entry("10")
    If err.Number <> 0 Then TrcEntryError = vbNullString
End Property

Private Property Let TrcEntryError(Optional ByRef entry As Collection, ByRef s As String):      entry.Add s, "10":              End Property

Private Property Get TrcEntryExcSecs(Optional ByRef entry As Collection) As String
    On Error Resume Next ' in case no value exists (the case for each begin entry)
    TrcEntryExcSecs = entry("4")
    If err.Number <> 0 Then TrcEntryExcSecs = Space$(Len(ExecSecsFrmt))
End Property

Private Property Let TrcEntryExcSecs(Optional ByRef entry As Collection, ByRef s As String):    entry.Add s, "4":               End Property

Private Property Get TrcEntryItem(Optional ByRef entry As Collection) As String:                TrcEntryItem = entry("8"):      End Property

Private Property Let TrcEntryItem(Optional ByRef entry As Collection, ByRef s As String):       entry.Add s, "8":               End Property

Private Property Get TrcEntryNo(Optional ByRef entry As Collection) As Long:                    TrcEntryNo = entry("1"):        End Property

Private Property Let TrcEntryNo(Optional ByRef entry As Collection, ByRef l As Long):           entry.Add l, "1":               End Property

Private Property Get TrcEntryTicks(Optional ByRef entry As Collection) As Currency:             TrcEntryTicks = entry("2"):     End Property

Private Property Let TrcEntryTicks(Optional ByRef entry As Collection, ByRef cy As Currency):   entry.Add cy, "2":              End Property

Private Property Get TrcPrecision() As Long:                                                    TrcPrecision = lPrecision:      End Property

Public Property Let TrcPrecision(ByVal l As Long):                                              lPrecision = l:                 End Property

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
    ExecSecs = Format(cy, ExecSecsFrmt)
End Function

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

Public Sub TrcCodeBegin(ByVal s As String): TrcCollectBegin s, CODE_BEGIN: End Sub

Public Sub TrcCodeEnd(ByVal s As String):   TrcCollectEnd s, CODE_END:     End Sub

Private Sub TrcCollectAdd( _
            ByVal itm As String, _
            ByVal ticks As Currency, _
            ByVal dir As String, _
            ByVal lvl As Long, _
   Optional ByVal err As String = vbNullString)
    cllTrace.Add TrcEntry(no:=cllTrace.Count + 1, ticks:=ticks, dir:=dir, itm:=itm, lvl:=lvl, err:=err)
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
    TrcCollectInit
    iTraceLevel = iTraceLevel + 1
    TrcCollectAdd itm:=itm, ticks:=cy, dir:=Trim(dir), lvl:=iTraceLevel
    cyOverheadTicks = cyOverheadTicks + (CurrentTicks - cy)
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
    cyOverheadTicks = cyOverheadTicks + (CurrentTicks - cy)

exit_proc:
    Exit Sub
    
on_error:
    MsgBox err.Description, vbOKOnly, "Error in " & ErrSrc(PROC)
End Sub

Private Sub TrcCollectInit()
    If Not cllTrace Is Nothing Then
        If cllTrace.Count <> 0 Then Exit Sub ' already initialized
    Else
        Set cllTrace = New Collection
    End If
    dtTraceBegin = Now()
    cyOverheadTicks = 0
    iTraceLevel = 0
    If lPrecision = 0 Then lPrecision = 6 ' default when another valus has not been set by the caller
    
End Sub

Public Sub TrcDsply()
    
    On Error GoTo on_error
    Const PROC = "TrcDsply"
    Dim dct         As Dictionary
    Dim v           As Variant
    Dim entry       As Collection
    Dim sTrace      As String
    
    TrcDsplyEndStack ' end procedures still on the stack

    If Not TrcEntriesAreConsistent(dct) Then
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
    Else
        sTrace = TrcDsplyHeader
        For Each v In cllTrace
            Set entry = v
            sTrace = sTrace & vbLf & TrcDsplyLine(entry)
        Next v
        sTrace = sTrace & vbLf & TrcDsplyFooter
        With fMsg
            .MaxFormWidthPrcntgOfScreenSize = 95
            .MsgTitle = "Execution Trace, displayed because the Conditional Compile Argument ""ExecTrace = 1""!"
            .MsgText(1) = sTrace:   .MsgMonoSpaced(1) = True
            .MsgLabel(2) = "About overhead, precision, etc.:": .MsgText(2) = TrcDsplyAbout
            .Setup
            .Show
        End With
    End If
    
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
      & Format(ExecSecs(0, cyOverheadTicks), ExecSecsFrmt) & "(" & cyOverheadTicksPcntg * 100 & "%)" _
      & " seconds. For a best possible execution time precision the average overhead of " _
      & Format((cyOverheadTicks / (cllTrace.Count / 2)) / cyFrequency, ExecSecsFrmt) & " secs " _
      & "had been deducted from each of the " & cllTrace.Count / 2 & " traced item's execution time." & vbLf _
      & "The displayed precision of the trace defaults to 0,000000 (6 decimals) which may " _
      & "be changed via the property ""TrcPrecision""." & vbLf _
      & "The displayed execution time will vary from execution to execution and can only be estimated as an average of many executions." & vbLf _
      & "When an error had been displayed the traced execution time includes the time " _
      & "of the user reaction and thus does not provide a meaningful result." & vbLf _

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

Private Function TrcDsplyFooter()
    TrcDsplyFooter = _
        Space$((Len(ExecSecsFrmt) * 2) + 2) _
      & PROC_END _
      & " End execution trace " _
      & Format(Now(), "hh:mm:ss")
End Function

Private Function TrcDsplyHeader() As String
    
    Const ELAPSED = "Elapsed"
    Dim sIndent As String: sIndent = Space$(Len(ExecSecsFrmt))

    TrcDsplyHeader = _
        ELAPSED _
      & VBA.Space$(Len(sIndent) - Len(ELAPSED) + 1) _
      & "Exec secs " _
      & PROC_BEGIN _
      & " Begin execution trace " _
      & Format(dtTraceBegin, "hh:mm:ss") _
      & " (exec time in seconds)"
End Function

Private Function TrcDsplyLine(ByVal entry As Collection) As String
' -------------------------------------------------------------
' Returns a trace line for being displayed.
' -------------------------------------------------------------
    TrcDsplyLine = _
        TrcEntryElpsdSecs(entry) _
      & " " _
      & TrcEntryExcSecs(entry) _
      & " " _
      & TrcDsplyLineIndentation(entry) _
      & TrcEntryDirective(entry) _
      & " " _
      & TrcEntryItem(entry) _
      & " " _
      & TrcEntryError(entry)
      
End Function

Private Function TrcEntriesAreConsistent(ByRef dct As Dictionary) As Boolean
' --------------------------------------------------------------------------
' Returns TRUE when for each begin entry there is a corresponding end entry,
' else the function returns FALSE and the collected items without a
' corresponding counterpart are returned as Dictionary (dct).
' For a best possible execution time precision the average overhead is
' deducted from each item's execution time.
' --------------------------------------------------------------------------

    Dim v                   As Variant
    Dim cllEndEntry         As Collection
    Dim cllBeginEntry       As Collection
    Dim bConsistent         As Boolean
    Dim cyBeginTicks        As Currency
    Dim cllEntry            As Collection
    Dim i                   As Long
    Dim j                   As Long
    Dim sComment            As String
    Dim cyAvrgOvrhdTicks    As Currency:    cyAvrgOvrhdTicks = cyOverheadTicks / (cllTrace.Count / 2)
    Dim cyTicksTotal        As Currency:    cyTicksTotal = TrcEntryTicks(cllTrace(cllTrace.Count)) - TrcEntryTicks(cllTrace(1))
    Dim cyTicksGross        As Currency
    Dim cyTicksNet          As Currency
    Dim cyTicks             As Currency
    
    If dct Is Nothing Then Set dct = New Dictionary
    
    cyOverheadTicksPcntg = cyOverheadTicks / cyTicksTotal
    
    '~~ Calculate the elapsed seconds for each entry
    cyBeginTicks = 0
    For Each v In cllTrace
        Set cllEntry = v
        If cyBeginTicks = 0 Then cyBeginTicks = TrcEntryTicks(cllEntry)
        cyTicks = TrcEntryTicks(cllEntry) - cyBeginTicks
        TrcEntryElpsdSecs(cllEntry) = Format((cyTicks / Frequency), ExecSecsFrmt)
    Next v
    Set cllEntry = Nothing
    
    '~~ Check for missing corresponding end entries while calculating the execution time for each end entry.
    For i = 1 To TrcEntryLastBegin
        If TrcEntryIsBegin(cllTrace(i), cllBeginEntry) Then
            bConsistent = False
            For j = i + 1 To cllTrace.Count
                If TrcEntryIsEnd(cllTrace(j), cllEndEntry) Then
                    If TrcEntryItem(cllBeginEntry) = TrcEntryItem(cllEndEntry) Then
                        If TrcEntryCallLvl(cllBeginEntry) = TrcEntryCallLvl(cllEndEntry) Then
                            '~~ Calculate the executesd seconds for the end entry
                            cyTicksGross = TrcEntryTicks(cllEndEntry) - TrcEntryTicks(cllBeginEntry)
                            cyTicksNet = cyTicksGross - (cyTicksGross * cyOverheadTicksPcntg)
                            TrcEntryExcSecs(cllEndEntry) = Format((cyTicksNet / cyFrequency), ExecSecsFrmt)
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
    
    '~~ Check for missing corresponding begin entries (if the end entry has no TrcEntryExcSecs)
    For Each v In cllTrace
        If TrcEntryIsEnd(v, cllEndEntry) Then
            If Trim(TrcEntryExcSecs(cllEndEntry)) = vbNullString Then
                '~~ No corresponding begin entry found
                Select Case TrcEntryDirective(cllEndEntry)
                    Case PROC_END: sComment = "No corresponding Begin of Procedure (BoP) code line in:  "
                    Case CODE_END: sComment = "No corresponding Begin of CodeTrace (BoC) code line in:  "
                End Select
                If Not dct.Exists(TrcEntryItem(cllBeginEntry)) Then dct.Add TrcEntryItem(cllBeginEntry), sComment
            End If
        End If
    Next v
    
    TrcEntriesAreConsistent = dct.Count = 0
    
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
    TrcEntryTicks(entry) = ticks
    TrcEntryDirective(entry) = dir
    TrcEntryItem(entry) = itm
    TrcEntryCallLvl(entry) = lvl
    TrcEntryError(entry) = err
    Set TrcEntry = entry
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

