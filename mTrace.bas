Attribute VB_Name = "mTrace"
Option Explicit
' --------------------------------------------------------------------------------
'
' Structure of a collected trace entry:
' -------------------------------------
' |  Entry Item | Origin, Transformation |   Type    |
' | ------------| -----------------------| --------- |
' | EntryNo     | = Count + 1,           | Long      |
' |             | before add Index       |           |
' | Ticks       | Collected              | Currency  |
' | ElapsedSecs | Computed               | Long      |
' | ExecSecs    | Computed               | Long      |
' | CallLevel   | = Number if items on   | Long      |
' |             | the stack -1 after push|           |
' |             | or item on stack before|           |
' |             | push                   |           |
' | Indentation | Computed based on      | String    |
' |             | CallLevel              |           |
' | DirectiveId | Collected              | String    |
' | Item        | Collected              | String    |
' | ExecError   | Collected              | String    |
' | TraceError  | Computed               | String    |

' --- Begin of declaration for the Execution Tracing
Private Declare PtrSafe Function getFrequency Lib "kernel32" _
Alias "QueryPerformanceFrequency" (cyFrequency As Currency) As Long
Private Declare PtrSafe Function getTickCount Lib "kernel32" _
Alias "QueryPerformanceCounter" (cyTickCount As Currency) As Long

Private Const BEGIN_ID      As String = ">"                   ' Begin procedure or code trace indicator
Private Const END_ID        As String = "<"                   ' End procedure or code trace indicator
Private Const COMMENT       As String = " !!! "

Private dicTrace            As Dictionary   ' For the collection of execution trance entries/lines
Private cllTrace            As Collection
Private cyFrequency         As Currency     ' Execution Trace Frequency (initialized with init)
Private iTraceItem          As Long         ' Execution Trace Call counter to unify key
Private lPrecisionDecimals  As Long         ' Execution Trace Default Precision (6=0,000000)
Private iSec                As Integer      ' Execution Trace digits left from decimal point
Private iDec                As Integer      ' Execution Trace decimal digits right from decimal point
Private sFormat             As String       ' Execution Trace tracking time presentation format
Private cyOverhead          As Currency     ' Execution Trace time accumulated by caused by the time tracking itself
Private dtTraceBeginTime    As Date         ' Execution Trace start time
Private iTraceLevel         As Long         ' Increased with each begin entry and decreased with each end entry

' --- End of declaration for the Execution Tracing

Private Property Get CallLevel(Optional ByRef entry As Collection) As Long:                 CallLevel = entry("5"):     End Property

Private Property Let CallLevel(Optional ByRef entry As Collection, ByRef l As Long):        entry.Add l, "5":           End Property

Private Property Get CODE_BEGIN() As String:                                 CODE_BEGIN = BEGIN_ID & " ":               End Property

Private Property Get CODE_END() As String:                                   CODE_END = END_ID & " ":                   End Property

Private Property Get CurrentTicks() As Currency:                                            getTickCount CurrentTicks:  End Property

Private Property Get Directive(Optional ByRef entry As Collection) As String:               Directive = entry("7"):     End Property

Private Property Let Directive(Optional ByRef entry As Collection, ByRef s As String):      entry.Add s, "7":           End Property

Private Property Get ElapsedSecs(Optional ByRef entry As Collection) As Long:               ElapsedSecs = entry("3"):   End Property

Private Property Let ElapsedSecs(Optional ByRef entry As Collection, ByRef l As Long):      entry.Add l, "3":           End Property

Private Property Get EntryNo(Optional ByRef entry As Collection) As Long:                   EntryNo = entry("1"):       End Property

Private Property Let EntryNo(Optional ByRef entry As Collection, ByRef l As Long):          entry.Add l, "1":           End Property

Private Property Get ExecError(Optional ByRef entry As Collection) As String:               ExecError = entry("9"):     End Property

Private Property Let ExecError(Optional ByRef entry As Collection, ByRef s As String):      entry.Add s, "9":           End Property

Private Property Get ExecSecs(Optional ByRef entry As Collection) As Long:                  ExecSecs = entry("4"):      End Property

Private Property Let ExecSecs(Optional ByRef entry As Collection, ByRef l As Long):         entry.Add l, "4":           End Property

Private Property Get INCOMPLETE() As String:  INCOMPLETE = COMMENT & "Incomplete trace" & COMMENT:                      End Property

Private Property Get Indentation(Optional ByRef entry As Collection) As String:             Indentation = entry("6"):   End Property

Private Property Let Indentation(Optional ByRef entry As Collection, ByRef s As String):    entry.Add s, "6":           End Property

Private Property Get Item(Optional ByRef entry As Collection) As String:                    Item = entry("8"):          End Property

Private Property Let Item(Optional ByRef entry As Collection, ByRef s As String):           entry.Add s, "8":           End Property

Private Property Get PROC_BEGIN() As String:                                 PROC_BEGIN = Repeat(BEGIN_ID, 2) & " ":    End Property

Private Property Get PROC_END() As String:                                   PROC_END = Repeat(END_ID, 2) & " ":        End Property

Private Property Get TickCount(Optional ByRef entry As Collection) As Currency:             TickCount = entry("2"):     End Property

Private Property Let TickCount(Optional ByRef entry As Collection, ByRef cy As Currency):   entry.Add cy, "2":          End Property

Private Property Get TraceError(Optional ByRef entry As Collection) As String:              TraceError = entry("10"):   End Property

Private Property Let TraceError(Optional ByRef entry As Collection, ByRef s As String):     entry.Add s, "10":          End Property

Private Sub DctTrcAdd(ByVal s As String, _
                   ByVal cy As Currency)
                   
    iTraceItem = iTraceItem + 1
    dicTrace.Add iTraceItem & s, cy
'    Debug.Print "Added to Trace: '" & iTraceItem & s
    
End Sub

Private Function ErrSrc(ByVal sProc As String) As String
    ErrSrc = "mTrace." & sProc
End Function

 Private Function IsBeginEntry( _
                  ByVal v As Collection, _
                  ByRef cll As Collection) As Boolean
' ---------------------------------------------------
' Returns TRUE and v as cll when the entry is a begin
' entry, else FALSE and cll = Nothing
' ---------------------------------------------------
    If InStr(Directive(v), BEGIN_ID) <> 0 Then
        IsBeginEntry = True
        Set cll = v
    End If
End Function

Private Function IsConsistent(ByRef cllReturn As Collection) As Boolean
' --------------------------------------------------------------------
' Returns TRUe when for each begin entry there is a corresponding end
' end entry, else the function returns FALSE and the items without a
' corresponding counterpart collected in cllReturn.
' --------------------------------------------------------------------

    Dim cll             As New Collection
    Dim v1              As Variant
    Dim v2              As Variant
    Dim cllEndEntry     As Collection
    Dim cllBeginEntry   As Collection
    Dim bConsistent     As Boolean
    Dim cyBeginTicks    As Currency
    Dim cllEntry        As Collection
    
    '~~ Calculate the elapsed seconds for each entry
    cyBeginTicks = 0
    For Each v1 In cllTrace
        Set cllEntry = v1
        If cyBeginTicks = 0 Then cyBeginTicks = TickCount(cllEntry)
        ElapsedSecs(cllEntry) = SecsExecuted(cyBeginTicks, TickCount(cllEntry))
    Next v1
    Set cllEntry = Nothing
    
    '~~ Check for corresponding End entries. A corresponding end entry when found will have an item ExecSecs
    For Each v1 In cllTrace
        If IsBeginEntry(v1, cllBeginEntry) Then
            bConsistent = False
            For Each v2 In cllTrace
                If IsEndEntry(v2, cllEndEntry) Then
                    If Item(cllBeginEntry) = Item(cllEndEntry) Then
                        If CallLevel(cllBeginEntry) = CallLevel(cllEndEntry) Then
                            '~~ Calculate the executesd seconds for the end entry
                            ExecSecs(cllEndEntry) = SecsExecuted(TickCount(cllBeginEntry), TickCount(cllEndEntry))
                            GoTo next_begin_entry
                        End If
                    End If
                End If
            Next v2
            '~~ No corresponding end enty found
            cll.Add Item(cllBeginEntry) & COMMENT & "No corresponding EoP/EoT code line!"
        End If

next_begin_entry:
    Next v1
    
    Set cllReturn = cll

End Function

Private Sub IsConsistent_Test()
    
    Dim endentry    As Collection
    Dim cll         As Collection
    Dim v           As Variant
    Dim entry       As Collection
    
    For Each v In cllTrace
        Set entry = v
        Debug.Print CallLevel(entry) & " " & Directive(entry) & " " & Item(entry)
    Next v
    
    If IsConsistent(cll) Then
        For Each v In cllTrace
            If IsEndEntry(v, endentry) Then
                Debug.Print "ElapsedSecs=" & ElapsedSecs(entry) & " " & "ExecSecs=" & ExecSecs(endentry) & " " & Item(endentry)
            End If
        Next v
    End If
    Set cllTrace = Nothing
    
End Sub

Private Function IsEndEntry( _
                 ByVal v As Collection, _
                 ByRef cll As Collection) As Boolean
' --------------------------------------------------
' Returns TRUE and v as cll when the entry is an end
' entry, else FALSE and cll = Nothing
' --------------------------------------------------
    If InStr(Directive(v), END_ID) <> 0 Then
        IsEndEntry = True
        Set cll = v
    End If
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

Private Function SecsExecuted( _
                 ByVal beginticks As Currency, _
                 ByVal endticks As Currency) As String
' ----------------------------------------------------
' Returns the difference between begin- and endticks
' as formatted seconds string (decimal = nanoseconds).
' ----------------------------------------------------
    Dim cy As Currency
    cy = (endticks - beginticks) / cyFrequency
    SecsExecuted = Format(cy, sFormat)
End Function

Private Sub TrcAdd( _
            ByVal itm As String, _
            ByVal ticks As Currency, _
            ByVal dir As String, _
            ByVal lvl As Long, _
   Optional ByVal err As String = vbNullString)
    cllTrace.Add TrcEntry(no:=cllTrace.Count + 1, ticks:=ticks, dir:=dir, itm:=itm, lvl:=lvl, err:=err)
End Sub

Private Sub TrcBegin(ByVal itm As String, _
                     ByVal dir As String)
' -----------------------------------------
' Keep a record of the current tick count
' at the begin of the execution of the
' procedure or code (item).
' -----------------------------------------
    Dim cy As Currency: cy = CurrentTicks
    
    DctTrcInit
    TrcInit
    
    DctTrcAdd dir & itm, cy
    
    iTraceLevel = iTraceLevel + 1
    TrcAdd itm:=itm, ticks:=cy, dir:=Trim(dir), lvl:=iTraceLevel

    cyOverhead = cyOverhead + (CurrentTicks - cy)
    
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
    TrcBeginLine = SecsExecuted(cyInitial, dicTrace.Items(iTt)) _
                 & "    " & sIndent _
                 & " " & Repeat("|  ", iIndent) _
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
    s = Replace(s, END_ID, BEGIN_ID)  ' turn the end item into a begin item string
    For j = i - 1 To 0 Step -1
        sKey = Split(TrcUnstripItemNo(dicTrace.Keys(j)), COMMENT)(0)
        sItem = Split(s, COMMENT)(0)
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
    TrcBegin s, CODE_BEGIN
End Sub

Public Sub TrcCodeEnd(ByVal s As String)
    TrcEnd s, CODE_END
End Sub

Private Function TrcDirectiveId(ByVal cll As Collection) As String
    TrcDirectiveId = cll(7)
End Function

Public Function TrcDsply() As String
' -------------------------------------------------
' Returns the execution trace a displayable string.
' -------------------------------------------------
    
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
                sTraceLine = Space$((Len(sFormat) * 2) + 1) & "    " & Repeat("|  ", iIndent) & sProcName
                If InStr(sTraceLine, PROC_END) <> 0 _
                Then sTraceLine = sTraceLine & INCOMPLETE & "The corresponding BoP statement for a EoP statement is missing." _
                Else sTraceLine = sTraceLine & INCOMPLETE & "The corresponding BoT statement for a EoT statement is missing."
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
    sTraceLine = Space$((Len(sFormat) * 2) + 2) & "<< End execution trace " & Format(Now(), "hh:mm:ss") & " (only " & Format(SecsExecuted(0, cyOverhead), "0.000000") & " seconds exec time were caused by the executuion trace itself)"
    sTrace = sTrace & vbLf & sTraceLine
    
exit_proc:
    TrcDsply = sTrace
    IsConsistent_Test
    
    Exit Function
    
on_error:
    MsgBox err.Description, vbOKOnly, "Error in " & ErrSrc(PROC)

End Function

Private Sub TrcEnd(ByVal s As String, _
          Optional ByVal dir As String = vbNullString, _
          Optional ByVal errinf As String = vbNullString)
' --------------------------------------------------------
' End of Trace. Keeps a record of the ticks count for the
' execution trace of the group of code lines named (s).
' --------------------------------------------------------
    
    On Error GoTo on_error
    Const PROC = "TrcEnd"
    Dim cy As Currency: cy = CurrentTicks
    
    If errinf <> vbNullString Then
        errinf = COMMENT & errinf & COMMENT
    End If
    DctTrcAdd dir & s & errinf, cy
    
    TrcAdd itm:=s, ticks:=cy, dir:=Trim(dir), lvl:=iTraceLevel, err:=errinf
    iTraceLevel = iTraceLevel - 1
    
    cyOverhead = cyOverhead + (CurrentTicks - cy)

exit_proc:
    Exit Sub
    
on_error:
    MsgBox err.Description, vbOKOnly, "Error in " & ErrSrc(PROC)
End Sub

Private Function TrcEndItemMissing(ByVal s As String) As String
' -------------------------------------------------------------------
' Returns a message string when a corresponding end item is missing.
' -------------------------------------------------------------------
    Dim i       As Long
    Dim sKey    As String
    Dim sInfo   As String

    TrcEndItemMissing = "missing"
    s = Replace(s, BEGIN_ID, END_ID)  ' turn the end item into a begin item string
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
        If Split(s, " ")(0) & " " = PROC_END _
        Then TrcEndItemMissing = INCOMPLETE & "The corresponding EoP statement for a BoP statement is missing." _
        Else TrcEndItemMissing = INCOMPLETE & "The corresponding EoT statement for a BoT statement is missing."
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
    
    TrcEndLine = SecsExecuted(cyInitial, cyEnd) & " " & _
                 SecsExecuted(cyStrt, cyEnd) & "    " & _
                 Repeat("|  ", iIndent) & _
                 sProcName

End Function

Private Sub TrcEndStack()
' --------------------------------------
' Completes the execution trace for
' items still on the stack.
' --------------------------------------
    Dim s As String
    Do Until mErrHndlr.StackIsEmpty
        s = mErrHndlr.StackPop
        If s <> vbNullString Then
            TrcEnd s
        End If
    Loop
End Sub

Private Function TrcEntry( _
      ByVal no As Long, _
      ByVal ticks As Currency, _
      ByVal dir As String, _
      ByVal itm As String, _
      ByVal lvl As Long, _
      ByVal err As String) As Collection
      
    Dim entry As New Collection
    
    EntryNo(entry) = no
    TickCount(entry) = ticks
    Directive(entry) = dir
    Item(entry) = itm
    CallLevel(entry) = lvl
    TraceError(entry) = err

    Set TrcEntry = entry
End Function

Public Sub TrcError(ByVal s As String)
' --------------------------------------
' Keep record of the error (s) raised
' during the execution of any procedure.
' --------------------------------------
    Dim cy As Currency: cy = CurrentTicks
    
    DctTrcInit
    TrcInit
    
    DctTrcAdd PROC_END & s, cy
    cyOverhead = cyOverhead + (CurrentTicks - cy)
    
End Sub

Private Sub DctTrcInit()
    If Not dicTrace Is Nothing Then
        If dicTrace.Count <> 0 Then Exit Sub
    Else
        Set dicTrace = New Dictionary
    End If
    dtTraceBeginTime = Now()
    iTraceItem = 0
    cyOverhead = 0
    iTraceLevel = 0

End Sub

Private Sub TrcInit()
    If Not cllTrace Is Nothing Then
        If cllTrace.Count <> 0 Then Exit Sub
    Else
        Set cllTrace = New Collection
    End If

End Sub

Private Function TrcIsBegItem(ByRef s As String) As Boolean
' ---------------------------------------------------------
' Returns TRUE if s is an execution trace begin item.
' Returns s with the call counter unstripped.
' ---------------------------------------------------------
Dim i As Single
    TrcIsBegItem = False
    i = InStr(1, s, BEGIN_ID)
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
        Case Trim(PROC_END), Trim(CODE_END)
            TrcIsEndItem = True
        Case Else
            TrcIsEndItem = False
    End Select
        
    If InStr(s, COMMENT) <> 0 Then
        sRight = COMMENT & Split(s, COMMENT)(1)
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
    TrcBegin s, PROC_BEGIN
End Sub

Public Sub TrcProcEnd(ByVal s As String, _
             Optional ByVal errinfo As String = vbNullString)
    TrcEnd s, PROC_END, errinfo
End Sub

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

