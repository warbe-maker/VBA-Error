Attribute VB_Name = "mTrc"
Option Explicit
' ------------------------------------------------------------------------
' Standard Module mTrc Procedure/code execution trace with result display.
'
' Uses: mErrHndlr
'       fMsg
'
' W. Rauschenberger, Berlin, Nov. 1 2020
' Data structure of any collected begin/end trace entry:

' | Entry Item          | Origin, Transformation |   Type    | Key |
' |---------------------|------------------------|-----------+-----+
' | NtryTcksSys         | Collected              | Currency  | TS  |
' | NtryTcksElpsd       | Computed               | Currency  | TE  |
' | NtryTcksGrss        | Computed               | Currency  | TG  |
' | NtryTcksOvrhdNtry   | Collected              | Currency  | TON |
' | NtryTcksOvrhdItm    | Computed               | Currency  | TOI |
' | NtryScsElpsd        | Computed               | Currency  | SE  |
' | NtryScsGrss         | Computed               | Currency  | SG  |
' | NtryScsOvrhdNtry    | Computed               | Currency  | SON |
' | NtryScsOvrhdItm     | Computed               | Currency  | SOI |
' | NtryScsNt           | Computed               | Currency  | SN  |
' |                     | gross - overhad        |           |     |
' | NtryCllLvl          | Computed               | Long      | CL  |
' | NtryDrctv           | Collected              | String    | D   |
' | NtryItm            | Collected              | String    | I   |
' | NtryError           | Computed               | String    | E   |

Public Enum enDisplayedInfo
    Detailed = 1
    Compact = 2
End Enum

Private Declare PtrSafe Function getFrequency Lib "kernel32" _
Alias "QueryPerformanceFrequency" (cySysFrequency As Currency) As Long
Private Declare PtrSafe Function getTickCount Lib "kernel32" _
Alias "QueryPerformanceCounter" (cyTickCount As Currency) As Long

Private Const DIR_BEGIN_ID  As String = ">"     ' Begin procedure or code trace indicator
Private Const DIR_END_ID    As String = "<"     ' End procedure or code trace indicator
Private Const COMMENT       As String = " !!! "

Private cllStck             As Collection       ' Trace stack
Private cllNtryLast         As Collection       '
Private cllTrc              As Collection       ' Collection of begin and end trace entries
Private cySysFrequency      As Currency         ' Execution Trace SysFrequency (initialized with init)
Private cyTcksOvrhd         As Currency         ' Overhead ticks caused by the collection of a traced item's entry
Private cyTcksOvrhdItm      As Currency         ' Execution Trace time accumulated by caused by the time tracking itself
Private dtTraceBegin        As Date             ' Initialized at start of execution trace
Private iTrcLvl             As Long             ' Increased with each begin entry and decreased with each end entry
Private lDisplayedInfo      As Long             ' Detailed or Compact, defaults to Compact
Private lPrecision          As Long             ' Precision for displayed seconds, defaults to 6 decimals (0,000000)
Private sFrmtScsElpsd       As String           ' -------------
Private sFrmtScsGrss        As String           ' Format String
Private sFrmtScsNt          As String           ' for Seconds
Private sFrmtScsOvrhdNtry   As String           ' -------------
Private sFrmtScsOvrhdItm    As String           ' -------------
Private sFrmtTcksElpsd      As String           ' -------------
Private sFrmtTcksGrss       As String           ' Format String
Private sFrmtTcksNt         As String           ' for
Private sFrmtTcksOvrhdNtry  As String           ' Ticks
Private sFrmtTcksOvrhdItm   As String           ' Ticks
Private sFrmtTcksSys        As String           ' -------------
Private sFirstTraceItem     As String

Private Property Get DIR_BEGIN_CODE() As String
    DIR_BEGIN_CODE = DIR_BEGIN_ID
End Property

Private Property Get DIR_BEGIN_PROC() As String
    DIR_BEGIN_PROC = Repeat(DIR_BEGIN_ID, 2)
End Property

Private Property Get DIR_END_CODE() As String
    DIR_END_CODE = DIR_END_ID
End Property

Private Property Get DIR_END_PROC() As String
    DIR_END_PROC = Repeat(DIR_END_ID, 2)
End Property

Private Property Get DisplayedInfo() As enDisplayedInfo
    DisplayedInfo = lDisplayedInfo
End Property

Public Property Let DisplayedInfo(ByVal l As enDisplayedInfo)
    lDisplayedInfo = l
End Property

Private Property Get DisplayedSecsPrecision() As Long
    DisplayedSecsPrecision = lPrecision
End Property

Public Property Let DisplayedSecsPrecision(ByVal l As Long)
    lPrecision = l
End Property

Private Property Get DsplyLnIndnttn(Optional ByRef entry As Collection) As String
    DsplyLnIndnttn = Repeat("|  ", NtryCllLvl(entry))
End Property

Private Property Get NtryCllLvl(Optional ByRef entry As Collection) As Long
    NtryCllLvl = entry("CL")
End Property

Private Property Let NtryCllLvl(Optional ByRef entry As Collection, ByRef l As Long)
    entry.Add l, "CL"
End Property

Private Property Get NtryDrctv(Optional ByRef entry As Collection) As String
    NtryDrctv = entry("D")
End Property

Private Property Let NtryDrctv(Optional ByRef entry As Collection, ByRef s As String)
    entry.Add s, "D"
End Property

Private Property Get NtryError(Optional ByRef entry As Collection) As String
    On Error Resume Next ' in case this has never been collected
    NtryError = entry("E")
    If err.Number <> 0 Then NtryError = vbNullString
End Property

Private Property Let NtryError(Optional ByRef entry As Collection, ByRef s As String)
    entry.Add s, "E"
End Property

Private Property Get NtryItm(Optional ByRef entry As Collection) As String
    NtryItm = entry("I")
End Property

Private Property Let NtryItm(Optional ByRef entry As Collection, ByRef s As String)
    entry.Add s, "I"
End Property

Private Property Get NtryScsElpsd(Optional ByRef entry As Collection) As Currency
    On Error Resume Next
    NtryScsElpsd = entry("SE")
    If err.Number <> 0 Then NtryScsElpsd = Space$(Len(sFrmtScsElpsd))
End Property

Private Property Let NtryScsElpsd(Optional ByRef entry As Collection, ByRef cy As Currency)
    entry.Add cy, "SE"
End Property

Private Property Get NtryScsGrss(Optional ByRef entry As Collection) As Currency
    On Error Resume Next ' in case no value exists (the case for each begin entry)
    NtryScsGrss = entry("SG")
    If err.Number <> 0 Then NtryScsGrss = Space$(Len(sFrmtScsGrss))
End Property

Private Property Let NtryScsGrss(Optional ByRef entry As Collection, ByRef cy As Currency)
    entry.Add cy, "SG"
End Property

Private Property Get NtryScsNt(Optional ByRef entry As Collection) As Double
    On Error Resume Next
    NtryScsNt = entry("SN")
    If err.Number <> 0 Then NtryScsNt = Space$(Len(sFrmtScsNt))
End Property

Private Property Let NtryScsNt(Optional ByRef entry As Collection, ByRef dbl As Double)
    entry.Add dbl, "SN"
End Property

Private Property Get NtryScsOvrhdItm(Optional ByRef entry As Collection) As Double
    On Error Resume Next
    NtryScsOvrhdItm = entry("SOI")
    If err.Number <> 0 Then NtryScsOvrhdItm = Space$(Len(sFrmtScsOvrhdItm))
End Property

Private Property Let NtryScsOvrhdItm(Optional ByRef entry As Collection, ByRef dbl As Double)
    entry.Add dbl, "SOI"
End Property

Private Property Get NtryScsOvrhdNtry(Optional ByRef entry As Collection) As Double
    On Error Resume Next
    NtryScsOvrhdNtry = entry("SON")
    If err.Number <> 0 Then NtryScsOvrhdNtry = Space$(Len(sFrmtScsOvrhdItm))
End Property

Private Property Let NtryScsOvrhdNtry(Optional ByRef entry As Collection, ByRef dbl As Double)
    entry.Add dbl, "SON"
End Property

Private Property Get NtryTcksElpsd(Optional ByRef entry As Collection) As Currency
    NtryTcksElpsd = entry("TE")
End Property

Private Property Let NtryTcksElpsd(Optional ByRef entry As Collection, ByRef cy As Currency)
    entry.Add cy, "TE"
End Property

Private Property Get NtryTcksGrss(Optional ByRef entry As Collection) As Currency
    On Error Resume Next
    NtryTcksGrss = entry("TG")
    If err.Number <> 0 Then NtryTcksGrss = 0
End Property

Private Property Let NtryTcksGrss(Optional ByRef entry As Collection, ByRef cy As Currency)
    entry.Add cy, "TG"
End Property

Private Property Get NtryTcksNt(Optional ByRef entry As Collection) As Currency
    On Error Resume Next
    NtryTcksNt = entry("TN")
    If err.Number <> 0 Then NtryTcksNt = 0
End Property

Private Property Let NtryTcksNt(Optional ByRef entry As Collection, ByRef cy As Currency)
    entry.Add cy, "TN"
End Property

Private Property Get NtryTcksOvrhdItm(Optional ByRef entry As Collection) As Currency
    On Error Resume Next
    NtryTcksOvrhdItm = entry("TOI")
    If err.Number <> 0 Then NtryTcksOvrhdItm = 0
End Property

Private Property Let NtryTcksOvrhdItm(Optional ByRef entry As Collection, ByRef cy As Currency)
    entry.Add cy, "TOI"
End Property

Private Property Get NtryTcksOvrhdNtry(Optional ByRef entry As Collection) As Currency
    On Error Resume Next
    NtryTcksOvrhdNtry = entry("TON")
    If err.Number <> 0 Then NtryTcksOvrhdNtry = 0
End Property

Private Property Let NtryTcksOvrhdNtry(Optional ByRef entry As Collection, ByRef cy As Currency)
    entry.Add cy, "TON"
End Property

Private Property Get NtryTcksSys(Optional ByRef entry As Collection) As Currency
    NtryTcksSys = entry("TS")
End Property

Private Property Let NtryTcksSys(Optional ByRef entry As Collection, ByRef cy As Currency)
    entry.Add cy, "TS"
End Property

Private Property Get SysCrrntTcks() As Currency
    getTickCount SysCrrntTcks
End Property

Private Property Get SysFrequency() As Currency
    If cySysFrequency = 0 Then
        getFrequency cySysFrequency
    End If
    SysFrequency = cySysFrequency
End Property

Public Sub BoC(ByVal s As String)
' -------------------------------
' Begin of Code trace.
' -------------------------------
#If ExecTrace Then
    Dim cll As Collection
    
    TrcBgn itm:=s, dir:=DIR_BEGIN_CODE, cll:=cll
    StckPush cll
#End If
End Sub

Public Sub BoP(ByVal s As String)
' -------------------------------
' Trace Begin of Procedure
' -------------------------------
#If ExecTrace Then
    Dim cll As Collection
    
    If TrcIsEmpty Then
        Initialize
        sFirstTraceItem = s
    Else
        If s = sFirstTraceItem Then
            '~~ A previous trace had not come to a regular end and thus will be erased
            Set cllTrc = Nothing
            Initialize
        End If
    End If
    TrcBgn itm:=s, dir:=DIR_BEGIN_PROC, cll:=cll
    StckPush cll

#End If
End Sub

Private Sub ComputeSecsGrssOvrhdNet()
' -----------------------------------
' Compite the seconds based on ticks.
' -----------------------------------
    Const PROC = "ComputeSecsGrssOvrhdNet"
    
    On Error GoTo eh
    Dim v   As Variant
    Dim cll As Collection
    
    For Each v In cllTrc
        Set cll = v
        NtryScsOvrhdNtry(cll) = CDbl(NtryTcksOvrhdNtry(cll)) / CDbl(SysFrequency)
        NtryScsElpsd(cll) = CDec(NtryTcksElpsd(cll)) / CDec(SysFrequency)
        If Not NtryIsBegin(cll) Then
            NtryScsGrss(cll) = CDec(NtryTcksGrss(cll)) / CDec(SysFrequency)
            NtryScsOvrhdItm(cll) = CDec(NtryTcksOvrhdItm(cll)) / CDec(SysFrequency)
            NtryScsNt(cll) = CDec(NtryTcksNt(cll)) / CDec(SysFrequency)
        End If
    Next v
    Set cll = Nothing

xt: Exit Sub

eh: ErrMsg errno:=err.Number, errsource:=ErrSrc(PROC), errdscrptn:=err.Description, errline:=Erl
#If Debugging Then
    Stop: Resume
#End If
    Set cllTrc = Nothing
End Sub

Private Function ComputeSecsOvrhdTtlEntry() As Double
' --------------------------------------------------
' Returns the total overhead seconds caused by the
' collection of the traced item's data.
' ----------------------------------------------
    Const PROC = "ComputeSecsOvrhdTtlEntry"
    
    On Error GoTo eh
    Dim v   As Variant
    Dim cll As Collection
    Dim dbl As Double
    
    For Each v In cllTrc
        Set cll = v
        dbl = dbl + NtryScsOvrhdNtry(cll)
    Next v
    ComputeSecsOvrhdTtlEntry = dbl

xt: Exit Function

eh: ErrMsg errno:=err.Number, errsource:=ErrSrc(PROC), errdscrptn:=err.Description, errline:=Erl
#If Debugging Then
    Stop: Resume
#End If
    Set cllTrc = Nothing
End Function

Private Sub ComputeTcksElpsd()
' ------------------------------
' Compute the elapsed ticks.
' ------------------------------
    Const PROC = "ComputeTcksElpsd"
    
    On Error GoTo eh
    Dim v               As Variant
    Dim cll             As Collection
    Dim cyTcksBegin     As Currency
    Dim cyTcksElapsed   As Currency
    
    For Each v In cllTrc
        Set cll = v
        If cyTcksBegin = 0 Then cyTcksBegin = NtryTcksSys(cll)
        cyTcksElapsed = NtryTcksSys(cll) - cyTcksBegin
        NtryTcksElpsd(cll) = cyTcksElapsed
    Next v
    Set cll = Nothing

xt: Exit Sub

eh: ErrMsg errno:=err.Number, errsource:=ErrSrc(PROC), errdscrptn:=err.Description, errline:=Erl
#If Debugging Then
    Stop: Resume
#End If
    Set cllTrc = Nothing
End Sub

Private Sub ComputeTcksNet()
' ----------------------------
' Compute the net ticks by
' deducting the total overhad
' ticks from the gross ticks.
' ----------------------------
    Const PROC = "ComputeTcksNet"
    
    On Error GoTo eh
    Dim v   As Variant
    Dim cll As Collection
    
    For Each v In cllTrc
        Set cll = v
        NtryTcksNt(cll) = NtryTcksGrss(cll) - NtryTcksOvrhdNtry(cll)
    Next v
    Set cll = Nothing

xt: Exit Sub

eh: ErrMsg errno:=err.Number, errsource:=ErrSrc(PROC), errdscrptn:=err.Description, errline:=Erl
#If Debugging Then
    Stop: Resume
#End If
    Set cllTrc = Nothing
End Sub

Private Function ComputeTcksOvrhdItem() As Currency
' -------------------------------------------------
' Compute the total overhead ticks caused by the
' collection of the traced item's data.
' -------------------------------------------------
    Const PROC = "ComputeTcksOvrhdItem"
    
    On Error GoTo eh
    Dim v   As Variant
    Dim cll As Collection
    Dim cy  As Currency
    
    For Each v In cllTrc
        Set cll = v
        cy = cy + NtryTcksOvrhdNtry(cll)
    Next v
    ComputeTcksOvrhdItem = cy

xt: Exit Function

eh: ErrMsg errno:=err.Number, errsource:=ErrSrc(PROC), errdscrptn:=err.Description, errline:=Erl
#If Debugging Then
    Stop: Resume
#End If
    Set cllTrc = Nothing
End Function

Public Sub Dsply()
    Const PROC = "Dsply"
    
    On Error GoTo eh
    Dim dct         As Dictionary
    Dim v           As Variant
    Dim entry       As Collection
    Dim sTrace      As String
    Dim lLenHeader  As Long
    
    NtryTcksOvrhdNtry(cllNtryLast) = cyTcksOvrhd

    If DisplayedInfo = 0 Then DisplayedInfo = Compact

    If Not DsplyNtryAllCnsstnt(dct) Then
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
        Set cllTrc = Nothing
        GoTo xt
    End If
        
    ComputeTcksNet
    ComputeSecsGrssOvrhdNet        ' Calculate gross and net execution seconds
    DsplyValuesFormatSet       ' Setup the format for the displayed ticks and sec values
    
    sTrace = DsplyHdr(lLenHeader)
    For Each v In cllTrc
        Set entry = v
        sTrace = sTrace & vbLf & DsplyLn(entry)
    Next v
    sTrace = sTrace & vbLf & DsplyFtr(lLenHeader)
    With fMsg
        .MaxFormWidthPrcntgOfScreenSize = 95
        .MsgTitle = "Execution Trace, displayed because the Conditional Compile Argument ""ExecTrace = 1""!"
        .MsgText(1) = sTrace:   .MsgMonoSpaced(1) = True
        .MsgLabel(2) = "About overhead, precision, etc.:": .MsgText(2) = DsplyAbout
        .Setup
        .Show
    End With
    
xt: mTrc.Terminate
    Exit Sub
    
eh: ErrMsg errno:=err.Number, errsource:=ErrSrc(PROC), errdscrptn:=err.Description, errline:=Erl
#If Debugging Then
    Stop: Resume
#End If
    Set cllTrc = Nothing
End Sub

Private Function DsplyAbout() As String
    
    Dim cyTcksOvrhdItm      As Currency
    Dim dblOvrhdPcntg       As Double
    Dim dblTtlScsOvrhdNtry  As Double
    
    dblTtlScsOvrhdNtry = ComputeSecsOvrhdTtlEntry
    cyTcksOvrhdItm = ComputeTcksOvrhdItem
    dblOvrhdPcntg = (dblTtlScsOvrhdNtry / NtryScsElpsd(NtryLst)) * 100
    
    DsplyAbout = "> The trace itself, i.e. the collection of the begin and end data for each traced item " & _
                 "(procedure or code) caused a performance loss of " & Format(dblTtlScsOvrhdNtry, sFrmtScsOvrhdItm) & _
                 " seconds (=" & Format(dblOvrhdPcntg, "0.00") & "%). " _
               & "For a best possible execution time precision the overhead per traced item " _
               & "has been deducted from each of the " & cllTrc.Count / 2 & " traced item's execution time." _
      & vbLf _
      & "> The precision (decimals) for the displayed seconds defaults to 0,000000 (6 decimals) which may " _
      & "be changed via the property ""DisplayedSecsPrecision""." _
      & vbLf _
      & "> The displayed execution time varies from execution to execution and can only be estimated " _
      & "as an average of many executions." _
      & vbLf _
      & "> When an error had been displayed the traced execution time includes the time " _
      & "of the user reaction and thus does not provide a meaningful result."

End Function

Private Function DsplyFtr(ByVal lLenHeaderData As Long) ' Displayed trace footer
    DsplyFtr = _
        Space$(lLenHeaderData) _
      & DIR_END_PROC _
      & " End execution trace " _
      & Format(Now(), "hh:mm:ss")
End Function

Private Function DsplyHdr( _
                 ByRef lLenHeaderData As Long) As String
' ------------------------------------------------------
' Compact:  Header1 = Elapsed Exec >> Start
' Detailed: Header1 =    Ticks   |   Secs
'           Header2 = xxx xxxx xx xxx xxx xx
'           Header3 " --- ---- -- --- --- -- >> Start
' ------------------------------------------------------
    Const PROC = "DsplyHdr"
    
    On Error GoTo eh
    Dim sIndent         As String: sIndent = Space$(Len(sFrmtScsElpsd))
    Dim sHeader1        As String
    Dim sHeader2        As String
    Dim sHeader2Ticks   As String
    Dim sHeader2Secs    As String
    Dim sHeader3        As String
    Dim sHeaderTrace    As String
    
    sHeaderTrace = _
      DIR_BEGIN_PROC _
    & " Begin execution trace " _
    & Format(dtTraceBegin, "hh:mm:ss") _
    & " (exec time in seconds)"

    Select Case DisplayedInfo
        Case Compact
            sHeader2Ticks = vbNullString
            sHeader2Secs = _
                    DsplyHdrCntrAbv("Elapsed", sFrmtScsElpsd) _
            & " " & DsplyHdrCntrAbv("Net", sFrmtScsNt)
            
            sHeader2 = sHeader2Secs & " "
            lLenHeaderData = Len(sHeader2)
            
            sHeader3 = _
                    Repeat$("-", Len(sFrmtScsElpsd)) _
            & " " & Repeat$("-", Len(sFrmtScsNt)) _
            & " " & sHeaderTrace
              
            sHeader1 = DsplyHdrCntrAbv(" Seconds ", sHeader2Secs, , , "-")
            
            DsplyHdr = _
              sHeader1 & vbLf _
            & sHeader2 & vbLf _
            & sHeader3
        
        Case Detailed
            sHeader2Ticks = _
                    DsplyHdrCntrAbv("System", sFrmtTcksSys) _
            & " " & DsplyHdrCntrAbv("Elapsed", sFrmtTcksElpsd) _
            & " " & DsplyHdrCntrAbv("Gross", sFrmtTcksGrss) _
            & " " & DsplyHdrCntrAbv("Oh Entry", sFrmtTcksOvrhdItm) _
            & " " & DsplyHdrCntrAbv("Oh Item", sFrmtTcksOvrhdItm) _
            & " " & DsplyHdrCntrAbv("Net", sFrmtTcksNt)
            
            sHeader2Secs = _
                    DsplyHdrCntrAbv("Elapsed", sFrmtScsElpsd) _
            & " " & DsplyHdrCntrAbv("Gross", sFrmtScsGrss) _
            & " " & DsplyHdrCntrAbv("Oh Entry", sFrmtScsOvrhdItm) _
            & " " & DsplyHdrCntrAbv("Oh Item", sFrmtScsOvrhdItm) _
            & " " & DsplyHdrCntrAbv("Net", sFrmtScsNt)
            
            sHeader2 = sHeader2Ticks & " " & sHeader2Secs & " "
            lLenHeaderData = Len(sHeader2)
            
            sHeader3 = _
                    Repeat$("-", Len(sFrmtTcksSys)) _
            & " " & Repeat$("-", Len(sFrmtTcksElpsd)) _
            & " " & Repeat$("-", Len(sFrmtTcksGrss)) _
            & " " & Repeat$("-", Len(sFrmtTcksOvrhdItm)) _
            & " " & Repeat$("-", Len(sFrmtTcksOvrhdItm)) _
            & " " & Repeat$("-", Len(sFrmtTcksNt)) _
            & " " & Repeat$("-", Len(sFrmtScsElpsd)) _
            & " " & Repeat$("-", Len(sFrmtScsGrss)) _
            & " " & Repeat$("-", Len(sFrmtScsOvrhdItm)) _
            & " " & Repeat$("-", Len(sFrmtScsOvrhdItm)) _
            & " " & Repeat$("-", Len(sFrmtScsNt)) _
            & " " & sHeaderTrace
              
            sHeader1 = _
                    DsplyHdrCntrAbv(" Ticks ", sHeader2Ticks, , , "-") _
            & " " & DsplyHdrCntrAbv(" Seconds ", sHeader2Secs, , , "-")
            
            DsplyHdr = _
                sHeader1 & vbLf _
              & sHeader2 & vbLf _
              & sHeader3
            
    End Select

xt: Exit Function

eh: ErrMsg errno:=err.Number, errsource:=ErrSrc(PROC), errdscrptn:=err.Description, errline:=Erl
#If Debugging Then
    Stop: Resume
#End If
    Set cllTrc = Nothing
End Function

Public Function DsplyHdrCntrAbv( _
                          ByVal s1 As String, _
                          ByVal s2 As String, _
                 Optional ByVal sLeft As String = vbNullString, _
                 Optional ByVal sRight As String = vbNullString, _
                 Optional ByVal sFillChar As String = " ") As String
' ---------------------------------------------------------------------------
' Returns s1 centered above s2 considering any characters left and right.
' ---------------------------------------------------------------------------
    
    If Len(s1) > Len(s2) Then
        DsplyHdrCntrAbv = s1
    Else
        DsplyHdrCntrAbv = s1 & String$(Int((Len(s2) - Len(s1 & sLeft & sRight)) / 2), sFillChar)
        DsplyHdrCntrAbv = String$(Len(s2) + -Len(sLeft & DsplyHdrCntrAbv & sRight), sFillChar) & DsplyHdrCntrAbv
        DsplyHdrCntrAbv = sLeft & DsplyHdrCntrAbv & sRight
    End If
    
End Function

Private Function DsplyLn(ByVal entry As Collection) As String
' -------------------------------------------------------------
' Returns a trace line for being displayed.
' -------------------------------------------------------------
    Const PROC = "DsplyLn"
    On Error GoTo eh
    
    Select Case DisplayedInfo
        Case Compact
            DsplyLn = _
                      DsplyValue(entry, NtryScsElpsd(entry), sFrmtScsElpsd) _
              & " " & DsplyValue(entry, NtryScsNt(entry), sFrmtScsNt) _
              & " " & DsplyLnIndnttn(entry) _
                    & NtryDrctv(entry) _
              & " " & NtryItm(entry) _
              & " " & NtryError(entry)
        Case Detailed
            DsplyLn = _
                        DsplyValue(entry, NtryTcksSys(entry), sFrmtTcksSys) _
                & " " & DsplyValue(entry, NtryTcksElpsd(entry), sFrmtTcksElpsd) _
                & " " & DsplyValue(entry, NtryTcksGrss(entry), sFrmtTcksGrss) _
                & " " & DsplyValue(entry, NtryTcksOvrhdNtry(entry), sFrmtTcksOvrhdItm) _
                & " " & DsplyValue(entry, NtryTcksOvrhdItm(entry), sFrmtTcksOvrhdItm) _
                & " " & DsplyValue(entry, NtryTcksNt(entry), sFrmtTcksNt) _
                & " " & DsplyValue(entry, NtryScsElpsd(entry), sFrmtScsElpsd) _
                & " " & DsplyValue(entry, NtryScsGrss(entry), sFrmtScsGrss) _
                & " " & DsplyValue(entry, NtryScsOvrhdNtry(entry), sFrmtScsOvrhdItm) _
                & " " & DsplyValue(entry, NtryScsOvrhdItm(entry), sFrmtScsOvrhdItm) _
                & " " & DsplyValue(entry, NtryScsNt(entry), sFrmtScsNt) _
                & " " & DsplyLnIndnttn(entry) _
                      & NtryDrctv(entry) _
                & " " & NtryItm(entry) _
                & " " & NtryError(entry)
    End Select

xt: Exit Function

eh: ErrMsg errno:=err.Number, errsource:=ErrSrc(PROC), errdscrptn:=err.Description, errline:=Erl
#If Debugging Then
    Stop: Resume
#End If
    Set cllTrc = Nothing
End Function

Private Function DsplyNtryAllCnsstnt(ByRef dct As Dictionary) As Boolean
' ----------------------------------------------------------------------
' Returns TRUE when for each begin entry there is a corresponding end
' entry and vice versa. Else the function returns FALSE and the items
' without a corresponding counterpart are returned as Dictionary (dct).
' The consistency check is based on the calculated execution ticks as
' the difference between the elapsed begin and end ticks.
' ----------------------------------------------------------------------
    Const PROC = "DsplyLn"
    
    On Error GoTo eh
    Dim v                   As Variant
    Dim cllEndEntry         As Collection
    Dim cllBeginEntry       As Collection
    Dim bConsistent         As Boolean
    Dim i                   As Long
    Dim j                   As Long
    Dim sComment            As String
    
    If dct Is Nothing Then Set dct = New Dictionary
        
    ComputeTcksElpsd ' Calculates the ticks elapsed since trace start
    
    '~~ Check for missing corresponding end entries while calculating the execution time for each end entry.
    For i = 1 To NtryLastBegin
        If NtryIsBegin(cllTrc(i), cllBeginEntry) Then
            Debug.Print NtryCllLvl(cllBeginEntry) & " " & NtryDrctv(cllBeginEntry) & " " & NtryItm(cllBeginEntry)
            bConsistent = False
            For j = i + 1 To cllTrc.Count
                If NtryIsEnd(cllTrc(j), cllEndEntry) Then
                    If NtryItm(cllBeginEntry) = NtryItm(cllEndEntry) Then
                        If NtryCllLvl(cllBeginEntry) = NtryCllLvl(cllEndEntry) Then
                            '~~ Calculate the executesd seconds for the end entry
                            NtryTcksGrss(cllEndEntry) = NtryTcksElpsd(cllEndEntry) - NtryTcksElpsd(cllBeginEntry)
                            NtryTcksOvrhdItm(cllEndEntry) = NtryTcksOvrhdNtry(cllBeginEntry) + NtryTcksOvrhdNtry(cllEndEntry)
                            GoTo next_begin_entry
                        End If
                    End If
                End If
            Next j
            '~~ No corresponding end entry found
            Select Case NtryDrctv(cllBeginEntry)
                Case DIR_BEGIN_PROC: sComment = "No corresponding End of Procedure (EoP) code line in:    "
                Case DIR_BEGIN_CODE: sComment = "No corresponding End of CodeTrace (EoC) code line in:    "
            End Select
            If Not dct.Exists(NtryItm(cllBeginEntry)) Then dct.Add NtryItm(cllBeginEntry), sComment
        End If

next_begin_entry:
    Next i
    
    '~~ Check for missing corresponding begin entries (if the end entry has no executed ticks)
    For Each v In cllTrc
        If NtryIsEnd(v, cllEndEntry) Then
            If NtryTcksGrss(cllEndEntry) = 0 Then
                '~~ No corresponding begin entry found
                Select Case NtryDrctv(cllEndEntry)
                    Case DIR_END_PROC: sComment = "No corresponding Begin of Procedure (BoP) code line in:  "
                    Case DIR_END_CODE: sComment = "No corresponding Begin of CodeTrace (BoC) code line in:  "
                End Select
                If Not dct.Exists(NtryItm(cllBeginEntry)) Then dct.Add NtryItm(cllBeginEntry), sComment
            End If
        End If
    Next v
    
    DsplyNtryAllCnsstnt = dct.Count = 0
    Exit Function

eh: ErrMsg errno:=err.Number, errsource:=ErrSrc(PROC), errdscrptn:=err.Description, errline:=Erl
#If Debugging Then
    Stop: Resume
#End If
End Function

Private Function DsplyTcksDffToScs( _
                             ByVal beginticks As Currency, _
                             ByVal endticks As Currency) As Currency
' ------------------------------------------------------------------
' Returns the difference between begin- and endticks as seconds.
' ------------------------------------------------------------------
    IIf endticks - beginticks > 0, DsplyTcksDffToScs = (endticks - beginticks) / cySysFrequency, DsplyTcksDffToScs = 0
End Function

Private Function DsplyValue(ByVal entry As Collection, _
                            ByVal value As Variant, _
                            ByVal frmt As String) As String
    If NtryIsBegin(entry) And value = 0 _
    Then DsplyValue = Space$(Len(frmt)) _
    Else DsplyValue = IIf(value >= 0, Format(value, frmt), Space$(Len(frmt)))

End Function

Private Function DsplyValueFormat( _
                 ByRef thisformat As String, _
                 ByVal forvalue As Variant) As String
    thisformat = String$(DsplyValueLength(forvalue), "0") & "." & String$(DisplayedSecsPrecision, "0")
End Function

Public Function DsplyValueLength(ByVal v As Variant) As Long

    If InStr(CStr(v), ".") <> 0 Then
        DsplyValueLength = Len(Split(CStr(v), ".")(0))
    ElseIf InStr(CStr(v), ",") <> 0 Then
        DsplyValueLength = Len(Split(CStr(v), ",")(0))
    Else
        DsplyValueLength = Len(v)
    End If

End Function

Private Sub DsplyValuesFormatSet()
    Dim cllLast As Collection:  Set cllLast = NtryLst
    DsplyValueFormat thisformat:=sFrmtTcksSys, forvalue:=NtryTcksSys(cllLast)
    DsplyValueFormat thisformat:=sFrmtTcksElpsd, forvalue:=NtryTcksElpsd(cllLast)
    DsplyValueFormat thisformat:=sFrmtTcksGrss, forvalue:=NtryTcksGrss(cllLast)
    DsplyValueFormat thisformat:=sFrmtTcksOvrhdNtry, forvalue:=NtryTcksOvrhdNtry(cllLast)
    DsplyValueFormat thisformat:=sFrmtTcksOvrhdItm, forvalue:=NtryTcksOvrhdItmMax
    DsplyValueFormat thisformat:=sFrmtTcksNt, forvalue:=NtryTcksNt(cllLast)
    DsplyValueFormat thisformat:=sFrmtScsElpsd, forvalue:=NtryScsElpsd(cllLast)
    DsplyValueFormat thisformat:=sFrmtScsGrss, forvalue:=NtryScsGrss(cllLast)
    DsplyValueFormat thisformat:=sFrmtScsOvrhdNtry, forvalue:=NtryScsOvrhdNtry(cllLast)
    DsplyValueFormat thisformat:=sFrmtScsOvrhdItm, forvalue:=NtryScsOvrhdItm(cllLast)
    DsplyValueFormat thisformat:=sFrmtScsNt, forvalue:=NtryScsNt(cllLast)
End Sub

Private Function Max(ParamArray va() As Variant) As Variant
' ------------------------------------------------------
' Returns the maximum value of all values provided (va).
' ------------------------------------------------------
    Dim v   As Variant
    
    Max = va(LBound(va)): If LBound(va) = UBound(va) Then Exit Function
    For Each v In va
        If v > Max Then Max = v
    Next v
    
End Function

Private Function NtryTcksOvrhdItmMax() As Double
    
    Dim cll As Collection
    Dim dbl As Double
    Dim v   As Variant
    
    For Each v In cllTrc
        Set cll = v
        NtryTcksOvrhdItmMax = Max(NtryTcksOvrhdItmMax, NtryTcksOvrhdItm(cll))
    Next v

End Function

Public Sub EoC(ByVal itm As String, _
      Optional ByVal errinf As String = vbNullString)
' ---------------------------------------------------
'
' ---------------------------------------------------
#If ExecTrace Then
    Dim cll As Collection
    
    If StckIsEmpty Then Exit Sub
    If cllTrc Is Nothing Then Exit Sub
    TrcEnd itm:=itm, dir:=DIR_END_CODE, cll:=cll
    StckPop cll
#End If
End Sub

Public Sub EoP(ByVal itm As String, _
      Optional ByVal errinf As String = vbNullString)
' ---------------------------------------------------
' Trace of the End of a Procedure.
' ---------------------------------------------------
#If ExecTrace Then
    Dim cll As Collection
    
    If StckIsEmpty Then Exit Sub        ' Nothing to trace any longer
    If cllTrc Is Nothing Then Exit Sub  ' No trace or trace has finished
    TrcEnd itm:=itm, dir:=DIR_END_PROC, errinf:=errinf, cll:=cll
    StckPop cll
    If StckIsEmpty Then
        Dsply
    End If

#End If
End Sub

Private Sub ErrMsg(ByVal errno As Long, _
                   ByVal errsource As String, _
                   ByVal errdscrptn As String, _
                   ByVal errline As Long)
' ----------------------------------------------
' Display of a module's error message.
' ----------------------------------------------
    MsgBox Prompt:="Error description" & vbLf & err.Description, _
           buttons:=vbOKOnly, _
           Title:="VB Runtime error " & errno & " in " & errsource & IIf(errline <> 0, " at line " & errline, "")
End Sub

Private Function ErrSrc(ByVal sProc As String) As String
    ErrSrc = "mTrc." & sProc
End Function

Public Sub Finish( _
   Optional ByVal errinf As String = vbNullString)
' -------------------------------------------------
' Finishes an unfinished traced items by means of
' the stack. All items on the the stack are
' processed via EoP/EoC.
' -------------------------------------------------
    Dim cll As Collection
    
    While Not StckIsEmpty
        Set cll = StckTop
        If NtryIsCode(cll) _
        Then mTrc.EoC itm:=NtryItm(cll), errinf:=errinf _
        Else mTrc.EoP itm:=NtryItm(cll), errinf:=errinf
        errinf = vbNullString
    Wend
    
End Sub

Private Sub Initialize()
' -------------------------------------------------------
' Public allows an error handler to initialize the trace.
' -------------------------------------------------------
    Set cllTrc = New Collection
    Set cllNtryLast = Nothing
    dtTraceBegin = Now()
    cyTcksOvrhdItm = 0
    iTrcLvl = 0
    cySysFrequency = 0
    sFirstTraceItem = vbNullString
    If lPrecision = 0 Then lPrecision = 6 ' default when another valus has not been set by the caller
    
End Sub

Private Function Ntry( _
                ByVal tcks As Currency, _
                ByVal dir As String, _
                ByVal itm As String, _
                ByVal lvl As Long, _
                ByVal err As String) As Collection
' ------------------------------------------------
' Return the arguments as items of a collection.
' This "entry" is used
' ------------------------------------------------
      
    Dim cll As New Collection
    
    NtryTcksSys(cll) = tcks
    NtryDrctv(cll) = dir
    NtryItm(cll) = itm
    NtryCllLvl(cll) = lvl
    NtryError(cll) = err
    Set Ntry = cll
    
End Function

 Private Function NtryIsBegin( _
                  ByVal v As Collection, _
         Optional ByRef cll As Collection = Nothing) As Boolean
' -------------------------------------------------------------
' Returns TRUE and v as cll when the entry is a begin entry,
' else FALSE and cll = Nothing
' ---------------------------------------------------
    If InStr(NtryDrctv(v), DIR_BEGIN_ID) <> 0 Then
        NtryIsBegin = True
        Set cll = v
    End If
End Function

Private Function NtryIsCode( _
           ByVal cll As Collection) As Boolean
    Dim drctv As String
    
    Select Case NtryDrctv(cll)
        Case DIR_BEGIN_CODE, DIR_END_CODE: NtryIsCode = True
    End Select
End Function

Private Function NtryIsEnd( _
                 ByVal v As Collection, _
                 ByRef cll As Collection) As Boolean
' --------------------------------------------------
' Returns TRUE and v as cll when the entry is an end
' entry, else FALSE and cll = Nothing
' --------------------------------------------------
    If InStr(NtryDrctv(v), DIR_END_ID) <> 0 Then
        NtryIsEnd = True
        Set cll = v
    End If
End Function

Private Function NtryLastBegin() As Long
    
    Dim i As Long
    
    For i = cllTrc.Count To 1 Step -1
        If NtryIsBegin(cllTrc(i)) Then
            NtryLastBegin = i
            Exit Function
        End If
    Next i
    
End Function

Private Function NtryLst() As Collection ' Retun last entry
    Set NtryLst = cllTrc(cllTrc.Count)
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

Private Function StckIsEmpty() As Boolean
    StckIsEmpty = cllStck Is Nothing
    If Not StckIsEmpty Then StckIsEmpty = cllStck.Count = 0
End Function

Private Sub StckPop(ByVal pop As Collection)
' ------------------------------------------
'
' ------------------------------------------
    Dim cllTop As Collection: Set cllTop = StckTop
    
    If NtryItm(pop) <> NtryItm(cllTop) And Not NtryIsCode(pop) And NtryIsCode(cllTop) Then
        '~~ There is an unfinished code trace still on the stack which needs to be finished first
        mTrc.EoC NtryItm(cllTop)
    End If
    
    If NtryItm(pop) = NtryItm(cllTop) Then
        cllStck.Remove cllStck.Count
    Else
        Debug.Print "Stack Pop ='" & NtryItm(pop) _
                  & "', Stack Top = '" & NtryItm(cllTop) _
                  & "', Stack Dir = '" & NtryDrctv(cllTop) _
                  & "', Stack Lvl = '" & NtryCllLvl(cllTop) _
                  & "', Stack Cnt = '" & cllStck.Count
        Stop
    End If
End Sub

Private Sub StckPush(ByVal cll As Collection)
      
    If cllStck Is Nothing _
    Then Set cllStck = New Collection
    cllStck.Add cll

End Sub

Private Function StckTop() As Collection
    Set StckTop = cllStck(cllStck.Count)
End Function

Public Sub Terminate()
' -----------------------------------------------------------------
' Should be called by any error handling when a new execution trace
' is about to begin with the very first procedure's execution.
' -----------------------------------------------------------------
    Set cllTrc = Nothing
    Set cllStck = Nothing
End Sub

Private Sub TrcAdd( _
            ByVal itm As String, _
            ByVal tcks As Currency, _
            ByVal dir As String, _
            ByVal lvl As Long, _
   Optional ByVal err As String = vbNullString, _
   Optional ByRef cll As Collection)
   
    If Not cllNtryLast Is Nothing Then
        '~~ When this is not the first entry added, save the overhead ticks caused by the previous entry
        '~~ Note: Would corrupt the overhead when saved with the entry itself because the overhead is
        '~~       the ticks caused by the collection of the entry
        NtryTcksOvrhdNtry(cllNtryLast) = cyTcksOvrhd
    End If
    
    Set cll = Ntry(tcks:=tcks, dir:=dir, itm:=itm, lvl:=lvl, err:=err)
    cllTrc.Add cll
    Set cllNtryLast = cll

End Sub

Private Sub TrcBgn(ByVal itm As String, _
                   ByVal dir As String, _
          Optional ByRef cll As Collection)
' -----------------------------------------
' Collect a trace begin entry with the
' current ticks count for the procedure or
' code (item).
' -----------------------------------------
    
    Dim cy  As Currency:    cy = SysCrrntTcks
    
    iTrcLvl = iTrcLvl + 1
    TrcAdd tcks:=cy, dir:=dir, itm:=itm, lvl:=iTrcLvl, cll:=cll
    cyTcksOvrhd = SysCrrntTcks - cy ' overhead ticks caused by the collection of the begin trace entry
    
End Sub

Private Sub TrcEnd( _
            ByVal itm As String, _
   Optional ByVal dir As String = vbNullString, _
   Optional ByVal errinf As String = vbNullString, _
   Optional ByRef cll As Collection)
' ------------------------------------------------
' Collect an end trace entry with the current
' ticks count for the procedure or code (item).
' ------------------------------------------------
    Const PROC = "TrcEnd"
    
    On Error GoTo eh
    Dim cy  As Currency:    cy = SysCrrntTcks

    If errinf <> vbNullString Then
        errinf = COMMENT & errinf & COMMENT
    End If
    
    TrcAdd itm:=itm, tcks:=cy, dir:=Trim(dir), lvl:=iTrcLvl, err:=errinf, cll:=cll
    iTrcLvl = iTrcLvl - 1
    cyTcksOvrhd = SysCrrntTcks - cy ' overhead ticks caused by the collection of the begin trace entry

xt: Exit Sub
    
eh: ErrMsg errno:=err.Number, errsource:=ErrSrc(PROC), errdscrptn:=err.Description, errline:=Erl
#If Debugging Then
    Stop: Resume
#End If
End Sub

Private Function TrcIsEmpty() As Boolean
    TrcIsEmpty = cllTrc Is Nothing
    If Not TrcIsEmpty Then TrcIsEmpty = cllTrc.Count = 0
End Function

