Attribute VB_Name = "mTrc"
Option Explicit
' ----------------------------------------------------------------------------
' Standard Module mTrc
'       Services to trace the execution of procedures and code snippets with
'       automated display of the trace result. Any trace activity is triggered
'       by the Conditional Compile Argument ExecTrace = 1. When not activated
'       this way the negative effect on the performance is close to absolutely
'       none. Even when activated the effect is less then 0.01% of the overall
'       execution time. Execution time is traced with the highest possible
'       precision.
'
' Uses: mMsg.Dsply to display the trace result
'
' W. Rauschenberger, Berlin, Nov. 2021
' ------------------------------------------------------------------------
Public Enum enDsplydInfo
    Detailed = 1
    Compact = 2
End Enum

Private Enum enTraceInfo
    enItmDrctv = 1
    enItmId
    enItmInf
    enItmLvl
    enItmTcksSys
    enPosItmArgs
End Enum

Private Enum enNtry
    enItm
    enScsElpsd
    enScsGrss
    enScsNt
    enScsOvrhdItm
    enScsOvrhdNtry
    enTcksElpsd
    enTcksGrss
    enTcksOvrhdItm
    enTcksOvrhdNtry
End Enum

Private Declare PtrSafe Function getFrequency Lib "kernel32" _
Alias "QueryPerformanceFrequency" (cySysFrequency As Currency) As Long
Private Declare PtrSafe Function getTickCount Lib "kernel32" _
Alias "QueryPerformanceCounter" (cyTickCount As Currency) As Long

Private Const DIR_BEGIN_ID  As String = ">"     ' Begin procedure or code trace indicator
Private Const DIR_END_ID    As String = "<"     ' End procedure or code trace indicator
Private Const TRC_INFO_DELIM As String = " !!! "

Private cllNtryLast         As Collection       '
Private cySysFrequency      As Currency         ' Execution Trace SysFrequency (initialized with init)
Private cyTcksOvrhd         As Currency         ' Overhead ticks caused by the collection of a traced item's entry
Private cyTcksOvrhdItm      As Currency         ' Execution Trace time accumulated by caused by the time tracking itself
Private cyTcksOvrhdTrc      As Currency         ' Overhead ticks caused by the collection of a traced item's entry
Private cyTcksOvrhdTrcStrt  As Currency         ' Overhead ticks caused by the collection of a traced item's entry
Private cyTcksPaused        As Currency         ' Accumulated with procedure Continue
Private cyTcksPauseStart    As Currency         ' Set with procedure Pause
Private dtTraceBegin        As Date             ' Initialized at start of execution trace
Private iTrcLvl             As Long             ' Increased with each begin entry and decreased with each end entry
Private lDisplayedInfo      As Long             ' Detailed or Compact, defaults to Compact
Private lPrecision          As Long             ' Precision for displayed seconds, defaults to 6 decimals (0,000000)
Private sFirstTraceItem     As String
Private sFrmtScsElpsd       As String           ' -------------
Private sFrmtScsGrss        As String           ' Format String
Private sFrmtScsNt          As String           ' for Seconds
Private sFrmtScsOvrhdItm    As String           ' -------------
Private sFrmtScsOvrhdNtry   As String           ' -------------
Private sFrmtTcksElpsd      As String           ' -------------
Private sFrmtTcksGrss       As String           '
Private sFrmtTcksNt         As String           ' Format String
Private sFrmtTcksOvrhdItm   As String           ' for Ticks
Private sFrmtTcksOvrhdNtry  As String           '
Private sFrmtTcksSys        As String           ' -------------
Private sTraceLogFile       As String           ' When not vbNullString the trace is written into file and tzhe display is suspended
Private Trace               As Collection       ' Collection of begin and end trace entries
Private TraceStack          As Collection       ' Trace stack

Private Property Get DIR_BEGIN_CODE() As String:            DIR_BEGIN_CODE = DIR_BEGIN_ID:                  End Property

Private Property Get DIR_BEGIN_PROC() As String:            DIR_BEGIN_PROC = VBA.String$(2, DIR_BEGIN_ID):  End Property

Private Property Get DIR_END_CODE() As String:              DIR_END_CODE = DIR_END_ID:                      End Property

Private Property Get DIR_END_PROC() As String:              DIR_END_PROC = VBA.String$(2, DIR_END_ID):      End Property

Private Property Get DisplayedInfo() As enDsplydInfo:       DisplayedInfo = lDisplayedInfo:                 End Property

Public Property Let DisplayedInfo(ByVal l As enDsplydInfo): lDisplayedInfo = l:                             End Property

Private Property Get DisplayedPrecision() As Long:          DisplayedPrecision = lPrecision:                End Property

Public Property Let DisplayedPrecision(ByVal l As Long):    lPrecision = l:                                 End Property

Private Property Get DsplyLnIndnttn(Optional ByRef trc_entry As Collection) As String
    DsplyLnIndnttn = RepeatStrng("|  ", ItmLvl(trc_entry))
End Property

Private Property Get ItmArgs(Optional ByRef trc_entry As Collection) As Variant
    ItmArgs = trc_entry("I")(enPosItmArgs)
End Property

Private Property Get ItmDrctv(Optional ByRef trc_entry As Collection) As String
    ItmDrctv = trc_entry("I")(enItmDrctv)
End Property

Private Property Get ItmId(Optional ByRef trc_entry As Collection) As String
    ItmId = trc_entry("I")(enItmId)
End Property

Private Property Get ItmInf(Optional ByRef trc_entry As Collection) As String
    On Error Resume Next ' in case this has never been collected
    ItmInf = trc_entry("I")(enItmInf)
    If Err.Number <> 0 Then ItmInf = vbNullString
End Property

Private Property Get ItmLvl(Optional ByRef trc_entry As Collection) As Long
    ItmLvl = trc_entry("I")(enItmLvl)
End Property

Private Property Get ItmTcksSys(Optional ByRef trc_entry As Collection) As Currency
    ItmTcksSys = trc_entry("I")(enItmTcksSys)
End Property

Private Property Let NtryItm(Optional ByVal trc_entry As Collection, ByVal v As Variant)
    trc_entry.Add v, "I"
End Property

Private Property Get NtryScsElpsd(Optional ByRef trc_entry As Collection) As Currency
    On Error Resume Next
    NtryScsElpsd = trc_entry("SE")
    If Err.Number <> 0 Then NtryScsElpsd = Space$(Len(sFrmtScsElpsd))
End Property

Private Property Let NtryScsElpsd(Optional ByRef trc_entry As Collection, ByRef cy As Currency)
    trc_entry.Add cy, "SE"
End Property

Private Property Get NtryScsGrss(Optional ByRef trc_entry As Collection) As Currency
    On Error Resume Next ' in case no value exists (the case for each begin entry)
    NtryScsGrss = trc_entry("SG")
    If Err.Number <> 0 Then NtryScsGrss = Space$(Len(sFrmtScsGrss))
End Property

Private Property Let NtryScsGrss(Optional ByRef trc_entry As Collection, ByRef cy As Currency)
    trc_entry.Add cy, "SG"
End Property

Private Property Get NtryScsNt(Optional ByRef trc_entry As Collection) As Double
    On Error Resume Next
    NtryScsNt = trc_entry("SN")
    If Err.Number <> 0 Then NtryScsNt = Space$(Len(sFrmtScsNt))
End Property

Private Property Let NtryScsNt(Optional ByRef trc_entry As Collection, ByRef dbl As Double)
    trc_entry.Add dbl, "SN"
End Property

Private Property Get NtryScsOvrhdItm(Optional ByRef trc_entry As Collection) As Double
    On Error Resume Next
    NtryScsOvrhdItm = trc_entry("SOI")
    If Err.Number <> 0 Then NtryScsOvrhdItm = Space$(Len(sFrmtScsOvrhdItm))
End Property

Private Property Let NtryScsOvrhdItm(Optional ByRef trc_entry As Collection, ByRef dbl As Double)
    trc_entry.Add dbl, "SOI"
End Property

Private Property Get NtryScsOvrhdNtry(Optional ByRef trc_entry As Collection) As Double
    On Error Resume Next
    NtryScsOvrhdNtry = trc_entry("SON")
    If Err.Number <> 0 Then NtryScsOvrhdNtry = Space$(Len(sFrmtScsOvrhdItm))
End Property

Private Property Let NtryScsOvrhdNtry(Optional ByRef trc_entry As Collection, ByRef dbl As Double)
    trc_entry.Add dbl, "SON"
End Property

Private Property Get NtryTcksElpsd(Optional ByRef trc_entry As Collection) As Currency
    NtryTcksElpsd = trc_entry("TE")
End Property

Private Property Let NtryTcksElpsd(Optional ByRef trc_entry As Collection, ByRef cy As Currency)
    trc_entry.Add cy, "TE"
End Property

Private Property Get NtryTcksGrss(Optional ByRef trc_entry As Collection) As Currency
    On Error Resume Next
    NtryTcksGrss = trc_entry("TG")
    If Err.Number <> 0 Then NtryTcksGrss = 0
End Property

Private Property Let NtryTcksGrss(Optional ByRef trc_entry As Collection, ByRef cy As Currency)
    trc_entry.Add cy, "TG"
End Property

Private Property Get NtryTcksNt(Optional ByRef trc_entry As Collection) As Currency
    On Error Resume Next
    NtryTcksNt = trc_entry("TN")
    If Err.Number <> 0 Then NtryTcksNt = 0
End Property

Private Property Let NtryTcksNt(Optional ByRef trc_entry As Collection, ByRef cy As Currency)
    trc_entry.Add cy, "TN"
End Property

Private Property Get NtryTcksOvrhdItm(Optional ByRef trc_entry As Collection) As Currency
    On Error Resume Next
    NtryTcksOvrhdItm = trc_entry("TOI")
    If Err.Number <> 0 Then NtryTcksOvrhdItm = 0
End Property

Private Property Let NtryTcksOvrhdItm(Optional ByRef trc_entry As Collection, ByRef cy As Currency)
    trc_entry.Add cy, "TOI"
End Property

Private Property Get NtryTcksOvrhdNtry(Optional ByRef trc_entry As Collection) As Currency
    On Error Resume Next
    NtryTcksOvrhdNtry = trc_entry("TON")
    If Err.Number <> 0 Then NtryTcksOvrhdNtry = 0
End Property

Private Property Let NtryTcksOvrhdNtry(Optional ByRef trc_entry As Collection, ByRef cy As Currency)
    If trc_entry Is Nothing Then Set trc_entry = New Collection
    trc_entry.Add cy, "TON"
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

Public Property Let TraceLogFile( _
                        Optional ByVal tl_append As Boolean = False, _
                                 ByVal tl_file As String)
    sTraceLogFile = tl_file
    With New FileSystemObject
        If Not tl_append Then
            If .FileExists(tl_file) Then .DeleteFile tl_file, True
        End If
    End With
End Property

Private Property Let TraceLogTxt( _
                        Optional ByVal tl_file As Variant, _
                        Optional ByVal tl_append As Boolean = True, _
                        Optional ByRef tl_split As String, _
                                 ByVal tl_string As String)
' ----------------------------------------------------------------------------
' Writes the string (tl_string) into the file (tl_file) which might be a file
' object or a file's full name.
' Note: tl_split is not used but specified to comply with Property Get.
' ----------------------------------------------------------------------------
    Const PROC = "TraceLogTxt-Let"
    
    On Error GoTo eh
    Dim fso As New FileSystemObject
    Dim ts  As TextStream
    Dim sFl As String
   
    tl_split = tl_split ' not used! just for coincidence with Get
    With fso
        If TypeName(tl_file) = "File" Then
            sFl = tl_file.Path
        Else
            '~~ tl_file is regarded a file's full name, created if not existing
            sFl = tl_file
            If Not .FileExists(sFl) Then .CreateTextFile sFl
        End If
        
        If tl_append _
        Then Set ts = .OpenTextFile(Filename:=sFl, IOMode:=ForAppending) _
        Else Set ts = .OpenTextFile(Filename:=sFl, IOMode:=ForWriting)
    End With
    
    ts.WriteLine tl_string

xt: ts.Close
    Set fso = Nothing
    Set ts = Nothing
    Exit Property
    
eh: Select Case ErrMsg(ErrSrc(PROC))
        Case vbResume:  Stop: Resume
        Case Else:      GoTo xt
    End Select
End Property

Private Function AppErr(ByVal app_err_no As Long) As Long
' ------------------------------------------------------------------------------
' Ensures that a programmed (i.e. an application) error numbers never conflicts
' with the number of a VB runtime error. Thr function returns a given positive
' number (app_err_no) with the vbObjectError added - which turns it into a
' negative value. When the provided number is negative it returns the original
' positive "application" error number e.g. for being used with an error message.
' ------------------------------------------------------------------------------
    If app_err_no >= 0 Then AppErr = app_err_no + vbObjectError Else AppErr = Abs(app_err_no - vbObjectError)
End Function

Public Sub BoC(ByVal boc_id As String, _
          ParamArray boc_arguments() As Variant)
' ----------------------------------------------
' Begin of code sequence trace.
' ----------------------------------------------
#If ExecTrace = 1 Then
    Dim cll             As Collection
    Dim vArguments()    As Variant
    
    cyTcksOvrhdTrcStrt = SysCrrntTcks
    vArguments = boc_arguments
    TrcBgn trc_id:=boc_id, trc_dir:=DIR_BEGIN_CODE, trc_args:=vArguments, trc_cll:=cll
    cyTcksOvrhdTrc = SysCrrntTcks - cyTcksOvrhdTrcStrt ' overhead ticks caused by the collection of the begin trace entry
#End If
End Sub

Public Sub BoP(ByVal bop_id As String, _
          ParamArray bop_arguments() As Variant)
' ----------------------------------------------
' Begin of procedure trace.
' ----------------------------------------------
#If ExecTrace = 1 Then
    Dim cll           As Collection
    Dim vArguments()  As Variant
    
    cyTcksOvrhdTrcStrt = SysCrrntTcks
    vArguments = bop_arguments
    If TrcIsEmpty Then
        Initialize
        sFirstTraceItem = bop_id
    Else
        If bop_id = sFirstTraceItem Then
            '~~ A previous trace had not come to a regular end and thus will be erased
            Set Trace = Nothing
            Initialize
        End If
    End If
    TrcBgn trc_id:=bop_id, trc_dir:=DIR_BEGIN_PROC, trc_args:=vArguments, trc_cll:=cll
    cyTcksOvrhdTrc = SysCrrntTcks - cyTcksOvrhdTrcStrt ' overhead ticks caused by the collection of the begin trace entry
#End If
End Sub

Public Sub BoP_ErH(ByVal bopeh_id As String, _
                   ByVal bopeh_args As Variant)
' ---------------------------------------------
' Begin of procedure trace, specifically for
' being used by the mErH module.
' ---------------------------------------------
#If ExecTrace = 1 Then
    Dim cll           As Collection
    
    cyTcksOvrhdTrcStrt = SysCrrntTcks
    If TrcIsEmpty Then
        Initialize
        sFirstTraceItem = bopeh_id
    Else
        If bopeh_id = sFirstTraceItem Then
            '~~ A previous trace had not come to a regular end and thus will be erased
            Set Trace = Nothing
            Initialize
        End If
    End If
    TrcBgn trc_id:=bopeh_id, trc_dir:=DIR_BEGIN_PROC, trc_args:=bopeh_args, trc_cll:=cll
    cyTcksOvrhdTrc = SysCrrntTcks - cyTcksOvrhdTrcStrt ' overhead ticks caused by the collection of the begin trace entry
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
    
    For Each v In Trace
        Set cll = v
        NtryScsOvrhdNtry(cll) = CDbl(NtryTcksOvrhdNtry(cll)) / CDbl(SysFrequency)
        NtryScsElpsd(cll) = CDec(NtryTcksElpsd(cll)) / CDec(SysFrequency)
        If Not NtryIsBegin(cll) Then
            NtryScsGrss(cll) = CDec(NtryTcksGrss(cll)) / CDec(SysFrequency)
            NtryScsOvrhdItm(cll) = CDec(NtryTcksOvrhdItm(cll)) / CDec(SysFrequency)
            NtryScsNt(cll) = CDec(NtryTcksNt(cll)) / CDec(SysFrequency)
        End If
    Next v

xt: Set cll = Nothing
    Exit Sub

eh: Select Case ErrMsg(ErrSrc(PROC))
        Case vbYes: Stop: Resume
        Case vbNo:  Stop: Resume Next
        Case Else:  GoTo xt
    End Select
End Sub

Private Function ComputeSecsOvrhdTtlEntry() As Double
' --------------------------------------------------
' Returns the total overhead seconds caused by the
' collection of the traced item's data.
' --------------------------------------------------
    Const PROC = "ComputeSecsOvrhdTtlEntry"
    
    On Error GoTo eh
    Dim v   As Variant
    Dim cll As Collection
    Dim dbl As Double
    
    For Each v In Trace
        Set cll = v
        dbl = dbl + NtryScsOvrhdNtry(cll)
    Next v
    ComputeSecsOvrhdTtlEntry = dbl

xt: Exit Function

eh: Select Case ErrMsg(ErrSrc(PROC))
        Case vbYes: Stop: Resume
        Case vbNo:  Stop: Resume Next
        Case Else:  GoTo xt
    End Select
    Set Trace = Nothing
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
    
    For Each v In Trace
        Set cll = v
        If cyTcksBegin = 0 Then cyTcksBegin = ItmTcksSys(cll)
        cyTcksElapsed = ItmTcksSys(cll) - cyTcksBegin
        NtryTcksElpsd(cll) = cyTcksElapsed
    Next v

xt: Set cll = Nothing
    Exit Sub

eh: Select Case ErrMsg(ErrSrc(PROC))
        Case vbYes: Stop: Resume
        Case vbNo:  Stop: Resume Next
        Case Else:  GoTo xt
    End Select
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
    
    For Each v In Trace
        Set cll = v
        NtryTcksNt(cll) = NtryTcksGrss(cll) - NtryTcksOvrhdNtry(cll)
    Next v

xt: Set cll = Nothing
    Exit Sub

eh: Select Case ErrMsg(ErrSrc(PROC))
        Case vbYes: Stop: Resume
        Case vbNo:  Stop: Resume Next
        Case Else:  GoTo xt
    End Select
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
    
    For Each v In Trace
        Set cll = v
        cy = cy + NtryTcksOvrhdNtry(cll)
    Next v
    ComputeTcksOvrhdItem = cy

xt: Exit Function

eh: Select Case ErrMsg(ErrSrc(PROC))
        Case vbYes: Stop: Resume
        Case vbNo:  Stop: Resume Next
        Case Else:  GoTo xt
    End Select
    Set Trace = Nothing
End Function

Public Sub Continue()
    cyTcksPaused = cyTcksPaused + (SysCrrntTcks - cyTcksPauseStart)
End Sub

Public Sub Dsply()
    Const PROC = "Dsply"
    
    On Error GoTo eh
    Dim dct         As Dictionary
    Dim v           As Variant
    Dim cllTrcEntry As Collection
    Dim sTrace      As String
    Dim lLenHeader  As Long
    Dim SctnLabel   As TypeMsgLabel
    Dim SctnText    As TypeMsgText
    Dim TrcResult   As TypeMsg
    
    If TrcIsEmpty Or sTraceLogFile <> vbNullString Then GoTo xt
    
    NtryTcksOvrhdNtry(cllNtryLast) = cyTcksOvrhd
    If DisplayedInfo = 0 Then DisplayedInfo = Compact

    If Not DsplyNtryAllCnsstnt(dct) Then
        mTrc.Terminate
        GoTo xt
    End If
        
    ComputeTcksNet
    ComputeSecsGrssOvrhdNet    ' Calculate gross and net execution seconds
    DsplyValuesFormatSet       ' Setup the format for the displayed ticks and sec values
    
    sTrace = DsplyHdr(lLenHeader)
    For Each v In Trace
        Set cllTrcEntry = v
        sTrace = sTrace & vbLf & DsplyLn(cllTrcEntry)
        If DsplyArgs(cllTrcEntry) <> vbNullString Then
        End If
    Next v
    sTrace = sTrace & vbLf & DsplyFtr(lLenHeader)
    With TrcResult.Section(1).Text
        .Text = sTrace
        .MonoSpaced = True
    End With
    With TrcResult.Section(2)
        .Label.Text = "About overhead, precision, etc.:"
        .Label.FontColor = rgbBlue
        .Text.Text = DsplyAbout
        .Text.MonoSpaced = True
        .Text.FontSize = 8
    End With
    
    mMsg.Dsply dsply_title:="Execution Trace " & Format(Now(), "MM-DD hh:mm:ss") & _
                            " (displayed because the Conditional Compile Argument ""ExecTrace = 1""!)" _
             , dsply_msg:=TrcResult _
             , dsply_width_min:=60 _
             , dsply_modeless:=True
            
xt: mTrc.Terminate
    Exit Sub
    
eh: Select Case ErrMsg(ErrSrc(PROC))
        Case vbYes: Stop: Resume
        Case vbNo:  Resume Next
        Case Else:  GoTo xt
    End Select
End Sub

Private Function DsplyAbout() As String
    
    Dim cyTcksOvrhdItm      As Currency
    Dim dblOvrhdPcntg       As Double
    Dim dblTtlScsOvrhdNtry  As Double
    
    dblTtlScsOvrhdNtry = ComputeSecsOvrhdTtlEntry
    cyTcksOvrhdItm = ComputeTcksOvrhdItem
    dblOvrhdPcntg = (dblTtlScsOvrhdNtry / NtryScsElpsd(NtryLst)) * 100
    
    DsplyAbout = "> The trace itself, i.e. the collection of the begin and end data for each traced procedure or code snippet " & vbLf _
               & "  caused a performance loss of " & Format$(dblTtlScsOvrhdNtry, sFrmtScsOvrhdItm) & " seconds (=" & Format$(dblOvrhdPcntg, "0.00") & "%). " _
               & "For a best possible execution time precision" & vbLf _
               & "  the overhead per traced item has been deducted from each of the " & Trace.Count / 2 & " traced item's execution time." & vbLf _
               & "> The precision (decimals) for the displayed seconds defaults to 0,000000 (6 decimals)." & vbLf _
               & "  This may be changed via the 'DisplayedPrecision' Property." & vbLf _
               & "> Though the traced execution time comes with the highest possible precisssion it will vary from execution" & vbLf _
               & "  to execution because of different system conditions. For an estimation of the average execution time and" & vbLf _
               & "  the possible time spread, the trace will have to be repeated several times." & vbLf _
               & "> When an error had been displayed the trace will have paused and continued when the a reply button is pressed." & vbLf _
               & "  For a best possible correct execution time trace any paused time is subtracted."

End Function

Private Function DsplyArgName(ByVal s As String) As Boolean
    If Right(s, 1) = ":" _
    Or Right(s, 1) = "=" _
    Or Right(s, 2) = ": " _
    Or Right(s, 2) = " :" _
    Or Right(s, 2) = "= " _
    Or Right(s, 2) = " =" _
    Or Right(s, 3) = " : " _
    Or Right(s, 3) = " = " _
    Then DsplyArgName = True
End Function

Private Function DsplyArgs(ByVal trc_entry As Collection) As String
' -------------------------------------------------------------
' Returns a string with the collection of the traced arguments
' Any entry ending with a ":" or "=" is an arguments name with
' its value in the subsequent item.
' -------------------------------------------------------------
    Dim va()    As Variant
    Dim i       As Long
    Dim sL      As String
    Dim sR      As String
    
    On Error Resume Next
    va = ItmArgs(trc_entry)
    If Err.Number <> 0 Then Exit Function
    i = LBound(va)
    If Err.Number <> 0 Then Exit Function
    
    For i = i To UBound(va)
        If DsplyArgs = vbNullString Then
            ' This is the very first argument
            If DsplyArgName(va(i)) Then
                ' The element is the name of an argument followed by a subsequent value
                DsplyArgs = "|  " & va(i) & CStr(va(i + 1))
                i = i + 1
            Else
                sL = ">": sR = "<"
                DsplyArgs = "|  Argument values: " & sL & va(i) & sR
            End If
        Else
            If DsplyArgName(va(i)) Then
                ' The element is the name of an argument followed by a subsequent value
                DsplyArgs = DsplyArgs & ", " & va(i) & CStr(va(i + 1))
                i = i + 1
            Else
                sL = ">": sR = "<"
                DsplyArgs = DsplyArgs & "  " & sL & va(i) & sR
            End If
        End If
    Next i
End Function

Private Function DsplyFtr(ByVal lLenHeaderData As Long) ' Displayed trace footer
    DsplyFtr = _
        Space$(lLenHeaderData) _
      & DIR_END_PROC _
      & " End execution trace " _
      & Format$(Now(), "hh:mm:ss")
End Function

Private Function DsplyHdr(ByRef lLenHeaderData As Long) As String
' ---------------------------------------------------------------
' Compact:  Header1 = Elapsed Exec >> Start
' Detailed: Header1 =    Ticks   |   Secs
'           Header2 = xxx xxxx xx xxx xxx xx
'           Header3 " --- ---- -- --- --- -- >> Start
' ---------------------------------------------------------------
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
    & Format$(dtTraceBegin, "hh:mm:ss") _
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
                    RepeatStrng$("-", Len(sFrmtScsElpsd)) _
            & " " & RepeatStrng$("-", Len(sFrmtScsNt)) _
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
                    RepeatStrng$("-", Len(sFrmtTcksSys)) _
            & " " & RepeatStrng$("-", Len(sFrmtTcksElpsd)) _
            & " " & RepeatStrng$("-", Len(sFrmtTcksGrss)) _
            & " " & RepeatStrng$("-", Len(sFrmtTcksOvrhdItm)) _
            & " " & RepeatStrng$("-", Len(sFrmtTcksOvrhdItm)) _
            & " " & RepeatStrng$("-", Len(sFrmtTcksNt)) _
            & " " & RepeatStrng$("-", Len(sFrmtScsElpsd)) _
            & " " & RepeatStrng$("-", Len(sFrmtScsGrss)) _
            & " " & RepeatStrng$("-", Len(sFrmtScsOvrhdItm)) _
            & " " & RepeatStrng$("-", Len(sFrmtScsOvrhdItm)) _
            & " " & RepeatStrng$("-", Len(sFrmtScsNt)) _
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

eh: Select Case ErrMsg(ErrSrc(PROC))
        Case vbYes: Stop: Resume
        Case vbNo:  Stop: Resume Next
        Case Else:  GoTo xt
    End Select
    Set Trace = Nothing
End Function

Public Function DsplyHdrCntrAbv(ByVal s1 As String, _
                                ByVal s2 As String, _
                       Optional ByVal sLeft As String = vbNullString, _
                       Optional ByVal sRight As String = vbNullString, _
                       Optional ByVal sFillChar As String = " ") As String
' ------------------------------------------------------------------------
' Returns s1 centered above s2 considering any characters left and right.
' ------------------------------------------------------------------------
    
    If Len(s1) > Len(s2) Then
        DsplyHdrCntrAbv = s1
    Else
        DsplyHdrCntrAbv = s1 & String$(Int((Len(s2) - Len(s1 & sLeft & sRight)) / 2), sFillChar)
        DsplyHdrCntrAbv = String$(Len(s2) + -Len(sLeft & DsplyHdrCntrAbv & sRight), sFillChar) & DsplyHdrCntrAbv
        DsplyHdrCntrAbv = sLeft & DsplyHdrCntrAbv & sRight
    End If
    
End Function

Private Function DsplyLn(ByVal trc_entry As Collection) As String
' -------------------------------------------------------------
' Returns a trace line for being displayed.
' -------------------------------------------------------------
    Const PROC = "DsplyLn"
    On Error GoTo eh
    Dim lLenData    As Long
    Dim sArgs       As String: sArgs = DsplyArgs(trc_entry)
    Dim sArgsLine   As String
    
    Select Case DisplayedInfo
        Case Compact
            DsplyLn = _
                      DsplyValue(trc_entry, NtryScsElpsd(trc_entry), sFrmtScsElpsd) _
              & " " & DsplyValue(trc_entry, NtryScsNt(trc_entry), sFrmtScsNt)
            
            If sArgs <> vbNullString _
            Then sArgsLine = vbLf & " " & String(Len(DsplyLn), " ") & DsplyLnIndnttn(trc_entry) & sArgs
            
            DsplyLn = DsplyLn _
              & " " & DsplyLnIndnttn(trc_entry) _
                    & ItmDrctv(trc_entry) _
              & " "
              
              lLenData = Len(DsplyLn)
              DsplyLn = DsplyLn _
                    & ItmId(trc_entry) _
              & " " & ItmInf(trc_entry)
                          
        Case Detailed
            DsplyLn = _
                        DsplyValue(trc_entry, ItmTcksSys(trc_entry), sFrmtTcksSys) _
                & " " & DsplyValue(trc_entry, NtryTcksElpsd(trc_entry), sFrmtTcksElpsd) _
                & " " & DsplyValue(trc_entry, NtryTcksGrss(trc_entry), sFrmtTcksGrss) _
                & " " & DsplyValue(trc_entry, NtryTcksOvrhdNtry(trc_entry), sFrmtTcksOvrhdItm) _
                & " " & DsplyValue(trc_entry, NtryTcksOvrhdItm(trc_entry), sFrmtTcksOvrhdItm) _
                & " " & DsplyValue(trc_entry, NtryTcksNt(trc_entry), sFrmtTcksNt) _
                & " " & DsplyValue(trc_entry, NtryScsElpsd(trc_entry), sFrmtScsElpsd) _
                & " " & DsplyValue(trc_entry, NtryScsGrss(trc_entry), sFrmtScsGrss) _
                & " " & DsplyValue(trc_entry, NtryScsOvrhdNtry(trc_entry), sFrmtScsOvrhdItm) _
                & " " & DsplyValue(trc_entry, NtryScsOvrhdItm(trc_entry), sFrmtScsOvrhdItm) _
                & " " & DsplyValue(trc_entry, NtryScsNt(trc_entry), sFrmtScsNt) _
                & " "

            lLenData = Len(DsplyLn)
            
            If sArgs <> vbNullString _
            Then sArgsLine = vbLf & String(Len(DsplyLn), " ") & DsplyLnIndnttn(trc_entry) & sArgs
                
            DsplyLn = DsplyLn _
                  & DsplyLnIndnttn(trc_entry) _
                  & ItmDrctv(trc_entry) _
            & " " & ItmId(trc_entry) _
            & " " & ItmInf(trc_entry)
    
    End Select
    DsplyLn = DsplyLn & sArgsLine

xt: Exit Function

eh: Select Case ErrMsg(ErrSrc(PROC))
        Case vbYes: Stop: Resume
        Case vbNo:  Stop: Resume Next
        Case Else:  GoTo xt
    End Select
    Set Trace = Nothing
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
    Dim v               As Variant
    Dim cllEndEntry     As Collection
    Dim cllBeginEntry   As Collection
    Dim bConsistent     As Boolean
    Dim i               As Long
    Dim j               As Long
    Dim sComment        As String
    Dim sTrace          As String
    Dim SctnLabel       As TypeMsgLabel
    Dim SctnText        As TypeMsgText
    
    If dct Is Nothing Then Set dct = New Dictionary
        
    ComputeTcksElpsd ' Calculates the ticks elapsed since trace start
    
    '~~ Check for missing corresponding end entries while calculating the execution time for each end entry.
    For i = 1 To NtryLastBegin
        If NtryIsBegin(Trace(i), cllBeginEntry) Then
            bConsistent = False
            For j = i + 1 To Trace.Count
                If NtryIsEnd(Trace(j), cllEndEntry) Then
                    If ItmId(cllBeginEntry) = ItmId(cllEndEntry) Then
                        If ItmLvl(cllBeginEntry) = ItmLvl(cllEndEntry) Then
                            '~~ Calculate the executesd seconds for the end entry
                            NtryTcksGrss(cllEndEntry) = NtryTcksElpsd(cllEndEntry) - NtryTcksElpsd(cllBeginEntry)
                            NtryTcksOvrhdItm(cllEndEntry) = NtryTcksOvrhdNtry(cllBeginEntry) + NtryTcksOvrhdNtry(cllEndEntry)
                            GoTo next_begin_entry
                        End If
                    End If
                End If
            Next j
            '~~ No corresponding end entry found
            Select Case ItmDrctv(cllBeginEntry)
                Case DIR_BEGIN_PROC: sComment = "No corresponding End of Procedure (EoP) code line in:    "
                Case DIR_BEGIN_CODE: sComment = "No corresponding End of CodeTrace (EoC) code line in:    "
            End Select
            If Not dct.Exists(ItmId(cllBeginEntry)) Then dct.Add ItmId(cllBeginEntry), sComment
        End If

next_begin_entry:
    Next i
    
    '~~ Check for missing corresponding begin entries (if the end entry has no executed ticks)
    For Each v In Trace
        If NtryIsEnd(v, cllEndEntry) Then
            If NtryTcksGrss(cllEndEntry) = 0 Then
                '~~ No corresponding begin entry found
                Select Case ItmDrctv(cllEndEntry)
                    Case DIR_END_PROC: sComment = "Corresponding Begin of Procedure (BoP) code line missing in: "
                    Case DIR_END_CODE: sComment = "Corresponding Begin of CodeTrace (BoC) code line missing in: "
                End Select
                If Not dct.Exists(ItmId(cllEndEntry)) Then dct.Add ItmId(cllEndEntry), sComment
            End If
        End If
    Next v
    
    If dct.Count > 0 Then
        For Each v In dct
            sTrace = sTrace & dct(v) & v & vbLf
        Next v
        With fMsg
            .MsgTitle = "Inconsistent begin/end trace code lines!"
            
            SctnLabel.Text = "Due to the following inconsistencies the display of the trace result became useless/impossible:"
            SctnText.Text = sTrace:   SctnText.MonoSpaced = True
            .MsgLabel(1) = SctnLabel
            .MsgText(1) = SctnText
            
            .Setup
            .Show
        End With
    Else
        DsplyNtryAllCnsstnt = True
    End If

xt: Exit Function

eh: Select Case ErrMsg(ErrSrc(PROC))
        Case vbYes: Stop: Resume
        Case vbNo:  Stop: Resume Next
        Case Else:  GoTo xt
    End Select
End Function

Private Function DsplyTcksDffToScs(ByVal BeginTicks As Currency, _
                                   ByVal endticks As Currency) As Currency
' ------------------------------------------------------------------------
' Returns the difference between begin- and endticks as seconds.
' ------------------------------------------------------------------------
    IIf endticks - BeginTicks > 0, DsplyTcksDffToScs = (endticks - BeginTicks) / cySysFrequency, DsplyTcksDffToScs = 0
End Function

Private Function DsplyValue( _
                      ByVal dv_entry As Collection, _
                      ByVal dv_value As Variant, _
                      ByVal dv_frmt As String) As String
    If NtryIsBegin(dv_entry) And dv_value = 0 _
    Then DsplyValue = Space$(Len(dv_frmt)) _
    Else DsplyValue = IIf(dv_value >= 0, Format$(dv_value, dv_frmt), Space$(Len(dv_frmt)))
End Function

Private Property Get DsplyValueFormat(Optional ByVal forvalue As Variant = 999.999) As String
    Select Case CLng(forvalue)
        Case Is < 10:       DsplyValueFormat = "0" & "." & String$(DisplayedPrecision, "0")
        Case Is < 100:      DsplyValueFormat = "00" & "." & String$(DisplayedPrecision, "0")
        Case Is < 1000:     DsplyValueFormat = "000" & "." & String$(DisplayedPrecision, "0")
        Case Is < 10000:    DsplyValueFormat = "0000" & "." & String$(DisplayedPrecision, "0")
    End Select
End Property

Private Sub DsplyValuesFormatSet()
    Dim cllLast As Collection:  Set cllLast = NtryLst
    sFrmtTcksSys = DsplyValueFormat(ItmTcksSys(cllLast))
    sFrmtTcksElpsd = DsplyValueFormat(NtryTcksElpsd(cllLast))
    sFrmtTcksGrss = DsplyValueFormat(NtryTcksGrss(cllLast))
    sFrmtTcksOvrhdNtry = DsplyValueFormat(NtryTcksOvrhdNtry(cllLast))
    sFrmtTcksOvrhdItm = DsplyValueFormat(NtryTcksOvrhdItmMax)
    sFrmtTcksNt = DsplyValueFormat(NtryTcksNt(cllLast))
    sFrmtScsElpsd = DsplyValueFormat(NtryScsElpsd(cllLast))
    sFrmtScsGrss = DsplyValueFormat(NtryScsGrss(cllLast))
    sFrmtScsOvrhdNtry = DsplyValueFormat(NtryScsOvrhdNtry(cllLast))
    sFrmtScsOvrhdItm = DsplyValueFormat(NtryScsOvrhdItm(cllLast))
    sFrmtScsNt = DsplyValueFormat(NtryScsNt(cllLast))
End Sub

Public Sub EoC(ByVal eoc_id As String, _
      Optional ByVal eoc_inf As String = vbNullString)
' ----------------------------------------------------
' End of the trace of a code sequence.
' ----------------------------------------------------
#If ExecTrace = 1 Then
    Dim cll As Collection
    
    cyTcksOvrhdTrcStrt = SysCrrntTcks
    If StckIsEmpty(TraceStack) Then Exit Sub
    If Trace Is Nothing Then Exit Sub
    TrcEnd trc_id:=eoc_id, trc_dir:=DIR_END_CODE, trc_inf:=eoc_inf, trc_cll:=cll
    cyTcksOvrhdTrc = SysCrrntTcks - cyTcksOvrhdTrcStrt ' overhead ticks caused by the collection of the begin trace entry

#End If
End Sub

Public Sub EoP(ByVal eop_id As String, _
      Optional ByVal eop_inf As String = vbNullString)
' ----------------------------------------------------
' End of the trace of a procedure.
' ----------------------------------------------------
#If ExecTrace = 1 Then
    Dim cll As Collection
    
    cyTcksOvrhdTrcStrt = SysCrrntTcks
    If StckIsEmpty(TraceStack) Then Exit Sub        ' Nothing to trace any longer. Stack has been emptied after an error to finish the trace
    If Trace Is Nothing Then Exit Sub  ' No trace or trace has finished
    
    TrcEnd trc_id:=eop_id, trc_dir:=DIR_END_PROC, trc_inf:=eop_inf, trc_cll:=cll
    If StckIsEmpty(TraceStack) Then
        Dsply
    End If
    cyTcksOvrhdTrc = SysCrrntTcks - cyTcksOvrhdTrcStrt ' overhead ticks caused by the collection of the end-of-trace entry
#End If
End Sub

'Private Sub ErrMsgMatter(ByVal err_source As String, _
'                         ByVal err_no As Long, _
'                         ByVal err_line As Long, _
'                         ByVal err_dscrptn As String, _
'                Optional ByRef msg_title As String, _
'                Optional ByRef msg_type As String, _
'                Optional ByRef msg_line As String, _
'                Optional ByRef msg_no As Long, _
'                Optional ByRef msg_details As String, _
'                Optional ByRef msg_dscrptn As String, _
'                Optional ByRef msg_info As String)
'' -------------------------------------------------------------
'' Returns all matter to build a proper error message.
'' msg_line:    at line <eline>
'' msg_no:      1 to n
'' msg_title:   <etype> <enumber> in <esource> [at line <eline>]
'' msg_details: (at line <eline>)
'' msg_dscrptn: the error description
'' msg_info:    any text which follows the description
''              concatenated by a ||
'' -------------------------------------------------------------
'    If InStr(1, err_source, "DAO") <> 0 _
'    Or InStr(1, err_source, "ODBC Teradata Driver") <> 0 _
'    Or InStr(1, err_source, "ODBC") <> 0 _
'    Or InStr(1, err_source, "Oracle") <> 0 Then
'        msg_type = "Database Error "
'    Else
'      msg_type = IIf(err_no > 0, "VB-Runtime Error ", "Application Error ")
'    End If
'
'    msg_line = IIf(err_line <> 0, "at line " & err_line, vbNullString)     ' Message error line
'    msg_no = IIf(err_no < 0, err_no - vbObjectError, err_no)                ' Message error number
'    msg_title = msg_type & msg_no & " in " & err_source & " " & msg_line             ' Message title
'    msg_details = IIf(err_line <> 0, msg_type & msg_no & " in " & err_source & " (at line " & err_line & ")", msg_type & msg_no & " in " & err_source)
'    msg_dscrptn = IIf(InStr(err_dscrptn, CONCAT) <> 0, Split(err_dscrptn, CONCAT)(0), err_dscrptn)
'    If InStr(err_dscrptn, CONCAT) <> 0 Then msg_info = Split(err_dscrptn, CONCAT)(1)
'
'End Sub

Private Function ErrMsg(ByVal err_source As String, _
               Optional ByVal err_no As Long = 0, _
               Optional ByVal err_dscrptn As String = vbNullString, _
               Optional ByVal err_line As Long = 0) As Variant
' ------------------------------------------------------------------------------
' Universal error message display service including a debugging option
' (Conditional Compile Argument 'Debugging = 1') and an optional additional
' "about the error" information which may be connected to an error message by
' two vertical bars (||).
'
' A copy of this function is used in each procedure with an error handling
' (On error Goto eh).
'
' The function considers the Common VBA Error Handling Component (ErH) which
' may be installed (Conditional Compile Argument 'ErHComp = 1') and/or the
' Common VBA Message Display Component (mMsg) installed (Conditional Compile
' Argument 'MsgComp = 1'). Only when none of the two is installed the error
' message is displayed by means of the VBA.MsgBox.
'
' Usage: Example with the Conditional Compile Argument 'Debugging = 1'
'
'        Private/Public <procedure-name>
'            Const PROC = "<procedure-name>"
'
'            On Error Goto eh
'            ....
'        xt: Exit Sub/Function/Property
'
'        eh: Select Case ErrMsg(ErrSrc(PROC))
'               Case vbResume:  Stop: Resume
'               Case Else:      GoTo xt
'            End Select
'        End Sub/Function/Property
'
'        The above may appear a lot of code lines but will be a godsend in case
'        of an error!
'
' Uses:  - For programmed application errors (Err.Raise AppErr(n), ....) the
'          function AppErr will be used which turns the positive number into a
'          negative one. The error message will regard a negative error number
'          as an 'Application Error' and will use AppErr to turn it back for
'          the message into its original positive number. Together with the
'          ErrSrc there will be no need to maintain numerous different error
'          numbers for a VB-Project.
'        - The caller provides the source of the error through the module
'          specific function ErrSrc(PROC) which adds the module name to the
'          procedure name.
'
' W. Rauschenberger Berlin, Nov 2021
' ------------------------------------------------------------------------------
#If ErHComp = 1 Then
    '~~ ------------------------------------------------------------------------
    '~~ When the Common VBA Error Handling Component (mErH) is installed in the
    '~~ VB-Project (which includes the mMsg component) the mErh.ErrMsg service
    '~~ is preferred since it provides some enhanced features like a path to the
    '~~ error.
    '~~ ------------------------------------------------------------------------
    ErrMsg = mErH.ErrMsg(err_source, err_no, err_dscrptn, err_line)
    GoTo xt
#ElseIf MsgComp = 1 Then
    '~~ ------------------------------------------------------------------------
    '~~ When only the Common Message Services Component (mMsg) is installed but
    '~~ not the mErH component the mMsg.ErrMsg service is preferred since it
    '~~ provides an enhanced layout and other features.
    '~~ ------------------------------------------------------------------------
    ErrMsg = mMsg.ErrMsg(err_source, err_no, err_dscrptn, err_line)
    GoTo xt
#End If
    '~~ -------------------------------------------------------------------
    '~~ When neither the mMsg nor the mErH component is installed the error
    '~~ message is displayed by means of the VBA.MsgBox
    '~~ -------------------------------------------------------------------
    Dim ErrBttns    As Variant
    Dim ErrAtLine   As String
    Dim ErrDesc     As String
    Dim ErrLine     As Long
    Dim ErrNo       As Long
    Dim ErrSrc      As String
    Dim ErrText     As String
    Dim ErrTitle    As String
    Dim ErrType     As String
    Dim ErrAbout    As String
        
    '~~ Obtain error information from the Err object for any argument not provided
    If err_no = 0 Then err_no = Err.Number
    If err_line = 0 Then ErrLine = Erl
    If err_source = vbNullString Then err_source = Err.Source
    If err_dscrptn = vbNullString Then err_dscrptn = Err.Description
    If err_dscrptn = vbNullString Then err_dscrptn = "--- No error description available ---"
    
    If InStr(err_dscrptn, "||") <> 0 Then
        ErrDesc = Split(err_dscrptn, "||")(0)
        ErrAbout = Split(err_dscrptn, "||")(1)
    Else
        ErrDesc = err_dscrptn
    End If
    
    '~~ Determine the type of error
    Select Case err_no
        Case Is < 0
            ErrNo = AppErr(err_no)
            ErrType = "Application Error "
        Case Else
            ErrNo = err_no
            If (InStr(1, err_dscrptn, "DAO") <> 0 _
            Or InStr(1, err_dscrptn, "ODBC Teradata Driver") <> 0 _
            Or InStr(1, err_dscrptn, "ODBC") <> 0 _
            Or InStr(1, err_dscrptn, "Oracle") <> 0) _
            Then ErrType = "Database Error " _
            Else ErrType = "VB Runtime Error "
    End Select
    
    If err_source <> vbNullString Then ErrSrc = " in: """ & err_source & """"   ' assemble ErrSrc from available information"
    If err_line <> 0 Then ErrAtLine = " at line " & err_line                    ' assemble ErrAtLine from available information
    ErrTitle = Replace(ErrType & ErrNo & ErrSrc & ErrAtLine, "  ", " ")         ' assemble ErrTitle from available information
       
    ErrText = "Error: " & vbLf & _
              ErrDesc & vbLf & vbLf & _
              "Source: " & vbLf & _
              err_source & ErrAtLine
    If ErrAbout <> vbNullString _
    Then ErrText = ErrText & vbLf & vbLf & _
                  "About: " & vbLf & _
                  ErrAbout
    
#If Debugging Then
    ErrBttns = vbYesNo
    ErrText = ErrText & vbLf & vbLf & _
              "Debugging:" & vbLf & _
              "Yes    = Resume Error Line" & vbLf & _
              "No     = Terminate"
#Else
    ErrBttns = vbCritical
#End If
    
    ErrMsg = MsgBox(Title:=ErrTitle _
                  , Prompt:=ErrText _
                  , Buttons:=ErrBttns)
xt: Exit Function

End Function

Private Function ErrSrc(ByVal sProc As String) As String
    ErrSrc = "mTrc." & sProc
End Function

Public Sub Finish(Optional ByRef inf As String = vbNullString)
' ------------------------------------------------------------
' Finishes an unfinished traced items by means of the stack.
' All items on the the stack are processed via EoP/EoC.
' ------------------------------------------------------------
    Dim cll As Collection
    
    While Not StckIsEmpty(TraceStack)
        Set cll = StckTop(TraceStack)
        If NtryIsCode(cll) _
        Then mTrc.EoC eoc_id:=ItmId(cll), eoc_inf:=inf _
        Else mTrc.EoP eop_id:=ItmId(cll), eop_inf:=inf
        inf = vbNullString
    Wend
    
End Sub

Private Sub Initialize()
' -------------------------------------------------------
' Public allows an error handler to initialize the trace.
' -------------------------------------------------------
    Set Trace = New Collection
    Set cllNtryLast = Nothing
    dtTraceBegin = Now()
    cyTcksOvrhdItm = 0
    iTrcLvl = 0
    cySysFrequency = 0
    sFirstTraceItem = vbNullString
    If lPrecision = 0 Then lPrecision = 6 ' default when another valus has not been set by the caller
    
End Sub

Private Function itm( _
               ByVal itm_drctv As String, _
               ByVal itm_id As String, _
               ByVal itm_inf As String, _
               ByVal itm_lvl As Long, _
               ByVal itm_tckssys As Currency, _
               ByVal itm_args As Variant) As Variant()
' -------------------------------------------------------
' Returns an array with the arguments ordered by their
' enumerated position.
' -------------------------------------------------------
    Dim av(1 To 6) As Variant
    
    av(enItmDrctv) = itm_drctv
    av(enItmId) = itm_id
    av(enItmInf) = itm_inf
    av(enItmLvl) = itm_lvl
    av(enItmTcksSys) = itm_tckssys
    av(enPosItmArgs) = itm_args
    itm = av
    
End Function

Private Function Max(ParamArray va() As Variant) As Variant
' ---------------------------------------------------------
' Returns the maximum value of all values provided (va).
' ---------------------------------------------------------
    Dim v   As Variant
    
    Max = va(LBound(va)): If LBound(va) = UBound(va) Then Exit Function
    For Each v In va
        If v > Max Then Max = v
    Next v
    
End Function

Private Function Ntry(ByVal ntry_tcks As Currency, _
                      ByVal ntry_dir As String, _
                      ByVal ntry_id As String, _
                      ByVal ntry_lvl As Long, _
                      ByVal ntry_inf As String, _
                      ByVal ntry_args As Variant) As Collection
' ------------------------------------------------------
' Return the arguments as elements in an array as an
' item in a collection.
' ------------------------------------------------------
    Const PROC = "Ntry"
    
    On Error GoTo eh
    Dim cll As New Collection
    Dim VarItm  As Variant
    
    VarItm = itm(itm_drctv:=ntry_dir, itm_id:=ntry_id, itm_inf:=ntry_inf, itm_lvl:=ntry_lvl, itm_tckssys:=ntry_tcks, itm_args:=ntry_args)
    NtryItm(cll) = VarItm
'    NtryTestDsply ntry_tcks:=ntry_tcks, ntry_dir:=ntry_dir, ntry_id:=ntry_id, ntry_lvl:=ntry_lvl, ntry_inf:=ntry_inf
    Set Ntry = cll
    
xt: Exit Function
    
eh: Select Case ErrMsg(ErrSrc(PROC))
        Case vbResume:  Stop: Resume
        Case Else:      GoTo xt
    End Select
End Function

Private Function NtryIsBegin(ByVal v As Collection, _
                     Optional ByRef cll As Collection = Nothing) As Boolean
' -------------------------------------------------------------------------
' Returns TRUE and v as cll when the entry is a begin entry, else FALSE and
' cll = Nothing
' -------------------------------------------------------------------------
    If InStr(ItmDrctv(v), DIR_BEGIN_ID) <> 0 Then
        NtryIsBegin = True
        Set cll = v
    End If
End Function

Private Function NtryIsCode(ByVal cll As Collection) As Boolean
    
    Select Case ItmDrctv(cll)
        Case DIR_BEGIN_CODE, DIR_END_CODE: NtryIsCode = True
    End Select

End Function

Private Function NtryIsEnd(ByVal v As Collection, _
                           ByRef cll As Collection) As Boolean
' ------------------------------------------------------------
' Returns TRUE and v as cll when the entry is an end entry,
' else FALSE and cll = Nothing
' ------------------------------------------------------------
    If InStr(ItmDrctv(v), DIR_END_ID) <> 0 Then
        NtryIsEnd = True
        Set cll = v
    End If
End Function

Private Function NtryLastBegin() As Long
    
    Dim i As Long
    
    For i = Trace.Count To 1 Step -1
        If NtryIsBegin(Trace(i)) Then
            NtryLastBegin = i
            Exit Function
        End If
    Next i
    
End Function

Private Function NtryLst() As Collection ' Retun last entry
    Set NtryLst = Trace(Trace.Count)
End Function

Private Function NtryTcksOvrhdItmMax() As Double
    
    Dim cll As Collection
    Dim v   As Variant
    
    For Each v In Trace
        Set cll = v
        NtryTcksOvrhdItmMax = Max(NtryTcksOvrhdItmMax, NtryTcksOvrhdItm(cll))
    Next v

End Function

Private Sub NtryTestDsply( _
                    ByVal ntry_tcks As Currency, _
                    ByVal ntry_dir As String, _
                    ByVal ntry_id As String, _
                    ByVal ntry_lvl As Long, _
                    ByVal ntry_inf As String)
                    
    If TraceStack Is Nothing Then
        Debug.Print ntry_tcks, ntry_lvl, "(1)", ntry_dir, ntry_id, ntry_inf
    Else
        Debug.Print ntry_tcks, ntry_lvl, "(" & TraceStack.Count & ")", ntry_dir, ntry_id, ntry_inf
    End If
End Sub

Public Sub Pause()
    cyTcksPauseStart = SysCrrntTcks
End Sub

Private Function RepeatStrng( _
                       ByVal rs_s As String, _
                       ByVal rs_n As Long) As String
' --------------------------------------------------
' Returns the string (s) concatenated (n) times.
' !! VBA.String in not an alternative because it  !!
' !! it not supports leading andr trailing spaces !!
' --------------------------------------------------
    Dim i   As Long
    For i = 1 To rs_n: RepeatStrng = RepeatStrng & rs_s:  Next i
End Function

Private Sub StckAdjust(ByVal trc_id As String)
    Dim cllNtry As Collection
    Dim i       As Long
    
    For i = TraceStack.Count To 1 Step -1
        Set cllNtry = TraceStack(i)
        If ItmId(cllNtry) = trc_id Then
            Exit For
        Else
            TraceStack.Remove (TraceStack.Count)
            iTrcLvl = iTrcLvl - 1
        End If
    Next i

End Sub

Private Function StckEd(ByVal stck_id As String, _
                        ByVal stck_lvl As Long) As Boolean
' --------------------------------------------------------
' Returns TRUE when last item pushed to the stack is
' identical with the item (stck_id) and level (stck_lvl).
' --------------------------------------------------------
    Const PROC = "StckEd"
    
    On Error GoTo eh
    Dim v       As Variant
    Dim cllNtry As Collection
    Dim i       As Long
    
    For i = TraceStack.Count To 1 Step -1
        Set cllNtry = TraceStack(i)
        If ItmId(cllNtry) = stck_id Then ' And ItmLvl(cllNtry) = stck_lvl Then
            StckEd = True
            Exit Function
        End If
    Next i

xt: Exit Function

eh: Select Case ErrMsg(ErrSrc(PROC))
        Case vbResume:  Stop: Resume
        Case Else:      GoTo xt
    End Select
End Function

Private Function StckIsEmpty(ByVal stck As Collection) As Boolean
    StckIsEmpty = stck Is Nothing
    If Not StckIsEmpty Then StckIsEmpty = stck.Count = 0
End Function

Private Sub StckPop(ByRef stck As Collection, _
                    ByVal stck_item As Variant, _
           Optional ByRef stck_ppd As Collection)
' ----------------------------------------------------------------------------
'
' ----------------------------------------------------------------------------
    Const PROC = "StckPop"
    
    On Error GoTo eh
    Dim cllTop  As Collection: Set cllTop = StckTop(stck)
    Dim cll     As Collection: Set cll = stck_item
    
    While ItmId(cll) <> ItmId(cllTop) And Not StckIsEmpty(TraceStack)
        '~~ Finish any unfinished code trace still on the stack which needs to be finished first
        If NtryIsCode(cllTop) Then
            mTrc.EoC eoc_id:=ItmId(cllTop), eoc_inf:="ended by stack!!"
        Else
            mTrc.EoP eop_id:=ItmId(cllTop), eop_inf:="ended by stack!!"
        End If
        If Not StckIsEmpty(TraceStack) Then Set cllTop = StckTop(stck)
    Wend
    
    If StckIsEmpty(TraceStack) Then GoTo xt
    
    If ItmId(cll) = ItmId(cll) Then
        Set stck_ppd = cllTop
        TraceStack.Remove TraceStack.Count
        Set cllTop = StckTop(TraceStack)
    Else
        Debug.Print "Stack Pop ='" & ItmId(cll) _
                  & "', Stack Top = '" & ItmId(cllTop) _
                  & "', Stack Dir = '" & ItmDrctv(cllTop) _
                  & "', Stack Lvl = '" & ItmLvl(cllTop) _
                  & "', Stack Cnt = '" & TraceStack.Count
        Stop
    End If

xt: Exit Sub

eh: Select Case ErrMsg(ErrSrc(PROC))
        Case vbResume:  Stop: Resume
        Case Else:      GoTo xt
    End Select
End Sub

Private Sub StckPush(ByRef stck As Collection, _
                     ByVal stck_item As Variant)
    If stck Is Nothing Then Set stck = New Collection
    stck.Add stck_item
End Sub

Private Function StckTop(ByVal stck As Collection) As Collection
    If Not StckIsEmpty(stck) _
    Then Set StckTop = stck(stck.Count)
End Function

Public Sub Terminate()
' -----------------------------------------------------------------
' Should be called by any error handling when a new execution trace
' is about to begin with the very first procedure's execution.
' -----------------------------------------------------------------
    Set Trace = Nothing
    Set TraceStack = Nothing
    cyTcksPaused = 0
End Sub

Private Sub TraceLogBgn(ByVal tl_ticks As Currency, _
                        ByVal tl_dir As String, _
                        ByVal tl_id As String)
' ----------------------------------------------------------------------------
' Write an begin trace line to the trace log file (sTraceLogFile) - provided
' one had been specified. When the TraceStack contains the very first item
' write an additional Begion trace header line.
' ----------------------------------------------------------------------------
    Dim LogText As String
    
    sFrmtScsElpsd = DsplyValueFormat(99.99999)
    
    If Not sTraceLogFile = vbNullString Then
        If TraceStack.Count = 1 Then
            TraceLogTxt(sTraceLogFile) = vbNullString ' empty line in new file
            LogText = String(Len(sFrmtScsElpsd) + 2, " ") & tl_dir & Format(Now(), " hh:mm:ss") & " Begin execution trace "
            TraceLogTxt(sTraceLogFile) = LogText
        End If
        LogText = String(Len(sFrmtScsElpsd) + 2, " ") & RepeatStrng("|  ", TraceStack.Count) & tl_dir & " " & ThisWorkbook.Name & " " & tl_id
        TraceLogTxt(sTraceLogFile) = LogText
    End If
End Sub

Private Sub TraceLogEnd(ByVal tl_ticks_begin As Currency, _
                        ByVal tl_ticks_end As Currency, _
                        ByVal tl_dir As String, _
                        ByVal tl_id As String)
' ----------------------------------------------------------------------------
' Write an end trace line to the trace log file (sTraceLogFile) - provided
' one had been specified - with the execution time calculated in seconds.
' When the TraceStack is empty write an additional End trace footer line.
' ----------------------------------------------------------------------------
    Dim LogText         As String
    Dim TicksElapsed    As Currency: TicksElapsed = tl_ticks_end - tl_ticks_begin
    Dim SecsElapsed     As Currency: SecsElapsed = CDec(TicksElapsed) / CDec(SysFrequency)
    
    sFrmtScsElpsd = DsplyValueFormat(99.99999)
  
    If Not sTraceLogFile = vbNullString Then
        LogText = " " & Format(SecsElapsed, sFrmtScsElpsd) & " " & RepeatStrng("|  ", TraceStack.Count + 1) & tl_dir & " " & ThisWorkbook.Name & " " & tl_id
        TraceLogTxt(sTraceLogFile) = LogText
        If TraceStack.Count = 0 Then
            LogText = String(Len(sFrmtScsElpsd) + 2, " ") & tl_dir & Format(Now(), " hh:mm:ss") & " End execution trace "
            TraceLogTxt(sTraceLogFile) = LogText
        End If
    End If
End Sub

Public Property Let TraceLogInfo(ByVal tl_inf As String)
' ----------------------------------------------------------------------------
' Write an info line (tl_inf) to the trace log file (sTraceLogFile) provided
' one had been specified for this execution trace
' ----------------------------------------------------------------------------
Dim LogText As String
    If Not sTraceLogFile = vbNullString And TraceStack.Count > 1 Then
        LogText = String(Len(sFrmtScsElpsd) + 2, " ") & RepeatStrng("|  ", TraceStack.Count) & "|  " & ThisWorkbook.Name & " " & tl_inf
        TraceLogTxt(tl_file:=sTraceLogFile, tl_append:=True) = LogText
    End If
End Property

Private Sub TrcAdd(ByVal trc_id As String, _
                   ByVal trc_tcks As Currency, _
                   ByVal trc_dir As String, _
                   ByVal trc_lvl As Long, _
          Optional ByVal trc_args As Variant, _
          Optional ByVal trc_inf As String = vbNullString, _
          Optional ByRef trc_ntry As Collection)
' ----------------------------------------------------
' Adds an entry to the trace collection.
' ----------------------------------------------------
    Const PROC = "TrcAdd"
    
    Static sLastDrctv   As String
    Static sLastId      As String
    Static lLastLvl     As String
    
    On Error GoTo eh
    Dim bAlreadyAdded   As Boolean
    
    If Not cllNtryLast Is Nothing Then
        '~~ When this is not the first entry added the overhead ticks caused by the previous entry is saved.
        '~~ Saving it with the next entry avoids a wrong overhead when saved with the entry itself because.
        '~~ Its maybe nitpicking but worth the try to get execution time figures as correct/exact as possible.
        If sLastId = trc_id And lLastLvl = trc_lvl And sLastDrctv = trc_dir Then bAlreadyAdded = True
        If Not bAlreadyAdded Then
            NtryTcksOvrhdNtry(cllNtryLast) = cyTcksOvrhdTrc
        Else
            Debug.Print ItmId(cllNtryLast) & " already added"
        End If
    End If
    
    If Not bAlreadyAdded Then
        Set trc_ntry = Ntry(ntry_tcks:=trc_tcks, ntry_dir:=trc_dir, ntry_id:=trc_id, ntry_lvl:=trc_lvl, ntry_inf:=trc_inf, ntry_args:=trc_args)
        Trace.Add trc_ntry
        Set cllNtryLast = trc_ntry
        sLastDrctv = trc_dir
        sLastId = trc_id
        lLastLvl = trc_lvl
    Else
        Debug.Print ItmId(cllNtryLast) & " already added"
    End If

xt: Exit Sub

eh: Select Case ErrMsg(ErrSrc(PROC))
        Case vbResume:  Stop: Resume
        Case Else:      GoTo xt
    End Select
End Sub

Private Sub TrcBgn(ByVal trc_id As String, _
                   ByVal trc_dir As String, _
          Optional ByVal trc_args As Variant, _
          Optional ByRef trc_cll As Collection)
' ----------------------------------------------------------------------------
' Collect a trace begin entry with the current ticks count for the procedure
' or code (item).
' ----------------------------------------------------------------------------
    Const PROC = "TrcEnd"
    
    On Error GoTo eh
    Dim cy  As Currency:    cy = SysCrrntTcks - cyTcksPaused
    Dim i As Long
           
    iTrcLvl = iTrcLvl + 1
    TrcAdd trc_id:=trc_id _
         , trc_tcks:=cy _
         , trc_dir:=trc_dir _
         , trc_lvl:=iTrcLvl _
         , trc_inf:=vbNullString _
         , trc_args:=trc_args _
         , trc_ntry:=trc_cll
    StckPush TraceStack, trc_cll
    TraceLogBgn tl_ticks:=cy, tl_dir:=trc_dir, tl_id:=trc_id

xt: Exit Sub
    
eh: Select Case ErrMsg(ErrSrc(PROC))
        Case vbResume:  Stop: Resume
        Case Else:      GoTo xt
    End Select
End Sub

Private Sub TrcEnd(ByVal trc_id As String, _
          Optional ByVal trc_dir As String = vbNullString, _
          Optional ByVal trc_inf As String = vbNullString, _
          Optional ByRef trc_cll As Collection)
' ------------------------------------------------------
' Collect an end trace entry with the current ticks
' count for the procedure or code (item).
' ------------------------------------------------------
    Const PROC = "TrcEnd"
    
    On Error GoTo eh
    Dim cy  As Currency:    cy = SysCrrntTcks - cyTcksPaused
    Dim Top As Collection:  Set Top = StckTop(TraceStack)
    Dim itm As Collection
    
    If trc_inf <> vbNullString Then
        trc_inf = TRC_INFO_DELIM & trc_inf & TRC_INFO_DELIM
    End If
    
    '~~ Any end trace for an item not on the stack is ignored. On the other hand,
    '~~ if on the stack but not the last item the stack is adjusted because this
    '~~ indicates a begin without a corresponding end trace statement.
    If Not StckEd(stck_id:=trc_id, stck_lvl:=iTrcLvl) Then
        Exit Sub
    Else
        StckAdjust trc_id
    End If
    
    If ItmId(Top) <> trc_id And ItmLvl(Top) = iTrcLvl Then
        StckPop TraceStack, Top
    End If

    TrcAdd trc_id:=trc_id _
         , trc_tcks:=cy _
         , trc_dir:=trc_dir _
         , trc_lvl:=iTrcLvl _
         , trc_inf:=vbNullString _
         , trc_ntry:=trc_cll
         
    StckPop stck:=TraceStack, stck_item:=trc_cll, stck_ppd:=itm
    TraceLogEnd tl_ticks_begin:=ItmTcksSys(itm), tl_ticks_end:=cy, tl_dir:=trc_dir, tl_id:=trc_id
    
    iTrcLvl = iTrcLvl - 1

xt: Exit Sub
    
eh: Select Case ErrMsg(ErrSrc(PROC))
        Case vbResume:  Stop: Resume
        Case Else:      GoTo xt
    End Select
End Sub

Private Function TrcIsEmpty() As Boolean
    TrcIsEmpty = Trace Is Nothing
    If Not TrcIsEmpty Then TrcIsEmpty = Trace.Count = 0
End Function

Private Function TrcLast() As Collection
    If Trace.Count <> 0 _
    Then Set TrcLast = Trace(Trace.Count)
End Function

