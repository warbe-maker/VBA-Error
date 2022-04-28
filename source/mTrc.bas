Attribute VB_Name = "mTrc"
Option Explicit
' ----------------------------------------------------------------------------
' Standard Module  m T r c :
' Services to trace the execution  of procedures and code snippets. The
' elapsed execution time of traced items comes with the highest possible
' precision. The trace result is written to a log file which ensures at least
' a partial trace when the execution exceptionally terminates.
'
' When this module is installed the sevices are triggered/activated by the
' Conditional Compile Argument 'ExecTrace = 1'. When the module is installed
' and the Conditional Compile Argument is turned to 'ExecTrace = 0' all
' services are disabled thus avoiding any effect on the performance - which is
' already very little when the services are active.
'
' Public services:
' - BoC             Indicates the (B)egin (o)f the execution trace of a (C)ode
'                   snippet.
' - BoP             Indicates the (B)egin (o)f the execution trace of a
'                   (P)rocedure.
' - BoP_ErH         Exclusively used by the mErH module.
' - Continue        Commands the execution trace to continue taking the
'                   execution time when it had been paused. Pause and Continue
'                   is used by the mErH module for example to avoid useless
'                   execution time taking while waiting for the users reply.
' - Dsply           Displays the content of the trace log file. Available only
'                   when the mMsg/fMsg modules are installed and this is
'                   indicated by the Conditional Compile Argument
'                   'MsgComp = 1'. Without mMsg/fMsg the trace result log
'                   will be viewed with any appropriate text file viewer.
' - EoC             Indicates the (E)nd (o)f the execution trace of a (C)ode
'                   snippet.
' - EoP             Indicates the (E)nd (o)f the execution trace of a
'                   (P)rocedure.
' - Pause           Stops the execution traces time taking, e.g. while an
'                   error  message is displayed.
' - LogFile         Provides the full name of a desired trace log file which
'                   defaults to "ExecTrace.log" in ThisWorkbook's parent
'                   folder.
' - LogInfo         Explicitely writes an entry to the trace lof file by
'                   considering the nesting level (i.e. the indentation).
'
' Optionally may use:
' - mMsg/fMsg 1)    To enable the Dsply service for which the VBA.MsgBox
'                   is inappropriate, also displays more comprehensive
'                   error messages.
' - mErH 2)         To display proper error messages by providing additional
'                   information such like the 'path to the error'. 2)
'
' Requires:         Reference to 'Microsoft Scripting Runtime'
'
' See: https://github.com/warbe-maker/Common-VBA-Execution-Trace-Service
'
' 1) See https://github.com/warbe-maker/Common-VBA-Message-Service for how to
'    install an use.
' 2) See https://github.com/warbe-maker/Common-VBA-Error-Services for how to
'    install an use.
'
' W. Rauschenberger, Berlin, Feb. 2022
' ----------------------------------------------------------------------------
#If Not MsgComp = 1 Then
    ' ------------------------------------------------------------------------
    ' The 'minimum error handling' aproach implemented with this module and
    ' provided by the ErrMsg function uses the VBA.MsgBox to display an error
    ' message which includes a debugging option to resume the error line
    ' provided the Conditional Compile Argument 'Debugging = 1'.
    ' This declaration allows the mTrc module to work completely autonomous.
    ' It becomes obsolete when the mMsg/fMsg module is installed 1) which must
    ' be indicated by the Conditional Compile Argument MsgComp = 1
    '
    ' 1) See https://github.com/warbe-maker/Common-VBA-Message-Service for
    '    how to install an use.
    ' ------------------------------------------------------------------------
    Private Const vbResumeOk As Long = 7 ' Buttons value in mMsg.ErrMsg (pass on not supported)
    Private Const vbResume   As Long = 6 ' return value (equates to vbYes)
#End If

Public Enum enDsplydInfo
    Detailed = 1
    Compact = 2
End Enum

Private Enum enTraceInfo
    enItmDir = 1
    enItmId
    enItmInf
    enItmLvl
    enItmTcks
    enPosItmArgs
End Enum

Private Declare PtrSafe Function getFrequency Lib "kernel32" _
Alias "QueryPerformanceFrequency" (cySysFrequency As Currency) As Long
Private Declare PtrSafe Function getTickCount Lib "kernel32" _
Alias "QueryPerformanceCounter" (cyTickCount As Currency) As Long

Private Const DIR_BEGIN_ID      As String = ">"     ' Begin procedure or code trace indicator
Private Const DIR_END_ID        As String = "<"     ' End procedure or code trace indicator
Private Const TRC_INFO_DELIM    As String = " !!! "
Private Const TRC_LOG_SEC_FRMT  As String = "00.0000 "

Private cyTcksAtStart       As Currency         ' Trace log to file
Private LastNtry            As Collection       '
Private cySysFrequency      As Currency         ' Execution Trace SysFrequency (initialized with init)
Private cyTcksOvrhdItm      As Currency         ' Execution Trace time accumulated by caused by the time tracking itself
Private cyTcksOvrhdTrc      As Currency         ' Overhead ticks caused by the collection of a traced item's entry
Private cyTcksOvrhdTrcStrt  As Currency         ' Overhead ticks caused by the collection of a traced item's entry
Private cyTcksPaused        As Currency         ' Accumulated with procedure Continue
Private cyTcksPauseStart    As Currency         ' Set with procedure Pause
Private dtTraceBegin        As Date             ' Initialized at start of execution trace
Private iTrcLvl             As Long             ' Increased with each begin entry and decreased with each end entry
Private sFirstTraceItem     As String
Private sLogFile            As String           ' When not vbNullString the trace is written into file and tzhe display is suspended
Private sLogTitle           As String
Private Trace               As Collection       ' Collection of begin and end trace entries
Private TraceStack          As Collection       ' Trace stack for the trace log written to a file

Private Property Get DIR_BEGIN_CODE() As String:            DIR_BEGIN_CODE = DIR_BEGIN_ID:                  End Property

Private Property Get DIR_BEGIN_PROC() As String:            DIR_BEGIN_PROC = VBA.String$(2, DIR_BEGIN_ID):  End Property

Private Property Get DIR_END_CODE() As String:              DIR_END_CODE = DIR_END_ID:                      End Property

Private Property Get DIR_END_PROC() As String:              DIR_END_PROC = VBA.String$(2, DIR_END_ID):      End Property

Private Property Get ItmArgs(Optional ByRef trc_entry As Collection) As Variant
    ItmArgs = trc_entry("I")(enPosItmArgs)
End Property

Private Property Get ItmDir(Optional ByRef trc_entry As Collection) As String
    ItmDir = trc_entry("I")(enItmDir)
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

Private Property Get ItmTcks(Optional ByRef trc_entry As Collection) As Currency
    ItmTcks = trc_entry("I")(enItmTcks)
End Property

Private Property Let NtryItm(Optional ByVal trc_entry As Collection, ByVal v As Variant)
    trc_entry.Add v, "I"
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

Public Sub Dsply()
' ----------------------------------------------------------------------------
' Display service, available only when the mMsg component is installed.
' ----------------------------------------------------------------------------
#If MsgComp = 1 Or ErHComp = 1 Then
    mMsg.Box Prompt:=LogTxt(mTrc.LogFile) _
           , Title:="Trasce log provided by the Common VBA Execution Trace Service (displayed by mTrc.Dsply)" _
           , box_monospaced:=True
#Else
    VBA.MsgBox "The mMsg.Box service is not available and the VBA.MsgBox is inappropriate " & _
               "to display a file's content." & vbLf & _
               "Either the Common VBA Message Service is not installed or it is installed " & _
               "but neither of the Conditional Compile Arguments MsgComp, ErHComp is set to 1." & vbLf & vbLf & _
               "The display of the trace result will be done by any text file viewer."
#End If
End Sub

Private Property Get SplitStr(ByRef s As String)
' ----------------------------------------------------------------------------
' Returns the split string in string (s) used by VBA.Split() to turn the
' string into an array.
' ----------------------------------------------------------------------------
    If InStr(s, vbCrLf) <> 0 Then SplitStr = vbCrLf _
    Else If InStr(s, vbLf) <> 0 Then SplitStr = vbLf _
    Else If InStr(s, vbCr) <> 0 Then SplitStr = vbCr
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

Public Property Get LogFile(Optional ByVal tl_append As Boolean = False) As String
    tl_append = tl_append
    LogFile = sLogFile
End Property

Public Property Let LogFile(Optional ByVal tl_append As Boolean = False, _
                                     ByVal tl_file As String)
' ----------------------------------------------------------------------------
' Determines the file to which the execution trace is written to.
' When the Conditional Compile Argument 'ExecTrace = 0' an existing file with
' the provided name is deleted.
' ----------------------------------------------------------------------------
    Dim fso As New FileSystemObject
    
#If ExecTrace = 1 Then
    sLogFile = tl_file
    With fso
        If tl_append = False Then
            If .FileExists(tl_file) Then .DeleteFile tl_file, True
        End If
    End With
#Else
    With fso
        If .FileExists(tl_file) Then .DeleteFile tl_file, True
    End With
#End If
    
    Set fso = Nothing

End Property

Private Function LogInfoLvl() As Long
    
    If ItmDir(LastNtry) Like DIR_END_ID & "*" _
    Then LogInfoLvl = ItmLvl(LastNtry) - 1 _
    Else LogInfoLvl = ItmLvl(LastNtry)

End Function

Public Property Let LogInfo(ByVal tl_inf As String)
' ----------------------------------------------------------------------------
' Write an info line (tl_inf) to the trace log file (sLogFile)
' ----------------------------------------------------------------------------
#If ExecTrace = 1 Then
    Dim LogText As String
    
    If sLogFile <> vbNullString Then
        LogText = LogLinePrefix & String(Len(TRC_LOG_SEC_FRMT) * 2, " ") & RepeatStrng("|  ", LogInfoLvl) & "|  " & tl_inf
        LogTxt(tl_file:=sLogFile, tl_append:=True) = LogText
    End If
#End If
End Property

Public Property Get LogTitle() As String:        LogTitle = sLogTitle:  End Property

Public Property Let LogTitle(ByVal s As String): sLogTitle = s:              End Property

Public Property Get LogTxt( _
                       Optional ByVal tl_file As Variant, _
                       Optional ByVal tl_append As Boolean = True) As String
' ----------------------------------------------------------------------------
' Returns the content of the Trace Log File (tl_file) as string. When the file
' doesn't exist a vbNullString is returned.
' ----------------------------------------------------------------------------
    Const PROC = "Txt-Get"
    
    On Error GoTo eh
    Dim fso As New FileSystemObject
    Dim ts  As TextStream
    Dim s   As String
    Dim sFl As String
   
    tl_append = tl_append ' not used! for declaration compliance and dead code check only
    
    With fso
        If TypeName(tl_file) = "File" Then
            sFl = tl_file.Path
        Else
            '~~ tl_file is regarded a file's full name, created if not existing
            sFl = tl_file
            If Not .FileExists(sFl) Then GoTo xt
        End If
        Set ts = .OpenTextFile(FileName:=sFl, IOMode:=ForReading)
    End With
    
    If Not ts.AtEndOfStream Then
        s = ts.ReadAll
        If VBA.Right$(s, 2) = vbCrLf Then
            s = VBA.Left$(s, Len(s) - 2)
        End If
    Else
        LogTxt = vbNullString
    End If
    If LogTxt = vbCrLf Then LogTxt = vbNullString Else LogTxt = s

xt: Set fso = Nothing
    Exit Property

eh: Select Case ErrMsg(ErrSrc(PROC))
        Case vbResume:  Stop: Resume
        Case Else:      GoTo xt
    End Select
End Property

Private Property Let LogTxt(Optional ByVal tl_file As Variant, _
                                 Optional ByVal tl_append As Boolean = True, _
                                          ByVal tl_string As String)
' ----------------------------------------------------------------------------
' Writes the string (tl_string) into the file (tl_file) which might be a file
' object or a file's full name.
' ----------------------------------------------------------------------------
    Const PROC = "LogTxt-Let"
    
    On Error GoTo eh
    Dim fso As New FileSystemObject
    Dim ts  As TextStream
    Dim sFl As String
   
    With fso
        If TypeName(tl_file) = "File" Then
            sFl = tl_file.Path
        Else
            '~~ tl_file is regarded a file's full name, created if not existing
            sFl = tl_file
            If Not .FileExists(sFl) Then .CreateTextFile sFl
        End If
        
        If tl_append _
        Then Set ts = .OpenTextFile(FileName:=sFl, IOMode:=ForAppending) _
        Else Set ts = .OpenTextFile(FileName:=sFl, IOMode:=ForWriting)
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
' ----------------------------------------------------------------------------
' Begin of code sequence trace.
' ----------------------------------------------------------------------------
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
' ----------------------------------------------------------------------------
' Begin of procedure trace.
' ----------------------------------------------------------------------------
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
' ----------------------------------------------------------------------------
' Begin of procedure trace, specifically for being used by the mErH module.
' ----------------------------------------------------------------------------
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

Public Sub Continue()
' ----------------------------------------------------------------------------
' Continues with counting the execution time
' ----------------------------------------------------------------------------
    cyTcksPaused = cyTcksPaused + (SysCrrntTcks - cyTcksPauseStart)
End Sub

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
'    If StckIsEmpty(TraceStack) Then
'        Dsply
'    End If
    cyTcksOvrhdTrc = SysCrrntTcks - cyTcksOvrhdTrcStrt ' overhead ticks caused by the collection of the end-of-trace entry
#End If
End Sub

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
    If err_source = vbNullString Then err_source = Err.source
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

'Public Sub Finish(Optional ByRef inf As String = vbNullString)
'' ----------------------------------------------------------------------------
'' Finishes an unfinished traced item by means of the stack. All items on the
'' stack are processed via EoP/EoC.
'' ----------------------------------------------------------------------------
'    Dim cll As Collection
'
'    While Not StckIsEmpty(TraceStack)
'        Set cll = StckTop(TraceStack)
'        If NtryIsCode(cll) _
'        Then mTrc.EoC eoc_id:=ItmId(cll), eoc_inf:=inf _
'        Else mTrc.EoP eop_id:=ItmId(cll), eop_inf:=inf
'        inf = vbNullString
'    Wend
'
'End Sub

Private Sub Initialize()
    
    Set Trace = New Collection
    Set LastNtry = Nothing
    dtTraceBegin = Now()
    cyTcksOvrhdItm = 0
    iTrcLvl = 0
    cySysFrequency = 0
    sFirstTraceItem = vbNullString
    
End Sub

Private Function itm(ByVal itm_drctv As String, _
                     ByVal itm_id As String, _
                     ByVal itm_inf As String, _
                     ByVal itm_lvl As Long, _
                     ByVal itm_tckssys As Currency, _
                     ByVal itm_args As Variant) As Variant()
' ----------------------------------------------------------------------------
' Returns an array with the arguments ordered by their enumerated position.
' ----------------------------------------------------------------------------
    Dim av(1 To 6) As Variant
    
    av(enItmDir) = itm_drctv
    av(enItmId) = itm_id
    av(enItmInf) = itm_inf
    av(enItmLvl) = itm_lvl
    av(enItmTcks) = itm_tckssys
    av(enPosItmArgs) = itm_args
    itm = av
    
End Function

Private Function Max(ParamArray va() As Variant) As Variant
' ----------------------------------------------------------------------------
' Returns the maximum value of all values provided (va).
' ----------------------------------------------------------------------------
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
' ----------------------------------------------------------------------------
' Return the arguments as elements in an array as an item in a collection.
' ----------------------------------------------------------------------------
    Const PROC = "Ntry"
    
    On Error GoTo eh
    Dim cll As New Collection
    Dim VarItm  As Variant
    
    VarItm = itm(itm_drctv:=ntry_dir, itm_id:=ntry_id, itm_inf:=ntry_inf, itm_lvl:=ntry_lvl, itm_tckssys:=ntry_tcks, itm_args:=ntry_args)
    NtryItm(cll) = VarItm
    Set Ntry = cll
    
xt: Exit Function
    
eh: Select Case ErrMsg(ErrSrc(PROC))
        Case vbResume:  Stop: Resume
        Case Else:      GoTo xt
    End Select
End Function

Private Function NtryIsCode(ByVal cll As Collection) As Boolean
    
    Select Case ItmDir(cll)
        Case DIR_BEGIN_CODE, DIR_END_CODE: NtryIsCode = True
    End Select

End Function

Public Sub Pause()
    cyTcksPauseStart = SysCrrntTcks
End Sub

Private Function RepeatStrng( _
                       ByVal rs_s As String, _
                       ByVal rs_n As Long) As String
' ----------------------------------------------------------------------------
' Returns the string (s) concatenated (n) times. VBA.String in not appropriate
' because it does not support leading and trailing spaces.
' ----------------------------------------------------------------------------
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

Private Function StckEd(ByVal stck_id As String) As Boolean
' ----------------------------------------------------------------------------
' Returns TRUE when last item pushed to the stack is identical with the item
' (stck_id) and level (stck_lvl).
' ----------------------------------------------------------------------------
    Const PROC = "StckEd"
    
    On Error GoTo eh
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
' Pops the item (stck_item) from the stack (stck) when it is the top item.
' When the top item is not identical with the provided item (stck_item) the
' pop is skipped.
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
    
    If ItmId(cll) = ItmId(cllTop) Then
        Set stck_ppd = cllTop
        TraceStack.Remove TraceStack.Count
        Set cllTop = StckTop(TraceStack)
    Else
        '~~ There is nothing to pop because the top item is not the one requested to pop
        Debug.Print "Stack Pop ='" & ItmId(cll) _
                  & "', Stack Top = '" & ItmId(cllTop) _
                  & "', Stack Dir = '" & ItmDir(cllTop) _
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
' ----------------------------------------------------------------------------
' Should be called by any error handling when a new execution trace is about
' to begin with the very first procedure's execution.
' ----------------------------------------------------------------------------
    Set Trace = Nothing
    Set TraceStack = Nothing
    cyTcksPaused = 0
End Sub

Private Function LogLinePrefix() As String
    LogLinePrefix = Format(Now(), "YY-MM-DD hh:mm:ss ")
End Function

Private Sub LogBgn(ByVal tl_ntry As Collection)
' ----------------------------------------------------------------------------
' Write an begin trace line to the trace log file (sLogFile). If none
' had been provided via the LogFile property the trace log file defaults
' to ExceTrace.log in the Workbook's parent folder.
' ----------------------------------------------------------------------------
    Const PROC = "LogBgn"
    
    On Error GoTo eh
    Dim LogText             As String
    Dim ElapsedSecsTotal    As String
    
    StckPush TraceStack, tl_ntry
    
    If TraceStack.Count = 1 Then
        If mTrc.LogFile = vbNullString Then
            '~~ Provide a default log file not appended when non had been specified
            mTrc.LogFile(False) = Replace(ThisWorkbook.FullName, ThisWorkbook.Name, "ExecTrace.log")
        End If
        
        '~~ When the very first trace entry had been pushed on the stack
        '~~ Provide a trace separator and a trace header
        cyTcksAtStart = ItmTcks(tl_ntry)
    
        If LogTxt(sLogFile) <> vbNullString Then
            LogTxt(sLogFile) = vbNullString ' empty separator line when appended
        End If
        
        '~~ Service header
        LogText = LogLinePrefix & "Execution trace by 'Common VBA Execution Trace Service' (https://github.com/warbe-maker/Common-VBA-Execution-Trace-Service)"
        LogTxt(sLogFile) = LogText
        
        LogText = LogLinePrefix & String((Len(TRC_LOG_SEC_FRMT)) * 2, " ") & ItmDir(tl_ntry)
        If LogTitle = vbNullString _
        Then LogText = LogText & " Begin execution trace " _
        Else LogText = LogText & " " & LogTitle
        LogTxt(sLogFile) = LogText
        '~~ Keep the ticks at start for the calculation of the elepased ticks with each entry
    End If
        
    ElapsedSecsTotal = LogElapsedSecsTotal(ItmTcks(tl_ntry))
    LogText = LogLinePrefix & ElapsedSecsTotal & String(Len(TRC_LOG_SEC_FRMT), " ") & RepeatStrng("|  ", ItmLvl(tl_ntry)) & ItmDir(tl_ntry) & " " & ItmId(tl_ntry)
    LogTxt(sLogFile) = LogText
    
xt: Exit Sub

eh: Select Case ErrMsg(ErrSrc(PROC))
        Case vbResume:  Stop: Resume
        Case Else:      GoTo xt
    End Select
End Sub

Private Function LogElapsedSecs(ByVal et_ticks_end As Currency, _
                                     ByVal et_ticks_start As Currency) As String
    LogElapsedSecs = Format(CDec(et_ticks_end - et_ticks_start) / CDec(SysFrequency), TRC_LOG_SEC_FRMT)
End Function

Private Function LogSecsOverhead()
    LogSecsOverhead = Format(CDec(cyTcksOvrhdTrc / CDec(SysFrequency)), TRC_LOG_SEC_FRMT)
End Function

Private Function LogElapsedSecsTotal(ByVal et_ticks As Currency) As String
    LogElapsedSecsTotal = Format(CDec(et_ticks - cyTcksAtStart) / CDec(SysFrequency), TRC_LOG_SEC_FRMT)
End Function

Private Sub LogEnd(ByVal tl_ntry As Collection)
' ----------------------------------------------------------------------------
' Write an end trace line to the trace log file (sLogFile) - provided one
' had been specified - with the execution time calculated in seconds. When the
' TraceStack is empty write an additional End trace footer line.
' ----------------------------------------------------------------------------
    Const PROC = "LogEnd"
    
    On Error GoTo eh
    Dim BgnNtry             As Collection
    Dim LogText             As String
    Dim ElapsedSecs         As String
    Dim ElapsedSecsTotal    As String
    
    StckPop TraceStack, tl_ntry, BgnNtry
    ElapsedSecsTotal = LogElapsedSecsTotal(ItmTcks(tl_ntry))
    ElapsedSecs = LogElapsedSecs(et_ticks_end:=ItmTcks(tl_ntry), et_ticks_start:=ItmTcks(BgnNtry))
    
    If Not sLogFile = vbNullString Then
        LogText = LogLinePrefix & ElapsedSecsTotal & ElapsedSecs & RepeatStrng("|  ", ItmLvl(tl_ntry)) & ItmDir(tl_ntry) & " " & ItmId(tl_ntry) & ItmInf(tl_ntry)
        LogTxt(sLogFile) = LogText
        If TraceStack.Count = 1 Then
            LogText = LogLinePrefix & ElapsedSecsTotal & ElapsedSecs & ItmDir(tl_ntry) & " "
            If LogTitle = vbNullString _
            Then LogText = LogText & "End execution trace " _
            Else LogText = LogText & LogTitle
            LogTxt(sLogFile) = LogText
            '~~ Service footer
            LogText = LogLinePrefix & String(Len(ElapsedSecsTotal & ElapsedSecs), " ") & "Impact on the overall performance (caused by the trace itself): " & LogSecsOverhead & "seconds!"
            LogTxt(sLogFile) = LogText
            LogText = LogLinePrefix & "Execution trace by 'Common VBA Execution Trace Service' (https://github.com/warbe-maker/Common-VBA-Execution-Trace-Service)"
            LogTxt(sLogFile) = LogText
        End If
    End If

xt: Exit Sub

eh: Select Case ErrMsg(ErrSrc(PROC))
        Case vbResume:  Stop: Resume
        Case Else:      GoTo xt
    End Select
End Sub

Private Sub TrcAdd(ByVal trc_id As String, _
                   ByVal trc_tcks As Currency, _
                   ByVal trc_dir As String, _
                   ByVal trc_lvl As Long, _
          Optional ByVal trc_args As Variant, _
          Optional ByVal trc_inf As String = vbNullString, _
          Optional ByRef trc_ntry As Collection)
' ----------------------------------------------------------------------------
' - Adds an entry to the collection of trace result lines (Trace)
' - With each trc_dir = > a trace-begin entry is pushed to the stack and the
'   trace record is written to the Trace-File-Log
' - With each trc_dir = < the trace-begin entry is popped a n entry is popped from the stack and a trace record
'   is written to the Trace-File-Log with the executed seconds calculated
' ----------------------------------------------------------------------------
    Const PROC = "TrcAdd"
    
    Static sLastDrctv   As String
    Static sLastId      As String
    Static lLastLvl     As String
    
    On Error GoTo eh
    Dim bAlreadyAdded   As Boolean
    
    If Not LastNtry Is Nothing Then
        '~~ When this is not the first entry added the overhead ticks caused by the previous entry is saved.
        '~~ Saving it with the next entry avoids a wrong overhead when saved with the entry itself because.
        '~~ Its maybe nitpicking but worth the try to get execution time figures as correct/exact as possible.
        If sLastId = trc_id And lLastLvl = trc_lvl And sLastDrctv = trc_dir Then bAlreadyAdded = True
        If Not bAlreadyAdded Then
            NtryTcksOvrhdNtry(LastNtry) = cyTcksOvrhdTrc
        Else
            Debug.Print ItmId(LastNtry) & " already added"
        End If
    End If
    
    If Not bAlreadyAdded Then
        Set trc_ntry = Ntry(ntry_tcks:=trc_tcks, ntry_dir:=trc_dir, ntry_id:=trc_id, ntry_lvl:=trc_lvl, ntry_inf:=trc_inf, ntry_args:=trc_args)
        Trace.Add trc_ntry
        
        If trc_dir Like DIR_BEGIN_CODE & "*" Then
            LogBgn trc_ntry
        Else
            LogEnd trc_ntry
        End If
        
        Set LastNtry = trc_ntry
        sLastDrctv = trc_dir
        sLastId = trc_id
        lLastLvl = trc_lvl
    Else
        Debug.Print ItmId(LastNtry) & " already added"
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
           
    iTrcLvl = iTrcLvl + 1
    TrcAdd trc_id:=trc_id _
         , trc_tcks:=cy _
         , trc_dir:=trc_dir _
         , trc_lvl:=iTrcLvl _
         , trc_inf:=vbNullString _
         , trc_args:=trc_args _
         , trc_ntry:=trc_cll
    StckPush TraceStack, trc_cll

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
' ----------------------------------------------------------------------------
' Collect an end trace entry with the current ticks count for the procedure or
' code (item).
' ----------------------------------------------------------------------------
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
    If Not StckEd(stck_id:=trc_id) Then
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
         , trc_inf:=trc_inf _
         , trc_ntry:=trc_cll
         
    StckPop stck:=TraceStack, stck_item:=trc_cll, stck_ppd:=itm
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

