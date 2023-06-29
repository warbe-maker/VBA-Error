Attribute VB_Name = "mTrc"
Option Explicit
' ----------------------------------------------------------------------------
' Standard Module  mTrc: Common VBA Execution Trace Service to trace the
' ====================== execution of procedures and code snippets with the
' highest possible precision regarding the measured elapsed execution time.
' The trace log is written to a file which ensures at least a partial trace
' in case the execution terminates by exception. When this module is installed
' the availability of the sevice is triggered/activated by the Conditional
' Compile Argument 'XcTrc_mTrc = 1'. When the Conditional Compile
' Argument is 0 all services are disabled even when the module is installed,
' avoiding any effect on the performance though the effect very little anyway.
'
' Public services:
' ----------------
' BoC      Indicates the (B)egin (o)f the execution trace of a (C)ode
'          snippet.
' BoP      Indicates the (B)egin (o)f the execution trace of a (P)rocedure.
' BoP_ErH  Exclusively used by the mErH module.
' Continue Commands the execution trace to continue taking the execution
'          time when it had been paused. Pause and Continue is used by the
'          mErH module for example to avoid useless execution time counting
'          while waiting for the users reply.
' Dsply    Displays the content of the trace-log-file (FileFullName).
' EoC      Indicates the (E)nd (o)f the execution trace of a (C)ode snippet.
' EoP      Indicates the (E)nd (o)f the execution trace of a (P)rocedure.
' LogInfo  Explicitly writes an entry to the trace lof file by considering the
'          current nesting level.
' Pause    Stops the execution traces time taking, e.g. while an error message
'          is displayed.
' README   Displays the README in the corresponding public GitHub rebo.
'
' Public Properties:
' ------------------
' FileFullName r/w Specifies the full name of the trace-log-file, defaults to
'                  Path & "\" & FileName when not specified.
' FileName         Specifies the trace-log-file's name, defaults to
'                  "ExecTrace.log" when not specified.
' KeepDays     w   Specifies the number of days a trace-log-file is kept until
'                  it is deleted and re-created.
' Path         r/w Specifies the path to the trace-log-file, defaults to
'                  ThisWorkbook.Path when not specified.
' Title        w   Specifies a trace-log title
'
' Uses (for test purpose only!):
' ------------------------------
' mMsg/fMsg Supports a more comprehensive and well designed error message.
'           See https://github.com/warbe-maker/VBA-Message
'           for how to install and use it in any module.
'
' mErH      Privides an error message with additional information and options.
'           See https://github.com/warbe-maker/VBA-Error
'           for how to install an use in any module.
'
' Requires:
' ---------
' Reference to 'Microsoft Scripting Runtime'
'
' W. Rauschenberger, Berlin, June 2023
' See: https://github.com/warbe-maker/VBA-Trace
' ----------------------------------------------------------------------------
Private Const GITHUB_REPO_URL As String = "https://github.com/warbe-maker/VBA-Trace"

Private fso As New FileSystemObject

#If Not MsgComp = 1 Then
    ' ------------------------------------------------------------------------
    ' The 'minimum error handling' aproach implemented with this module and
    ' provided by the ErrMsg function uses the VBA.MsgBox to display an error
    ' message which includes a debugging option to resume the error line
    ' provided the Cond. Comp. Arg. 'Debugging = 1'.
    ' This declaration allows the mTrc module to work completely autonomous.
    ' It becomes obsolete when the mMsg/fMsg module is installed 1) which must
    ' be indicated by the Cond. Comp. Arg. MsgComp = 1
    '
    ' 1) See https://github.com/warbe-maker/VBA-Message for install and use.
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
    enItmLvl
    enItmTcks
    enItmArgs
End Enum

Private Declare PtrSafe Function apiShellExecute Lib "shell32.dll" _
    Alias "ShellExecuteA" _
    (ByVal hWnd As Long, _
    ByVal lpOperation As String, _
    ByVal lpFile As String, _
    ByVal lpParameters As String, _
    ByVal lpDirectory As String, _
    ByVal nShowCmd As Long) _
    As Long
'***App Window Constants***
Private Const WIN_NORMAL = 1         'Open Normal

'***Error Codes***
Private Const ERROR_SUCCESS = 32&
Private Const ERROR_NO_ASSOC = 31&
Private Const ERROR_OUT_OF_MEM = 0&
Private Const ERROR_FILE_NOT_FOUND = 2&
Private Const ERROR_PATH_NOT_FOUND = 3&
Private Const ERROR_BAD_FORMAT = 11&

Private Declare PtrSafe Function getFrequency Lib "kernel32" _
Alias "QueryPerformanceFrequency" (cySysFrequency As Currency) As Long
Private Declare PtrSafe Function getTickCount Lib "kernel32" _
Alias "QueryPerformanceCounter" (cyTickCount As Currency) As Long

Private Const DIR_BEGIN_ID      As String = ">"     ' Begin procedure or code trace indicator
Private Const DIR_END_ID        As String = "<"     ' End procedure or code trace indicator
Private Const TRC_INFO_DELIM    As String = " !!! "
Private Const TRC_LOG_SEC_FRMT  As String = "00.0000 "

Private bLogToFileSuspended As Boolean
Private cySysFrequency      As Currency         ' Execution Trace SysFrequency (initialized with init)
Private cyTcksAtStart       As Currency         ' Trace log to file
Private cyTcksOvrhdItm      As Currency         ' Execution Trace time accumulated by caused by the time tracking itself
Private cyTcksOvrhdTrc      As Currency         ' Overhead ticks caused by the collection of a traced item's entry
Private cyTcksOvrhdTrcStrt  As Currency         ' Overhead ticks caused by the collection of a traced item's entry
Private cyTcksPaused        As Currency         ' Accumulated with procedure Continue
Private cyTcksPauseStart    As Currency         ' Set with procedure Pause
Private dtTraceBegin        As Date             ' Initialized at start of execution trace
Private iTrcLvl             As Long             ' Increased with each begin entry and decreased with each end entry
Private LastNtry            As Collection       '
Private lKeepDays           As Long
Private oLogFile            As TextStream
Private sFileFullName       As String           ' Defaults to the When vbNullString the trace is written into file and the display is suspended
Private sFileName           As String
Private sFirstTraceItem     As String
Private sTitle              As String
Private sPath               As String
Private TraceStack          As Collection       ' Trace stack for the trace log written to a file

Private Property Get DIR_BEGIN_CODE() As String:            DIR_BEGIN_CODE = DIR_BEGIN_ID:                  End Property

Private Property Get DIR_BEGIN_PROC() As String:            DIR_BEGIN_PROC = VBA.String$(2, DIR_BEGIN_ID):  End Property

Private Property Get DIR_END_CODE() As String:              DIR_END_CODE = DIR_END_ID:                      End Property

Private Property Get DIR_END_PROC() As String:              DIR_END_PROC = VBA.String$(2, DIR_END_ID):      End Property

Public Property Get FileFullName() As String

    If sFileFullName = vbNullString _
    Then FileFullName = Path & "\" & FileName _
    Else FileFullName = sFileFullName
    
    With fso
        If Not .FileExists(FileFullName) Then .CreateTextFile FileFullName
    End With

End Property

Public Property Let FileFullName(ByVal s As String)
' ----------------------------------------------------------------------------
' Specifies the trace-log-file's name and location, thereby maintaining the
' Path and the FileName property.
' ----------------------------------------------------------------------------
    Dim lDaysAge As Long
    
    With fso
        If sFileFullName <> s Then
            '~~ Either the trace-log-file's name has yet not been initialized
            '~~ or the name is not/no longer the one previously used
            If .FileExists(sFileFullName) Then .DeleteFile sFileFullName
        End If
        
        sFileFullName = s
        sPath = .GetParentFolderName(s)
        If Not .FileExists(FileFullName) Then .CreateTextFile FileFullName
        sFileName = .GetFileName(s)
        
        '~~ In case the file already existed it may have passed the KeepDays limit
        lDaysAge = VBA.DateDiff("d", .GetFile(sFileFullName).DateLastAccessed, Now())
        If lDaysAge > KeepDays Then
            .DeleteFile sFileFullName
            .CreateTextFile sFileFullName
        End If
    End With

End Property

Public Property Get FileName() As String

    If sFileName = vbNullString _
    Then FileName = "ExecTrace.log" _
    Else FileName = sFileName
    
End Property

Public Property Let FileName(ByVal s As String)
    If fso.GetExtensionName(s) = vbNullString Then
        s = s & ".log"
    End If
    sFileName = s
End Property

Private Property Get ItmArgs(Optional ByRef t_entry As Collection) As Variant
    ItmArgs = t_entry("I")(enItmArgs)
End Property

Private Property Get ItmDir(Optional ByRef t_entry As Collection) As String
    ItmDir = t_entry("I")(enItmDir)
End Property

Private Property Get ItmId(Optional ByRef t_entry As Collection) As String
    ItmId = t_entry("I")(enItmId)
End Property

Private Property Get ItmLvl(Optional ByRef t_entry As Collection) As Long
    ItmLvl = t_entry("I")(enItmLvl)
End Property

Private Property Get ItmTcks(Optional ByRef t_entry As Collection) As Currency
    ItmTcks = t_entry("I")(enItmTcks)
End Property

Private Property Get KeepDays() As Long

    If lKeepDays = 0 _
    Then KeepDays = 10 _
    Else KeepDays = lKeepDays
    
End Property

Public Property Let KeepDays(ByVal l As Long):  lKeepDays = l:  End Property

Private Property Let Log(ByVal tl_string As String)
' ----------------------------------------------------------------------------
' Writes the string (tl_string) to the FileFullName provided one exists.
' Precondition: A FileFullName is specified (mTrc.FileFullName() = "xx").
' ----------------------------------------------------------------------------
    Const PROC = "Log-Let"
    
    On Error GoTo eh
    Dim oFile   As TextStream
    
    With fso
        If Not .FileExists(FileFullName) Then .CreateTextFile (FileFullName)
        Set oFile = .OpenTextFile(FileFullName, ForAppending)
        oFile.WriteLine tl_string
    End With
    
xt: Exit Property
    
eh: Select Case ErrMsg(ErrSrc(PROC))
        Case vbResume:  Stop: Resume
        Case Else:      GoTo xt
    End Select
End Property

Public Property Let LogInfo(ByVal tl_inf As String)
' ----------------------------------------------------------------------------
' Write an info line (tl_inf) to the trace-log-file (FileFullName)
' ----------------------------------------------------------------------------
    
    Log = LogNow & String(Len(TRC_LOG_SEC_FRMT) * 2, " ") & RepeatStrng("|  ", LogInfoLvl) & "|  " & tl_inf

End Property

Public Property Get LogSuspended() As Boolean:              LogSuspended = bLogToFileSuspended:             End Property

Public Property Let LogSuspended(ByVal b As Boolean):       bLogToFileSuspended = b:                        End Property

Private Property Let NtryItm(Optional ByVal t_entry As Collection, ByVal v As Variant)
    t_entry.Add v, "I"
End Property

Private Property Get NtryTcksOvrhdNtry(Optional ByRef t_entry As Collection) As Currency
    On Error Resume Next
    NtryTcksOvrhdNtry = t_entry("TON")
    If Err.Number <> 0 Then NtryTcksOvrhdNtry = 0
End Property

Private Property Let NtryTcksOvrhdNtry(Optional ByRef t_entry As Collection, ByRef cy As Currency)
    If t_entry Is Nothing Then Set t_entry = New Collection
    t_entry.Add cy, "TON"
End Property

Public Property Get Path() As Variant
    If sPath = vbNullString _
    Then Path = ThisWorkbook.Path _
    Else Path = sPath
End Property

Public Property Let Path(ByVal v As Variant)
' -----------------------------------------------------------------------------------
' Specifies the location (folder) for the log file based on the provided information
' which may be a string, a Workbook, or a folder object.
' -----------------------------------------------------------------------------------
    Const PROC = "Path-Let"
    Dim wbk As Workbook
    Dim fld As Folder
    
    Select Case VarType(v)
        Case VarType(v) = vbString
            sPath = v
        Case VarType(v) = vbObject
            If TypeOf v Is Workbook Then
                Set wbk = v
                sPath = wbk.Path
            ElseIf TypeOf v Is Folder Then
                Set fld = v
                sPath = fld.Path
            Else
                Err.Raise AppErr(1), ErrSrc(PROC), "The provided argument is neither a string specifying a " & _
                                                   "folder's path, nor a Workbook object, nor a Folder object!"
            End If
    End Select
    FileFullName = Path & "\" & FileName ' re-establishes it when not already existing
    
End Property

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

Public Property Let Title(ByVal s As String):               sTitle = s:                                  End Property

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

Public Sub BoC(ByVal b_id As String, _
      Optional ByVal b_args As String = vbNullString)
' ----------------------------------------------------------------------------
' Begin of code sequence trace.
' ----------------------------------------------------------------------------
    Dim cll As Collection
    
    cyTcksOvrhdTrcStrt = SysCrrntTcks
    TrcBgn t_id:=b_id, t_dir:=DIR_BEGIN_CODE, t_args:=b_args, t_cll:=cll
    cyTcksOvrhdTrc = SysCrrntTcks - cyTcksOvrhdTrcStrt ' overhead ticks caused by the collection of the begin trace entry

End Sub

Public Sub BoP(ByVal b_id As String, _
      Optional ByVal b_args As String = vbNullString)
' ----------------------------------------------------------------------------
' Begin of procedure trace.
' ----------------------------------------------------------------------------
    Dim cll As Collection
        
    cyTcksOvrhdTrcStrt = SysCrrntTcks
    If sFirstTraceItem = vbNullString Then
        Initialize
        sFirstTraceItem = b_id
    Else
        If b_id = sFirstTraceItem Then
            '~~ A previous trace had not come to a regular end and thus will be erased
            Initialize
        End If
    End If
    TrcBgn t_id:=b_id, t_dir:=DIR_BEGIN_PROC, t_args:=b_args, t_cll:=cll
    cyTcksOvrhdTrc = SysCrrntTcks - cyTcksOvrhdTrcStrt ' overhead ticks caused by the collection of the begin trace entry

End Sub

Public Sub BoP_ErH(ByVal b_id As String, _
          Optional ByVal b_args As String = vbNullString)
' ----------------------------------------------------------------------------
' Begin of procedure trace, specifically for being used by the mErH module.
' ----------------------------------------------------------------------------
    Dim cll As Collection
    
    cyTcksOvrhdTrcStrt = SysCrrntTcks
    If sFirstTraceItem = vbNullString Then
        Initialize
        sFirstTraceItem = b_id
    Else
        If b_id = sFirstTraceItem Then
            '~~ A previous trace had not come to a regular end and thus will be erased
            Initialize
        End If
    End If
    TrcBgn t_id:=b_id, t_dir:=DIR_BEGIN_PROC, t_args:=b_args, t_cll:=cll
    cyTcksOvrhdTrc = SysCrrntTcks - cyTcksOvrhdTrcStrt ' overhead ticks caused by the collection of the begin trace entry

End Sub

Public Sub Continue()
' ----------------------------------------------------------------------------
' Continues with counting the execution time
' ----------------------------------------------------------------------------
    cyTcksPaused = cyTcksPaused + (SysCrrntTcks - cyTcksPauseStart)
End Sub

Public Sub Dsply()
' ----------------------------------------------------------------------------
' Display service using ShellRun to open the log-file by means of the
' application associated with the log-file's file-extenstion.
' ----------------------------------------------------------------------------
    ShellRun FileFullName, WIN_NORMAL
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

Private Function DsplyArgs(ByVal t_entry As Collection) As String
' ------------------------------------------------------------------------------
' Returns a string with the collection of the traced arguments. Any entry ending
' with a ":" or "=" is an arguments name with its value in the subsequent item.
' ------------------------------------------------------------------------------
    Dim va()    As Variant
    Dim i       As Long
    Dim sL      As String
    Dim sR      As String
    
    On Error Resume Next
    va = ItmArgs(t_entry)
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
' ------------------------------------------------------------------------------
' End of the trace of a code sequence.
' ------------------------------------------------------------------------------
    Dim cll As Collection
    
    cyTcksOvrhdTrcStrt = SysCrrntTcks
    If StckIsEmpty(TraceStack) Then Exit Sub
    TrcEnd t_id:=eoc_id, t_dir:=DIR_END_CODE, t_args:=eoc_inf, t_cll:=cll
    cyTcksOvrhdTrc = SysCrrntTcks - cyTcksOvrhdTrcStrt ' overhead ticks caused by the collection of the begin trace entry

End Sub

Public Sub EoP(ByVal e_id As String, _
      Optional ByVal e_args As String = vbNullString)
' ------------------------------------------------------------------------------
' End of the trace of a procedure.
' ------------------------------------------------------------------------------
    Dim cll As Collection
    
    cyTcksOvrhdTrcStrt = SysCrrntTcks
    If StckIsEmpty(TraceStack) Then Exit Sub        ' Nothing to trace any longer. Stack has been emptied after an error to finish the trace
    
    TrcEnd t_id:=e_id, t_dir:=DIR_END_PROC, t_args:=e_args, t_cll:=cll
    cyTcksOvrhdTrc = SysCrrntTcks - cyTcksOvrhdTrcStrt ' overhead ticks caused by the collection of the end-of-trace entry
End Sub

Private Function ErrMsg(ByVal err_source As String, _
               Optional ByVal err_no As Long = 0, _
               Optional ByVal err_dscrptn As String = vbNullString, _
               Optional ByVal err_line As Long = 0) As Variant
' ------------------------------------------------------------------------------
' Universal error message display service which displays:
' - a debugging option button (Conditional Compile Argument 'Debugging = 1')
' - an optional additional "About:" section when the err_dscrptn has an
'   additional string concatenated by two vertical bars (||)
' - the error message by means of the Common VBA Message Service (fMsg/mMsg)
'   Common Component
'   mMsg (Conditional Compile Argument "MsgComp = 1") is installed.
'
' Uses:
' - AppErr  For programmed application errors (Err.Raise AppErr(n), ....)
'           to turn them into a negative and in the error message back into
'           its origin positive number.
' - ErrSrc  To provide an unambiguous procedure name by prefixing is with
'           the module name.
'
' W. Rauschenberger Berlin, Apr 2023
'
' See: https://github.com/warbe-maker/VBA-Error
' ------------------------------------------------------------------------------
#If ErHComp = 1 Then
    '~~ When Common VBA Error Services (mErH) is availabel in the VB-Project
    '~~ (which includes the mMsg component) the mErh.ErrMsg service is invoked.
    ErrMsg = mErH.ErrMsg(err_source, err_no, err_dscrptn, err_line): GoTo xt
    GoTo xt
#ElseIf MsgComp = 1 Then
    '~~ When (only) the Common Message Service (mMsg, fMsg) is available in the
    '~~ VB-Project, mMsg.ErrMsg is invoked for the display of the error message.
    ErrMsg = mMsg.ErrMsg(err_source, err_no, err_dscrptn, err_line): GoTo xt
    GoTo xt
#End If
    '~~ When neither of the Common Component is available in the VB-Project
    '~~ the error message is displayed by means of the VBA.MsgBox
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
    
    '~~ Consider extra information is provided with the error description
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
            If err_dscrptn Like "*DAO*" _
            Or err_dscrptn Like "*ODBC*" _
            Or err_dscrptn Like "*Oracle*" _
            Then ErrType = "Database Error " _
            Else ErrType = "VB Runtime Error "
    End Select
    
    If err_source <> vbNullString Then ErrSrc = " in: """ & err_source & """"   ' assemble ErrSrc from available information"
    If err_line <> 0 Then ErrAtLine = " at line " & err_line                    ' assemble ErrAtLine from available information
    ErrTitle = Replace(ErrType & ErrNo & ErrSrc & ErrAtLine, "  ", " ")         ' assemble ErrTitle from available information
       
    ErrText = "Error: " & vbLf & ErrDesc & vbLf & vbLf & "Source: " & vbLf & err_source & ErrAtLine
    If ErrAbout <> vbNullString Then ErrText = ErrText & vbLf & vbLf & "About: " & vbLf & ErrAbout
    
#If Debugging = 1 Then
    ErrBttns = vbYesNo
    ErrText = ErrText & vbLf & vbLf & "Debugging:" & vbLf & "Yes    = Resume Error Line" & vbLf & "No     = Terminate"
#Else
    ErrBttns = vbCritical
#End If
    ErrMsg = MsgBox(Title:=ErrTitle, Prompt:=ErrText, Buttons:=ErrBttns)
xt:
End Function

Private Function ErrSrc(ByVal sProc As String) As String
    ErrSrc = "mTrc." & sProc
End Function

Public Sub Initialize()
' ----------------------------------------------------------------------------
' - Initializes defaults - when yet no other had been specified
' - Initializes the means for a new trace log
' ----------------------------------------------------------------------------
        
    Set LastNtry = Nothing
    dtTraceBegin = Now()
    cyTcksOvrhdItm = 0
    iTrcLvl = 0
    cySysFrequency = 0
    sFirstTraceItem = vbNullString
    lKeepDays = 10
    
End Sub

Private Function Itm(ByVal i_id As String, _
                     ByVal i_dir As String, _
                     ByVal i_lvl As Long, _
                     ByVal i_tcks As Currency, _
                     ByVal i_args As String) As Variant()
' ----------------------------------------------------------------------------
' Returns an array with the arguments ordered by their enumerated position.
' ----------------------------------------------------------------------------
    Dim av(1 To 5) As Variant
    
    av(enItmId) = i_id
    av(enItmDir) = i_dir
    av(enItmLvl) = i_lvl
    av(enItmTcks) = i_tcks
    av(enItmArgs) = i_args
    Itm = av
    
End Function

Private Sub LogBgn(ByVal l_ntry As Collection, _
          Optional ByVal l_args As String = vbNullString)
' ----------------------------------------------------------------------------
' Writes a begin trace line to the trace-log-file (sFileFullName).
' ----------------------------------------------------------------------------
    Const PROC = "LogBgn"
    
    On Error GoTo eh
    Dim sLogText            As String
    Dim ElapsedSecsTotal    As String
    Dim ElapsedSecs         As String
    Dim TopNtry             As Collection
    Dim s                   As String
    
    Set TopNtry = StckTop(TraceStack)
    If TopNtry Is Nothing _
    Then ElapsedSecsTotal = vbNullString _
    Else ElapsedSecsTotal = LogElapsedSecsTotal(ItmTcks(l_ntry))
    StckPush TraceStack, l_ntry
    
    If TraceStack.Count = 1 Then
        cyTcksAtStart = ItmTcks(l_ntry)
    
        '~~ Trace-Log header
        s = LogText(l_elpsd_total:="Elapsed " _
                  , l_elpsd_secs:="Length  " _
                  , l_strng:="Execution trace by 'Common VBA Execution Trace Service (mTrc)' (" & GITHUB_REPO_URL & ")")
        If fso.GetFile(FileFullName).Size > 0 Then Log = String(Len(s), "=")
        Log = s
        
        '~~ Trace-Log title
        If sTitle = vbNullString Then
            Log = LogText(l_elpsd_total:="seconds " _
                        , l_elpsd_secs:="seconds " _
                        , l_strng:=ItmDir(l_ntry) & " Begin execution trace ")
        Else
            Log = LogText(l_elpsd_total:="seconds " _
                        , l_elpsd_secs:="seconds " _
                        , l_strng:=ItmDir(l_ntry) & " " & sTitle)
        End If
        '~~ Keep the ticks at start for the calculation of the elepased ticks with each entry
    End If
        
    ElapsedSecsTotal = LogElapsedSecsTotal(ItmTcks(l_ntry))
    Log = LogText(l_elpsd_total:=ElapsedSecsTotal _
                , l_elpsd_secs:=ElapsedSecs _
                , l_strng:=RepeatStrng("|  ", ItmLvl(l_ntry)) & ItmDir(l_ntry) & " " & ItmId(l_ntry) _
                , l_args:=l_args)
    
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

Private Function LogElapsedSecsTotal(ByVal et_ticks As Currency) As String
    LogElapsedSecsTotal = Format(CDec(et_ticks - cyTcksAtStart) / CDec(SysFrequency), TRC_LOG_SEC_FRMT)
End Function

Private Sub LogEnd(ByVal l_ntry As Collection, _
          Optional ByVal l_args As String = vbNullString)
' ------------------------------------------------------------------------------
' Write an end trace line to the trace-log-file (sFileFullName) - provided one
' had been specified - with the execution time calculated in seconds. When the
' TraceStack is empty write an additional End trace footer line.
' ------------------------------------------------------------------------------
    Const PROC = "LogEnd"
    
    On Error GoTo eh
    Dim BgnNtry             As Collection
    Dim sLogText            As String
    Dim ElapsedSecs         As String
    Dim ElapsedSecsTotal    As String
    
    StckPop TraceStack, l_ntry, BgnNtry
    ElapsedSecsTotal = LogElapsedSecsTotal(ItmTcks(l_ntry))
    ElapsedSecs = LogElapsedSecs(et_ticks_end:=ItmTcks(l_ntry), et_ticks_start:=ItmTcks(BgnNtry))
    
    Log = LogText(l_elpsd_total:=ElapsedSecsTotal _
                , l_elpsd_secs:=ElapsedSecs _
                , l_strng:=RepeatStrng("|  ", ItmLvl(l_ntry)) _
                        & ItmDir(l_ntry) _
                        & " " _
                        & ItmId(l_ntry) _
               , l_args:=ItmArgs(l_ntry))
    
    If TraceStack.Count = 1 Then
        
        '~~ Trace bottom title
        If sTitle = vbNullString Then
            Log = LogText(l_elpsd_total:=ElapsedSecsTotal _
                        , l_elpsd_secs:=ElapsedSecs _
                        , l_strng:=ItmDir(l_ntry) & " " & "End execution trace ")
        Else
            Log = LogText(l_elpsd_total:=ElapsedSecsTotal _
                        , l_elpsd_secs:=ElapsedSecs _
                        , l_strng:=ItmDir(l_ntry) & " " & sTitle)
        End If
        
        '~~ Trace footer and summary
        Log = LogText(l_strng:="Execution trace by 'Common VBA Execution Trace Service (mTrc)' (" & GITHUB_REPO_URL & ")")
        Log = LogText(l_strng:="Impact on the overall performance (caused by the trace itself): " & Format(LogSecsOverhead * 1000, "#0.0") & " milliseconds!")
    End If
xt: Exit Sub

eh: Select Case ErrMsg(ErrSrc(PROC))
        Case vbResume:  Stop: Resume
        Case Else:      GoTo xt
    End Select
End Sub

Private Sub LogEntry(ByVal t_id As String, _
                     ByVal t_dir As String, _
                     ByVal t_lvl As Long, _
                     ByVal t_tcks As Currency, _
            Optional ByVal t_args As String = vbNullString, _
            Optional ByRef t_ntry As Collection)
' ----------------------------------------------------------------------------
' Writes an entry to the trace-log-file by meand of LogBgn or LogEnd.
' ----------------------------------------------------------------------------
    Const PROC = "LogEntry"
    
    On Error GoTo eh
    Static sLastDrctv   As String
    Static sLastId      As String
    Static lLastLvl     As String
    Dim bAlreadyAdded   As Boolean
    
    If Not LastNtry Is Nothing Then
        '~~ When this is not the first entry added the overhead ticks caused by the previous entry is saved.
        '~~ Saving it with the next entry avoids a wrong overhead when saved with the entry itself because.
        '~~ Its maybe nitpicking but worth the try to get execution time figures as correct/exact as possible.
        If sLastId = t_id And lLastLvl = t_lvl And sLastDrctv = t_dir Then bAlreadyAdded = True
        If Not bAlreadyAdded Then
            NtryTcksOvrhdNtry(LastNtry) = cyTcksOvrhdTrc
        Else
            Debug.Print ItmId(LastNtry) & " already added"
        End If
    End If
    
    If Not bAlreadyAdded Then
        Set t_ntry = Ntry(n_tcks:=t_tcks, n_dir:=t_dir, n_id:=t_id, n_lvl:=t_lvl, n_args:=t_args)
        If t_dir Like DIR_BEGIN_CODE & "*" _
        Then LogBgn t_ntry, t_args _
        Else LogEnd t_ntry, t_args
        
        Set LastNtry = t_ntry
        sLastDrctv = t_dir
        sLastId = t_id
        lLastLvl = t_lvl
    Else
        Debug.Print ItmId(LastNtry) & " already added"
    End If

xt: Exit Sub

eh: Select Case ErrMsg(ErrSrc(PROC))
        Case vbResume:  Stop: Resume
        Case Else:      GoTo xt
    End Select
End Sub

Private Function LogInfoLvl() As Long
    
    If ItmDir(LastNtry) Like DIR_END_ID & "*" _
    Then LogInfoLvl = ItmLvl(LastNtry) - 1 _
    Else LogInfoLvl = ItmLvl(LastNtry)

End Function

Private Function LogNow() As String
    LogNow = Format(Now(), "YY-MM-DD hh:mm:ss ")
End Function

Private Function LogSecsOverhead()
    LogSecsOverhead = Format(CDec(cyTcksOvrhdTrc / CDec(SysFrequency)), TRC_LOG_SEC_FRMT)
End Function

Private Function LogText(Optional ByVal l_elpsd_total As String = vbNullString, _
                         Optional ByVal l_elpsd_secs As String = vbNullString, _
                         Optional ByVal l_strng As String = vbNullString, _
                         Optional ByVal l_args As String = vbNullString) As String
' ----------------------------------------------------------------------------
' Returns the uniformed assemled log text.
' ----------------------------------------------------------------------------
    LogText = LogNow
    
    If l_elpsd_total = vbNullString Then l_elpsd_total = String((Len(TRC_LOG_SEC_FRMT)), " ")
    LogText = LogText & l_elpsd_total
    
    If l_elpsd_secs = vbNullString Then l_elpsd_secs = String((Len(TRC_LOG_SEC_FRMT)), " ")
    LogText = LogText & l_elpsd_secs
    
    LogText = LogText & l_strng
    
    If l_args <> vbNullString Then
        If InStr(l_args, "!!") <> 0 _
        Then LogText = LogText & l_args _
        Else LogText = LogText & " (" & l_args & ")"
    End If
    
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

Public Sub NewFile()
' ----------------------------------------------------------------------------
' Deletes an existing trace-log-file and creates a new one. Mainly used for
' testing where a new trace-log-file is usually appropriate.
' ----------------------------------------------------------------------------
    
    With fso
        If .FileExists(FileFullName) Then .DeleteFile FileFullName, True
        .CreateTextFile FileFullName
    End With
    
End Sub

Private Function Ntry(ByVal n_id As String, _
                      ByVal n_dir As String, _
                      ByVal n_lvl As Long, _
                      ByVal n_tcks As Currency, _
                      ByVal n_args As Variant) As Collection
' ----------------------------------------------------------------------------
' Return the arguments as elements in an array as an item in a collection.
' ----------------------------------------------------------------------------
    Const PROC = "Ntry"
    
    On Error GoTo eh
    Dim cll As New Collection
    Dim VarItm  As Variant
    
    VarItm = Itm(i_id:=n_id, i_dir:=n_dir, i_lvl:=n_lvl, i_tcks:=n_tcks, i_args:=n_args)
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

Public Sub README(Optional ByVal r_bookmark As String = vbNullString)
' ------------------------------------------------------------------------------
'
' ------------------------------------------------------------------------------
    If r_bookmark = vbNullString Then
        ShellRun GITHUB_REPO_URL
    Else
        ShellRun GITHUB_REPO_URL & Replace("#" & r_bookmark, "##", "#")
    End If

End Sub

Private Function RepeatStrng(ByVal rs_s As String, _
                             ByVal rs_n As Long) As String
' ----------------------------------------------------------------------------
' Returns the string (s) concatenated (n) times. VBA.String in not appropriate
' because it does not support leading and trailing spaces.
' ----------------------------------------------------------------------------
    Dim i   As Long
    
    For i = 1 To rs_n: RepeatStrng = RepeatStrng & rs_s:  Next i

End Function

Private Function ShellRun(ByVal oue_string As String, _
                 Optional ByVal oue_show_how As Long = WIN_NORMAL) As String
' ----------------------------------------------------------------------------
' Opens a folder, email-app, url, or even an Access instance.
'
' Usage Examples: - Open a folder:  ShellRun("C:\TEMP\")
'                 - Call Email app: ShellRun("mailto:user@tutanota.com")
'                 - Open URL:       ShellRun("http://.......")
'                 - Unknown:        ShellRun("C:\TEMP\Test") (will call
'                                   "Open With" dialog)
'                 - Open Access DB: ShellRun("I:\mdbs\xxxxxx.mdb")
' Copyright:      This code was originally written by Dev Ashish. It is not to
'                 be altered or distributed, except as part of an application.
'                 You are free to use it in any application, provided the
'                 copyright notice is left unchanged.
' Courtesy of:    Dev Ashish
' ----------------------------------------------------------------------------

    Dim lRet            As Long
    Dim varTaskID       As Variant
    Dim stRet           As String
    Dim hWndAccessApp   As Long
    
    '~~ First try ShellExecute
    lRet = apiShellExecute(hWndAccessApp, vbNullString, oue_string, vbNullString, vbNullString, oue_show_how)
    
    Select Case True
        Case lRet = ERROR_OUT_OF_MEM:       stRet = "Execution failed: Out of Memory/Resources!"
        Case lRet = ERROR_FILE_NOT_FOUND:   stRet = "Execution failed: File not found!"
        Case lRet = ERROR_PATH_NOT_FOUND:   stRet = "Execution failed: Path not found!"
        Case lRet = ERROR_BAD_FORMAT:       stRet = "Execution failed: Bad File Format!"
        Case lRet = ERROR_NO_ASSOC          ' Try the OpenWith dialog
            varTaskID = Shell("rundll32.exe shell32.dll,OpenAs_RunDLL " & oue_string, WIN_NORMAL)
            lRet = (varTaskID <> 0)
        Case lRet > ERROR_SUCCESS:          lRet = -1
    End Select
    
    ShellRun = lRet & IIf(stRet = vbNullString, vbNullString, ", " & stRet)

End Function

Private Sub StckAdjust(ByVal t_id As String)
    Dim cllNtry As Collection
    Dim i       As Long
    
    For i = TraceStack.Count To 1 Step -1
        Set cllNtry = TraceStack(i)
        If ItmId(cllNtry) = t_id Then
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
        If ItmId(cllNtry) = stck_id Then
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
            EoC eoc_id:=ItmId(cllTop), eoc_inf:="ended by stack!!"
        Else
            EoP e_id:=ItmId(cllTop), e_args:="ended by stack!!"
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
    Set TraceStack = Nothing
    cyTcksPaused = 0
End Sub

Private Sub TrcBgn(ByVal t_id As String, _
                   ByVal t_dir As String, _
          Optional ByVal t_args As String = vbNullString, _
          Optional ByRef t_cll As Collection)
' ----------------------------------------------------------------------------
' Collect a trace begin entry with the current ticks count for the procedure
' or code (item).
' ----------------------------------------------------------------------------
    Const PROC = "TrcEnd"
    
    On Error GoTo eh
           
    iTrcLvl = iTrcLvl + 1
    LogEntry t_id:=t_id _
           , t_dir:=t_dir _
           , t_lvl:=iTrcLvl _
           , t_tcks:=SysCrrntTcks - cyTcksPaused _
           , t_args:=t_args _
           , t_ntry:=t_cll
    StckPush TraceStack, t_cll

xt: Exit Sub
    
eh: Select Case ErrMsg(ErrSrc(PROC))
        Case vbResume:  Stop: Resume
        Case Else:      GoTo xt
    End Select
End Sub

Private Sub TrcEnd(ByVal t_id As String, _
                   ByVal t_dir As String, _
          Optional ByVal t_args As String = vbNullString, _
          Optional ByRef t_cll As Collection)
' ----------------------------------------------------------------------------
' Collect an end trace entry with the current ticks count for the procedure or
' code (item).
' ----------------------------------------------------------------------------
    Const PROC = "TrcEnd"
    
    On Error GoTo eh
    Dim Top As Collection:  Set Top = StckTop(TraceStack)
    Dim Itm As Collection
    
    '~~ Any end trace for an item not on the stack is ignored. On the other hand,
    '~~ if on the stack but not the last item the stack is adjusted because this
    '~~ indicates a begin without a corresponding end trace statement.
    If Not StckEd(t_id) _
    Then Exit Sub _
    Else StckAdjust t_id
    
    If ItmId(Top) <> t_id And ItmLvl(Top) = iTrcLvl Then StckPop TraceStack, Top

    LogEntry t_id:=t_id _
           , t_dir:=t_dir _
           , t_lvl:=iTrcLvl _
           , t_tcks:=SysCrrntTcks - cyTcksPaused _
           , t_args:=t_args _
           , t_ntry:=t_cll
         
    StckPop stck:=TraceStack, stck_item:=t_cll, stck_ppd:=Itm
    iTrcLvl = iTrcLvl - 1

xt: Exit Sub
    
eh: Select Case ErrMsg(ErrSrc(PROC))
        Case vbResume:  Stop: Resume
        Case Else:      GoTo xt
    End Select
End Sub

