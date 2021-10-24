Attribute VB_Name = "mErH"
Option Explicit
Option Private Module
' ------------------------------------------------------------------------------
' Standard  Module mErH: Global error handling for any VBA Project.
'
' Methods:
' - AppErr   Converts a positive number into a negative error number ensuring it
'            not conflicts with a VB Runtime Error. A negative error number is
'            turned back into the original positive Application  Error Number.
' - ErrMsg   Either passes on the error to the caller or when the Entry-Procedure
'            is reached, displays the error with a complete path from the
'            Entry-Procedure to the procedure with the error.
' - BoP      Maintains the call stack at the Begin of a Procedure (optional when
'            using this common error handler)
' - EoP      Maintains the call stack at the End of a Procedure, triggers the
'            display of the Execution Trace when the 'Entry-Procedure' is
'            finished and the  Conditional Compile Argument ExecTrace = 1
' - ErrDsply Displays the error message in a proper formated manner
'            The local Conditional Compile Argument "AlternativeMsgBox = 1"
'            enforces the use of the Alternative VBA MsgBox which provideds an
'            improved readability.
'
' Note: When never a mErH.BoP/mErH.EoP procedure had been executed the ErrMsg
'       is displayed with the procedure the error occoured. Else the error is
'       passed on back up to the first procedure with a mErH.BoP/mErH.EoP code
'       line executed and displayed when it had been reached.
'
' Uses: fMsg
'       mTrc (optionally, when the Conditional Compile Argument ExecTrace = 1)
'
' Requires: Reference to "Microsoft Scripting Runtime"
'
' For further details see the Github blog post: "Comprehensive Common VBA Error Handler"
' https://warbe-maker.github.io/vba/common/2020/10/02/Comprehensive-Common-VBA-Error-Handler.html
'
' W. Rauschenberger, Berlin, Nov 2020
' ------------------------------------------------------------------------------

Public Const CONCAT         As String = "||"

Private cllErrPath          As Collection
Private cllErrorPath        As Collection   ' managed by ErrPath... procedures exclusively
Private dctStck             As Dictionary
Private sErrHndlrEntryProc  As String
Private lSubsequErrNo       As Long         ' possibly different from the initial error number if it changes when passed on
Private vErrsAsserted       As Variant      ' possibly provided with BoTP
Private vErrReply           As Variant
Private vArguments()        As Variant      ' The last procedures (with BoP) provided arguments
Private cllRecentErrors     As Collection
Private bPassOnToEntryProc  As Boolean      ' Used with the debugging option (Conditional Compile Argument Debugging=1)

' Debuggin option buttons (Compile Argument Debugging = 1) -----------------------------------------------
Public Property Get DebugOptResumeErrorLine() As String
    DebugOptResumeErrorLine = "Debugging Option:" & vbLf & vbLf & _
                              "Stop and" & vbLf & _
                              "resume error line"
End Property

Public Property Get DebugOptResumeNext() As String
    DebugOptResumeNext = "Debugging Option:" & vbLf & vbLf & vbLf & _
                         "Resume Next"
End Property

Public Property Get DebugOptPassOnToEntryProc() As String
    DebugOptPassOnToEntryProc = "Debugging Option:" & vbLf & vbLf & _
                                "Pass on error to" & vbLf & _
                                "'Entry-Procedure' *)"
End Property

Public Property Get DebugOptCleanExit() As String
    DebugOptCleanExit = "Debugging Option:" & vbLf & vbLf & _
                        "Clean exit" & vbLf & _
                        "and continue"
End Property
' Debuggin option buttons (Compile Argument Debugging = 1) -----------------------------------------------

Public Property Get ErrMsgDefaultButton() As String:    ErrMsgDefaultButton = "Terminate execution":                                                  End Property

Public Property Get ErrReply() As Variant
    ErrReply = vErrReply
End Property

Public Property Get MostRecentError() As Long
    If Not cllRecentErrors Is Nothing Then
        If cllRecentErrors.Count <> 0 Then
            MostRecentError = cllRecentErrors(cllRecentErrors.Count)
        End If
    End If
End Property

Private Property Let MostRecentError(ByVal lErrNo As Long)
    If cllRecentErrors Is Nothing Then Set cllRecentErrors = New Collection
    cllRecentErrors.Add lErrNo
End Property

Public Property Get RecentErrors() As Collection
    Set RecentErrors = cllRecentErrors
End Property

Private Property Get StckEntryProc() As String
    If Not StckIsEmpty _
    Then StckEntryProc = dctStck.Items()(0) _
    Else StckEntryProc = vbNullString
End Property

Private Function AppErr(ByVal app_err_no As Long) As Long
' ------------------------------------------------------------------------------
' Ensures that a programmed (i.e. an application) error numbers never conflicts
' with the number of a VB runtime error. Thr function returns a given positive
' number (app_err_no) with the vbObjectError added - which turns it into a
' negative value. When the provided number is negative it returns the original
' positive "application" error number e.g. for being used with an error message.
' ------------------------------------------------------------------------------
    AppErr = IIf(app_err_no < 0, app_err_no - vbObjectError, vbObjectError - app_err_no)
End Function

Public Sub BoP(ByVal bop_id As String, _
          ParamArray bop_arguments() As Variant)
' -------------------------------------------------
' Trace and stack Begin of Procedure.
' The traced_arguments argument is passed on to the
' mTrc.BoP and displayed with the error message in
' case.
' -------------------------------------------------
    Const PROC = "BoP"
    
    On Error GoTo eh
    
    If StckIsEmpty Then
        Set vErrsAsserted = Nothing
        Set cllRecentErrors = Nothing: Set cllRecentErrors = New Collection
    End If
    
    StckPush bop_id
#If ExecTrace Then
    vArguments = bop_arguments
    mTrc.BoP_ErH bop_id, vArguments    ' start of the procedure's execution trace
#End If

xt: Exit Sub

eh: MsgBox Err.Description, vbOKOnly, "Error in " & ErrSrc(PROC)
    Stop: Resume
End Sub

Public Sub BoTP(ByVal botp_id As String, _
           ParamArray botp_errs_asserted() As Variant)
' ----------------------------------------------------
' Trace and stack Begin of Procedure and keep a record
' of any asserted errors (error numbers).
' ----------------------------------------------------
    Const PROC = "BoTP"
    
    On Error GoTo eh
    
    mErH.BoP botp_id
    vErrsAsserted = botp_errs_asserted

xt: Exit Sub

eh: MsgBox Err.Description, vbOKOnly, "Error in " & ErrSrc(PROC)
    Stop: Resume
End Sub

Public Sub EoP(ByVal eop_id As String)
' ------------------------------------
' Trace and stack End of Procedure
' ------------------------------------
#If ExecTrace Then
    mTrc.EoP eop_id
#End If
    mErH.StckPop eop_id
End Sub

Private Function ErrArgName(ByVal s As String) As Boolean
    If Right(s, 1) = ":" _
    Or Right(s, 1) = "=" _
    Or Right(s, 2) = ": " _
    Or Right(s, 2) = " :" _
    Or Right(s, 2) = "= " _
    Or Right(s, 2) = " =" _
    Or Right(s, 3) = " : " _
    Or Right(s, 3) = " = " _
    Then ErrArgName = True
End Function

Private Function ErrArgs() As String
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
    va = vArguments
    If Err.Number <> 0 Then Exit Function
    i = LBound(va)
    If Err.Number <> 0 Then Exit Function
    
    For i = i To UBound(va)
        If ErrArgs = vbNullString Then
            ' This is the very first argument
            If ErrArgName(va(i)) Then
                ' The element is the name of an argument followed by a subsequent value
                ErrArgs = va(i) & CStr(va(i + 1))
                i = i + 1
            Else
                sL = ">": sR = "<"
                ErrArgs = "Argument values: " & sL & va(i) & sR
            End If
        Else
            If ErrArgName(va(i)) Then
                ' The element is the name of an argument followed by a subsequent value
                ErrArgs = ErrArgs & ", " & va(i) & CStr(va(i + 1))
                i = i + 1
            Else
                sL = ">": sR = "<"
                ErrArgs = ErrArgs & "  " & sL & va(i) & sR
            End If
        End If
    Next i

End Function

Private Function ErrBttns( _
           ByVal bttns As Variant) As Long
' ----------------------------------------
' Returns the number of specified bttns.
' -----------------------------------------
    Dim v As Variant
    
    For Each v In Split(bttns, ",")
        If IsNumeric(v) Then
            Select Case v
                Case vbOKOnly:                              ErrBttns = ErrBttns + 1
                Case vbOKCancel, vbYesNo, vbRetryCancel:    ErrBttns = ErrBttns + 2
                Case vbAbortRetryIgnore, vbYesNoCancel:     ErrBttns = ErrBttns + 3
            End Select
        Else
            Select Case v
                Case vbNullString, vbLf, vbCr, vbCrLf
                Case Else:  ErrBttns = ErrBttns + 1
            End Select
        End If
    Next v

End Function

Private Function ErrDsply( _
                    ByVal err_source As String, _
                    ByVal err_number As Long, _
                    ByVal err_dscrptn As String, _
                    ByVal err_line As Long, _
           Optional ByVal err_buttons As Variant = vbOKOnly) As Variant
' ------------------------------------------------------------------------------
' Displays the error message. The displayed path to the error may be
' provided as the error is passed on to the 'Entry-Procedure' or based on
' all passed BoP/EoP services. In the first case the path to the error
' may be pretty complete, in the second case the extent of detail
' depends on which (how many) procedures do call the BoP/EoP service.
'
' W. Rauschenberger, Berlin, Nov 2020
' ------------------------------------------------------------------------------
    
    Dim sErrPath    As String
    Dim sTitle      As String
    Dim sLine       As String
    Dim sDetails    As String
    Dim sDscrptn    As String
    Dim sInfo       As String
    Dim sSource     As String
    Dim sType       As String
    Dim lNo         As Long
    Dim ErrMsgText  As TypeMsg
    Dim SctnText    As TypeMsgText
    Dim SctnLabel   As TypeMsgLabel
    
    ErrMsgMatter err_source:=err_source _
               , err_no:=err_number _
               , err_line:=err_line _
               , err_dscrptn:=err_dscrptn _
               , msg_title:=sTitle _
               , msg_line:=sLine _
               , msg_details:=sDetails _
               , msg_source:=sSource _
               , msg_dscrptn:=sDscrptn _
               , msg_info:=sInfo _
               , msg_type:=sType _
               , msg_no:=lNo
#If Debugging Then
    sErrPath = ErrPathErrMsg(msg_details:=sType & lNo & " " & sLine _
                           , err_source:=err_source)
#Else
    If sLine = vbNullString Then sLine = "at line ? *)"
    sErrPath = ErrPathErrMsg(msg_details:=sType & lNo & " " & sLine _
                           , err_source:=err_source)
#End If
    If err_line = 0 Then
        '~~ In case no error line is provided with the error message (commonly the case)
        '~~ a hint regarding the Conditional Compile Argument which may be used to get
        '~~ an option which supports 'resuming' it will be displayed.
#If Debugging Then
        If Not bPassOnToEntryProc Then
            If sInfo <> vbNullString Then sInfo = sInfo & vbLf & vbLf
            sInfo = sInfo & "*) Missing the full path to the error?" & vbLf & _
                            "   Use 'passon on' button pass the error on to the 'Entry-Procedure'!" & vbLf & _
                            "   All passed procedures with an 'mErH.ErrMsg ErrSrc(PROC) ...' will" & vbLf & _
                            "   support gathering the path-to-the-error."
        End If
#Else
        If sInfo <> vbNullString Then sInfo = sInfo & vbLf & vbLf
        sInfo = sInfo & "*) Missing the code line which raised the error? Set the Conditional Compile Argument 'Debugging = 1'." & vbLf & vbLf & _
                        "   The addtionally displayed button <Debugging Option: Stop and resume error line>" & vbLf & _
                        "   and the the error handling:      'If mErH.ErrMsg(ErrSrc(PROC) = mErH.DebugOptResumeErrorLine Then Stop: Resume'" & vbLf & vbLf & _
                        "   will make debugging extremely quick and easy."
#End If
    End If
                       
    '~~ Display the error message via the Common Component procedure mMsg.Dsply
    With ErrMsgText.Section(1)
        With .Label
            .Text = "Error description:"
            .FontColor = rgbBlue
        End With
        .Text.Text = sDscrptn
    End With
    With ErrMsgText.Section(2)
        With .Label
            .Text = "Error source:"
            .FontColor = rgbBlue
        End With
        If ErrArgs = vbNullString _
        Then .Text.Text = sSource & " " & sLine: SctnText.MonoSpaced = True _
        Else .Text.Text = sSource & " " & sLine & vbLf & "(with arguments: " & ErrArgs & ")"
        .Text.MonoSpaced = True
    End With
    With ErrMsgText.Section(3)
        With .Label
            .Text = "Error path (call stack):"
            .FontColor = rgbBlue
        End With
        .Text.Text = sErrPath
        .Text.MonoSpaced = True
    End With
    With ErrMsgText.Section(4)
        If sInfo = vbNullString Then
            .Label.Text = vbNullString
            .Text.Text = vbNullString
        Else
            .Label.Text = "About this error:"
            .Text.Text = sInfo
            .Text.FontSize = 8.5
        End If
        .Label.FontColor = rgbBlue
    End With
    
    mMsg.Dsply dsply_title:=sTitle _
             , dsply_msg:=ErrMsgText _
             , dsply_buttons:=err_buttons
    
    ErrDsply = mMsg.RepliedWith
    
End Function

Private Function ErrHndlrFailed( _
        ByVal err_number As Long, _
        ByVal err_source As String, _
        ByVal err_buttons As Variant) As Boolean
' ------------------------------------------
'
' ------------------------------------------

    If err_number = 0 Then
        MsgBox Prompt:="The error handling has been called with an error number = 0 !" & vbLf & vbLf & _
                       "This indicates that in procedure" & vbLf & _
                       ">>>>> " & err_source & " <<<<<" & vbLf & _
                       "an ""Exit ..."" statement before the call of the error handling is missing!" _
             , Buttons:=vbExclamation _
             , Title:="Exit ... statement missing in " & err_source & "!"
        ErrHndlrFailed = True
        Exit Function
    End If
    
    If IsNumeric(err_buttons) Then
        '~~ When err_buttons is a numeric value, only the VBA MsgBox values for the button argument are supported
        Select Case err_buttons
            Case vbOKOnly, vbOKCancel, vbYesNo, vbRetryCancel, vbYesNoCancel, vbAbortRetryIgnore
            Case Else
                MsgBox "When the errbuttons argument is a numeric value Only the valid VBA MsgBox vaulues are supported. " & _
                       "For valid values please refer to:" & vbLf & _
                       "https://docs.microsoft.com/en-us/office/vba/Language/Reference/User-Interface-Help/msgbox-function" _
                       , vbOKOnly, "Only the valid VBA MsgBox vaulues are supported!"
                ErrHndlrFailed = True
                Exit Function
        End Select
    End If

End Function

Private Function ErrIsAsserted(ByVal err_no As Long) As Boolean
' -------------------------------------------------------------
' Returns TRUE when err_no is an asserted error number, which
' had been provided with BoTP.
' -------------------------------------------------------------
    Dim i As Long
    
    On Error GoTo xt ' no asserted errors provided
    For i = LBound(vErrsAsserted) To UBound(vErrsAsserted)
        If vErrsAsserted(i) = err_no Then
            ErrIsAsserted = True
            Exit Function
        End If
    Next i
xt:
End Function

Public Function ErrMsg( _
                  ByVal err_source As String, _
         Optional ByVal err_number As Long = 0, _
         Optional ByVal err_dscrptn As String = vbNullString, _
         Optional ByVal err_line As Long = 0, _
         Optional ByVal err_buttons As Variant = vbNullString, _
         Optional ByRef err_reply As Variant) As Variant
' ------------------------------------------------------------------------------
' When the buttons (err_buttons) argument specifies more than one button the
' error message is immediately displayed and the users choice is returned to the
' caller, else when the caller (err_source) is the 'Entry-Procedure' the error
' is displayed with the path to the error, else the error is passed on to the
' Entry Procedure whereby the path to the error is composed/assembled.
' ------------------------------------------------------------------------------
    
    Static lInitErrNo       As Long
    Static lInitErrLine     As Long
    Static sInitErrSource   As String
    Static sInitErrDscrptn  As String
    Static sInitErrInfo     As String
    Dim sDetails            As String
    Dim sType               As String
    Dim lNo                 As Long
    Dim sLine               As String
        
    If err_number = 0 Then err_number = Err.Number
    If err_dscrptn = vbNullString Then err_dscrptn = Err.Description
    If err_line = 0 Then err_line = Erl
    
    If ErrHndlrFailed(err_number, err_source, err_buttons) Then GoTo xt
    If cllErrPath Is Nothing Then Set cllErrPath = New Collection
    ErrMsgMatter err_source:=err_source, err_no:=err_number, err_line:=err_line, err_dscrptn:=err_dscrptn, msg_details:=sDetails
    
    If sInitErrSource = vbNullString Then
        '~~ This is the initial/first execution of the error handler within the error raising procedure.
        sInitErrInfo = sDetails
        lInitErrLine = err_line
        lInitErrNo = err_number
        sInitErrSource = err_source
        sInitErrDscrptn = err_dscrptn
        MostRecentError = err_number
        bPassOnToEntryProc = False
    ElseIf err_number <> lInitErrNo _
        And err_number <> lSubsequErrNo _
        And err_source <> sInitErrSource Then
        '~~ In the rare case when the error number had changed during the process of passing it back up to the 'Entry-Procedure'
        lSubsequErrNo = err_number
        sInitErrInfo = sDetails
    End If
    
    MsgManageButtons err_buttons
    ErrMsgMatter err_source:=sInitErrSource, _
                 err_no:=lInitErrNo, _
                 err_line:=lInitErrLine, _
                 err_dscrptn:=sInitErrDscrptn, _
                 msg_type:=sType, _
                 msg_line:=sLine, _
                 msg_no:=lNo
    
    '~~ ---------------------------------------------------------------------------
    '~~ Pass on the error to the 'Entry-Procedure' when
    '~~ - the user has no choice because only a single reply button is displayed
    '~~   and the 'Entry-Procedure' is know but yet not reached
    '~~   or
    '~~ - the debugging option button 'Pass-On-To-Entry-Procedure' had been pressed
    '~~ ---------------------------------------------------------------------------
    If (ErrBttns(err_buttons) = 1 _
        And sErrHndlrEntryProc <> vbNullString _
        And StckEntryProc <> err_source) _
    Or (bPassOnToEntryProc = True _
        And sErrHndlrEntryProc <> vbNullString _
        And StckEntryProc <> err_source) _
    Then
        ErrPathAdd err_source
#If ExecTrace Then
        mTrc.EoP err_source, sType & lNo & " " & sLine
#End If
        mErH.StckPop Itm:=err_source
        sInitErrInfo = vbNullString
        Err.Raise err_number, err_source, err_dscrptn
    End If
       
    '~~ Either when the user has the choice between several replies because of the debugging option
    '~~ or the 'Entry-Procedure' is unknown or has been reached the error is nor displayed
    If (ErrBttns(err_buttons) > 1 _
        Or StckEntryProc = err_source _
        Or StckEntryProc = vbNullString) _
    Or (ErrBttns(err_buttons) > 1 _
        And bPassOnToEntryProc = True _
        And StckEntryProc = err_source) _
    Then
        If Not ErrPathIsEmpty Then ErrPathAdd err_source
        If bPassOnToEntryProc = True Then
            err_buttons = vbOKOnly ' reset to default button because the 'Entry-Procedure' has been reached by an explicit debugging-pass-on
        End If
        '~~ Display the error message
#If ExecTrace Then
    mTrc.Pause
#End If

#If Test Then
        '~~ When the Conditional Compile Argument Test = 1 and the error number is an asserted one
        '~~ the display of the error message is suspended thereby avoiding a user interaction
        If Not ErrIsAsserted(lInitErrNo) _
        Then vErrReply = ErrDsply(err_source:=sInitErrSource, err_number:=lInitErrNo, err_dscrptn:=sInitErrDscrptn, err_line:=lInitErrLine, err_buttons:=err_buttons)
#Else
        vErrReply = ErrDsply(err_source:=sInitErrSource, err_number:=lInitErrNo, err_dscrptn:=sInitErrDscrptn, err_line:=lInitErrLine, err_buttons:=err_buttons)
#End If
        ErrMsg = vErrReply
        err_reply = vErrReply
#If ExecTrace Then
    mTrc.Continue
#End If
        Select Case vErrReply
            Case DebugOptResumeErrorLine, DebugOptResumeNext
            Case DebugOptPassOnToEntryProc
                 bPassOnToEntryProc = True
                 ErrPathAdd err_source
                 Err.Raise lInitErrNo, sInitErrSource, sInitErrDscrptn
            Case DebugOptCleanExit                                      'Goto proc_exit' is with the caller
                 On Error GoTo -1
                 Err.Clear
            Case Else: ErrPathErase
        End Select
#If ExecTrace Then
        mTrc.EoP err_source, sType & lNo & " " & sLine
#End If
        mErH.StckPop Itm:=err_source
        sInitErrInfo = vbNullString
        sInitErrSource = vbNullString
        sInitErrDscrptn = vbNullString
        lInitErrNo = 0
    End If
    
xt: Exit Function
End Function

Private Sub ErrMsgMatter(ByVal err_source As String, _
                         ByVal err_no As Long, _
                         ByVal err_line As Long, _
                         ByVal err_dscrptn As String, _
                Optional ByRef msg_title As String, _
                Optional ByRef msg_type As String, _
                Optional ByRef msg_line As String, _
                Optional ByRef msg_no As Long, _
                Optional ByRef msg_details As String, _
                Optional ByRef msg_dscrptn As String, _
                Optional ByRef msg_info As String, _
                Optional ByRef msg_source As String)
' -------------------------------------------------------
' Returns all the matter to build a proper error message.
' -------------------------------------------------------
                
    If InStr(1, err_source, "DAO") <> 0 _
    Or InStr(1, err_source, "ODBC Teradata Driver") <> 0 _
    Or InStr(1, err_source, "ODBC") <> 0 _
    Or InStr(1, err_source, "Oracle") <> 0 Then
        msg_type = "Database Error "
    Else
      msg_type = IIf(err_no > 0, "VB-Runtime Error ", "Application Error ")
    End If
   
    msg_line = IIf(err_line <> 0, "at line " & err_line, vbNullString)     ' Message error line
    msg_no = IIf(err_no < 0, err_no - vbObjectError, err_no)                ' Message error number
    msg_title = msg_type & msg_no & " in " & err_source & " " & msg_line             ' Message title
    msg_details = IIf(err_line <> 0, msg_type & msg_no & " in " & err_source & " (at line " & err_line & ")", msg_type & msg_no & " in " & err_source)
    msg_dscrptn = IIf(InStr(err_dscrptn, CONCAT) <> 0, Split(err_dscrptn, CONCAT)(0), err_dscrptn)
    If InStr(err_dscrptn, CONCAT) <> 0 Then msg_info = Split(err_dscrptn, CONCAT)(1)
    msg_source = Application.ActiveWindow.Caption & ":  " & err_source
    
End Sub

Private Sub ErrPathAdd(ByVal s As String)
    
    If cllErrorPath Is Nothing Then Set cllErrorPath = New Collection _

    If Not ErrPathItemExists(s) Then
        cllErrorPath.Add s ' avoid duplicate recording of the same procedure/item
    End If
End Sub

Private Sub ErrPathErase()
    Set cllErrorPath = Nothing
End Sub

Private Function ErrPathErrMsg(ByVal msg_details As String, _
                               ByVal err_source) As String
' ------------------------------------------------------------------
' Returns the error path for being displayed in the error message.
' ------------------------------------------------------------------
    
    Dim i   As Long
    Dim j   As Long
    Dim s   As String
    
    ErrPathErrMsg = vbNullString
    If Not ErrPathIsEmpty Then
        '~~ When the error path is not empty and not only contains the error source procedure
        For i = cllErrorPath.Count To 1 Step -1
            s = cllErrorPath(i)
            If i = cllErrorPath.Count _
            Then ErrPathErrMsg = s _
            Else ErrPathErrMsg = ErrPathErrMsg & vbLf & Space$(j * 2) & "|_" & s
            j = j + 1
        Next i
    Else
        '~~ When the error path is empty the stack may provide an alternative information
        If Not StckIsEmpty Then
            For i = 0 To dctStck.Count - 1
                If ErrPathErrMsg <> vbNullString Then
                   ErrPathErrMsg = ErrPathErrMsg & vbLf & Space$((i - 1) * 2) & "|_" & dctStck.Items()(i)
                Else
                   ErrPathErrMsg = dctStck.Items()(i)
                End If
            Next i
        End If
    End If
    ErrPathErrMsg = ErrPathErrMsg & ": " & msg_details
End Function

Private Function ErrPathIsEmpty() As Boolean
    ErrPathIsEmpty = cllErrorPath Is Nothing
    If Not ErrPathIsEmpty Then ErrPathIsEmpty = cllErrorPath.Count = 0
End Function

Private Function ErrPathItemExists(ByVal s As String) As Boolean

    Dim v As Variant
    
    For Each v In cllErrorPath
        If InStr(v & " ", s & " ") <> 0 Then
            ErrPathItemExists = True
            Exit Function
        End If
    Next v
    
End Function

Private Function ErrSrc(ByVal sProc As String) As String
    ErrSrc = "mErH." & sProc
End Function

Private Sub MsgAddButtons(ByRef v1 As Variant, _
                          ByRef v2 As Variant)
' ----------------------------------------------
' Returns v1 followed by v2 whereby both may be
' an msg_buttons argument, i.e. a string, a
' Dictionary or a Collection. When v1 is a
' Dictionary or Collection v2 must be a string
' or long and vice versa.
' ----------------------------------------------
    
    Dim dct As New Dictionary
    Dim cll As New Collection
    Dim v   As Variant
    
    Select Case TypeName(v1)
        Case "Dictionary"
            Select Case TypeName(v2)
                Case "String", "Long": v1.Add v2, v2
                Case Else ' Not added !
            End Select
        Case "Collection"
            Select Case TypeName(v2)
                Case "String", "Long": v1.Add v2
                Case Else ' Not added !
            End Select
        Case "String", "Long"
            Select Case TypeName(v2)
                Case "String"
                    v1 = v1 & "," & v2
                Case "Dictionary"
                    dct.Add v1, v1
                    For Each v In v2
                        dct.Add v, v
                    Next v
                    Set v2 = dct
                Case "Collection"
                    cll.Add v1
                    For Each v In v2
                        cll.Add v
                    Next v
                    Set v2 = cll
            End Select
    End Select
    
End Sub

Private Sub MsgManageButtons(ByRef err_buttons As Variant)

    If err_buttons = vbNullString _
    Then err_buttons = ErrMsgDefaultButton _
    Else MsgAddButtons ErrMsgDefaultButton, err_buttons ' add the default button before the errbuttons specified
    
'~~ Special features are only available with the Alternative VBA MsgBox
#If Debugging Or Test Then
    MsgAddButtons err_buttons, vbLf ' errbuttons in new row
#End If
#If Debugging Then
    MsgAddButtons err_buttons, DebugOptResumeErrorLine
    MsgAddButtons err_buttons, DebugOptResumeNext
    MsgAddButtons err_buttons, DebugOptCleanExit
    MsgAddButtons err_buttons, DebugOptPassOnToEntryProc
#End If

End Sub

Private Function StckBottom() As String
    If Not StckIsEmpty Then StckBottom = dctStck.Items()(0)
End Function

Private Sub StckErase()
    If Not dctStck Is Nothing Then dctStck.RemoveAll
End Sub

Public Function StckIsEmpty() As Boolean
    StckIsEmpty = dctStck Is Nothing
    If Not StckIsEmpty Then StckIsEmpty = dctStck.Count = 0
End Function

Private Function StckPop( _
          Optional ByVal Itm As String = vbNullString) As String
' -----------------------------------------------------------
' Returns the popped of the stack. When itm is provided and
' is not on the top of the stack pop is suspended.
' -----------------------------------------------------------
    Const PROC = "StckPop"
    
    On Error GoTo eh

    If Not StckIsEmpty Then
        If Itm <> vbNullString And StckTop = Itm Then
            StckPop = dctStck.Items()(dctStck.Count - 1) ' Return the poped item
            dctStck.Remove dctStck.Count                  ' Remove item itm from stack
        ElseIf Itm = vbNullString Then
            dctStck.Remove dctStck.Count                  ' Unwind! Remove item itm from stack
        End If
    End If
    
xt: Exit Function

eh: MsgBox Err.Description, vbOKOnly, "Error in " & ErrSrc(PROC)
End Function

Private Sub StckPush(ByVal s As String)

    If dctStck Is Nothing Then Set dctStck = New Dictionary
    If dctStck.Count = 0 Then
        sErrHndlrEntryProc = s ' First pushed = bottom item = 'Entry-Procedure'
#If ExecTrace Then
        mTrc.Terminate ' ensures any previous trace is erased
#End If
    End If
    dctStck.Add dctStck.Count + 1, s

End Sub

Private Function StckTop() As String
    If Not StckIsEmpty Then StckTop = dctStck.Items()(dctStck.Count - 1)
End Function

