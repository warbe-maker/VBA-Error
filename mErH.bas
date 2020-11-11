Attribute VB_Name = "mErH"
Option Explicit
Option Private Module
' -----------------------------------------------------------------------------------------------
' Standard  Module mErrHndlr: Global error handling for any VBA Project.
'
' Methods: - AppErr   Converts a positive number into a negative error number ensuring it not
'                     conflicts with a VB Runtime Error. A negative error number is turned back into the
'                     original positive Application  Error Number.
'          - ErrMsg Either passes on the error to the caller or when the entry procedure is
'                     reached, displays the error with a complete path from the entry procedure
'                     to the procedure with the error.
'          - BoP      Maintains the call stack at the Begin of a Procedure (optional when using
'                     this common error handler)
'          - EoP      Maintains the call stack at the End of a Procedure, triggers the display of
'                     the Execution Trace when the entry procedure is finished and the
'                     Conditional Compile Argument ExecTrace = 1
'          - ErrDsply Displays the error message in a proper formated manner
'                     The local Conditional Compile Argument "AlternativeMsgBox = 1" enforces the use
'                     of the Alternative VBA MsgBox which provideds an improved readability.
'
' Usage:   Private/Public Sub/Function any()
'              Const PROC = "any"  ' procedure's name as error source
'
'              On Error GoTo eh
'              mErH.BoP ErrSrc(PROC)   ' puts the procedure on the call stack
'
'              ' <any code>
'
'          xt: ' <any "finally" code like re-protecting an unprotected sheet for instance>
'                               mErH.EoP ErrSrc(PROC)   ' takes the procedure off from the call stack
'                               Exit Sub/Function
'
'           eh: mErH.ErrMsg Err.Number, ErrSrc(PROC), Err.Description, Erl
'           End ....
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
'          For further details see the Github blog post
'          "A comprehensive common VBA Error Handler inspired by the best of the web"
' https://warbe-maker.github.io/vba/common/2020/10/02/Comprehensive-Common-VBA-Error-Handler.html
'
' W. Rauschenberger, Berlin, Nov 2020
' -----------------------------------------------------------------------------------------------

Public Const CONCAT         As String = "||"

Public Enum StartupPosition         ' ---------------------------
    Manual = 0                      ' Used to position the
    CenterOwner = 1                 ' final setup message form
    CenterScreen = 2                ' horizontally and vertically
    WindowsDefault = 3              ' centered on the screen
End Enum                            ' ---------------------------

Public Type tSection                ' ------------------
       sLabel As String             ' Structure of the
       sText As String              ' UserForm's
       bMonspaced As Boolean        ' message area which
End Type                            ' consists of
Public Type tMessage                ' three message
       section(1 To 3) As tSection  ' sections
End Type                            ' -------------------

Private cllErrPath          As Collection
Private cllErrorPath        As Collection   ' managed by ErrPath... procedures exclusively
Private dctStck             As Dictionary
Private sErrHndlrEntryProc  As String
Private lSubsequErrNo       As Long ' a number possibly different from lInitialErrNo when it changes when passed on to the Entry Procedure
Private lInitialErrLine     As Long
Private lInitialErrNo       As Long
Private sInitialErrSource   As String
Private sInitialErrDscrptn  As String
Private sInitialErrInfo     As String

' Test button, displayed with Conditional Compile Argument Test = 1
Public Property Get ExitAndContinue() As String:        ExitAndContinue = "Exit procedure" & vbLf & "and continue" & vbLf & "with next":    End Property

' Debugging button, displayed with Conditional Compile Argument Debugging = 1
Public Property Get ResumeError() As String:            ResumeError = "Resume" & vbLf & "error code line":                                  End Property

' Test button, displayed with Conditional Compile Argument Test = 1
Public Property Get ResumeNext() As String:             ResumeNext = "Continue with code line" & vbLf & "following the error line":         End Property

' Default error message button
Public Property Get ErrMsgDefaultButton() As String:    ErrMsgDefaultButton = "Terminate execution":                                                  End Property

Private Property Get StckEntryProc() As String
    If Not StckIsEmpty _
    Then StckEntryProc = dctStck.Items()(0) _
    Else StckEntryProc = vbNullString
End Property

Public Function AppErr(ByVal errno As Long) As Long
' -----------------------------------------------------------------
' Used with Err.Raise AppErr(<l>).
' When the error number <l> is > 0 it is considered an "Application
' Error Number and vbObjectErrror is added to it into a negative
' number in order not to confuse with a VB runtime error.
' When the error number <l> is negative it is considered an
' Application Error and vbObjectError is added to convert it back
' into its origin positive number.
' ------------------------------------------------------------------
    If errno < 0 Then
        AppErr = errno - vbObjectError
    Else
        AppErr = vbObjectError + errno
    End If
End Function

Public Sub BoP(ByVal s As String)
' ----------------------------------
' Trace and stack Begin of Procedure
' ----------------------------------
    Const PROC = "BoP"
    
    On Error GoTo eh
    
    StckPush s
#If ExecTrace Then
    mTrc.BoP s    ' start of the procedure's execution trace
#End If

xt: Exit Sub

eh: MsgBox Err.Description, vbOKOnly, "Error in " & ErrSrc(PROC)
    Stop: Resume
End Sub

Public Sub EoP(ByVal s As String)
' --------------------------------
' Trace and stack End of Procedure
' --------------------------------
#If ExecTrace Then
    mTrc.EoP s
    mErH.StckPop s
#End If
End Sub

Public Function ErrMsg(ByVal errnumber As Long, _
                         ByVal errsource As String, _
                         ByVal errdscrptn As String, _
                         ByVal errline As Long, _
                Optional ByVal errbuttons As Variant = vbNullString) As Variant
' ----------------------------------------------------------------------
' When the errbuttons argument specifies more than one button the error
' message is immediately displayed and the users choice is returned,
' else when the caller (errsource) is the "Entry Procedure" the error
' is displayed with the path to the error,
' else the error is passed on to the "Entry Procedure" whereby the
' .ErrorPath string is assebled.
' ----------------------------------------------------------------------
    
    Static sLine    As String   ' provided error line (if any) for the the finally displayed message
    Dim sDetails    As String
    
    If ErrHndlrFailed(errnumber, errsource, errbuttons) Then Exit Function
    If cllErrPath Is Nothing Then Set cllErrPath = New Collection
    If errline <> 0 Then sLine = errline Else sLine = "0"
    ErrHndlrManageButtons errbuttons
    ErrMsgMatter errsource:=errsource, errno:=errnumber, errline:=errline, errdscrptn:=errdscrptn, msgdetails:=sDetails
    
    If sInitialErrSource = vbNullString Then
        '~~ This is the initial/first execution of the error handler within the error raising procedure.
        sInitialErrInfo = sDetails
        lInitialErrLine = errline
        lInitialErrNo = errnumber
        sInitialErrSource = errsource
        sInitialErrDscrptn = errdscrptn
    ElseIf errnumber <> lInitialErrNo _
        And errnumber <> lSubsequErrNo _
        And errsource <> sInitialErrSource Then
        '~~ In the rare case when the error number had changed during the process of passing it back up to the entry procedure
        lSubsequErrNo = errnumber
        sInitialErrInfo = sDetails
    End If
    
    If ErrBttns(errbuttons) = 1 _
    And sErrHndlrEntryProc <> vbNullString _
    And StckEntryProc <> errsource Then
        '~~ When the user has no choice to press any button but the only one displayed button
        '~~ and the Entry Procedure is known but yet not reached the error is passed on back
        '~~ up to the Entry Procedure whereupon the path to the error is assembled
        ErrPathAdd errsource
#If ExecTrace Then
        mTrc.EoP errsource, sInitialErrInfo
#End If
        mErH.StckPop Itm:=errsource
        sInitialErrInfo = vbNullString
        Err.Raise errnumber, errsource, errdscrptn
    End If
    
    If ErrBttns(errbuttons) > 1 _
    Or StckEntryProc = errsource _
    Or StckEntryProc = vbNullString Then
        '~~ When the user has the choice between several errbuttons displayed
        '~~ or the Entry Procedure is unknown or has been reached
        If Not ErrPathIsEmpty Then ErrPathAdd errsource
        '~~ Display the error message
        ErrMsg = ErrDsply(errnumber:=lInitialErrNo, errline:=lInitialErrLine, errbuttons:=errbuttons)
        Select Case ErrMsg
            Case ResumeError, ResumeNext, ExitAndContinue
            Case Else: ErrPathErase
        End Select
#If ExecTrace Then
        mTrc.EoP errsource, sInitialErrInfo
#End If
        mErH.StckPop Itm:=errsource
        sInitialErrInfo = vbNullString
        sInitialErrSource = vbNullString
        sInitialErrDscrptn = vbNullString
        lInitialErrNo = 0
    End If
    
    '~~ Each time a known Entry Procedure is reached the execution trace
    '~~ maintained by the BoP and mErH.EoP and the BoC and EoC statements is displayed
    If StckEntryProc = errsource _
    Or StckEntryProc = vbNullString Then
        Select Case ErrMsg
            Case ResumeError, ResumeNext, ExitAndContinue
            Case vbOK
            Case Else: StckErase
        End Select
    End If
    mErH.StckPop errsource
    
End Function

Private Sub ErrHndlrManageButtons(ByRef errbuttons As Variant)

    If errbuttons = vbNullString _
    Then errbuttons = ErrMsgDefaultButton _
    Else ErrHndlrAddButtons ErrMsgDefaultButton, errbuttons ' add the default button before the errbuttons specified
    
'~~ Special features are only available with the Alternative VBA MsgBox
#If Debugging Or Test Then
    ErrHndlrAddButtons errbuttons, vbLf ' errbuttons in new row
#End If
#If Debugging Then
    ErrHndlrAddButtons errbuttons, ResumeError
#End If
#If Test Then
     ErrHndlrAddButtons errbuttons, ResumeNext
     ErrHndlrAddButtons errbuttons, ExitAndContinue
#End If

End Sub
Private Function ErrHndlrFailed( _
        ByVal errnumber As Long, _
        ByVal errsource As String, _
        ByVal errbuttons As Variant) As Boolean
' ------------------------------------------
'
' ------------------------------------------

    If errnumber = 0 Then
        MsgBox "The error handling has been called with an error number = 0 !" & vbLf & vbLf & _
               "This indicates that in procedure" & vbLf & _
               ">>>>> " & errsource & " <<<<<" & vbLf & _
               "an ""Exit ..."" statement before the call of the error handling is missing!" _
               , vbExclamation, _
               "Exit ... statement missing in " & errsource & "!"
                ErrHndlrFailed = True
        Exit Function
    End If
    
    If IsNumeric(errbuttons) Then
        '~~ When errbuttons is a numeric value, only the VBA MsgBox values for the button argument are supported
        Select Case errbuttons
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

Private Sub ErrHndlrAddButtons(ByRef v1 As Variant, _
                               ByRef v2 As Variant)
' ---------------------------------------------------
' Returns v1 followed by v2 whereby both may be a
' errbuttons argument which means  a string, a
' Dictionary or a Collection. When v1 is a Dictionary
' or Collection v2 must be a string or long and vice
' versa.
' ---------------------------------------------------
    
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

Public Function ErrDsply( _
                ByVal errnumber As Long, _
                ByVal errline As Long, _
       Optional ByVal errbuttons As Variant = vbOKOnly) As Variant
' -------------------------------------------------------------
' Displays the error message either by means of VBA MsgBox or,
' when the Conditional Compile Argument AlternativeMsgBox = 1 by
' means of the Alternative VBA MsgBox (UserForm fMsg). In any
' case the path to the error may be displayed, provided the
' entry procedure has BoP/EoP code lines.
'
' W. Rauschenberger, Berlin, Sept 2020
' -------------------------------------------------------------
    
    Dim sErrPath    As String
    Dim sTitle      As String
    Dim sErrLine    As String
    Dim sDetails    As String
    Dim sDscrptn    As String
    Dim sInfo       As String
    
    ErrMsgMatter errsource:=sInitialErrSource, errno:=errnumber, errline:=errline, errdscrptn:=sInitialErrDscrptn, msgtitle:=sTitle, msgline:=sErrLine, msgdetails:=sDetails
    sErrPath = ErrPathErrMsg(sDetails)
    '~~ Display the error message by means of the Common UserForm fMsg
    With fMsg
        .msgtitle = sTitle
        .MsgLabel(1) = "Error description:":            .MsgText(1) = sDscrptn
        If Not ErrPathIsEmpty Then
            .MsgLabel(2) = "Error path (call stack):":  .MsgText(2) = sErrPath:   .MsgMonoSpaced(2) = True
        Else
            .MsgLabel(2) = "Error source:":             .MsgText(2) = sInitialErrSource & sErrLine
        End If
        If sInfo <> vbNullString Then
            .MsgLabel(3) = "Info:":                     .MsgText(3) = sInfo
        End If
        .MsgButtons = errbuttons
        .Setup
        .Show
        If ErrBttns(errbuttons) = 1 Then
            ErrDsply = errbuttons ' a single reply errbuttons return value cannot be obtained since the form is unloaded with its click
        Else
            ErrDsply = .ReplyValue ' when more than one button is displayed the form is unloadhen the return value is obtained
        End If
    End With

End Function

Private Function ErrBttns( _
                 ByVal bttns As Variant) As Long
' ------------------------------------------------
' Returns the number of specified bttns.
' ------------------------------------------------
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

Private Sub ErrMsgMatter(ByVal errsource As String, _
                         ByVal errno As Long, _
                         ByVal errline As Long, _
                         ByVal errdscrptn As String, _
                 Optional ByRef msgtitle As String, _
                 Optional ByRef msgtype As String, _
                 Optional ByRef msgline As String, _
                 Optional ByRef msgno As Long, _
                 Optional ByRef msgdetails As String, _
                 Optional ByRef msgdscrptn As String, _
                 Optional ByRef msginfo As String)
' -------------------------------------------------------
' Returns all the matter to build a proper error message.
' -------------------------------------------------------
                
    If InStr(1, errsource, "DAO") <> 0 _
    Or InStr(1, errsource, "ODBC Teradata Driver") <> 0 _
    Or InStr(1, errsource, "ODBC") <> 0 _
    Or InStr(1, errsource, "Oracle") <> 0 Then
        msgtype = "Database Error "
    Else
      msgtype = IIf(errno > 0, "VB-Runtime Error ", "Application Error ")
    End If
   
    msgline = IIf(errline <> 0, "at line " & errline, vbNullString)     ' Message error line
    msgno = IIf(errno < 0, errno - vbObjectError, errno)                ' Message error number
    msgtitle = msgtype & msgno & " in " & errsource & " " & msgline             ' Message title
    msgdetails = IIf(errline <> 0, msgtype & msgno & " in " & errsource & " (at line " & errline & ")", msgtype & msgno & " in " & errsource)
    msgdscrptn = IIf(InStr(errdscrptn, CONCAT) <> 0, Split(errdscrptn, CONCAT)(0), errdscrptn)
    If InStr(errdscrptn, CONCAT) <> 0 Then msginfo = Split(errdscrptn, CONCAT)(1)

End Sub

Private Sub ErrPathAdd(ByVal s As String)
    
    If cllErrorPath Is Nothing Then Set cllErrorPath = New Collection _

    If Not ErrPathItemExists(s) Then
        Debug.Print s & " added to path"
        cllErrorPath.Add s ' avoid duplicate recording of the same procedure/item
    End If
End Sub

Private Function ErrPathItemExists(ByVal s As String) As Boolean

    Dim v As Variant
    
    For Each v In cllErrorPath
        If InStr(v & " ", s & " ") <> 0 Then
            ErrPathItemExists = True
            Exit Function
        End If
    Next v
    
End Function

Private Sub ErrPathErase()
    Set cllErrorPath = Nothing
End Sub

Private Function ErrPathErrMsg(ByVal msgdetails As String) As String
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
            s = cllErrorPath.TrcEntryItem(i)
            If i = cllErrorPath.Count _
            Then ErrPathErrMsg = s _
            Else ErrPathErrMsg = ErrPathErrMsg & vbLf & Space(j * 2) & "|_" & s
            j = j + 1
        Next i
    End If
    ErrPathErrMsg = ErrPathErrMsg & vbLf & Space(j * 2) & "|_" & sInitialErrSource & " " & msgdetails
End Function

Private Function ErrPathIsEmpty() As Boolean
    ErrPathIsEmpty = cllErrorPath Is Nothing
    If Not ErrPathIsEmpty Then ErrPathIsEmpty = cllErrorPath.Count = 0
End Function

Private Function ErrSrc(ByVal sProc As String) As String
    ErrSrc = "mErH." & sProc
End Function

Public Function Space(ByVal l As Long) As String
' --------------------------------------------------
' Unifies the VB differences SPACE$ and Space$ which
' lead to code diferences where there aren't any.
' --------------------------------------------------
    Space = VBA.Space$(l)
End Function

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
        sErrHndlrEntryProc = s ' First pushed = bottom item = entry procedure
#If ExecTrace Then
        mTrc.Terminate ' ensures any previous trace is erased
#End If
    End If
    dctStck.Add dctStck.Count + 1, s

End Sub

Private Function StckTop() As String
    If Not StckIsEmpty Then StckTop = dctStck.Items()(dctStck.Count - 1)
End Function

