Attribute VB_Name = "mMsg"
Option Explicit
' ----------------------------------------------------------------------------------
' Standard Module mMsg  Interface for the Common VBA Message Service (fMsg UserForm)
'
' Public services:
' - Dsply               Exposes all properties and methods for the display of any
'                       kind of message
' - Box                 In analogy to the MsgBox, provides a simple message but with
'                       all the fexibility for the display of up to 49 reply buttons
' - Buttons             Supports the specification of the design buttons displayed
'                       in 7 rows by 7 buttons each
'
' See details at to:
' https://warbe-maker.github.io/warbe-maker.github.io/vba/common/2020/11/17/Common-VBA-Message-Services.html
'
' W. Rauschenberger, Berlin Jan 2021 (last revision)
' ----------------------------------------------------------------------------------
Public Const END_OF_PROGRESS As String = "EndOfProgress"

Public ProgressText As String

Public Type TypeMsgLabel
        FontBold As Boolean
        FontColor As XlRgbColor
        FontItalic As Boolean
        FontName As String
        FontSize As Long
        FontUnderline As Boolean
        MonoSpaced As Boolean ' overwrites any FontName
        Text As String
End Type
Public Type TypeMsgText
        FontBold As Boolean
        FontColor As XlRgbColor
        FontItalic As Boolean
        FontName As String
        FontSize As Long
        FontUnderline As Boolean
        MonoSpaced As Boolean ' overwrites any FontName
        Text As String
End Type
Public Type TypeMsgSect
       Label As TypeMsgLabel
       Text As TypeMsgText
End Type
Public Type TypeMsg
    Section(1 To 4) As TypeMsgSect
End Type

Private bModeless       As Boolean
Public DisplayDone      As Boolean
Public RepliedWith      As Variant
Private MsgForms        As Dictionary   ' Collection of active form instances with their caption as the key

Public Function Form(Optional ByVal frm_caption As String) As fMsg
' -------------------------------------------------------------------------
' When an fMsg instance, identified by the caption (frm_caption), exists in
' the MsgForms Dictionary and this instance still exists it is returned,
' else the no longer existing instance is removed from the dictionary, a
' new one is created, added, and returned.
' -------------------------------------------------------------------------
    Const PROC = "Form"
    
    On Error GoTo eh
    Dim MsgForm As fMsg

    If MsgForms Is Nothing Then Set MsgForms = New Dictionary
    If Not MsgForms.Exists(frm_caption) Then
        Set MsgForm = New fMsg
        MsgForms.Add frm_caption, MsgForm
    Else
        On Error Resume Next
        Set MsgForm = MsgForms(frm_caption)
        Select Case Err.Number
            Case 0
            Case 13
                '~~ The fMsg instance no longer exists
                MsgForms.Remove frm_caption
                Set MsgForm = New fMsg
                MsgForms.Add frm_caption, MsgForm
            Case Else
                '~~ Unknown error!
                Err.Raise AppErr(1), ErrSrc(PROC), "Unknown/unrecognized error!"
        End Select
        On Error Resume Next
        
    End If

xt: Set Form = MsgForm
    Exit Function

eh: ErrMsg ErrSrc(PROC)
End Function

Public Property Get Modeless() As Boolean:          Modeless = bModeless:   End Property
Public Property Let Modeless(ByVal b As Boolean):   bModeless = b:          End Property

Private Function Max(ParamArray va() As Variant) As Variant
' --------------------------------------------------------
' Returns the maximum value of all values provided (va).
' --------------------------------------------------------
    Dim v As Variant
    
    Max = va(LBound(va)): If LBound(va) = UBound(va) Then Exit Function
    For Each v In va
        If v > Max Then Max = v
    Next v
    
End Function

Public Function Progress( _
                   ByVal prgrs_title As String, _
                   ByRef prgrs_msg As String, _
          Optional ByVal prgrs_header As String = vbNullString, _
          Optional ByVal prgrs_footer As String = "Process in progress! Please wait.", _
          Optional ByVal prgrs_msg_append As Boolean = True, _
          Optional ByVal prgrs_msg_monospaced As Boolean = False, _
          Optional ByVal prgrs_min_width As Long = 400, _
          Optional ByVal prgrs_max_width As Long = 80, _
          Optional ByVal prgrs_max_height As Long = 70) As Variant
' -------------------------------------------------------------------------------------
' Progress display using an instance of the Common VBA Message Form. When no instance
' for the provided title (prgrs_title) exists one is created, else the existing
' instance is used.
'
' See: https://warbe-maker.github.io/vba/common/2020/11/17/Common-VBA-Message-Form.html
'
' W. Rauschenberger, Berlin, May 2021
' -------------------------------------------------------------------------------------
    Const PROC = "Progress"
   
    On Error GoTo eh
    Dim Msg     As TypeMsg
    Dim MsgForm As fMsg

    Set MsgForm = Form(prgrs_title)
    With Msg.Section(1)
        With .Label
            .Text = prgrs_header
            .MonoSpaced = prgrs_msg_monospaced
            .FontBold = True
        End With
        With .Text
            .Text = prgrs_msg
            .MonoSpaced = prgrs_msg_monospaced
        End With
    End With
    With Msg.Section(2).Text
        .Text = prgrs_footer
        .FontColor = rgbBlue
        .FontSize = 8
        .FontBold = True
    End With
    
    If Trim(MsgForm.MsgTitle) <> Trim(prgrs_title) Then
        With MsgForm
            '~~ A new title starts a new progress message
            .DsplyFrmsWthCptnTestOnly = False
            .MsgHeightMaxSpecAsPoSS = prgrs_max_height ' percentage of screen height
            .MsgWidthMaxSpecAsPoSS = prgrs_max_width   ' percentage of screen width
            .MsgWidthMaxSpecInPt = prgrs_min_width     ' defaults to 300 pt. the absolute minimum is 200 pt
            .MsgTitle = prgrs_title
            .MsgLabel(1) = Msg.Section(1).Label
            .MsgText(1) = Msg.Section(1).Text
            .MsgText(2) = Msg.Section(2).Text
            .MsgButtons = vbNullString
            .ProgressFollowUp = True

            '+------------------------------------------------------------------------+
            '|| Setup prior showing the form improves the performance significantly  ||
            '|| and avoids any flickering message window with its setup.             ||
            '|| For testing purpose it may be appropriate to out-comment the Setup.  ||
            .Setup '                                                                 ||
            '+------------------------------------------------------------------------+
            .show vbModeless
            GoTo xt
        End With
    Else
        '~~ Another progress message with the same title is appended or relpaces the message in the provided section
        Application.ScreenUpdating = False
        MsgForm.Progress prgrs_text:=prgrs_msg _
                   , prgrs_append:=prgrs_msg_append _
                   , prgrs_footer:=Msg.Section(2).Text.Text
    End If
      
xt: Exit Function

eh: ErrMsg ErrSrc(PROC)
End Function

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

Public Function Box(ByVal box_title As String, _
           Optional ByVal box_msg As String = vbNullString, _
           Optional ByVal box_monospaced As Boolean = False, _
           Optional ByVal box_buttons As Variant = vbOKOnly, _
           Optional ByVal box_button_default = 1, _
           Optional ByVal box_returnindex As Boolean = False, _
           Optional ByVal box_min_width As Long = 400, _
           Optional ByVal box_max_width As Long = 80, _
           Optional ByVal box_max_height As Long = 70, _
           Optional ByVal box_min_button_width = 70) As Variant
' -------------------------------------------------------------------------------------
' Common VBA Message Display: A service using the Common VBA Message Form as an
' alternative MsgBox.
' Please Note: This Box service is a kind of backward compatibility with the VBA.MsgBox
'              with equivalent arguments:      VBA.MsgBox | mMsg.Box
'                                              ---------- + ------------------------
'                                              Title      | box_title
'                                              Prompt     | box_msg
'                                              Buttons    | box_buttons
'              and explicit                               | box_button_default
'              and some additional arguments concerning the message size.
'
' See: https://warbe-maker.github.io/vba/common/2020/11/17/Common-VBA-Message-Form.html
'
' W. Rauschenberger, Berlin, Nov 2020
' -------------------------------------------------------------------------------------
    Dim Message As TypeMsgText
    
    Message.Text = box_msg
    Message.MonoSpaced = box_monospaced
    
    With fMsg
        .MsgHeightMaxSpecAsPoSS = box_max_height    ' percentage of screen height
        .MsgWidthMaxSpecAsPoSS = box_max_width      ' percentage of screen width
        .MsgWidthMaxSpecInPt = box_min_width        ' defaults to 300 pt. the absolute minimum is 200 pt
        .MinButtonWidth = box_min_button_width
        .MsgTitle = box_title
        .MsgText(1) = Message
        .MsgButtons = box_buttons
        .DefaultBttn = box_button_default
        '+------------------------------------------------------------------------+
        '|| Setup prior showing the form improves the performance significantly  ||
        '|| and avoids any flickering message window with its setup.             ||
        '|| For testing purpose it may be appropriate to out-comment the Setup.  ||
        .Setup '                                                                 ||
        '+------------------------------------------------------------------------+
        .show
    End With
    Box = RepliedWith

End Function

Public Function ButtonsString(ByVal msg_buttons As Variant) As String
' ------------------------------------------------------------------------------
' Returns the button captions (msg_buttons) which may be provided as komma
' delimited string, array, collection, or Dictionary, as komma delimited string.
' ------------------------------------------------------------------------------
    
    On Error GoTo eh
    Dim v As Variant
    
    Debug.Print TypeName(msg_buttons)
    
    If IsArray(msg_buttons) Then
        ButtonsString = Join(msg_buttons, ",")
    Else
        Select Case VarType(msg_buttons)
            Case vbArray
                ButtonsString = Split(msg_buttons, ",")
            Case vbString
            Case Else
                Select Case TypeName(msg_buttons)
                    Case "Dictionary"
                    Case "Collection"
                        For Each v In msg_buttons
                            ButtonsString = v & ","
                        Next v
                        ButtonsString = Left(ButtonsString, Len(ButtonsString) - 1)
                End Select
        End Select
    End If

xt: Exit Function
eh: Debug.Print Err.Description: Stop: Resume
End Function

Public Function ButtonsArray(ByVal msg_buttons As Variant) As Variant
' ------------------------------------------------------------------------------
' Returns the button captions (msg_buttons) which may be provided as komma
' delimited string, array, collection, or Dictionary, as komma delimited string.
' ------------------------------------------------------------------------------
    
    Dim va()    As Variant
    Dim i       As Long
    Dim dct     As Dictionary
    Dim cll     As Collection
    
    Debug.Print TypeName(msg_buttons)
    Select Case VarType(msg_buttons)
        Case vbArray:   ButtonsArray = msg_buttons
        Case vbString: ButtonsArray = Split(msg_buttons, ",")
        Case Else
            Select Case TypeName(msg_buttons)
                Case "Dictionary"
                    Set dct = msg_buttons
                    ReDim va(dct.Count - 1)
                    For i = 0 To dct.Count - 1
                        va(i) = dct.Items()(i)
                    Next i
                    ButtonsArray = va
                Case "Collection"
                    Set cll = msg_buttons
                    ReDim va(cll.Count - 1)
                    For i = 0 To cll.Count - 1
                        va(i) = cll.Item(i + 1)
                    Next i
                    ButtonsArray = va
            End Select
    End Select

End Function

Public Sub ButtonsAdd(ByVal msg_buttons As Variant, _
                      ByRef to_collection As Collection)
' --------------------------------------------------------------------------
' Adds the buttons (msg_buttons) - provided either as a comma delimited
' string, an array, a Collection, or a Dictionary to the Collection
' (to_collection). Allows to concatenate bunches of button captions.
' --------------------------------------------------------------------------
    Dim arry()  As Variant
    Dim cll     As Collection
    Dim dct     As Dictionary
    Dim i       As Long
    
    If IsArray(msg_buttons) Then
        For i = LBound(msg_buttons) To UBound(msg_buttons)
            to_collection.Add msg_buttons(i)
        Next i
    Else
        Select Case TypeName(msg_buttons)
            Case "String"
                ButtonsAdd Split(msg_buttons, ","), to_collection ' call recursively with the array as argument
            Case "Collection"
                Set cll = msg_buttons
                ReDim arry(cll.Count - 1)
                For i = 1 To cll.Count
                    arry(i - 1) = cll.Item(i)
                Next i
                ButtonsAdd arry, to_collection ' call recursively with the array as argument
            Case "Dictionary"
                Set dct = msg_buttons
                ReDim arry(dct.Count - 1)
                For i = 1 To dct.Count
                    arry(i - 1) = dct.Items()(i - 1)
                Next i
                ButtonsAdd arry, to_collection ' call recursively with the array as argument
        End Select
    End If
End Sub

Public Function Buttons(ParamArray msg_buttons() As Variant) As Collection
' --------------------------------------------------------------------------
' Returns a collection of the items provided (msg_buttons). When more
' than 7 items are provided the function adds a button row break.
' The function considers a possible kind of mistake when the ParamArray
' contains only one item which is a comma delimited string.
' So instead of 3 argument "A", "B", "C" only one "A,B,C" is provided.
' --------------------------------------------------------------------------
    
    Dim cll     As New Collection
    Dim i       As Long
    Dim j       As Long         ' buttons in a row counter
    Dim k       As Long: k = 1  ' button rows counter
    Dim l       As Long         ' total buttons count
    Dim va1     As Variant      ' array of button captions from a comma delimeted string
    Dim va2()   As Variant      ' array of button captions either from va1 or from msg_butttons
    Dim s       As String
    
    On Error Resume Next
    i = LBound(msg_buttons)
    If Err.Number <> 0 Then GoTo xt
    
    '~~ Transpose the the buttons argument into an array considering that
    '~~  the ParaArray may contain only one comma delimited string.
    If LBound(msg_buttons) = UBound(msg_buttons) Then
        s = msg_buttons(LBound(msg_buttons))
        va1 = Split(s, ",")
        ReDim va2(UBound(va1))
        For i = LBound(va1) To UBound(va1)
            va2(i) = va1(i)
        Next i
    Else
        ReDim va2(UBound(msg_buttons))
        For i = LBound(msg_buttons) To UBound(msg_buttons)
            va2(i) = msg_buttons(i)
        Next i
    End If
    
    '~~ Return the array (va2) as Collection
    For i = LBound(va2) To UBound(va2)
        If VarType(va2(i)) = vbEmpty Then GoTo nxt
        If (k = 7 And j = 7) Or l = 49 Then GoTo xt
        Select Case va2(i)
            Case vbLf, vbCrLf, vbCr
                cll.Add va2(i)
                j = 0
                k = k + 1
            Case vbOKOnly, vbOKCancel, vbAbortRetryIgnore, vbYesNoCancel, vbYesNo, vbRetryCancel
                If j = 7 Then
                    cll.Add vbLf
                    j = 0
                    k = k + 1
                End If
                cll.Add va2(i)
                j = j + 1
                l = l + 1
            Case Else
                If TypeName(va2(i)) = "String" Then
                    ' Any invalid buttons value will be ignored without notice
                    If j = 7 Then
                        cll.Add vbLf
                        j = 0
                        k = k + 1
                    End If
                    cll.Add va2(i)
                    j = j + 1
                    l = l + 1
                End If
        End Select
nxt: Next i
    
xt: Set Buttons = cll

End Function
                                     
Public Function Dsply(ByVal dsply_title As String, _
                      ByRef dsply_msg As TypeMsg, _
             Optional ByVal dsply_buttons As Variant = vbOKOnly, _
             Optional ByVal dsply_button_default = 1, _
             Optional ByVal dsply_reply_with_index As Boolean = False, _
             Optional ByVal dsply_modeless As Boolean = False, _
             Optional ByVal dsply_min_width As Long = 0, _
             Optional ByVal dsply_max_width As Long = 0, _
             Optional ByVal dsply_max_height As Long = 0, _
             Optional ByVal dsply_min_button_width = 0) As Variant
' ------------------------------------------------------------------------------
' Common VBA Message Display: A service using the Common VBA Message Form as an
' alternative to the VBA.MsgBox.
'
' Argument               | Description
' ---------------------- + ----------------------------------------------------
' dsply_title            | String, Title
' dsply_msg              | UDT, Message
' dsply_buttons          | Button captions as Collection
' dsply_button_default   | Default button, either the index or the caption,
'                        | defaults to 1 (= the first displayed button)
' dsply_reply_with_index | Defaults to False, when True the index of the
'                        | of the pressed button is returned else the caption
'                        | or the VBA.MsgBox button value respectively
' dsply_modeless         | The message is displayed modeless, defaults to False
'                        | = vbModal
' dsply_min_width        | Overwrites the default when not 0
' dsply_max_width        | Overwrites the default when not 0
' dsply_max_height       | Overwrites the default when not 0
' dsply_min_button_width | Overwrites the default when not 0
'
' See: https://github.com/warbe-maker/Common-VBA-Message-Service
'
' W. Rauschenberger, Berlin, Nov 2020
' -------------------------------------------------------------------------------------
    Const PROC = "Dsply"
    
    On Error GoTo eh
    Dim i       As Long
    
    With fMsg
        .ReplyWithIndex = dsply_reply_with_index
        If dsply_max_height > 0 Then .MsgHeightMaxSpecAsPoSS = dsply_max_height ' percentage of screen height
        If dsply_max_width > 0 Then .MsgWidthMaxSpecAsPoSS = dsply_max_width   ' percentage of screen width
        If dsply_min_width > 0 Then .MsgWidthMinSpecInPt = dsply_min_width                     ' defaults to 300 pt. the absolute minimum is 200 pt
        If dsply_min_button_width > 0 Then .MinButtonWidth = dsply_min_button_width
        .MsgTitle = dsply_title
        For i = 1 To fMsg.NoOfDesignedMsgSects
            '~~ Save the label and the text udt into a Dictionary by transfering it into an array
            .MsgLabel(i) = dsply_msg.Section(i).Label
            .MsgText(i) = dsply_msg.Section(i).Text
        Next i
        
        .MsgButtons = dsply_buttons
        .DefaultBttn = dsply_button_default
        '+------------------------------------------------------------------------+
        '|| Setup prior showing the form improves the performance significantly  ||
        '|| and avoids any flickering message window with its setup.             ||
        '|| For testing purpose it may be appropriate to out-comment the Setup.  ||
        '+------------------------------------------------------------------------+
        .Setup '                                                                 ||
        If dsply_modeless Then
            DisplayDone = False
            .show vbModeless
            .top = 1
            .Left = 1
        Else
            .show vbModal
        End If
    End With
    Dsply = RepliedWith

xt: Exit Function

eh: ErrMsg ErrSrc(PROC)
#If Debugging Then
    Stop: Resume
#End If
End Function

Public Function ReplyString( _
          ByVal vReply As Variant) As String
' ------------------------------------------
' Returns the Dsply or Box return value as
' string. An invalid value is ignored.
' ------------------------------------------

    If VarType(vReply) = vbString Then
        ReplyString = vReply
    Else
        Select Case vReply
            Case vbAbort:   ReplyString = "Abort"
            Case vbCancel:  ReplyString = "Cancel"
            Case vbIgnore:  ReplyString = "Ignore"
            Case vbNo:      ReplyString = "No"
            Case vbOK:      ReplyString = "Ok"
            Case vbRetry:   ReplyString = "Retry"
            Case vbYes:     ReplyString = "Yes"
        End Select
    End If
    
End Function


Private Sub ErrMsg( _
             ByVal err_source As String, _
    Optional ByVal err_no As Long = 0, _
    Optional ByVal err_dscrptn As String = vbNullString, _
    Optional ByVal err_line As Long = 0)
' ------------------------------------------------------------------------------
' This 'Common VBA Component' uses only a kind of minimum error handling!
' ------------------------------------------------------------------------------
    Dim ErrNo   As Long
    Dim ErrDesc As String
    Dim ErrType As String
    Dim errline As Long
    Dim AtLine  As String
    
    If err_no = 0 Then err_no = Err.Number
    If err_no < 0 Then
        ErrNo = AppErr(err_no)
        ErrType = "Applicatin error "
    Else
        ErrNo = err_no
        ErrType = "Runtime error "
    End If
    If err_dscrptn = vbNullString Then ErrDesc = Err.Description Else ErrDesc = err_dscrptn
    If err_line = 0 Then errline = Erl
    If err_line <> 0 Then AtLine = " at line " & err_line
    MsgBox Title:=ErrType & ErrNo & " in " & err_source _
         , Prompt:="Error : " & ErrDesc & vbLf & _
                   "Source: " & err_source & AtLine _
         , Buttons:=vbCritical
End Sub
Private Function ErrSrc(ByVal s As String) As String
    ErrSrc = "mMsg." & s
End Function

