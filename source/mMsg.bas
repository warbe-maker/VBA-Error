Attribute VB_Name = "mMsg"
Option Explicit
' ------------------------------------------------------------------------------
' Standard Module mMsg
'               Message display services using the fMsg form.
'
' Public services:
' ------------------------------------------------------------------------------
' - Box         In analogy to the MsgBox, provides a simple message but with all
'               the fexibility for the display of up to 49 reply buttons.
' - Buttons     Supports the specification of the design buttons displayed in 7
'               rows by 7 buttons each
' - Dsply       Exposes all properties and methods for the display of any kind
'               of message
' - Monitor     Uses modeless instances of the fMsg form - any instance is
'               identified by the window title - to display the progress of a
'               process or monitor intermediate results.
'
' See details at:
' https://warbe-maker.github.io/warbe-maker.github.io/vba/common/2020/11/17/Common-VBA-Message-Services.html
' https://github.com/warbe-maker/Common-VBA-Message-Service
'
' W. Rauschenberger, Berlin Jan 2021 (last revision)
' ------------------------------------------------------------------------------
' ------------------------------------------------------------
' Means to get and calculate the display devices DPI in points
Const SM_XVIRTUALSCREEN                 As Long = &H4C&
Const SM_YVIRTUALSCREEN                 As Long = &H4D&
Const SM_CXVIRTUALSCREEN                As Long = &H4E&
Const SM_CYVIRTUALSCREEN                As Long = &H4F&
Const LOGPIXELSX                        As Long = 88
Const LOGPIXELSY                        As Long = 90
Const TWIPSPERINCH                      As Long = 1440
Private Declare PtrSafe Function GetSystemMetrics32 Lib "user32" Alias "GetSystemMetrics" (ByVal nIndex As Long) As Long
Private Declare PtrSafe Function GetDC Lib "user32" (ByVal hWnd As Long) As Long
Private Declare PtrSafe Function GetDeviceCaps Lib "gdi32" (ByVal hDC As Long, ByVal nIndex As Long) As Long
Private Declare PtrSafe Function ReleaseDC Lib "user32" (ByVal hWnd As Long, ByVal hDC As Long) As Long
' ------------------------------------------------------------
Public Const MSG_WIDTH_MIN_LIMIT_PERCENTAGE    As Long = 15
Public Const MSG_WIDTH_MAX_LIMIT_PERCENTAGE    As Long = 95
Public Const MSG_HEIGHT_MIN_LIMIT_PERCENTAGE   As Long = 20
Public Const MSG_HEIGHT_MAX_LIMIT_PERCENTAGE   As Long = 85

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

Public Property Get Modeless() As Boolean:          Modeless = bModeless:   End Property

Public Property Let Modeless(ByVal b As Boolean):   bModeless = b:          End Property

Private Property Get ScreenHeight() As Single
    ConvertPixelsToPoints y_dpi:=GetSystemMetrics32(SM_CYVIRTUALSCREEN), y_pts:=ScreenHeight
End Property

Private Property Get ScreenWidth() As Single
    ConvertPixelsToPoints x_dpi:=GetSystemMetrics32(SM_CXVIRTUALSCREEN), x_pts:=ScreenWidth
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

Public Sub AssertWidthAndHeight(ByRef width_min As Long, _
                                ByRef width_max As Long, _
                                ByRef height_min As Long, _
                                ByRef height_max As Long)
' ------------------------------------------------------------------------------
' Returns all provided arguments in pt. When any value is not asserted valid
' a corresponding return code is returned.
' When the min width is greater than the max width it is set equal with the max
' When the height min is greater than the height max it is set to the max limit.
' A min width below the min width limit is set to the min limit
' A max width 0 or above the max width limit is set to the max limit
' A min height below the min height limit is set to the height min height limit
' A max height 0 or above the max height limit is set to the max height limit.
' Note: Public for test purpose only
' ------------------------------------------------------------------------------
    Const PROC  As String = "AssertWidthAndHeight"

    '~~ Convert all limits from percentage to pt
    Dim MsgWidthMaxLimitPt  As Long:    MsgWidthMaxLimitPt = Pnts(MSG_WIDTH_MAX_LIMIT_PERCENTAGE, "w")
    Dim MsgWidthMinLimitPt  As Long:    MsgWidthMinLimitPt = Pnts(MSG_WIDTH_MIN_LIMIT_PERCENTAGE, "w")
    Dim MsgHeightMaxLimitPt As Long:    MsgHeightMaxLimitPt = Pnts(MSG_HEIGHT_MAX_LIMIT_PERCENTAGE, "h")
    Dim MsgHeightMinLimitPt As Long:    MsgHeightMinLimitPt = Pnts(MSG_HEIGHT_MIN_LIMIT_PERCENTAGE, "h")
    
    '~~ Convert all percentage arguments into pt arguments
    If width_max <= 100 Then width_max = Pnts(width_max, "w")
    If width_min <= 100 Then width_min = Pnts(width_min, "w")
    If height_max <= 100 Then height_max = Pnts(height_max, "h")
    If height_min <= 100 Then height_min = Pnts(height_min, "h")
        
    '~~ Set all invalid, improper, or useless arguments to sensible values
    If width_min > width_max Then width_min = width_max
    If height_min > height_max Then height_min = height_max
    If width_min < MsgWidthMinLimitPt Then width_min = MsgWidthMinLimitPt
    If width_max = 0 Or width_max > MsgWidthMaxLimitPt Then width_max = MsgWidthMaxLimitPt
    If height_min < MsgHeightMinLimitPt Then height_min = MsgHeightMinLimitPt
    If height_max = 0 Or height_max > MsgHeightMaxLimitPt Then height_max = MsgHeightMaxLimitPt
    
End Sub

Public Function Box(ByVal box_title As String, _
           Optional ByVal box_msg As String = vbNullString, _
           Optional ByVal box_monospaced As Boolean = False, _
           Optional ByVal box_buttons As Variant = vbOKOnly, _
           Optional ByVal box_buttons_width_min = 70, _
           Optional ByVal box_button_default = 1, _
           Optional ByVal box_returnindex As Boolean = False, _
           Optional ByVal box_width_min As Long = 300, _
           Optional ByVal box_width_max As Long = 85, _
           Optional ByVal box_height_min As Long = 20, _
           Optional ByVal box_height_max As Long = 85) As Variant
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
    Const PROC = "Box§"
    
    On Error GoTo eh
    Dim Message As TypeMsgText
    Dim MsgForm As fMsg

    Message.Text = box_msg
    Message.MonoSpaced = box_monospaced

    AssertWidthAndHeight box_width_min _
                       , box_width_max _
                       , box_height_min _
                       , box_height_max
    
    '~~ In order to avoid any interferance with modeless displayed fMsg form
    '~~ all services create and use their own instance identified by the message title.
    Set MsgForm = Form(box_title)
    With MsgForm
        .MsgTitle = box_title
        .MsgText(1) = Message
        .MsgButtons = box_buttons
        .MsgHeightMax = box_height_max    ' percentage of screen height
        .MsgHeightMin = box_height_min    ' percentage of screen height
        .MsgWidthMax = box_width_max      ' percentage of screen width
        .MsgWidthMin = box_width_min        ' defaults to 400 pt. the absolute minimum is 200 pt
        .MinButtonWidth = box_buttons_width_min
        .MsgButtonDefault = box_button_default
        '+------------------------------------------------------------------------+
        '|| Setup prior showing the form improves the performance significantly  ||
        '|| and avoids any flickering message window with its setup.             ||
        '|| For testing purpose it may be appropriate to out-comment the Setup.  ||
        .Setup '                                                                 ||
        '+------------------------------------------------------------------------+
        .show
    End With
    Box = RepliedWith

xt: Exit Function

eh: If ErrMsg(ErrSrc(PROC)) = vbYes Then: Stop: Resume
End Function

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

Private Sub ConvertPixelsToPoints(Optional ByVal x_dpi As Single, _
                                  Optional ByVal y_dpi As Single, _
                                  Optional ByRef x_pts As Single, _
                                  Optional ByRef y_pts As Single)
' ------------------------------------------------------------------------------
' Returns pixels (device dependent) to points (used by Excel).
' ------------------------------------------------------------------------------
    
    Dim hDC            As Long
    Dim RetVal         As Long
    Dim PixelsPerInchX As Long
    Dim PixelsPerInchY As Long
 
    On Error Resume Next
    hDC = GetDC(0)
    PixelsPerInchX = GetDeviceCaps(hDC, LOGPIXELSX)
    PixelsPerInchY = GetDeviceCaps(hDC, LOGPIXELSY)
    RetVal = ReleaseDC(0, hDC)
    If Not IsMissing(x_dpi) And Not IsMissing(x_pts) Then x_pts = x_dpi * TWIPSPERINCH / 20 / PixelsPerInchX
    If Not IsMissing(y_dpi) And Not IsMissing(y_pts) Then y_pts = y_dpi * TWIPSPERINCH / 20 / PixelsPerInchY

End Sub

                                     
Public Function Dsply(ByVal dsply_title As String, _
                      ByRef dsply_msg As TypeMsg, _
             Optional ByVal dsply_buttons As Variant = vbOKOnly, _
             Optional ByVal dsply_button_default = 1, _
             Optional ByVal dsply_button_width_min = 0, _
             Optional ByVal dsply_button_reply_with_index As Boolean = False, _
             Optional ByVal dsply_modeless As Boolean = False, _
             Optional ByVal dsply_width_min As Long = 15, _
             Optional ByVal dsply_width_max As Long = 85, _
             Optional ByVal dsply_height_min As Long = 25, _
             Optional ByVal dsply_height_max As Long = 85) As Variant
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
' dsply_button_reply_with_index | Defaults to False, when True the index of the
'                        | of the pressed button is returned else the caption
'                        | or the VBA.MsgBox button value respectively
' dsply_modeless         | The message is displayed modeless, defaults to False
'                        | = vbModal
' dsply_width_min        | Overwrites the default when not 0
' dsply_width_max        | Overwrites the default when not 0
' dsply_height_max       | Overwrites the default when not 0
' dsply_button_width_min | Overwrites the default when not 0
'
' See: https://github.com/warbe-maker/Common-VBA-Message-Service
'
' W. Rauschenberger, Berlin, Nov 2020
' -------------------------------------------------------------------------------------
    Const PROC = "Dsply"
    
    On Error GoTo eh
    Dim i       As Long
    Dim MsgForm As fMsg

    AssertWidthAndHeight dsply_width_min _
                       , dsply_width_max _
                       , dsply_height_min _
                       , dsply_height_max
    
    Set MsgForm = Form(dsply_title)
    
    With MsgForm
        .ReplyWithIndex = dsply_button_reply_with_index
        If dsply_height_max > 0 Then .MsgHeightMax = dsply_height_max ' percentage of screen height
        If dsply_width_max > 0 Then .MsgWidthMax = dsply_width_max    ' percentage of screen width
        If dsply_width_min > 0 Then .MsgWidthMin = dsply_width_min      ' defaults to 300 pt. the absolute minimum is 200 pt
        If dsply_button_width_min > 0 Then .MinButtonWidth = dsply_button_width_min
        .MsgTitle = dsply_title
        For i = 1 To .NoOfDesignedMsgSects
            '~~ Save the label and the text udt into a Dictionary by transfering it into an array
            .MsgLabel(i) = dsply_msg.Section(i).Label
            .MsgText(i) = dsply_msg.Section(i).Text
        Next i
        
        .MsgButtons = dsply_buttons
        .MsgButtonDefault = dsply_button_default
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

eh: If ErrMsg(ErrSrc(PROC)) = vbYes Then: Stop: Resume
End Function

Public Function ErrMsg(ByVal err_source As String, _
              Optional ByVal err_no As Long = 0, _
              Optional ByVal err_dscrptn As String = vbNullString, _
              Optional ByVal err_line As Long = 0) As Variant
' ------------------------------------------------------------------------------
' Displays a proper designe error message providing the option to resume the
' error line when the Conditional Compile Argument Debugging = 1.
' ------------------------------------------------------------------------------
    Dim ErrNo       As Long
    Dim ErrDesc     As String
    Dim ErrType     As String
    Dim ErrLine     As Long
    Dim ErrAtLine   As String
    Dim ErrBttns    As Long
    Dim ErrMsgText  As TypeMsg
    
    '~~ Obtain error information from the Err object for any argument not provided
    If err_no = 0 Then err_no = Err.Number
    If err_line = 0 Then err_line = Erl
    If err_source = vbNullString Then err_source = Err.Source
    If err_dscrptn = vbNullString Then err_dscrptn = Err.Description
    If err_dscrptn = vbNullString Then err_dscrptn = "--- No error description available ---"
    
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
    
    If err_line <> 0 Then ErrAtLine = " at line " & err_line
    
    If err_dscrptn = vbNullString Then err_dscrptn = "--- No error message available ---"
    With ErrMsgText.Section(1)
        .Label.Text = "Error:"
        .Label.FontColor = rgbBlue
        .Text.Text = err_dscrptn
    End With
    With ErrMsgText.Section(2)
        .Label.Text = "Source:"
        .Label.FontColor = rgbBlue
        .Text.Text = err_source & ErrAtLine
    End With

#If Debugging Then
    ErrBttns = vbYesNoCancel
    With ErrMsgText.Section(3)
        .Label.Text = "Debugging: (Conditional Compile Argument 'Debugging = 1')"
        .Label.FontColor = rgbBlue
        .Text.MonoSpaced = True
        .Text.Text = "Yes    = Resume error line" & vbLf & _
                     "No     = Resume Next" & vbLf & _
                     "Cancel = Terminate"
    End With
    With ErrMsgText.Section(4)
        .Label.Text = "Use the debugging options as follows:"
        .Label.FontColor = rgbBlue
        .Text.MonoSpaced = True
        .Text.Text = "    Private Sub Any()                   " & vbLf & _
                     "        Const PROC = ""Any""            " & vbLf & _
                     "        On Error Goto eh                " & vbLf & _
                     "        ' any code                      " & vbLf & _
                     "    xt: Exit Sub                        " & vbLf & vbLf & _
                     "    eh: Select Case ErrMsg(ErrSrc(PROC))" & vbLf & _
                     "            Case vbYes: Stop: Resume    " & vbLf & _
                     "            Case vbNo:  Resume Next     " & vbLf & _
                     "            Case Else:  Goto xt         " & vbLf & _
                     "         End Select                     " & vbLf & _
                     "    End Sub                             "
    End With
#Else
    ErrBttns = vbCritical
#End If
    
    ErrMsg = Dsply(dsply_title:=ErrType & ErrNo & " in " & err_source & ErrAtLine _
                 , dsply_msg:=ErrMsgText _
                 , dsply_buttons:=ErrBttns)
End Function

Private Function ErrSrc(ByVal sProc As String) As String
    ErrSrc = "mMsg." & sProc
End Function

Public Function Form(ByVal frm_caption As String, _
            Optional ByVal frm_unload As Boolean = False, _
            Optional ByVal frm_caller As String = vbNullString) As fMsg
' -------------------------------------------------------------------------
' Returns an instance of the fMsg UserForm which is uniquely identified by
' by the caption (frm_caption). When the instance is already collected in
' the MsgForms Dictionary and it effectively exists this instance is
' returned. Else the no longer existing instance is removed from the
' dictionary and a new instance is created, stored in the Dictionary and
' returned.
' -------------------------------------------------------------------------
    Const PROC = "Form"
    
    On Error GoTo eh
    Static MsgForms As Dictionary   ' Collection of active form instances with their caption as the key
    Dim MsgForm     As fMsg

    If frm_caller <> vbNullString Then frm_caller = "(" & frm_caller & ")"
    
    If MsgForms Is Nothing Then Set MsgForms = New Dictionary
    
    If frm_unload Then
        If MsgForms.Exists(frm_caption) Then
            On Error Resume Next
            Unload MsgForm(frm_caption)
            MsgForms.Remove frm_caption
        End If
        Exit Function
    End If
    
    If Not MsgForms.Exists(frm_caption) Then
        '~~ There is no evidence of an already existing fMsg instance
        Set MsgForm = New fMsg
        Debug.Print "fMsg instance titled ||" & frm_caption & "|| new created, initalized and returned to caller " & frm_caller
        MsgForms.Add frm_caption, MsgForm
        Debug.Print "fMsg instance titled ||" & frm_caption & "|| saved to Dictionary"
    Else
        '~~ An fMsg instance exists in the Dictionary, it may however no longer exist in the system
        On Error Resume Next
        Set MsgForm = MsgForms(frm_caption)
        Select Case Err.Number
            Case 0
                Debug.Print "fMsg instance titled ||" & frm_caption & "|| returned to caller " & frm_caller
            Case 13
                '~~ The fMsg instance no longer exists
                If MsgForms.Exists(frm_caption) Then
                    MsgForms.Remove frm_caption
                    Debug.Print "fMsg instance titled ||" & frm_caption & "|| removed from Dictionary"
                End If
                Set MsgForm = New fMsg
                Debug.Print "fMsg instance titled ||" & frm_caption & "|| created, initialized and returned to caller " & frm_caller
                MsgForms.Add frm_caption, MsgForm
                Debug.Print "fMsg instance titled ||" & frm_caption & "|| saved to Dictionary"
            Case Else
                '~~ Unknown error!
                Debug.Print "Unexpectd error number " & Err.Number & "!"
                Err.Raise AppErr(1), ErrSrc(PROC), "Unknown/unrecognized error!"
        End Select
        On Error GoTo -1
        
    End If

xt: Set Form = MsgForm
    Exit Function

eh: ErrMsg ErrSrc(PROC)
End Function

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

Public Function Monitor( _
                   ByVal mntr_title As String, _
                   ByRef mntr_msg As String, _
          Optional ByVal mntr_header As String = vbNullString, _
          Optional ByVal mntr_buttons As Variant = vbNullString, _
          Optional ByVal mntr_footer As String = "Process in progress! Please wait.", _
          Optional ByVal mntr_msg_append As Boolean = True, _
          Optional ByVal mntr_msg_monospaced As Boolean = False, _
          Optional ByVal mntr_width_min As Long = 40, _
          Optional ByVal mntr_width_max As Long = 85, _
          Optional ByVal mntr_height_min As Long = 20, _
          Optional ByVal mntr_height_max As Long = 85) As Variant
' -------------------------------------------------------------------------------------
' Displays an instance of the Common VBA Message Form (fMsg) modeless for a series of
' progress messages. When no instance for the provided title (mntr_title) exists one
' is created, else the existing instance is used. This allows multiple modeless message
' windows at the same time. Arguments:
' - mntr_title ........: Text displayed at the window handele bar. Identifies the
'                         progress. I.e. a different title would display a process in
'                         another instance of the fMsg form.
' - mntr_msg ..........: The process message displayed.
' - mntr_header .......: The text displayed above the mntr_msg
' - mntr_footer .......: The text displayed below the mntr_msg.
'                         Defaults to "Process in progress! Please wait."
' - mntr_msg_append ...: Defaults to True. Any text provided with mntr_msg is
'                         appended to the text already displayed
' - mntr_msg_monospaced: Displays the mntr_msg monospaced
' - mntr_width_min ....: Defaults to 400
' - mntr_width_max ....: Defaults to 80% of the screen size
' - mntr_height_max ...: Defaults to 70% of the screen size
'
' See: https://warbe-maker.github.io/vba/common/2020/11/17/Common-VBA-Message-Form.html
'
' W. Rauschenberger, Berlin, May 2021
' -------------------------------------------------------------------------------------
    Const PROC = "Monitor"
   
    On Error GoTo eh
    Dim msg     As TypeMsg
    Dim MsgForm As fMsg

    AssertWidthAndHeight mntr_width_min _
                       , mntr_width_max _
                       , mntr_height_min _
                       , mntr_height_max
    
    Set MsgForm = Form(mntr_title)
    msg.Section(1).Label.Text = mntr_header
    msg.Section(1).Label.MonoSpaced = mntr_msg_monospaced
    msg.Section(1).Label.FontBold = True
    msg.Section(1).Text.Text = mntr_msg
    msg.Section(1).Text.MonoSpaced = mntr_msg_monospaced
    
    msg.Section(2).Text.Text = mntr_footer
    msg.Section(2).Text.FontColor = rgbBlue
    msg.Section(2).Text.FontSize = 8
    msg.Section(2).Text.FontBold = True
    
    If Trim(MsgForm.MsgTitle) <> Trim(mntr_title) Then
        With MsgForm
            '~~ A new title starts a new progress message
            .MsgTitle = mntr_title
            .MsgLabel(1) = msg.Section(1).Label
            .MsgText(1) = msg.Section(1).Text
            .MsgText(2) = msg.Section(2).Text
            .MsgButtons = mntr_buttons
            .MsgWidthMin = mntr_width_min   ' pt min width
            .MsgWidthMax = mntr_width_max   ' pt max width
            .MsgHeightMin = mntr_height_min ' pt min height
            .MsgHeightMax = mntr_height_max ' pt max height
            .MonitorMode = True
            .DsplyFrmsWthCptnTestOnly = False

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
        MsgForm.Monitor mntr_text:=mntr_msg _
                      , mntr_append:=mntr_msg_append _
                      , mntr_footer:=msg.Section(2).Text.Text
    End If
      
xt: Exit Function

eh: ErrMsg ErrSrc(PROC)
End Function

Public Function Pnts(ByVal pt_value As Long, _
                     ByVal pt_dimension As String) As Single
' ------------------------------------------------------------------------------
' Returns a value as pt considering a dimensions pt. A value <= 100 is regarded
' a percentage value and transformed to pt. A value > 100 is regarded a pt value
' already.
' ------------------------------------------------------------------------------
    If pt_value <= 100 Then
        If UCase(Left(pt_dimension, 1)) = "W" _
        Then Pnts = RoundUp(ScreenWidth * (pt_value / 100)) _
        Else Pnts = RoundUp(ScreenHeight * (pt_value / 100))
    Else
        Pnts = pt_value
    End If
End Function

Public Function Prcnt(ByVal pc_value As Long, _
                     ByVal pc_dimension As String) As Single
' ------------------------------------------------------------------------------
' Returns a value as percentage considering a screen dimensions pt. A value
' <= 100 is regarded a percentage already, a value > 100 is regarded a pt value
' and transformed to a percentage.
' ------------------------------------------------------------------------------
    If pc_value > 100 Then
        If UCase(Left(pc_dimension, 1)) = "W" _
        Then Prcnt = RoundUp(pc_value / (ScreenWidth / 100)) _
        Else Prcnt = RoundUp(pc_value / (ScreenHeight / 100))
    Else
        Prcnt = pc_value
    End If
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

Public Function RoundUp(ByVal v As Variant) As Variant
' -------------------------------------------------------------------------------------
' Returns (v) rounded up to the next integer. Note: to round down omit the "+ 0.5").
' -------------------------------------------------------------------------------------
    RoundUp = Int(v) + (v - Int(v) + 0.5) \ 1
End Function

