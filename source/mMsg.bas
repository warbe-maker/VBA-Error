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
Const SM_XVIRTUALSCREEN                         As Long = &H4C&
Const SM_YVIRTUALSCREEN                         As Long = &H4D&
Const SM_CXVIRTUALSCREEN                        As Long = &H4E&
Const SM_CYVIRTUALSCREEN                        As Long = &H4F&
Const LOGPIXELSX                                As Long = 88
Const LOGPIXELSY                                As Long = 90
Const TWIPSPERINCH                              As Long = 1440
Private Declare PtrSafe Function GetSystemMetrics32 Lib "user32" Alias "GetSystemMetrics" (ByVal nIndex As Long) As Long
Private Declare PtrSafe Function GetDC Lib "user32" (ByVal hWnd As Long) As Long
Private Declare PtrSafe Function GetDeviceCaps Lib "gdi32" (ByVal hDC As Long, ByVal nIndex As Long) As Long
Private Declare PtrSafe Function ReleaseDC Lib "user32" (ByVal hWnd As Long, ByVal hDC As Long) As Long
' ------------------------------------------------------------
Public Const MSG_WIDTH_MIN_LIMIT_PERCENTAGE     As Long = 25
Public Const MSG_WIDTH_MAX_LIMIT_PERCENTAGE     As Long = 98
Public Const MSG_HEIGHT_MIN_LIMIT_PERCENTAGE    As Long = 20
Public Const MSG_HEIGHT_MAX_LIMIT_PERCENTAGE    As Long = 95

Public Const END_OF_PROGRESS                    As String = "EndOfProgress"

' Extension of the VBA.MsgBox constants for the Debugging option of the ErrMsg service
' to display additional debugging buttons
Public Const vbResumeOk                         As Long = 7 ' Buttons value in mMsg.ErrMsg (pass on not supported)
Public Const vbResume                           As Long = 6 ' return value (equates to vbYes)

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

Public Property Get ScreenHeight() As Single
'    Debug.Print "Screen-Height: " & GetSystemMetrics32(SM_CYVIRTUALSCREEN) & " dpi"
    ConvertPixelsToPoints y_dpi:=GetSystemMetrics32(SM_CYVIRTUALSCREEN), y_pts:=ScreenHeight
End Property

Public Property Get ScreenWidth() As Single
'    Debug.Print "Screen-Width: " & GetSystemMetrics32(SM_CXVIRTUALSCREEN) & " dpi"
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
    If width_max <> 0 And width_max <= 100 Then width_max = Pnts(width_max, "w")
    If width_min <> 0 And width_min <= 100 Then width_min = Pnts(width_min, "w")
    If height_max <> 0 And height_max <= 100 Then height_max = Pnts(height_max, "h")
    If height_min <> 0 And height_min <= 100 Then height_min = Pnts(height_min, "h")
        
    '~~ Provide sensible values for all invalid, improper, or useless
    If width_min > width_max Then width_min = width_max
    If height_min > height_max Then height_min = height_max
    If width_min < MsgWidthMinLimitPt Then width_min = MsgWidthMinLimitPt
    If width_max <= width_min Then width_max = width_min
    If width_max > MsgWidthMaxLimitPt Then width_max = MsgWidthMaxLimitPt
    If height_min < MsgHeightMinLimitPt Then height_min = MsgHeightMinLimitPt
    If height_max = 0 Or height_max < height_min Then height_max = height_min
    If height_max > MsgHeightMaxLimitPt Then height_max = MsgHeightMaxLimitPt
    
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
    Set MsgForm = MsgInstance(box_title)
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
        .Show
    End With
    Box = RepliedWith

xt: Exit Function

eh: If ErrMsg(ErrSrc(PROC)) = vbYes Then: Stop: Resume
End Function

Public Function Buttons(ByRef bttns_collection As Collection, _
                        ParamArray bttns() As Variant) As Collection
' --------------------------------------------------------------------------
' Returns a collection comprising of the provided collection
' (bttns_collection) with the provided items (bttns) added and only the
' provided items (bttns) as collection (bttns_coollection. The function may
' thus be used to have a collection with the provided items and one with the
' provided items added to the provided collection.
' Example: Set cll = Buttons(cll, "A", "B") ' mode add
' Example: Buttons cll, "A", "B"
' --------------------------------------------------------------------------
    
    Dim i               As Long
    Dim s               As String
    Dim cllOnly         As New Collection
    Dim lOnlyBttnsInRow As Long                 ' buttons in a row counter
    Dim lOnlyBttns      As Long                 ' total buttons in cllOnly
    Dim lOnlyRows       As Long: lOnlyRows = 1  ' button rows counter
    Dim cllAdd          As Collection
    Dim lAddBttnsInRow  As Long                 ' buttons in a row counter (excludes break items)
    Dim lAddBttns       As Long                 ' total buttons in cllAdd
    Dim lAddRows        As Long: lAddRows = 1   ' button rows counter
    Dim v1              As Variant
    Dim v2              As Variant
    Dim cllBttns        As New Collection
    
    If bttns_collection Is Nothing Then
        Set bttns_collection = New Collection
        Set cllAdd = New Collection
    Else
        Set cllAdd = bttns_collection
        '~~ Count the buttons already specified in cllAdd
        For Each v1 In cllAdd
            If v1 = vbLf Or v1 = vbCrLf Or v1 = vbCr Then
                lAddBttnsInRow = 0
            Else
                lAddBttnsInRow = lAddBttnsInRow + 1
                lAddRows = lAddRows + 1
                lAddBttns = lAddBttns + 1
            End If
        Next v1
    End If
    
    On Error Resume Next
    i = LBound(bttns)
    If Err.Number <> 0 Then GoTo xt
    
    '~~ Transpose the the buttons argument (bttns) into a collection considering
    '~~ that an item may be one with sub-items in a comma delimited string.
    For Each v1 In bttns
        If InStr(v1, ",") <> 0 Then
            '~~ Comma deliomited string
            For Each v2 In Split(v1, ",")
                cllBttns.Add v2
            Next v2
        Else
            cllBttns.Add v1
        End If
    Next v1
    
    '~~ Prepare the cllAdd and the cllOnly Collection
    For Each v1 In cllBttns
        If VarType(v1) = vbEmpty Then GoTo nx  ' skip empty items
        If (lOnlyRows = 7 And lOnlyBttnsInRow = 7) Or lAddBttns = 49 Then GoTo xt ' max possible buttons reached
        Select Case v1
            
            Case vbLf, vbCrLf, vbCr
                cllOnly.Add v1: lOnlyBttnsInRow = 0
                cllAdd.Add v1:  lAddBttnsInRow = 0
            
            Case vbOKOnly, vbOKCancel, vbYesNo, vbRetryCancel, vbResumeOk
                '~~ Two more buttons
                If lOnlyBttnsInRow = 6 Then
                    cllOnly.Add vbLf
                    lOnlyBttnsInRow = 0
                End If
                If lAddBttnsInRow = 6 Then
                    cllAdd.Add vbLf
                    lAddBttnsInRow = 0
                End If
                cllOnly.Add v1: lOnlyBttnsInRow = lOnlyBttnsInRow + 2
                cllAdd.Add v1:  lAddBttnsInRow = lAddBttnsInRow + 2
                lAddBttns = lAddBttns + 2
            
            Case vbAbortRetryIgnore, vbYesNoCancel
                '~~ Three more buttons
                If lOnlyBttnsInRow = 5 Then
                    cllOnly.Add vbLf
                    lOnlyBttnsInRow = 0
                End If
                If lAddBttnsInRow = 5 Then
                    cllAdd.Add vbLf
                    lAddBttnsInRow = 0
                End If
                cllOnly.Add v1: lOnlyBttnsInRow = lOnlyBttnsInRow + 2
                cllAdd.Add v1:  lAddBttnsInRow = lAddBttnsInRow + 3
                lAddBttns = lAddBttns + 3
            
            Case Else
                If TypeName(v1) = "String" Then
                    ' Any invalid buttons value will be ignored without notice
                    If lOnlyBttnsInRow = 7 Then
                        cllOnly.Add vbLf
                        lOnlyBttnsInRow = 0
                    End If
                    If lAddBttnsInRow = 7 Then
                        cllAdd.Add vbLf
                        lAddBttnsInRow = 0
                    End If
                    cllOnly.Add v1: lOnlyBttnsInRow = lOnlyBttnsInRow + 1:  lOnlyBttns = lOnlyBttns + 1
                    cllAdd.Add v1:  lAddBttnsInRow = lAddBttnsInRow + 1:    lAddBttns = lAddBttns + 1
                End If
        End Select
nx: Next v1
    
xt: Set Buttons = cllAdd
    Set bttns_collection = cllOnly
    Exit Function

End Function

'Public Function ButtonsArray(ByVal msg_buttons As Variant) As Variant
'' ------------------------------------------------------------------------------
'' Returns the button captions (msg_buttons) which may be provided as komma
'' delimited string, array, collection, or Dictionary, as komma delimited string.
'' ------------------------------------------------------------------------------
'
'    Dim va()    As Variant
'    Dim i       As Long
'    Dim dct     As Dictionary
'    Dim cll     As Collection
'
'    Debug.Print TypeName(msg_buttons)
'    Select Case VarType(msg_buttons)
'        Case vbArray:   ButtonsArray = msg_buttons
'        Case vbString: ButtonsArray = Split(msg_buttons, ",")
'        Case Else
'            Select Case TypeName(msg_buttons)
'                Case "Dictionary"
'                    Set dct = msg_buttons
'                    ReDim va(dct.Count - 1)
'                    For i = 0 To dct.Count - 1
'                        va(i) = dct.Items()(i)
'                    Next i
'                    ButtonsArray = va
'                Case "Collection"
'                    Set cll = msg_buttons
'                    ReDim va(cll.Count - 1)
'                    For i = 0 To cll.Count - 1
'                        va(i) = cll.Item(i + 1)
'                    Next i
'                    ButtonsArray = va
'            End Select
'    End Select
'
'End Function

Public Function ButtonsNumeric(ByVal bn_num_buttons As Long) As Long
' -------------------------------------------------------------------------------------
' Returns the Buttons argument (bn_num_buttons) with additional options removed.
' In order to mimic the Buttons argument of the VBA.MsgBox any values added for other
' options but the display of the buttons are unstripped (i.e. the values are deducted).
' -------------------------------------------------------------------------------------
    Const PROC = "ButtonsNumeric"
    
    On Error GoTo eh
        
    While bn_num_buttons >= vbCritical                  ' 16
        Select Case bn_num_buttons
            '~~ VBA.MsgBox Display options
            Case Is >= vbMsgBoxRtlReading                ' 1048576  not implemented
                bn_num_buttons = bn_num_buttons - vbMsgBoxRtlReading
    
            Case Is >= vbMsgBoxRight                     ' 524288   not implemented
                bn_num_buttons = bn_num_buttons - vbMsgBoxRight
    
            Case Is >= vbMsgBoxSetForeground             ' 65536    not implemented
                bn_num_buttons = bn_num_buttons - vbMsgBoxSetForeground
            
            '~~ Display of a Help button
            Case Is >= vbMsgBoxHelpButton                ' 16384    not implemented
                bn_num_buttons = bn_num_buttons - vbMsgBoxHelpButton
    
            Case Is >= vbSystemModal                     ' 4096     not implemented
                bn_num_buttons = bn_num_buttons - vbSystemModal
    
            Case Is >= vbDefaultButton4                  ' 768
                bn_num_buttons = bn_num_buttons - vbDefaultButton4
            
            Case Is >= vbDefaultButton3                  ' 512
                bn_num_buttons = bn_num_buttons - vbDefaultButton3
            
            Case Is >= vbDefaultButton2                  ' 256
                bn_num_buttons = bn_num_buttons - vbDefaultButton2
                
            Case Is >= vbInformation                      ' 64
                bn_num_buttons = bn_num_buttons - vbInformation
            
            Case Is >= vbExclamation                    ' 48
                bn_num_buttons = bn_num_buttons - vbExclamation
            
            Case Is >= vbQuestion                       ' 32
                bn_num_buttons = bn_num_buttons - vbQuestion
    
            Case Is >= vbCritical                       ' 16
                bn_num_buttons = bn_num_buttons - vbCritical
    
        End Select
    Wend
    ButtonsNumeric = bn_num_buttons

xt: Exit Function

eh:
End Function

'Public Function ButtonsString(ByVal msg_buttons As Variant) As String
'' ------------------------------------------------------------------------------
'' Returns the button captions (msg_buttons) which may be provided as komma
'' delimited string, array, collection, or Dictionary, as komma delimited string.
'' ------------------------------------------------------------------------------
'    Const PROC = "ButtonsString"
'
'    On Error GoTo eh
'    Dim v As Variant
'
'    Debug.Print TypeName(msg_buttons)
'
'    If IsArray(msg_buttons) Then
'        ButtonsString = Join(msg_buttons, ",")
'    Else
'        Select Case VarType(msg_buttons)
'            Case vbArray
'                ButtonsString = Split(msg_buttons, ",")
'            Case vbString
'            Case Else
'                Select Case TypeName(msg_buttons)
'                    Case "Dictionary"
'                    Case "Collection"
'                        For Each v In msg_buttons
'                            ButtonsString = v & ","
'                        Next v
'                        ButtonsString = Left(ButtonsString, Len(ButtonsString) - 1)
'                End Select
'        End Select
'    End If
'
'xt: Exit Function
'
'eh: If ErrMsg(ErrSrc(PROC)) = vbYes Then: Stop: Resume
'End Function

Private Sub ConvertPixelsToPoints(Optional ByVal x_dpi As Single, _
                                  Optional ByVal y_dpi As Single, _
                                  Optional ByRef x_pts As Single, _
                                  Optional ByRef y_pts As Single)
' ------------------------------------------------------------------------------
' Returns pixels (device dependent) to points.
' Results verified by: https://pixelsconverter.com/px-to-pt.
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
    If Not IsMissing(x_dpi) And Not IsMissing(x_pts) Then
        x_pts = x_dpi * TWIPSPERINCH / 20 / PixelsPerInchX
'        If Not x_pts = 0 Then Debug.Print x_dpi & " dpi = " & x_pts & " pt"
    End If
    If Not IsMissing(y_dpi) And Not IsMissing(y_pts) Then
        y_pts = y_dpi * TWIPSPERINCH / 20 / PixelsPerInchY
'        If Not y_pts = 0 Then Debug.Print y_dpi & " dpi = " & y_pts & " pt"
    End If
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
    
    Set MsgForm = MsgInstance(dsply_title)
    
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
            .Show vbModeless
            .top = 1
            .Left = 1
        Else
            .Show vbModal
        End If
    End With
    Dsply = RepliedWith

xt: Exit Function

eh: If ErrMsg(ErrSrc(PROC)) = vbYes Then: Stop: Resume
End Function

Public Function ErrMsg(ByVal err_source As String, _
              Optional ByVal err_number As Long = 0, _
              Optional ByVal err_dscrptn As String = vbNullString, _
              Optional ByVal err_line As Long = 0, _
              Optional ByVal err_buttons As Variant = vbOKOnly) As Variant
' ------------------------------------------------------------------------------
' Displays an error message.
'
' W. Rauschenberger, Berlin, Nov 2020
' ------------------------------------------------------------------------------
    
    Dim ErrNo       As Long
    Dim ErrDesc     As String
    Dim ErrType     As String
    Dim ErrLine     As Long
    Dim ErrAtLine   As String
    Dim ErrBttns    As Long
    Dim ErrMsgText  As TypeMsg
    Dim ErrAbout    As String
    Dim ErrTitle    As String
    Dim ErrButtons  As Collection
    
    '~~ Obtain error information from the Err object for any argument not provided
    If err_number = 0 Then err_number = Err.Number
    If err_line = 0 Then err_line = Erl
    If err_source = vbNullString Then err_source = Err.Source
    If err_dscrptn = vbNullString Then err_dscrptn = Err.Description
    If err_dscrptn = vbNullString Then err_dscrptn = "--- No error description available ---"
        
    '~~ Determine type of error
    Select Case err_number
        Case Is < 0
            ErrNo = AppErr(err_number)
            ErrType = "Application Error "
        Case Else
            ErrNo = err_number
            If (InStr(1, err_dscrptn, "DAO") <> 0 _
            Or InStr(1, err_dscrptn, "ODBC Teradata Driver") <> 0 _
            Or InStr(1, err_dscrptn, "ODBC") <> 0 _
            Or InStr(1, err_dscrptn, "Oracle") <> 0) _
            Then ErrType = "Database Error " _
            Else ErrType = "VB Runtime Error "
    End Select
    
    '~~ Prepare error line info when an error line is available
    If err_line <> 0 Then ErrAtLine = " at line " & err_line
    
    '~~ Prepare Error Description which might have additional information connected
    If InStr(err_dscrptn, "||") = 0 Then
        ErrDesc = err_dscrptn
    Else
        ErrDesc = Split(err_dscrptn, "||")(0)
        ErrAbout = Split(err_dscrptn, "||")(1)
    End If
    
    '~~ Prepare Error Title
    ErrTitle = ErrType & " " & ErrNo & " in: '" & err_source & "'" & ErrAtLine
    
    '~~ Prepare the Error Reply Buttons
#If Debugging = 1 Then
    mMsg.Buttons ErrButtons, vbResumeOk
#Else
    mMsg.Buttons ErrButtons, err_buttons
#End If
        
    '~~ Display the error message by means of the mMsg's Dsply function
    With ErrMsgText.Section(1)
        With .Label
            .Text = "Error description:"
            .FontColor = rgbBlue
        End With
        .Text.Text = ErrDesc
    End With
    With ErrMsgText.Section(2)
        With .Label
            .Text = "Error source:"
            .FontColor = rgbBlue
        End With
        .Text.Text = err_source
    End With
    With ErrMsgText.Section(3)
        If ErrAbout = vbNullString Then
            .Label.Text = vbNullString
            .Text.Text = vbNullString
        Else
            .Label.Text = "About this error:"
            .Text.Text = ErrAbout
        End If
        .Label.FontColor = rgbBlue
    End With
#If Debugging = 1 Then
    With ErrMsgText.Section(4)
        With .Label
            .Text = "About Debugging:"
            .FontColor = rgbBlue
        End With
        .Text.Text = "The additional debugging option button is displayed because the " & _
                     "Conditional Compile Argument 'Debugging = 1'."
        .Text.FontSize = 8
    End With
#End If
    mMsg.Dsply dsply_title:=ErrTitle _
             , dsply_msg:=ErrMsgText _
             , dsply_buttons:=ErrButtons
    ErrMsg = mMsg.RepliedWith
    
End Function

Private Function ErrSrc(ByVal sProc As String) As String
    ErrSrc = "mMsg." & sProc
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
    
    Set MsgForm = MsgInstance(mntr_title)
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
            .Show vbModeless
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

eh: If ErrMsg(ErrSrc(PROC)) = vbYes Then: Stop: Resume
End Function

Public Function MsgInstance(ByVal fi_key As String, _
                   Optional ByVal fi_unload As Boolean = False) As fMsg
' -------------------------------------------------------------------------
' Returns an instance of the UserForm fMsg which is uniquely identified by
' a uniqe string (fi_key) which may be the title of the message or anything
' else including an object . An already existing or new created instance is
' maintained in a static Dictionary with (fi_key) as the key and returned
' to the caller. When (fi_unload) is TRUE only a possibly already existing
' Userform identified by (fi_key) is unloaded.
'
' Requires: Reference to the "Microsoft Scripting Runtime".
' Usage   : The fMsg has to be replaced by the name of the desired UserForm
' -------------------------------------------------------------------------
    Const PROC = "MsgInstance"
    
    On Error GoTo eh
    Static Instances As Dictionary    ' Collection of (possibly still)  active form instances
    
    If Instances Is Nothing Then Set Instances = New Dictionary
    
    If fi_unload Then
        If Instances.Exists(fi_key) Then
            On Error Resume Next
            Unload Instances(fi_key) ' The instance may be already unloaded
            Instances.Remove fi_key
        End If
        Exit Function
    End If
    
    If Not Instances.Exists(fi_key) Then
        '~~ There is no evidence of an already existing instance
        Set MsgInstance = New fMsg
        Instances.Add fi_key, MsgInstance
    Else
        '~~ An instance identified by fi_key exists in the Dictionary.
        '~~ It may however have already been unloaded.
        On Error Resume Next
        Set MsgInstance = Instances(fi_key)
        Select Case Err.Number
            Case 0
            Case 13
                If Instances.Exists(fi_key) Then
                    '~~ The apparently no longer existing instance is removed from the Dictionarys
                    Instances.Remove fi_key
                End If
                Set MsgInstance = New fMsg
                Instances.Add fi_key, MsgInstance
            Case Else
                '~~ Unknown error!
                Err.Raise 1 + vbObjectError, ErrSrc(PROC), "Unknown/unrecognized error!"
        End Select
        On Error GoTo -1
    End If

xt: Exit Function

eh: If ErrMsg(ErrSrc(PROC)) = vbYes Then: Stop: Resume
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
        Then Prcnt = Int(pc_value / (ScreenWidth / 100)) _
        Else Prcnt = Int(pc_value / (ScreenHeight / 100))
    Else
        Prcnt = pc_value
    End If
End Function

Public Function RoundUp(ByVal v As Variant) As Variant
' -------------------------------------------------------------------------------------
' Returns (v) rounded up to the next integer. Note: to round down omit the "+ 0.5").
' -------------------------------------------------------------------------------------
    RoundUp = Int(v) + (v - Int(v) + 0.5) \ 1
End Function

Public Function StackIsEmpty(ByVal stck As Collection) As Boolean
' ----------------------------------------------------------------------------
' Returns TRUE when the stack (stck) is empty.
' ----------------------------------------------------------------------------
    If stck Is Nothing _
    Then StackIsEmpty = True _
    Else StackIsEmpty = stck.Count = 0
End Function

Public Function StackPop(ByVal stck As Collection) As Variant
' ----------------------------------------------------------------------------
' Common Stack Pop service. Returns the last item pushed on the stack (stck)
' and removes the item from the stack. When the stack (stck) is empty a
' vbNullString is returned.
' ----------------------------------------------------------------------------
    Const PROC = "StckPop"
    
    On Error GoTo eh
    If StackIsEmpty(stck) Then GoTo xt
    
    On Error Resume Next
    Set StackPop = stck(stck.Count)
    If Err.Number <> 0 _
    Then StackPop = stck(stck.Count)
    stck.Remove stck.Count

xt: Exit Function

eh: If ErrMsg(ErrSrc(PROC)) = vbYes Then: Stop: Resume
End Function

Public Sub StackPush(ByRef stck As Collection, _
                     ByVal stck_item As Variant)
' ----------------------------------------------------------------------------
' Common Stack Push service. Pushes (adds) an item (stck_item) to the stack
' (stck). When the provided stack (stck) is Nothing the stack is created.
' ----------------------------------------------------------------------------
    Const PROC = "StckPush"
    
    On Error GoTo eh
    If stck Is Nothing Then Set stck = New Collection
    stck.Add stck_item

xt: Exit Sub

eh: If ErrMsg(ErrSrc(PROC)) = vbYes Then: Stop: Resume
End Sub

