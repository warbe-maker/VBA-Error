Attribute VB_Name = "mMsg"
Option Explicit
' ------------------------------------------------------------------------------
' Standard Module mMsg
'               Message display services using the fMsg form.
'
' Public services:
' - Box         In analogy to the MsgBox, provides a simple message but with all
'               the fexibility for the display of up to 49 reply buttons.
' - Buttons     Supports the specification of the buttons displayed in a matrix
'               of 7 x 7 buttons (max 7 buttons in max 7 rows)
' - Dsply       Exposes all properties and methods for the display of any kind
'               of message
' - Monitor     Uses modeless instances of the fMsg form - any instance is
'               identified by the window title - to display the progress of a
'               process or monitor intermediate results.
' - MsgInstance Creates (when not existing) and returns an fMsg object
'               identified by the Title
'
' Uses:         fMsg
'
' Requires:     Reference to "Microsoft Scripting Runtime"
'
' See: https://github.com/warbe-maker/Common-VBA-Message-Service
'
' W. Rauschenberger, Berlin Mar 2022 (last revision)
' ------------------------------------------------------------------------------
Const LOGPIXELSX                                As Long = 88        ' -------------
Const LOGPIXELSY                                As Long = 90        ' Constants for
Const SM_CXVIRTUALSCREEN                        As Long = &H4E&     ' calculating
Const SM_CYVIRTUALSCREEN                        As Long = &H4F&     ' the
Const SM_XVIRTUALSCREEN                         As Long = &H4C&     ' display's
Const SM_YVIRTUALSCREEN                         As Long = &H4D&     ' DPI in points
Const TWIPSPERINCH                              As Long = 1440      ' -------------

' Timer means
Private Declare PtrSafe Function getFrequency Lib "kernel32" _
Alias "QueryPerformanceFrequency" (TimerSystemFrequency As Currency) As Long
Private Declare PtrSafe Function getTickCount Lib "kernel32" _
Alias "QueryPerformanceCounter" (cyTickCount As Currency) As Long

#If VBA7 Then
    Public Declare PtrSafe Sub Sleep Lib "kernel32" (ByVal ms As LongPtr)
#Else
    Public Declare PtrSafe Sub Sleep Lib "kernel32" (ByVal ms As Long)
#End If
Private Declare PtrSafe Function GetSystemMetrics32 Lib "user32" Alias "GetSystemMetrics" (ByVal nIndex As Long) As Long
Private Declare PtrSafe Function GetDC Lib "user32" (ByVal hWnd As Long) As Long
Private Declare PtrSafe Function GetDeviceCaps Lib "gdi32" (ByVal hDC As Long, ByVal nIndex As Long) As Long
Private Declare PtrSafe Function ReleaseDC Lib "user32" (ByVal hWnd As Long, ByVal hDC As Long) As Long
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
Public Const WIN_NORMAL = 1         'Open Normal
Public Const WIN_MAX = 3            'Open Maximized
Public Const WIN_MIN = 2            'Open Minimized

'***Error Codes***
Private Const ERROR_SUCCESS = 32&
Private Const ERROR_NO_ASSOC = 31&
Private Const ERROR_OUT_OF_MEM = 0&
Private Const ERROR_FILE_NOT_FOUND = 2&
Private Const ERROR_PATH_NOT_FOUND = 3&
Private Const ERROR_BAD_FORMAT = 11&

' Declarations for making a UserForm resizable
Private Declare PtrSafe Function GetForegroundWindow Lib "User32.dll" () As Long
Private Declare PtrSafe Function GetWindowLong Lib "User32.dll" Alias "GetWindowLongA" (ByVal hWnd As Long, ByVal nIndex As Long) As Long
Private Declare PtrSafe Function SetWindowLong Lib "User32.dll" Alias "SetWindowLongA" (ByVal hWnd As Long, ByVal nIndex As Long, ByVal dwNewLong As Long) As Long
Private Const WS_THICKFRAME As Long = &H40000
Private Const GWL_STYLE As Long = -16

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
        FontBold        As Boolean
        FontColor       As XlRgbColor
        FontItalic      As Boolean
        FontName        As String
        FontSize        As Long
        FontUnderline   As Boolean
        MonoSpaced      As Boolean  ' FontName defaults to "Courier New"
        Text            As String
        OpenWhenClicked As String   ' this extra option is the purpose of this sepcific Type
End Type

Public Type TypeMsgText
        FontBold        As Boolean
        FontColor       As XlRgbColor
        FontItalic      As Boolean
        FontName        As String
        FontSize        As Long
        FontUnderline   As Boolean
        MonoSpaced      As Boolean  ' FontName defaults to "Courier New"
        Text            As String
End Type

Public Type TypeMsgSect:    Label As TypeMsgLabel:  Text As TypeMsgText:    End Type
Public Type TypeMsg:        Section(1 To 4) As TypeMsgSect:                 End Type

Public Enum enStartupPosition     ' ---------------------------
    enManual = 0                  ' Used to position the
    enCenterOwner = 1             ' final setup message form
    enCenterScreen = 2            ' horizontally and vertically
    enWindowsDefault = 3          ' centered on the screen
End Enum                          ' ---------------------------

Public Enum KindOfText  ' Used with the Get/Let Text Property
    enMonHeader
    enMonFooter
    enMonStep
    enSectText
End Enum

Private bModeless       As Boolean
Public DisplayDone      As Boolean
Public RepliedWith      As Variant

Private fMonitor                As fMsg
Private MsgText1                As TypeMsgText  ' common text element
Private TextMonitorHeader       As TypeMsgText
Private TextMonitorFooter       As TypeMsgText
Private TextMonitorStep         As TypeMsgText
Private TextMsg                 As TypeMsgText
Private TextLabel               As TypeMsgText
Private TextSection             As TypeMsg

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

Public Sub AssertWidthAndHeight(Optional ByRef width_min As Long = 0, _
                                Optional ByRef WIDTH_MAX As Long = 0, _
                                Optional ByRef height_min As Long = 0, _
                                Optional ByRef height_max As Long = 0)
' ------------------------------------------------------------------------------
' Returns all provided arguments in pt.
' When the min width is greater than the max width it is set equal with the max
' When the height min is greater than the height max it is set to the max limit.
' A min width below the min width limit is set to the min limit
' A max width 0 or above the max width limit is set to the max limit
' A min height below the min height limit is set to the height min height limit
' A max height 0 or above the max height limit is set to the max height limit.
' Note: Public for test purpose only
' ------------------------------------------------------------------------------

    '~~ Convert all limits from percentage to pt
    Dim MsgWidthMaxLimitPt  As Long:    MsgWidthMaxLimitPt = Pnts(MSG_WIDTH_MAX_LIMIT_PERCENTAGE, "w")
    Dim MsgWidthMinLimitPt  As Long:    MsgWidthMinLimitPt = Pnts(MSG_WIDTH_MIN_LIMIT_PERCENTAGE, "w")
    Dim MsgHeightMaxLimitPt As Long:    MsgHeightMaxLimitPt = Pnts(MSG_HEIGHT_MAX_LIMIT_PERCENTAGE, "h")
    Dim MsgHeightMinLimitPt As Long:    MsgHeightMinLimitPt = Pnts(MSG_HEIGHT_MIN_LIMIT_PERCENTAGE, "h")
    
    '~~ Convert all percentage arguments into pt arguments
    If WIDTH_MAX <> 0 And WIDTH_MAX <= 100 Then WIDTH_MAX = Pnts(WIDTH_MAX, "w")
    If width_min <> 0 And width_min <= 100 Then width_min = Pnts(width_min, "w")
    If height_max <> 0 And height_max <= 100 Then height_max = Pnts(height_max, "h")
    If height_min <> 0 And height_min <= 100 Then height_min = Pnts(height_min, "h")
        
    '~~ Provide sensible values for all invalid, improper, or useless
    If width_min > WIDTH_MAX Then width_min = WIDTH_MAX
    If height_min > height_max Then height_min = height_max
    If width_min < MsgWidthMinLimitPt Then width_min = MsgWidthMinLimitPt
    If WIDTH_MAX <= width_min Then WIDTH_MAX = width_min
    If WIDTH_MAX > MsgWidthMaxLimitPt Then WIDTH_MAX = MsgWidthMaxLimitPt
    If height_min < MsgHeightMinLimitPt Then height_min = MsgHeightMinLimitPt
    If height_max = 0 Or height_max < height_min Then height_max = height_min
    If height_max > MsgHeightMaxLimitPt Then height_max = MsgHeightMaxLimitPt
    
End Sub

Public Function Box(ByVal Prompt As String, _
           Optional ByVal Buttons As Variant = vbOKOnly, _
           Optional ByVal Title As String = vbNullString, _
           Optional ByVal box_monospaced As Boolean = False, _
           Optional ByVal box_button_default = 1, _
           Optional ByVal box_return_index As Boolean = False, _
           Optional ByVal box_modeless As Boolean = False, _
           Optional ByVal box_width_min As Long = 300, _
           Optional ByVal box_width_max As Long = 85, _
           Optional ByVal box_height_min As Long = 20, _
           Optional ByVal box_height_max As Long = 85, _
           Optional ByVal box_pos As Variant = 3) As Variant
' -------------------------------------------------------------------------------------
' Display of a message string analogous to the VBA.Msgbox (why the first three
' arguments are identical.
' box_button_default
'
' See: https://github.com/warbe-maker/Common-VBA-Message-Service
'
' W. Rauschenberger, Berlin, Feb 2022
' -------------------------------------------------------------------------------------
    Const PROC = "Box"
    
    On Error GoTo eh
    Dim Message As TypeMsgText
    Dim MsgForm As fMsg

    If Not IsValidMsgButtonsArg(Buttons) _
    Then Err.Raise AppErr(1), ErrSrc(PROC), _
                   "The provided buttons argument is neither empty (defaults to vbOkOnly), a string " & _
                   "(optionally comma separated), a valid VBA.MsgBox value (vbYesNo, vbRetryCancel, " & _
                   "etc. plus any extra options - which may or may not be implemented), an Array, a " & _
                   "Collection, or a Dictionary! When an Array, Collection, or Dictionary at least " & _
                   "one of its items in incorrect!"

    '~~ Defaults
    If Title = vbNullString Then Title = Application.Name
    
    Message.Text = Prompt
    Message.MonoSpaced = box_monospaced

    AssertWidthAndHeight box_width_min _
                       , box_width_max _
                       , box_height_min _
                       , box_height_max
    
    
    '~~ In order to avoid any interferance with modeless displayed fMsg form
    '~~ all services create and use their own instance identified by the message title.
    Set MsgForm = MsgInstance(Title)
    With MsgForm
'        .VisualizeForTest = True
        .MsgTitle = Title
        .Text(enSectText, 1) = Message
        .MsgBttns = mMsg.Buttons(Buttons)   ' Provide the buttons as Collection
        .MsgHeightMax = box_height_max      ' percentage of screen height
        .MsgHeightMin = box_height_min      ' percentage of screen height
        .MsgWidthMax = box_width_max        ' percentage of screen width
        .MsgWidthMin = box_width_min        ' defaults to 400 pt. the absolute minimum is 200 pt
        .MsgButtonDefault = box_button_default
        '+------------------------------------------------------------------------+
        '|| Setup prior showing the form is much faster and avoids flickering.   ||
        '|| For testing purpose it may be appropriate to out-comment the Setup.  ||
        '+------------------------------------------------------------------------+
        .Setup
        If box_modeless Then
            DisplayDone = False
            .Show vbModeless
            .PositionOnScreen box_pos
        Else
            .PositionOnScreen box_pos
            .Show vbModal
        End If
    End With
    Box = RepliedWith

xt: Exit Function

eh: If ErrMsg(ErrSrc(PROC)) = vbYes Then: Stop: Resume
End Function

Public Function BttnsArgs(ByVal ba_arg As Long, _
                 Optional ByRef ba_rtl_reading As Boolean, _
                 Optional ByRef ba_box_right As Boolean, _
                 Optional ByRef ba_set_foreground As Boolean, _
                 Optional ByRef ba_help_button As Boolean, _
                 Optional ByRef ba_system_modal As Boolean, _
                 Optional ByRef ba_default_button As Long, _
                 Optional ByRef ba_information As Boolean, _
                 Optional ByRef ba_exclamation As Boolean, _
                 Optional ByRef ba_question As Boolean, _
                 Optional ByRef ba_critical As Boolean) As Long
' -------------------------------------------------------------------------------------
' Returns the Buttons argument (ba_arg) with all the options removed by returning them
' as optional arguments. In order to mimic the Buttons argument of the VBA.MsgBox any
' values added for other options but the display of the buttons are unstripped/deducted.
' I.e. the values are deducted and the corresponding argument is returtned instead).
' -------------------------------------------------------------------------------------
    Dim l As Long
    
    l = ba_arg - (Abs(Int(ba_arg / 16) * 16))
    Select Case l
        Case vbOKOnly, vbOKCancel, vbAbortRetryIgnore, vbYesNoCancel, vbYesNo, vbRetryCancel
        Case Else
            BttnsArgs = l ' may be a wromg value and thus need to be validated further
            Exit Function
    End Select

    While ba_arg >= vbCritical                          ' 16
        Select Case ba_arg
            '~~ VBA.MsgBox Display options
            Case Is >= vbMsgBoxRtlReading               ' 1048576  not implemented
                ba_arg = ba_arg - vbMsgBoxRtlReading
                ba_rtl_reading = True
            
            Case Is >= vbMsgBoxRight                    ' 524288   not implemented
                ba_arg = ba_arg - vbMsgBoxRight
                ba_box_right = True
            
            Case Is >= vbMsgBoxSetForeground            ' 65536    not implemented
                ba_arg = ba_arg - vbMsgBoxSetForeground
                ba_set_foreground = True
            
            Case Is >= vbMsgBoxHelpButton               ' 16384    not implemented: Display of a Help button
                ba_arg = ba_arg - vbMsgBoxHelpButton
                ba_help_button = True
            
            Case Is >= vbSystemModal                    ' 4096     not implemented
                ba_arg = ba_arg - vbSystemModal
                ba_system_modal = True
            
            Case Is >= vbDefaultButton4                 ' 768
                ba_arg = ba_arg - vbDefaultButton4
                ba_default_button = 4
            Case Is >= vbDefaultButton3                 ' 512
                ba_arg = ba_arg - vbDefaultButton3
                ba_default_button = 3
            
            Case Is >= vbDefaultButton2                 ' 256
                ba_arg = ba_arg - vbDefaultButton2
                ba_default_button = 2
            
            Case Is >= vbInformation                    ' 64
                ba_arg = ba_arg - vbInformation
                ba_information = True
            
            Case Is >= vbExclamation                    ' 48
                ba_arg = ba_arg - vbExclamation
                ba_exclamation = True
            
            Case Is >= vbQuestion                       ' 32
                ba_arg = ba_arg - vbQuestion
                ba_question = True
            
            Case Is >= vbCritical                       ' 16
                ba_arg = ba_arg - vbCritical
                ba_critical = True
        End Select
    Wend
    BttnsArgs = ba_arg

End Function

Private Function BttnsNo(ByVal v As Variant) As Long
    Select Case v
        Case vbYesNo, vbRetryCancel, vbResumeOk:    BttnsNo = 2
        Case vbAbortRetryIgnore, vbYesNoCancel:     BttnsNo = 3
        Case Else:                                  BttnsNo = 1
    End Select
End Function

Public Function Buttons(ParamArray bttns() As Variant) As Collection
' --------------------------------------------------------------------------
' Returns the provided items (bttns) as Collection. If an item is a
' Collection its items are extracted and included at the corresponding
' position. When the consequtive number of buttons exceeds 7 a vbLf is
' included to indicate a new row. When the number of rows is exieeded any
' subsequent items are ignored.
' --------------------------------------------------------------------------
    Const PROC          As String = "Buttons"
    
    On Error GoTo eh
    Static StackItems   As Collection
    Static QueueResult  As Collection
    Static cllResult    As Collection
    Static lBttnsInRow  As Long         ' buttons in a row counter (excludes break items)
    Static lBttns       As Long         ' total buttons in cllAdd
    Static lRows        As Long         ' button rows counter
    Static SubItemsDone As Long
    Dim cll             As Collection
    Dim dct             As Dictionary
    Dim i               As Long
    Dim sDelimiter      As String
    
    If cllResult Is Nothing Then
        Set StackItems = New Collection
        Set QueueResult = New Collection
        Set cllResult = New Collection
        lBttnsInRow = 0
        lBttns = 0
        lRows = 0
        SubItemsDone = 0
    End If
    If UBound(bttns) = -1 Then GoTo xt
    If UBound(bttns) = 0 Then
        If TypeName(bttns(0)) = "Nothing" Then GoTo xt
        '~~ When only one item is provided it may be a Collection, a Dictionary, a single string or numeric item, or
        '~~ a string with comma or semicolon delimited items
        If lRows > 7 Then GoTo xt
        If TypeName(bttns(0)) = "Collection" Then
            Set cll = bttns(0)
            For i = cll.Count To 1 Step -1
                StckPush StackItems, cll(i)
            Next i
        ElseIf TypeName(bttns(0)) = "Dictionary" Then
            Set dct = bttns(0)
            For i = dct.Count - 1 To 0 Step -1
                StckPush StackItems, dct.Items()(i)
            Next i
        ElseIf IsNumeric(bttns(0)) _
            Or (TypeName(bttns(0)) = "String" And bttns(0) <> vbNullString) Then
            '~~ Any other item but Collection, Numeric or String is ignored
            Select Case bttns(0)
                Case vbLf, vbCr, vbCrLf
                    If lRows < 7 And lBttnsInRow <> 0 Then
                        '~~ Exceeding rows or empty rows are ignored
                        cllResult.Add bttns(0)
                        lBttnsInRow = 0
                        lRows = lRows + 1
                    End If
                Case Else
                    '~~ The string may still be a comma or semicolon delimited string of items
                    sDelimiter = vbNullString
                    If InStr(bttns(0), ",") <> 0 Then sDelimiter = ","
                    If InStr(bttns(0), ";") <> 0 Then sDelimiter = ";"
                    If sDelimiter <> vbNullString Then
                        '~~ The comma or semicolon delimited items are pushed on the stack in reverse order
                        For i = UBound(Split(bttns(0), sDelimiter)) To 0 Step -1
                            StckPush StackItems, Trim(Split(bttns(0), sDelimiter)(i))
                        Next i
                    Else
                        '~~ This is a single buttons caption specified by a numeric value or a string
                        If lRows = 0 Then lRows = 1
                        
                        If lRows < 7 _
                        And lBttnsInRow + BttnsNo(bttns(0)) > 7 Then
                            '~~ Insert a row break
                            cllResult.Add vbLf
                            lRows = lRows + 1
                            lBttnsInRow = 0
                        End If
                        If lRows <= 7 _
                        And lBttnsInRow + BttnsNo(bttns(0)) <= 7 Then
                            '~~ Any excessive buttons spec is ignored
                            If bttns(0) = "B50" Then Stop
                            cllResult.Add bttns(0)
                            lBttnsInRow = lBttnsInRow + BttnsNo(bttns(0))
                        End If
                    End If
            End Select
        End If
        ' items other than Collection, Dictionary, Numeric or String are ignored
    Else
        '~~ More than one item in ParamArray
        For i = UBound(bttns) To 0 Step -1
            StckPush StackItems, bttns(i)
        Next i
    End If
    
    While Not StckIsEmpty(StackItems)
        Set cllResult = Buttons(StckPop(StackItems))
    Wend

xt: If Not StckIsEmpty(StackItems) Then Exit Function
    Set Buttons = cllResult
    Set cllResult = Nothing
    Set StackItems = Nothing
    Exit Function
        
eh: If ErrMsg(ErrSrc(PROC)) = vbYes Then: Stop: Resume
End Function

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
             Optional ByVal dsply_button_reply_with_index As Boolean = False, _
             Optional ByVal dsply_modeless As Boolean = False, _
             Optional ByVal dsply_width_min As Long = 15, _
             Optional ByVal dsply_width_max As Long = 85, _
             Optional ByVal dsply_height_min As Long = 25, _
             Optional ByVal dsply_height_max As Long = 85, _
             Optional ByVal dsply_pos As Variant = 3) As Variant
' ------------------------------------------------------------------------------
' Common VBA Message Display: A service using the Common VBA Message Form as an
' alternative to the VBA.MsgBox.
'
' Argument                      | Description
' ----------------------------- + ----------------------------------------------
' dsply_title                   | String, Title
' dsply_msg                     | UDT, Message
' dsply_buttons                 | Button captions as Collection
' dsply_button_default          | Default button, either the index or the
'                               | caption, defaults to 1 (= the first displayed
'                               | button)
' dsply_button_reply_with_index | Defaults to False, when True the index of the
'                               | of the pressed button is returned else the
'                               | caption or the VBA.MsgBox button value
'                               | respectively
' dsply_modeless                | The message is displayed modeless, defaults
'                               | to False = vbModal
' dsply_width_min               | Overwrites the default when not 0
' dsply_width_max               | Overwrites the default when not 0
' dsply_height_max              | Overwrites the default when not 0
' dsply_button_width_min       | Overwrites the default when not 0
'
' See: https://github.com/warbe-maker/Common-VBA-Message-Service
'
' W. Rauschenberger, Berlin, Apr 2022
' ------------------------------------------------------------------------------
    Const PROC = "Dsply"
    
    On Error GoTo eh
    Dim i       As Long
    Dim MsgForm As fMsg

#If ExecTrace = 1 Then
    mTrc.Pause
#End If
    
    If Not IsValidMsgButtonsArg(dsply_buttons) _
    Then Err.Raise AppErr(1), ErrSrc(PROC), _
                   "The provided buttons argument is neither empty (defaults to vbOkOnly), a string " & _
                   "(optionally comma separated), a valid VBA.MsgBox value (vbYesNo, vbRetryCancel, " & _
                   "etc. plus any extra options - which may or may not be implemented), an Array, a " & _
                   "Collection, or a Dictionary! When an Array, Collection, or Dictionary at least " & _
                   "one of its items in incorrect!"
    
    AssertWidthAndHeight dsply_width_min _
                       , dsply_width_max _
                       , dsply_height_min _
                       , dsply_height_max
    
    Set MsgForm = MsgInstance(dsply_title)
    
    With MsgForm
        .ReplyWithIndex = dsply_button_reply_with_index
        '~~ Use dimensions when explicitely specified
        If dsply_height_max > 0 Then .MsgHeightMax = dsply_height_max   ' percentage of screen height
        If dsply_width_max > 0 Then .MsgWidthMax = dsply_width_max      ' percentage of screen width
        If dsply_width_min > 0 Then .MsgWidthMin = dsply_width_min      ' defaults to 300 pt. the absolute minimum is 200 pt
        .MsgTitle = dsply_title
        For i = 1 To .NoOfDesignedMsgSects
            '~~ Save the label and the text udt into a Dictionary by transfering it into an array
            .MsgLabel(i) = dsply_msg.Section(i).Label
            .Text(enSectText, i) = dsply_msg.Section(i).Text
        Next i
        
        .MsgBttns = dsply_buttons
        .MsgButtonDefault = dsply_button_default
        '+------------------------------------------------------------------------+
        '|| Setup prior showing the form is much faster and avoids flickering.   ||
        '|| For testing - indicated by VisualizerControls = True and             ||
        '|| dsply_modeless = True - prior Setup is suspended.                    ||
        '+------------------------------------------------------------------------+
        If Not .VisualizeForTest Then .Setup
        If dsply_modeless Then
            DisplayDone = False
            .Show vbModeless
            .PositionOnScreen dsply_pos
        Else
            .PositionOnScreen dsply_pos
            .Show vbModal
        End If
    End With
    Dsply = RepliedWith
    
xt:
#If ExecTrace = 1 Then
    mTrc.Continue
#End If
    Exit Function

eh: If ErrMsg(ErrSrc(PROC)) = vbYes Then: Stop: Resume
End Function

Public Function ErrMsg(ByVal err_source As String, _
              Optional ByVal err_number As Long = 0, _
              Optional ByVal err_dscrptn As String = vbNullString, _
              Optional ByVal err_line As Long = 0, _
              Optional ByVal err_buttons As Variant = vbOKOnly, _
              Optional ByVal err_pos As Variant = 3) As Variant
' ------------------------------------------------------------------------------
' Displays an error message.
'
' W. Rauschenberger, Berlin, Nov 2020
' ------------------------------------------------------------------------------
    
    Dim ErrNo       As Long
    Dim ErrDesc     As String
    Dim ErrType     As String
    Dim ErrAtLine   As String
    Dim ErrMsgText  As TypeMsg
    Dim ErrAbout    As String
    Dim ErrTitle    As String
    Dim ErrButtons  As Collection
    
    '~~ Obtain error information from the Err object for any argument not provided
    If err_number = 0 Then err_number = Err.Number
    If err_line = 0 Then err_line = Erl
    If err_source = vbNullString Then err_source = Err.source
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
    ErrTitle = ErrType & ErrNo & " in: '" & err_source & "'" & ErrAtLine
    Debug.Print ErrTitle
    
    '~~ Prepare the Error Reply Buttons
#If Debugging = 1 Then
    Set ErrButtons = mMsg.Buttons(vbResumeOk)
#Else
    Set ErrButtons = mMsg.Buttons(err_buttons)
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
            .Text = "About 'Resume Error Line':"
            .FontColor = rgbBlue
        End With
        .Text.Text = "The additional debugging option button is displayed because the " & _
                     "Conditional Compile Argument 'Debugging = 1'. Pressing this button " & _
                     "and twice F8 ends up at the code line which raised the error"
    End With
#End If
    mMsg.Dsply dsply_title:=ErrTitle _
             , dsply_msg:=ErrMsgText _
             , dsply_buttons:=ErrButtons _
             , dsply_pos:=err_pos
    ErrMsg = mMsg.RepliedWith
    
End Function

Private Function ErrSrc(ByVal sProc As String) As String
    ErrSrc = "mMsg." & sProc
End Function

Private Function GetPanesIndex(ByVal Rng As Range) As Integer
    Dim sR As Long:          sR = ActiveWindow.SplitRow
    Dim sc As Long:          sc = ActiveWindow.SplitColumn
    Dim r As Long:            r = Rng.row
    Dim c As Long:            c = Rng.Column
    Dim Index As Integer: Index = 1

    Select Case True
    Case sR = 0 And sc = 0: Index = 1
    Case sR = 0 And sc > 0 And c > sc: Index = 2
    Case sR > 0 And sc = 0 And r > sR: Index = 2
    Case sR > 0 And sc > 0 And r > sR: If c > sc Then Index = 4 Else Index = 3
    Case sR > 0 And sc > 0 And c > sc: If r > sR Then Index = 4 Else Index = 2
    End Select

    GetPanesIndex = Index
End Function

Public Function IsValidMsgButtonsArg(ByVal v_arg As Variant) As Boolean
' -------------------------------------------------------------------------------------
' Returns TRUE when the buttons argument (v_arg) is valid. When v_arg is an Array,
' a Collection, or a Dictionary, TRUE is returned when all items are valid.
' -------------------------------------------------------------------------------------
    Dim i As Long
    Dim v As Variant
    
    Select Case VarType(v_arg)
        Case vbString, vbEmpty
            IsValidMsgButtonsArg = True
        Case Else
            Select Case True
                Case IsArray(v_arg), TypeName(v_arg) = "Collection", TypeName(v_arg) = "Dictionary"
                     For Each v In v_arg
                        If Not IsValidMsgButtonsArg(v) Then Exit Function
                     Next v
                    IsValidMsgButtonsArg = True
                Case IsNumeric(v_arg)
                    Select Case BttnsArgs(v_arg) ' The numeric buttons argument with all additional option 'unstripped'
                        Case vbOKOnly, vbOKCancel, vbYesNo, vbRetryCancel, vbYesNoCancel, vbAbortRetryIgnore, vbYesNo, vbResumeOk
                            IsValidMsgButtonsArg = True
                    End Select
            End Select
    End Select

End Function

Public Sub MakeFormResizable()
' ----------------------------------------------------------------------------
' Written: February 14, 2011
' Author:  Leith Ross
'
' NOTE:  This code should be executed within the UserForm_Activate() event.
' ----------------------------------------------------------------------------
    Dim lStyle As Long
    Dim hWnd As Long
    Dim RetVal
  
    hWnd = GetForegroundWindow
    'Get the basic window style
     lStyle = GetWindowLong(hWnd, GWL_STYLE) Or WS_THICKFRAME
    'Set the basic window styles
     RetVal = SetWindowLong(hWnd, GWL_STYLE, lStyle)

End Sub

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

Public Sub Monitor(ByVal mon_title As String, _
                   ByRef mon_text As TypeMsgText, _
          Optional ByVal mon_steps_displayed As Long = 10, _
          Optional ByVal mon_height_max As Long = 80, _
          Optional ByVal mon_pos As Variant = 3, _
          Optional ByVal mon_steps_monospaced As Boolean = False, _
          Optional ByVal mon_width_max As Long = 80, _
          Optional ByVal mon_width_min As Long = 30)
' ------------------------------------------------------------------------------
' Service for the monitoring of a process step. The title (mon_title) identifies
' the instance of the fMsg UserForm, the process monitored respectively.
' ------------------------------------------------------------------------------
    Const PROC = "Monitor"
    
    On Error GoTo eh
    Set fMonitor = MsgInstance(mon_title)
    With fMonitor
        If Not .MonitorIsInitialized Then
            AssertWidthAndHeight width_min:=mon_width_min _
                               , WIDTH_MAX:=mon_width_max _
                               , height_max:=mon_height_max
            .MonitorProcess = mon_title
            .MonitorStepsDisplayed = mon_steps_displayed
            .SetupDone = True ' Bypass regular message setup
            .MsgHeightMax = mon_height_max
            .MsgWidthMax = mon_width_max
            .MsgWidthMin = mon_width_min
            .MonitorInit
            .Show False
            .PositionOnScreen mon_pos
        End If
        .Text(enMonStep) = mon_text
        .MonitorStep
    End With
    
xt: Exit Sub

eh: If ErrMsg(ErrSrc(PROC)) = vbYes Then: Stop: Resume
End Sub

Public Sub MonitorFooter(ByVal mon_title As String, _
                         ByRef mon_text As TypeMsgText, _
                Optional ByVal mon_steps_displayed As Long = 10, _
                Optional ByVal mon_height_max As Long = 80, _
                Optional ByVal mon_pos As String = "5,5", _
                Optional ByVal mon_steps_monospaced As Boolean = False, _
                Optional ByVal mon_width_max As Long = 80, _
                Optional ByVal mon_width_min As Long = 30)
' ------------------------------------------------------------------------------
' Adds or modifies a footer (mon_text) in a monitored process identified by
' the title (mon_title).
' ------------------------------------------------------------------------------
    Const PROC = "MonitorFooter"
    
    On Error GoTo eh
    
    Set fMonitor = MsgInstance(mon_title)
    With fMonitor
        If Not .MonitorIsInitialized Then
            AssertWidthAndHeight width_min:=mon_width_min _
                               , WIDTH_MAX:=mon_width_max _
                               , height_max:=mon_height_max
            .MonitorProcess = mon_title
            .MonitorStepsDisplayed = mon_steps_displayed
            .SetupDone = True ' Bypass regular message setup
            .MsgHeightMax = mon_height_max
            .MsgWidthMax = mon_width_max
            .MsgWidthMin = mon_width_min
            .MonitorInit
            .Show False
            .PositionOnScreen mon_pos
        End If
        .Text(enMonFooter) = mon_text
        .MonitorFooter
    End With
    
xt: Exit Sub

eh: If ErrMsg(ErrSrc(PROC)) = vbYes Then: Stop: Resume
End Sub

Public Sub MonitorHeader(ByVal mon_title As String, _
                         ByRef mon_text As TypeMsgText, _
                Optional ByVal mon_steps_displayed As Long = 10, _
                Optional ByVal mon_height_max As Long = 80, _
                Optional ByVal mon_pos As String = "5,5", _
                Optional ByVal mon_steps_monospaced As Boolean = False, _
                Optional ByVal mon_width_max As Long = 80, _
                Optional ByVal mon_width_min As Long = 30)
' ------------------------------------------------------------------------------
'
' ------------------------------------------------------------------------------
    Const PROC = "MonitorHeader"
    
    On Error GoTo eh
    
    Set fMonitor = MsgInstance(mon_title)
    With fMonitor
        If Not .MonitorIsInitialized Then
            AssertWidthAndHeight width_min:=mon_width_min _
                               , WIDTH_MAX:=mon_width_max _
                               , height_max:=mon_height_max
            .MonitorProcess = mon_title
            .MonitorStepsDisplayed = mon_steps_displayed
            .SetupDone = True ' Bypass regular message setup
            .MsgHeightMax = mon_height_max
            .MsgWidthMax = mon_width_max
            .MsgWidthMin = mon_width_min
            .MonitorInit
            .Show False
            .PositionOnScreen mon_pos
        End If
        .Text(enMonHeader) = mon_text
        .MonitorHeader
    End With
    
xt: Exit Sub

eh: If ErrMsg(ErrSrc(PROC)) = vbYes Then: Stop: Resume
End Sub

'Public Function MonitorInit(ByVal mon_title As String, _
'                   Optional ByVal mon_steps_displayed As Long = 10, _
'                   Optional ByVal mon_height_max As Long = 80, _
'                   Optional ByVal mon_pos As Range = Nothing, _
'                   Optional ByVal mon_steps_monospaced As Boolean = False, _
'                   Optional ByVal mon_width_max As Long = 80, _
'                   Optional ByVal mon_width_min As Long = 30) As fMsg
'' ------------------------------------------------------------------------------
'' Establish a monitor window for n (mon_steps) steps by creating the
'' corresponding number of - st first invisible - text boxes
'' ------------------------------------------------------------------------------
'    Const PROC = "MonitorInit"
'
'    On Error GoTo eh
'    Dim t       As TypeMsgText
'
'    AssertWidthAndHeight width_min:=mon_width_min _
'                       , WIDTH_MAX:=mon_width_max _
'                       , height_max:=mon_height_max
'
'    Set fMonitor = mMsg.MsgInstance(mon_title)
'    With fMonitor
'        .MonitorProcess = mon_title
'        .MonitorStepsDisplayed = mon_steps_displayed
'        .SetupDone = True ' Bypass regular message setup
'        .MsgHeightMax = mon_height_max
'        .MsgWidthMax = mon_width_max
'        .MsgWidthMin = mon_width_min
'        .MonitorInit
'        .PositionOnScreen = mon_pos
'    End With
'    fMonitor.Show False
'    Set MonitorInit = fMonitor
'
'xt: Exit Function
'
'eh: If ErrMsg(ErrSrc(PROC)) = vbYes Then: Stop: Resume
'End Function

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
    Static cyStart      As Currency
    Static Instances    As Dictionary    ' Collection of (possibly still)  active form instances
    Dim MsecsElapsed    As Currency
    Dim MsecsWait       As Long
    
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
        '~~ When there is no evidence of an already existing instance a new one is established.
        '~~ In order not to interfere with any prior established instance a minimum wait time
        '~~ of 10 milliseconds is maintained.
        MsecsElapsed = (Time() - cyStart) / CDec(TimeTicksFrequency)
        MsecsWait = 10 - MsecsElapsed
        If MsecsWait > 0 Then Sleep MsecsWait
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

Public Function ShellRun(ByVal oue_string As String, _
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

Private Sub ShowAtRange(ByVal sar_form As Object, _
                        ByVal sar_rng As Range)
' ----------------------------------------------------------------------------
'
' ----------------------------------------------------------------------------
    Dim PosLeft As Single
    Dim PosTop  As Single

    If ActiveWindow.FreezePanes Then
       PosLeft = ActiveWindow.Panes(GetPanesIndex(sar_rng)).PointsToScreenPixelsX(sar_rng.Left)
       PosTop = ActiveWindow.Panes(GetPanesIndex(sar_rng)).PointsToScreenPixelsY(sar_rng.Top + sar_rng.Height)
    Else
       PosLeft = ActiveWindow.ActivePane.PointsToScreenPixelsX(sar_rng.Left)
       PosTop = ActiveWindow.ActivePane.PointsToScreenPixelsY(sar_rng.Top + sar_rng.Height)
    End If

    ConvertPixelsToPoints PosLeft, PosTop, PosLeft, PosTop

    With sar_form
       .StartupPosition = 0
       .Left = PosLeft
       .Top = PosTop
    End With

End Sub

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

Private Function StckIsEmpty(ByVal stck As Collection) As Boolean
' ----------------------------------------------------------------------------
' Common Stack Empty check service. Returns True when either there is no stack
' (stck Is Nothing) or when the stack is empty (items count is 0).
' ----------------------------------------------------------------------------
    StckIsEmpty = stck Is Nothing
    If Not StckIsEmpty Then StckIsEmpty = stck.Count = 0
End Function

Private Function StckPop(ByVal stck As Collection) As Variant
' ----------------------------------------------------------------------------
' Common Stack Pop service. Returns the last item pushed on the stack (stck)
' and removes the item from the stack. When the stack (stck) is empty a
' vbNullString is returned.
' ----------------------------------------------------------------------------
    Const PROC = "StckPop"
    
    On Error GoTo eh
    If StckIsEmpty(stck) Then GoTo xt
    
    On Error Resume Next
    Set StckPop = stck(stck.Count)
    If Err.Number <> 0 _
    Then StckPop = stck(stck.Count)
    stck.Remove stck.Count

xt: Exit Function

eh: If ErrMsg(ErrSrc(PROC)) = vbYes Then: Stop: Resume
End Function

Private Sub StckPush(ByRef stck As Collection, _
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

Private Function Time() As Currency:                getTickCount Time:                  End Function

Private Function TimeTicksFrequency() As Currency:  getFrequency TimeTicksFrequency:    End Function

