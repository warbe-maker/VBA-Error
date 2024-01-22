Attribute VB_Name = "mMsg"
Option Explicit
' ------------------------------------------------------------------------------
' Standard Module mMsg: Message display services using the fMsg form.
' =====================
' Public services:
' ----------------
' Box         In analogy to the MsgBox, provides a simple message but with all
'             the fexibility for the display of up to 49 reply buttons.
' Buttons     Supports the specification of the buttons displayed in a matrix
'             of 7 x 7 buttons (max 7 buttons in max 7 rows)
' Dsply       Exposes all properties and methods for the display of any kind
'             of message
' Monitor     Uses modeless instances of the fMsg form - any instance is
'             identified by the window title - to display the progress of a
'             process or monitor intermediate results.
' MsgInstance Creates (when not existing) and returns an fMsg object
'             identified by the Title
'
' Uses:       fMsg
'
' Requires:   Reference to "Microsoft Scripting Runtime"
'
' W. Rauschenberger, Berlin Jan 2024
' See: https://github.com/warbe-maker/VBA-Message
' ------------------------------------------------------------------------------
Public Const MSG_LIMIT_WIDTH_MIN_PERCENTAGE     As Long = 15
Public Const MSG_LIMIT_WIDTH_MAX_PERCENTAGE     As Long = 95
Public Const MSG_LIMIT_HEIGHT_MIN_PERCENTAGE    As Long = 20
Public Const MSG_LIMIT_HEIGHT_MAX_PERCENTAGE    As Long = 80

' Extension of the VBA.MsgBox constants for the Debugging option of the ErrMsg service
' to display additional debugging buttons
Public Const vbResumeOk                         As Long = 7 ' Buttons value in mMsg.ErrMsg (pass on not supported)
Public Const vbResume                           As Long = 6 ' return value (equates to vbYes)

Public Enum enScreen
    enHeightDPI         ' VerticalResolution (pixelsY)"
    enHeightInches      ' HeightInches (inchesY)
    enHeightPPI         ' PixelsPerInchY (ppiY)
    enHeightWinDPI      ' WinDPIy (dpiY)
    
    enWidthDPI          ' HorizontalResolution (pixelsX)
    enWidthInches       ' WidthInches (inchesX)
    enWidthPPI          ' PixelsPerInchX (ppiX)
    enWidthWinDPI       ' WinDPIx (dpiX)
    
    enAdjustmentfactor  ' AdjustmentFactor (zoomFac)
    enDiagonalInches    ' DiagonalInches (inchesDiag)
    enDiagonalPPI       ' PixelsPerInch (ppiDiag)
    enDisplayName       ' DisplayName
    enHelp              ' Help
    enIsPrimary         ' IsPrimary
    enUpdate            ' Update
    enWinDPI            ' WinDPI (dpiWin)
End Enum

Public Enum enLabelPos ' pending implementation
    enLabelAboveSectionText = 1
    enLposLeftAlignedRight
    enLposLeftAlignedLeft
    enLposLeftAlignedCenter
End Enum

Public Enum enDsplyDimension
    enDsplyDimensionWidth
    enDsplyDimensionHeight
End Enum

Public Type udtMsgLabel
        FontBold        As Boolean
        FontColor       As XlRgbColor
        FontItalic      As Boolean
        FontName        As String
        FontSize        As Long
        FontUnderline   As Boolean
        MonoSpaced      As Boolean  ' FontName defaults to "Courier New"
        Text            As String
        OnClickAction   As String   ' this extra option is only available when the control is implemented as msForms.Label
End Type

Public Type udtMsgText
        FontBold        As Boolean
        FontColor       As XlRgbColor
        FontItalic      As Boolean
        FontName        As String
        FontSize        As Long
        FontUnderline   As Boolean
        MonoSpaced      As Boolean  ' FontName defaults to "Courier New"
        Text            As String
        OnClickAction   As String   ' this extra option is only available when the control is implemented as msForms.Label
End Type

Public Type udtMsgSect:    Label As udtMsgLabel:  Text As udtMsgText:   End Type
Public Type udtMsg:        Section(1 To 8) As udtMsgSect:               End Type '!!! 8 = Public Property NoOfMsgSects !!!

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
Public RepliedWith          As Variant  ' provided by the UseForm when a button has been pressed/clicked

Private Const SM_CMONITORS              As Long = 80    ' number of display monitors
Private Const MONITOR_CCHDEVICENAME     As Long = 32    ' device name fixed length
Private Const MONITOR_PRIMARY           As Long = 1
Private Const MONITOR_DEFAULTTONULL     As Long = 0
Private Const MONITOR_DEFAULTTOPRIMARY  As Long = 1
Private Const MONITOR_DEFAULTTONEAREST  As Long = 2
Private Type RECT
    Left As Long
    Top As Long
    Right As Long
    Bottom As Long
End Type
Private Type MONITORINFOEX
   cbSize As Long
   rcMonitor As RECT
   rcWork As RECT
   dwFlags As Long
   szDevice As String * MONITOR_CCHDEVICENAME
End Type
Private Enum DevCap     ' GetDeviceCaps nIndex (video displays)
    HORZSIZE = 4        ' width in millimeters
    VERTSIZE = 6        ' height in millimeters
    HORZRES = 8         ' width in pixels
    VERTRES = 10        ' height in pixels
    BITSPIXEL = 12      ' color bits per pixel
    LOGPIXELSX = 88     ' horizontal DPI (assumed by Windows)
    LOGPIXELSY = 90     ' vertical DPI (assumed by Windows)
    COLORRES = 108      ' actual color resolution (bits per pixel)
    VREFRESH = 116      ' vertical refresh rate (Hz)
End Enum

Private Const ERROR_BAD_FORMAT = 11&
Private Const ERROR_FILE_NOT_FOUND = 2&
Private Const ERROR_NO_ASSOC = 31&
Private Const ERROR_OUT_OF_MEM = 0&
Private Const ERROR_PATH_NOT_FOUND = 3&
Private Const ERROR_SUCCESS = 32&
Private Const GITHUB_REPO_URL       As String = "https://github.com/warbe-maker/VBA-Message"
Private Const SM_CXVIRTUALSCREEN    As Long = &H4E&     ' calculating
Private Const SM_CYVIRTUALSCREEN    As Long = &H4F&     ' the
Private Const SM_XVIRTUALSCREEN     As Long = &H4C&     ' display's
Private Const SM_YVIRTUALSCREEN     As Long = &H4D&     ' DPI in points
Private Const TWIPSPERINCH          As Long = 1440      ' -------------

Private Declare PtrSafe Function CreateDC Lib "gdi32" Alias "CreateDCA" (ByVal lpDriverName As String, ByVal lpDeviceName As String, ByVal lpOutput As String, lpInitData As LongPtr) As LongPtr
Private Declare PtrSafe Function DeleteDC Lib "gdi32" (ByVal hDC As LongPtr) As Long
Private Declare PtrSafe Function GetActiveWindow Lib "user32" () As LongPtr
Private Declare PtrSafe Function GetDC Lib "user32" (ByVal hWnd As LongPtr) As LongPtr
Private Declare PtrSafe Function GetDeviceCaps Lib "gdi32" (ByVal hDC As LongPtr, ByVal nIndex As Long) As Long
Private Declare PtrSafe Function getFrequency Lib "kernel32" Alias "QueryPerformanceFrequency" (TimerSystemFrequency As Currency) As Long
Private Declare PtrSafe Function GetMonitorInfo Lib "user32" Alias "GetMonitorInfoA" (ByVal hMonitor As LongPtr, ByRef lpMI As MONITORINFOEX) As Boolean
Private Declare PtrSafe Function GetSystemMetrics Lib "user32" (ByVal nIndex As Long) As Long
Private Declare PtrSafe Function getTickCount Lib "kernel32" Alias "QueryPerformanceCounter" (cyTickCount As Currency) As Long
Private Declare PtrSafe Function MonitorFromWindow Lib "user32" (ByVal hWnd As LongPtr, ByVal dwFlags As Long) As LongPtr
Private Declare PtrSafe Function ReleaseDC Lib "user32" (ByVal hWnd As LongPtr, ByVal hDC As LongPtr) As Long

#If VBA7 Then
    Private Declare PtrSafe Sub Sleep Lib "kernel32" (ByVal ms As LongPtr)
#Else
    Private Declare PtrSafe Sub Sleep Lib "kernel32" (ByVal ms As Long)
#End If
Private Declare PtrSafe Function GetSystemMetrics32 Lib "user32" Alias "GetSystemMetrics" (ByVal nIndex As Long) As Long
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
Private bModeLess           As Boolean
Private lPixelsPerInchX     As Long
Private lPixelsPerInchY     As Long
Private fMonitor            As fMsg
Public MsgInstances         As Dictionary    ' Collection of (possibly still)  active form instances

Public Property Get DsplyWidthDPI() As Variant:         DsplyWidthDPI = Screen(enWidthDPI):                                 End Property

Public Property Get DsplyHeightDPI() As Variant:        DsplyHeightDPI = Screen(enHeightDPI):                               End Property

Public Property Get DsplyHeightPT() As Single:          DsplyDPItoPT DsplyHeightDPI, DsplyWidthDPI, x_pts:=DsplyHeightPT:   End Property

Public Property Get DsplyWidthPT() As Single:           DsplyDPItoPT DsplyHeightDPI, DsplyWidthDPI, y_pts:=DsplyWidthPT:    End Property

Private Property Get ModeLess() As Boolean:             ModeLess = bModeLess:                                               End Property

Private Property Let ModeLess(ByVal b As Boolean):      bModeLess = b:                                                      End Property

Public Property Get NoOfMsgSects() As Long:             NoOfMsgSects = 8:                                                   End Property

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

Public Sub AssertWidthAndHeight(Optional ByRef a_width_min As Long = 0, _
                                Optional ByRef a_width_max As Long = 0, _
                                Optional ByRef a_height_min As Long = 0, _
                                Optional ByRef a_height_max As Long = 0)
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

    '~~ Convert all default limits from percentage - i.e. a value < 100 - to pt
    Dim MsgMaxWidthLimitPt  As Long:    MsgMaxWidthLimitPt = ValueAsPt(MSG_LIMIT_WIDTH_MAX_PERCENTAGE, mMsg.enDsplyDimensionWidth)
    Dim MsgMinWidthLimitPt  As Long:    MsgMinWidthLimitPt = ValueAsPt(MSG_LIMIT_WIDTH_MIN_PERCENTAGE, mMsg.enDsplyDimensionWidth)
    Dim MsgMaxHeightLimitPt As Long:    MsgMaxHeightLimitPt = ValueAsPt(MSG_LIMIT_HEIGHT_MAX_PERCENTAGE, mMsg.enDsplyDimensionHeight)
    Dim MsgMinHeightLimitPt As Long:    MsgMinHeightLimitPt = ValueAsPt(MSG_LIMIT_HEIGHT_MIN_PERCENTAGE, mMsg.enDsplyDimensionHeight)
    
    '~~ Convert all percentage arguments - i.e. a value < 100 - to pt arguments
    If a_width_max <> 0 And a_width_max <= 100 Then a_width_max = ValueAsPt(a_width_max, mMsg.enDsplyDimensionWidth)
    If a_width_min <> 0 And a_width_min <= 100 Then a_width_min = ValueAsPt(a_width_min, mMsg.enDsplyDimensionWidth)
    If a_height_max <> 0 And a_height_max <= 100 Then a_height_max = ValueAsPt(a_height_max, mMsg.enDsplyDimensionHeight)
    If a_height_min <> 0 And a_height_min <= 100 Then a_height_min = ValueAsPt(a_height_min, mMsg.enDsplyDimensionHeight)
        
    '~~ Provide sensible values for all values invalid, improper, or useless
    If a_width_min > a_width_max Then a_width_min = a_width_max
    If a_height_min > a_height_max Then a_height_min = a_height_max
    If a_width_min < MsgMinWidthLimitPt Then a_width_min = MsgMinWidthLimitPt
    If a_width_max <= a_width_min Then a_width_max = a_width_min
    If a_width_max > MsgMaxWidthLimitPt Then a_width_max = MsgMaxWidthLimitPt
    If a_height_min < MsgMinHeightLimitPt Then a_height_min = MsgMinHeightLimitPt
    If a_height_max = 0 Or a_height_max < a_height_min Then a_height_max = a_height_min
    If a_height_max > MsgMaxHeightLimitPt Then a_height_max = MsgMaxHeightLimitPt
    
End Sub

Public Function Box(ByVal Prompt As String, _
           Optional ByVal Buttons As Variant = vbOKOnly, _
           Optional ByVal Title As String = vbNullString, _
           Optional ByVal box_buttons_app_run As Dictionary = Nothing, _
           Optional ByVal box_monospaced As Boolean = False, _
           Optional ByVal box_button_default = 1, _
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
    Dim Message As udtMsgText
    Dim MsgForm As fMsg

    If Not BttnArgsAreValid(Buttons) _
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
        .MsgTitle = Title
        .MsgText(enSectText, 1) = Message
        .MsgBttns = mMsg.Buttons(Buttons)   ' Provide the buttons as Collection
        
        '~~ All width and height specifications by the user are "outside" dimensions !
        .FormHeightOutsideMax = mMsg.ValueAsPt(box_height_max, enDsplyDimensionHeight)  ' percentage of screen height in pt
        .FormHeightOutsideMin = mMsg.ValueAsPt(box_height_min, enDsplyDimensionHeight)  ' percentage of screen height in pt
        .FormWidthOutsideMax = mMsg.ValueAsPt(box_width_max, enDsplyDimensionWidth)     ' percentage of screen width in pt
        .FormWidthOutsideMin = mMsg.ValueAsPt(box_width_min, enDsplyDimensionWidth)     ' percentage of screen width in pt
        
        .MsgButtonDefault = box_button_default
        .ModeLess = box_modeless
        If box_buttons_app_run Is Nothing Then Set box_buttons_app_run = New Dictionary
        .ApplicationRunArgs = box_buttons_app_run
        '+------------------------------------------------------------------------+
        '|| Setup prior showing the form is much faster and avoids flickering.   ||
        '|| For testing purpose it may be appropriate to out-comment the Setup.  ||
        '+------------------------------------------------------------------------+
        .Setup
        If box_modeless Then
            .Show vbModeless
            .PositionOnScreen box_pos
        Else
            .PositionOnScreen box_pos
            .Show vbModal
            Box = mMsg.RepliedWith
        End If
    End With

xt: Exit Function

eh: If ErrMsg(ErrSrc(PROC)) = vbYes Then: Stop: Resume
End Function

Public Sub BttnAppRun(ByRef b_dct As Dictionary, _
                      ByVal b_button As String, _
                      ByVal b_wb As Workbook, _
                      ByVal b_service_name As String, _
                 ParamArray b_arguments() As Variant)
' --------------------------------------------------------------------------
' Returns a Dictionary (b_dct) with Application.Run information for the
' button identified by its caption string (b_button) added with the
' button's caption as the key and all other arguments (b_wb,
' b_service_name, b_arguments) as Collection as item.
'
' Attention!
' - Application.Run supports only positional arguments. When only some of
'   the optional arguments are used only those after the last one may be
'   omitted but not those in between. An error is raised when empty
'   arguments are dedected.
' - When Run information is provided for a button already existing in the
'   Dictionary (b_dct) it is replaced.
' - When the message form is displayed "Modal", which is the default, any
'   provided Application.Run information is ignored.
' --------------------------------------------------------------------------
    Const PROC = "BttnAppRun"
    
    On Error GoTo eh
    Dim v   As Variant
    Dim cll As New Collection
    
    If b_dct Is Nothing Then Set b_dct = New Dictionary
    
    cll.Add b_wb
    cll.Add b_service_name
    For Each v In b_arguments
        If TypeName(v) = "Error" Then
            Err.Raise Number:=AppErr(1) _
                    , Source:=ErrSrc(PROC) _
                    , Description:="The ParamArray argument (b_arguments) contains empty elements but empty elements " & _
                                   "are not supported/possible!" & "||" & _
                                   "Application.Run supports only positional but not named arguments. When only some of " & _
                                   "the optional arguments of the called service are used only those after the last one " & _
                                   "may be omitted but not those in between."
        Else
            cll.Add v
        End If
    Next v
    If b_dct.Exists(b_button) Then b_dct.Remove b_button
    b_dct.Add b_button, cll
    Set cll = Nothing
    
xt: Exit Sub

eh: Select Case ErrMsg(ErrSrc(PROC))
        Case vbResume:  Stop: Resume
        Case Else:      GoTo xt
    End Select
End Sub

Public Function BttnArg(ByVal b_arg As Long, _
                 Optional ByRef b_rtl_reading As Boolean, _
                 Optional ByRef b_box_right As Boolean, _
                 Optional ByRef b_set_foreground As Boolean, _
                 Optional ByRef b_help_button As Boolean, _
                 Optional ByRef b_system_modal As Boolean, _
                 Optional ByRef b_default_button As Long, _
                 Optional ByRef b_information As Boolean, _
                 Optional ByRef b_exclamation As Boolean, _
                 Optional ByRef b_question As Boolean, _
                 Optional ByRef b_critical As Boolean) As Long
' -------------------------------------------------------------------------------------
' Returns the Buttons argument (b_arg) with all the options removed by returning them
' as optional arguments. In order to mimic the Buttons argument of the VBA.MsgBox any
' values added for other options but the display of the buttons are unstripped/deducted.
' I.e. the values are deducted and the corresponding argument is returtned instead).
' -------------------------------------------------------------------------------------
    Dim l As Long
    
    l = b_arg - (Abs(Int(b_arg / 16) * 16))
    Select Case l
        Case vbOKOnly, vbOKCancel, vbAbortRetryIgnore, vbYesNoCancel, vbYesNo, vbRetryCancel
        Case Else
            BttnArg = l ' may be a wromg value and thus need to be validated further
            Exit Function
    End Select

    While b_arg >= vbCritical                          ' 16
        Select Case b_arg
            '~~ VBA.MsgBox Display options
            Case Is >= vbMsgBoxRtlReading               ' 1048576  not implemented
                b_arg = b_arg - vbMsgBoxRtlReading
                b_rtl_reading = True
            
            Case Is >= vbMsgBoxRight                    ' 524288   not implemented
                b_arg = b_arg - vbMsgBoxRight
                b_box_right = True
            
            Case Is >= vbMsgBoxSetForeground            ' 65536    not implemented
                b_arg = b_arg - vbMsgBoxSetForeground
                b_set_foreground = True
            
            Case Is >= vbMsgBoxHelpButton               ' 16384    not implemented: Display of a Help button
                b_arg = b_arg - vbMsgBoxHelpButton
                b_help_button = True
            
            Case Is >= vbSystemModal                    ' 4096     not implemented
                b_arg = b_arg - vbSystemModal
                b_system_modal = True
            
            Case Is >= vbDefaultButton4                 ' 768
                b_arg = b_arg - vbDefaultButton4
                b_default_button = 4
            Case Is >= vbDefaultButton3                 ' 512
                b_arg = b_arg - vbDefaultButton3
                b_default_button = 3
            
            Case Is >= vbDefaultButton2                 ' 256
                b_arg = b_arg - vbDefaultButton2
                b_default_button = 2
            
            Case Is >= vbInformation                    ' 64
                b_arg = b_arg - vbInformation
                b_information = True
            
            Case Is >= vbExclamation                    ' 48
                b_arg = b_arg - vbExclamation
                b_exclamation = True
            
            Case Is >= vbQuestion                       ' 32
                b_arg = b_arg - vbQuestion
                b_question = True
            
            Case Is >= vbCritical                       ' 16
                b_arg = b_arg - vbCritical
                b_critical = True
        End Select
    Wend
    BttnArg = b_arg

End Function

Public Function BttnArgsAreValid(ByVal v_arg As Variant) As Boolean
' -------------------------------------------------------------------------------------
' Returns TRUE when all items of the argument (v_arg) are valid, i.e. a string or one
' of the valid MsgBox button values. When the argument is an Array, a Collection, or a
' Dictionary the function is called recursively for each item.
' -------------------------------------------------------------------------------------
    Dim v As Variant
    
    BttnArgsAreValid = VarType(v_arg) = vbString Or VarType(v_arg) = vbEmpty
    If Not BttnArgsAreValid Then
        Select Case True
            Case IsArray(v_arg), TypeName(v_arg) = "Collection", TypeName(v_arg) = "Dictionary"
                 For Each v In v_arg
                    If Not BttnArgsAreValid(v) Then
                        Exit Function
                    End If
                 Next v
                BttnArgsAreValid = True
            Case IsNumeric(v_arg)
                Select Case BttnArg(v_arg) ' The numeric buttons argument with all additional option 'unstripped'
                    Case vbOKOnly, vbOKCancel, vbYesNo, vbRetryCancel, vbYesNoCancel, vbAbortRetryIgnore, vbYesNo, vbResumeOk
                        BttnArgsAreValid = True
                End Select
        End Select
    End If

End Function

Private Function BttnsNo(ByVal v As Variant) As Long
    Select Case v
        Case vbYesNo, vbRetryCancel, vbResumeOk:    BttnsNo = 2
        Case vbAbortRetryIgnore, vbYesNoCancel:     BttnsNo = 3
        Case Else:                                  BttnsNo = 1
    End Select
End Function

Public Function Buttons(ParamArray Bttns() As Variant) As Collection
' --------------------------------------------------------------------------
' Returns the provided items (bttns) as Collection. If an item is a
' Collection its items are extracted and included at the corresponding
' position. When the consequtive number of buttons exceeds 7 a vbLf is
' included to indicate a new row. When the number of rows is exeeded any
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
    If UBound(Bttns) = -1 Then GoTo xt
    If UBound(Bttns) = 0 Then
        If TypeName(Bttns(0)) = "Nothing" Then GoTo xt
        '~~ When only one item is provided it may be a Collection, a Dictionary, a single string or numeric item, or
        '~~ a string with comma or semicolon delimited items
        If lRows > 7 Then GoTo xt
        If TypeName(Bttns(0)) = "Collection" Then
            Set cll = Bttns(0)
            For i = cll.Count To 1 Step -1
                StckPush StackItems, cll(i)
            Next i
        ElseIf TypeName(Bttns(0)) = "Dictionary" Then
            Set dct = Bttns(0)
            For i = dct.Count - 1 To 0 Step -1
                StckPush StackItems, dct.Items()(i)
            Next i
        ElseIf IsNumeric(Bttns(0)) _
            Or (TypeName(Bttns(0)) = "String" And Bttns(0) <> vbNullString) Then
            '~~ Any other item but Collection, Numeric or String is ignored
            Select Case Bttns(0)
                Case vbLf, vbCr, vbCrLf
                    If lRows < 7 And lBttnsInRow <> 0 Then
                        '~~ Exceeding rows or empty rows are ignored
                        cllResult.Add Bttns(0)
                        lBttnsInRow = 0
                        lRows = lRows + 1
                    End If
                Case Else
                    '~~ The string may still be a comma or semicolon delimited string of items
                    sDelimiter = vbNullString
                    If InStr(Bttns(0), ",") <> 0 And InStr(Bttns(0), "\,") = 0 Then
                        '~~ Comma delimited string with the comma not escaped
                        sDelimiter = ","
                    ElseIf InStr(Bttns(0), ";") <> 0 Then
                        sDelimiter = ";"
                    End If
                    If sDelimiter <> vbNullString Then
                        '~~ The comma or semicolon delimited items are pushed on the stack in reverse order
                        For i = UBound(Split(Bttns(0), sDelimiter)) To 0 Step -1
                            StckPush StackItems, Trim(Split(Bttns(0), sDelimiter)(i))
                        Next i
                    Else
                        '~~ This is a single buttons caption specified by a numeric value or a string
                        If lRows = 0 Then lRows = 1
                        
                        If lRows < 7 _
                        And lBttnsInRow + BttnsNo(Bttns(0)) > 7 Then
                            '~~ Insert a row break
                            cllResult.Add vbLf
                            lRows = lRows + 1
                            lBttnsInRow = 0
                        End If
                        If lRows <= 7 _
                        And lBttnsInRow + BttnsNo(Bttns(0)) <= 7 Then
                            '~~ Any excessive buttons spec is ignored
                            If Bttns(0) = "B50" Then Stop
                            cllResult.Add Bttns(0)
                            lBttnsInRow = lBttnsInRow + BttnsNo(Bttns(0))
                        End If
                    End If
            End Select
        End If
        ' items other than Collection, Dictionary, Numeric or String are ignored
    Else
        '~~ More than one item in ParamArray
        For i = UBound(Bttns) To 0 Step -1
            StckPush StackItems, Bttns(i)
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

Private Sub DsplyDPItoPT(Optional ByVal x_dpi As Single, _
                         Optional ByVal y_dpi As Single, _
                         Optional ByRef x_pts As Single, _
                         Optional ByRef y_pts As Single)
' ------------------------------------------------------------------------------
' Returns pixels (device dependent) to points.
' Results verified by: https://pixelsconverter.com/px-to-pt.
' ------------------------------------------------------------------------------
    
    Dim hDC            As Variant
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

Sub DisplayMonitorInfo()
    MsgBox "Monitor Size (dpi) is: " & Screen(enWidthDPI) & " x " & Screen(enHeightDPI), vbInformation, " (width x height dpi) "
End Sub

               
Public Function Dsply(ByVal dsply_title As String, _
                      ByRef dsply_msg As udtMsg, _
             Optional ByVal dsply_Label_spec As String = vbNullString, _
             Optional ByVal dsply_buttons As Variant = vbOKOnly, _
             Optional ByVal dsply_buttons_app_run As Dictionary = Nothing, _
             Optional ByVal dsply_button_default = 1, _
             Optional ByVal dsply_button_reply_with_index As Boolean = False, _
             Optional ByVal dsply_modeless As Boolean = False, _
             Optional ByVal dsply_width_min As Long = 250, _
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
' dsply_Label_spec              | Label width and position
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

#If mTrc = 1 Then
    mTrc.Pause
#ElseIf clsTrc = 1 Then
    Trc.Pause
#End If
    
    If Not BttnArgsAreValid(dsply_buttons) _
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
        .LabelAllSpec = dsply_Label_spec    ' !!! has to be provided first
        .ReplyWithIndex = dsply_button_reply_with_index
        
        '~~ All width and height specifications by the user are "outside" dimensions !
        If dsply_height_max > 0 Then .FormHeightOutsideMax = dsply_height_max   ' percentage of screen height in pt
        If dsply_width_max > 0 Then .FormWidthOutsideMax = dsply_width_max     ' percentage of screen width in pt
        If dsply_width_min > 0 Then .FormWidthOutsideMin = dsply_width_min      ' percentage of screen width in pt
        
        .MsgTitle = dsply_title
        For i = 1 To NoOfMsgSects
            '~~ Save the Label and the text udt into a Dictionary by transfering it into an array
            .MsgLabel(i) = dsply_msg.Section(i).Label
            .MsgText(enSectText, i) = dsply_msg.Section(i).Text
        Next i
        
        If TypeName(dsply_buttons) = "Collection" _
        Then .MsgBttns = dsply_buttons _
        Else .MsgBttns = mMsg.Buttons(dsply_buttons)
        
        .MsgButtonDefault = dsply_button_default
        .ModeLess = dsply_modeless
        If dsply_buttons_app_run Is Nothing Then Set dsply_buttons_app_run = New Dictionary
        .ApplicationRunArgs = dsply_buttons_app_run

        '+------------------------------------------------------------------------+
        '|| Setup prior showing the form is much faster and avoids flickering.   ||
        '|| For testing - indicated by VisualizerControls = True and             ||
        '|| dsply_modeless = True - prior Setup is suspended.                    ||
        '+------------------------------------------------------------------------+
        If Not .VisualizeForTest Then .Setup
        If dsply_modeless Then
            .PositionOnScreen dsply_pos
            .Show vbModeless
        Else
            .PositionOnScreen dsply_pos
            .Show vbModal
        End If
    End With
    Dsply = mMsg.RepliedWith
    
xt:
#If mTrc = 1 Then
    mTrc.Continue
#ElseIf clsTrc = 1 Then
    Trc.Continue
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
    Dim ErrMsgText  As udtMsg
    Dim ErrAbout    As String
    Dim ErrTitle    As String
    Dim ErrButtons  As Collection
    Dim iSect       As Long
    
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
    Set ErrButtons = mMsg.Buttons(vbResumeOk)
        
    '~~ Display the error message by means of the mMsg's Dsply function
    iSect = 1
    With ErrMsgText.Section(iSect)
        With .Label
            .Text = "Error:"
            .FontColor = rgbBlue
            .FontBold = True
        End With
        .Text.Text = ErrDesc
    End With
    
    iSect = iSect + 1
    With ErrMsgText.Section(iSect)
        With .Label
            .Text = "Source:"
            .FontBold = True
            .FontColor = rgbBlue
        End With
        .Text.Text = err_source
    End With
    
    iSect = iSect + 1
    With ErrMsgText.Section(iSect)
        With .Label
            .FontBold = True
            .FontColor = rgbBlue
        End With
        If ErrAbout = vbNullString Then
            .Label.Text = vbNullString
            .Text.Text = vbNullString
        Else
            .Label.Text = "About:"
            .Text.Text = ErrAbout
        End If
    End With
    iSect = iSect + 1
    With ErrMsgText.Section(iSect)
        With .Label
            .Text = "Resume Error" & Chr$(160) & "Line"
            .FontBold = True
        End With
        .Text.Text = "Debugging option! Button displayed because the " & _
                     "Cond. Comp. Argument 'Debugging = 1'. Pressing this button " & _
                     "and twice F8 leads straight to the code line which raised the error."
    End With
    mMsg.Dsply dsply_title:=ErrTitle _
             , dsply_msg:=ErrMsgText _
             , dsply_Label_spec:="R40" _
             , dsply_buttons:=ErrButtons _
             , dsply_pos:=err_pos _
             , dsply_width_min:=15
    ErrMsg = mMsg.RepliedWith
    
End Function

Private Function ErrSrc(ByVal sProc As String) As String
    ErrSrc = "mMsg." & sProc
End Function

Public Function LabelPos(ByVal l_spec As String) As enLabelPos
    Const PROC = "LabelPos"
    
    On Error GoTo eh
    Select Case True
        Case l_spec = vbNullString:     LabelPos = enLabelAboveSectionText:  GoTo xt
        Case InStr(l_spec, "L") <> 0:   LabelPos = enLposLeftAlignedLeft
        Case InStr(l_spec, "C") <> 0:   LabelPos = enLposLeftAlignedCenter
        Case InStr(l_spec, "R") <> 0:   LabelPos = enLposLeftAlignedRight
        Case Else:                      Err.Raise AppErr(1), ErrSrc(PROC), "The Label position specification'l_char is neither a vbNullString (the default = top pos) nor L, R, or C!"
    End Select

xt: Exit Function

eh: If ErrMsg(ErrSrc(PROC)) = vbYes Then: Stop: Resume
End Function

Public Function LabelWidth(ByVal l_spec As String) As Long
    If l_spec <> vbNullString _
    Then LabelWidth = CInt(Replace(Replace(Replace(UCase(l_spec), "L", vbNullString), "C", vbNullString), "R", vbNullString))
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

Public Sub Monitor(ByVal mon_title As String, _
                   ByRef mon_text As udtMsgText, _
          Optional ByVal mon_steps_displayed As Long = 10, _
          Optional ByVal mon_height_max As Long = 80, _
          Optional ByVal mon_pos As Variant = 3, _
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
            AssertWidthAndHeight a_width_min:=mon_width_min _
                               , a_width_max:=mon_width_max _
                               , a_height_max:=mon_height_max
            .MonitorProcess = mon_title
            .MonitorStepsDisplayed = mon_steps_displayed
            .SetupDone = True ' Bypass regular message setup
            .FormHeightOutsideMax = mMsg.ValueAsPt(mon_height_max, enDsplyDimensionHeight)
            .FormWidthOutsideMax = mMsg.ValueAsPt(mon_width_max, enDsplyDimensionWidth)
            .FormWidthOutsideMin = mMsg.ValueAsPt(mon_width_min, enDsplyDimensionWidth)
            .MonitorInit
            .Show False
            .PositionOnScreen mon_pos
        End If
        .MsgText(enMonStep) = mon_text
        .MonitorStep
    End With
    
xt: Exit Sub

eh: If ErrMsg(ErrSrc(PROC)) = vbYes Then: Stop: Resume
End Sub

Public Sub MonitorFooter(ByVal mon_title As String, _
                         ByRef mon_text As udtMsgText, _
                Optional ByVal mon_steps_displayed As Long = 10, _
                Optional ByVal mon_height_max As Long = 80, _
                Optional ByVal mon_pos As String = "5,5", _
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
            AssertWidthAndHeight a_width_min:=mon_width_min _
                               , a_width_max:=mon_width_max _
                               , a_height_max:=mon_height_max
            .MonitorProcess = mon_title
            .MonitorStepsDisplayed = mon_steps_displayed
            .SetupDone = True ' Bypass regular message setup
            .FormHeightOutsideMax = mMsg.ValueAsPt(mon_height_max, enDsplyDimensionHeight)
            .FormWidthOutsideMax = mMsg.ValueAsPt(mon_width_max, enDsplyDimensionWidth)
            .FormWidthOutsideMin = mMsg.ValueAsPt(mon_width_min, enDsplyDimensionWidth)
            .MonitorInit
            .Show False
            .PositionOnScreen mon_pos
        End If
        .MsgText(enMonFooter) = mon_text
        .MonitorFooter
    End With
    
xt: Exit Sub

eh: If ErrMsg(ErrSrc(PROC)) = vbYes Then: Stop: Resume
End Sub

Public Sub MonitorHeader(ByVal mon_title As String, _
                         ByRef mon_text As udtMsgText, _
                Optional ByVal mon_steps_displayed As Long = 10, _
                Optional ByVal mon_height_max As Long = 80, _
                Optional ByVal mon_pos As String = "5,5", _
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
            AssertWidthAndHeight a_width_min:=mon_width_min _
                               , a_width_max:=mon_width_max _
                               , a_height_max:=mon_height_max
            .MonitorProcess = mon_title
            .MonitorStepsDisplayed = mon_steps_displayed
            .SetupDone = True ' Bypass regular message setup
            .FormHeightOutsideMax = mMsg.ValueAsPt(mon_height_max, enDsplyDimensionHeight)
            .FormWidthOutsideMax = mMsg.ValueAsPt(mon_width_max, enDsplyDimensionWidth)
            .FormWidthOutsideMin = mMsg.ValueAsPt(mon_width_min, enDsplyDimensionWidth)
            .MonitorInit
            .Show False
            .PositionOnScreen mon_pos
        End If
        .MsgText(enMonHeader) = mon_text
        .MonitorHeader
    End With
    
xt: Exit Sub

eh: If ErrMsg(ErrSrc(PROC)) = vbYes Then: Stop: Resume
End Sub

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
    Dim MsecsElapsed    As Currency
    Dim MsecsWait       As Long
    
    If MsgInstances Is Nothing Then Set MsgInstances = New Dictionary
    
    If fi_unload Then
        If MsgInstances.Exists(fi_key) Then
            On Error Resume Next
            Unload MsgInstances(fi_key) ' The instance may be already unloaded
            MsgInstances.Remove fi_key
        End If
        Exit Function
    End If
    
    If Not MsgInstances.Exists(fi_key) Then
        '~~ When there is no evidence of an already existing instance a new one is established.
        '~~ In order not to interfere with any prior established instance a minimum wait time
        '~~ of 10 milliseconds is maintained.
        MsecsElapsed = (TicksCount() - cyStart) / CDec(TicksFrequency)
        MsecsWait = 10 - MsecsElapsed
        If MsecsWait > 0 Then Sleep MsecsWait
        Set MsgInstance = Nothing
        Set MsgInstance = New fMsg
        MsgInstances.Add fi_key, MsgInstance
    Else
        '~~ An instance identified by fi_key exists in the Dictionary.
        '~~ It may however have already been unloaded.
        On Error Resume Next
        Set MsgInstance = MsgInstances(fi_key)
        Select Case Err.Number
            Case 0
            Case 13
                If MsgInstances.Exists(fi_key) Then
                    '~~ The apparently no longer existing instance is removed from the Dictionarys
                    MsgInstances.Remove fi_key
                End If
                Set MsgInstance = New fMsg
                MsgInstances.Add fi_key, MsgInstance
            Case Else
                '~~ Unknown error!
                Err.Raise 1 + vbObjectError, ErrSrc(PROC), "Unknown/unrecognized error!"
        End Select
        On Error GoTo -1
    End If

xt: Exit Function

eh: If ErrMsg(ErrSrc(PROC)) = vbYes Then: Stop: Resume
End Function

Public Sub README(Optional ByVal r_bookmark As String = vbNullString)
    
    If r_bookmark = vbNullString Then
        mBasic.ShellRun GITHUB_REPO_URL
    Else
        r_bookmark = Replace("#" & r_bookmark, "##", "#") ' add # if missing
        mBasic.ShellRun GITHUB_REPO_URL & r_bookmark
    End If
        
End Sub

Private Function RoundUp(ByVal v As Variant) As Variant
' -------------------------------------------------------------------------------------
' Returns (v) rounded up to the next integer. Note: to round down omit the "+ 0.5").
' -------------------------------------------------------------------------------------
    RoundUp = Int(v) + (v - Int(v) + 0.5) \ 1
End Function

Public Function Screen(ByVal item As enScreen) As Variant
' -------------------------------------------------------------------------
' Return display screen Item for monitor displaying ActiveWindow
' Patterned after Excel's built-in information functions CELL and INFO
' Supported Item values (each must be a string, but alphabetic case is ignored):
' HorizontalResolution or pixelsX
' VerticalResolution or pixelsY
' WidthInches or inchesX
' HeightInches or inchesY
' DiagonalInches or inchesDiag
' PixelsPerInchX or ppiX
' PixelsPerInchY or ppiY
' PixelsPerInch or ppiDiag
' WinDPIX or dpiX
' WinDPIY or dpiY
' WinDPI or dpiWin ' DPI assumed by Windows
' AdjustmentFactor or zoomFac ' adjustment to match actual size (ppiDiag/dpiWin)
' IsPrimary ' True if primary display
' DisplayName ' name recognized by CreateDC
' Update ' update cells referencing this UDF and return date/time
' Help ' display all recognized Item string values
' EXAMPLE: =Screen("pixelsX")
' Function Returns #VALUE! for invalid Item
' -------------------------------------------------------------------------
    Dim xHSizeSq        As Double
    Dim xVSizeSq        As Double
    Dim xPix            As Double
    Dim xDot            As Double
    Dim hWnd            As LongPtr
    Dim hDC             As LongPtr
    Dim hMonitor        As LongPtr
    Dim tMonitorInfo    As MONITORINFOEX
    Dim nMonitors       As Integer
    Dim vResult         As Variant
    Dim sItem           As String
    
    Application.Volatile
    nMonitors = GetSystemMetrics(SM_CMONITORS)
    If nMonitors < 2 Then
        nMonitors = 1                                       ' in case GetSystemMetrics failed
        hWnd = 0
    Else
        hWnd = GetActiveWindow()
        hMonitor = MonitorFromWindow(hWnd, MONITOR_DEFAULTTONULL)
        If hMonitor = 0 Then
            Debug.Print "ActiveWindow does not intersect a monitor"
            hWnd = 0
        Else
            tMonitorInfo.cbSize = Len(tMonitorInfo)
            If GetMonitorInfo(hMonitor, tMonitorInfo) = False Then
                Debug.Print "GetMonitorInfo failed"
                hWnd = 0
            Else
                hDC = CreateDC(tMonitorInfo.szDevice, 0, 0, 0)
                If hDC = 0 Then
                    Debug.Print "CreateDC failed"
                    hWnd = 0
                End If
            End If
        End If
    End If
    If hWnd = 0 Then
        hDC = GetDC(hWnd)
        tMonitorInfo.dwFlags = MONITOR_PRIMARY
        tMonitorInfo.szDevice = "PRIMARY" & vbNullChar
    End If
    Select Case item
        Case enAdjustmentfactor:    xHSizeSq = GetDeviceCaps(hDC, DevCap.HORZSIZE) ^ 2
                                    xVSizeSq = GetDeviceCaps(hDC, DevCap.VERTSIZE) ^ 2
                                    xPix = GetDeviceCaps(hDC, DevCap.HORZRES) ^ 2 + GetDeviceCaps(hDC, DevCap.VERTRES) ^ 2
                                    xDot = GetDeviceCaps(hDC, DevCap.LOGPIXELSX) ^ 2 * xHSizeSq + GetDeviceCaps(hDC, DevCap.LOGPIXELSY) ^ 2 * xVSizeSq
                                    vResult = 25.4 * Sqr(xPix / xDot)
        Case enWidthDPI:            vResult = GetDeviceCaps(hDC, DevCap.HORZRES)
        Case enHeightDPI:           vResult = GetDeviceCaps(hDC, DevCap.VERTRES)
        Case enWidthInches:         vResult = GetDeviceCaps(hDC, DevCap.HORZSIZE) / 25.4
        Case enHeightInches:        vResult = GetDeviceCaps(hDC, DevCap.VERTSIZE) / 25.4
        Case enDiagonalInches:      vResult = Sqr(GetDeviceCaps(hDC, DevCap.HORZSIZE) ^ 2 + GetDeviceCaps(hDC, DevCap.VERTSIZE) ^ 2) / 25.4
        Case enWidthPPI:      vResult = 25.4 * GetDeviceCaps(hDC, DevCap.HORZRES) / GetDeviceCaps(hDC, DevCap.HORZSIZE)
        Case enHeightPPI:      vResult = 25.4 * GetDeviceCaps(hDC, DevCap.VERTRES) / GetDeviceCaps(hDC, DevCap.VERTSIZE)
        Case enDiagonalPPI:       xHSizeSq = GetDeviceCaps(hDC, DevCap.HORZSIZE) ^ 2
                                    xVSizeSq = GetDeviceCaps(hDC, DevCap.VERTSIZE) ^ 2
                                    xPix = GetDeviceCaps(hDC, DevCap.HORZRES) ^ 2 + GetDeviceCaps(hDC, DevCap.VERTRES) ^ 2
                                    vResult = 25.4 * Sqr(xPix / (xHSizeSq + xVSizeSq))
        Case enWidthWinDPI:     vResult = GetDeviceCaps(hDC, DevCap.LOGPIXELSX)
        Case enHeightWinDPI:     vResult = GetDeviceCaps(hDC, DevCap.LOGPIXELSY)
        Case enWinDPI:      xHSizeSq = GetDeviceCaps(hDC, DevCap.HORZSIZE) ^ 2
                                    xVSizeSq = GetDeviceCaps(hDC, DevCap.VERTSIZE) ^ 2
                                    xDot = GetDeviceCaps(hDC, DevCap.LOGPIXELSX) ^ 2 * xHSizeSq + GetDeviceCaps(hDC, DevCap.LOGPIXELSY) ^ 2 * xVSizeSq
                                    vResult = Sqr(xDot / (xHSizeSq + xVSizeSq))
        Case enIsPrimary:           vResult = CBool(tMonitorInfo.dwFlags And MONITOR_PRIMARY)
        Case enDisplayName:         vResult = tMonitorInfo.szDevice & vbNullChar
                                    vResult = Left(vResult, (InStr(1, vResult, vbNullChar) - 1))
        Case enUpdate:              vResult = Now
        Case enHelp:                vResult = "HorizontalResolution (pixelsX), VerticalResolution (pixelsY), " _
                                            & "WidthInches (inchesX), HeightInches (inchesY), DiagonalInches (inchesDiag), " _
                                            & "PixelsPerInchX (ppiX), PixelsPerInchY (ppiY), PixelsPerInch (ppiDiag), " _
                                            & "WinDPIX (dpiX), WinDPIY (dpiY), WinDPI (dpiWin), " _
                                            & "AdjustmentFactor (zoomFac), IsPrimary, DisplayName, Update, Help"
        Case Else:                  vResult = CVErr(xlErrValue)                         ' return #VALUE! error (2015)
    End Select
    
    If hWnd = 0 _
    Then ReleaseDC hWnd, hDC _
    Else DeleteDC hDC
    Screen = vResult
    
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

Private Function StackIsEmpty(ByVal stck As Collection) As Boolean
' ----------------------------------------------------------------------------
' Returns TRUE when the stack (stck) is empty.
' ----------------------------------------------------------------------------
    If stck Is Nothing _
    Then StackIsEmpty = True _
    Else StackIsEmpty = stck.Count = 0
End Function

Private Function StackPop(ByVal stck As Collection) As Variant
' ----------------------------------------------------------------------------
' Common Stack Pop service. Returns the last item pushed on the stack (stck)
' and removes the item from the stack. When the stack (stck) is empty a
' vbNullString is returned.
' ----------------------------------------------------------------------------
    Const PROC = "StckPop"
    
    On Error GoTo eh
    If StackIsEmpty(stck) Then GoTo xt
    
    If IsObject(stck(stck.Count)) _
    Then Set StackPop = stck(stck.Count) _
    Else StackPop = stck(stck.Count)
    stck.Remove stck.Count

xt: Exit Function

eh: If ErrMsg(ErrSrc(PROC)) = vbYes Then: Stop: Resume
End Function

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

Private Function TicksCount() As Currency:      getTickCount TicksCount:        End Function

Private Function TicksFrequency() As Currency:  getFrequency TicksFrequency:    End Function

Public Function ValueAsPercentage(ByVal p_value As Long, _
                                  ByVal p_dimension As enDsplyDimension) As Single
' ------------------------------------------------------------------------------
' Returns a value (p_value) as percentage considering a screen width or height
' dimensions (p_dimension), whereby a value <= 100 is considered a percentage
' already, a value > 100 is regarded a pt value and transformed to a percentage.
' ------------------------------------------------------------------------------
    If p_value > 100 Then
        Select Case p_dimension
            Case enDsplyDimensionWidth:     ValueAsPercentage = Int(p_value / (DsplyWidthPT / 100))
            Case enDsplyDimensionHeight:    ValueAsPercentage = Int(p_value / (DsplyHeightPT / 100))
        End Select
    Else
        ValueAsPercentage = p_value
    End If
End Function

Public Function ValueAsPt(ByVal p_value As Long, _
                          ByVal p_dimension As enDsplyDimension) As Single
' ------------------------------------------------------------------------------
' Returns a value (p_value) as pt considering a dimensions width or height
' (p_dimension), whereby a value <= 100 is considered a percentage and therefore
' is computed into pt. A p_value > 100 is regarded a pt value already.
' ------------------------------------------------------------------------------
    If p_value <= 100 Then
        Select Case p_dimension
            Case enDsplyDimensionWidth:    ValueAsPt = RoundUp(DsplyWidthPT * (p_value / 100))
            Case enDsplyDimensionHeight:   ValueAsPt = RoundUp(DsplyHeightPT * (p_value / 100))
        End Select
    Else
        ValueAsPt = p_value
    End If
End Function

