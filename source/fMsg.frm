VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} fMsg 
   ClientHeight    =   11598
   ClientLeft      =   150
   ClientTop       =   390
   ClientWidth     =   12390
   OleObjectBlob   =   "fMsg.frx":0000
End
Attribute VB_Name = "fMsg"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Option Explicit
' -------------------------------------------------------------------------------
' UserForm fMsg Provides all means for a message with up to 5 separated text
'               sections, either proportional- or mono-spaced, with an optional
'               label, and up to 7 reply buttons.
'
' Design:       Since the implementation is merely design driven its setup is
'               essential. Design changes must adhere to the concept.
'
' Public Properties:
' - IndicateFrameCaptions Test option, indicated the frame names
' - MinButtonWidth
' - MsgTitle               The title displayed in the window handle bar
' - MinButtonWidth         Minimum button width in pt
' - MsgButtonDefault       The number of the default button
' - MsgBttns               Buttons to be displayed, Collection provided by the
'                          mMsg.Buttons service
' - MsgHeightMax           Percentage of screen height
' - MsgHeightMin           Percentage of screen height
' - MsgLabel               A section's label
' - MsgWidthMax            Percentage of screen width
' - MsgWidthMin            Defaults to 400 pt. the absolute minimum is 200 pt
' - Text                   A section's text or a monitor header, monitor footer
'                          or monitor step text
' - VisualizeForTest       Test option, visualizes the controls via a specific
'                          BackColor
'
' Uses:         Module mMsg to pass on the clicked reply button to the caller.
'               Note: The UserForm cannot be used directly unless the implemen-
'               tation is mimicked.
'
' Requires:     Reference to "Microsoft Scripting Runtime"
'
' See details at:
' https://warbe-maker.github.io/warbe-maker.github.io/vba/common/2020/11/17/Common-VBA-Message-Services.html
'
' W. Rauschenberger Berlin, Mar 2022 (last revision)
' --------------------------------------------------------------------------
Const DFLT_BTTN_MIN_WIDTH           As Single = 70              ' Default minimum reply button width
Const DFLT_LBL_MONOSPACED_FONT_NAME As String = "Courier New"   ' Default monospaced font name
Const DFLT_LBL_MONOSPACED_FONT_SIZE As Single = 9               ' Default monospaced font size
Const DFLT_LBL_PROPSPACED_FONT_NAME As String = "Calibri"       ' Default proportional spaced font name
Const DFLT_LBL_PROPSPACED_FONT_SIZE As Single = 9               ' Default proportional spaced font size
Const DFLT_TXT_MONOSPACED_FONT_NAME As String = "Courier New"   ' Default monospaced font name
Const DFLT_TXT_MONOSPACED_FONT_SIZE As Single = 10              ' Default monospaced font size
Const DFLT_TXT_PROPSPACED_FONT_NAME As String = "Tahoma"        ' Default proportional spaced font name
Const DFLT_TXT_PROPSPACED_FONT_SIZE As Single = 10              ' Default proportional spaced font size
Const HSPACE_BTTN_AREA              As Single = 15              ' Minimum left and right margin for the centered buttons area
Const HSPACE_BTTNS                As Single = 4               ' Horizontal space between reply buttons
Const HSPACE_LEFT                   As Single = 0               ' Left margin for labels and text boxes
Const HSPACE_RIGHT                  As Single = 15              ' Horizontal right space for labels and text boxes
Const HSPACE_LEFTRIGHT_BUTTONS      As Long = 8                 ' The margin before the left most and after the right most button
Const MARGIN_RIGHT_MSG_AREA         As String = 7
Const NEXT_ROW                      As String = vbLf            ' Reply button row break
Const SCROLL_V_WIDTH                As Single = 20              ' Additional horizontal space required for a frame with a vertical scrollbar
Const SCROLL_H_HEIGHT               As Single = 20              ' Additional vertical space required for a frame with a horizontal scroll barr
Const TEST_WITH_FRAME_BORDERS       As Boolean = False          ' For test purpose only! Display frames with visible border
Const TEST_WITH_FRAME_CAPTIONS      As Boolean = False          ' For test purpose only! Display frames with their test captions (erased by default)
Const VSPACE_AREAS                  As Single = 12              ' Vertical space between message area and replies area
Const VSPACE_BOTTOM                 As Single = 30              ' Space occupied by the title bar
Const VSPACE_BTTN_ROWS              As Single = 5               ' Vertical space between button rows
Const VSPACE_LABEL                  As Single = 0               ' Vertical space between the section-label and the following section-text
Const VSPACE_SECTIONS               As Single = 7               ' Vertical space between displayed message sections
Const VSPACE_TEXTBOXES              As Single = 18              ' Vertical bottom marging for all textboxes
Const VSPACE_TOP                    As Single = 2               ' Top position for the first displayed control
Const VISLZE_BCKCLR_AREA            As Long = &HC0E0FF          ' -------------
Const VISLZE_BCKCLR_BTTNS_FRM       As Long = &HFFFFC0          ' Backcolors
Const VISLZE_BCKCLR_BTTNS_ROW_FRM   As Long = &HC0FFC0          ' for the
Const VISLZE_BCKCLR_MON_STEPS_FRM   As Long = &HFFFFC0          ' visualization
Const VISLZE_BCKCLR_MSEC_FRM        As Long = &HFFFFC0          ' of the
Const VISLZE_BCKCLR_MSEC_LBL        As Long = &HC0FFFF          ' controls
Const VISLZE_BCKCLR_MSEC_TBX        As Long = &H80C0FF          ' during test
Const VISLZE_BCKCLR_MSEC_TBX_FRM    As Long = &HC0FFC0          ' -------------

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
' -------------------------------------------------------------------------------

'Api Declarations
Private Declare PtrSafe Function GetCursorInfo Lib "user32" (ByRef pci As CursorInfo) As Boolean
Private Declare PtrSafe Function LoadCursor Lib "user32" Alias "LoadCursorA" (ByVal hInstance As Long, ByVal lpCursorName As Long) As Long
Private Declare PtrSafe Function SetCursor Lib "user32" (ByVal hCursor As Long) As Long
Private Declare PtrSafe Sub Sleep Lib "kernel32" (ByVal dwMilliseconds As Long)

'You can use the default cursors in windows
Public Enum CursorTypes
    IDC_ARROW = 32512
    IDC_IBEAM = 32513
    IDC_WAIT = 32514
    IDC_CROSS = 32515
    IDC_UPARROW = 32516
    IDC_SIZE = 32640
    IDC_ICON = 32641
    IDC_SIZENWSE = 32642
    IDC_SIZENESW = 32643
    IDC_SIZEWE = 32644
    IDC_SIZENS = 32645
    IDC_SIZEALL = 32646
    IDC_NO = 32648
    IDC_HAND = 32649
    IDC_APPSTARTING = 32650
End Enum

'Needed for GetCursorInfo
Private Type POINT
    X As Long
    Y As Long
End Type

'Needed for GetCursorInfo
Private Type CursorInfo
    cbSize As Long
    flags As Long
    hCursor As Long
    ptScreenPos As POINT
End Type

' Timer means
Private Declare PtrSafe Function getFrequency Lib "kernel32" _
Alias "QueryPerformanceFrequency" (TimerSystemFrequency As Currency) As Long
Private Declare PtrSafe Function getTickCount Lib "kernel32" _
Alias "QueryPerformanceCounter" (cyTickCount As Currency) As Long
' -------------------------------------------------------------------------------

Private Enum enMsgFormUsage
    usage_progress_display = 1
'    usage_message_display = 2
End Enum

Public Enum MSFormControls
    ' List of all the MSForms Controls.
    CheckBox
    ComboBox
    CommandButton
    Frame
    Image
    Label
    ListBox
    MultiPage
    OptionButton
    ScrollBar
    SpinButton
    TabStrip
    TextBox
    ToggleButton
End Enum

Private AppliedBttns            As Dictionary   ' Dictionary of applied buttons (key=CommandButton, item=row)
Private AppliedBttnsRetVal      As Dictionary   ' Dictionary of the applied buttons' reply value (key=CommandButton)
Private bDoneMonoSpacedSects    As Boolean
Private bDoneMsgArea            As Boolean
Private bDonePropSpacedSects    As Boolean
Private bDoneTitle              As Boolean
Private bFormEvents             As Boolean
Private bIndicateFrameCaptions  As Boolean
Private bMonitorInitialized     As Boolean
Private bModeLess               As Boolean
Private bReplyWithIndex         As Boolean
Private bSetUpDone              As Boolean
Private bVisualizeForTest       As Boolean
Private cllDsgnBttnRows         As Collection       ' Collection of the designed reply button row frames
Private cllDsgnRowBttns         As Collection       ' Collection of a designed reply button row's buttons
Private cllMsgBttns             As New Collection
Private cllSteps                As Collection
Private cyTimerTicksBegin       As Currency
Private cyTimerTicksEnd         As Currency
Private dctApplicationRunArgs   As New Dictionary   ' Dictionary will be available with each instance of this UserForm
Private dctAreas                As New Dictionary   ' Collection of the two primary/top frames
Private dctBttns                As New Dictionary   ' Collection of the collection of the designed reply buttons of a certain row
Private dctBttnsRow             As New Dictionary   ' Established/created Button Row's Frame
Private dctBttnsRowBttns        As New Dictionary   ' Established/created Command Buttons in a specific BttnsRow
Private dctMonoSpaced           As New Dictionary
Private dctMonoSpacedTbx        As New Dictionary
Private dctMsectsFrm            As New Dictionary   ' Established/created Message Sections Frame
Private dctMsectsLbl            As New Dictionary   ' Established/created Message Sections Label
Private dctMsectsTbx            As New Dictionary   ' Established/created Message Sections TextBox
Private dctMsectsTbxFrm         As New Dictionary   ' Established/created Message Sections TextBox Frame
Private dctSectsLabel           As New Dictionary   ' MsectFrm specific label either provided via properties MsgLabel or Msg
Private dctSectsMonoSpaced      As New Dictionary   ' MsectFrm specific monospace option either provided via properties MsgMonospaced or Msg
Private dctSectsText            As New Dictionary
Private frmBttnsArea            As MSForms.Frame
Private frmBttnsFrm             As MSForms.Frame    ' Set with CollectDesignControls
Private frmMsectsArea           As MSForms.Frame
Private frmSteps                As MSForms.Frame
Private lBackColor              As Long
Private sMonitorProcess         As String
Private lNoOfDesignedMsgSects   As Long             ' Set with CollectDesignControls (number of message sections designed)
Private lMonitorStepsDisplayed  As Long
Private lSetupRowButtons        As Long             ' number of buttons setup in a row
Private lSetupRows              As Long             ' number of setup button rows
Private lStepsDisplayed         As Long
Private MsectFrm                As MSForms.Frame    ' Set with ProvideSectionItems for a certain section
Private MsectLbl                As MSForms.Label    ' Set with ProvideSectionItems for a certain section
Private MsectTbx                As MSForms.TextBox  ' Set with ProvideSectionItems for a certain section
Private MsectTbxFrm             As MSForms.Frame    ' Set with ProvideSectionItems for a certain section
Private MsgSectLbl              As TypeMsgLabel     ' Label section of the TypeMsg UDT
Private MsgSectTxt              As TypeMsgText      ' Text section of the TypeMsg UDT
Private siHmarginFrames         As Single           ' Test property, value defaults to 0
Private siMaxButtonHeight       As Single
Private siMaxButtonWidth        As Single
Private siMsgHeightMax          As Single       ' Maximum message height in pt
Private siMsgHeightMin          As Single       ' Minimum message height in pt
Private siMsgWidthMax           As Single       ' Maximum message width in pt
Private siMsgWidthMin           As Single       ' Minimum message width in pt
Private siVmarginFrames         As Single       ' Test property, value defaults to 0
Private sMsgTitle               As String
Private tbxFooter               As MSForms.TextBox
Private tbxHeader               As MSForms.TextBox
Private tbxStep                 As MSForms.TextBox
Private TextMonitorFooter       As TypeMsgText
Private TextMonitorHeader       As TypeMsgText
Private TextMonitorStep         As TypeMsgText
Private TextSection             As TypeMsg
Private TimerSystemFrequency    As Currency
Private TitleWidth              As Single
Private VirtualScreenHeightPts  As Single
Private VirtualScreenLeftPts    As Single
Private VirtualScreenTopPts     As Single
Private VirtualScreenWidthPts   As Single
Private vMsgButtonDefault       As Variant          ' Index or caption of the default button
'Private vReplyValue             As Variant
Private iSectionsPropSpaced     As Long             ' number of prop. spaced sections setup
Private iSectionsMonoSpaced     As Long             ' number of mono-spaced sections setup

Public Property Let ModeLess(ByVal b As Boolean): bModeLess = b:    End Property
Private Sub UserForm_Initialize()
    Const PROC = "UserForm_Initialize"
    
    On Error GoTo eh
    ' Get the display screen's dimensions and position in pts
    GetScreenMetrics VirtualScreenLeftPts _
                   , VirtualScreenTopPts _
                   , VirtualScreenWidthPts _
                   , VirtualScreenHeightPts
    Initialize
    CollectDesignControls

xt: Exit Sub
    
eh: If ErrMsg(ErrSrc(PROC)) = vbYes Then: Stop: Resume
End Sub

Private Sub Initialize()
    Const PROC = "Initialize"
    
    On Error GoTo eh
    
    If bSetUpDone Then GoTo xt
    Set dctMonoSpaced = New Dictionary
    Set dctMonoSpacedTbx = New Dictionary
    
    bFormEvents = False
    siMsgWidthMin = 20      ' Default minimum message window width
    siHmarginFrames = 0     ' Ensures proper command buttons framing, may be used for test purpose
    Me.VmarginFrames = 0    ' Ensures proper command buttons framing and vertical positioning of controls
    SetupDone = False
    bDoneTitle = False
    bDoneMonoSpacedSects = False
    bDonePropSpacedSects = False
    bDoneMsgArea = False
    vMsgButtonDefault = 1

xt: Exit Sub
    
eh: If ErrMsg(ErrSrc(PROC)) = vbYes Then: Stop: Resume
End Sub

Private Sub UserForm_Terminate()
    Set AppliedBttns = Nothing
    Set AppliedBttnsRetVal = Nothing
    Set dctAreas = Nothing
    Set cllDsgnBttnRows = Nothing
    Set dctBttns = Nothing
    Set dctMsectsFrm = Nothing
    Set dctMsectsLbl = Nothing
    Set dctMsectsTbx = Nothing
    Set dctMsectsTbxFrm = Nothing
    Set cllDsgnRowBttns = Nothing
    Set dctMonoSpaced = Nothing
    Set dctMonoSpacedTbx = Nothing
    Set dctSectsLabel = Nothing
    Set dctSectsMonoSpaced = Nothing
    Set dctSectsText = Nothing
End Sub

Public Property Get MonitorIsInitialized() As Boolean: MonitorIsInitialized = Not cllSteps Is Nothing:  End Property

Private Property Get AppliedButtonRetVal(Optional ByVal Button As MSForms.CommandButton) As Variant
    AppliedButtonRetVal = AppliedBttnsRetVal(Button)
End Property

Private Property Let AppliedButtonRetVal(Optional ByVal Button As MSForms.CommandButton, ByVal v As Variant)
    AppliedBttnsRetVal.Add Button, v
End Property

Private Property Get AppliedButtonRowHeight() As Single
    AppliedButtonRowHeight = siMaxButtonHeight + 2
End Property

Private Property Get ClickedButtonIndex(Optional ByVal cmb As MSForms.CommandButton) As Long
    
    Dim i   As Long
    Dim v   As Variant
    
    For Each v In AppliedBttnsRetVal
        i = i + 1
        If v Is cmb Then
            ClickedButtonIndex = i
            Exit For
        End If
    Next v

End Property

Private Property Let FormWidth(ByVal considered_width As Single)
' ------------------------------------------------------------------------------
' The FormWidth property ensures
' - it is not less than the minimum specified width
' - it does not exceed the specified or the default maximum value
' - it may expand up to the maximum but never shrink
' ------------------------------------------------------------------------------
    Dim new_width As Single
    new_width = Max(Me.Width, TitleWidth, siMsgWidthMin, considered_width + 15)
    Me.Width = Min(new_width, siMsgWidthMax + Max(ScrollVscrollWidth(frmMsectsArea), ScrollVscrollWidth(frmBttnsArea)))
End Property

Private Property Get FormWidthMaxUsable()
    FormWidthMaxUsable = siMsgWidthMax - 15
End Property

Public Property Let IndicateFrameCaptions(ByVal b As Boolean):                              bIndicateFrameCaptions = b:                                         End Property

Private Property Get MaxRowsHeight() As Single:                                             MaxRowsHeight = siMaxButtonHeight + (siVmarginFrames * 2):          End Property

Private Property Get MaxWidthMsgArea() As Single
' ------------------------------------------------------------------------------
' The maximum usable message area width considers the specified maximum form
' width and the InsideWidth
' ------------------------------------------------------------------------------
    MaxWidthMsgArea = Me.InsideWidth
End Property

Private Property Get MonitorHeightExSteps() As Single
    MonitorHeightExSteps = ContentHeight(frmSteps.Parent) - frmSteps.Height
End Property

Private Property Get MonitorHeightMaxSteps()
    MonitorHeightMaxSteps = Me.MsgHeightMax - MonitorHeightExSteps
End Property

Private Property Get MonoSpaced(Optional ByVal var_ctl As Variant) As Boolean
    MonoSpaced = dctMonoSpaced.Exists(var_ctl)
End Property

Private Property Let MonoSpaced( _
                 Optional ByVal var_ctl As Variant, _
                          ByVal b As Boolean)
    If b Then
        If Not dctMonoSpaced.Exists(var_ctl) Then dctMonoSpaced.Add var_ctl, var_ctl.Name
    Else
        If dctMonoSpaced.Exists(var_ctl) Then dctMonoSpaced.Remove var_ctl
    End If
End Property

Private Property Get MonoSpacedTbx(Optional ByVal tbx As MSForms.TextBox) As Boolean
    MonoSpacedTbx = dctMonoSpacedTbx.Exists(tbx)
End Property

Private Property Let MonoSpacedTbx( _
                 Optional ByVal tbx As MSForms.TextBox, _
                          ByVal b As Boolean)
    If b Then
        If Not dctMonoSpacedTbx.Exists(tbx) Then dctMonoSpacedTbx.Add tbx, tbx.Name
    Else
        If dctMonoSpacedTbx.Exists(tbx) Then dctMonoSpacedTbx.Remove tbx
    End If
End Property

Private Property Get MSFormsCtlType(ByVal msf_enum As MSFormControls) As String
' ------------------------------------------------------------------------------
' Returns the control Type of the provided ProgID. See
' https://docs.microsoft.com/en-us/office/vba/language/reference/user-interface-help/add-method-microsoft-forms
' ------------------------------------------------------------------------------
    Select Case msf_enum
        Case CheckBox:      MSFormsCtlType = "CheckBox"
        Case ComboBox:      MSFormsCtlType = "ComboBox"
        Case CommandButton: MSFormsCtlType = "CommandButton"
        Case Frame:         MSFormsCtlType = "Frame"
        Case Image:         MSFormsCtlType = "Image"
        Case Label:         MSFormsCtlType = "Label"
        Case ListBox:       MSFormsCtlType = "ListBox"
        Case MultiPage:     MSFormsCtlType = "MultiPage"
        Case OptionButton:  MSFormsCtlType = "OptionButton"
        Case ScrollBar:     MSFormsCtlType = "ScrollBar"
        Case SpinButton:    MSFormsCtlType = "SpinButton"
        Case TabStrip:      MSFormsCtlType = "TabStrip"
        Case TextBox:       MSFormsCtlType = "TextBox"
        Case ToggleButton:  MSFormsCtlType = "ToggleButton"
    End Select
End Property

Private Property Get MSFormsProgID(Optional mfc As MSFormControls) As String
' ------------------------------------------------------------------------------
' Returns the ProgID for the control (mfc). See
' https://docs.microsoft.com/en-us/office/vba/language/reference/user-interface-help/add-method-microsoft-forms
' ------------------------------------------------------------------------------
    Select Case mfc
      Case MSFormControls.CheckBox:       MSFormsProgID = "Forms.CheckBox.1"
      Case MSFormControls.ComboBox:       MSFormsProgID = "Forms.ComboBox.1"
      Case MSFormControls.CommandButton:  MSFormsProgID = "Forms.CommandButton.1"
      Case MSFormControls.Frame:          MSFormsProgID = "Forms.Frame.1"
      Case MSFormControls.Image:          MSFormsProgID = "Forms.Image.1"
      Case MSFormControls.Label:          MSFormsProgID = "Forms.Label.1"
      Case MSFormControls.ListBox:        MSFormsProgID = "Forms.ListBox.1"
      Case MSFormControls.MultiPage:      MSFormsProgID = "Forms.MultiPage.1"
      Case MSFormControls.OptionButton:   MSFormsProgID = "Forms.OptionButton.1"
      Case MSFormControls.ScrollBar:      MSFormsProgID = "Forms.ScrollBar.1"
      Case MSFormControls.SpinButton:     MSFormsProgID = "Forms.SpinButton.1"
      Case MSFormControls.TabStrip:       MSFormsProgID = "Forms.TabStrip.1"
      Case MSFormControls.TextBox:        MSFormsProgID = "Forms.TextBox.1"
      Case MSFormControls.ToggleButton:   MSFormsProgID = "Forms.ToggleButton.1"
    End Select
End Property

Public Property Let MsgBttns(ByVal cll As Collection):  Set cllMsgBttns = cll:          End Property

Public Property Let MsgButtonDefault(ByVal v As Variant)
    vMsgButtonDefault = v
End Property

Public Property Get MsgHeightMax() As Single:           MsgHeightMax = siMsgHeightMax:  End Property

Public Property Let MsgHeightMax(ByVal si As Single):   siMsgHeightMax = si:            End Property

Public Property Get MsgHeightMin() As Single:           MsgHeightMin = siMsgHeightMin:  End Property

Public Property Let MsgHeightMin(ByVal si As Single):   siMsgHeightMin = si:            End Property

Public Property Get MsgLabel(Optional ByVal lbl_section As Long = 1) As TypeMsgLabel
' ------------------------------------------------------------------------------
' Returns the text for the Label of section (lbl_sectionn).
' ------------------------------------------------------------------------------
    Dim vArry() As Variant
    
    If dctSectsLabel Is Nothing Then
        MsgLabel.Text = vbNullString
    ElseIf Not dctSectsLabel.Exists(lbl_section) Then
        MsgLabel.Text = vbNullString
    Else
        vArry = dctSectsLabel(lbl_section)
        MsgLabel.FontBold = vArry(0)
        MsgLabel.FontColor = vArry(1)
        MsgLabel.FontItalic = vArry(2)
        MsgLabel.FontName = vArry(3)
        MsgLabel.FontSize = vArry(4)
        MsgLabel.FontUnderline = vArry(5)
        MsgLabel.MonoSpaced = vArry(6)
        MsgLabel.Text = vArry(7)
        MsgLabel.OpenWhenClicked = vArry(8)
    End If

End Property

Public Property Let MsgLabel(Optional ByVal lbl_section As Long = 1, _
                                      ByRef lbl_text As TypeMsgLabel)
' ------------------------------------------------------------------------------
' Provide the text (lbl_text) as section (lbl_section) text, section label,
' monitor header, footer, or step (lbl_kind_of_text).
' ------------------------------------------------------------------------------
    Dim vArry(0 To 8)   As Variant
    
    vArry(0) = lbl_text.FontBold
    vArry(1) = lbl_text.FontColor
    vArry(2) = lbl_text.FontItalic
    vArry(3) = lbl_text.FontName
    vArry(4) = lbl_text.FontSize
    vArry(5) = lbl_text.FontUnderline
    vArry(6) = lbl_text.MonoSpaced
    vArry(7) = lbl_text.Text
    vArry(8) = lbl_text.OpenWhenClicked
    dctSectsLabel.Add lbl_section, vArry

End Property

Public Property Get MsgTitle() As String:               MsgTitle = Me.Caption:                                          End Property

Public Property Let MsgTitle(ByVal s As String):        sMsgTitle = s:                                                  End Property

Public Property Get MsgWidthMax() As Single:            MsgWidthMax = siMsgWidthMax:                                    End Property

Public Property Let MsgWidthMax(ByVal si As Single):    siMsgWidthMax = si:                                             End Property

Public Property Get MsgWidthMin() As Single:            MsgWidthMin = siMsgWidthMin:                                    End Property

Public Property Let MsgWidthMin(ByVal si As Single):    siMsgWidthMin = si:                                             End Property

Public Property Let NewHeight(Optional ByRef nh_frame_form As Object, _
                              Optional ByVal nh_for_visible_only As Boolean = True, _
                              Optional ByVal nh_y_action As fmScrollAction = fmScrollActionBegin, _
                              Optional ByVal nh_threshold_height_diff As Single = 5, _
                                       ByVal nh_height As Single)
' ------------------------------------------------------------------------------
' Mimics a height change event. Applies a vertical scroll-bar when the content
' height of the frame or form (nh_frame_form) is greater than the height of
' the frame or form by considering a threshold (nh_threshold_height_diff) in
' order to avoid a usesless scroll-bar for a redicolous height difference. In
' case the new height is less the the frame's height a vertical scrollbar is
' removed.
' ------------------------------------------------------------------------------
    Const PROC = "NewHeight"
    
    On Error GoTo eh
    Dim siContentHeight As Single:  siContentHeight = ContentHeight(nh_frame_form, nh_for_visible_only)
    
    If nh_frame_form Is Nothing Then Err.Raise AppErr(1), ErrSrc(PROC), "The required argument 'nh_frame_form' is Nothing!"
    If Not IsFrameOrForm(nh_frame_form) _
    Then Err.Raise AppErr(1), ErrSrc(PROC), "The provided argument is neither a Frame nor a Form!"
        
'    Debug.Print "nh_frame_form.Name   = " & nh_frame_form.Name
'    Debug.Print "nh_frame_form.Height = " & nh_frame_form.Height
'    Debug.Print "siContentHeight      = " & siContentHeight
    nh_frame_form.Height = nh_height
'    Debug.Print "New frame height     = " & nh_frame_form.Height
'    Debug.Print "V-Scroll             = " & ScrollVscrollApplied(nh_frame_form)
    
    If siContentHeight - nh_frame_form.Height > nh_threshold_height_diff Then
        Debug.Print "apply a vertical scroll-bar for " & nh_frame_form.Name
        ScrollVscrollApply sva_frame_form:=nh_frame_form, sva_content_height:=siContentHeight, sva_y_action:=nh_y_action
    ElseIf ScrollVscrollApplied(nh_frame_form) Then
        ScrollVscrollRemove nh_frame_form
    End If
    
xt: Exit Property
    
eh: If ErrMsg(ErrSrc(PROC)) = vbYes Then: Stop: Resume
End Property

Private Property Let NewWidth(Optional ByRef nw_frame_form As Object, _
                              Optional ByVal nw_for_visible_only As Boolean = True, _
                              Optional ByVal nw_threshold_width_diff As Single = 5, _
                                       ByVal nw_width As Single)
' ------------------------------------------------------------------------------
' Mimics a width change event. Applies a horizontal scroll-bar when the content
' height of the frame or form (nw_frame_form) is greater than the width of
' the frame or form by considering a threshold (nw_threshold_width_diff) in
' order to avoid a usesless scroll-bar for a redicolous width difference. In
' case the new width is less the frame's width a horizontal scrollbar is removed.
' ------------------------------------------------------------------------------
    Const PROC = "NewWidth"
    
    On Error GoTo eh
    Dim siContentWidth  As Single:  siContentWidth = ContentWidth(nw_frame_form, nw_for_visible_only)
    
    If nw_frame_form Is Nothing Then Err.Raise AppErr(1), ErrSrc(PROC), "The required argument 'nw_frame_form' is Nothing!"
    If Not IsFrameOrForm(nw_frame_form) Then Err.Raise AppErr(2), ErrSrc(PROC), "The provided argument 'nw_frame_form' is neither a Frame nor a Form!"
    
'    Debug.Print "nw_frame_form.Name  = " & nw_frame_form.Name
'    Debug.Print "nw_frame_form.Width = " & nw_frame_form.Width
'    Debug.Print "siContentWidth      = " & siContentWidth
    nw_frame_form.Width = nw_width
'    Debug.Print "New frame width     = " & nw_frame_form.Width
'    Debug.Print "H-Scroll            = " & ScrollHscrollApplied(nw_frame_form)

    If siContentWidth - nw_frame_form.Width > nw_threshold_width_diff Then
        ScrollHscrollApply sha_frame_form:=nw_frame_form, sha_content_width:=siContentWidth
    ElseIf ScrollHscrollApplied(nw_frame_form) Then
        ScrollHscrollRemove nw_frame_form
    End If
    
xt: Exit Property
    
eh: If ErrMsg(ErrSrc(PROC)) = vbYes Then: Stop: Resume
End Property

Public Property Get NoOfDesignedMsgSects() As Long ' -----------------------
    NoOfDesignedMsgSects = lNoOfDesignedMsgSects   ' Global definition !!!!!
End Property                                       ' -----------------------

Private Property Let NoOfDesignedMsgSects(ByVal l As Long)
    lNoOfDesignedMsgSects = l
End Property

Private Property Get PrcntgHeightfrmBttnsArea() As Single
    PrcntgHeightfrmBttnsArea = Round(frmBttnsArea.Height / (frmMsectsArea.Height + frmBttnsArea.Height), 2)
End Property

Private Property Get PrcntgHeightMsgArea() As Single
    PrcntgHeightMsgArea = Round(frmMsectsArea.Height / (frmMsectsArea.Height + frmBttnsArea.Height), 2)
End Property

'Public Property Get ReplyValue() As Variant:                ReplyValue = vReplyValue:                                   End Property

Public Property Let ReplyWithIndex(ByVal b As Boolean):     bReplyWithIndex = b:                                        End Property

Public Property Let SetupDone(ByVal b As Boolean):          bSetUpDone = b:         End Property

Private Property Get SysFrequency() As Currency
    If TimerSystemFrequency = 0 Then getFrequency TimerSystemFrequency
    SysFrequency = TimerSystemFrequency
End Property

Public Property Get Text(Optional ByVal txt_kind_of_text As KindOfText, _
                         Optional ByVal txt_section As Long = 1) As TypeMsgText
' ------------------------------------------------------------------------------
' Returns the text (txt_kind_of_text) as section-text or -label, monitor-header,
' -footer, or -step.
' ------------------------------------------------------------------------------
    Dim vArry() As Variant
    
    Select Case txt_kind_of_text
        Case enMonHeader:    Text = TextMonitorHeader
        Case enMonFooter:    Text = TextMonitorFooter
        Case enMonStep:      Text = TextMonitorStep
        Case enSectText
            If dctSectsText Is Nothing Then
                Text.Text = vbNullString
            ElseIf Not dctSectsText.Exists(txt_section) Then
                Text.Text = vbNullString
            Else
                vArry = dctSectsText(txt_section)
                Text.FontBold = vArry(0)
                Text.FontColor = vArry(1)
                Text.FontItalic = vArry(2)
                Text.FontName = vArry(3)
                Text.FontSize = vArry(4)
                Text.FontUnderline = vArry(5)
                Text.MonoSpaced = vArry(6)
                Text.Text = vArry(7)
            End If
    End Select
End Property

Public Property Let Text(Optional ByVal txt_kind_of_text As KindOfText, _
                         Optional ByVal txt_section As Long = 1, _
                                  ByRef txt_text As TypeMsgText)
' ------------------------------------------------------------------------------
' Provide the text (txt_text) as section (txt_section) text, section label,
' monitor header, footer, or step (txt_kind_of_text).
' ------------------------------------------------------------------------------
    Dim vArry(0 To 7)   As Variant
    
    vArry(0) = txt_text.FontBold
    vArry(1) = txt_text.FontColor
    vArry(2) = txt_text.FontItalic
    vArry(3) = txt_text.FontName
    vArry(4) = txt_text.FontSize
    vArry(5) = txt_text.FontUnderline
    vArry(6) = txt_text.MonoSpaced
    vArry(7) = txt_text.Text
    Select Case txt_kind_of_text
        Case enMonHeader:    TextMonitorHeader = txt_text
        Case enMonFooter:    TextMonitorFooter = txt_text
        Case enMonStep:      TextMonitorStep = txt_text
        Case enSectText:     dctSectsText.Add txt_section, vArry
    End Select

End Property

Private Property Get TimerSecsElapsed() As Currency:        TimerSecsElapsed = TimerTicksElapsed / SysFrequency:        End Property

Private Property Get TimerSysCurrentTicks() As Currency:    getTickCount TimerSysCurrentTicks:                          End Property

Private Property Get TimerTicksElapsed() As Currency:       TimerTicksElapsed = cyTimerTicksEnd - cyTimerTicksBegin:    End Property

Public Property Get VisualizeForTest() As Boolean:          VisualizeForTest = bVisualizeForTest:                       End Property

Public Property Let VisualizeForTest(ByVal b As Boolean)
    bVisualizeForTest = b
    CollectDesignControls ' do again to ensure visualization
End Property

Public Property Get VmarginFrames() As Single:              VmarginFrames = siVmarginFrames:                            End Property

Public Property Let VmarginFrames(ByVal si As Single):      siVmarginFrames = AdjustToVgrid(si):                        End Property

Public Function AddControl(ByVal ac_ctl As MSFormControls _
                , Optional ByVal ac_in As MSForms.Frame = Nothing _
                , Optional ByVal ac_name As String = vbNullString _
                , Optional ByVal ac_visible As Boolean = False) As MSForms.Control
' ------------------------------------------------------------------------------
' Returns the type of control (ac_ctl) added to the to the userform or - when
' provided - to the frame (ac_in), optionally named (ac_name) and by default
' invisible (ac_visible).
' ------------------------------------------------------------------------------
    Const PROC = "AddControl"
    
    On Error GoTo eh
    Dim ctl As MSForms.Control
    Dim frm As MSForms.Frame
    
    If ac_in Is Nothing Then
        If Not CtlExists(ac_name) Then
            Set ctl = Me.Controls.Add(bstrProgID:=MSFormsProgID(ac_ctl) _
                                    , Name:=ac_name _
                                    , Visible:=ac_visible)
            Set AddControl = ctl
        End If
    Else
        If Not IsFrameOrForm(ac_in) _
        Then Err.Raise AppErr(1), ErrSrc(PROC), "The object in argument 'ac_in' is neither a Frame nor a UserForm!"
        
        If Not CtlExists(ac_name) Then
            If ac_ctl = Frame Then Stop
            If ac_ctl = Frame Then
                Set frm = ac_in.Controls.Add(bstrProgID:=MSFormsProgID(ac_ctl) _
                                           , Name:=ac_name _
                                           , Visible:=ac_visible)
                Set AddControl = frm
            Else
                Set ctl = ac_in.Controls.Add(bstrProgID:=MSFormsProgID(ac_ctl) _
                                           , Name:=ac_name _
                                           , Visible:=ac_visible)
                Set AddControl = ctl
            End If
        End If
    End If
    Set AddControl = ctl
    
xt: Exit Function
    
eh: If ErrMsg(ErrSrc(PROC)) = vbYes Then: Stop: Resume
End Function

Private Function AddCursor(CursorType As CursorTypes)
' -------------------------------------------------------------------------------
' To set a cursor
' -------------------------------------------------------------------------------
    If Not IsCursorType(CursorType) Then
        SetCursor LoadCursor(0, CursorType)
        Sleep 200 ' wait a bit, needed for rendering
    End If
End Function

Private Sub AdjustedParentsWidthAndHeight(ByVal ctrl As MSForms.Control)
' ------------------------------------------------------------------------------
' Adjust the width and height of all parent frames starting with the parent of
' the provided control (ctrl) by considering the control's width and height and
' a possibly already applied vertical and/or horizontal scroll-bar.
' ------------------------------------------------------------------------------
    Dim FrmParent   As Variant
    
    On Error Resume Next
    Set FrmParent = ctrl.Parent
    If Err.Number <> 0 Then
        On Error GoTo eh
        GoTo xt
    End If
    
    Do
        If IsForm(FrmParent) Then
            If Not ScrollVscrollApplied(FrmParent) Then
                FrmParent.Width = ContentWidth(FrmParent) + 5
                FrmParent.Height = ctrl.Top + ContentHeight(FrmParent) + 30
            End If
        ElseIf IsFrameOrForm(FrmParent) Then
            If Not ScrollVscrollApplied(FrmParent) Then
                FrmParent.Width = ContentWidth(FrmParent)
                FrmParent.Height = ContentHeight(FrmParent)
            End If
        End If
        If IsForm(FrmParent) Then Exit Do
        Set FrmParent = FrmParent.Parent
    Loop
                
xt: Exit Sub
eh:
End Sub

Private Sub AdjustTopPositions()
' ------------------------------------------------------------------------------
' - Adjusts each visible control's top position considering its current height.
' ------------------------------------------------------------------------------
    Const PROC = "AdjustTopPositions"
    
    On Error GoTo eh
    Dim i                   As Long
    Dim MaxTextFrameWidth   As Single
    Dim TopPosTextFrame     As Single
    Dim TopPosNextSect      As Single
    
    MaxTextFrameWidth = MaxUsedWidthTextFrames
    TopPosNextSect = 0 ' The top sections top position
    frmMsectsArea.Top = 0
    
    For i = 1 To lNoOfDesignedMsgSects
        TopPosTextFrame = 0
        If dctMsectsFrm(i).Visible Then
            ProvideSectionItems i
            '~~ Adjust top positions of message section and its items
            
            '~~ Top pos section label
            If MsectLbl.Visible Then
                '~~ Adjust the section label
                With MsectLbl
                    .Top = 0
                    TopPosTextFrame = AdjustToVgrid(.Top + .Height)
                    MsectLbl.Width = Me.Width - .Left - 5
                End With
            End If
            
            '~~ Top pos TextBox
            MsectTbx.Top = 0
            '~~ Top pos Text Frasme
            MsectTbxFrm.Top = TopPosTextFrame
            
            '~~ Top pos Message Section
            With MsectFrm
                .Top = TopPosNextSect
                .Width = MaxTextFrameWidth
                If MsectTbxFrm.Visible And Not ScrollVscrollApplied(MsectTbxFrm) And .Height <> MsectTbxFrm.Top + MsectTbxFrm.Height Then
                    .Height = MsectTbxFrm.Top + MsectTbxFrm.Height
                End If
                TopPosNextSect = AdjustToVgrid(.Top + .Height + VSPACE_SECTIONS)
            End With
            TimedDoEvents ErrSrc(PROC)   ' to properly h-align the text
            AdjustedParentsWidthAndHeight MsectFrm
        End If
    Next i
    
    '~~ Top position Message Area
    If frmBttnsArea.Visible And frmMsectsArea.Visible Then
        frmBttnsArea.Top = frmMsectsArea.Top + frmMsectsArea.Height + VSPACE_AREAS
        If Me.Height <> frmBttnsArea.Top + frmBttnsArea.Height + VSPACE_AREAS Then
            Me.Height = frmBttnsArea.Top + frmBttnsArea.Height + VSPACE_AREAS
        End If
    ElseIf frmBttnsArea.Visible And Not frmMsectsArea.Visible Then
        frmBttnsArea.Top = VSPACE_AREAS
        FrameCenterHorizontal frmBttnsArea
        If Me.Height <> frmBttnsArea.Top + frmBttnsArea.Height + VSPACE_AREAS Then
            Me.Height = frmBttnsArea.Top + frmBttnsArea.Height + VSPACE_AREAS
            Me.Height = Me.Height + VSPACE_BOTTOM
        End If
    ElseIf Not frmBttnsArea.Visible And frmMsectsArea.Visible And Me.Height <> frmMsectsArea.Top + frmMsectsArea.Height + VSPACE_AREAS Then
        Me.Height = frmMsectsArea.Top + frmMsectsArea.Height + VSPACE_AREAS
        Me.Height = Me.Height + VSPACE_BOTTOM
    End If
    
xt: Exit Sub
    
eh: If ErrMsg(ErrSrc(PROC)) = vbYes Then: Stop: Resume
End Sub

Public Function AdjustToVgrid(ByVal atvg_si As Single, _
                     Optional ByVal atvg_threshold As Single = 1.5, _
                     Optional ByVal atvg_grid As Single = 6) As Single
' -------------------------------------------------------------------------------
' Returns the value (atvg_si) as a Single value which is a multiple of the grid
' value (atvg_grid), which defaults to 6. To avoid irritating vertical spacing
' a certain threshold (atvg_threshold) is considered which defaults to 1.5.
' The returned value can be used to vertically align a control's top position to
' the grid or adjust its height to the grid.
' Examples for the function of the threshold:
'  7.5 < si >= 0   results to 6
' 13.5 < si >= 7.5 results in 12
' -------------------------------------------------------------------------------
    AdjustToVgrid = (Int((atvg_si - atvg_threshold) / atvg_grid) * atvg_grid) + atvg_grid
End Function

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

Private Function AppliedBttnRows() As Dictionary
' ------------------------------------------------------------------------------
' Returns a Dictionary of the visible button rows with the row
' frame as the key and the applied/visible buttons therein as item.
' ------------------------------------------------------------------------------
    Const PROC = "AppliedBttnRows"
    
    On Error GoTo eh
    Dim dct         As New Dictionary
    Dim lRow        As Long
    Dim frmRow      As MSForms.Frame
    Dim v           As Variant
    Dim lButtons    As Long
    Dim cmb         As MSForms.CommandButton
    
    For lRow = 1 To dctBttnsRow.Count
        Set frmRow = dctBttnsRow(lRow)
        If frmRow.Visible Then
            lButtons = 0
            For Each v In dctBttns
                If Split(v, "-")(0) = lRow Then
                    Set cmb = dctBttns(v)
                    If cmb.Visible Then lButtons = lButtons + 1
                End If
            Next v
            dct.Add frmRow, lButtons
        End If
    Next lRow
    Set AppliedBttnRows = dct

xt: Exit Function
    
eh: If ErrMsg(ErrSrc(PROC)) = vbYes Then: Stop: Resume
End Function

Public Sub AutoSizeTextBox(ByRef as_tbx As MSForms.TextBox, _
                           ByVal as_text As String, _
                  Optional ByVal as_width_limit As Single = 0, _
                  Optional ByVal as_append As Boolean = False, _
                  Optional ByVal as_append_margin As String = vbNullString)
' ------------------------------------------------------------------------------
' Common AutoSize service for an MsForms.TextBox providing a width and height
' for the TextBox (as_tbx) by considering:
' - When a width limit is provided (as_width_limit > 0) the width is regarded a
'   fixed maximum and thus the height is auto-sized by means of WordWrap=True.
' - When no width limit is provided (the default) WordWrap=False and thus the
'   width of the TextBox is determined by the longest line.
' - When a maximum width is provided (as_width_max > 0) and the parent of the
'   TextBox is a frame a horizontal scrollbar is applied for the parent frame.
' - When a maximum height is provided (as_heightmax > 0) and the parent of the
'   TextBox is a frame a vertical scrollbar is applied for the parent frame.
' - When a minimum width (as_width_min > 0) or a minimum height (as_height_min
'   > 0) is provided the size of the textbox is set correspondingly. This
'   option is specifically usefull when text is appended to avoid much flicker.
'
' Uses: AdjustToVgrid
'
' W. Rauschenberger Berlin April 2022
' ------------------------------------------------------------------------------
    
    With as_tbx
        .MultiLine = True
        If as_width_limit > 0 Then
            '~~ AutoSize the height of the TextBox considering the limited width
            '~~ (applied for proportially spaced text where the width determines the height)
            .WordWrap = True
            .AutoSize = False
            .Width = as_width_limit - 2 ' the readability space is added later
            
            If Not as_append Then
                .Value = as_text
            Else
                If .Value = vbNullString Then
                    .Value = as_text
                Else
                    .Value = .Value & as_append_margin & vbLf & as_text
                End If
            End If
            .AutoSize = True
        Else
            '~~ AutoSize the height and width of the TextBox
            '~~ (applied for mono-spaced text where the longest line defines the width)
            .MultiLine = True
            .WordWrap = False ' the means to limit the width
            .AutoSize = True
            If Not as_append Then
                .Value = as_text
            Else
                If .Value = vbNullString Then
                    .Value = as_text
                Else
                    .Value = .Value & vbLf & as_text
                End If
            End If
        End If
        .Width = .Width + 2 ' readability space
        .Height = AdjustToVgrid(.Height, 0)
    End With
        
xt: Exit Sub

End Sub

Private Function BttnsArea() As MSForms.Frame
' ------------------------------------------------------------------------------
' Returns the Buttons area Frame, created if yet not existing.
' ------------------------------------------------------------------------------

    If frmBttnsArea Is Nothing Then
        Set frmBttnsArea = Me.AddControl(ac_ctl:=Frame, ac_visible:=True, ac_name:="frBttnsArea")
        With frmBttnsArea
            .Top = 0
            If Not frmMsectsArea Is Nothing Then
                .Top = AdjustToVgrid(frmMsectsArea.Top + frmMsectsArea.Height) + VSPACE_AREAS
            End If
            .Left = 0
            .Height = 50
            .Width = Me.InsideWidth
            .Visible = True
        End With
        VisualizeCtl frmBttnsArea, VISLZE_BCKCLR_AREA
    End If
    Set BttnsArea = frmBttnsArea

End Function

Private Function BttnsFrm() As MSForms.Frame
' ------------------------------------------------------------------------------
' Returns the Frame of the message buttons, created in the BttnsArea if yet
' not existing.
' ------------------------------------------------------------------------------
    
    If frmBttnsFrm Is Nothing Then
        Set frmBttnsFrm = AddControl(ac_ctl:=Frame, ac_in:=BttnsArea, ac_name:="frBttnsFrame")
        VisualizeCtl frmBttnsFrm, VISLZE_BCKCLR_BTTNS_FRM
    End If
    Set BttnsFrm = frmBttnsFrm

End Function

Private Function BttnsRowBttn(ByVal brb_row As Long, _
                              ByVal brb_bttn As Long) As MSForms.CommandButton
' ------------------------------------------------------------------------------
' Returns the Command Button (brb_bttn) in Row (brb_row), created in Button
' row (brb_row) when not existing.
' ------------------------------------------------------------------------------
    Dim cmb As MSForms.CommandButton
    If Not dctBttnsRowBttns.Exists(brb_row & "-" & brb_bttn) Then
        Set cmb = AddControl(ac_ctl:=CommandButton, ac_in:=BttnsRowFrm(brb_row), ac_name:="cb" & brb_row & brb_bttn)
        With cmb
            .Top = 0
            .Height = 18
            .Width = DFLT_BTTN_MIN_WIDTH
            .Left = (.Width * (brb_bttn - 1)) + (HSPACE_BTTNS * (brb_bttn - 1))
        End With
        dctBttnsRowBttns.Add brb_row & "-" & brb_bttn, cmb
    End If
    Set BttnsRowBttn = dctBttnsRowBttns(brb_row & "-" & brb_bttn)

End Function

Private Function BttnsRowFrm(ByVal brf_row As Long) As MSForms.Frame
' ------------------------------------------------------------------------------
' Returns the Frame of the buttons row (brf_row), created in the BttnsFrm if yet
' not existing.
' ------------------------------------------------------------------------------
    Dim frm As MSForms.Frame
    
    If Not dctBttnsRow.Exists(brf_row) Then
        Set frm = AddControl(ac_ctl:=Frame, ac_in:=BttnsFrm, ac_name:="frBttnsRow" & brf_row)
        VisualizeCtl frm, VISLZE_BCKCLR_BTTNS_ROW_FRM
        dctBttnsRow.Add brf_row, frm
    End If
    Set BttnsRowFrm = dctBttnsRow(brf_row)

End Function

Private Sub ButtonClicked(ByVal cmb As MSForms.CommandButton)
' ------------------------------------------------------------------------------
' Provides the clicked button's (cmb) caption string or value for the caller
' via mMsg.Replied and additionally via the ReplyValue Property. When there is
' only one applied reply button the form is unloaded with the click. Otherwise the form is just hidden waiting for
' the caller to obtain the return value via the ReplyValue Property which is
' either the clicked button's (cmb) caption stringor index which then unloads the form.
' ------------------------------------------------------------------------------
    On Error Resume Next
    If bModeLess Then
        '~~ When the form is displayed "Modelss" there may be an Application.Run action provided
        '~~ for the clicked button
        ApplicationRunViaButton cmb.Caption
    Else
        '~~ When the form is displayed "Modal" the clicked button is returned and the form is unloaded
        If bReplyWithIndex _
        Then mMsg.RepliedWith = ClickedButtonIndex(cmb) _
        Else mMsg.RepliedWith = AppliedButtonRetVal(cmb)  ' global variable of calling module mMsg
        Unload Me
    End If
    
End Sub

Private Function ButtonsProvided() As Boolean
    ButtonsProvided = cllMsgBttns.Count > 0
End Function

' ------------------------------------------------------------
' The reply button click event is the only code using the
' control's name - which unfortunately this cannot be avioded.
' ------------------------------------------------------------
Private Sub cmb11_Click():  ButtonClicked Me.cmb11:   End Sub

Private Sub cmb12_Click():  ButtonClicked Me.cmb12:   End Sub

Private Sub cmb13_Click():  ButtonClicked Me.cmb13:   End Sub

Private Sub cmb14_Click():  ButtonClicked Me.cmb14:   End Sub

Private Sub cmb15_Click():  ButtonClicked Me.cmb15:   End Sub

Private Sub cmb16_Click():  ButtonClicked Me.cmb16:   End Sub

Private Sub cmb17_Click():  ButtonClicked Me.cmb17:   End Sub

Private Sub cmb21_Click():  ButtonClicked Me.cmb21:   End Sub

Private Sub cmb22_Click():  ButtonClicked Me.cmb22:   End Sub

Private Sub cmb23_Click():  ButtonClicked Me.cmb23:   End Sub

Private Sub cmb24_Click():  ButtonClicked Me.cmb24:   End Sub

Private Sub cmb25_Click():  ButtonClicked Me.cmb25:   End Sub

Private Sub cmb26_Click():  ButtonClicked Me.cmb26:   End Sub

Private Sub cmb27_Click():  ButtonClicked Me.cmb27:   End Sub

Private Sub cmb31_Click():  ButtonClicked Me.cmb31:   End Sub

Private Sub cmb32_Click():  ButtonClicked Me.cmb32:   End Sub

Private Sub cmb33_Click():  ButtonClicked Me.cmb33:   End Sub

Private Sub cmb34_Click():  ButtonClicked Me.cmb34:   End Sub

Private Sub cmb35_Click():  ButtonClicked Me.cmb35:   End Sub

Private Sub cmb36_Click():  ButtonClicked Me.cmb36:   End Sub

Private Sub cmb37_Click():  ButtonClicked Me.cmb37:   End Sub

Private Sub cmb41_Click():  ButtonClicked Me.cmb41:   End Sub

Private Sub cmb42_Click():  ButtonClicked Me.cmb42:   End Sub

Private Sub cmb43_Click():  ButtonClicked Me.cmb43:   End Sub

Private Sub cmb44_Click():  ButtonClicked Me.cmb44:   End Sub

Private Sub cmb45_Click():  ButtonClicked Me.cmb45:   End Sub

Private Sub cmb46_Click():  ButtonClicked Me.cmb46:   End Sub

Private Sub cmb47_Click():  ButtonClicked Me.cmb47:   End Sub

Private Sub cmb51_Click():  ButtonClicked Me.cmb51:   End Sub

Private Sub cmb52_Click():  ButtonClicked Me.cmb52:   End Sub

Private Sub cmb53_Click():  ButtonClicked Me.cmb53:   End Sub

Private Sub cmb54_Click():  ButtonClicked Me.cmb54:   End Sub

Private Sub cmb55_Click():  ButtonClicked Me.cmb55:   End Sub

Private Sub cmb56_Click():  ButtonClicked Me.cmb56:   End Sub

Private Sub cmb57_Click():  ButtonClicked Me.cmb57:   End Sub

Private Sub cmb61_Click():  ButtonClicked Me.cmb61:   End Sub

Private Sub cmb62_Click():  ButtonClicked Me.cmb62:   End Sub

Private Sub cmb63_Click():  ButtonClicked Me.cmb63:   End Sub

Private Sub cmb64_Click():  ButtonClicked Me.cmb64:   End Sub

Private Sub cmb65_Click():  ButtonClicked Me.cmb65:   End Sub

Private Sub cmb66_Click():  ButtonClicked Me.cmb66:   End Sub

Private Sub cmb67_Click():  ButtonClicked Me.cmb67:   End Sub

Private Sub cmb71_Click():  ButtonClicked Me.cmb71:   End Sub

Private Sub cmb72_Click():  ButtonClicked Me.cmb72:   End Sub

Private Sub cmb73_Click():  ButtonClicked Me.cmb73:   End Sub

Private Sub cmb74_Click():  ButtonClicked Me.cmb74:   End Sub

Private Sub cmb75_Click():  ButtonClicked Me.cmb75:   End Sub

Private Sub cmb76_Click():  ButtonClicked Me.cmb76:   End Sub

Private Sub cmb77_Click():  ButtonClicked Me.cmb77:   End Sub

Private Sub Collect(ByRef col_into As Variant, _
                    ByVal col_with_parent As Variant, _
                    ByVal col_cntrl_type As String, _
                    ByVal col_set_height As Single, _
                    ByVal col_set_width As Single, _
           Optional ByVal col_set_visible As Boolean = False)
' ------------------------------------------------------------------------------
' Setup of a Collection (col_into) with all type (col_cntrl_type) controls
' with a parent (col_with_parent) as Collection (col_into) by assigning the
' an initial height (col_set_height) and width (col_set_width).
' ------------------------------------------------------------------------------
    Const PROC = "Collect"
    
    On Error GoTo eh
    Dim ctl As MSForms.Control
    Dim v   As Variant
    Dim i   As Long
    
    lBackColor = Me.BackColor

    Select Case TypeName(col_with_parent)
        Case "Dictionary"
            '~~ Parent is each frame in the collection
            For Each v In col_with_parent
                For Each ctl In Me.Controls
                    If TypeName(ctl) = col_cntrl_type And ctl.Parent Is col_with_parent(v) Then
                        With ctl
                            .Visible = col_set_visible
                            .Height = col_set_height
                            .Width = col_set_width
                        End With
                        i = col_into.Count + 1
                        col_into.Add i, ctl
                    End If
               Next ctl
            Next v
        Case Else
            For Each ctl In Me.Controls
                If TypeName(ctl) = col_cntrl_type And ctl.Parent Is col_with_parent Then
                    With ctl
                        .Visible = col_set_visible
                        .Height = col_set_height
                        .Width = col_set_width
                    End With
                    Select Case TypeName(col_into)
                        Case "Dictionary"
                            i = col_into.Count + 1
                            col_into.Add i, ctl
'                            Debug.Print col_into.Count & ": " & ctl.Name
                        Case Else
                            Set col_into = ctl
                            Exit For
                    End Select
                End If
            Next ctl
    End Select

xt: Exit Sub
    
eh: If ErrMsg(ErrSrc(PROC)) = vbYes Then: Stop: Resume
End Sub

Private Sub CollectDesignControls()
' ----------------------------------------------------------------------
' Collects all designed controls without considering any control's name.
' ----------------------------------------------------------------------
    Const PROC = "CollectDesignControls"
    
    On Error GoTo eh
    Dim v           As Variant
    Dim frm         As MSForms.Frame
    Dim lRow        As Long
    Dim lBttn       As Long
    Dim cmb         As MSForms.CommandButton
    Dim sKey        As String
    Dim PntsWidth   As Single:  PntsWidth = mMsg.Pnts(Me.MsgWidthMin - mMsg.Prcnt(Me.Width - Me.InsideWidth, "W"), "W")
    
    Collect col_into:=NewDict(dctAreas) _
          , col_cntrl_type:="Frame" _
          , col_with_parent:=Me _
          , col_set_height:=10 _
          , col_set_width:=PntsWidth
    Set frmMsectsArea = dctAreas(1)
    Set frmBttnsArea = dctAreas(2)
    VisualizeCtl frmMsectsArea, VISLZE_BCKCLR_AREA
    VisualizeCtl frmBttnsArea, VISLZE_BCKCLR_AREA
    
    Collect col_into:=NewDict(dctMsectsFrm) _
          , col_cntrl_type:="Frame" _
          , col_with_parent:=frmMsectsArea _
          , col_set_height:=50 _
          , col_set_width:=frmMsectsArea.Width - siHmarginFrames
    For Each v In dctMsectsFrm
        VisualizeCtl dctMsectsFrm(v), VISLZE_BCKCLR_MSEC_FRM
    Next v
    NoOfDesignedMsgSects = dctMsectsFrm.Count
    
    Collect col_into:=NewDict(dctMsectsLbl) _
          , col_cntrl_type:="Label" _
          , col_with_parent:=dctMsectsFrm _
          , col_set_height:=15 _
          , col_set_width:=frmMsectsArea.Width - (siHmarginFrames * 2)
    For Each v In dctMsectsLbl
        VisualizeCtl dctMsectsLbl(v), VISLZE_BCKCLR_MSEC_LBL
    Next v
    
    Collect col_into:=NewDict(dctMsectsTbxFrm) _
          , col_cntrl_type:="Frame" _
          , col_with_parent:=dctMsectsFrm _
          , col_set_height:=20 _
          , col_set_width:=frmMsectsArea.Width - (siHmarginFrames * 2)
    For Each v In dctMsectsTbxFrm
        VisualizeCtl dctMsectsTbxFrm(v), VISLZE_BCKCLR_MSEC_TBX_FRM
    Next v
    
    Collect col_into:=NewDict(dctMsectsTbx) _
          , col_cntrl_type:="TextBox" _
          , col_with_parent:=dctMsectsTbxFrm _
          , col_set_height:=20 _
          , col_set_width:=frmMsectsArea.Width - (siHmarginFrames * 3)
    For Each v In dctMsectsTbx
        VisualizeCtl dctMsectsTbx(v), VISLZE_BCKCLR_MSEC_TBX
    Next v
        
    Collect col_into:=frmBttnsFrm _
          , col_cntrl_type:="Frame" _
          , col_with_parent:=frmBttnsArea _
          , col_set_height:=10 _
          , col_set_width:=10 _
          , col_set_visible:=True ' minimum is one button
    VisualizeCtl frmBttnsFrm, VISLZE_BCKCLR_BTTNS_FRM
    
    Collect col_into:=NewDict(dctBttnsRow) _
          , col_cntrl_type:="Frame" _
          , col_with_parent:=frmBttnsFrm _
          , col_set_height:=10 _
          , col_set_width:=10 _
          , col_set_visible:=False ' minimum is one button
    For Each v In dctBttnsRow
        VisualizeCtl dctBttnsRow(v), VISLZE_BCKCLR_BTTNS_ROW_FRM
    Next v
        
    NewDict dctBttns
    For lRow = 1 To dctBttnsRow.Count
        Set frm = dctBttnsRow(lRow)
        For lBttn = 0 To frm.Controls.Count - 1
            Set cmb = frm.Controls(lBttn)
            sKey = lRow & "-" & lBttn + 1
            dctBttns.Add sKey, cmb
'            Debug.Print "Button " & sKey & ": " & cmb.Name
        Next lBttn
    Next lRow
    
    NewDict AppliedBttns
    NewDict AppliedBttnsRetVal
       
xt: Exit Sub
    
eh: If ErrMsg(ErrSrc(PROC)) = vbYes Then: Stop: Resume
End Sub

Public Function ContentHeight(ByVal ch_frame_form As Variant, _
                     Optional ByVal ch_visible_only As Boolean = True) As Single
' ------------------------------------------------------------------------------
' Returns the height of the largest control in a Frame or Form (ch_frame_form)
' which is the maximum value of the Controls Top + Height.
' ------------------------------------------------------------------------------
    Const PROC = "ContzentHeight"
    
    On Error GoTo eh
    Dim ctl As MSForms.Control
    Dim i   As Long
    
    If Not IsFrameOrForm(ch_frame_form) _
    Then Err.Raise AppErr(1), ErrSrc(PROC), "The provided argument is neither a Frame nor a Form - and thus has no controls!"
    
    For Each ctl In ch_frame_form.Controls
        With ctl
            If .Parent Is ch_frame_form Then
                If ch_visible_only Then
                    If ctl.Visible Then
                        ContentHeight = Max(ContentHeight, .Top + .Height)
                        i = i + 1
                    End If
                Else
                    ContentHeight = Max(ContentHeight, .Top + .Height)
                    i = i + 1
                End If
            End If
        End With
    Next ctl
        
xt: Exit Function
    
eh: If ErrMsg(ErrSrc(PROC)) = vbYes Then: Stop: Resume
End Function

Public Function ContentWidth(ByVal cw_frame_form As Variant, _
                    Optional ByVal cw_visible_only As Boolean = True) As Single
' ------------------------------------------------------------------------------
' Returns the width of the largest control in a Frame or Form (cw_frame_form)
' which is the maximum value of the Controls Left + Width.
' ------------------------------------------------------------------------------
    Const PROC = "ContentWidth"
    
    On Error GoTo eh
    Dim ctl As MSForms.Control
    Dim i   As Long
    
    If Not IsFrameOrForm(cw_frame_form) _
    Then Err.Raise AppErr(1), ErrSrc(PROC), "The provided argument is neither a Frame nor a Form - and thus has no controls!"

    For Each ctl In cw_frame_form.Controls
        With ctl
            If .Parent Is cw_frame_form Then
                If cw_visible_only Then
                    If ctl.Visible Then
                        ContentWidth = Max(ContentWidth, (.Left + .Width))
                        i = i + 1
                    End If
                Else
                    ContentWidth = Max(ContentWidth, (.Left + .Width))
                    i = i + 1
                End If
            End If
        End With
    Next ctl
    
xt: Exit Function
    
eh: If ErrMsg(ErrSrc(PROC)) = vbYes Then: Stop: Resume
End Function

Private Sub ConvertPixelsToPoints(ByVal x_dpi As Single, ByVal y_dpi As Single, _
                                  ByRef x_pts As Single, ByRef y_pts As Single)
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
    x_pts = x_dpi * TWIPSPERINCH / 20 / PixelsPerInchX
    y_pts = y_dpi * TWIPSPERINCH / 20 / PixelsPerInchY

End Sub

Private Function CtlExists(ByVal ce_name As String) As Boolean
    Dim ctl As MSForms.Control
    For Each ctl In Me.Controls
        If ctl.Name = ce_name Then
            CtlExists = True
            Exit For
        End If
    Next ctl
End Function

Private Function ErrMsg(ByVal err_source As String, _
               Optional ByVal err_no As Long = 0, _
               Optional ByVal err_dscrptn As String = vbNullString, _
               Optional ByVal err_line As Long = 0) As Variant
' ------------------------------------------------------------------------------
' Minimum error message display where neither mErH.ErrMsg nor mMsg.ErrMsg is
' appropriate. This is the case here because this component is used by the other
' two components which implies the danger of a loop.
'
' W. Rauschenberger Berlin, Nov 2021
' ------------------------------------------------------------------------------
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
    ErrBttns = vbYesNoCancel
    ErrText = ErrText & vbLf & vbLf & _
              "Debugging:" & vbLf & _
              "Yes    = Resume error line" & vbLf & _
              "No     = Resume Next (skip error line)" & vbLf & _
              "Cancel = Terminate"
#Else
    ErrBttns = vbCritical
#End If
    
    ErrMsg = MsgBox(Title:=ErrTitle _
                  , Prompt:=ErrText _
                  , Buttons:=ErrBttns)
xt: Exit Function

End Function

Private Function ErrSrc(ByVal sProc As String) As String
    ErrSrc = "fMsg." & sProc
End Function

Private Sub FrameCenterHorizontal(ByVal center_frame As MSForms.Frame, _
                         Optional ByVal within_frame As MSForms.Frame = Nothing, _
                         Optional ByVal left_margin As Single = 0)
' ------------------------------------------------------------------------------
' Center the frame (center_frame) horizontally within the frame (within_frame)
' - which defaults to the UserForm when not provided.
' ------------------------------------------------------------------------------
    
    If within_frame Is Nothing Then
        center_frame.Left = (Me.InsideWidth - center_frame.Width) / 2
    Else
        center_frame.Left = (within_frame.Width - center_frame.Width) / 2
    End If
    If center_frame.Left = 0 Then center_frame.Left = left_margin
End Sub

Private Sub GetScreenMetrics(ByRef left_pts As Single, _
                             ByRef top_pts As Single, _
                             ByRef width_pts As Single, _
                             ByRef height_pts As Single)
' ------------------------------------------------------------
' Get coordinates of top-left corner and size of entire screen
' (stretched over all monitors) and convert to Points.
' ------------------------------------------------------------
    
    ConvertPixelsToPoints x_dpi:=GetSystemMetrics32(SM_XVIRTUALSCREEN), x_pts:=left_pts, _
                          y_dpi:=GetSystemMetrics32(SM_YVIRTUALSCREEN), y_pts:=top_pts
                          
    ConvertPixelsToPoints x_dpi:=GetSystemMetrics32(SM_CXVIRTUALSCREEN), x_pts:=width_pts, _
                          y_dpi:=GetSystemMetrics32(SM_CYVIRTUALSCREEN), y_pts:=height_pts

End Sub

Private Sub HandCursorForLink(ByVal hc_section As Long)
    If MsgLabel(hc_section).OpenWhenClicked <> vbNullString _
    Then AddCursor IDC_HAND
End Sub

Private Sub IndicateFrameCaptionsSetup(Optional ByVal b As Boolean = True)
' ----------------------------------------------------------------------------
' When False (the default) captions are removed from all frames, else they
' remain visible for testing purpose
' ----------------------------------------------------------------------------
            
    Dim ctl As MSForms.Control
       
    If Not b Then
        For Each ctl In Me.Controls
            If TypeName(ctl) = "Frame" Then
                ctl.Caption = vbNullString
            End If
        Next ctl
    End If

End Sub

Private Function IsCursorType(CursorType As CursorTypes) As Boolean
' -------------------------------------------------------------------------------
' To determine if the cursor is already set
' -------------------------------------------------------------------------------
    Dim CursorHandle    As Long:        CursorHandle = LoadCursor(ByVal 0&, CursorType)
    Dim Cursor          As CursorInfo:  Cursor.cbSize = Len(Cursor)
    Dim CursorInfo      As Boolean:     CursorInfo = GetCursorInfo(Cursor)

    If Not CursorInfo Then
        IsCursorType = False
        Exit Function
    End If

    IsCursorType = (Cursor.hCursor = CursorHandle)
End Function

Private Function IsForm(ByVal v As Object) As Boolean
    Dim o As Object
    On Error Resume Next
    Set o = v.Parent
    IsForm = Err.Number <> 0
End Function

Private Function IsFrameOrForm(ByVal v As Object) As Boolean
    IsFrameOrForm = TypeOf v Is MSForms.UserForm Or TypeOf v Is MSForms.Frame
End Function

Private Function IsUserForm(ByVal is_obj As Object) As Boolean
      IsUserForm = TypeOf is_obj Is MSForms.UserForm
End Function

Private Sub laMsgSection1Labe2_Click():     OpenClickedLabelItem 2: End Sub

Private Sub laMsgSection1Labe3_Click():     OpenClickedLabelItem 3: End Sub

Private Sub laMsgSection1Labe4_Click():     OpenClickedLabelItem 4: End Sub

Private Sub laMsgSection1Label_Click():     OpenClickedLabelItem 1: End Sub

Private Sub laMsgSection1Label_MouseMove(ByVal Button As Integer, ByVal Shift As Integer, ByVal X As Single, ByVal Y As Single):        HandCursorForLink 1:    End Sub

Private Sub laMsgSection2Label_MouseMove(ByVal Button As Integer, ByVal Shift As Integer, ByVal X As Single, ByVal Y As Single):        HandCursorForLink 2:    End Sub

Private Sub laMsgSection3Label_MouseMove(ByVal Button As Integer, ByVal Shift As Integer, ByVal X As Single, ByVal Y As Single):        HandCursorForLink 3:    End Sub

Private Sub laMsgSection4Label_MouseMove(ByVal Button As Integer, ByVal Shift As Integer, ByVal X As Single, ByVal Y As Single):        HandCursorForLink 4:    End Sub

Private Function Max(ParamArray va() As Variant) As Variant
' ----------------------------------------------------------------------------
' Returns the maximum value of all values provided (va).
' ----------------------------------------------------------------------------
    Dim v As Variant
    
    Max = va(LBound(va)): If LBound(va) = UBound(va) Then Exit Function
    For Each v In va
        If v > Max Then Max = v
    Next v
    
End Function

Private Function MaxUsedWidthTextFrames() As Single
' ------------------------------------------------------------------------------
' Return the maximum width of all existing and visible TextBox Frames.
' ------------------------------------------------------------------------------
    Const PROC = "MaxUsedWidthTextFrames"
    
    On Error GoTo eh
    Dim frm As MSForms.Frame
    Dim v   As Variant
    
    For Each v In dctMsectsTbxFrm
        If TypeOf dctMsectsTbxFrm(v) Is MSForms.Frame Then
            Set frm = dctMsectsTbxFrm(v)
            If frm.Visible Then
                MaxUsedWidthTextFrames = Max(MaxUsedWidthTextFrames, frm.Width)
            End If
        End If
    Next v

xt: Exit Function

eh: If ErrMsg(ErrSrc(PROC)) = vbYes Then: Stop: Resume
End Function

Private Function MaxWidthMsgSect(ByVal frm_area As MSForms.Frame) As Single
' ------------------------------------------------------------------------------
' The maximum usable message section width depends on the maximum message area
' width whether or not the area frame (frm_area) has a vertical scrollbar. A
' vertical scrollbar reduces the available space by the required space for the
' vertical scrollbar.
' ------------------------------------------------------------------------------
    If frm_area.ScrollBars = fmScrollBarsVertical Or frm_area.ScrollBars = fmScrollBarsBoth _
    Then MaxWidthMsgSect = MaxWidthMsgArea - SCROLL_V_WIDTH _
    Else MaxWidthMsgSect = MaxWidthMsgArea
End Function

Private Function MaxWidthMsgTextFrame(ByVal frm_area As MSForms.Frame, _
                                      ByVal frm_section As MSForms.Frame) As Single
' ------------------------------------------------------------------------------
' The maximum usable message text width depends on the maximum message section
' width and whether or not the section (frm_section) has a vertical scrollbar
' which reduces the available space by its required width.
' ------------------------------------------------------------------------------
    If frm_section.ScrollBars = fmScrollBarsVertical Or frm_section.ScrollBars = fmScrollBarsBoth _
    Then MaxWidthMsgTextFrame = MaxWidthMsgSect(frm_area) - SCROLL_V_WIDTH _
    Else MaxWidthMsgTextFrame = MaxWidthMsgSect(frm_area)
End Function

Private Function MaxWidthSectTxtBox(ByVal frm_text As MSForms.Frame) As Single
' ------------------------------------------------------------------------------
' The maximum with of a sections text-box depends on whether or not the frame of
' the TextBox (frm_text) has a vertical scrollbar which reduces the available
' space by its required width.
' ------------------------------------------------------------------------------
    If frm_text.ScrollBars = fmScrollBarsVertical Or frm_text.ScrollBars = fmScrollBarsBoth _
    Then MaxWidthSectTxtBox = frm_text.Width - SCROLL_V_WIDTH _
    Else MaxWidthSectTxtBox = frm_text.Width
End Function

Private Function Min(ParamArray va() As Variant) As Variant
' ----------------------------------------------------------------------------
' Returns the minimum (smallest) of all provided values.
' ----------------------------------------------------------------------------
    Dim v   As Variant
    
    Min = va(LBound(va)): If LBound(va) = UBound(va) Then Exit Function
    For Each v In va
        If v < Min Then Min = v
    Next v
    
End Function

Public Sub MonitorHeader()
' ------------------------------------------------------------------------------
' Establishes an (optional) header in the monitor window.
' ------------------------------------------------------------------------------
    Const PROC = "MonitorHeader"
    
    Dim siTop As Single
    
    On Error GoTo eh
    If cllSteps Is Nothing Then Me.MonitorInit
    
    If tbxHeader Is Nothing Then
        Set tbxHeader = AddControl(ac_ctl:=TextBox, ac_visible:=False, ac_name:="tbMonitorHeader")
        With tbxHeader
            .Top = 6
            .Left = 0
            .Height = 18
            .Width = Me.InsideWidth
            .Visible = True
        End With
        VisualizeCtl tbxHeader, VISLZE_BCKCLR_MSEC_TBX
    End If
    SetupTextFont tbxHeader, enMonHeader
    If TextMonitorHeader.MonoSpaced Then
        AutoSizeTextBox as_tbx:=tbxHeader _
                      , as_text:=TextMonitorHeader.Text
    Else
        AutoSizeTextBox as_tbx:=tbxHeader _
                      , as_text:=TextMonitorHeader.Text _
                      , as_width_limit:=Me.InsideWidth
    End If
    
    With tbxHeader
        siTop = AdjustToVgrid(.Top + .Height)
    End With
    
    '~~ Adjust the subsequent steps' top position
    With frmSteps
        .Top = siTop
        siTop = AdjustToVgrid(.Top + .Height)
    End With
    
    If Not tbxFooter Is Nothing Then tbxFooter.Top = siTop
    Me.Height = ContentHeight(frmSteps.Parent) + 35

xt: Exit Sub

eh: If ErrMsg(ErrSrc(PROC)) = vbYes Then: Stop: Resume
End Sub

Private Sub MonitorEstablishStep(ByRef ms_top As Single)
' ------------------------------------------------------------------------------
' Adds a monitor step TextBox to the frmSteps Frame, enqueues it into the
' cllSteps queue and adjusts the top position to (ms_top).
' ------------------------------------------------------------------------------
    Const PROC = "MonitorEstablishStep"
    Const CTL_NAME As String = "tbMonitorStep"
    
    Set tbxStep = AddControl(ac_ctl:=TextBox _
                           , ac_in:=frmSteps _
                           , ac_visible:=False _
                           , ac_name:=CTL_NAME & cllSteps.Count + 1)
    SetupTextFont tbxStep, enMonStep
    With tbxStep
        .Top = ms_top
        .Left = 0
        .Visible = True
        .Height = 12
        .Width = Me.InsideWidth
         ms_top = AdjustToVgrid(.Top + .Height)
    End With
    VisualizeCtl tbxStep, VISLZE_BCKCLR_MON_STEPS_FRM
    Qenqueue cllSteps, tbxStep
    TimedDoEvents ErrSrc(PROC)
    
End Sub

Public Property Let MonitorProcess(ByVal s As String):          sMonitorProcess = s:            End Property
Public Property Let MonitorStepsDisplayed(ByVal l As Long):     lMonitorStepsDisplayed = l:     End Property

Public Sub MonitorFooter()
' ------------------------------------------------------------------------------
' Establishes a footer in the monitor window when none has yet been established
' and displays the provided text.
' ------------------------------------------------------------------------------
    Const PROC = "MonitorFooter"
    
    On Error GoTo eh
    Dim siTop As Single
    
    If cllSteps Is Nothing Then Me.MonitorInit
    
    '~~ Establsh monitor footer
    If TextMonitorFooter.Text <> vbNullString Then
        siTop = AdjustToVgrid(frmSteps.Top + frmSteps.Height) + 6
        If tbxFooter Is Nothing Then
            Set tbxFooter = AddControl(ac_ctl:=TextBox, ac_visible:=True, ac_name:="tbMonitorFooter")
            With tbxFooter
                .Left = 0
                .Height = 18
                .Width = Me.InsideWidth
            End With
            VisualizeCtl tbxFooter, VISLZE_BCKCLR_MSEC_TBX
        End If
        SetupTextFont tbxFooter, enMonFooter
        With tbxFooter
            .Top = AdjustToVgrid(frmSteps.Top + frmSteps.Height + 6)
            .Value = TextMonitorFooter.Text
        End With
        Me.Height = ContentHeight(tbxFooter.Parent) + 35
    End If

xt: Exit Sub

eh: If ErrMsg(ErrSrc(PROC)) = vbYes Then: Stop: Resume
End Sub

Public Sub MonitorInit()
' ------------------------------------------------------------------------------
' Setup the number of monitored steps to be displayed (mon_steps_displayed) - at
' first invisible - and the footer (mon_footer). The fMsg instance is identified
' through the title (mon_title).
' ------------------------------------------------------------------------------
    Const PROC = "MonitorInit"
    
    On Error GoTo eh
    Dim ctl                 As MSForms.Control
    Dim siTop               As Single
    Dim i                   As Long
    Dim siStepsHeightMax    As Single
    Dim siNetHeight         As Single           ' The height of the setup header and footer
        
    If Not bMonitorInitialized Then
        Set cllSteps = Nothing
        Set cllSteps = New Collection
        For Each ctl In Me.Controls
            ctl.Visible = False
        Next ctl
        siTop = 6
        
        With Me
            .Caption = sMonitorProcess
            .Width = .MsgWidthMin
                        
            '~~ Establish the number of visualized monitor steps in its dedicated frame
            Set frmSteps = AddControl(ac_ctl:=Frame, ac_name:="frMonitorSteps")
            With frmSteps
                .Top = siTop
                .Visible = True
                .Width = Me.InsideWidth
                .BorderColor = Me.BackColor
                .BorderStyle = fmBorderStyleSingle
                siTop = 0
                For i = 1 To lMonitorStepsDisplayed
                    MonitorEstablishStep siTop
                Next i
                .Height = ContentHeight(frmSteps, False)
                '~~ The maximum height for the steps frame is the max formheight minus the height of header and footer
                siNetHeight = ContentHeight(frmSteps.Parent) - frmSteps.Height
                siStepsHeightMax = Me.MsgHeightMax - siNetHeight
                NewHeight(frmSteps, False) = Min(siStepsHeightMax, .Height)
            End With
            VisualizeCtl frmSteps, VISLZE_BCKCLR_MON_STEPS_FRM
            NewHeight(frmSteps.Parent) = Min(.MsgHeightMax, ContentHeight(frmSteps.Parent))
            NewWidth(frmSteps) = Min(.MsgWidthMax, ContentWidth(frmSteps.Parent))
        End With
        bMonitorInitialized = True
    End If

xt: Exit Sub

eh: If ErrMsg(ErrSrc(PROC)) = vbYes Then: Stop: Resume
End Sub

Public Sub MonitorStep()
' ------------------------------------------------------------------------------
' Displays a monitored step. Note that the height of the steps frame (frmSteps)
' is already adjusted to the number of steps to be displayed. However, when one
' or another step's height is more than one line the height needs to be ajusted.
' ------------------------------------------------------------------------------
    Const PROC = "MonitorStep"
    
    On Error GoTo eh
    Dim siMaxWidth          As Single
    Dim tbx                 As MSForms.TextBox
    Dim i                   As Long
    Dim siTop               As Single
    Dim siNetHeight         As Single
    
    If cllSteps Is Nothing Then Me.MonitorInit
    
    siTop = 0
    If TextMonitorStep.Text <> vbNullString Then
        If lStepsDisplayed < lMonitorStepsDisplayed Then
            If lStepsDisplayed > 0 Then
                Set tbx = cllSteps(lStepsDisplayed)
                siTop = AdjustToVgrid(tbx.Top + tbx.Height)
            End If
            Set tbx = cllSteps(lStepsDisplayed + 1)
            SetupTextFont tbx, enMonStep
            tbx.Visible = True
            tbx.Top = siTop
            
            If TextMonitorStep.MonoSpaced Then
                AutoSizeTextBox as_tbx:=tbx _
                              , as_text:=TextMonitorStep.Text _
                              , as_width_limit:=0
            Else
                AutoSizeTextBox as_tbx:=tbx _
                              , as_width_limit:=Me.InsideWidth _
                              , as_text:=TextMonitorStep.Text
            End If
            MonitorStepsAdjustTopPosition
            NewWidth(frmSteps, False) = Min(Me.MsgWidthMax, ContentWidth(frmSteps, False)) ' applies a horizontal scroll-bar or adjust its width
            NewWidth(frmSteps.Parent) = ContentWidth(frmSteps.Parent)
            
            siNetHeight = Me.Height - (frmSteps.Height - frmSteps.Top)
            NewHeight(frmSteps, False, fmScrollActionBegin) = Min(MonitorHeightMaxSteps, ContentHeight(frmSteps, False))
            
            lStepsDisplayed = lStepsDisplayed + 1
            
            If Not tbxFooter Is Nothing Then
                tbxFooter.Top = AdjustToVgrid(frmSteps.Top + frmSteps.Height + 6)
                Me.Height = tbxFooter.Top + tbxFooter.Height + 35
            Else
                Me.Height = ContentHeight(frmSteps.Parent) + 35
            End If
        Else
            '~~ All steps are displayed each display of a new process step
            '~~ scrolls the displayed steps by dequeueing the top item and
            '~~ enqueueing as the new step
            siTop = 0
            Set tbx = Qdequeue(cllSteps)
            tbx.Value = vbNullString
            Qenqueue cllSteps, tbx
            
            For i = 1 To lMonitorStepsDisplayed
                Set tbx = cllSteps(i)
                tbx.Top = siTop
                siTop = AdjustToVgrid(tbx.Top + tbx.Height)
                siMaxWidth = Max(siMaxWidth, tbx.Width)
                TimedDoEvents ErrSrc(PROC)
            Next i
            
            If TextMonitorStep.MonoSpaced Then
                AutoSizeTextBox as_tbx:=tbx _
                              , as_text:=TextMonitorStep.Text _
                              , as_width_limit:=0
            Else
                AutoSizeTextBox as_tbx:=tbx _
                              , as_width_limit:=Me.InsideWidth _
                              , as_text:=TextMonitorStep.Text
            End If
            MonitorStepsAdjustTopPosition
            NewWidth(frmSteps, False) = Min(Me.MsgWidthMax, ContentWidth(frmSteps, False)) ' applies a horizontal scroll-bar or adjust its width
            NewWidth(frmSteps.Parent) = ContentWidth(frmSteps.Parent) + 20
            
            siNetHeight = Me.Height - (frmSteps.Height - frmSteps.Top)
            NewHeight(frmSteps, False, fmScrollActionEnd) = Min(MonitorHeightMaxSteps, ContentHeight(frmSteps, False)) + ScrollHscrollHeight(frmSteps)
        
            If Not tbxFooter Is Nothing Then
                tbxFooter.Top = AdjustToVgrid(frmSteps.Top + frmSteps.Height + 6)
                Me.Height = tbxFooter.Top + tbxFooter.Height + 35
            Else
                Me.Height = ContentHeight(frmSteps.Parent) + 35
            End If
        End If
    End If
        
    TimedDoEvents ErrSrc(PROC)
    NewWidth(frmSteps) = Min(Me.MsgWidthMax, ContentWidth(frmSteps.Parent) + 15)
    Me.Height = ContentHeight(frmSteps.Parent) + 35
    
xt: Exit Sub

eh: If ErrMsg(ErrSrc(PROC)) = vbYes Then: Stop: Resume
End Sub

Private Sub MonitorStepsAdjustTopPosition()
' ------------------------------------------------------------------------------
' - Adjusts each visible control's top position considering its current height.
' ------------------------------------------------------------------------------
    Const PROC = "AdjustTopPositions"
    
    On Error GoTo eh
    Dim siTop   As Single
    Dim ctl     As MSForms.Control
    Dim v       As Variant
    
    siTop = 0
    For Each v In cllSteps
        Set ctl = v
        ctl.Top = siTop
        siTop = AdjustToVgrid(ctl.Top + ctl.Height)
    Next v

xt: Exit Sub
    
eh: If ErrMsg(ErrSrc(PROC)) = vbYes Then: Stop: Resume
End Sub

Private Function MsectsArea() As MSForms.Frame
' ------------------------------------------------------------------------------
' Returns the Frame of the message area section, created if yet not existing.
' ------------------------------------------------------------------------------
    
    If frmMsectsArea Is Nothing Then
        Set frmMsectsArea = Me.AddControl(ac_ctl:=Frame, ac_visible:=True)
        With frmMsectsArea
            .Top = 0
            .Left = 0
            .Height = 50
            .Width = Me.InsideWidth
            .Visible = True
        End With
        VisualizeCtl frmMsectsArea, VISLZE_BCKCLR_AREA
    End If
    Set MsectsArea = frmMsectsArea

End Function

Private Function MsectsFrm(ByVal est_section As Long) As MSForms.Frame
' ------------------------------------------------------------------------------
' Returns the Frame of the message section (est_section), created if yet not
' existing.
' ------------------------------------------------------------------------------
    Dim frm         As MSForms.Frame
    Dim frmAbove    As MSForms.Frame
    
    If Not dctMsectsFrm.Exists(est_section) Then
        Set frm = AddControl(ac_ctl:=Frame, ac_visible:=True, ac_in:=MsectsArea, ac_name:="frMsgSection" & est_section)
        With frm
            .Top = 0
            If est_section > 1 Then
                Set frmAbove = MsectsFrm(est_section - 1)
                .Top = AdjustToVgrid(frmAbove.Top + frmAbove.Height)
            End If
            .Left = 0
            .Height = 18
            .Width = Me.InsideWidth
            .Visible = True
        End With
        VisualizeCtl frm, VISLZE_BCKCLR_MSEC_FRM
        dctMsectsFrm.Add est_section, frm
    End If
    Set MsectsFrm = dctMsectsFrm(est_section)

End Function

Private Function MsectsLbl(ByVal est_section As Long) As MSForms.Label
' ------------------------------------------------------------------------------
' Returns the Label of the message section (est_section), created in the
' corresponding MsectsFrm when not yet existing.
' ------------------------------------------------------------------------------
    Const PROC      As String = "MsectsLbl"
    Const NAME_LBL  As String = "laMsgSection[sect]Label"
    
    On Error GoTo eh
    Dim lbl As MSForms.Label
    
    If Not dctMsectsLbl.Exists(est_section) Then
        Set lbl = AddControl(ac_ctl:=Label _
                           , ac_visible:=True _
                           , ac_in:=MsectsFrm(est_section) _
                           , ac_name:=Replace(NAME_LBL, "[sect]", est_section) _
                            )
        With lbl
            .Top = 0
            .Left = 10
            .Height = 12
            .Width = Me.InsideWidth
            .Visible = False
        End With
        VisualizeCtl lbl, VISLZE_BCKCLR_MSEC_LBL
        dctMsectsLbl.Add est_section, lbl
    End If
    Set MsectsLbl = dctMsectsLbl(est_section)
   
xt: Exit Function
    
eh: If ErrMsg(ErrSrc(PROC)) = vbYes Then: Stop: Resume
End Function

Private Function MsectsTbx(ByVal est_section As Long) As MSForms.TextBox
' ------------------------------------------------------------------------------
' Returns the TextBox of the section (est_section), created in the Frame
' (est_in_frame) when not yet existing.
' ------------------------------------------------------------------------------
    Const PROC      As String = "MsectsTbx"
    Const NAME_TBX  As String = "tbMsgSection[sect]Text"
    
    On Error GoTo eh
    Dim tbx As MSForms.TextBox
    
    If Not dctMsectsTbx.Exists(est_section) Then
        Set tbx = AddControl(ac_ctl:=TextBox _
                           , ac_visible:=True _
                           , ac_in:=MsectsTbxFrm(est_section) _
                           , ac_name:=Replace(NAME_TBX, "[sect]", est_section) _
                            )
        With tbx
            .Top = 0
            .Left = 0
            .Height = 18
            .Width = Me.InsideWidth
            .Visible = True
        End With
        VisualizeCtl tbx, VISLZE_BCKCLR_MSEC_TBX
        dctMsectsTbx.Add est_section, tbx
    End If
    Set MsectsTbx = dctMsectsTbx(est_section)

xt: Exit Function
    
eh: If ErrMsg(ErrSrc(PROC)) = vbYes Then: Stop: Resume
End Function

Private Function MsectsTbxFrm(ByVal est_section As Long) As MSForms.Frame
' ------------------------------------------------------------------------------
' Returns the frame of the TextBox of the section (est_section), created in the
' corresponding MsectsFrm when not yet existing. The Frame's top
' position is 0 or, when there is a visible above Label underneath it.
' ------------------------------------------------------------------------------
    Const PROC          As String = "MsectsTbxFrm"
    Const NAME_TBX_FRM  As String = "frMsgSection[sect]Text"
    
    On Error GoTo eh
    Dim frm As MSForms.Frame
    Dim lbl As MSForms.Label
    
    If Not dctMsectsTbxFrm.Exists(est_section) Then
        Set frm = AddControl(ac_ctl:=Frame _
                           , ac_visible:=True _
                           , ac_in:=MsectsFrm(est_section) _
                           , ac_name:=Replace(NAME_TBX_FRM, "{sect]", est_section) _
                            )
        With frm
            .Top = 0
            If dctMsectsLbl.Exists(est_section) Then
                Set lbl = dctMsectsLbl(est_section)
                If lbl.Visible Then .Top = AdjustToVgrid(lbl.Top + lbl.Height)
            End If
            .Left = 0
            .Height = 50
            .Width = Me.InsideWidth
            .Visible = True
        End With
        VisualizeCtl frm, VISLZE_BCKCLR_MSEC_TBX_FRM
        dctMsectsTbxFrm.Add est_section, frm
    End If
    Set MsectsTbxFrm = dctMsectsTbxFrm(est_section)

xt: Exit Function
    
eh: If ErrMsg(ErrSrc(PROC)) = vbYes Then: Stop: Resume
End Function

Private Sub OpenClickedLabelItem(ByVal oc_section As Long)
    Dim sItem As String
    sItem = MsgLabel(oc_section).OpenWhenClicked
    mMsg.ShellRun sItem, WIN_NORMAL
End Sub

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

Private Sub ShowAtRange(ByVal sar_rng As Range)
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

    With Me
       .StartupPosition = 0
       .Left = PosLeft
       .Top = PosTop
    End With

End Sub

Public Sub PositionOnScreen(Optional ByVal pos_top_left As Variant = 3)
' ------------------------------------------------------------------------------
' Positions the form on the display, defaults to "Windows Default" and may be
' the following: - enManual (0)         = No initial setting specified
'                - enCenterOwner (1)    = Center on the item to which the
'                                           UserForm belongs
'                - enCenterScreen (2)   = Center on the whole screen.
'                - enWindowsDefault (3) = Position in upper-left corner of
'                                           screen (default)
'                - a range object specifying top and left
'                - a string in the form <top>;<left>
' ------------------------------------------------------------------------------
    Const PROC = "PositionOnScreen"
    
    On Error GoTo eh
    Dim pos_top     As Long
    Dim pos_left    As Long
    
    On Error Resume Next
    Select Case True
        Case TypeName(pos_top_left) = "Range"
            ShowAtRange pos_top_left
        Case TypeName(pos_top_left) = "String"
            If InStr(pos_top_left, ";") = 0 _
            Then Err.Raise AppErr(1), ErrSrc(PROC), "The provided argument is not a string with two values seperated with a comma!"
            If Not IsNumeric(Trim(Split(pos_top_left, ";")(0))) _
            Then Err.Raise AppErr(1), ErrSrc(PROC), "In the provided string argument the value preceeding the comma is not a numeric value!"
            If Not IsNumeric(Trim(Split(pos_top_left, ";")(1))) _
            Then Err.Raise AppErr(1), ErrSrc(PROC), "In the provided string argument the value following the comma is not a numeric value!"
            pos_top = CLng(Trim(Split(pos_top_left, ";")(0)))
            pos_left = CLng(Trim(Split(pos_top_left, ";")(1)))
            With Me
                .StartupPosition = enManual
                .Left = pos_left
                .Top = pos_top
            End With
        Case IsNumeric(pos_top_left)
            Me.StartupPosition = pos_top_left
    End Select
    
    '~~ First make sure the bottom right fits,
    '~~ then check if the top-left is still on the screen (which gets priority).
    With Me
        If ((.Left + .Width) > (VirtualScreenLeftPts + VirtualScreenWidthPts)) Then .Left = ((VirtualScreenLeftPts + VirtualScreenWidthPts) - .Width)
        If ((.Top + .Height) > (VirtualScreenTopPts + VirtualScreenHeightPts)) Then .Top = ((VirtualScreenTopPts + VirtualScreenHeightPts) - .Height)
        If (.Left < VirtualScreenLeftPts) Then .Left = VirtualScreenLeftPts
        If (.Top < VirtualScreenTopPts) Then .Top = VirtualScreenTopPts
    End With
    
xt: Exit Sub

eh: If ErrMsg(ErrSrc(PROC)) = vbYes Then: Stop: Resume
End Sub

Private Function NewDict(ByRef dct As Dictionary) As Dictionary
' ------------------------------------------------------------------------------
' Returns the Dictionary (dct), getting rid of an old.
' ------------------------------------------------------------------------------
    Set dct = Nothing
    Set dct = New Dictionary
    Set NewDict = dct
End Function

Private Sub ProvideSectionItems(ByVal i As Long)
' ------------------------------------------------------------------------------
' Provides the section specific controls, the section's message text string and
' the message label
' ------------------------------------------------------------------------------
    Const PROC = "ProvideSectionItems"
    
    On Error GoTo eh
    If i = 0 Then Stop
    Set MsectFrm = MsectsFrm(i)
    Set MsectLbl = MsectsLbl(i)
    Set MsectTbxFrm = MsectsTbxFrm(i)
    Set MsectTbx = MsectsTbx(i)
    MsgSectTxt = Text(enSectText, i)
    MsgSectLbl = MsgLabel(i)

xt: Exit Sub

eh: If ErrMsg(ErrSrc(PROC)) = vbYes Then: Stop: Resume
End Sub

Private Function Qdequeue(ByRef qu As Collection) As Variant
    Const PROC = "DeQueue"
    
    On Error GoTo eh
    If qu Is Nothing Then GoTo xt
    If QisEmpty(qu) Then GoTo xt
    On Error Resume Next
    Set Qdequeue = qu(1)
    If Err.Number <> 0 _
    Then Qdequeue = qu(1)
    qu.Remove 1

xt: Exit Function

eh: If ErrMsg(ErrSrc(PROC)) = vbYes Then: Stop: Resume
End Function

Private Sub Qenqueue(ByRef qu As Collection, ByVal qu_item As Variant)
    If qu Is Nothing Then Set qu = New Collection
    qu.Add qu_item
End Sub

Private Function QisEmpty(ByVal qu As Collection) As Boolean
    If Not qu Is Nothing _
    Then QisEmpty = qu.Count = 0 _
    Else QisEmpty = True
End Function

Private Function ScrollHscrollApplied(ByVal sa_frame_form As Variant) As Boolean
' ------------------------------------------------------------------------------
' Returns True when the control (sa_frame_form) has a horizontal scrollbar applied. When
' no control is provided it is the UserForm which is ment.
' ------------------------------------------------------------------------------
    If IsFrameOrForm(sa_frame_form) Then
        Select Case sa_frame_form.ScrollBars
            Case fmScrollBarsBoth, fmScrollBarsHorizontal: ScrollHscrollApplied = True
        End Select
    End If
End Function

Private Sub ScrollHscrollApply(ByRef sha_frame_form As Variant, _
                               ByVal sha_content_width, _
                      Optional ByVal sha_x_action As fmScrollAction = fmScrollActionBegin)
' ------------------------------------------------------------------------------
' - Apllies a horizontal scroll-bar when the width of the content of the frame
'   (sha_frame_form) is greater than its content
' - Adjust the scroll-bar's width by considering an already displayed vertical
'   scroll-bar
' ------------------------------------------------------------------------------
    Const PROC = "ScrollHscrollApply"
    
    On Error GoTo eh
    If Not IsFrameOrForm(sha_frame_form) _
    Then Err.Raise AppErr(1), ErrSrc(PROC), "The provided argument is neither a Frame nor a Form!"
        
    With sha_frame_form
        If Not ScrollHscrollApplied(sha_frame_form) Then
            Debug.Print "apply a horizontal scroll-bar for " & sha_frame_form.Name
            Select Case .ScrollBars
                Case fmScrollBarsBoth
                    .KeepScrollBarsVisible = fmScrollBarsBoth
                Case fmScrollBarsHorizontal
                    .KeepScrollBarsVisible = fmScrollBarsHorizontal
                Case fmScrollBarsVertical
                    .ScrollBars = fmScrollBarsBoth
                    .KeepScrollBarsVisible = fmScrollBarsBoth
                Case fmScrollBarsNone
                    .ScrollBars = fmScrollBarsHorizontal
                    .KeepScrollBarsVisible = fmScrollBarsHorizontal
            End Select
            If Not ScrollVscrollApplied(sha_frame_form) Then
                .Height = ContentHeight(sha_frame_form) + ScrollHscrollHeight(sha_frame_form)
            Else
                .Height = .Height + ScrollHscrollHeight(sha_frame_form)
            End If
        End If
       .ScrollWidth = sha_content_width
       .Scroll xAction:=sha_x_action
    End With

xt: Exit Sub
    
eh: If ErrMsg(ErrSrc(PROC)) = vbYes Then: Stop: Resume
End Sub

Private Function ScrollHscrollHeight(ByVal sh_frame_form As Variant) As Single
    If IsFrameOrForm(sh_frame_form) Then
        If ScrollHscrollApplied(sh_frame_form) Then ScrollHscrollHeight = SCROLL_H_HEIGHT
    End If
End Function

Private Sub ScrollHscrollRemove(ByRef shr_frame_form As Variant)
' ------------------------------------------------------------------------------
' Removes a vertical scroll-bar.
' ------------------------------------------------------------------------------
    Const PROC = "ScrollHscrollRemove"
    
    On Error GoTo eh
    If Not IsFrameOrForm(shr_frame_form) _
    Then Err.Raise AppErr(1), ErrSrc(PROC), "The provided argument is neither a Frame nor a Form!"
    
    With shr_frame_form
        If ScrollHscrollApplied(shr_frame_form) Then
            '~~ Establish the vertical scroll-bar, added to the horizontal scroll-bar when one is already applied
            Select Case .ScrollBars
                Case fmScrollBarsBoth, fmScrollBarsVertical
                    .KeepScrollBarsVisible = fmScrollBarsHorizontal
                    .ScrollBars = fmScrollBarsHorizontal
                Case fmScrollBarsHorizontal
                    .KeepScrollBarsVisible = fmScrollBarsNone
                    .ScrollBars = fmScrollBarsNone
            End Select
        End If
    End With
    
xt: Exit Sub
    
eh: If ErrMsg(ErrSrc(PROC)) = vbYes Then: Stop: Resume
End Sub

Private Function ScrollVscrollApplied(Optional ByVal sa_frame_form As Variant = Nothing) As Boolean
' ------------------------------------------------------------------------------
' Returns True when the control (sa_frame_form) has a vertical scrollbar applied. When no
' control is provided it is the UserForm which is ment.
' ------------------------------------------------------------------------------
    If IsFrameOrForm(sa_frame_form) Then
        Select Case sa_frame_form.ScrollBars
            Case fmScrollBarsBoth, fmScrollBarsVertical: ScrollVscrollApplied = True
        End Select
    End If
End Function

Private Sub ScrollVscrollApply(ByRef sva_frame_form As Variant, _
                               ByVal sva_content_height As Single, _
                      Optional ByVal sva_y_action As fmScrollAction = fmScrollActionBegin)
' ------------------------------------------------------------------------------
' - Apllies a vertical scroll-bar when the height of the content of the frame
'   (sva_frame_form) is greater than its content
' - Adjust the scroll-bar's height by considering an already displayed
'   horizontal scroll-bar
' ------------------------------------------------------------------------------
    Const PROC = "ScrollVscrollApply"
    
    On Error GoTo eh
    If Not IsFrameOrForm(sva_frame_form) _
    Then Err.Raise AppErr(1), ErrSrc(PROC), "The provided argument is neither a Frame nor a Form!"
    
    With sva_frame_form
        If Not ScrollVscrollApplied(sva_frame_form) Then
            '~~ Establish the vertical scroll-bar, added to the horizontal scroll-bar when one is already applied
            Select Case .ScrollBars
                Case fmScrollBarsBoth
                    .KeepScrollBarsVisible = fmScrollBarsBoth
                Case fmScrollBarsHorizontal
                    .ScrollBars = fmScrollBarsBoth
                    .KeepScrollBarsVisible = fmScrollBarsBoth
                Case fmScrollBarsVertical
                    .KeepScrollBarsVisible = fmScrollBarsVertical
                Case fmScrollBarsNone
                    .ScrollBars = fmScrollBarsVertical
                    .KeepScrollBarsVisible = fmScrollBarsVertical
            End Select
        End If
        .Scroll yAction:=sva_y_action
        .ScrollHeight = sva_content_height
        If Not ScrollHscrollApplied(sva_frame_form) Then
            .Width = ContentWidth(sva_frame_form) + ScrollVscrollWidth(sva_frame_form)
        End If
    End With
    Debug.Print "Height Frame   = " & sva_frame_form.Height
    Debug.Print "Height Scroll  = " & sva_content_height

xt: Exit Sub
    
eh: If ErrMsg(ErrSrc(PROC)) = vbYes Then: Stop: Resume
End Sub

Private Sub ScrollVscrollMsgSectionOrArea(ByVal exceeding_height As Single)
' ------------------------------------------------------------------------------
' Either because the message area occupies 60% or more of the total height or
' because both, the message area and the buttons area us about the same height,
' it - or only the section text occupying 65% or more - will be reduced by the
' exceeding height amount (exceeding_height) and will get a vertical scrollbar.
' ------------------------------------------------------------------------------
    Const PROC = "ScrollVscrollMsgSectionOrArea"
    
    On Error GoTo eh
    Dim v               As Variant
    Dim frm             As MSForms.Frame
    Dim bScrollApplied  As Boolean
    
    '~~ Find a/the message section text which occupies 65% or more of the message area's height,
    For Each v In dctMsectsFrm
        Set frm = dctMsectsFrm(v)
        If frm.Visible Then
            ProvideSectionItems v
            
            If MsectTbxFrm.Height >= frmMsectsArea.Height * 0.65 _
            Or ScrollVscrollApplied(MsectTbxFrm) Then
                ' ------------------------------------------------------------------------------
                ' There is a section which occupies 65% of the overall height or has already a
                ' vertical scrollbar applied. Assigning a new frame height applies a vertical
                ' scrollbar if none is applied yet or just adjusts the scrollbar's height to the
                ' frame's content height
                ' ------------------------------------------------------------------------------
                If MsectTbxFrm.Height - exceeding_height > 0 Then
                    If MsectTbxFrm.Height <> MsectTbxFrm.Height - exceeding_height Then
                        NewHeight(MsectTbxFrm) = MsectTbxFrm.Height - exceeding_height
                        AdjustedParentsWidthAndHeight MsectTbx
                        AdjustTopPositions
                        ProvideSectionItems v ' reset
                        bScrollApplied = True
                        Exit For
                    End If
                End If
            End If
        End If ' visible
    Next v
    
    If Not ScrollVscrollApplied(frmMsectsArea) And Not bScrollApplied And MsectsArea.Height <> ContentHeight(frmMsectsArea) - exceeding_height Then
        '~~ None of the message sections has a dominating height. Becaue the overall message area
        '~~ occupies >=60% of the height it is now reduced to fit the maximum message height
        '~~ thereby receiving a vertical scroll-bar
        NewHeight(frmMsectsArea) = ContentHeight(frmMsectsArea) - exceeding_height
        AdjustedParentsWidthAndHeight MsectTbx
        AdjustTopPositions
    End If

xt: Exit Sub

eh: If ErrMsg(ErrSrc(PROC)) = vbYes Then: Stop: Resume
End Sub

Private Sub ScrollVscrollRemove(ByRef sr_frame_form As Variant)
' ------------------------------------------------------------------------------
' Removes a vertical scroll-bar.
' ------------------------------------------------------------------------------
    Const PROC = "ScrollVscrollRemove"
    
    On Error GoTo eh
    If Not IsFrameOrForm(sr_frame_form) _
    Then Err.Raise AppErr(1), ErrSrc(PROC), "The provided argument is neither a Frame nor a Form!"
    
    With sr_frame_form
        If Not ScrollVscrollApplied(sr_frame_form) Then
            '~~ Establish the vertical scroll-bar, added to the horizontal scroll-bar when one is already applied
            Select Case .ScrollBars
                Case fmScrollBarsBoth, fmScrollBarsHorizontal
                    .KeepScrollBarsVisible = fmScrollBarsHorizontal
                    .ScrollBars = fmScrollBarsHorizontal
                Case fmScrollBarsVertical
                    .KeepScrollBarsVisible = fmScrollBarsNone
                    .ScrollBars = fmScrollBarsNone
            End Select
        End If
    End With
    
xt: Exit Sub
    
eh: If ErrMsg(ErrSrc(PROC)) = vbYes Then: Stop: Resume
End Sub

Private Sub ScrollVscrollWhereApplicable()
' ------------------------------------------------------------------------------
' Reduce the height of the message area and or the height of the buttons area to
' have the message form not exceeds the specified maximum height. The area which
' uses 60% or more of the overall height is the one being reduced. Otherwise
' both are reduced proportionally.
' When one of the message sections within the to be reduced message area
' occupies 80% or more of the overall message area height only this section
' is reduced and gets a verticall scrollbar.
' The reduced frames are returned (frame_msg, frame_bttns).
' ------------------------------------------------------------------------------
    Const PROC = "ScrollVscrollWhereApplicable"
    
    On Error GoTo eh
    Dim TotalExceedingHeight    As Single
    
    '~~ When the message form's height exceeds the specified maximum height
    If Me.Height > siMsgHeightMax Then
        With Me
            TotalExceedingHeight = .Height - siMsgHeightMax
            If TotalExceedingHeight < 20 Then GoTo xt ' not worth any intervention
            .Height = siMsgHeightMax     '~~ Reduce the height to the max height specified
            
            If PrcntgHeightMsgArea >= 0.6 Then
                '~~ Either the message area as a whole or the dominating message section - if ther is any -
                '~~ will be height reduced and applied with a vertical scroll bar
                ScrollVscrollMsgSectionOrArea TotalExceedingHeight
            ElseIf PrcntgHeightfrmBttnsArea >= 0.6 Then
                '~~ Only the buttons area will be reduced and applied with a vertical scrollbar.
'                Debug.Print frmBttnsArea.Height - TotalExceedingHeight
                NewHeight(frmBttnsArea) = frmBttnsArea.Height - TotalExceedingHeight
            Else
                '~~ Both, the message area and the buttons area will be
                '~~ height reduced proportionally and applied with a vertical scrollbar
'                Debug.Print frmMsectsArea.Height - (TotalExceedingHeight * PrcntgHeightMsgArea)
                NewHeight(frmMsectsArea) = frmMsectsArea.Height - (TotalExceedingHeight * PrcntgHeightMsgArea)
'                Debug.Print frmBttnsArea.Height - (TotalExceedingHeight * PrcntgHeightfrmBttnsArea)
                NewHeight(frmBttnsArea) = frmBttnsArea.Height - (TotalExceedingHeight * PrcntgHeightfrmBttnsArea)
            End If
        End With
    End If ' height exceeds specified maximum
   
xt: Exit Sub
    
eh: If ErrMsg(ErrSrc(PROC)) = vbYes Then: Stop: Resume
End Sub

Private Function ScrollVscrollWidth(ByVal sw_frame_form As Variant) As Single
    If IsFrameOrForm(sw_frame_form) Then
        If ScrollVscrollApplied(sw_frame_form) Then ScrollVscrollWidth = SCROLL_V_WIDTH
    End If
End Function

Public Sub Setup()
    Const PROC = "Setup"
    
    On Error GoTo eh
    IndicateFrameCaptionsSetup bIndicateFrameCaptions ' may be True for test purpose
    
    '~~ Start the setup as if there wouldn't be any message - which might be the case
    Me.StartupPosition = 2
    Me.Height = 200                             ' just to start with - specifically for test purpose
    Me.Width = siMsgWidthMin
    
'    PositionOnScreen pos_top_left:=True  ' in case of test best pos to start with
    frmMsectsArea.Visible = False
    frmBttnsArea.Top = VSPACE_AREAS
    
    '~~ ----------------------------------------------------------------------------------------
    '~~ The  p r i m a r y  setup of the title, the message sections and the reply buttons
    '~~ returns their individual widths which determines the minimum required message form width
    '~~ This setup ends width the final message form width and all elements adjusted to it.
    '~~ ----------------------------------------------------------------------------------------
    
    '~~ Setup of the title, the first element which potentially effects the final message width
    If Not bDoneTitle _
    Then Setup1_Title setup_title:=sMsgTitle _
                    , setup_width_min:=siMsgWidthMin _
                    , setup_width_max:=siMsgWidthMax
    
    '~~ Setup of any monospaced message sections, the second element which potentially effects the final message width.
    '~~ In case the section width exceeds the maximum width specified a horizontal scrollbar is applied.
    Setup2_MsgSectsMonoSpaced
    If iSectionsMonoSpaced > 0 Then VisualizeSetupStep "Step2 Message Section(s) Monospaced"

    '~~ Setup the reply buttons. This is the third element which may effect the final message's width.
    '~~ In case the widest buttons row exceeds the maximum width specified for the message
    '~~ a horizontal scrollbar is applied.
    If ButtonsProvided Then
        Setup3_Bttns
        SizeAndPosition2Bttns1
        SizeAndPosition2Bttns2Rows
        SizeAndPosition2Bttns3Frame
        AdjustTopPositions
        SizeAndPosition2Bttns4Area
        VisualizeSetupStep "Step3 Buttons"
    End If
    
    ' -----------------------------------------------------------------------------------------------
    ' At this point the form has reached its final width (all proportionally spaced message sections
    ' are adjusted to it). However, the message height is only final in case there are just buttons
    ' but no message. The setup of proportional spaced message sections determines the final message
    ' height. When it exeeds the maximum height specified one or two vertical scrollbars are applied.
    ' -----------------------------------------------------------------------------------------------
    Setup4_MsgSectsPropSpaced
    AdjustTopPositions
                
    ' -----------------------------------------------------------------------------------------------
    ' When the message form height exceeds the specified or the default message height the height of
    ' the message area and or the buttons area is reduced and a vertical is applied.
    ' When both areas are about the same height (neither is taller the than 60% of the total heigth)
    ' both will get a vertical scrollbar, else only the one which uses 60% or more of the height.
    ' -----------------------------------------------------------------------------------------------
    ScrollVscrollWhereApplicable
    
    '~~ Final form width adjustment
    '~~ When the message area or the buttons area has a vertical scrollbar applied
    '~~ the scrollbar may not be visible when the width as a result exeeds the specified
    '~~ message form width. In order not to interfere again with the width of all content
    '~~ the message form width is extended (over the specified maximum) in order to have
    '~~ the vertical scrollbar visible
    AdjustTopPositions
    
    If MsectsArea.Visible Then
        Me.Width = Max(ContentWidth(BttnsArea.Parent), ContentWidth(MsectsArea.Parent)) + 15
    End If
    Me.Height = Max(ContentHeight(BttnsArea.Parent), ContentHeight(MsectsArea.Parent)) + 35
    
    PositionOnScreen
    bSetUpDone = True ' To indicate for the Activate event that the setup had already be done beforehand
    
    VisualizeSetupStep "Setup finalized"

xt: Exit Sub

eh: If ErrMsg(ErrSrc(PROC)) = vbYes Then: Stop: Resume
End Sub

Private Sub Setup1_Title(ByVal setup_title As String, _
                         ByVal setup_width_min As Single, _
                         ByVal setup_width_max As Single)
' ------------------------------------------------------------------------------
' Setup the message form for the provided title (setup_title) optimized with the
' provided minimum width (setup_width_min) and the provided maximum width
' (setup_width_max) by using a certain factor (setup_factor) for the calculation
' of the width required to display an untruncated title - as long as the maximum
' widht is not exeeded.
' The correction of the template length label is a function (percentage) of the
' lenght.
' ------------------------------------------------------------------------------
    Const PROC = "Setup1_Title"
    
    On Error GoTo eh
    Dim Correction  As Single
    Dim siWidth     As Single
    
    With Me
        .Width = setup_width_min
        '~~ The extra title label is only used to adjust the form width and remains hidden
        With .laMsgTitle
            With .Font
                .Bold = False
                .Name = Me.Font.Name
                .Size = 8    ' Value which comes to a length close to the length required
            End With
            .Caption = vbNullString
            .AutoSize = True
            .Caption = " " & setup_title    ' some left margin
        End With
        .Caption = setup_title
        Correction = .laMsgTitle.Width * 0.11 ' The correction is a percentage of the length of the title template Label control
        siWidth = .laMsgTitle.Width + 44 + Correction
        .Width = Min(setup_width_max, siWidth)
        .Width = Max(.Width, setup_width_min)
        TitleWidth = .Width
    End With
    bDoneTitle = True
    
xt: Exit Sub
    
eh: If ErrMsg(ErrSrc(PROC)) = vbYes Then: Stop: Resume
End Sub

Private Sub Setup2_MsgSectsMonoSpaced()
' --------------------------------------------------------------------------------------
' Setup of all sections for which a text is provided indicated mono-spaced.
' Note: The number of message sections is only determined by the number of elements in
'       MsgText.
' --------------------------------------------------------------------------------------
    Const PROC = "Setup2_MsgSectsMonoSpaced"
    
    On Error GoTo eh
    Dim i           As Long
    
    For i = 1 To UBound(TextSection.Section)
        With Me.Text(enSectText, i)
            If .MonoSpaced And .Text <> vbNullString Then
                ProvideSectionItems i
                SetupMsgSect
                iSectionsMonoSpaced = iSectionsMonoSpaced + 1
            End If
        End With
    Next i
    bDoneMonoSpacedSects = True

xt: Exit Sub
    
eh: If ErrMsg(ErrSrc(PROC)) = vbYes Then: Stop: Resume
End Sub

Private Sub Setup3_Bttn(ByVal sb_row As Long, _
                        ByVal sb_button As Long, _
                        ByVal sb_caption As String, _
                        ByVal sb_ret_value As Variant)
' -------------------------------------------------------------------------------
' Setup an applied reply button's (sb_row, sb_button) visibility and caption,
' calculate the maximum width and height, keep a record of the setup
' reply sb_index's return value.
' -------------------------------------------------------------------------------
    Const PROC = "Setup3_Bttn"
    
    On Error GoTo eh
    Dim cmb As MSForms.CommandButton
    
    If sb_row = 0 Then sb_row = 1
    Set cmb = dctBttns(sb_row & "-" & sb_button)
    
    With cmb
        .AutoSize = True
        .WordWrap = False ' the longest line determines the sb_index's width
        .Caption = sb_caption
        .AutoSize = False
        .Height = .Height + 1 ' safety margin to ensure proper multilin caption display
        siMaxButtonHeight = Max(siMaxButtonHeight, .Height)
        siMaxButtonWidth = Max(siMaxButtonWidth, .Width, DFLT_BTTN_MIN_WIDTH)
    End With
    AppliedBttns.Add cmb, sb_row
    AppliedButtonRetVal(cmb) = sb_ret_value ' keep record of the setup sb_index's reply value
    cmb.Visible = True
    BttnsRowFrm(sb_row).Visible = True
    
xt: Exit Sub
    
eh: If ErrMsg(ErrSrc(PROC)) = vbYes Then: Stop: Resume
End Sub

Private Sub Setup3_Bttns()
' -------------------------------------------------------------------------------
' Setup the reply buttons based on the comma delimited string of button captions
' and row breaks indicated by a vbLf, vbCr, or vbCrLf.
' ---------------------------------------------------------------------
    Const PROC = "Setup3_Bttns"
    
    On Error GoTo eh
    Dim v           As Variant
    Dim BttnRow     As MSForms.Frame
    Dim Bttn        As MSForms.CommandButton

    If cllMsgBttns.Count = 0 Then GoTo xt
    frmBttnsArea.Visible = True
    frmBttnsFrm.Visible = True

    lSetupRows = 1
    lSetupRowButtons = 0
    Set BttnRow = dctBttnsRow(1)
    Set Bttn = dctBttns(1 & "-" & 1)
    
    Me.Height = 100 ' just to start with
    frmBttnsArea.Top = VSPACE_AREAS
    BttnsFrm.Top = frmBttnsArea.Top
    BttnRow.Top = BttnsFrm.Top
    Bttn.Top = BttnRow.Top
    Bttn.Width = DFLT_BTTN_MIN_WIDTH
    
    For Each v In cllMsgBttns
        If IsNumeric(v) Then v = mMsg.BttnsArgs(v)
        Select Case v
            Case vbOKOnly, vbOKCancel, vbYesNo, vbRetryCancel, vbYesNoCancel, vbAbortRetryIgnore, vbYesNo, vbResumeOk
                Setup3_BttnsFromValue v
            Case Else
                If v <> vbNullString Then
                    If v = vbLf Or v = vbCr Or v = vbCrLf Then
                        '~~ prepare for the next row
                        If lSetupRows <= 7 Then ' ignore exceeding rows
                            BttnsRowFrm(lSetupRows).Visible = True
                            lSetupRows = lSetupRows + 1
                            lSetupRowButtons = 0
                        Else
                            MsgBox "Setup of button row " & lSetupRows & " ignored! The maximum applicable rows is 7."
                        End If
                    Else
                        lSetupRowButtons = lSetupRowButtons + 1
                        If lSetupRowButtons <= 7 And lSetupRows <= 7 Then
                            Setup3_Bttn sb_row:=lSetupRows, sb_button:=lSetupRowButtons, sb_caption:=v, sb_ret_value:=v
                        Else
                            MsgBox "The setup of button " & lSetupRowButtons & " in row " & lSetupRows & " is ignored! The maximum applicable buttons per row is 7 " & _
                                   "and the maximum rows is 7 !"
                        End If
                    End If
                End If
        End Select
    Next v
    If lSetupRows <= 7 Then
        BttnsRowFrm(lSetupRows).Visible = True
    End If
    
xt: Exit Sub

eh: If ErrMsg(ErrSrc(PROC)) = vbYes Then: Stop: Resume
End Sub

Private Sub Setup3_BttnsFromValue(ByVal lButtons As Long)
' -------------------------------------------------------------------------------
' Setup a row of standard VB MsgBox reply command buttons
' -------------------------------------------------------------------------------
    Const PROC = "Setup3_BttnsFromValue"
    
    On Error GoTo eh
    Dim ResumeErrorLine As String: ResumeErrorLine = "Resume" & vbLf & "Error Line"
    Dim PassOn          As String: PassOn = "Pass on Error to" & vbLf & "Entry Procedure"
    
    Select Case lButtons
        Case vbOKOnly
            lSetupRowButtons = lSetupRowButtons + 1
            Setup3_Bttn sb_row:=lSetupRows, sb_button:=lSetupRowButtons, sb_caption:="Ok", sb_ret_value:=vbOK
        Case vbOKCancel
            lSetupRowButtons = lSetupRowButtons + 1
            Setup3_Bttn sb_row:=lSetupRows, sb_button:=lSetupRowButtons, sb_caption:="Ok", sb_ret_value:=vbOK
            lSetupRowButtons = lSetupRowButtons + 1
            Setup3_Bttn sb_row:=lSetupRows, sb_button:=lSetupRowButtons, sb_caption:="Cancel", sb_ret_value:=vbCancel
        Case vbYesNo
            lSetupRowButtons = lSetupRowButtons + 1
            Setup3_Bttn sb_row:=lSetupRows, sb_button:=lSetupRowButtons, sb_caption:="Yes", sb_ret_value:=vbYes
            lSetupRowButtons = lSetupRowButtons + 1
            Setup3_Bttn sb_row:=lSetupRows, sb_button:=lSetupRowButtons, sb_caption:="No", sb_ret_value:=vbNo
        Case vbRetryCancel
            lSetupRowButtons = lSetupRowButtons + 1
            Setup3_Bttn sb_row:=lSetupRows, sb_button:=lSetupRowButtons, sb_caption:="Retry", sb_ret_value:=vbRetry
            lSetupRowButtons = lSetupRowButtons + 1
            Setup3_Bttn sb_row:=lSetupRows, sb_button:=lSetupRowButtons, sb_caption:="Cancel", sb_ret_value:=vbCancel
        Case vbResumeOk
            lSetupRowButtons = lSetupRowButtons + 1
            Setup3_Bttn sb_row:=lSetupRows, sb_button:=lSetupRowButtons, sb_caption:=ResumeErrorLine, sb_ret_value:=vbResume
            lSetupRowButtons = lSetupRowButtons + 1
            Setup3_Bttn sb_row:=lSetupRows, sb_button:=lSetupRowButtons, sb_caption:="Ok", sb_ret_value:=vbOK
        Case vbYesNoCancel
            lSetupRowButtons = lSetupRowButtons + 1
            Setup3_Bttn sb_row:=lSetupRows, sb_button:=lSetupRowButtons, sb_caption:="Yes", sb_ret_value:=vbYes
            lSetupRowButtons = lSetupRowButtons + 1
            Setup3_Bttn sb_row:=lSetupRows, sb_button:=lSetupRowButtons, sb_caption:="No", sb_ret_value:=vbNo
            lSetupRowButtons = lSetupRowButtons + 1
            Setup3_Bttn sb_row:=lSetupRows, sb_button:=lSetupRowButtons, sb_caption:="Cancel", sb_ret_value:=vbCancel
        Case vbAbortRetryIgnore
            lSetupRowButtons = lSetupRowButtons + 1
            Setup3_Bttn sb_row:=lSetupRows, sb_button:=lSetupRowButtons, sb_caption:="Abort", sb_ret_value:=vbAbort
            lSetupRowButtons = lSetupRowButtons + 1
            Setup3_Bttn sb_row:=lSetupRows, sb_button:=lSetupRowButtons, sb_caption:="Retry", sb_ret_value:=vbRetry
            lSetupRowButtons = lSetupRowButtons + 1
            Setup3_Bttn sb_row:=lSetupRows, sb_button:=lSetupRowButtons, sb_caption:="Ignore", sb_ret_value:=vbIgnore
        Case vbResumeOk
            lSetupRowButtons = lSetupRowButtons + 1
            Setup3_Bttn sb_row:=lSetupRows, sb_button:=lSetupRowButtons, sb_caption:="Resume" & vbLf & "Error Line", sb_ret_value:=vbResume
            lSetupRowButtons = lSetupRowButtons + 1
            Setup3_Bttn sb_row:=lSetupRows, sb_button:=lSetupRowButtons, sb_caption:="Ok", sb_ret_value:=vbOK
    
        Case Else
            MsgBox "The value provided for the ""buttons"" argument is not a known VB MsgBox value"
    End Select
    If lSetupRows <> 0 Then
        BttnsRowFrm(lSetupRows).Visible = True
        BttnsFrm.Visible = True
    End If
xt: Exit Sub

eh: If ErrMsg(ErrSrc(PROC)) = vbYes Then: Stop: Resume
End Sub

Private Sub Setup4_MsgSectsPropSpaced()
' -------------------------------------------------------------------------------
' Loop through all provided message sections for which a text is provided and is
' not Monospaced and setup the section.
' Note: The number of message sections is only determined by the number of elements in
'       MsgText.
' -------------------------------------------------------------------------------
    Const PROC = "Setup4_MsgSectsPropSpaced"
    
    On Error GoTo eh
    Dim i As Long

    For i = i To UBound(TextSection.Section)
        With Me.Text(enSectText, i)
            If .Text <> vbNullString And Not .MonoSpaced Then
                ProvideSectionItems i
                SetupMsgSect
                iSectionsPropSpaced = iSectionsPropSpaced + 1
            End If
        End With
    Next i
    bDonePropSpacedSects = True
    bDoneMsgArea = True
    
xt: Exit Sub
    
eh: If ErrMsg(ErrSrc(PROC)) = vbYes Then: Stop: Resume
End Sub

Private Sub SetupMsgSect()
' -------------------------------------------------------------------------------
' Setup a message section with a label provided one is specified.
' -------------------------------------------------------------------------------
    Const PROC = "SetupMsgSect"
    
    On Error GoTo eh
                
    MsectFrm.Width = frmMsectsArea.Width
    MsectLbl.Width = MsectFrm.Width
    MsectTbxFrm.Width = MsectFrm.Width
    MsectTbx.Width = MsectFrm.Width
        
    If MsgSectTxt.Text <> vbNullString Then
    
        frmMsectsArea.Visible = True
        MsectFrm.Visible = True
        MsectTbxFrm.Visible = True
        MsectTbx.Visible = True
                
        If MsgSectLbl.Text <> vbNullString Then
            MsectLbl.Visible = True
'            Debug.Print MsectLbl.Name
            With MsectLbl
                .Left = 10
                .Width = Me.InsideWidth - (siHmarginFrames * 2)
                .Caption = MsgSectLbl.Text
                With .Font
                    If MsgSectLbl.MonoSpaced Then
                        If MsgSectLbl.FontName <> vbNullString Then .Name = MsgSectLbl.FontName Else .Name = DFLT_LBL_MONOSPACED_FONT_NAME
                        If MsgSectLbl.FontSize <> 0 Then .Size = MsgSectLbl.FontSize Else .Size = DFLT_LBL_MONOSPACED_FONT_SIZE
                    Else
                        If MsgSectLbl.FontName <> vbNullString Then .Name = MsgSectLbl.FontName Else .Name = DFLT_LBL_PROPSPACED_FONT_NAME
                        If MsgSectLbl.FontSize <> 0 Then .Size = MsgSectLbl.FontSize Else .Size = DFLT_LBL_PROPSPACED_FONT_SIZE
                    End If
                    If MsgSectLbl.FontItalic Then .Italic = True
                    If MsgSectLbl.FontBold Then .Bold = True
                    If MsgSectLbl.FontUnderline Then .Underline = True
                End With
                If MsgSectLbl.FontColor <> 0 Then .ForeColor = MsgSectLbl.FontColor Else .ForeColor = rgbBlack
            End With
            MsectTbxFrm.Top = MsectLbl.Top + MsectLbl.Height
            MsectLbl.Visible = True
        Else
            MsectTbxFrm.Top = 0
        End If
        
        If MsgSectTxt.MonoSpaced Then
            SetupMsgSectMonoSpaced  ' returns the maximum width required for monospaced section
        Else ' proportional spaced
            SetupMsgSectPropSpaced
        End If
        MsectTbx.SelStart = 0
        
    End If
    
xt: Exit Sub

eh: If ErrMsg(ErrSrc(PROC)) = vbYes Then: Stop: Resume
End Sub

Private Sub SetupMsgSectMonoSpaced(Optional ByVal msg_append As Boolean = False, _
                                   Optional ByVal msg_append_margin As String = vbNullString, _
                                   Optional ByVal msg_text As String = vbNullString)
' ------------------------------------------------------------------------------
' Setup the current monospaced message section. When a text is explicitely
' provided (msg_text) the section is setup with this one, else with the MsgText
' content. When an explicit text is provided the text either replaces the text,
' which the default or the text is appended when (msg_append = True).
' Note 1: All top and height adjustments - except the one for the text box
'         itself are finally done by SizeAndPosition services when all
'         elements had been set up.
' Note 2: The optional arguments (msg_append) and (msg_text) are used with the
'         Monitor service which ma replace or add the provided text
' ------------------------------------------------------------------------------
Const PROC = "SetupMsgSectMonoSpaced"
    
    On Error GoTo eh
    Dim MaxWidthAreaFrame   As Single
    Dim MaxWidthSectFrame   As Single
    Dim MaxWidthTextFrame   As Single
    
    MaxWidthAreaFrame = FormWidthMaxUsable - 4
    MaxWidthSectFrame = MaxWidthAreaFrame
    MaxWidthTextFrame = MaxWidthSectFrame
    
    '~~ Keep record of the controls which had been applied
    frmMsectsArea.Visible = True
    MsectFrm.Visible = True
    MsectTbxFrm.Visible = True: MonoSpaced(MsectTbxFrm) = True
    MsectTbx.Visible = True:    MonoSpaced(MsectTbx) = True:  MonoSpacedTbx(MsectTbx) = True
    
    If msg_text <> vbNullString Then MsgSectTxt.Text = msg_text
  
    With MsectTbx
        With .Font
            If MsgSectTxt.FontName <> vbNullString Then .Name = MsgSectTxt.FontName Else .Name = DFLT_LBL_MONOSPACED_FONT_NAME
            If MsgSectTxt.FontSize <> 0 Then .Size = MsgSectTxt.FontSize Else .Size = DFLT_LBL_MONOSPACED_FONT_SIZE
            If .Bold <> MsgSectTxt.FontBold Then .Bold = MsgSectTxt.FontBold
            If .Italic <> MsgSectTxt.FontItalic Then .Italic = MsgSectTxt.FontItalic
            If .Underline <> MsgSectTxt.FontUnderline Then .Underline = MsgSectTxt.FontUnderline
        End With
        If .ForeColor <> MsgSectTxt.FontColor And MsgSectTxt.FontColor <> 0 Then .ForeColor = MsgSectTxt.FontColor
    End With
    
    AutoSizeTextBox as_tbx:=MsectTbx _
                  , as_text:=MsgSectTxt.Text _
                  , as_width_limit:=0 _
                  , as_append:=msg_append _
                  , as_append_margin:=msg_append_margin
    
    With MsectTbx
        .SelStart = 0
        .Left = siHmarginFrames
        MsectTbxFrm.Left = siHmarginFrames
        MsectTbxFrm.Height = .Top + .Height
    End With ' MsectTbx
        
    '~~ The width may expand or shrink depending on the change of the displayed text
    '~~ However, it cannot expand beyond the maximum width calculated for the text frame
    NewWidth(MsectTbxFrm) = Min(MaxWidthTextFrame, MsectTbx.Width)
    MsectFrm.Width = Min(MaxWidthSectFrame, MsectTbxFrm.Width)
    frmMsectsArea.Width = Min(MaxWidthSectFrame, MsectFrm.Width)
    FormWidth = frmMsectsArea.Width
    AdjustTopPositions
'    AdjustedParentsWidthAndHeight MsectTbx
    
xt: Exit Sub
    
eh: If ErrMsg(ErrSrc(PROC)) = vbYes Then: Stop: Resume
End Sub

Private Sub SetupMsgSectPropSpaced(Optional ByVal msg_append As Boolean = False, _
                                   Optional ByVal msg_append_marging As String = vbNullString, _
                                   Optional ByVal msg_text As String = vbNullString)
' ------------------------------------------------------------------------------
' Setup the current proportional spaced section. When a text is explicitely
' provided (msg_text) the section is setup with this one, else with the property
' MsgText content. When an explicit text is provided the text either replaces
' the text, which the default or the text is appended when (msg_appen = True).
' Note 1: When this proportional spaced section is setup the message width is
'         regarded final. However, top and height adjustments - except the one
'         for the text box itself are finally done by SizeAndPosition
'         services when all elements had been set up.
' Note 2: The optional arguments (msg_append) and (msg_text) are used with the
'         Monitor service which ma replace or add the provided text
' ------------------------------------------------------------------------------
    Const PROC = "SetupMsgSectPropSpaced"
    
    On Error GoTo eh
    
    frmMsectsArea.Visible = True
    MsectFrm.Visible = True
    MsectTbxFrm.Visible = True
    MsectTbx.Visible = True

    '~~ For proportional spaced message sections the width is determined by the Message area's width
    frmMsectsArea.Width = Me.InsideWidth
    MsectFrm.Width = frmMsectsArea.Width
    MsectTbxFrm.Width = MsectFrm.Width - 5
'    Debug.Print "MsectTbxFrm.Width = " & MsectTbxFrm.Width
    
    frmBttnsArea.Top = frmMsectsArea.Top + frmMsectsArea.Height + 20
    Me.Height = frmBttnsArea.Top + frmBttnsArea.Height + 20
    
    If msg_text <> vbNullString Then MsgSectTxt.Text = msg_text
    
    With MsectTbx
        With .Font
            If MsgSectTxt.FontName <> vbNullString Then .Name = MsgSectTxt.FontName Else .Name = DFLT_LBL_PROPSPACED_FONT_NAME
            If MsgSectTxt.FontSize <> 0 Then .Size = MsgSectTxt.FontSize Else .Size = DFLT_LBL_PROPSPACED_FONT_SIZE
            If .Bold <> MsgSectTxt.FontBold Then .Bold = MsgSectTxt.FontBold
            If .Italic <> MsgSectTxt.FontItalic Then .Italic = MsgSectTxt.FontItalic
            If .Underline <> MsgSectTxt.FontUnderline Then .Underline = MsgSectTxt.FontUnderline
        End With
        If .ForeColor <> MsgSectTxt.FontColor And MsgSectTxt.FontColor <> 0 Then .ForeColor = MsgSectTxt.FontColor
    End With
    
    AutoSizeTextBox as_tbx:=MsectTbx _
                  , as_width_limit:=MsectTbxFrm.Width _
                  , as_text:=MsgSectTxt.Text _
                  , as_append:=msg_append _
                  , as_append_margin:=msg_append_marging
    
    With MsectTbx
        .SelStart = 0
        .Left = HSPACE_LEFT
        TimedDoEvents ErrSrc(PROC)    ' to properly h-align the text
    End With
    
    MsectTbxFrm.Height = MsectTbx.Top + MsectTbx.Height
    MsectFrm.Height = MsectTbxFrm.Top + MsectTbxFrm.Height
    frmMsectsArea.Height = ContentHeight(frmMsectsArea)
    frmBttnsArea.Top = frmMsectsArea.Top + frmMsectsArea.Height + 20
    Me.Height = frmBttnsArea.Top + frmBttnsArea.Height + 20
    
xt: Exit Sub
    
eh: If ErrMsg(ErrSrc(PROC)) = vbYes Then: Stop: Resume
End Sub

Private Sub SetupTextFont(ByVal ctl As MSForms.Control, _
                          ByVal kind_of_text As KindOfText)
' ------------------------------------------------------------------------------
' Setup the font properties for a Label or TextBox (ctl) according to the
' corresponding TypeMsgText type (kind_of_text).
' ------------------------------------------------------------------------------

    Dim txt As TypeMsgText
    txt = Me.Text(kind_of_text)
    
    With ctl.Font
        If .Bold <> txt.FontBold Then .Bold = txt.FontBold
        If .Italic <> txt.FontItalic Then .Italic = txt.FontItalic
        If .Underline <> txt.FontUnderline Then .Underline = txt.FontUnderline
        If txt.MonoSpaced Then
            .Name = DFLT_TXT_MONOSPACED_FONT_NAME
            If txt.FontSize = 0 _
            Then .Size = DFLT_TXT_MONOSPACED_FONT_SIZE _
            Else .Size = txt.FontSize
        Else
            If txt.FontName = vbNullString _
            Then .Name = DFLT_TXT_PROPSPACED_FONT_NAME _
            Else .Name = txt.FontName
            If txt.FontSize = 0 _
            Then .Size = DFLT_TXT_PROPSPACED_FONT_SIZE _
            Else .Size = txt.FontSize
        End If
    End With
    ctl.ForeColor = txt.FontColor
    If bVisualizeForTest Then ctl.BackColor = VISLZE_BCKCLR_MSEC_TBX
End Sub

Private Sub SizeAndPosition2Bttns1()
' ------------------------------------------------------------------------------
' Unify all applied/visible button's size by assigning the maximum width and
' height provided with their setup, and adjust their resulting left position.
' ------------------------------------------------------------------------------
    Const PROC = "SizeAndPosition2Bttns1"
    
    On Error GoTo eh
    Dim siLeft          As Single
    Dim frmRow          As MSForms.Frame    ' Frame for the buttons in a row
    Dim v               As Variant
    Dim lRow            As Long
    Dim lButton         As Long
    Dim HeightfrmBttnsArea As Single
    Dim cmb             As MSForms.CommandButton
    
    For lRow = 1 To dctBttnsRow.Count
        siLeft = HSPACE_LEFTRIGHT_BUTTONS
        Set frmRow = dctBttnsRow(lRow)
        If frmRow.Visible Then
            For Each v In dctBttns
                If Split(v, "-")(0) = lRow Then
                    lButton = Split(v, "-")(1)
                    Set cmb = dctBttns(v)
                    If cmb.Visible Then
                        With cmb
                            .Left = siLeft
                            .Width = siMaxButtonWidth
                            .Height = siMaxButtonHeight
                            .Top = siVmarginFrames
                            siLeft = .Left + .Width + HSPACE_BTTNS
                            If IsNumeric(vMsgButtonDefault) Then
                                If lButton = vMsgButtonDefault Then .Default = True
                            Else
                                If .Caption = vMsgButtonDefault Then .Default = True
                            End If
                        End With
                    End If
                End If
                HeightfrmBttnsArea = HeightfrmBttnsArea + siMaxButtonHeight + HSPACE_BTTNS
            Next v
        End If
        frmRow.Width = frmRow.Width + HSPACE_LEFTRIGHT_BUTTONS
    Next lRow
    Me.Height = frmMsectsArea.Top + frmMsectsArea.Height + HeightfrmBttnsArea
        
xt: Exit Sub
    
eh: If ErrMsg(ErrSrc(PROC)) = vbYes Then: Stop: Resume
End Sub

Private Sub SizeAndPosition2Bttns2Rows()
' ------------------------------------------------------------------------------
' Adjust all applied/visible button rows height to the maximum buttons height
' and the row frames width to the number of the displayed buttons considering a
' certain margin between the buttons (HSPACE_BTTNS) and a margin at the
' left and the right.
' ------------------------------------------------------------------------------
    Const PROC = "SizeAndPosition2Bttns2Rows"
    
    On Error GoTo eh
    Dim frmRow          As MSForms.Frame
    Dim siTop           As Single
    Dim v               As Variant
    Dim lButtons        As Long
    Dim siHeight        As Single
    Dim BttnsFrmWidth   As Single
    Dim dct             As Dictionary:      Set dct = AppliedBttnRows
    
    '~~ Adjust button row's width and height
    siHeight = AppliedButtonRowHeight
    siTop = siVmarginFrames
    For Each v In dct
        Set frmRow = v
        lButtons = dct(v)
        If frmRow.Visible Then
            With frmRow
                .Top = siTop
                .Height = siHeight
                '~~ Provide some extra space for the button's design
                BttnsFrmWidth = CInt((siMaxButtonWidth * lButtons) _
                               + (HSPACE_BTTNS * (lButtons - 1)) _
                               + (siHmarginFrames * 2)) - HSPACE_BTTNS + 7
                .Width = BttnsFrmWidth + (HSPACE_LEFTRIGHT_BUTTONS * 2)
                siTop = .Top + .Height + VSPACE_BTTN_ROWS
            End With
        End If
    Next v
    Set dct = Nothing

xt: Exit Sub
    
eh: If ErrMsg(ErrSrc(PROC)) = vbYes Then: Stop: Resume
End Sub

Private Sub SizeAndPosition2Bttns3Frame()
' ------------------------------------------------------------------------------
' Adjust the frame around all button row frames to the maximum width calculated
' by the adjustment of each of the rows frame.
' ------------------------------------------------------------------------------
    Const PROC = "SizeAndPosition2Bttns3Frame"
    
    On Error GoTo eh
    Dim v           As Variant
    Dim siWidth     As Single
    Dim siHeight    As Single
    Dim frm         As MSForms.Frame
    
    If frmBttnsFrm.Visible Then
        siWidth = ContentWidth(frmBttnsFrm)
        siHeight = ContentHeight(frmBttnsFrm)
        With frmBttnsFrm
            .Top = 0
            BttnsFrm.Height = siHeight
            BttnsFrm.Width = siWidth
            '~~ Center all button rows within the buttons frame
            For Each v In dctBttnsRow
                Set frm = dctBttnsRow(v)
                If frm.Visible Then
                    FrameCenterHorizontal center_frame:=frm, within_frame:=frmBttnsFrm
                End If
            Next v
        End With
    End If
    If BttnsArea.Height <> Max(Me.InsideHeight, ContentHeight(BttnsFrm)) Then
        NewHeight(BttnsArea) = Max(Me.InsideHeight, ContentHeight(BttnsFrm))
    End If
xt: Exit Sub
    
eh: If ErrMsg(ErrSrc(PROC)) = vbYes Then: Stop: Resume
End Sub

Private Sub SizeAndPosition2Bttns4Area()
' ------------------------------------------------------------------------------
' Adjust the buttons area frame in accordance with the buttons frame.
' ------------------------------------------------------------------------------
    Const PROC = "SizeAndPosition2Bttns4Area"
    
    On Error GoTo eh
    Dim siHeight    As Single
    Dim siWidth     As Single
    Dim frm         As MSForms.Frame: Set frm = BttnsArea
    
    siHeight = ContentHeight(frm) + ScrollHscrollHeight(frm)
    siWidth = ContentWidth(frm) + ScrollVscrollWidth(frm)
    If frm.Width <> Min(siWidth, (siMsgWidthMax - 10)) Then
        NewWidth(frm) = Min(siWidth, (siMsgWidthMax - 10))
    End If
    
    If frm.Height <> Min(siHeight, (siMsgHeightMax - 30)) Then
        NewHeight(frm) = Min(siHeight, (siMsgHeightMax - 30))
    End If
    siHeight = ContentHeight(frm) + ScrollHscrollHeight(frm)
    siWidth = ContentWidth(frm) + ScrollVscrollWidth(frm)
    If frm.Width <> Min(siWidth, (siMsgWidthMax - 10)) Then
        NewWidth(frm) = Min(siWidth, (siMsgWidthMax - 10))
    End If
    If frm.Height <> Min(siHeight, (siMsgHeightMax - 30)) Then
        NewHeight(frm) = Min(siHeight, (siMsgHeightMax - 30))
    End If
    
    If Not ScrollHscrollApplied(frm) Then
        frm.Width = BttnsFrm.Left + BttnsFrm.Width + ScrollVscrollWidth(frm)
    End If
    
    If Not ScrollHscrollApplied(frm) Then
        If Not ScrollVscrollApplied(frm) Then
'            frm.Height = BttnsFrm.Top + BttnsFrm.Height + ScrollHscrollHeight(frm)
            frm.Height = ContentHeight(frmBttnsFrm) + ScrollHscrollHeight(frm)
        End If
    End If
    
    FormWidth = frm.Width + ScrollVscrollWidth(frm)
    FrameCenterHorizontal center_frame:=frmBttnsArea, left_margin:=10
    Me.Height = ContentHeight(frm.Parent) + 35
    
xt: Exit Sub
    
eh: If ErrMsg(ErrSrc(PROC)) = vbYes Then: Stop: Resume
End Sub

Public Sub TimerBegin()
    cyTimerTicksBegin = TimerSysCurrentTicks
End Sub

Public Function TimerEnd() As Currency
    cyTimerTicksEnd = TimerSysCurrentTicks
    TimerEnd = TimerSecsElapsed * 1000
End Function

Private Sub UserForm_Activate()
' -------------------------------------------------------------------------------
' To avoid screen flicker the setup may has been done already. However for test
' purpose the Setup may run with the Activate event i.e. the .Show
' -------------------------------------------------------------------------------
    If Not bSetUpDone Then Setup
End Sub

Private Sub VisualizeCtl(ByVal vc_ctl As MSForms.Control, _
                         ByVal vc_backcolor As Long)
' ------------------------------------------------------------------------------
' Visualizes the Control (vc_ctl) with the BackColor (vc_backcolor) when
' bVisualizeForTest  is TRUE.
' ------------------------------------------------------------------------------
    
    With vc_ctl
        If bVisualizeForTest Then
            .BackColor = vc_backcolor
            .BorderStyle = fmBorderStyleNone
        Else
            .BackColor = Me.BackColor
            .BorderColor = Me.BackColor
            .BorderStyle = fmBorderStyleSingle
        End If
    End With
    
End Sub

Private Sub VisualizeSetupStep(ByVal vss_status As String)
' -------------------------------------------------------------------------------
' Display the Message form with visualized controls at the top left.
' -------------------------------------------------------------------------------
    With Me
        If .VisualizeForTest Then
            .Top = 10
            .Left = 10
            Application.StatusBar = "Setup step visualization for debug and test: " & vss_status
            Stop
        End If
    End With
End Sub

Private Function TimedDoEvents(ByVal tde_source As String) As String
' ---------------------------------------------------------------------------
' For the execution of a DoEvents statement. Provides the information in
' which procedure it had been executed and the msecs delay it has caused.
'
' Note: DoEvents every now and then is able to solve timing problems. When
'       looking at the description of its effect this often appears
'       miraculous. However, when it helps ... . But DoEvents allow keyboard
'       interaction while a process executes. In case of a loop - and when
'       the DoEvents lies within it, this may be a godsend. But it as well
'       may cause unpredictable results. This little procedure at least
'       documents in the Immediate window when (with milliseconds) and where
'       it had been executed.
' ---------------------------------------------------------------------------
    Dim s As String
    
    TimerBegin
    DoEvents
    s = Format(Now(), "hh:mm:ss") & ":" _
      & Right(Format(Timer, "0.000"), 3) _
      & " DoEvents paused the execution for " _
      & Format(TimerEnd, "00000") _
      & " msecs in '" & tde_source & "'"
    Debug.Print s
    TimedDoEvents = s
    
End Function

Public Property Let ApplicationRunArgs(ByVal dct As Dictionary)
    Set dctApplicationRunArgs = dct
End Property

Private Sub ApplicationRunViaButton(ByVal ar_button As String)
' --------------------------------------------------------------------------
' Performs an Application.Run for a button's caption.
' Preconditions: - Application.Run arguments had been provided by the caller
'                  via ApplicationRunArgsLetViaButton for the button
'                  (ar_button)
'                - The form has been displayed "Modeless"
' --------------------------------------------------------------------------
    Const PROC = "ApplicationRunViaButton"
    
    On Error GoTo eh
    Dim cll         As Collection
    Dim sService    As String
    Dim Msg         As TypeMsg
    Dim i           As Long
    Dim sKey        As String
    Dim sButton     As String
    
    sButton = Replace(Replace(ar_button, vbCrLf, "|"), vbLf, "|")
    Debug.Print "Button:   '" & sButton & "'"
    For i = 0 To dctApplicationRunArgs.Count - 1
        sKey = Replace(Replace(dctApplicationRunArgs.Keys()(i), vbCrLf, "|"), vbLf, "|")
        If sKey = sButton Then
            Set cll = dctApplicationRunArgs.Items()(i)
            sService = cll(1).Name & "!" & cll(2)
            Debug.Print "Key     : '" & sKey & "'"
            Debug.Print "sService:   " & sService
            
            Select Case cll.Count
                Case 2: Application.Run sService                 ' service call without arguments
                Case 3: Application.Run sService, cll(3)
                Case 4: Application.Run sService, cll(3), cll(4)
                Case 5: Application.Run sService, cll(3), cll(4), cll(5)
                Case 6: Application.Run sService, cll(3), cll(4), cll(5), cll(6)
                Case 7: Application.Run sService, cll(3), cll(4), cll(5), cll(6), cll(7)
                Case 8: Application.Run sService, cll(3), cll(4), cll(5), cll(6), cll(7), cll(8)
                Case 9: Application.Run sService, cll(3), cll(4), cll(5), cll(6), cll(7), cll(8), cll(9)
                Case 10: Application.Run sService, cll(3), cll(4), cll(5), cll(6), cll(7), cll(8), cll(9), cll(10)
                Case 11: Application.Run sService, cll(3), cll(4), cll(5), cll(6), cll(7), cll(8), cll(9), cll(10), cll(11)
                Case 12: Application.Run sService, cll(3), cll(4), cll(5), cll(6), cll(7), cll(8), cll(9), cll(10), cll(11), cll(12)
                Case 13: Application.Run sService, cll(3), cll(4), cll(5), cll(6), cll(7), cll(8), cll(9), cll(10), cll(11), cll(12), cll(13)
                Case 14: Application.Run sService, cll(3), cll(4), cll(5), cll(6), cll(7), cll(8), cll(9), cll(10), cll(11), cll(12), cll(13), cll(14)
                Case 15: Application.Run sService, cll(3), cll(4), cll(5), cll(6), cll(7), cll(8), cll(9), cll(10), cll(11), cll(12), cll(13), cll(14), cll(15)
                Case 16: Application.Run sService, cll(3), cll(4), cll(5), cll(6), cll(7), cll(8), cll(9), cll(10), cll(11), cll(12), cll(13), cll(14), cll(15), cll(16)
            End Select
            GoTo xt
        End If
    Next i
        
    With Msg.Section(1).Text
        .Text = "This button is useless when the form is displayed modeless," & vbLf & _
                "unless it is provided by the caller with an action to perform " & vbLf & _
                "- which is not the case!"
        .FontColor = rgbRed
        .FontBold = True
        .MonoSpaced = True
    End With
    With Msg.Section(2)
        With .Label
            .Text = "What this means:"
            .FontColor = rgbBlue
        End With
        .Text.Text = "In a modeless displayed form there should be no buttons other than those an 'Application.Run' service had been specified to be performed when clicked."
    End With
    
    mMsg.Dsply dsply_title:="No 'Application.Run' information provided for this button!" _
             , dsply_msg:=Msg
    
xt: Exit Sub

eh: Select Case ErrMsg(ErrSrc(PROC))
        Case vbResume:  Stop: Resume
        Case Else:      GoTo xt
    End Select
End Sub

