VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} fMsg 
   ClientHeight    =   10560
   ClientLeft      =   150
   ClientTop       =   390
   ClientWidth     =   12390
   OleObjectBlob   =   "fMsg.frx":0000
   StartUpPosition =   2  'Bildschirmmitte
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
' Uses:         Module mMsg to pass on the clicked reply button to the caller.
'               Note: The UserForm cannot be used directly unless the implemen-
'               tation is mimicked.
'
' Requires:     Reference to "Microsoft Scripting Runtime"
'
' See details at:
' https://warbe-maker.github.io/warbe-maker.github.io/vba/common/2020/11/17/Common-VBA-Message-Services.html
'
' W. Rauschenberger Berlin, April 2021 (last revision)
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
Const HSPACE_BUTTONS                As Single = 4               ' Horizontal space between reply buttons
Const HSPACE_LEFT                   As Single = 0               ' Left margin for labels and text boxes
Const HSPACE_RIGHT                  As Single = 15              ' Horizontal right space for labels and text boxes
Const HSPACE_LEFTRIGHT_BUTTONS      As Long = 8                 ' The margin before the left most and after the right most button
Const MARGIN_RIGHT_MSG_AREA         As String = 7
Const NEXT_ROW                      As String = vbLf            ' Reply button row break
Const VSCROLLBAR_WIDTH              As Single = 10              ' Additional horizontal space required for a frame with a vertical scrollbar
Const HSCROLLBAR_HEIGHT             As Single = 18              ' Additional vertical space required for a frame with a horizontal scroll barr
Const TEST_WITH_FRAME_BORDERS       As Boolean = False          ' For test purpose only! Display frames with visible border
Const TEST_WITH_FRAME_CAPTIONS      As Boolean = False          ' For test purpose only! Display frames with their test captions (erased by default)
Const VSPACE_AREAS                  As Single = 10              ' Vertical space between message area and replies area
Const VSPACE_BOTTOM                 As Single = 35              ' Vertical space at the bottom after the last displayed area
Const VSPACE_BTTN_ROWS              As Single = 5               ' Vertical space between button rows
Const VSPACE_LABEL                  As Single = 0               ' Vertical space between the section-label and the following section-text
Const VSPACE_SECTIONS               As Single = 7               ' Vertical space between displayed message sections
Const VSPACE_TEXTBOXES              As Single = 18              ' Vertical bottom marging for all textboxes
Const VSPACE_TOP                    As Single = 2               ' Top position for the first displayed control
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

Private Enum enStartupPosition      ' ---------------------------
    sup_Manual = 0                  ' Used to position the
'    sup_CenterOwner = 1             ' final setup message form
'    sup_CenterScreen = 2            ' horizontally and vertically
'    sup_WindowsDefault = 3          ' centered on the screen
End Enum                            ' ---------------------------

Private Enum enMsgFormUsage
    usage_progress_display = 1
'    usage_message_display = 2
End Enum

Private AppliedBttns                    As Dictionary   ' Dictionary of applied buttons (key=CommandButton, item=row)
Private AppliedBttnsRetVal              As Dictionary   ' Dictionary of the applied buttons' reply value (key=CommandButton)
Private bDoneMonoSpacedSects            As Boolean
Private bDoneMsgArea                    As Boolean
Private bDonePropSpacedSects            As Boolean
Private bDoneTitle                      As Boolean
Private bDsplyFrmsWthCptnTestOnly       As Boolean
Private bFormEvents                     As Boolean
Private bMonitorMode                    As Boolean
Private bReplyWithIndex                 As Boolean
Private cllDsgnAreas                    As Collection   ' Collection of the two primary/top frames
Private cllDsgnBttnRows                 As Collection   ' Collection of the designed reply button row frames
Private cllDsgnBttns                    As Collection   ' Collection of the collection of the designed reply buttons of a certain row
Private cllDsgnBttnsFrame               As Collection
Private cllDsgnMsgSects                 As Collection   '
Private cllDsgnMsgSectsLabel            As Collection
Private cllDsgnMsgSectsTextBox          As Collection   ' Collection of section textboxes
Private cllDsgnMsgSectsTextFrame        As Collection   ' Collection of section textframes
Private cllDsgnRowBttns                 As Collection   ' Collection of a designed reply button row's buttons
Private dctAppliedControls              As Dictionary   ' Dictionary of all applied controls (versus just designed)
Private dctMonoSpaced                   As Dictionary
Private dctMonoSpacedTbx                As Dictionary
Private dctSectsLabel                   As Dictionary   ' Sect specific label either provided via properties MsgLabel or Msg
Private dctSectsMonoSpaced              As Dictionary   ' Sect specific monospace option either provided via properties MsgMonospaced or Msg
Private dctSectsText                    As Dictionary   ' Sect specific text either provided via properties MsgText or Msg
Private lBackColor                      As Long
Private lSetupRowButtons                As Long         ' number of buttons setup in a row
Private lSetupRows                      As Long         ' number of setup button rows
Private SetUpDone                       As Boolean
Private siHmarginButtons                As Single
Private siHmarginFrames                 As Single       ' Test property, value defaults to 0
Private siMaxButtonHeight               As Single
Private siMaxButtonWidth                As Single
Private siMinButtonWidth                As Single
Private siMsgHeightMax                  As Single       ' Maximum message height in pt
Private siMsgHeightMin                  As Single       ' Minimum message height in pt
Private siMsgWidthMax                   As Single       ' Maximum message width in pt
Private siMsgWidthMin                   As Single       ' Minimum message width in pt
Private siTitleWidth                    As Single
Private siVmarginButtons                As Single
Private siVmarginFrames                 As Single       ' Test property, value defaults to 0
Private sMonoSpacedLabelDefaultFontName As String
Private sMonoSpacedLabelDefaultFontSize As Single
Private sMonoSpacedTextDefaultFontName  As String
Private sMonoSpacedTextDefaultFontSize  As Single
Private sMsgTitle                       As String
Private sTitleFontName                  As String
Private sTitleFontSize                  As String       ' Ignored when sTitleFontName is not provided
Private TitleWidth                      As Single
Private UsageType                       As enMsgFormUsage
Private vbuttons                        As Variant
Private VirtualScreenHeightPts          As Single
Private VirtualScreenLeftPts            As Single
Private VirtualScreenTopPts             As Single
Private VirtualScreenWidthPts           As Single
Private vMsgButtonDefault                    As Variant      ' Index or caption of the default button
Private vReplyValue                     As Variant

Private Sub UserForm_Initialize()
    Const PROC = "UserForm_Initialize"
    
    On Error GoTo eh
    Set dctMonoSpaced = New Dictionary
    Set dctMonoSpacedTbx = New Dictionary
    
    siMinButtonWidth = DFLT_BTTN_MIN_WIDTH
    siHmarginButtons = HSPACE_BUTTONS
    siVmarginButtons = VSPACE_BTTN_ROWS
    bFormEvents = False
    ' Get the display screen's dimensions and position in pts
    GetScreenMetrics VirtualScreenLeftPts _
                   , VirtualScreenTopPts _
                   , VirtualScreenWidthPts _
                   , VirtualScreenHeightPts
    sMonoSpacedTextDefaultFontName = DFLT_TXT_MONOSPACED_FONT_NAME
    sMonoSpacedTextDefaultFontSize = DFLT_TXT_MONOSPACED_FONT_SIZE
    sMonoSpacedLabelDefaultFontName = DFLT_LBL_MONOSPACED_FONT_NAME
    sMonoSpacedLabelDefaultFontSize = DFLT_LBL_MONOSPACED_FONT_SIZE
    bDsplyFrmsWthCptnTestOnly = False
    DsplyFrmsWthBrdrsTestOnly = False
    siHmarginFrames = 0     ' Ensures proper command buttons framing, may be used for test purpose
    Me.VmarginFrames = 0    ' Ensures proper command buttons framing and vertical positioning of controls
    SetUpDone = False
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
    Set cllDsgnAreas = Nothing
    Set cllDsgnBttnRows = Nothing
    Set cllDsgnBttns = Nothing
    Set cllDsgnBttnsFrame = Nothing
    Set cllDsgnMsgSects = Nothing
    Set cllDsgnMsgSectsLabel = Nothing
    Set cllDsgnMsgSectsTextBox = Nothing
    Set cllDsgnMsgSectsTextFrame = Nothing
    Set cllDsgnRowBttns = Nothing
    Set dctAppliedControls = Nothing
    Set dctMonoSpaced = Nothing
    Set dctMonoSpacedTbx = Nothing
    Set dctSectsLabel = Nothing
    Set dctSectsMonoSpaced = Nothing
    Set dctSectsText = Nothing
End Sub

Private Property Get AppliedButtonRetVal(Optional ByVal Button As MSForms.CommandButton) As Variant
    AppliedButtonRetVal = AppliedBttnsRetVal(Button)
End Property

Private Property Let AppliedButtonRetVal(Optional ByVal Button As MSForms.CommandButton, ByVal v As Variant)
    AppliedBttnsRetVal.Add Button, v
End Property

Private Property Get AppliedButtonRowHeight() As Single
    AppliedButtonRowHeight = siMaxButtonHeight + 2
End Property

Private Property Let AppliedControls( _
                      Optional ByVal msg_section As Long = 0, _
                               ByVal v As Variant)
' ------------------------------------------------------------------------------
' Register any applied (used) control and set it visible.
' ------------------------------------------------------------------------------
If dctAppliedControls Is Nothing Then Set dctAppliedControls = New Dictionary
    If Not dctAppliedControls.Exists(v) Then
        v.Visible = True
        dctAppliedControls.Add v, msg_section
    End If
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

Private Property Get DsgnBttn(Optional ByVal bttn_row As Long, Optional ByVal bttn_no As Long) As MSForms.CommandButton
    Set DsgnBttn = cllDsgnBttns(bttn_row)(bttn_no)
End Property

Private Property Get DsgnBttnRow(Optional ByVal row As Long) As MSForms.Frame:              Set DsgnBttnRow = cllDsgnBttnRows(row):                             End Property

Private Property Get DsgnBttnRows() As Collection:                                          Set DsgnBttnRows = cllDsgnBttnRows:                                 End Property

Private Property Get DsgnBttnsArea() As MSForms.Frame:                                      Set DsgnBttnsArea = cllDsgnAreas(2):                                End Property

Private Property Get DsgnBttnsFrame() As MSForms.Frame:                                     Set DsgnBttnsFrame = cllDsgnBttnsFrame(1):                          End Property

Private Property Get DsgnMsgArea() As MSForms.Frame:                                        Set DsgnMsgArea = cllDsgnAreas(1):                                  End Property

Private Property Get DsgnMsgSect(Optional msg_section As Long) As MSForms.Frame:            Set DsgnMsgSect = cllDsgnMsgSects(msg_section):                     End Property

Private Property Get DsgnMsgSectLabel(Optional msg_section As Long) As MSForms.Label:       Set DsgnMsgSectLabel = cllDsgnMsgSectsLabel(msg_section):           End Property

Private Property Get DsgnMsgSects() As Collection:                                          Set DsgnMsgSects = cllDsgnMsgSects:                                 End Property

Private Property Get DsgnMsgSectsTextFrame() As Collection:                                 Set DsgnMsgSectsTextFrame = cllDsgnMsgSectsTextFrame:               End Property

Private Property Get DsgnMsgSectTextBox(Optional msg_section As Long) As MSForms.TextBox:   Set DsgnMsgSectTextBox = cllDsgnMsgSectsTextBox(msg_section):       End Property

Private Property Get DsgnMsgSectTextFrame(Optional ByVal msg_section As Long):              Set DsgnMsgSectTextFrame = cllDsgnMsgSectsTextFrame(msg_section):   End Property

Public Property Let DsplyFrmsWthBrdrsTestOnly(ByVal dsply_borders As Boolean)
    
    Dim ctl         As MSForms.Control
    Dim lBackColor  As Long
    
    lBackColor = Me.BackColor
    For Each ctl In Me.Controls
        If TypeName(ctl) = "Frame" Or TypeName(ctl) = "TextBox" Then
            If dsply_borders Then
                ctl.BorderColor = -2147483638   ' active frame, allows with style none to hide the frame
                ctl.BorderStyle = fmBorderStyleSingle
            Else
                ctl.BorderColor = lBackColor
                ctl.BorderStyle = fmBorderStyleNone
            End If
        End If
    Next ctl
    
End Property

Public Property Let DsplyFrmsWthCptnTestOnly(ByVal b As Boolean):                           bDsplyFrmsWthCptnTestOnly = b:                                      End Property

Public Property Get FormContentWidth(Optional ByRef with_vertical_scrollbars As Boolean) As Single
' ------------------------------------------------------------------------------
' Returns the maximum width of the form's content by considering only
' applied/visible controls.
' ------------------------------------------------------------------------------
    
    Dim ctl As MSForms.Control
    Dim frm As MSForms.Frame
    Dim frm_ctl As MSForms.Control
    
    For Each ctl In Me.Controls
        With ctl
            If .Parent Is Me Then
                If IsApplied(ctl) Then
                    FormContentWidth = Max(FormContentWidth, .Left + .Width)
                    If TypeName(ctl) = "Frame" Then
                        Set frm = ctl
                        If ScrollVerticalApplied(frm) Then
                            with_vertical_scrollbars = True
                        End If
                    End If
                End If
            End If
        End With
    Next ctl
    
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
    Me.Width = Min(new_width, siMsgWidthMax + Max(ScrollVerticalWidth(DsgnMsgArea), ScrollVerticalWidth(DsgnBttnsArea)))
End Property

Private Property Get FormWidthMaxUsable()
    FormWidthMaxUsable = siMsgWidthMax - 15
End Property

Public Property Get FrameContentHeight(ByRef frm As MSForms.Frame) As Single
' ------------------------------------------------------------------------------
' Returns the height of the Frames (frm) content by considering only
' applied/visible controls.
' ------------------------------------------------------------------------------
    Dim ctl As MSForms.Control
    
    For Each ctl In frm.Controls
        If ctl.Parent Is frm Then
            If IsApplied(ctl) Then
                FrameContentHeight = Max(FrameContentHeight, ctl.top + ctl.Height)
            End If
        End If
    Next ctl

End Property

Public Property Get FrameContentWidth( _
                       Optional ByRef v As Variant) As Single
' ------------------------------------------------------------------------------
' Returns the maximum width of the frames (frm) content by considering only
' applied/visible controls.
' ------------------------------------------------------------------------------
    
    Dim ctl As MSForms.Control
    Dim frm As MSForms.Frame
    Dim frm_ctl As MSForms.Control
    
    If TypeName(v) = "Frame" Then Set frm_ctl = v Else Stop
    For Each ctl In frm_ctl.Controls
        With ctl
            If .Parent Is frm_ctl Then
                If IsApplied(ctl) Then
                    FrameContentWidth = Max(FrameContentWidth, .Left + .Width)
                End If
            End If
        End With
    Next ctl
    
End Property

Public Property Let FrameHeight( _
                 Optional ByRef frm As MSForms.Frame, _
                 Optional ByVal y_action As fmScrollAction = fmScrollActionBegin, _
                          ByVal frm_height As Single)
' ------------------------------------------------------------------------------
' Mimics a frame's height change event. When the height of the frame (frm) is
' changed (frm_height) to less than the frame's content height and no vertical
' scrollbar is applied one is applied with the frame content's height. If one
' is already applied just the height is adjusted to the frame content's height.
' When the height becomes more than the frame's content height a vertical
' scrollbar becomes obsolete and is removed.
' ------------------------------------------------------------------------------
    Const PROC          As String = "FrameHieght"
    
    On Error GoTo eh
    Dim yAction         As fmScrollAction
    Dim ContentHeight   As Single:          ContentHeight = FrameContentHeight(frm)
    
    frm.Height = frm_height
    If frm.Height < ContentHeight Then
        '~~ Apply a vertical scrollbar if none is applied yet, adjust its height otherwise
        If Not ScrollVerticalApplied(frm) Then
            ScrollVerticalApply frm, ContentHeight, yAction
        Else
            frm.ScrollHeight = ContentHeight
            frm.Scroll yAction:=y_action
        End If
    Else
        '~~ With the frame's height is greater or equal its content height
        '~~ a vertical scrollbar becomes obsolete and is removed
        With frm
            Select Case .ScrollBars
                Case fmScrollBarsBoth:      .ScrollBars = fmScrollBarsHorizontal
                Case fmScrollBarsVertical:  .ScrollBars = fmScrollBarsNone
            End Select
        End With
    End If
    
xt: Exit Property
    
eh: If ErrMsg(ErrSrc(PROC)) = vbYes Then: Stop: Resume
End Property

Public Property Let FrameWidth( _
                 Optional ByRef frm As MSForms.Frame, _
                          ByVal frm_width As Single)
' ------------------------------------------------------------------------------
' Mimics a frame's width change event. When the width of the frame (frm) is
' changed (frm_width) a horizontal scrollbar will be applied - or adjusted to
' the frame content's width. I.e. this property must only be used when a
' horizontal scrollbar is applicable/desired in case.
' ------------------------------------------------------------------------------
    Dim ContentWidth As Single: ContentWidth = FrameContentWidth(frm)
    
    frm.Width = frm_width
    If frm_width < ContentWidth Then
        '~~ Apply a horizontal scrollbar if none is applied yet, adjust its width otherwise
        If Not ScrollHorizontalApplied(frm) Then
            ScrollHorizontalApply frm, ContentWidth
        Else
            frm.ScrollWidth = ContentWidth
        End If
    Else
        '~~ With the frame's width greater or equal its content width
        '~~ a horizontal scrollbar becomes obsolete and is removed
        With frm
            Select Case .ScrollBars
                Case fmScrollBarsBoth:          .ScrollBars = fmScrollBarsVertical
                Case fmScrollBarsHorizontal:    .ScrollBars = fmScrollBarsNone
            End Select
        End With
    End If
    
End Property

Public Property Let HmarginButtons(ByVal si As Single):                                     siHmarginButtons = si:                                              End Property

Public Property Let HmarginFrames(ByVal si As Single):                                      siHmarginFrames = si:                                               End Property

Public Property Get IsApplied(Optional ByVal v As Variant) As Boolean
    If dctAppliedControls Is Nothing _
    Then IsApplied = False _
    Else IsApplied = dctAppliedControls.Exists(v)
End Property

Private Property Get MaxRowsHeight() As Single:                                             MaxRowsHeight = siMaxButtonHeight + (siVmarginFrames * 2):          End Property

Private Property Get MaxWidthBttnsArea() As Single
    MaxWidthBttnsArea = FormWidthMaxUsable - (HSPACE_BTTN_AREA * 2)
End Property

Private Property Get MaxWidthMsgArea() As Single
' ------------------------------------------------------------------------------
' The maximum usable message area width considers the specified maximum form
' width and the InsideWidth
' ------------------------------------------------------------------------------
    MaxWidthMsgArea = Me.InsideWidth
End Property

Public Property Let MinButtonWidth(ByVal si As Single):                                     siMinButtonWidth = si:                                                          End Property

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

Public Property Let MsgButtonDefault(ByVal vDefault As Variant)
    vMsgButtonDefault = vDefault
End Property

Public Property Let MsgButtons(ByVal v As Variant)
        
    Select Case VarType(v)
        Case vbLong, vbString:  vbuttons = v
        Case vbEmpty:           vbuttons = vbOKOnly
        Case Else
            If IsArray(v) Then
                vbuttons = v
            ElseIf TypeName(v) = "Collection" Or TypeName(v) = "Dictionary" Then
                Set vbuttons = v
            End If
    End Select
End Property

Public Property Get MsgHeightMax() As Single:           MsgHeightMax = siMsgHeightMax:  End Property

Public Property Let MsgHeightMax(ByVal si As Single):   siMsgHeightMax = si:            End Property

Public Property Get MsgHeightMin() As Single:           MsgHeightMin = siMsgHeightMin:  End Property

Public Property Let MsgHeightMin(ByVal si As Single):   siMsgHeightMin = si:            End Property

Public Property Get MsgLabel( _
              Optional ByVal msg_section As Long) As TypeMsgLabel
' ------------------------------------------------------------------------------
' Transfers a section's message UDT stored as array back into the UDT.
' ------------------------------------------------------------------------------
    Dim vArry() As Variant
    
    If dctSectsLabel Is Nothing Then
        MsgLabel.Text = vbNullString
    Else
        If dctSectsLabel.Exists(msg_section) Then
            vArry = dctSectsLabel(msg_section)
            MsgLabel.FontBold = vArry(0)
            MsgLabel.FontColor = vArry(1)
            MsgLabel.FontItalic = vArry(2)
            MsgLabel.FontName = vArry(3)
            MsgLabel.FontSize = vArry(4)
            MsgLabel.FontUnderline = vArry(5)
            MsgLabel.MonoSpaced = vArry(6)
            MsgLabel.Text = vArry(7)
        Else
            MsgLabel.Text = vbNullString
        End If
    End If
End Property

Public Property Let MsgLabel( _
              Optional ByVal msg_section As Long, _
                       ByRef msg_label As TypeMsgLabel)
' ------------------------------------------------------------------------------
' Transfers a message label UDT (msg_label) into an array and stores it in the
' Dictionary (dctSectsLabel) with the section (msg_section) as the key.
' ------------------------------------------------------------------------------
    Dim vArry(7)    As Variant
    
    If dctSectsLabel Is Nothing Then Set dctSectsLabel = New Dictionary
    If Not dctSectsLabel.Exists(msg_section) Then
        vArry(0) = msg_label.FontBold
        vArry(1) = msg_label.FontColor
        vArry(2) = msg_label.FontItalic
        vArry(3) = msg_label.FontName
        vArry(4) = msg_label.FontSize
        vArry(5) = msg_label.FontUnderline
        vArry(6) = msg_label.MonoSpaced
        vArry(7) = msg_label.Text
        dctSectsLabel.Add msg_section, vArry
    End If
End Property

Public Property Get MsgMonoSpaced(Optional ByVal msg_section As Long) As Boolean
    Dim vArry() As Variant
    
    If dctSectsText Is Nothing Then
        MsgMonoSpaced = False
    Else
        With dctSectsText
            If .Exists(msg_section) Then
                vArry = .Item(msg_section)
                MsgMonoSpaced = vArry(6)
            Else
                MsgMonoSpaced = False
            End If
        End With
    End If
End Property

Public Property Get MsgText( _
             Optional ByVal msg_section As Long) As TypeMsgText
' ------------------------------------------------------------------------------
' Transferes message UDT stored as array back into the UDT.
' ------------------------------------------------------------------------------
    Dim vArry() As Variant
    
    If dctSectsText Is Nothing Then
        MsgText.Text = vbNullString
    Else
        If dctSectsText.Exists(msg_section) Then
            vArry = dctSectsText(msg_section)
            MsgText.FontBold = vArry(0)
            MsgText.FontColor = vArry(1)
            MsgText.FontItalic = vArry(2)
            MsgText.FontName = vArry(3)
            MsgText.FontSize = vArry(4)
            MsgText.FontUnderline = vArry(5)
            MsgText.MonoSpaced = vArry(6)
            MsgText.Text = vArry(7)
        Else
            MsgText.Text = vbNullString
        End If
    End If

End Property

Public Property Let MsgText( _
             Optional ByVal msg_section As Long, _
                      ByRef msg_msg As TypeMsgText)
' ------------------------------------------------------------------------------
' Transfers a message UDT into an array and stores it in a Dictionary with the
' section number (msg_section) as the key.
' ------------------------------------------------------------------------------
    Dim vArry(7)    As Variant
    
    If dctSectsText Is Nothing Then Set dctSectsText = New Dictionary
    If Not dctSectsText.Exists(msg_section) Then
        vArry(0) = msg_msg.FontBold
        vArry(1) = msg_msg.FontColor
        vArry(2) = msg_msg.FontItalic
        vArry(3) = msg_msg.FontName
        vArry(4) = msg_msg.FontSize
        vArry(5) = msg_msg.FontUnderline
        vArry(6) = msg_msg.MonoSpaced
        vArry(7) = msg_msg.Text
        dctSectsText.Add msg_section, vArry
    End If
End Property

Public Property Get MsgTitle() As String:               MsgTitle = Me.Caption:          End Property

Public Property Let MsgTitle(ByVal s As String):        sMsgTitle = s:                  End Property

Public Property Get MsgWidthMax() As Single:            MsgWidthMax = siMsgWidthMax:    End Property

Public Property Let MsgWidthMax(ByVal si As Single):    siMsgWidthMax = si:             End Property

Public Property Get MsgWidthMin() As Single:            MsgWidthMin = siMsgWidthMin:    End Property

Public Property Let MsgWidthMin(ByVal si As Single):    siMsgWidthMin = si:             End Property

Public Property Get NoOfDesignedMsgSects() As Long ' -----------------------
    NoOfDesignedMsgSects = 4                       ' Global definition !!!!!
End Property                                       ' -----------------------

Private Property Get PrcntgHeightBttnsArea() As Single
    PrcntgHeightBttnsArea = Round(DsgnBttnsArea.Height / (DsgnMsgArea.Height + DsgnBttnsArea.Height), 2)
End Property

Private Property Get PrcntgHeightMsgArea() As Single
    PrcntgHeightMsgArea = Round(DsgnMsgArea.Height / (DsgnMsgArea.Height + DsgnBttnsArea.Height), 2)
End Property

Public Property Let MonitorMode(ByVal b As Boolean):                                        bMonitorMode = b:                                              End Property

Public Property Get ReplyValue() As Variant:                                                ReplyValue = vReplyValue:                                           End Property

Public Property Let ReplyWithIndex(ByVal b As Boolean):                                     bReplyWithIndex = b:                                                End Property

Private Property Get ScrollBarHeight(Optional ByVal frm As MSForms.Frame) As Single
    If frm.ScrollBars = fmScrollBarsBoth Or frm.ScrollBars = fmScrollBarsHorizontal Then ScrollBarHeight = 14
End Property

Private Property Get ScrollBarWidth(Optional ByVal frm As MSForms.Frame) As Single
    If frm.ScrollBars = fmScrollBarsBoth Or frm.ScrollBars = fmScrollBarsVertical Then ScrollBarWidth = 12
End Property

Public Property Get VmarginButtons() As Single:                                             VmarginButtons = siVmarginButtons:                                  End Property

Public Property Let VmarginButtons(ByVal si As Single):                                     siVmarginButtons = si:                                              End Property

Public Property Get VmarginFrames() As Single:                                              VmarginFrames = siVmarginFrames:                                    End Property

Public Property Let VmarginFrames(ByVal si As Single):                                      siVmarginFrames = VgridPos(si):                                     End Property

Public Function AppErr(ByVal lNo As Long) As Long
' ------------------------------------------------------------------------------
' Converts a positive (i.e. an "application" error number into a negative number
' by adding vbObjectError. Converts a negative number back into a positive i.e.
' the original programmed application error number.
' Usage example:
'    Err.Raise mErH.AppErr(1), .... ' when an application error is detected
'    If Err.Number < 0 Then    ' when the error is displayed
'       MsgBox "Application error " & AppErr(Err.Number)
'    Else
'       MsgBox "VB Rutime Error " & Err.Number
'    End If
' ------------------------------------------------------------------------------
    AppErr = IIf(lNo < 0, AppErr = lNo - vbObjectError, AppErr = vbObjectError + lNo)
End Function

Private Function AppliedBttnRows() As Dictionary
' ------------------------------------------------------------------------------
' Returns a Dictionary of the applied/used/visible butoon rows with the row
' frame as the key and the applied/visible buttons therein as item.
' ------------------------------------------------------------------------------
    
    Dim dct             As New Dictionary
    Dim ButtonRows      As Long
    Dim ButtonRowsRow   As MSForms.Frame
    Dim v               As Variant
    Dim ButtonsInRow    As Long
    
    For ButtonRows = 1 To DsgnBttnRows.Count
        Set ButtonRowsRow = DsgnBttnRows(ButtonRows)
        If IsApplied(ButtonRowsRow) Then
            ButtonsInRow = 0
            For Each v In DsgnRowBttns(ButtonRows)
                If IsApplied(v) Then ButtonsInRow = ButtonsInRow + 1
            Next v
            dct.Add ButtonRowsRow, ButtonsInRow
        End If
    Next ButtonRows
    Set AppliedBttnRows = dct

End Function

Public Sub AutoSizeTextBox( _
                     ByRef as_tbx As MSForms.TextBox, _
                     ByVal as_text As String, _
            Optional ByVal as_width_limit As Single = 0, _
            Optional ByVal as_width_min As Single = 0, _
            Optional ByVal as_height_min As Single = 0, _
            Optional ByVal as_width_max As Single = 0, _
            Optional ByVal as_height_max As Single = 0, _
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
' Uses: FrameWidth, FrameContentWidth, ScrollHorizontalApply,
'       FrameHeight, FrameContentHeight, ScrollVerticalApply
'
' W. Rauschenberger Berlin June 2021
' ------------------------------------------------------------------------------
    
    With as_tbx
        .MultiLine = True
        If as_width_limit > 0 Then
            '~~ AutoSize the height of the TextBox considering the limited width
            .WordWrap = True
            .AutoSize = False
            .Width = as_width_limit - 7 ' the readability space is added later
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
        .Width = .Width + 7   ' readability space
        .Height = .Height + 7 ' redability space
        If as_width_min > 0 And .Width < as_width_min Then .Width = as_width_min
        If as_height_min > 0 And .Height < as_height_min Then .Height = as_height_min
        .Parent.Height = .top + .Height + 2
        .Parent.Width = .Left + .Width + 2
    End With
    
    '~~ When the parent of the TextBox is a frame scrollbars may have become applicable
    '~~ provided a mximimum with and/or height has been provided
    With as_tbx
        If TypeName(.Parent) = "Frame" Then
            '~~ When a max width is provided and exceeded a horizontal scrollbar is applied
            '~~ by the assignment of a frame width which is less than the frame's content width
            If as_width_max > 0 Then
                FrameWidth(.Parent) = Min(as_width_max, .Width + 2 + ScrollBarWidth(.Parent))
            End If
            '~~ When a max height is provided and exceeded a vertical scrollbar is applied
            '~~ by the assignment of a frame height which is less then the frame's content height
            If as_height_max > 0 Then
                FrameHeight(.Parent) = Min(as_height_max, .Height + ScrollBarHeight(.Parent))
            End If
        End If
    End With
    
xt: Exit Sub

End Sub

Private Sub ButtonClicked(ByVal cmb As MSForms.CommandButton)
' ------------------------------------------------------------------------------
' Return the value of the clicked reply button (button). When there is only one
' applied reply button the form is unloaded with the click of it. Otherwise the
' form is just hidden waiting for the caller to obtain the return value or
' index which then unloads the form.
' ------------------------------------------------------------------------------
    On Error Resume Next
    If bReplyWithIndex Then
        vReplyValue = ClickedButtonIndex(cmb)
        mMsg.RepliedWith = ClickedButtonIndex(cmb)
    Else
        vReplyValue = AppliedButtonRetVal(cmb)  ' global variable of calling module mMsg
        mMsg.RepliedWith = AppliedButtonRetVal(cmb)  ' global variable of calling module mMsg
    End If
    
    DisplayDone = True ' in case the form has been displayed modeless this will indicate the end of the wait loop
    Unload Me
    
End Sub

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

Private Sub Collect(ByRef cllct_into As Variant, _
                    ByVal cllct_with_parent As Variant, _
                    ByVal cllct_cntrl_type As String, _
                    ByVal cllct_set_height As Single, _
                    ByVal cllct_set_width As Single, _
           Optional ByVal cllct_set_visible As Boolean = False)
' ------------------------------------------------------------------------------
' Setup of a Collection (cllct_into) with all type (cllct_cntrl_type) controls
' with a parent (cllct_with_parent) as Collection (cllct_into) by assigning the
' an initial height (cllct_set_height) and width (cllct_set_width).
' ------------------------------------------------------------------------------
    Const PROC = "Collect"
    
    On Error GoTo eh
    Dim ctl         As MSForms.Control
    Dim v           As Variant
    
    lBackColor = Me.BackColor
    
    Set cllct_into = New Collection

    Select Case TypeName(cllct_with_parent)
        Case "Collection"
            '~~ Parent is each frame in the collection
            For Each v In cllct_with_parent
                For Each ctl In Me.Controls
                    If TypeName(ctl) = cllct_cntrl_type And ctl.Parent Is v Then
                        With ctl
                            If Not TypeName(ctl) = "CommandButton" Then
                                .BackColor = lBackColor
                            End If
                            .Visible = cllct_set_visible
                            .Height = cllct_set_height
                            .Width = cllct_set_width
                        End With
                        cllct_into.Add ctl
                    End If
               Next ctl
            Next v
        Case Else
            For Each ctl In Me.Controls
                If TypeName(ctl) = cllct_cntrl_type And ctl.Parent Is cllct_with_parent Then
                    With ctl
                        If Not TypeName(ctl) = "CommandButton" Then
                            .BackColor = lBackColor
                        End If
                        .Visible = cllct_set_visible
                        .Height = cllct_set_height
                        .Width = cllct_set_width
                    End With
                    Select Case TypeName(cllct_into)
                        Case "Collection"
                            cllct_into.Add ctl
                        Case Else
                            Set cllct_into = ctl
                    End Select
                End If
            Next ctl
    End Select

xt: Exit Sub
    
eh: If ErrMsg(ErrSrc(PROC)) = vbYes Then: Stop: Resume
End Sub

Private Sub CollectDesignControls()
' ----------------------------------------------------------------------
' Collects all designed controls without concidering any control's name.
' ----------------------------------------------------------------------
    Const PROC = "CollectDesignControls"
    
    On Error GoTo eh
    Dim v As Variant

    Collect cllct_into:=cllDsgnAreas _
          , cllct_cntrl_type:="Frame" _
          , cllct_with_parent:=Me _
          , cllct_set_height:=10 _
          , cllct_set_width:=Me.Width - siHmarginFrames
    
    Collect cllct_into:=cllDsgnMsgSects _
          , cllct_cntrl_type:="Frame" _
          , cllct_with_parent:=DsgnMsgArea _
          , cllct_set_height:=50 _
          , cllct_set_width:=DsgnMsgArea.Width - siHmarginFrames
    
    Collect cllct_into:=cllDsgnMsgSectsLabel _
          , cllct_cntrl_type:="Label" _
          , cllct_with_parent:=cllDsgnMsgSects _
          , cllct_set_height:=15 _
          , cllct_set_width:=DsgnMsgArea.Width - (siHmarginFrames * 2)
    
    Collect cllct_into:=cllDsgnMsgSectsTextFrame _
          , cllct_cntrl_type:="Frame" _
          , cllct_with_parent:=cllDsgnMsgSects _
          , cllct_set_height:=20 _
          , cllct_set_width:=DsgnMsgArea.Width - (siHmarginFrames * 2)
    
    Collect cllct_into:=cllDsgnMsgSectsTextBox _
          , cllct_cntrl_type:="TextBox" _
          , cllct_with_parent:=cllDsgnMsgSectsTextFrame _
          , cllct_set_height:=20 _
          , cllct_set_width:=DsgnMsgArea.Width - (siHmarginFrames * 3)
        
    Collect cllct_into:=cllDsgnBttnsFrame _
          , cllct_cntrl_type:="Frame" _
          , cllct_with_parent:=DsgnBttnsArea _
          , cllct_set_height:=10 _
          , cllct_set_width:=10 _
          , cllct_set_visible:=True ' minimum is one button
    
    Collect cllct_into:=cllDsgnBttnRows _
          , cllct_cntrl_type:="Frame" _
          , cllct_with_parent:=cllDsgnBttnsFrame _
          , cllct_set_height:=10 _
          , cllct_set_width:=10 _
          , cllct_set_visible:=False ' minimum is one button
        
    Set cllDsgnBttns = New Collection
    For Each v In cllDsgnBttnRows
        Collect cllct_into:=cllDsgnRowBttns _
              , cllct_cntrl_type:="CommandButton" _
              , cllct_with_parent:=v _
              , cllct_set_height:=10 _
              , cllct_set_width:=siMinButtonWidth
        cllDsgnBttns.Add cllDsgnRowBttns
    Next v
    
    ProvideDictionary dctAppliedControls ' provides a clean or new dictionary for collection applied controls
    ProvideDictionary AppliedBttns
    ProvideDictionary AppliedBttnsRetVal

xt: Exit Sub
    
eh: If ErrMsg(ErrSrc(PROC)) = vbYes Then: Stop: Resume
End Sub

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

Private Sub DisplayFramesWithCaptions( _
                        Optional ByVal b As Boolean = True)
' ---------------------------------------------------------
' When False (the default) captions are removed from all
' frames Else they remain visible for testing purpose
' ---------------------------------------------------------
            
    Dim ctl As MSForms.Control
       
    If Not b Then
        For Each ctl In Me.Controls
            If TypeName(ctl) = "Frame" Then
                ctl.Caption = vbNullString
            End If
        Next ctl
    End If

End Sub

Private Function DsgnRowBttns(ByVal ButtonRow As Long) As Collection
' --------------------------------------------------------------------
' Return a collection of applied/use/visible buttons in row buttonrow.
' --------------------------------------------------------------------
    Set DsgnRowBttns = cllDsgnBttns(ButtonRow)
End Function

Private Function ErrMsg(ByVal err_source As String, _
               Optional ByVal err_no As Long = 0, _
               Optional ByVal err_dscrptn As String = vbNullString, _
               Optional ByVal err_line As Long = 0) As Variant
' ------------------------------------------------------------------------------
' Common, minimum VBA error handling providing the means to resume the error
' line when the Conditional Compile Argument Debugging=1.
' Usage: When this procedure is copied into any desired module the statement
'        If ErrMsg(ErrSrc(PROC) = vbYes Then: Stop: Resume
'        is appropriate
'        The caller provides the source of the error through ErrSrc(PROC) where
'        ErrSrc is a procedure available in the module using this ErrMsg and
'        PROC is the constant identifying the procedure
' Uses: AppErr to translate a negative programmed application error into its
'              original positive number
' ------------------------------------------------------------------------------
    Dim ErrNo   As Long
    Dim ErrDesc As String
    Dim ErrType As String
    Dim errline As Long
    Dim AtLine  As String
    Dim Buttons As Long
    
    If err_no = 0 Then err_no = Err.Number
    If err_no < 0 Then
        ErrNo = AppErr(err_no)
        ErrType = "Applicatin error "
    Else
        ErrNo = err_no
        ErrType = "Runtime error "
    End If
    
    If err_line = 0 Then errline = Erl
    If err_line <> 0 Then AtLine = " at line " & err_line
    
    If err_dscrptn = vbNullString Then err_dscrptn = Err.Description
    If err_dscrptn = vbNullString Then err_dscrptn = "--- No error message available ---"
    ErrDesc = "Error: " & vbLf & err_dscrptn & vbLf & vbLf & "Source: " & vbLf & err_source & AtLine

    
#If Debugging Then
    Buttons = vbYesNo
    ErrDesc = ErrDesc & vbLf & vbLf & "Debugging: Yes=Resume error line, No=Continue"
#Else
    Buttons = vbCritical
#End If
    
    ErrMsg = MsgBox(Title:=ErrType & ErrNo & " in " & err_source _
                  , Prompt:=ErrDesc _
                  , Buttons:=Buttons)
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

Private Function Max(ParamArray va() As Variant) As Variant
' ---------------------------------------------------------
' Returns the maximum value of all values provided (va).
' ---------------------------------------------------------
    Dim v   As Variant
    
    Max = va(LBound(va)): If LBound(va) = UBound(va) Then Exit Function
    For Each v In va
        If v > Max Then Max = v
    Next v
    
End Function

Private Function MaxWidthMsgSect(ByVal frm_area As MSForms.Frame) As Single
' ------------------------------------------------------------------------------
' The maximum usable message section width depends on the maximum message area
' width whether or not the area frame (frm_artea) has a vertical scrollbar. A
' vertical scrollbar reduces the available spece by the required space for the
' vertical scrollbar.
' ------------------------------------------------------------------------------
    If frm_area.ScrollBars = fmScrollBarsVertical Or frm_area.ScrollBars = fmScrollBarsBoth _
    Then MaxWidthMsgSect = MaxWidthMsgArea - VSCROLLBAR_WIDTH _
    Else MaxWidthMsgSect = MaxWidthMsgArea
End Function

Private Function MaxWidthMsgTextBox(ByVal frm_text As MSForms.Frame) As Single
' ------------------------------------------------------------------------------
' The maximum with of a sections text-box depends on whether or not the frame of
' the TextBox (frm_text) has a vertical scrollbar which reduces the available
' space by its required width.
' ------------------------------------------------------------------------------
    If frm_text.ScrollBars = fmScrollBarsVertical Or frm_text.ScrollBars = fmScrollBarsBoth _
    Then MaxWidthMsgTextBox = frm_text.Width - VSCROLLBAR_WIDTH _
    Else MaxWidthMsgTextBox = frm_text.Width
End Function

Private Function MaxWidthMsgTextFrame( _
                           ByVal frm_area As MSForms.Frame, _
                           ByVal frm_section As MSForms.Frame) As Single
' ------------------------------------------------------------------------------
' The maximum usable message text width depends on the maximum message section
' width and whether or not the section (frm_section) has a vertical scrollbar
' which reduces the available space by its required width.
' ------------------------------------------------------------------------------
    If frm_section.ScrollBars = fmScrollBarsVertical Or frm_section.ScrollBars = fmScrollBarsBoth _
    Then MaxWidthMsgTextFrame = MaxWidthMsgSect(frm_area) - VSCROLLBAR_WIDTH _
    Else MaxWidthMsgTextFrame = MaxWidthMsgSect(frm_area)
End Function

Private Function Min(ParamArray va() As Variant) As Variant
' ------------------------------------------------------
' Returns the minimum (smallest) of all provided values.
' ------------------------------------------------------
    Dim v   As Variant
    
    Min = va(LBound(va)): If LBound(va) = UBound(va) Then Exit Function
    For Each v In va
        If v < Min Then Min = v
    Next v
    
End Function

Public Sub Monitor( _
              ByVal mntr_text As String, _
     Optional ByVal mntr_append As Boolean = True, _
     Optional ByVal mntr_footer As String)
' ------------------------------------------------------------------------------
' Replaces the MessageForms first section with the provided text (mntr_text) or
' appends it when (mntr_append) = True.
' ------------------------------------------------------------------------------
    
    UsageType = usage_progress_display
    
    If MsgText(1).MonoSpaced _
    Then SetupMsgSectMonoSpaced msg_section:=1 _
                              , msg_append:=mntr_append _
                              , msg_text:=mntr_text _
    Else SetupMsgSectPropSpaced msg_section:=1 _
                              , msg_append:=mntr_append _
                              , msg_text:=mntr_text
    SetupMsgSectPropSpaced msg_section:=2 _
                              , msg_text:=mntr_footer
    
    SizeAndPosition1MsgSects
    SizeAndPosition3Areas

    '~~ When the message form height exceeds the specified or the default message height
    '~~ height reduction and application of vertical scrollbars is due. The message area
    '~~ or the buttons area or both will be reduced to meet the limit and a vertical
    '~~ scrollbar will be setup. When both areas are about the same height (neither is
    '~~ taller the than 60% of the total heigth, both will get a vertical scrollbar,
    '~~ else only the one which uses 60% or more of the height.
    ScrollVerticalWhereApplicable
    SizeAndPosition1MsgSects
    SizeAndPosition3Areas

End Sub

Public Sub PositionMessageOnScreen( _
           Optional ByVal pos_top_left As Boolean = False)
' ------------------------------------------------------------------------------
'
' ------------------------------------------------------------------------------
    
    On Error Resume Next
        
    With Me
        .StartupPosition = sup_Manual
        If pos_top_left Then
            .Left = 5
            .top = 5
        Else
            .Left = (VirtualScreenWidthPts - .Width) / 2
            .top = (VirtualScreenHeightPts - .Height) / 4
        End If
    End With
    
    '~~ First make sure the bottom right fits,
    '~~ then check if the top-left is still on the screen (which gets priority).
    With Me
        If ((.Left + .Width) > (VirtualScreenLeftPts + VirtualScreenWidthPts)) Then .Left = ((VirtualScreenLeftPts + VirtualScreenWidthPts) - .Width)
        If ((.top + .Height) > (VirtualScreenTopPts + VirtualScreenHeightPts)) Then .top = ((VirtualScreenTopPts + VirtualScreenHeightPts) - .Height)
        If (.Left < VirtualScreenLeftPts) Then .Left = VirtualScreenLeftPts
        If (.top < VirtualScreenTopPts) Then .top = VirtualScreenTopPts
    End With
    
End Sub

'Private Sub ProvideCollection(ByRef cll As Collection)
'' ----------------------------------------------------
'' Provides a clean/new Collection.
'' ----------------------------------------------------
'    If Not cll Is Nothing Then Set cll = Nothing
'    Set cll = New Collection
'End Sub

Private Sub ProvideDictionary(ByRef dct As Dictionary)
' ----------------------------------------------------
' Provides a clean or new Dictionary.
' ----------------------------------------------------
    If Not dct Is Nothing Then dct.RemoveAll Else Set dct = New Dictionary
End Sub

Private Function ScrollHorizontalApplied(ByRef frm As MSForms.Frame) As Boolean
' ------------------------------------------------------------------------------
' Returns True when the frame (frm) has already a horizontal scrollbar applied.
' ------------------------------------------------------------------------------
    Select Case frm.ScrollBars
        Case fmScrollBarsBoth, fmScrollBarsHorizontal: ScrollHorizontalApplied = True
    End Select
End Function

Private Sub ScrollHorizontalApply( _
                            ByRef scroll_frame As MSForms.Frame, _
                            ByVal scrolled_width As Single, _
                   Optional ByVal x_action As fmScrollAction = fmScrollActionBegin)
' ------------------------------------------------------------------------------
' Apply a horizontal scrollbar is applied to the frame (scroll_frame) and
' adjusted to the frame content's width (scrolled_width). In case a horizontal
' scrollbar is already applied only its width is adjusted.
' ------------------------------------------------------------------------------
    Const PROC = "ScrollHorizontalApply"
    
    On Error GoTo eh
        
    With scroll_frame
        Select Case .ScrollBars
            Case fmScrollBarsBoth
                '~~ The already displayed horizonzal scrollbar's width is adjusted
                .KeepScrollBarsVisible = fmScrollBarsBoth
                .ScrollWidth = scrolled_width
                .Scroll xAction:=x_action
            Case fmScrollBarsHorizontal
                '~~ Already displayed (no vertical scrollbar yet)
                '~~ No need to adjust the height for the scrollbar
                .KeepScrollBarsVisible = fmScrollBarsHorizontal
                .ScrollWidth = scrolled_width
                .Scroll xAction:=x_action
                .Height = FrameContentHeight(scroll_frame) + HSCROLLBAR_HEIGHT
            Case fmScrollBarsVertical
                '~~ Add a horizontal scrollbar to the already displayed vertical
                .ScrollBars = fmScrollBarsBoth
                .KeepScrollBarsVisible = fmScrollBarsBoth
                .ScrollWidth = scrolled_width
                .Scroll xAction:=x_action
            Case fmScrollBarsNone
                '~~ Add a horizontal scrollbar
                .ScrollBars = fmScrollBarsHorizontal
                .KeepScrollBarsVisible = fmScrollBarsHorizontal
                .ScrollWidth = scrolled_width
                .Scroll xAction:=x_action
                .Height = FrameContentHeight(scroll_frame) + HSCROLLBAR_HEIGHT
        End Select
    End With

xt: Exit Sub
    
eh: If ErrMsg(ErrSrc(PROC)) = vbYes Then: Stop: Resume
End Sub

Private Function ScrollHorizontalHeight(ByVal frm As MSForms.Frame) As Single
    If ScrollHorizontalApplied(frm) Then ScrollHorizontalHeight = 12
End Function

Private Function ScrollVerticalApplied(ByRef frm As MSForms.Frame) As Boolean
' ------------------------------------------------------------------------------
' Returns True when the frame (frm) has already a vertical scrollbar applied.
' ------------------------------------------------------------------------------
    Select Case frm.ScrollBars
        Case fmScrollBarsBoth, fmScrollBarsVertical: ScrollVerticalApplied = True
    End Select
End Function

Private Sub ScrollVerticalApply( _
                          ByRef scroll_frame As MSForms.Frame, _
                          ByVal scrolled_height As Single, _
                 Optional ByVal y_action As fmScrollAction = fmScrollActionBegin)
' ------------------------------------------------------------------------------
' A vertical scrollbar is applied to the frame (scroll_frame) and adjusted to
' the frame content's height (scrolled_height). In case a vertical scrollbar is
' already applied only its width is adjusted.
' ------------------------------------------------------------------------------
    Const PROC = "ScrollVerticalApply"
    
    On Error GoTo eh
        
    With scroll_frame
        Select Case .ScrollBars
            Case fmScrollBarsBoth
                '~~ The already displayed horizonzal scrollbar's width is adjusted
                .KeepScrollBarsVisible = fmScrollBarsBoth
                .ScrollHeight = scrolled_height
                .Scroll yAction:=y_action
            Case fmScrollBarsHorizontal
                '~~ Already displayed (no vertical scrollbar yet)
                '~~ No need to adjust the height for the scrollbar
                .ScrollBars = fmScrollBarsBoth
                .KeepScrollBarsVisible = fmScrollBarsBoth
                .ScrollHeight = scrolled_height
                .Scroll yAction:=y_action
                .Width = FrameContentWidth(scroll_frame) + VSCROLLBAR_WIDTH
            Case fmScrollBarsVertical
                '~~ Add a horizontal scrollbar to the already displayed vertical
                .KeepScrollBarsVisible = fmScrollBarsVertical
                .ScrollHeight = scrolled_height
                .Scroll yAction:=y_action
            Case fmScrollBarsNone
                '~~ Add a horizontal scrollbar
                .ScrollBars = fmScrollBarsVertical
                .KeepScrollBarsVisible = fmScrollBarsVertical
                .ScrollHeight = scrolled_height
                .Scroll yAction:=y_action
                .Width = FrameContentWidth(scroll_frame) + VSCROLLBAR_WIDTH
'                scroll_frame.Parent.Width = .Width
        End Select
    End With

xt: Exit Sub
    
eh: If ErrMsg(ErrSrc(PROC)) = vbYes Then: Stop: Resume
End Sub

Private Sub ScrollVerticalMsgSectionOrArea(ByVal exceeding_height As Single)
' ------------------------------------------------------------------------------
' Either because the message area occupies 60% or more of the total height or
' because both, the message area and the buttons area us about the same height,
' it - or only the section text occupying 65% or more - will be reduced by the
' exceeding height amount (exceeding_height) and will get a vertical scrollbar.
' ------------------------------------------------------------------------------
    Const PROC = "ScrollVerticalMsgSectionOrArea"
    
    On Error GoTo eh
    Dim MsgArea             As MSForms.Frame:   Set MsgArea = DsgnMsgArea
    Dim MsgSect             As MSForms.Frame
    Dim MsgSectTextFrame    As MSForms.Frame
    Dim MsgSectTextBox      As MSForms.TextBox
    Dim i                   As Long
    
    '~~ Find a/the message section text which occupies 65% or more of the message area's height,
    For i = 1 To DsgnMsgSectsTextFrame.Count
        Set MsgSect = DsgnMsgSect(i)
        Set MsgSectTextFrame = DsgnMsgSectTextFrame(i)
        Set MsgSectTextBox = DsgnMsgSectTextBox(i)
        If MsgSectTextFrame.Height >= MsgArea.Height * 0.65 _
        Or ScrollVerticalApplied(MsgSectTextFrame) Then
            ' ------------------------------------------------------------------------------
            ' There is a section which occupies 65% of the overall height or has already a
            ' vertical scrollbar applied. Assigning a new frame height applies a vertical
            ' scrollbar if none is applied yet or just adjusts the scrollbar's height to the
            ' frame's content height
            ' ------------------------------------------------------------------------------
            If UsageType = usage_progress_display Then
                FrameHeight(MsgSectTextFrame, fmScrollActionEnd) = MsgSectTextFrame.Height - exceeding_height
            Else
                If MsgSectTextFrame.Height - exceeding_height > 0 Then
                    FrameHeight(MsgSectTextFrame) = MsgSectTextFrame.Height - exceeding_height
                End If
            End If
            ' MsgSect.Width = MsgSectTextFrame.Width + 1
            GoTo xt
        End If
    Next i
    
    FrameHeight(MsgArea) = MsgArea.Height - exceeding_height ' will apply a vertical scrollbar
    FormWidth = MsgArea.Width
xt: Exit Sub

eh: If ErrMsg(ErrSrc(PROC)) = vbYes Then: Stop: Resume
End Sub

Private Sub ScrollVerticalWhereApplicable()
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
    Const PROC = "ScrollVerticalWhereApplicable"
    
    On Error GoTo eh
    Dim BttnsArea               As MSForms.Frame:   Set BttnsArea = DsgnBttnsArea
    Dim BttnsFrame              As MSForms.Frame:   Set BttnsFrame = DsgnBttnsFrame
    Dim MsgArea                 As MSForms.Frame:   Set MsgArea = DsgnMsgArea
    Dim TotalExceedingHeight    As Single
    
    '~~ When the message form's height exceeds the specified maximum height
    If Me.Height > siMsgHeightMax Then
        With Me
            TotalExceedingHeight = .Height - siMsgHeightMax
            .Height = siMsgHeightMax     '~~ Reduce the height to the max height specified
            
            If PrcntgHeightMsgArea >= 0.6 Then
                '~~ Either the message area as a whole or the dominating message section - if theres is any -
                '~~ will be height reduced and applied with a vertical scroll bar
                ScrollVerticalMsgSectionOrArea TotalExceedingHeight
            ElseIf PrcntgHeightBttnsArea >= 0.6 Then
                '~~ Only the buttons area will be reduced and applied with a vertical scrollbar.
                FrameHeight(BttnsArea) = BttnsArea.Height - TotalExceedingHeight
            Else
                '~~ Both, the message area and the buttons area will be
                '~~ height reduced proportionally and applied with a vertical scrollbar
                FrameHeight(MsgArea) = MsgArea.Height * PrcntgHeightMsgArea
                FrameHeight(BttnsArea) = BttnsArea.Height * PrcntgHeightBttnsArea
            End If
        End With
    End If ' height exceeds specified maximum
   
xt: Exit Sub
    
eh: If ErrMsg(ErrSrc(PROC)) = vbYes Then: Stop: Resume
End Sub

Private Function ScrollVerticalWidth(ByVal frm As MSForms.Frame) As Single
    If ScrollVerticalApplied(frm) Then ScrollVerticalWidth = 15
End Function

Public Sub Setup()
    Const PROC = "Setup"
    
    On Error GoTo eh
    Dim BttnsArea       As MSForms.Frame
    Dim ContentWidth    As Single
    
    CollectDesignControls
    Set BttnsArea = DsgnBttnsArea
            
    DisplayFramesWithCaptions bDsplyFrmsWthCptnTestOnly ' may be True for test purpose
    
    '~~ Start the setup as if there wouldn't be any message - which might be the case
    Me.StartupPosition = 2
    Me.Height = 200                             ' just to start with - specifically for test purpose
    Me.Width = siMsgWidthMin
    
'    PositionMessageOnScreen pos_top_left:=True  ' in case of test best pos to start with
    DsgnMsgArea.Visible = False
    DsgnBttnsArea.top = VSPACE_AREAS
    
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
    
    '~~ Setup the reply buttons, the third element which potentially effects the final message width.
    '~~ In case the widest buttons row exceeds the maximum width specified a horizontal scrollbar is applied.
    Setup3_Bttns vbuttons
    
    SizeAndPosition2Bttns1
    SizeAndPosition2Bttns2Rows
    SizeAndPosition2Bttns3Frame
    SizeAndPosition2Bttns4Area
    
    ' -----------------------------------------------------------------------------------------------
    ' At this point the form has reached its final width (all proportionally spaced message sections
    ' are adjusted to it). However, the message height is only final in case there are just buttons
    ' but no message. The setup of proportional spaced message sections determines the final message
    ' height. When it exeeds the maximum height specified one or two vertical scrollbars are applied.
    ' -----------------------------------------------------------------------------------------------
    Setup4_MsgSectsPropSpaced
        
    If IsApplied(DsgnMsgArea) Then SizeAndPosition1MsgSects
    SizeAndPosition3Areas
            
    ' -----------------------------------------------------------------------------------------------
    ' When the message form height exceeds the specified or the default message height the height of
    ' the message area and or the buttons area is reduced and a vertical is applied.
    ' When both areas are about the same height (neither is taller the than 60% of the total heigth)
    ' both will get a vertical scrollbar, else only the one which uses 60% or more of the height.
    ' -----------------------------------------------------------------------------------------------
    ScrollVerticalWhereApplicable
    If IsApplied(DsgnMsgArea) Then SizeAndPosition1MsgSects
    SizeAndPosition2Bttns4Area
    SizeAndPosition3Areas
    
    '~~ Final form width adjustment
    '~~ When the message area or the buttons area has a vertical scrollbar applied
    '~~ the scrollbar may not be visible when the width as a result exeeds the specified
    '~~ message form width. In order not to interfere again with the width of all content
    '~~ the message form width is extended (over the specified maximum) in order to have
    '~~ the vertical scrollbar visible
    FormWidth = Me.FormContentWidth + ScrollVerticalWidth(DsgnMsgArea)
'    Debug.Print "Me.FormContentWidth             : " & Me.FormContentWidth
'    Debug.Print "ScrollVerticalWidth(DsgnMsgArea): " & ScrollVerticalWidth(DsgnMsgArea)
'    Debug.Print "Me.Width                        : " & Me.Width
    PositionMessageOnScreen
    SetUpDone = True ' To indicate for the Activate event that the setup had already be done beforehand
    
xt: Exit Sub

eh: If ErrMsg(ErrSrc(PROC)) = vbYes Then: Stop: Resume
End Sub

Public Sub Setup1_Title( _
                ByVal setup_title As String, _
                ByVal setup_width_min As Single, _
                ByVal setup_width_max As Single)
' ------------------------------------------------------------------------------
' Setup the message form for the provided title (setup_title) optimized with the
' provided minimum width (setup_width_min) and the provided maximum width
' (setup_width_max) by using a certain factor (setup_factor) for the calculation
' of the width required to display an untruncated title - as long as the maximum
' widht is not exeeded.
' ------------------------------------------------------------------------------
    Const PROC = "Setup1_Title"
    Const FACTOR = 1.5
    
    On Error GoTo eh
    Dim Correction    As Single
    
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
        Correction = (CInt(.laMsgTitle.Width)) / 2800
        .Width = Min(setup_width_max, .laMsgTitle.Width * (FACTOR - Correction))
        .Width = Max(.Width, setup_width_min)
        TitleWidth = .Width
    End With
    bDoneTitle = True
    
xt: Exit Sub
    
eh: If ErrMsg(ErrSrc(PROC)) = vbYes Then: Stop: Resume
End Sub

Private Sub Setup2_MsgSectsMonoSpaced()
    Const PROC = "Setup2_MsgSectsMonoSpaced"
    
    On Error GoTo eh
    Dim i       As Long
    Dim Message As TypeMsgText
    
    For i = 1 To Me.NoOfDesignedMsgSects
        Message = Me.MsgText(i)
        If Message.MonoSpaced And Message.Text <> vbNullString Then
            SetupMsgSect i
        End If
    Next i
    bDoneMonoSpacedSects = True
    
xt: Exit Sub
    
eh: If ErrMsg(ErrSrc(PROC)) = vbYes Then: Stop: Resume
End Sub

Private Sub Setup3_Bttns(ByVal vbuttons As Variant)
' --------------------------------------------------------------------------------------
' Setup and position the applied reply buttons and calculate the max reply button width.
' Note: When the provided vButtons argument is a string it wil be converted into a
'       collection and the procedure is performed recursively with it.
' --------------------------------------------------------------------------------------
    Const PROC = "Setup3_Bttns"
    
    On Error GoTo eh
    Dim BttnsArea As MSForms.Frame:   Set BttnsArea = DsgnBttnsArea
    
    AppliedControls = BttnsArea
    AppliedControls = DsgnBttnsFrame
    lSetupRows = 1
    
    '~~ Setup all reply button by calculatig their maximum width and height
    Select Case TypeName(vbuttons)
        Case "Long":        SetupBttnsFromValue vbuttons ' buttons are specified by one single VBA.MsgBox button value only
        Case "String":      SetupBttnsFromString vbuttons
        Case "Collection":  SetupBttnsFromCollection vbuttons
        Case "Dictionary":  SetupBttnsFromCollection vbuttons
        Case Else
            '~~ Because vbuttons is not provided by a known/accepted format
            '~~ the message will be setup with an Ok only button", vbExclamation
            Setup3_Bttns vbOKOnly
    End Select
    
xt: Exit Sub
    
eh: If ErrMsg(ErrSrc(PROC)) = vbYes Then: Stop: Resume
End Sub

Private Sub Setup4_MsgSectsPropSpaced()
    Const PROC = "Setup4_MsgSectsPropSpaced"
    
    On Error GoTo eh
    Dim i       As Long
    Dim Message As TypeMsgText
    
    For i = 1 To Me.NoOfDesignedMsgSects
        Message = MsgText(i)
        If Not Message.MonoSpaced And Message.Text <> vbNullString Then
            SetupMsgSect i
        End If
    Next i
    bDonePropSpacedSects = True
    bDoneMsgArea = True
    
xt: Exit Sub
    
eh: If ErrMsg(ErrSrc(PROC)) = vbYes Then: Stop: Resume
End Sub

Private Sub SetupBttnsFromCollection(ByVal cllButtons As Collection)
' ---------------------------------------------------------------------
' Setup the reply buttons based on the comma delimited string of button
' captions and row breaks indicated by a vbLf, vbCr, or vbCrLf.
' ---------------------------------------------------------------------
    Const PROC = "SetupBttnsFromCollection"
    
    On Error GoTo eh
    Dim v           As Variant
    Dim BttnsArea   As MSForms.Frame
    Dim BttnsFrame  As MSForms.Frame
    Dim BttnRow     As MSForms.Frame
    Dim Bttn        As MSForms.CommandButton

    lSetupRows = 1
    lSetupRowButtons = 0
    Set BttnsArea = DsgnBttnsArea
    Set BttnsFrame = DsgnBttnsFrame
    Set BttnRow = DsgnBttnRow(1)
    Set Bttn = DsgnBttn(1, 1)
    
    Me.Height = 100 ' just to start with
    BttnsArea.top = VSPACE_AREAS
    BttnsFrame.top = BttnsArea.top
    BttnRow.top = BttnsFrame.top
    Bttn.top = BttnRow.top
    Bttn.Width = DFLT_BTTN_MIN_WIDTH
    
    For Each v In cllButtons
        Select Case v
            Case vbOKOnly
                SetupBttnsFromValue v
            Case vbOKCancel, vbYesNo, vbRetryCancel
                SetupBttnsFromValue v
            Case vbYesNoCancel, vbAbortRetryIgnore
                SetupBttnsFromValue v
            Case vbYesNo
                SetupBttnsFromValue v
            Case Else
                If v <> vbNullString Then
                    If v = vbLf Or v = vbCr Or v = vbCrLf Then
                        '~~ prepare for the next row
                        If lSetupRows <= 7 Then ' ignore exceeding rows
                            AppliedControls(lSetupRows) = DsgnBttnRow(lSetupRows)
                            lSetupRows = lSetupRows + 1
                            lSetupRowButtons = 0
                        Else
                            MsgBox "Setup of button row " & lSetupRows & " ignored! The maximum applicable rows is 7."
                        End If
                    Else
                        lSetupRowButtons = lSetupRowButtons + 1
                        If lSetupRowButtons <= 7 Then
                            SetupButton ButtonRow:=lSetupRows, buttonindex:=lSetupRowButtons, buttoncaption:=v, buttonreturnvalue:=v
                        Else
                            MsgBox "The setup of a button " & lSetupRowButtons & " in row " & lSetupRows & " is ignored! The maximum applicable buttons per row is 7."
                        End If
                    End If
                End If
        End Select
    Next v
    If lSetupRows <= 7 Then
        AppliedControls(lSetupRows) = DsgnBttnRow(lSetupRows)
    End If
    
xt: Exit Sub

eh: If ErrMsg(ErrSrc(PROC)) = vbYes Then: Stop: Resume
End Sub

Private Sub SetupBttnsFromString(ByVal buttons_string As String)
    
    Dim cll As New Collection
    Dim v   As Variant
    
    For Each v In Split(buttons_string, ",")
        cll.Add v
    Next v
    Setup3_Bttns cll
    
End Sub

Private Sub SetupBttnsFromValue(ByVal lButtons As Long)
' -------------------------------------------------------
' Setup a row of standard VB MsgBox reply command buttons
' -------------------------------------------------------
    Const PROC = "SetupBttnsFromValue"
    
    On Error GoTo eh
    
    Select Case lButtons
        Case vbOKOnly
            lSetupRowButtons = lSetupRowButtons + 1
            SetupButton ButtonRow:=lSetupRows, buttonindex:=lSetupRowButtons, buttoncaption:="Ok", buttonreturnvalue:=vbOK
        Case vbOKCancel
            lSetupRowButtons = lSetupRowButtons + 1
            SetupButton ButtonRow:=lSetupRows, buttonindex:=lSetupRowButtons, buttoncaption:="Ok", buttonreturnvalue:=vbOK
            lSetupRowButtons = lSetupRowButtons + 1
            SetupButton ButtonRow:=lSetupRows, buttonindex:=lSetupRowButtons, buttoncaption:="Cancel", buttonreturnvalue:=vbCancel
        Case vbYesNo
            lSetupRowButtons = lSetupRowButtons + 1
            SetupButton ButtonRow:=lSetupRows, buttonindex:=lSetupRowButtons, buttoncaption:="Yes", buttonreturnvalue:=vbYes
            lSetupRowButtons = lSetupRowButtons + 1
            SetupButton ButtonRow:=lSetupRows, buttonindex:=lSetupRowButtons, buttoncaption:="No", buttonreturnvalue:=vbNo
        Case vbRetryCancel
            lSetupRowButtons = lSetupRowButtons + 1
            SetupButton ButtonRow:=lSetupRows, buttonindex:=lSetupRowButtons, buttoncaption:="Retry", buttonreturnvalue:=vbRetry
            lSetupRowButtons = lSetupRowButtons + 1
            SetupButton ButtonRow:=lSetupRows, buttonindex:=lSetupRowButtons, buttoncaption:="Cancel", buttonreturnvalue:=vbCancel
        Case vbYesNoCancel
            lSetupRowButtons = lSetupRowButtons + 1
            SetupButton ButtonRow:=lSetupRows, buttonindex:=lSetupRowButtons, buttoncaption:="Yes", buttonreturnvalue:=vbYes
            lSetupRowButtons = lSetupRowButtons + 1
            SetupButton ButtonRow:=lSetupRows, buttonindex:=lSetupRowButtons, buttoncaption:="No", buttonreturnvalue:=vbNo
            lSetupRowButtons = lSetupRowButtons + 1
            SetupButton ButtonRow:=lSetupRows, buttonindex:=lSetupRowButtons, buttoncaption:="Cancel", buttonreturnvalue:=vbCancel
        Case vbAbortRetryIgnore
            lSetupRowButtons = lSetupRowButtons + 1
            SetupButton ButtonRow:=lSetupRows, buttonindex:=lSetupRowButtons, buttoncaption:="Abort", buttonreturnvalue:=vbAbort
            lSetupRowButtons = lSetupRowButtons + 1
            SetupButton ButtonRow:=lSetupRows, buttonindex:=lSetupRowButtons, buttoncaption:="Retry", buttonreturnvalue:=vbRetry
            lSetupRowButtons = lSetupRowButtons + 1
            SetupButton ButtonRow:=lSetupRows, buttonindex:=lSetupRowButtons, buttoncaption:="Ignore", buttonreturnvalue:=vbIgnore
        Case Else
            MsgBox "The value provided for the ""buttons"" argument is not a known VB MsgBox value"
    End Select
    AppliedControls(lSetupRows) = DsgnBttnRow(lSetupRows)
    AppliedControls = DsgnBttnsFrame
    
xt: Exit Sub

eh: If ErrMsg(ErrSrc(PROC)) = vbYes Then: Stop: Resume
End Sub

Private Sub SetupButton(ByVal ButtonRow As Long, _
                        ByVal buttonindex As Long, _
                        ByVal buttoncaption As String, _
                        ByVal buttonreturnvalue As Variant)
' -----------------------------------------------------------------
' Setup an applied reply buttonindex's (buttonindex) visibility and
' caption, calculate the maximum buttonindex width and height,
' keep a record of the setup reply buttonindex's return value.
' -----------------------------------------------------------------
    Const PROC = "SetupButton"
    
    On Error GoTo eh
    Dim cmb As MSForms.CommandButton:   Set cmb = DsgnBttn(ButtonRow, buttonindex)
    
    With cmb
        .AutoSize = True
        .WordWrap = False ' the longest line determines the buttonindex's width
        .Caption = buttoncaption
        .AutoSize = False
        .Height = .Height + 1 ' safety margin to ensure proper multilin caption display
        siMaxButtonHeight = Max(siMaxButtonHeight, .Height)
        siMaxButtonWidth = Max(siMaxButtonWidth, .Width, siMinButtonWidth)
    End With
    AppliedBttns.Add cmb, ButtonRow
    AppliedButtonRetVal(cmb) = buttonreturnvalue ' keep record of the setup buttonindex's reply value
    AppliedControls = cmb
    AppliedControls(ButtonRow) = DsgnBttnRow(ButtonRow)
    
xt: Exit Sub
    
eh: If ErrMsg(ErrSrc(PROC)) = vbYes Then: Stop: Resume
End Sub

Private Sub SetupMsgSect(ByVal msg_section As Long)
' -------------------------------------------------------------
' Setup a message section with its label when one is specified
' and return the message's width when greater than any other.
' Note: All height adjustments except the one for the text box
'       are done by the SizeAndPosition
' -------------------------------------------------------------
    Const PROC = "SetupMsgSect"
    
    On Error GoTo eh
    Dim SectMessage      As TypeMsgText:     SectMessage = Me.MsgText(msg_section)
    Dim SectLabel        As TypeMsgLabel:    SectLabel = Me.MsgLabel(msg_section)
    Dim AreaFrame           As MSForms.Frame:   Set AreaFrame = DsgnMsgArea
    Dim MsgSect     As MSForms.Frame:   Set MsgSect = DsgnMsgSect(msg_section)
    Dim la                  As MSForms.Label:   Set la = DsgnMsgSectLabel(msg_section)
    Dim MsgSectTextBox   As MSForms.TextBox: Set MsgSectTextBox = DsgnMsgSectTextBox(msg_section)
    Dim MsgSectTextFrame As MSForms.Frame:   Set MsgSectTextFrame = DsgnMsgSectTextFrame(msg_section)
        
    MsgSect.Width = AreaFrame.Width
    la.Width = MsgSect.Width
    MsgSectTextFrame.Width = MsgSect.Width
    MsgSectTextBox.Width = MsgSect.Width
        
    If SectMessage.Text <> vbNullString Then
    
        AppliedControls = AreaFrame
        AppliedControls(msg_section) = MsgSect
        AppliedControls(msg_section) = MsgSectTextFrame
        AppliedControls(msg_section) = MsgSectTextBox
                
        If SectLabel.Text <> vbNullString Then
            Set la = DsgnMsgSectLabel(msg_section)
            With la
                .Width = Me.InsideWidth - (siHmarginFrames * 2)
                .Caption = SectLabel.Text
                With .Font
                    If SectLabel.MonoSpaced Then
                        If SectLabel.FontName <> vbNullString Then .Name = SectLabel.FontName Else .Name = DFLT_LBL_MONOSPACED_FONT_NAME
                        If SectLabel.FontSize <> 0 Then .Size = SectLabel.FontSize Else .Size = DFLT_LBL_MONOSPACED_FONT_SIZE
                    Else
                        If SectLabel.FontName <> vbNullString Then .Name = SectLabel.FontName Else .Name = DFLT_LBL_PROPSPACED_FONT_NAME
                        If SectLabel.FontSize <> 0 Then .Size = SectLabel.FontSize Else .Size = DFLT_LBL_PROPSPACED_FONT_SIZE
                    End If
                    If SectLabel.FontItalic Then .Italic = True
                    If SectLabel.FontBold Then .Bold = True
                    If SectLabel.FontUnderline Then .Underline = True
                End With
                If SectLabel.FontColor <> 0 Then .ForeColor = SectLabel.FontColor Else .ForeColor = rgbBlack
            End With
            MsgSectTextFrame.top = la.top + la.Height
            AppliedControls(msg_section) = la
        Else
            MsgSectTextFrame.top = 0
        End If
        
        If SectMessage.MonoSpaced Then
            SetupMsgSectMonoSpaced msg_section  ' returns the maximum width required for monospaced section
        Else ' proportional spaced
            SetupMsgSectPropSpaced msg_section
        End If
        MsgSectTextBox.SelStart = 0
        
    End If
    
xt: Exit Sub

eh: If ErrMsg(ErrSrc(PROC)) = vbYes Then: Stop: Resume
End Sub

Private Sub SetupMsgSectMonoSpaced( _
                             ByVal msg_section As Long, _
                    Optional ByVal msg_append As Boolean = False, _
                    Optional ByVal msg_append_margin As String = vbNullString, _
                    Optional ByVal msg_text As String = vbNullString)
' ------------------------------------------------------------------------------
' Setup the provided monospaced message section (msg_section). When a text is
' explicitely provided (msg_text) setup the sectiuon with this one, else with
' the property MsgText content. When an explicit text is provided the text
' either replaces the text, which the default or the text is appended when
' (msg_appen = True).
' Note 1: All top and height adjustments - except the one for the text box
'         itself are finally done by SizeAndPosition services when all
'         elements had been set up.
' Note 2: The optional arguments (msg_append) and (msg_text) are used with the
'         Monitor service which ma replace or add the provided text
' ------------------------------------------------------------------------------
Const PROC = "SetupMsgSectMonoSpaced"
    
    On Error GoTo eh
    Dim MsgSectText         As TypeMsgText
    Dim AreaFrame           As MSForms.Frame:   Set AreaFrame = DsgnMsgArea
    Dim MsgSectTextFrame    As MSForms.Frame:   Set MsgSectTextFrame = DsgnMsgSectTextFrame(msg_section)
    Dim MsgSectTextBox      As MSForms.TextBox: Set MsgSectTextBox = DsgnMsgSectTextBox(msg_section)
    Dim MsgSect        As MSForms.Frame:   Set MsgSect = DsgnMsgSect(msg_section)
    Dim MaxWidthAreaFrame   As Single:          MaxWidthAreaFrame = FormWidthMaxUsable - 4
    Dim MaxWidthSectFrame   As Single:          MaxWidthSectFrame = MaxWidthAreaFrame
    Dim MaxWidthTextFrame   As Single:          MaxWidthTextFrame = MaxWidthSectFrame
    Dim TextBoxValue        As String
    
    '~~ Keep record of the controls which had been applied
    AppliedControls(msg_section) = AreaFrame
    AppliedControls(msg_section) = MsgSect
    AppliedControls(msg_section) = MsgSectTextFrame:    MonoSpaced(MsgSectTextFrame) = True
    AppliedControls(msg_section) = MsgSectTextBox:      MonoSpaced(MsgSectTextBox) = True:  MonoSpacedTbx(MsgSectTextBox) = True
    
    If msg_text = vbNullString Then MsgSectText = MsgText(msg_section) Else MsgSectText.Text = msg_text

    With MsgSectTextBox
        With .Font
            If MsgSectText.FontName <> vbNullString Then .Name = MsgSectText.FontName Else .Name = DFLT_LBL_MONOSPACED_FONT_NAME
            If MsgSectText.FontSize <> 0 Then .Size = MsgSectText.FontSize Else .Size = DFLT_LBL_MONOSPACED_FONT_SIZE
            If .Bold <> MsgSectText.FontBold Then .Bold = MsgSectText.FontBold
            If .Italic <> MsgSectText.FontItalic Then .Italic = MsgSectText.FontItalic
            If .Underline <> MsgSectText.FontUnderline Then .Underline = MsgSectText.FontUnderline
        End With
        If .ForeColor <> MsgSectText.FontColor And MsgSectText.FontColor <> 0 Then .ForeColor = MsgSectText.FontColor
    End With
    
    AutoSizeTextBox as_tbx:=MsgSectTextBox _
                  , as_text:=MsgSectText.Text _
                  , as_width_limit:=0 _
                  , as_append:=msg_append _
                  , as_append_margin:=msg_append_margin
    
    With MsgSectTextBox
        .SelStart = 0
        .Left = siHmarginFrames
        MsgSectTextFrame.Left = siHmarginFrames
        MsgSectTextFrame.Height = .top + .Height
    End With ' MsgSectTextBox
        
    '~~ The width may expand or shrink depending on the change of the displayed text
    '~~ However, it cannot expand beyond the maximum width calculated for the text frame
    FrameWidth(MsgSectTextFrame) = Min(MaxWidthTextFrame, MsgSectTextBox.Width)
    MsgSect.Width = Min(MaxWidthSectFrame, MsgSectTextFrame.Width)
    AreaFrame.Width = Min(MaxWidthSectFrame, MsgSect.Width)
    FormWidth = AreaFrame.Width
                    
xt: Exit Sub
    
eh: If ErrMsg(ErrSrc(PROC)) = vbYes Then: Stop: Resume
End Sub

Private Sub SetupMsgSectPropSpaced( _
                             ByVal msg_section As Long, _
                    Optional ByVal msg_append As Boolean = False, _
                    Optional ByVal msg_append_marging As String = vbNullString, _
                    Optional ByVal msg_text As String = vbNullString)
' ------------------------------------------------------------------------------
' Setup the provided section (msg_section) proportional spaced. When a text is
' explicitely provided (msg_text) setup the sectiuon with this one, else with
' the property MsgText content. When an explicit text is provided the text
' either replaces the text, which the default or the text is appended when
' (msg_appen = True).
' Note 1: When this proportional spaced section is setup the message width is
'         regarded final. However, top and height adjustments - except the one
'         for the text box itself are finally done by SizeAndPosition
'         services when all elements had been set up.
' Note 2: The optional arguments (msg_append) and (msg_text) are used with the
'         Monitor service which ma replace or add the provided text
' ------------------------------------------------------------------------------
    
    Dim MsgSectText         As TypeMsgText
    Dim MsgArea             As MSForms.Frame:   Set MsgArea = DsgnMsgArea
    Dim MsgSect             As MSForms.Frame:   Set MsgSect = DsgnMsgSect(msg_section)
    Dim MsgSectTextFrame    As MSForms.Frame:   Set MsgSectTextFrame = DsgnMsgSectTextFrame(msg_section)
    Dim MsgSectTextBox      As MSForms.TextBox: Set MsgSectTextBox = DsgnMsgSectTextBox(msg_section)
    
    '~~ For the setup of proportional spaced message sections the message from width is regarded final
    Dim MaxWidthSectFrame   As Single:          MaxWidthSectFrame = Me.InsideWidth - 2
    Dim MaxWidthTextFrame   As Single:          MaxWidthTextFrame = MaxWidthSectFrame
    Dim MaxWidthTextBox     As Single:          MaxWidthTextBox = MaxWidthTextFrame
    
    AppliedControls(msg_section) = MsgArea
    AppliedControls(msg_section) = MsgSect
    AppliedControls(msg_section) = MsgSectTextFrame
    AppliedControls(msg_section) = MsgSectTextBox

    '~~ For proportional spaced message sections the width is determined by the area width
    MsgArea.Width = MaxWidthMsgArea
    If msg_text = vbNullString Then MsgSectText = MsgText(msg_section) Else MsgSectText.Text = msg_text
    
    With MsgSectTextBox
        With .Font
            If MsgSectText.FontName <> vbNullString Then .Name = MsgSectText.FontName Else .Name = DFLT_LBL_PROPSPACED_FONT_NAME
            If MsgSectText.FontSize <> 0 Then .Size = MsgSectText.FontSize Else .Size = DFLT_LBL_PROPSPACED_FONT_SIZE
            If .Bold <> MsgSectText.FontBold Then .Bold = MsgSectText.FontBold
            If .Italic <> MsgSectText.FontItalic Then .Italic = MsgSectText.FontItalic
            If .Underline <> MsgSectText.FontUnderline Then .Underline = MsgSectText.FontUnderline
        End With
        If .ForeColor <> MsgSectText.FontColor And MsgSectText.FontColor <> 0 Then .ForeColor = MsgSectText.FontColor
    End With
    
    AutoSizeTextBox as_tbx:=MsgSectTextBox _
                  , as_width_limit:=MaxWidthTextBox _
                  , as_text:=MsgSectText.Text _
                  , as_append:=msg_append _
                  , as_append_margin:=msg_append_marging
    
    With MsgSectTextBox
        .SelStart = 0
        .Left = HSPACE_LEFT
        DoEvents    ' to properly h-align the text
    End With
    
    MsgSectTextFrame.Height = MsgSectTextBox.top + MsgSectTextBox.Height
    MsgSect.Height = MsgSectTextFrame.top + MsgSectTextFrame.Height
    MsgArea.Height = FrameContentHeight(MsgArea)

End Sub

Private Sub SizeAndPosition1MsgSects()
' ------------------------------------------------------------------------------
' - Adjusts each applied section texts frame (textbox-frame, section-frame) in
'   accordance with its height and position the frames vertically
' - Adjust the applied lables' and the frames' top position in accordance to
'   their occupied height.
' - Consider vertical and horizontal scrollbars.
' ------------------------------------------------------------------------------
    Const PROC = "SizeAndPosition1MsgSects"
    
    On Error GoTo eh
    Dim i                   As Long
    Dim MsgArea             As MSForms.Frame:   Set MsgArea = DsgnMsgArea
    Dim MsgSect             As MSForms.Frame
    Dim MsgSectLabel        As MSForms.Label
    Dim MsgSectTextBox      As MSForms.TextBox
    Dim MsgSectTextFrame    As MSForms.Frame
    Dim TopForNextControl   As Single
    Dim TopNextSect         As Single
    Dim ContentHeight       As Single
    Dim ContentWidth        As Single
    
    TopNextSect = 6
    For i = 1 To cllDsgnMsgSects.Count
        TopForNextControl = 0
        If IsApplied(DsgnMsgSect(i)) Then
            Set MsgSect = DsgnMsgSect(i)
            Set MsgSectLabel = DsgnMsgSectLabel(i)
            Set MsgSectTextFrame = DsgnMsgSectTextFrame(i)
            Set MsgSectTextBox = DsgnMsgSectTextBox(i)
            
            '~~ Adjust the message section's label in case one is applied
            '~~ Note: The label's width cannot exceed the below txt-box's width
            If IsApplied(MsgSectLabel) Then
                With MsgSectLabel
                    .top = TopForNextControl
                    TopForNextControl = VgridPos(.top + .Height)
                    MsgSectLabel.Width = Me.Width - .Left - 5
                End With
            End If

            If IsApplied(MsgSectTextBox) Then
                MsgSectTextBox.top = siVmarginFrames
                With MsgSectTextFrame
                    .top = TopForNextControl
                    TopForNextControl = .top + .Height + siVmarginFrames
                End With
                
                '~~ Adjust the dimensions of message-text-frame considering possibly applied scrollbars
                If Not ScrollHorizontalApplied(MsgSectTextFrame) Then
                    MsgSectTextFrame.Width = MsgSectTextBox.Width + ScrollVerticalWidth(MsgSectTextFrame)
                End If
                If Not ScrollVerticalApplied(MsgSectTextFrame) Then
                    MsgSectTextFrame.Height = MsgSectTextBox.Height + ScrollHorizontalHeight(MsgSectTextFrame) - 4
                End If
                
                '~~ Adjust the dimensiona of the message-section-frame considering possibly applied scrollbars
                If Not ScrollHorizontalApplied(MsgSect) Then
                    MsgSect.Width = Me.Width - MsgSect.Left - 5
'                    MsgSect.Width = MsgSectTextFrame.Left + MsgSectTextFrame.Width + ScrollVerticalWidth(MsgSect)
                
                End If
                If Not ScrollVerticalApplied(MsgSect) Then
                    MsgSect.Height = MsgSectTextFrame.top + MsgSectTextFrame.Height + ScrollHorizontalHeight(MsgSect)
                End If
               
                DoEvents
            End If
                        
            '~~ Adjust the section-frame's top position
            With MsgSect
                .top = TopNextSect
                DoEvents
                TopNextSect = VgridPos(.top + .Height + siVmarginFrames + VSPACE_SECTIONS) ' the next section if any
            End With

        End If ' IsApplied(MsgSect)
    Next i
    
    '~~ Adjust dimensions of the message-area-frame
    If Not ScrollHorizontalApplied(MsgArea) Then
        MsgArea.Width = FrameContentWidth(MsgArea) + ScrollVerticalWidth(MsgArea)
    End If
    If Not ScrollVerticalApplied(MsgArea) Then
        MsgArea.Height = FrameContentHeight(MsgArea) + ScrollHorizontalHeight(MsgArea)
    End If

xt: Exit Sub
    
eh: If ErrMsg(ErrSrc(PROC)) = vbYes Then: Stop: Resume
End Sub

Private Sub SizeAndPosition2Bttns1()
' ------------------------------------------------------------------------------
' Unify all applied/visible button's size by assigning the maximum width and
' height provided with their setup, and adjust their resulting left position.
' ------------------------------------------------------------------------------
    Const PROC = "SizeAndPosition2Bttns1"
    
    On Error GoTo eh
    Dim cllButtonRows   As Collection:      Set cllButtonRows = DsgnBttnRows
    Dim siLeft          As Single
    Dim frRow           As MSForms.Frame    ' Frame for the buttons in a row
    Dim vButton         As Variant
    Dim lRow            As Long
    Dim lButton         As Long
    
    For lRow = 1 To cllButtonRows.Count
        siLeft = HSPACE_LEFTRIGHT_BUTTONS
        Set frRow = cllButtonRows(lRow)
        If IsApplied(frRow) Then
            For Each vButton In DsgnRowBttns(lRow)
                If IsApplied(vButton) Then
                    lButton = lButton + 1
                    With vButton
                        .Left = siLeft
                        .Width = siMaxButtonWidth
                        .Height = siMaxButtonHeight
                        .top = siVmarginFrames
                        siLeft = .Left + .Width + siHmarginButtons
                        If IsNumeric(vMsgButtonDefault) Then
                            If lButton = vMsgButtonDefault Then .Default = True
                        Else
                            If .Caption = vMsgButtonDefault Then .Default = True
                        End If
                    End With
                End If
            Next vButton
        End If
        frRow.Width = frRow.Width + HSPACE_LEFTRIGHT_BUTTONS
    Next lRow
        
xt: Exit Sub
    
eh: If ErrMsg(ErrSrc(PROC)) = vbYes Then: Stop: Resume
End Sub

Private Sub SizeAndPosition2Bttns2Rows()
' ------------------------------------------------------------------------------
' Adjust all applied/visible button rows height to the maximum buttons height
' and the row frames width to the number of the displayed buttons considering a
' certain margin between the buttons (siHmarginButtons) and a margin at the
' left and the right.
' ------------------------------------------------------------------------------
    Const PROC = "SizeAndPosition2Bttns2Rows"
    
    On Error GoTo eh
    Dim BttnsFrame      As MSForms.Frame:   Set BttnsFrame = DsgnBttnsFrame
    Dim BttnRows        As Collection:      Set BttnRows = DsgnBttnRows
    Dim BttnRowFrame    As MSForms.Frame
    Dim siTop           As Single
    Dim v               As Variant
    Dim lButtons        As Long
    Dim siHeight        As Single
    Dim BttnsFrameWidth As Single
    Dim dct             As Dictionary:      Set dct = AppliedBttnRows
    Dim ContentWidth    As Single
    Dim ContentHeight   As Single
    
    '~~ Adjust button row's width and height
    siHeight = AppliedButtonRowHeight
    siTop = siVmarginFrames
    For Each v In dct
        Set BttnRowFrame = v
        lButtons = dct(v)
        If IsApplied(BttnRowFrame) Then
            With BttnRowFrame
                .top = siTop
                .Height = siHeight
                '~~ Provide some extra space for the button's design
                BttnsFrameWidth = CInt((siMaxButtonWidth * lButtons) _
                               + (siHmarginButtons * (lButtons - 1)) _
                               + (siHmarginFrames * 2)) - siHmarginButtons + 7
                .Width = BttnsFrameWidth + (HSPACE_LEFTRIGHT_BUTTONS * 2)
                siTop = .top + .Height + siVmarginButtons
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
    Dim AreaFrame       As MSForms.Frame: Set AreaFrame = DsgnBttnsArea
    Dim BttnsFrame      As MSForms.Frame: Set BttnsFrame = DsgnBttnsFrame
    Dim v               As Variant
    Dim ContentWidth    As Single
    Dim ContentHeight   As Single
    
    If IsApplied(BttnsFrame) Then
        ContentWidth = FrameContentWidth(BttnsFrame)
        ContentHeight = FrameContentHeight(BttnsFrame)
        With BttnsFrame
            .top = 0
            BttnsFrame.Height = ContentHeight
            BttnsFrame.Width = ContentWidth
            '~~ Center all button rows within the buttons frame
            For Each v In DsgnBttnRows
                If IsApplied(v) Then
                    FrameCenterHorizontal center_frame:=v, within_frame:=BttnsFrame
                End If
            Next v
        End With
    End If
    AreaFrame.Height = FrameContentHeight(BttnsFrame)

xt: Exit Sub
    
eh: If ErrMsg(ErrSrc(PROC)) = vbYes Then: Stop: Resume
End Sub

Private Sub SizeAndPosition2Bttns4Area()
' ------------------------------------------------------------------------------
' Adjust the buttons area frame in accordance with the buttons frame.
' ------------------------------------------------------------------------------
    Const PROC = "SizeAndPosition2Bttns4Area"
    
    On Error GoTo eh
    Dim BttnsArea       As MSForms.Frame:   Set BttnsArea = DsgnBttnsArea
    Dim BttnsFrame      As MSForms.Frame:   Set BttnsFrame = DsgnBttnsFrame
    Dim ContentHeight   As Single
    Dim ContentWidth    As Single
    
    ContentHeight = FrameContentHeight(BttnsArea)
    ContentWidth = FrameContentWidth(BttnsArea)
            
    FrameWidth(BttnsArea) = Min(ContentWidth, siMsgWidthMax)
    
    If Not ScrollHorizontalApplied(BttnsArea) Then
'        If Not ScrollHorizontalApplied(BttnsArea) Then
            BttnsArea.Width = BttnsFrame.Left + BttnsFrame.Width + ScrollVerticalWidth(BttnsArea)
'        End If
    End If
    
    If Not ScrollHorizontalApplied(BttnsArea) Then
        If Not ScrollVerticalApplied(BttnsArea) Then
            BttnsArea.Height = BttnsFrame.top + BttnsFrame.Height + ScrollHorizontalHeight(BttnsArea)
        End If
    End If
    
    FormWidth = BttnsArea.Width + ScrollVerticalWidth(BttnsArea)
    
    FrameCenterHorizontal center_frame:=BttnsArea, left_margin:=10
'    With BttnsArea
'        FormWidth = .Left + .Width
'    End With
    
xt: Exit Sub
    
eh: If ErrMsg(ErrSrc(PROC)) = vbYes Then: Stop: Resume
End Sub

Private Sub SizeAndPosition3Areas()
' ------------------------------------------------------------------------------
'
' ------------------------------------------------------------------------------
    Const PROC = "SizeAndPosition3Areas"
    
    On Error GoTo eh
    Dim TopNextArea As Single
    Dim AreaFrame   As MSForms.Frame
    Dim MsgArea     As MSForms.Frame: Set MsgArea = DsgnMsgArea
    Dim BttnsArea   As MSForms.Frame: Set BttnsArea = DsgnBttnsArea
    
    TopNextArea = siVmarginFrames
    If IsApplied(MsgArea) Then
        With MsgArea
            .top = TopNextArea
            TopNextArea = VgridPos(.top + .Height + VSPACE_AREAS)
        End With
        Set AreaFrame = MsgArea
    Else
        TopNextArea = 15
    End If
    
    If IsApplied(BttnsArea) Then
        With BttnsArea
            .top = TopNextArea
            TopNextArea = VgridPos(.top + .Height + VSPACE_AREAS)
        End With
        Set AreaFrame = BttnsArea
    End If
    
    '~~ Adjust the final height of the message form
    Me.Height = VgridPos(AreaFrame.top + AreaFrame.Height + VSPACE_BOTTOM)
            
xt: Exit Sub
    
eh: If ErrMsg(ErrSrc(PROC)) = vbYes Then: Stop: Resume
End Sub

Private Sub UserForm_Activate()
' ------------------------------------------------------------
' To avoid screen flicker the setup may has been done already.
' However for test purpose the Setup may run with the Activate
' event i.e. the .Show
' ------------------------------------------------------------
    If Not SetUpDone Then Setup
End Sub

Public Function VgridPos(ByVal si As Single) As Single
' --------------------------------------------------------------
' Returns an integer of which the remainder (Int(si) / 6) is 0.
' Background: A controls content is only properly displayed
' when the top position of it is aligned to such a position.
' --------------------------------------------------------------
    Dim i As Long
    
    For i = 0 To 6
        If Int(si) = 0 Then
            VgridPos = 0
        Else
            If Int(si) < 6 Then si = 6
            If (Int(si) + i) Mod 6 = 0 Then
                VgridPos = Int(si) + i
                Exit For
            End If
        End If
    Next i

End Function

