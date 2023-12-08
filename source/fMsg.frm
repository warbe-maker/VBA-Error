VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} fMsg 
   ClientHeight    =   14805
   ClientLeft      =   240
   ClientTop       =   390
   ClientWidth     =   15150
   OleObjectBlob   =   "fMsg.frx":0000
End
Attribute VB_Name = "fMsg"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
' -------------------------------------------------------------------------------
' UserForm fMsg: Provides all means for the setup of all message variants with
' ==============  up to 8 sections, either proportional- or mono-spaced, each
' with an optional Label positioned above or to the left of the section text, and
' up to 7 rows each with 7 reply buttons.
'
' Public Properties:
' ------------------
' - IndicateFrameCaptions Test option, indicated the frame names
' - MinButtonWidth
' - MsgTitle               The title displayed in the window handle bar
' - MinButtonWidth         Minimum button width in pt
' - MsgButtonDefault       The number of the default button
' - MsgBttns               Buttons to be displayed, Collection provided by the
'                          mMsg.Buttons service
' - FormHeightOutsideMax   Percentage of screen height
' - FormHeightOutsideMin   Percentage of screen height
' - MsgLabel               A section's Label
' - FormWidthInsideMax     Percentage of screen width
' - FormWidthInsideMin     Defaults to 400 pt. the absolute minimum is 200 pt
' - MsgText                A section's text or a monitor header, monitor footer
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
' W. Rauschenberger Berlin, Oct 2023
' --------------------------------------------------------------------------
Private Const DFLT_BTTN_MIN_WIDTH           As Single = 70              ' Default minimum reply button width
Private Const DFLT_LBL_MONOSPACED_FONT_NAME As String = "Courier New"   ' Default monospaced Font name
Private Const DFLT_LBL_MONOSPACED_FONT_SIZE As Single = 9               ' Default monospaced Font size
Private Const DFLT_LBL_PROPSPACED_FONT_NAME As String = "Calibri"       ' Default proportional spaced Font name
Private Const DFLT_LBL_PROPSPACED_FONT_SIZE As Single = 9               ' Default proportional spaced Font size
Private Const DFLT_TXT_MONOSPACED_FONT_NAME As String = "Courier New"   ' Default monospaced Font name
Private Const DFLT_TXT_MONOSPACED_FONT_SIZE As Single = 10              ' Default monospaced Font size
Private Const DFLT_TXT_PROPSPACED_FONT_NAME As String = "Tahoma"        ' Default proportional spaced Font name
Private Const DFLT_TXT_PROPSPACED_FONT_SIZE As Single = 10              ' Default proportional spaced Font size
Private Const H_MARGIN_BTTNS                As Single = 4               ' Horizontal space between reply buttons
Private Const SCROLL_V_WIDTH                As Single = 18              ' Additional horizontal space required for a frame with a vertical scrollbar
Private Const SCROLL_H_HEIGHT               As Single = 14              ' Additional vertical space required for a frame with a horizontal scroll barr
Private Const VSPACE_BTTN_ROWS              As Single = 5               ' Vertical space between button rows
Private Const VISUALIZE_CLR_AREA            As Long = &HC0E0FF          ' light orange Backcolors for the
Private Const VISUALIZE_CLR_MON_STEPS_FRM   As Long = &HFFFFC0          '              visualization
Private Const VISUALIZE_CLR_MSEC_FRM        As Long = &HFFFFC0          '              of controls
Private Const VISUALIZE_CLR_MSEC_LBL        As Long = rgbLightYellow    ' light yellow during
Private Const VISUALIZE_CLR_MSEC_TEXT_FRM   As Long = rgbLightGreen     ' light green  test (only!)
Private Const VISUALIZE_CLR_BTTNS_FRM       As Long = &H80C0FF          ' light green  test (only!)
Private Const VISUALIZE_CLR_BTTNS_ROW_FRM   As Long = &HC0FFFF          ' light yellow test (only!)
Private Const TEMP_TBX_NAME                 As String = "tbxTemp"
Private Const SCROLL_VER_THRESHOLD          As Single = 5
Private Const SCROLL_HOR_THRESHOLD          As Single = 5
' Means to get and calculate the display devices DPI in points
Private Const SM_XVIRTUALSCREEN             As Long = &H4C&
Private Const SM_YVIRTUALSCREEN             As Long = &H4D&
Private Const SM_CXVIRTUALSCREEN            As Long = &H4E&
Private Const SM_CYVIRTUALSCREEN            As Long = &H4F&
Private Const LOGPIXELSX                    As Long = 88
Private Const LOGPIXELSY                    As Long = 90
Private Const TWIPSPERINCH                  As Long = 1440
Private Const WIN_NORMAL                    As Long = 1             ' Shell Open Normal

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
Private Enum CursorTypes
    IDC_HAND = 32649
End Enum

'Needed for GetCursorInfo
Private Type POINT
    x As Long
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

Private Enum MSFormControls
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

Private Msg                                 As udtMsg           ' The message tranferred from the caller
Private AppliedBttns                        As Dictionary       ' Dictionary of applied buttons (key=CommandButton, item=row)
Private AppliedBttnsRetVal                  As Dictionary       ' Dictionary of the applied buttons' reply value (key=CommandButton)
Private bButtonsSetup                       As Boolean
Private bFormEvents                         As Boolean
Private bIndicateFrameCaptions              As Boolean
Private bModeLess                           As Boolean
Private bMonitorInitialized                 As Boolean
Private bReplyWithIndex                     As Boolean
Private bSetUpDone                          As Boolean
Private bSetupDoneMonoSpacedSects           As Boolean
Private bSetupDonePropSpacedSects           As Boolean
Private bSetupDoneTitle                     As Boolean
Private bVisualizeForTest                   As Boolean
Private cllDsgnBttnRows                     As Collection       ' Collection of the designed reply button row frames
Private cllDsgnRowBttns                     As Collection       ' Collection of a designed reply button row's buttons
Private cllMsectsActive                     As New Collection
Private cllMsgBttns                         As New Collection
Private cllSteps                            As Collection
Private cyTimerTicksBegin                   As Currency
Private cyTimerTicksEnd                     As Currency
Private dctApplicationRunArgs               As New Dictionary   ' Dictionary will be available with each instance of this UserForm
Private dctAreas                            As New Dictionary   ' Collection of the two primary/top frames
Private dctBttns                            As New Dictionary   ' Collection of the collection of the designed reply buttons of a certain row
Private dctBttnsRowFrm                      As New Dictionary   ' Established/created Button Row's Frame
Private dctMonoSpaced                       As New Dictionary
Private dctMonoSpacedTbx                    As New Dictionary
Private dctMsectFrm                         As New Dictionary   ' Established/created Message Sections Frame
Private dctMsectLabel                       As New Dictionary   ' Established/created Message Sections Label
Private dctMsectTextFrm                     As New Dictionary   ' Established/created Message Sections TextBox Frame
Private dctMsectTextLbl                     As New Dictionary   ' Established/created Message Sections Text
Private dctMsectTextTbx                     As New Dictionary   ' Established/created Message Sections TextBox
Private dctSectLabelSetup                   As New Dictionary
Private dctSectsLabel                       As New Dictionary   ' frmMsect specific Label either provided via properties MsgLabel or Msg
Private dctSectsMonoSpaced                  As New Dictionary   ' frmMsect specific monospace option either provided via properties MsgMonospaced or Msg
Private dctSectsText                        As New Dictionary
Private dctSectTextSetup                    As New Dictionary
Private dctVisualizedForTest                As New Dictionary
Private frmBarea                            As MsForms.Frame    ' The buttons area frame
Private frmBttnsFrm                         As MsForms.Frame    ' Set with CollectDesignControls
Private frmMarea                            As MsForms.Frame    ' The message area frame
Private frmMsect                            As MsForms.Frame    ' A message section's fram
Private frmMsectText                        As MsForms.Frame    ' A message section's TextBox frame
Private frmSteps                            As MsForms.Frame
Private lBackColor                          As Long
Private lblMsectText                        As MsForms.Label    ' Set with MsectItems for a certain section
Private lLabelAllPos                        As enLabelPos       ' "global" Label position
Private lMaxNoOfMsgSects                    As Long             ' Set with CollectDesignControls (number of message sections designed)
Private lMonitorStepsDisplayed              As Long
Private lSetupRowButtons                    As Long             ' number of buttons setup in a row
Private lSetupRows                          As Long             ' number of setup button rows
Private lStepsDisplayed                     As Long
Private MsgSectLbl                          As udtMsgLabel      ' Label section of the udtMsg UDT
Private MsgSectTxt                          As udtMsgText       ' Text section of the udtMsg UDT
Private siBareaFrmWidth                     As Single
Private siAreasFrmWidth                     As Single
Private siAreasFrmWidthMax                  As Single
Private siAreasFrmWidthMin                  As Single
Private siFormHeightInside                  As Single
Private siFormHeightInsideMax               As Single
Private siFormHeightInsideMin               As Single
Private siFormHeightOutsideMax              As Single           ' The maximum (default or specified) message height in pt
Private siFormHeightOutsideMin              As Single           ' The minimum (default or specified) message height in pt
Private siFormWidthInside                   As Single
Private siFormWidthInsideMax                As Single
Private siFormWidthInsideMin                As Single
Private siFormWidthOutside                  As Single
Private siFormWidthOutsideMax               As Single
Private siFormWidthOutsideMin               As Single
Private siHmarginFrames                     As Single           ' Test property, value defaults to 0
Private siMsectLabelWidthAll                As Single           ' "global" Label width spec
Private siMaxButtonHeight                   As Single
Private siMaxButtonWidth                    As Single
Private siMsectFrmWidth                     As Single
Private siMsectFrmWidthMax                  As Single
Private siMsectFrmWidthMin                  As Single
Private siMsectTextFrmWidthOnlyText         As Single
Private siMsectTextFrmWidthOnlyTextMax      As Single
Private siMsectTextFrmWidthOnlyTextMin      As Single
Private siMsectTextFrmWidthWithLposLbl      As Single
Private siMsectTextFrmWidthWithLposLblMax   As Single
Private siMsectTextFrmWidthWithLposLblMin   As Single
Private sMonitorProcess                     As String
Private sMsgTitle                           As String
Private tbxFooter                           As MsForms.TextBox
Private tbxHeader                           As MsForms.TextBox
Private tbxStep                             As MsForms.TextBox
Private tbxTemp                             As MsForms.TextBox
Private TextMonitorFooter                   As udtMsgText
Private TextMonitorHeader                   As udtMsgText
Private TextMonitorStep                     As udtMsgText
Private TimerSystemFrequency                As Currency
Private VirtualScreenHeightPts              As Single
Private VirtualScreenLeftPts                As Single
Private VirtualScreenTopPts                 As Single
Private VirtualScreenWidthPts               As Single
Private vMsgButtonDefault                   As Variant          ' Index or caption of the default button

Private Function MsectTextFrmWidthMax(ByVal m_sect As Long) As Single
    If MsectHasLabelLeft(m_sect) _
    Then MsectTextFrmWidthMax = siMsectTextFrmWidthWithLposLblMax _
    Else MsectTextFrmWidthMax = siMsectTextFrmWidthOnlyTextMax
End Function

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
    lMaxNoOfMsgSects = mMsg.NoOfMsgSects       ' Global definition !!!!!
    
xt: Exit Sub
    
eh: If ErrMsg(ErrSrc(PROC)) = vbYes Then: Stop: Resume
End Sub

Private Sub UserForm_Terminate()
    Set AppliedBttns = Nothing
    Set AppliedBttnsRetVal = Nothing
    Set dctAreas = Nothing
    Set cllDsgnBttnRows = Nothing
    Set dctBttns = Nothing
    Set dctMsectFrm = Nothing
    Set dctMsectLabel = Nothing
    Set dctMsectTextTbx = Nothing
    Set dctMsectTextFrm = Nothing
    Set cllDsgnRowBttns = Nothing
    Set dctMonoSpaced = Nothing
    Set dctMonoSpacedTbx = Nothing
    Set dctSectsLabel = Nothing
    Set dctSectsMonoSpaced = Nothing
    Set dctSectsText = Nothing
    If bModeLess Then
        Application.EnableEvents = True
    End If
End Sub

Public Property Let ApplicationRunArgs(ByVal dct As Dictionary):    Set dctApplicationRunArgs = dct:                                        End Property

Private Property Get AppliedButtonRetVal(Optional ByVal cmb As MsForms.CommandButton) As Variant
    AppliedButtonRetVal = AppliedBttnsRetVal(cmb)
End Property

Private Property Let AppliedButtonRetVal(Optional ByVal cmb As MsForms.CommandButton, _
                                                  ByVal v As Variant)
    AppliedBttnsRetVal.Add cmb, v
End Property

Private Property Get AreasFrmMarginBottom() As Single:              AreasFrmMarginBottom = 12:                                              End Property

Private Property Get AreasFrmMarginLeft() As Single:                AreasFrmMarginLeft = 4:                                                 End Property

Private Property Get AreasFrmMarginRight() As Single:               AreasFrmMarginRight = 4:                                                End Property

Private Property Get AreasFrmMarginTop() As Single:                 AreasFrmMarginTop = 12:                                              End Property

Private Property Get AreasFrmWidthMax() As Single:                  AreasFrmWidthMax = siAreasFrmWidthMax:                                  End Property

Private Property Get AreasFrmWidthMin() As Single:                  AreasFrmWidthMin = siAreasFrmWidthMin:                                  End Property

Private Property Get AreasMarginVertical() As Single:               AreasMarginVertical = 8:                                                End Property

Private Property Let BareaFrmWidth(ByVal b_si As Single)
' ------------------------------------------------------------------------------
' When the width expands, the width of all depending frames are adjusted.
' ------------------------------------------------------------------------------
    Dim si As Single
    If WidthExpanded(b_si, siAreasFrmWidthMin, siAreasFrmWidthMax, siBareaFrmWidth, si) Then
        siBareaFrmWidth = si
        siAreasFrmWidth = si
        FormWidthInside = si + FormMarginLeft + FormMarginRight ' triggers the others
    End If
    BareaFrmCenter
    
End Property

Private Property Get BttnRowMarginLeft() As Single:     BttnRowMarginLeft = 8:  End Property

Private Property Get BttnRowMarginRight() As Single:    BttnRowMarginRight = 8:  End Property

Private Property Get BttnsFrmMarginBottom() As Single:              BttnsFrmMarginBottom = 2:                                               End Property

Private Property Get BttnsFrmMarginTop() As Single:                 BttnsFrmMarginTop = 2:                                                  End Property

Private Property Let FormHeightInside(ByVal si As Single)
    siFormHeightInside = si
    Me.Height = FormHeightOutside
End Property

Public Property Get FormHeightInsideMax() As Single:                FormHeightInsideMax = siFormHeightInsideMax:                            End Property

Public Property Let FormHeightInsideMax(ByVal si As Single):        siFormHeightInsideMax = si:                                             End Property

Public Property Get FormHeightInsideMin() As Single:                FormHeightInsideMin = siFormHeightInsideMin:                            End Property

Public Property Let FormHeightInsideMin(ByVal si As Single):        siFormHeightInsideMin = si:                                             End Property

Private Property Get FormHeightOutside() As Single
    FormHeightOutside = siFormHeightInside + HeightDiffFormOutIn
End Property

Public Property Get FormHeightOutsideMax() As Single:               FormHeightOutsideMax = siFormHeightOutsideMax:                          End Property

Public Property Let FormHeightOutsideMax(ByVal si As Single)
    siFormHeightOutsideMax = si
    siFormHeightInsideMax = siFormHeightOutsideMax - HeightDiffFormOutIn
End Property

Private Property Get FormHeightOutsideMin() As Single:              FormHeightOutsideMin = siFormHeightOutsideMin:                          End Property

Public Property Let FormHeightOutsideMin(ByVal si As Single):       siFormHeightOutsideMin = si:                                            End Property

Private Property Get FormMarginBottom() As Single:                  FormMarginBottom = 8:                                                   End Property

Private Property Get FormMarginLeft() As Single:                    FormMarginLeft = 4:                                                     End Property

Private Property Get FormMarginRight() As Single:                   FormMarginRight = 4:                                                    End Property

Private Property Get FormMarginTop() As Single:                     FormMarginTop = 8:                                                      End Property

Private Property Get FormWidthDiffOutIn() As Single:                FormWidthDiffOutIn = CInt(Me.Width - Me.InsideWidth):                   End Property

Private Property Let FormWidthInside(Optional ByVal f_sect As Long = 0, _
                                              ByVal f_si As Single)
' ------------------------------------------------------------------------------
' The message form's width is directly triggered by the Title setup.
' When this setup expands the initial minimum width all subordinate frame widths
' are adjusted correspondingly.
' See function WidthExpanded for documentation of hte width subject.
' ------------------------------------------------------------------------------
    Dim si As Single
    
    If WidthExpanded(f_si, siFormWidthInsideMin, siFormWidthInsideMax, siFormWidthInside, si) Then
        siFormWidthInside = si ' has expanded
        '~~ Subordinate frame's width adjustments
        siAreasFrmWidth = si - FormMarginLeft - FormMarginRight
        siMsectFrmWidth = siAreasFrmWidth - AreasFrmMarginLeft - AreasFrmMarginRight
        siMsectTextFrmWidthOnlyText = siMsectFrmWidth - MsectFrmMarginLeft - MsectFrmMarginRight
        siMsectTextFrmWidthWithLposLbl = siMsectTextFrmWidthOnlyText - siMsectLabelWidthAll - MsectLabelTextMargin
        
        '~~ The final form's outside width assignment triggers the adjustment of all active subordinate frames
        FormWidthOutside = siFormWidthInside + FormWidthDiffOutIn
    End If
    If f_sect <> 0 Then
        If MsgHasText Then MareaFrm
        MsectFrm f_sect
        If Not MsectHasOnlyLabel(f_sect) Then MsectTextFrm f_sect
    End If
    If bButtonsSetup Then BareaFrm

End Property

Private Property Let FormWidthOutside(ByVal si As Single)
' ------------------------------------------------------------------------------
' The message form's outside width may expand along with the setup of the Title,
' the reply Buttons, and/or the setup of monospaced sections. In case, all
' subordinate active frames are expandend correspondingly.
' See function WidthExpanded for documentation of hte width subject.
' ------------------------------------------------------------------------------
    Dim v As Variant
    
    If si > siFormWidthOutside Then
        siFormWidthOutside = si
        Me.Width = si
        For Each v In MsectsActive
            MsectFrm(v).Width = siMsectFrmWidth
            If Not MsectHasOnlyLabel(v) Then
                If MsectTxtWithLftPosLbl(v) Then
                    MsectTextFrm(v).Width = siMsectTextFrmWidthWithLposLbl
                Else
                    MsectTextFrm(v).Width = siMsectTextFrmWidthOnlyText
                End If
            End If
        Next v
    End If
    
End Property

Public Property Get FormWidthOutsideMax() As Single ' Public for testing purpose only!
    FormWidthOutsideMax = siFormWidthOutsideMax
End Property

Public Property Let FormWidthOutsideMax(ByVal si As Single)
' ------------------------------------------------------------------------------
' The maximum outside message form's width is assigned during initialization.
' This determines the maximum width of all subsequent frames throughout all
' setups.
' See function WidthExpanded for documentation of hte width subject.
' ------------------------------------------------------------------------------
    siFormWidthOutsideMax = si
    siFormWidthInsideMax = FormWidthOutsideMax - HeightDiffFormOutIn
    siAreasFrmWidthMax = siFormWidthInsideMax - FormMarginLeft - FormMarginRight
    siAreasFrmWidthMax = siFormWidthInsideMax - FormMarginLeft - FormMarginRight
    siMsectFrmWidthMax = siAreasFrmWidthMax - AreasFrmMarginLeft - AreasFrmMarginRight
    siMsectTextFrmWidthOnlyTextMax = siMsectFrmWidthMax - MsectFrmMarginLeft - MsectFrmMarginRight
    siMsectTextFrmWidthWithLposLblMax = siMsectFrmWidthMax - siMsectLabelWidthAll - MsectFrmMarginRight

End Property

Public Property Get FormWidthOutsideMin() As Single:                FormWidthOutsideMin = siFormWidthOutsideMin:                            End Property

Public Property Let FormWidthOutsideMin(ByVal si As Single)
' ------------------------------------------------------------------------------
' The minimum outside message form's width is assigned during initialization.
' This determines the minimum width of all subsequent frames and also their
' initial width. The width may be expanded by the setup of the message form's
' Title, its reply Buttons, and its monospaced sections. Once these setups
' had determined the widths it will be final for the setup of all prop-spaced
' sections.
' See function WidthExpanded for documentation of hte width subject.
' ------------------------------------------------------------------------------
    siFormWidthOutsideMin = si
    Me.Width = si
    '~~ Subsequent min width ajustments
    siFormWidthInsideMin = si - HeightDiffFormOutIn
    siAreasFrmWidthMin = siFormWidthInsideMin - FormMarginLeft - FormMarginRight
    siMsectFrmWidthMin = AreasFrmWidthMin - AreasFrmMarginLeft - AreasFrmMarginRight
    siMsectTextFrmWidthOnlyTextMin = siMsectFrmWidthMin - MsectFrmMarginLeft - MsectFrmMarginRight
    siMsectTextFrmWidthWithLposLblMin = siMsectFrmWidthMin - siMsectLabelWidthAll - MsectFrmMarginRight
    '~~ Make the (implicitely) specified min width the current effective widths
    siFormWidthInside = siFormWidthInsideMin
    siAreasFrmWidth = AreasFrmWidthMin
    siMsectFrmWidth = siMsectFrmWidthMin
    siMsectTextFrmWidthOnlyText = siMsectTextFrmWidthOnlyTextMin
    siMsectTextFrmWidthWithLposLbl = siMsectTextFrmWidthWithLposLblMin
    siBareaFrmWidth = AreasFrmWidthMin
    
End Property

Private Property Get HeightDiffFormOutIn() As Single:               HeightDiffFormOutIn = CInt(Me.Height - Me.InsideHeight):                End Property

Private Property Get LabelAllPos() As enLabelPos:                   LabelAllPos = lLabelAllPos:                                             End Property

Private Property Let LabelAllPos(ByVal en As enLabelPos):           lLabelAllPos = en:                                                      End Property

Public Property Let LabelAllSpec(ByVal l_spec As String)
    LabelAllPos = mMsg.LabelPos(l_spec)
    MsectLabelWidthAll = mMsg.LabelWidth(l_spec)
End Property

Private Property Get MaxRowsHeight() As Single:                     MaxRowsHeight = siMaxButtonHeight + (BttnsRowVerticalMargin * 2):       End Property

Public Property Let ModeLess(ByVal b As Boolean):                   bModeLess = b:                                                          End Property

Private Property Get MonitorHeightExSteps() As Single
    MonitorHeightExSteps = ContentHeight(frmSteps.Parent) - frmSteps.Height
End Property

Private Property Get MonitorHeightMaxSteps()
    MonitorHeightMaxSteps = Me.FormHeightOutsideMax - MonitorHeightExSteps
End Property

Public Property Get MonitorIsInitialized() As Boolean:              MonitorIsInitialized = Not cllSteps Is Nothing:                         End Property

Public Property Let MonitorProcess(ByVal s As String):              sMonitorProcess = s:                                                    End Property

Public Property Let MonitorStepsDisplayed(ByVal l As Long):         lMonitorStepsDisplayed = l:                                             End Property

Private Property Get MsectFrmMarginBottom() As Single:              MsectFrmMarginBottom = 2:                                               End Property

Private Property Get MsectFrmMarginLeft() As Single:                MsectFrmMarginLeft = 2:                                                 End Property

Private Property Get MsectFrmMarginRight() As Single:               MsectFrmMarginRight = 2:                                                End Property

Private Property Get MsectFrmMarginTop() As Single:                 MsectFrmMarginTop = 2:                                                  End Property

Private Property Get MsectLabelAbove(Optional ByVal m_sect As Long) As Boolean
    MsectLabelAbove = LabelAllPos = enLabelAboveSectionText And MsectLabelIsActive(m_sect)
End Property

Private Property Get MsectLabelTextBoxDiffHeight() As Single:       MsectLabelTextBoxDiffHeight = 6:                                        End Property

Private Property Get MsectLabelTextBoxDiffWidth() As Single:        MsectLabelTextBoxDiffWidth = 9:                                         End Property

Private Property Get MsectLabelTextMargin() As Single:              MsectLabelTextMargin = 4:                                               End Function

Private Property Get MsectLabelTop() As Single:                     MsectLabelTop = 4:                                                      End Property

Private Property Let MsectLabelWidthAll(ByVal si As Single):        siMsectLabelWidthAll = si:                                              End Property

Private Property Get MsectTextFrmMarginBottom() As Single:          MsectTextFrmMarginBottom = 2:                                           End Property

Private Property Get MsectTextFrmMarginLeft() As Single:            MsectTextFrmMarginLeft = 2:                                             End Property

Private Property Get MsectTextFrmMarginRight() As Single:           MsectTextFrmMarginRight = 2:                                            End Property

Private Property Get MsectTextFrmMarginTop() As Single:             MsectTextFrmMarginTop = 2:                                              End Property

Private Property Get MsectTextFrmWidth(Optional ByVal w_sect As Long) As Single
    
    If MsectTxtWithLftPosLbl(w_sect) _
    Then MsectTextFrmWidth = siMsectTextFrmWidthWithLposLbl _
    Else MsectTextFrmWidth = siMsectTextFrmWidthOnlyText

End Property

Private Property Let MsectTextFrmWidth(Optional ByVal w_sect As Long, _
                                               ByVal si As Single)
' ------------------------------------------------------------------------------
' The width of a text frame will only be increased when a vertical scroll-bar
' had been applied. However, this increase will cause correponding increases
' for all the other frames
' ------------------------------------------------------------------------------
    If si > siMsectTextFrmWidthWithLposLbl Then
        If MsectTxtWithLftPosLbl(w_sect) Then
            siMsectTextFrmWidthWithLposLbl = si
'            siMsectFrmWidth =
        Else
            siMsectTextFrmWidthOnlyText = si
        End If
    End If
    
End Property

Private Property Let MsectTextFrmWidthOnlyText(ByVal si As Single)
' ------------------------------------------------------------------------------
' A section's text frame width - in this case the width of the text with a left
' positioned label - may expand only with the setup of a monospaced section.
' In case it does, related widths are adjusted correspondingly.
' ------------------------------------------------------------------------------
    Dim siWidth As Single
    If WidthExpanded(si, siMsectTextFrmWidthOnlyTextMin, siMsectTextFrmWidthOnlyTextMax, siMsectTextFrmWidthOnlyText, siWidth) Then
        siMsectTextFrmWidthOnlyText = siWidth
        '~~ A new expanded width, which may only have been caused by a monospaced section setup -
        '~~ will expand all subsequent frames
        siMsectFrmWidth = siMsectTextFrmWidthOnlyText + MsectFrmMarginLeft + MsectFrmMarginRight
        siAreasFrmWidth = siMsectFrmWidth + AreasFrmMarginLeft + AreasFrmMarginRight
        FormWidthInside = siAreasFrmWidth + FormMarginLeft + FormMarginRight
    End If
    
End Property

Private Property Let MsectTextFrmWidthWithLposLbl(ByVal m_si As Single)
' ------------------------------------------------------------------------------
' A section's text frame width - in this case the width of the text with a left
' positioned label - may expand only with the setup of a monospaced section.
' In case it does, related widths are adjusted correspondingly.
' ------------------------------------------------------------------------------
    Dim si As Single
    If WidthExpanded(m_si, siMsectTextFrmWidthWithLposLblMin, siMsectTextFrmWidthWithLposLblMax, siMsectTextFrmWidthWithLposLbl, si) Then
        siMsectTextFrmWidthWithLposLbl = si
        siMsectFrmWidth = si + siMsectLabelWidthAll + MsectLabelTextMargin + MsectFrmMarginLeft + MsectFrmMarginRight
        siMsectTextFrmWidthOnlyText = siMsectFrmWidth - MsectFrmMarginLeft - MsectFrmMarginRight
        siAreasFrmWidth = siMsectFrmWidth + AreasFrmMarginLeft + AreasFrmMarginRight
        FormWidthInside = siAreasFrmWidth + FormMarginLeft + FormMarginRight
    End If
End Property

Private Property Get MsectTextMarginRight() As Single:              MsectTextMarginRight = 4:                                               End Property

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

Public Property Let MsgBttns(ByVal cll As Collection):              Set cllMsgBttns = cll:                                                  End Property

Private Property Get MsgButtonDefault() As Variant:                 MsgButtonDefault = vMsgButtonDefault:                                   End Property

Public Property Let MsgButtonDefault(ByVal v As Variant):           vMsgButtonDefault = v:                                                  End Property

Private Property Get MsgHasButtons():                               MsgHasButtons = cllMsgBttns.Count <> 0:                                 End Property

Private Property Get MsgHasText():                                  MsgHasText = MsectsActive.Count <> 0:                                   End Property

Private Sub MsgGet()
    Dim i As Long
    For i = 1 To NoOfMsgSects
        With Msg.Section(i)
            .Label = MsgLabel(i)
            .Text = MsgText(enSectText, i)
        End With
    Next i
End Sub

Public Property Get MsgLabel(Optional ByVal m_sect As Long = 1) As udtMsgLabel
' ------------------------------------------------------------------------------
' Returns the text for the Label of section (m_sectn).
' ------------------------------------------------------------------------------
    Dim vArry() As Variant
    
    If dctSectsLabel Is Nothing Then
        MsgLabel.Text = vbNullString
    ElseIf Not dctSectsLabel.Exists(m_sect) Then
        MsgLabel.Text = vbNullString
    Else
        vArry = dctSectsLabel(m_sect)
        MsgLabel.FontBold = vArry(0)
        MsgLabel.FontColor = vArry(1)
        MsgLabel.FontItalic = vArry(2)
        MsgLabel.FontName = vArry(3)
        MsgLabel.FontSize = vArry(4)
        MsgLabel.FontUnderline = vArry(5)
        MsgLabel.MonoSpaced = vArry(6)
        MsgLabel.Text = vArry(7)
        MsgLabel.OnClickAction = vArry(8)
    End If

End Property

Public Property Let MsgLabel(Optional ByVal m_sect As Long = 1, _
                                      ByRef m_udt As udtMsgLabel)
' ------------------------------------------------------------------------------
' Provide the text (m_udt) as section (m_sect) text, section Label,
' monitor header, footer, or step (lbl_kind_of_text).
' ------------------------------------------------------------------------------
    Dim vArry(0 To 8)   As Variant
    
    vArry(0) = m_udt.FontBold
    vArry(1) = m_udt.FontColor
    vArry(2) = m_udt.FontItalic
    vArry(3) = m_udt.FontName
    vArry(4) = m_udt.FontSize
    vArry(5) = m_udt.FontUnderline
    vArry(6) = m_udt.MonoSpaced
    vArry(7) = m_udt.Text
    vArry(8) = m_udt.OnClickAction
    If dctSectsLabel.Exists(m_sect) Then
        dctSectsLabel.Remove m_sect
    End If
    dctSectsLabel.Add m_sect, vArry

End Property

Private Property Get MsgText(Optional ByVal t_kind As KindOfText, _
                             Optional ByVal t_sect As Long = 1) As udtMsgText
' ------------------------------------------------------------------------------
' Returns the text (t_kind) as section-text or -Label, monitor-header,
' -footer, or -step.
' ------------------------------------------------------------------------------
    Dim vArry() As Variant
    
    Select Case t_kind
        Case enMonHeader:    MsgText = TextMonitorHeader
        Case enMonFooter:    MsgText = TextMonitorFooter
        Case enMonStep:      MsgText = TextMonitorStep
        Case enSectText
            With MsgText
                If dctSectsText Is Nothing Then
                    .Text = vbNullString
                ElseIf Not dctSectsText.Exists(t_sect) Then
                    .Text = vbNullString
                Else
                    vArry = dctSectsText(t_sect)
                    .FontBold = vArry(0)
                    .FontColor = vArry(1)
                    .FontItalic = vArry(2)
                    .FontName = vArry(3)
                    .FontSize = vArry(4)
                    .FontUnderline = vArry(5)
                    .MonoSpaced = vArry(6)
                    .Text = vArry(7)
                    .OnClickAction = vArry(8)
                End If
            End With
    End Select
    
End Property

Public Property Let MsgText(Optional ByVal t_kind As KindOfText, _
                            Optional ByVal t_sect As Long = 1, _
                                     ByRef t_udt As udtMsgText)
' ------------------------------------------------------------------------------
' Provide the text (t_udt) as section (txt_section) text, section Label,
' monitor header, footer, or step (txt_kind_of_text).
' ------------------------------------------------------------------------------
    Dim vArry(0 To 8)   As Variant
    
    With t_udt
        vArry(0) = .FontBold
        vArry(1) = .FontColor
        vArry(2) = .FontItalic
        vArry(3) = .FontName
        vArry(4) = .FontSize
        vArry(5) = .FontUnderline
        vArry(6) = .MonoSpaced
        vArry(7) = .Text
        vArry(8) = .OnClickAction
    End With
    Select Case t_kind
        Case enMonHeader:    TextMonitorHeader = t_udt
        Case enMonFooter:    TextMonitorFooter = t_udt
        Case enMonStep:      TextMonitorStep = t_udt
        Case enSectText
            If dctSectsText.Exists(t_sect) Then dctSectsText.Remove (t_sect)
            dctSectsText.Add t_sect, vArry
    End Select

    If t_sect = NoOfMsgSects Then
        MsgGet
    End If
    
End Property

Public Property Get MsgTitle() As String:                           MsgTitle = Me.Caption:                                                  End Property

Public Property Let MsgTitle(ByVal s As String):                    sMsgTitle = s:                                                          End Property

Private Property Let NewWidth(Optional ByRef n_frame As Object, _
                              Optional ByVal n_for_visible_only As Boolean = True, _
                              Optional ByVal n_width_threshold As Single = 5, _
                              Optional ByVal n_direct_child_only = True, _
                                       ByVal n_width As Single)
' ------------------------------------------------------------------------------
' Asigns a frame or form (n_frame) a new width (n_width) and a horizontal
' scroll-bar when the new width is less than the frame's content width by
' considering a threshold (n_width_threshold) avoiding a usesless scroll-bar for
' a redicolous width difference. In case the new width (n_width) is less the
' frame's content width, a horizontal scrollbar is removed.
' ------------------------------------------------------------------------------
    Const PROC = "NewWidth"
    
    On Error GoTo eh
    Dim siContentWidth  As Single
    
    siContentWidth = ContentWidth(n_frame, n_for_visible_only, n_direct_child_only)
    
    If n_frame Is Nothing Then Err.Raise AppErr(1), ErrSrc(PROC), "The required argument 'n_frame' is Nothing!"
    If Not IsFrameOrForm(n_frame) Then Err.Raise AppErr(2), ErrSrc(PROC), "The provided argument 'n_frame' is neither a Frame nor a Form!"
    
    n_frame.Width = n_width

    If siContentWidth - n_frame.Width > n_width_threshold Then
        ScrollHorApply s_frame:=n_frame, s_content_width:=siContentWidth
    ElseIf ScrollHorApplied(n_frame) Then
        ScrollHorRemove n_frame
    End If
    
xt: Exit Property
    
eh: If ErrMsg(ErrSrc(PROC)) = vbYes Then: Stop: Resume
End Property

Public Property Let ReplyWithIndex(ByVal b As Boolean):             bReplyWithIndex = b:                                                    End Property

Public Property Let SetupDone(ByVal b As Boolean):                  bSetUpDone = b:                                                         End Property

Private Property Get SysFrequency() As Currency
    If TimerSystemFrequency = 0 Then getFrequency TimerSystemFrequency
    SysFrequency = TimerSystemFrequency
End Property

Private Property Get TimerSecsElapsed() As Currency:                TimerSecsElapsed = TimerTicksElapsed / SysFrequency:                    End Property

Private Property Get TimerSysCurrentTicks() As Currency:            getTickCount TimerSysCurrentTicks:                                      End Property

Private Property Get TimerTicksElapsed() As Currency:               TimerTicksElapsed = cyTimerTicksEnd - cyTimerTicksBegin:                End Property

Public Property Get VisualizeForTest() As Boolean:                  VisualizeForTest = bVisualizeForTest:                                   End Property

Public Property Let VisualizeForTest(ByVal b As Boolean):           bVisualizeForTest = b:                                                  End Property

Private Function AddControl(ByVal ac_ctl As MSFormControls _
                 , Optional ByVal ac_in As MsForms.Frame = Nothing _
                 , Optional ByVal ac_name As String = vbNullString _
                 , Optional ByVal ac_visible As Boolean = False) As MsForms.Control
' ------------------------------------------------------------------------------
' Returns the type of control (ac_ctl) added to the to the userform or - when
' provided - to the frame (ac_in), optionally named (ac_name) and by default
' invisible (ac_visible).
' ------------------------------------------------------------------------------
    Const PROC = "AddControl"
    
    On Error GoTo eh
    Dim ctl As MsForms.Control
    Dim frm As MsForms.Frame
    
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

Private Sub AdjustedParentsWidthAndHeight(ByVal ctrl As MsForms.Control)
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
            If Not ScrollVerApplied(FrmParent) Then
                FrmParent.Width = ContentWidth(FrmParent) + 5
                FrmParent.Height = ctrl.Top + ContentHeight(FrmParent) + 30
            End If
        ElseIf IsFrameOrForm(FrmParent) Then
            If Not ScrollVerApplied(FrmParent) Then
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

Private Function AdjustToVgrid(ByVal atvg_si As Single, _
                      Optional ByVal atvg_threshold As Single = 1.5, _
                      Optional ByVal atvg_grid As Single = 4) As Single
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
    Dim i           As Long
    Dim j           As Long
    Dim sKey        As String
    Dim sButton     As String
    Dim Msg         As mMsg.udtMsg
    
    sButton = Replace(Replace(ar_button, vbCrLf, "|"), vbLf, "|")
    For i = 0 To dctApplicationRunArgs.Count - 1
        sKey = Replace(Replace(dctApplicationRunArgs.Keys()(i), vbCrLf, "|"), vbLf, "|")
        If sKey = sButton Then
            Set cll = dctApplicationRunArgs.Items()(i)
            sService = cll(1).Name & "!" & cll(2)
            
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
        
    With Msg
        j = j + 1
        With .Section(j).Text
            .Text = "Although the message is displayed modeless, this button has not been provided with Application.Run arguments *) which means that the button is useless (has no function)."
            .FontColor = rgbRed
            .FontBold = True
        End With
        j = j + 1
        With .Section(j).Label
            .Text = "*) In a modeless displayed form there should be no buttons other than those which had been provided " & _
                    "with 'Application.Run' arguments which specify which makro to execute when clicked (click this for help)."
            .OnClickAction = "https://github.com/warbe-maker/VBA-Message#the-buttonapprun-service"
            .FontColor = rgbBlue
        End With
    End With
    
    mMsg.Dsply dsply_title:="No 'Application.Run' information provided for this button!" _
             , dsply_msg:=Msg
    
xt: Exit Sub

eh: Select Case ErrMsg(ErrSrc(PROC))
        Case vbResume:  Stop: Resume
        Case Else:      GoTo xt
    End Select
End Sub

Private Function AppliedBttnRows() As Dictionary
' ------------------------------------------------------------------------------
' Returns a Dictionary of the visible button rows with the row
' frame as the key and the applied/visible buttons therein as item.
' ------------------------------------------------------------------------------
    Const PROC = "AppliedBttnRows"
    
    On Error GoTo eh
    Dim dct         As New Dictionary
    Dim lRow        As Long
    Dim frmRow      As MsForms.Frame
    Dim v           As Variant
    Dim lButtons    As Long
    Dim cmb         As MsForms.CommandButton
    
    For lRow = 1 To dctBttnsRowFrm.Count
        Set frmRow = dctBttnsRowFrm(lRow)
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

Private Function AppliedButtonRowHeight() As Single:         AppliedButtonRowHeight = siMaxButtonHeight + 2:    End Function

Private Sub AutoSizeApplyLabelProperties(ByVal a_sect As Long, _
                                         ByRef a_lbl As MsForms.Label, _
                                         ByRef a_text As String)
' ------------------------------------------------------------------------------
' Returns the section's (a_sect) Label control (a_lbl) and a temporary TextBox
' control (a_tbx) with the sections text properties for the Label, and thereby
' the to be autosized text (a_text).
' ------------------------------------------------------------------------------
    Const PROC = "AutoSizeApplyLabelProperties"
    
    On Error GoTo eh
    Dim frmSect As MsForms.Frame:   Set frmSect = MsectFrm(a_sect)
    
    MsgSectLbl = MsgLabel(a_sect)
    a_text = MsgSectLbl.Text
    TempTbx a_lbl, t_remove:=True    ' remove any if existing
    TempTbx a_lbl                    ' provide tbxTemp

    With frmSect
        With .Font
            .Name = MsectLabelFontName(a_sect)
            .Size = MsectLabelFontSize(a_sect)
            If .Bold <> MsgSectLbl.FontBold Then .Bold = MsgSectLbl.FontBold
            If .Italic <> MsgSectLbl.FontItalic Then .Italic = MsgSectLbl.FontItalic
            If .Underline <> MsgSectLbl.FontUnderline Then .Underline = MsgSectLbl.FontUnderline
        End With
    End With
    
    With a_lbl
        .Top = 0
        With .Font
            .Name = MsectLabelFontName(a_sect)
            .Size = MsectLabelFontSize(a_sect)
            If .Bold <> MsgSectLbl.FontBold Then .Bold = MsgSectLbl.FontBold
            If .Italic <> MsgSectLbl.FontItalic Then .Italic = MsgSectLbl.FontItalic
            If .Underline <> MsgSectLbl.FontUnderline Then .Underline = MsgSectLbl.FontUnderline
        End With
        If .ForeColor <> MsgSectLbl.FontColor And MsgSectLbl.FontColor <> 0 Then .ForeColor = MsgSectLbl.FontColor
        .Left = MsectFrmMarginLeft
    End With
    
    '~~ Both controls are assigned the same text properties in order to get an appropriate
    '~~ autosize result
    With tbxTemp
        .Top = a_lbl.Top
        .Left = MsectFrmMarginLeft
        With .Font
            .Name = a_lbl.Font.Name
            .Size = a_lbl.Font.Size
            If a_lbl.Font.Bold Then .Bold = a_lbl.Font.Bold
            If a_lbl.Font.Italic Then .Italic = a_lbl.Font.Italic
            If a_lbl.Font.Underline Then .Underline = a_lbl.Font.Underline
        End With
        .ForeColor = a_lbl.ForeColor
    End With
                                            
xt: Exit Sub
    
eh: If ErrMsg(ErrSrc(PROC)) = vbYes Then: Stop: Resume
End Sub

Private Sub AutoSizeApplyTextProperties(ByVal a_sect As Long, _
                                        ByRef a_lbl As MsForms.Label, _
                                        ByRef a_text As String, _
                                        ByVal a_monospaced As Boolean)
' ------------------------------------------------------------------------------
'
' ------------------------------------------------------------------------------
    Const PROC = "AutoSizeApplyTextProperties"
    
    On Error GoTo eh
    Dim frmText As MsForms.Frame: Set frmText = MsectTextFrm(a_sect)
    
    a_text = Msg.Section(a_sect).Text.Text
    TempTbx a_lbl, t_remove:=True    ' remove any if existing
    TempTbx a_lbl                    ' provide tbxTemp
    
    With frmText
        With .Font
            .Name = MsectTextFontName(a_sect)
            .Size = MsectTextFontSize(a_sect)
            DoEvents
            If .Bold <> Msg.Section(a_sect).Text.FontBold Then .Bold = Msg.Section(a_sect).Text.FontBold
            If .Italic <> Msg.Section(a_sect).Text.FontItalic Then .Italic = Msg.Section(a_sect).Text.FontItalic
            If .Underline <> Msg.Section(a_sect).Text.FontUnderline Then .Underline = Msg.Section(a_sect).Text.FontUnderline
        End With
        DoEvents
    End With
    DoEvents
    
    With a_lbl
        With .Font
            .Name = frmText.Font.Name
            .Size = frmText.Font.Size
            DoEvents
            If .Bold <> Msg.Section(a_sect).Text.FontBold Then .Bold = Msg.Section(a_sect).Text.FontBold
            If .Italic <> Msg.Section(a_sect).Text.FontItalic Then .Italic = Msg.Section(a_sect).Text.FontItalic
            If .Underline <> Msg.Section(a_sect).Text.FontUnderline Then .Underline = Msg.Section(a_sect).Text.FontUnderline
        End With
        DoEvents
        If .ForeColor <> Msg.Section(a_sect).Text.FontColor And MsgSectTxt.FontColor <> 0 Then .ForeColor = Msg.Section(a_sect).Text.FontColor
    End With
    DoEvents
    
    With tbxTemp
        .Top = a_lbl.Top
        .Left = a_lbl.Left
        With .Font
            .Name = frmText.Font.Name
            .Size = frmText.Font.Size
            DoEvents
            If .Bold <> Msg.Section(a_sect).Text.FontBold Then .Bold = Msg.Section(a_sect).Text.FontBold
            If .Italic <> Msg.Section(a_sect).Text.FontItalic Then .Italic = Msg.Section(a_sect).Text.FontItalic
            If .Underline <> Msg.Section(a_sect).Text.FontUnderline Then .Underline = Msg.Section(a_sect).Text.FontUnderline
        End With
        DoEvents
        .ForeColor = a_lbl.ForeColor
    End With
    DoEvents
    
xt: Exit Sub
    
eh: If ErrMsg(ErrSrc(PROC)) = vbYes Then: Stop: Resume
End Sub

Private Sub AutoSizeLabelViaTextBox(ByVal a_sect As Long, _
                                    ByRef a_lbl As MsForms.Label, _
                                    ByVal a_text As String, _
                                    ByVal a_width_limit As Single, _
                           Optional ByVal a_align As Long = fmTextAlignLeft)
' ------------------------------------------------------------------------------
' Background:
' The mMsg/fMsg component uses Label controls for the text which provides the
' advantage of a click-event. However, because a Label control does not consider
' line breaks the AutoSize is performed via/through a temporary TextBox control
' (tbxTemp). Thereby it must be considered that for an identical content a
' TextBox control's width differs from a Label control's width.
' ------------------------------------------------------------------------------
    Const PROC = "AutoSizeLabelViaTextBox"
    
    On Error GoTo eh
    
    With a_lbl
        .Visible = False
        tbxTemp.Visible = True
        If a_width_limit = 0 Then
            AutoSizeTextBox tbxTemp, a_text, a_width_limit
        Else
            AutoSizeTextBox tbxTemp, a_text, a_width_limit + MsectLabelTextBoxDiffWidth
        End If
        .WordWrap = True
        .AutoSize = False
        .Caption = a_text
        DoEvents
        .Height = tbxTemp.Height - MsectLabelTextBoxDiffHeight
        .Width = tbxTemp.Width - MsectLabelTextBoxDiffWidth
        .TextAlign = a_align ' !must be done before width adjustment!
        tbxTemp.Visible = False
        .Visible = True
    End With
    
xt: Exit Sub
    
eh: If ErrMsg(ErrSrc(PROC)) = vbYes Then: Stop: Resume
End Sub

Private Sub AutoSizeTextBox(ByRef a_tbx As MsForms.TextBox, _
                            ByVal a_text As String, _
                            ByVal a_width_limit As Single, _
                   Optional ByVal a_append As Boolean = False, _
                   Optional ByVal a_append_margin As String = vbNullString)
' ------------------------------------------------------------------------------
' Common AutoSize service for an MsForms.TextBox providing a width and height
' for the TextBox (a_tbx) by considering:
' - When a width limit is provided (a_width_limit > 0) the width is regarded a
'   fixed maximum and thus the height is auto-sized by means of WordWrap=True.
' - When no width limit is provided (the default) WordWrap=False and thus the
'   width of the TextBox is determined by the longest line.
' - When a maximum width is provided (a_width_max > 0) and the parent of the
'   TextBox is a frame a horizontal scrollbar is applied for the parent frame.
' - When a maximum height is provided (a_heightmax > 0) and the parent of the
'   TextBox is a frame a vertical scrollbar is applied for the parent frame.
' - When a minimum width (a_width_min > 0) or a minimum height (a_height_min
'   > 0) is provided the size of the textbox is set correspondingly. This
'   option is specifically usefull when text is appended to avoid much flicker.
'
' Uses: AdjustToVgrid
'
' W. Rauschenberger Berlin April 2022
' ------------------------------------------------------------------------------
    
    With a_tbx
        .Visible = True
        .MultiLine = True
        If a_width_limit > 0 Then
            '~~ AutoSize the height of the TextBox considering the limited width
            '~~ (applied for proportially spaced text where the width determines the height)
            .WordWrap = True
            .AutoSize = False
            .Width = a_width_limit - MsectTextMarginRight  ' plus a readability space
            
            If Not a_append Then
                .Value = a_text
            Else
                If .Value = vbNullString Then
                    .Value = a_text
                Else
                    .Value = .Value & a_append_margin & vbLf & a_text
                End If
            End If
            .AutoSize = True
        Else
            '~~ AutoSize the height  a n d  width of the TextBox
            '~~ (applied for mono-spaced text where the longest line defines the width)
            .MultiLine = True
            .WordWrap = False ' the means to limit the width
            .AutoSize = True
            If Not a_append Then
                .Value = a_text
            Else
                If .Value = vbNullString Then
                    .Value = a_text
                Else
                    .Value = .Value & vbLf & a_text
                End If
            End If
        End If
        .Height = AdjustToVgrid(.Height, 0)
    End With
        
xt: Exit Sub

End Sub

Private Function BareaFrm(Optional b_properties As Boolean = True) As MsForms.Frame
' ------------------------------------------------------------------------------
' Returns the Buttons area Frame, created if yet not existing.
' ------------------------------------------------------------------------------
    Const PROC = "BareaFrm"
    
    On Error GoTo eh
    
    If Not BareaFrmExists(BareaFrm) _
    Then Err.Raise AppErr(1), ErrSrc(PROC), "Userform design does not conform with expectations!"
    
    If cllMsgBttns.Count = 0 _
    Then Err.Raise AppErr(2), ErrSrc(PROC), "The Buttons-Area-Frame has been called/requested althoug ther are no Message-Buttons specified!"
        
    With BareaFrm
        If b_properties Then
            .Visible = True
            .Top = BareaFrmTop
            .Left = FormMarginLeft
            Select Case .ScrollBars
                Case fmScrollBarsNone
                    .Width = siAreasFrmWidth + AreasFrmMarginLeft + AreasFrmMarginRight
                    BttnsFrm
                    .Height = Max(50, ContentHeight(BareaFrm) + AreasFrmMarginTop + AreasFrmMarginBottom + ScrollHorHeight(frmBarea))
                Case fmScrollBarsHorizontal
                Case fmScrollBarsVertical
                    .ScrollHeight = ContentHeight(BareaFrm) + AreasFrmMarginTop + AreasFrmMarginBottom
                    .Width = siAreasFrmWidth + AreasFrmMarginRight + SCROLL_V_WIDTH
                    BttnsFrm
                    ' height is considered final once a vertical scroll-bar is applied
                Case fmScrollBarsBoth
                    ' width and height are considered final
                    BttnsFrm
                    .ScrollHeight = ContentHeight(BareaFrm) + AreasFrmMarginTop + AreasFrmMarginBottom
                    .ScrollWidth = ContentWidth(BareaFrm) + AreasFrmMarginLeft + AreasFrmMarginRight
            End Select
            VisualizationsForTestOnly frmBarea, VISUALIZE_CLR_AREA
        End If
    End With
    DoEvents

xt: Exit Function
    
eh: If ErrMsg(ErrSrc(PROC)) = vbYes Then: Stop: Resume
End Function

Private Sub BareaFrmAdjust()
    BttnsFrmAdjust
    BareaFrm.Width = ContentWidth(BareaFrm)
    BareaFrm.Height = ContentHeight(BareaFrm)
End Sub

Private Sub BareaFrmCenter()
    Dim si As Single
    
    If MsgHasButtons And bButtonsSetup Then
        With BareaFrm(False)
            si = (Me.InsideWidth - .Width) / 2
            If .Left <> si Then
                .Left = si
            End If
        End With
    End If

End Sub

Private Function BareaFrmExists(ByRef b_frm As MsForms.Frame) As Boolean
    If Not frmBarea Is Nothing Then
        Set b_frm = frmBarea
        BareaFrmExists = True
    End If
End Function

Private Function BareaFrmTop() As Single
    
    BareaFrmTop = FormMarginTop
    If MsgHasText Then
        With MareaFrm
            BareaFrmTop = AdjustToVgrid(.Top + .Height + AreasMarginVertical)
        End With
    End If

End Function

Private Function BttnsFrm() As MsForms.Frame
' ------------------------------------------------------------------------------
' Returns the Frame of the message buttons, created in the BareaFrm if yet
' not existing.
' ------------------------------------------------------------------------------
    Const PROC = "BttnsFrm"
    
    On Error GoTo eh
    Dim frmBttns    As MsForms.Frame
    Dim si          As Single
    
    If Not BttnsFrmExists(frmBttns) _
    Then Err.Raise AppErr(1), ErrSrc(PROC), "Userform design does not conform with expectations!"
    
    With frmBttns
        .Visible = True
        .Height = Max(20, ContentHeight(frmBttns) + BttnsFrmMarginTop + BttnsFrmMarginBottom)
        .Top = AreasFrmMarginTop
        .Width = Max(20, ContentWidth(frmBttns) + AreasFrmMarginLeft + AreasFrmMarginRight)
        si = (.Parent.Width - .Width) / 2 ' centered
        If si >= AreasFrmMarginLeft _
        Then .Left = si _
        Else .Left = AreasFrmMarginLeft
    End With
    VisualizationsForTestOnly frmBttns, rgbLightGreen
    Set BttnsFrm = frmBttns
    
xt: Exit Function
    
eh: If ErrMsg(ErrSrc(PROC)) = vbYes Then: Stop: Resume
End Function

Private Sub BttnsFrmAdjust()
    BttnsFrm.Width = ContentWidth(BttnsFrm)
    BttnsFrm.Height = ContentHeight(BttnsFrm)
End Sub

Private Function BttnsFrmExists(ByRef b_frm As MsForms.Frame) As Boolean
    If Not frmBttnsFrm Is Nothing Then
        Set b_frm = frmBttnsFrm
        BttnsFrmExists = True
    End If
End Function

Private Function BttnsRowFrm(ByVal b_row As Long) As MsForms.Frame
' ------------------------------------------------------------------------------
' Returns the Frame of the buttons row (b_row), created in the BttnsFrm if yet
' not existing.
' ------------------------------------------------------------------------------
    
    If Not BttnsRowFrmExists(b_row, BttnsRowFrm) Then
        Set BttnsRowFrm = AddControl(ac_ctl:=Frame, ac_in:=BttnsFrm, ac_name:="frBttnsRow" & b_row)
        dctBttnsRowFrm.Add b_row, BttnsRowFrm
    End If
    VisualizationsForTestOnly BttnsRowFrm, VISUALIZE_CLR_BTTNS_ROW_FRM
    
End Function

Private Function BttnsRowFrmExists(ByVal b_row As Long, _
                                   ByRef b_row_frm As MsForms.Frame) As Boolean
    If dctBttnsRowFrm.Exists(b_row) Then
        Set b_row_frm = dctBttnsRowFrm(b_row)
        BttnsRowFrmExists = True
    End If
End Function

Private Function BttnsRowVerticalMargin() As Single:              BttnsRowVerticalMargin = 0:                                               End Function

Private Sub ButtonClicked(ByVal cmb As MsForms.CommandButton)
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

Private Function ClickedButtonIndex(ByVal cmb As MsForms.CommandButton) As Long
    
    Dim i   As Long
    Dim v   As Variant
    
    For Each v In AppliedBttnsRetVal
        i = i + 1
        If v Is cmb Then
            ClickedButtonIndex = i
            Exit For
        End If
    Next v

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
           Optional ByVal col_set_visible As Boolean = False)
' ------------------------------------------------------------------------------
' Setup of a Collection (col_into) with all type (col_cntrl_type) controls
' with a parent (col_with_parent) as Collection (col_into) by assigning the
' an initial height (col_set_height) and width (col_set_width).
' ------------------------------------------------------------------------------
    Const PROC = "Collect"
    
    On Error GoTo eh
    Dim ctl As MsForms.Control
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
                    End With
                    Select Case TypeName(col_into)
                        Case "Dictionary"
                            i = col_into.Count + 1
                            col_into.Add i, ctl
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
    Dim frm         As MsForms.Frame
    Dim lRow        As Long
    Dim lBttn       As Long
    Dim cmb         As MsForms.CommandButton
    Dim sKey        As String
    
    Collect col_into:=NewDict(dctAreas) _
          , col_cntrl_type:="Frame" _
          , col_with_parent:=Me
    Set frmMarea = dctAreas(1)
    Set frmBarea = dctAreas(2)
    
    Collect col_into:=NewDict(dctMsectFrm) _
          , col_cntrl_type:="Frame" _
          , col_with_parent:=frmMarea
    
    Collect col_into:=NewDict(dctMsectLabel) _
          , col_cntrl_type:="Label" _
          , col_with_parent:=dctMsectFrm
    
    Collect col_into:=NewDict(dctMsectTextFrm) _
          , col_cntrl_type:="Frame" _
          , col_with_parent:=dctMsectFrm
    
    Collect col_into:=NewDict(dctMsectTextTbx) _
          , col_cntrl_type:="TextBox" _
          , col_with_parent:=dctMsectTextFrm
        
    Collect col_into:=NewDict(dctMsectTextLbl) _
          , col_cntrl_type:="Label" _
          , col_with_parent:=dctMsectTextFrm
        
    Collect col_into:=frmBttnsFrm _
          , col_cntrl_type:="Frame" _
          , col_with_parent:=frmBarea _
          , col_set_visible:=True ' minimum is one button
    
    Collect col_into:=NewDict(dctBttnsRowFrm) _
          , col_cntrl_type:="Frame" _
          , col_with_parent:=frmBttnsFrm _
          , col_set_visible:=False ' minimum is one button
        
    NewDict dctBttns
    For lRow = 1 To dctBttnsRowFrm.Count
        Set frm = dctBttnsRowFrm(lRow)
        For lBttn = 0 To frm.Controls.Count - 1
            Set cmb = frm.Controls(lBttn)
            sKey = lRow & "-" & lBttn + 1
            dctBttns.Add sKey, cmb
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
    Dim ctl As MsForms.Control
    Dim i   As Long
    Dim siMinTop    As Single
    Dim siMaxBottom As Single
    
    If Not IsFrameOrForm(ch_frame_form) _
    Then Err.Raise AppErr(1), ErrSrc(PROC), "The provided argument is neither a Frame nor a Form - and thus has no controls!"
    
    If ch_frame_form Is Me _
    Then siMinTop = Me.InsideWidth _
    Else siMinTop = ch_frame_form.Height
    
    For Each ctl In ch_frame_form.Controls
        With ctl
            If .Parent Is ch_frame_form Then
                If ch_visible_only Then
                    If ctl.Visible Then
                        siMinTop = Min(siMinTop, .Top)
                        siMaxBottom = Max(siMaxBottom, .Top + .Height)
                        i = i + 1
                    End If
                Else
                    siMinTop = Min(siMinTop, .Top)
                    siMaxBottom = Max(siMaxBottom, .Top + .Height)
                    i = i + 1
                End If
            End If
        End With
    Next ctl
    ContentHeight = siMaxBottom - siMinTop
    
xt: Exit Function
    
eh: If ErrMsg(ErrSrc(PROC)) = vbYes Then: Stop: Resume
End Function

Public Function ContentWidth(Optional ByVal c_frame_form As Variant = Nothing, _
                             Optional ByVal c_visible_only As Boolean = True, _
                             Optional ByVal c_direct_child_only As Boolean = True) As Single
' ------------------------------------------------------------------------------
' Returns the width of the largest control in a Frame or Form (c_frame_form)
' which is maximum right pos minus the minimum left positions.
' Note: Any margins of the frame are not included! Thus the surrounding frame
' width is always the content width plus a left and plus a right margin.
' ------------------------------------------------------------------------------
    Const PROC = "ContentWidth"
    
    On Error GoTo eh
    Dim ctl As MsForms.Control
    Dim i   As Long
    Dim siMinLeft   As Single
    Dim siMaxRight  As Single
    
    If c_frame_form Is Nothing Then Set c_frame_form = Me
    If Not IsFrameOrForm(c_frame_form) _
    Then Err.Raise AppErr(1), ErrSrc(PROC), "The provided argument is neither a Frame nor a Form - and thus has no controls!"

    If c_frame_form Is Me _
    Then siMinLeft = Me.InsideWidth _
    Else siMinLeft = c_frame_form.Width
    
    For Each ctl In c_frame_form.Controls
        With ctl
            If (c_direct_child_only And .Parent Is c_frame_form) _
            Or c_direct_child_only = False Then
                If c_visible_only Then
                    If ctl.Visible Then
                        siMinLeft = Min(siMinLeft, .Left)
                        siMaxRight = Max(siMaxRight, (.Left + .Width))
                        i = i + 1
                    End If
                Else
                    siMinLeft = Min(siMinLeft, .Left)
                    siMaxRight = Max(siMaxRight, (.Left + .Width))
                    i = i + 1
                End If
            End If
        End With
    Next ctl
    ContentWidth = siMaxRight - siMinLeft
    
xt: Exit Function
    
eh: If ErrMsg(ErrSrc(PROC)) = vbYes Then: Stop: Resume
End Function

Private Sub ConvertPixelsToPoints(ByVal x_dpi As Single, _
                                  ByVal y_dpi As Single, _
                                  ByRef x_pts As Single, _
                                  ByRef y_pts As Single)
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
    Dim ctl As MsForms.Control
    For Each ctl In Me.Controls
        If ctl.Name = ce_name Then
            CtlExists = True
            Exit For
        End If
    Next ctl
End Function

Private Sub CursorForLabelWithOnClickAction(ByVal hc_section As Long)
    If Msg.Section(hc_section).Label.OnClickAction <> vbNullString Then
        AddCursor IDC_HAND
    End If
End Sub

Private Sub CursorForTextWithOnClickAction(ByVal hc_section As Long)
    If Msg.Section(hc_section).Text.OnClickAction <> vbNullString Then
        AddCursor IDC_HAND
    End If
End Sub

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

Private Sub FinalizeHeights(ByRef f_finalized As Boolean)
' ------------------------------------------------------------------------------
' 1. Finalize heights and top positions.
' 2. When the resulting height exceeds the default or specified height,
'    either reduce a dominating frame's height or the message and the buttons
'    area height propotionally and apply a vertical scroll-bar.
' When neither is done or had already been done TRUE is returned (f_finalized).
' ------------------------------------------------------------------------------
    Const PROC = "FinalizeHeights"
    
    On Error GoTo eh
    Dim v   As Variant
    Dim si  As Single
    
    FinalizeHeightsAndTopPositions f_finalized
    If Not f_finalized Then
        '~~ When the message form's height exceeds the specified maximum height
        With Me
            If .InsideHeight > FormHeightInsideMax - SCROLL_VER_THRESHOLD Then
                '~~ Provide a vertical scroll-bar for the dominating frame (message section text frame,
                '~~ message area, buttons area or message area and buttons area proportionally
                ScrollVerForHeightExceedingFrames exceeding_height:=.InsideHeight - FormHeightInsideMax _
                                                , s_finalized:=f_finalized
                If Not f_finalized Then
'                    FinalizeHeightsAndTopPositions f_finalized
                End If
            End If
        End With ' height exceeds specified maximum
    End If
    FinalizeHeightsAndTopPositions f_finalized
    FormHeightAdjust
    
xt: Exit Sub
    
eh: If ErrMsg(ErrSrc(PROC)) = vbYes Then: Stop: Resume
End Sub

Private Sub FinalizeHeightsAndTopPositions(ByRef f_finalized As Boolean)
' ------------------------------------------------------------------------------
' Reduce the height of the message area - or a message section if one is
' dominating - and or the height of the buttons area in order to have the
' message form not exceeds the specified maximum height.
' The area which occupies 65% or more of the overall height is the one being
' reduced. Otherwise both are reduced proportionally.
' When one of the message sections within the to be reduced message area
' occupies 80% or more of the overall message area height only this section
' is reduced with an applied vertical scrollbar.
' Returns done (f_finalized) TRUE when no adjustments had been done.
' ------------------------------------------------------------------------------
    Const PROC = "FinalizeHeightsAndTopPositions"
    
    On Error GoTo eh
    Dim v           As Variant
    Dim frmSect     As MsForms.Frame
    Dim frmText     As MsForms.Frame
    Dim frmMarea    As MsForms.Frame
    Dim frmBarea    As MsForms.Frame
    Dim lblText     As MsForms.Label
    Dim lblLabel    As MsForms.Label
    Dim si          As Single
    Dim siTopNext   As Single:          siTopNext = FormMarginTop
    
    '~~ Adjustment of the final height and any reaulting Top position is
    '~~ done just by calling the frame
    If MsgHasText Then
        Set frmMarea = MareaFrm
        For Each v In MsectsActive ' Loop through all active/displayed message sections
            Select Case True
                Case MsectHasOnlyLabel(v):  MsectLabel v
                Case MsectHasOnlyText(v):   MsectTextFrm v
                Case Else:                  MsectLabel v
                                            MsectTextFrm v
            End Select
            MsectFrm v
        Next v
    End If
           
    If MsgHasButtons And bButtonsSetup Then
        BareaFrm
    End If
    
    si = ContentHeight(Me) + FormMarginTop + FormMarginBottom + HeightDiffFormOutIn
    If Me.Height <> si Then
        If Not ScrollVerApplied(Me) Then
            Me.Height = si
        End If
    End If

xt: Exit Sub
    
eh: If ErrMsg(ErrSrc(PROC)) = vbYes Then: Stop: Resume
End Sub

Private Sub FinalizeWidthFrame(ByVal f_frm As MsForms.Frame, _
                               ByVal f_margin_left As Single, _
                               ByVal f_margin_right As Single, _
                               ByRef f_finalized As Boolean, _
                               ByVal f_max_width As Single, _
                      Optional ByVal f_x_action As fmScrollAction = fmScrollActionNoChange)
' ------------------------------------------------------------------------------
' Provides a horizontal scroll-bar in case:
' - the content of the frame (f_frm) exceeds its maximum width (f_max_width),
' - a scroll-bar is yet not provided.
' When a horizontal scroll-bar is already applied for the frame (f_frm) only the scroll-width is adjusted.
' The width of the frame (f_frm) remains un-changed, i.e. is regarded final.
' ------------------------------------------------------------------------------
    Const PROC = "FinalizeWidthFrame"
    
    On Error GoTo eh
    Dim si                  As Single
    Dim siWidthContent      As Single ' the frame's content width not considering any left/right margin
    Dim siWidthContMax      As Single ' the maximum content which fits into the frame without a horizontal scroll-bar
    Dim bAppliedOrChanged   As Boolean
    
    siWidthContMax = f_max_width - f_margin_left - f_margin_right
    siWidthContent = ContentWidth(f_frm)
    With f_frm
        Debug.Print .Name
        If siWidthContent > siWidthContMax + SCROLL_HOR_THRESHOLD Then
            '~~ In cas a horizontal scroll-bar is already applied
            '~~ only the scroll-width is adjusted to the content width
            '~~ which has to consider a left margin
            ScrollHorApply s_frame:=f_frm _
                         , s_content_width:=siWidthContent + f_margin_left + f_margin_right _
                         , s_x_action:=f_x_action _
                         , s_applied_or_changed:=bAppliedOrChanged
            If bAppliedOrChanged Then
                f_finalized = False
            End If
        End If
    End With

xt: Exit Sub
    
eh: If ErrMsg(ErrSrc(PROC)) = vbYes Then: Stop: Resume
End Sub

Private Sub FinalizeWidthLabel(ByVal f_lbl As MsForms.Label, _
                               ByVal f_width As Single, _
                               ByVal f_margin_left As Single, _
                               ByVal f_margin_right As Single, _
                               ByVal f_finalized As Boolean)
    With f_lbl
        .Left = f_margin_left
        If .Width < f_width - 5 Then
            .TextAlign = fmTextAlignLeft
            .Font.Size = .Font.Size
            DoEvents
        End If
    End With
    DoEvents

End Sub

Private Sub FinalizeWidths(ByRef f_finalized As Boolean)
' ------------------------------------------------------------------------------
' Adjusts all visible frame control's width considering their content width and
' providing a horizontal scroll-bar for those which exceed the maximum width.
' ------------------------------------------------------------------------------
    Const PROC = "FinalizeWidths"
    
    On Error GoTo eh
    Dim frmBarea    As MsForms.Frame
    Dim frmMarea    As MsForms.Frame
    Dim frmSect     As MsForms.Frame
    Dim frmText     As MsForms.Frame
    Dim lblText     As MsForms.Label
    Dim si          As Single
    Dim v           As Variant
    
    If MsgHasText Then
        Set frmMarea = MareaFrm
        For Each v In MsectsActive ' Loop through all active/displayed message sections
            
            If MsectHasLabelAndText(v) _
            Or MsectHasOnlyText(v) Then
                Set frmText = MsectTextFrm(v)
                FinalizeWidthFrame f_frm:=frmText _
                                 , f_margin_left:=MsectTextFrmMarginLeft _
                                 , f_margin_right:=MsectTextFrmMarginRight _
                                 , f_finalized:=f_finalized _
                                 , f_max_width:=MsectTextFrmWidthMax(v)
            End If
        Next v
    End If
        
    If bButtonsSetup Then
        FinalizeWidthFrame f_frm:=BareaFrm _
                         , f_margin_left:=AreasFrmMarginLeft _
                         , f_margin_right:=AreasFrmMarginRight _
                         , f_finalized:=f_finalized _
                         , f_x_action:=fmScrollActionBegin _
                         , f_max_width:=siAreasFrmWidthMax
        
    End If
    si = ContentWidth() + FormMarginRight + ScrollVerWidth(Me) + FormWidthDiffOutIn
    If Me.Width <> si Then
        Me.Width = si
    End If
    
    If bButtonsSetup Then
        BareaFrmCenter
    End If

xt: Exit Sub
    
eh: If ErrMsg(ErrSrc(PROC)) = vbYes Then: Stop: Resume
End Sub

Private Function FormBottomSpace() As Single:           FormBottomSpace = 32:                                           End Function

Private Sub FormHeightAdjust()
    Me.Height = FormMarginTop + ContentHeight(Me) + FormMarginBottom + (Me.Height - Me.InsideHeight)
End Sub

Private Function FormWidthFinal() As Single
    FormWidthFinal = ContentWidth(Me) + FormWidthDiffOutIn + FormMarginLeft + FormMarginRight
End Function

Private Function FormWidthInsideMonospaced(ByVal m_sect As Long, _
                                           ByVal m_si As Single) As Single
' ------------------------------------------------------------------------------
'
' ------------------------------------------------------------------------------
    
    FormWidthInsideMonospaced = FormMarginLeft _
                              + MsectFrmMarginLeft _
                              + MsectTextFrmLeft(m_sect) _
                              + MsectTextFrmMarginLeft _
                              + MsectTextFrmMarginRight _
                              + m_si _
                              + MsectFrmMarginRight _
                              + FormMarginRight
      
End Function

Private Sub FrameCenterHorizontal(ByVal f_center As MsForms.Frame, _
                         Optional ByVal f_within As MsForms.Frame = Nothing, _
                         Optional ByVal f_margin_left As Single = 0)
' ------------------------------------------------------------------------------
' Center the frame (f_center) horizontally within the frame (f_within)
' - which defaults to the UserForm when not provided.
' ------------------------------------------------------------------------------
    
    If f_within Is Nothing _
    Then f_center.Left = (Me.InsideWidth - f_center.Width) / 2 _
    Else f_center.Left = (f_within.Width - f_center.Width) / 2
    If f_center.Left = 0 Then f_center.Left = f_margin_left

End Sub

Private Function FrameHeightExeedsThreshold(ByVal m_percentage_threshold As Single, _
                                            ByRef m_frm As MsForms.Frame) As Boolean
' ------------------------------------------------------------------------------
' Returns TRUE, the exceeding height (m_exceeding) and the section's text frame
' (m_frm) in the message area which exceeds the threshold percentage height
' (m_percentage_threshold) based on the message areas height. When none does,
' but the message area as a whole does - based on the message form's inside
' height does - the message area frame is returned analogous.
' ------------------------------------------------------------------------------
    Const PROC = "FrameHeightExeedsThreshold"
    
    On Error GoTo eh
    Dim i                   As Long
    Dim siThresholdHeight   As Single
    Dim frm                 As MsForms.Frame
    Dim frmMarea            As MsForms.Frame
    Dim frmBarea            As MsForms.Frame
    
    siThresholdHeight = HeightThreshold(m_percentage_threshold)
    
    If MsgHasText Then
        Set frmMarea = MareaFrm
        '~~ At first try to identify a message section's text frame which
        '~~ is the one with the dominating height
        For i = 1 To lMaxNoOfMsgSects
            If MsectFrmIsActive(i, frm) Then
                If MsectTextFrmIsActive(i, , frm) Then
                    If frm.Height >= siThresholdHeight Then
                        If frm.Height - siThresholdHeight > 100 Then
                            '~~ Only when the height can be reduced with a positive heoght result
                            Set m_frm = frm
                            FrameHeightExeedsThreshold = True
                            Exit Function
                        End If
                    End If
                End If
            End If ' Section is active
        Next i
        
        If frmMarea.Height >= siThresholdHeight Then
            '~~ When none of the active message section's text frame exceeded the
            '~~ threshold height but the message area as a whole does:
            Set m_frm = frmMarea
            FrameHeightExeedsThreshold = True
            Exit Function
        End If
    End If
        
    If MsgHasButtons Then
        Set frmBarea = BareaFrm
        If frmBarea.Height >= siThresholdHeight Then
            '~~ When neither a message section's text frame height is dominating (>= n %)
            '~~ nor the message area as a whole does it may be the buttons frame which does:
            Set m_frm = frmBarea
            FrameHeightExeedsThreshold = True
            Exit Function
        End If
    End If ' Buttons area frame is active
                                               
xt: Exit Function
    
eh: If ErrMsg(ErrSrc(PROC)) = vbYes Then: Stop: Resume
End Function

Private Function GetPanesIndex(ByVal rng As Range) As Integer
    Dim sR As Long:            sR = ActiveWindow.SplitRow
    Dim sc As Long:            sc = ActiveWindow.SplitColumn
    Dim r As Long:              r = rng.row
    Dim c As Long:              c = rng.Column
    Dim Index As Integer:   Index = 1

    Select Case True
    Case sR = 0 And sc = 0: Index = 1
    Case sR = 0 And sc > 0 And c > sc: Index = 2
    Case sR > 0 And sc = 0 And r > sR: Index = 2
    Case sR > 0 And sc > 0 And r > sR: If c > sc Then Index = 4 Else Index = 3
    Case sR > 0 And sc > 0 And c > sc: If r > sR Then Index = 4 Else Index = 2
    End Select

    GetPanesIndex = Index
End Function

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

Private Function HeightThreshold(ByVal h_percentage) As Single
    Dim v As Variant
    Dim si As Single
    For Each v In MsectsActive
        si = si + ContentHeight(MsectFrm(v)) - MsectFrmMarginTop
    Next v
    si = si + BareaFrm.Height
    HeightThreshold = si * (h_percentage / 100)
    
End Function

Private Sub IndicateCntrlCaptions(Optional ByVal b As Boolean = True)
' ----------------------------------------------------------------------------
' When False (the default) frame captions are removed, else they remain
' visible as designed.
' ----------------------------------------------------------------------------
            
    Dim ctl As MsForms.Control
       
    If Not b Then
        For Each ctl In Me.Controls
            If TypeName(ctl) = "Frame" Then
                With ctl
                    If .Visible And .Caption <> vbNullString Then
                        .Caption = vbNullString
                    End If
                End With
            End If
        Next ctl
    End If

End Sub

Private Sub Initialize()
' -------------------------------------------------------------------------------
'
' -------------------------------------------------------------------------------
    Const PROC = "Initialize"
    
    On Error GoTo eh
    
    If bSetUpDone Then GoTo xt
    Set dctMonoSpaced = New Dictionary
    Set dctMonoSpacedTbx = New Dictionary
    
    bSetupDoneMonoSpacedSects = False
    bSetupDonePropSpacedSects = False
    bSetupDoneTitle = False
    bFormEvents = False
    SetupDone = False
    siHmarginFrames = 0     ' Ensures proper command buttons framing, may be used for test purpose
    FormHeightOutsideMax = mMsg.ValueAsPt(mMsg.MSG_LIMIT_HEIGHT_MAX_PERCENTAGE, enDsplyDimensionHeight)
    FormHeightOutsideMin = mMsg.ValueAsPt(mMsg.MSG_LIMIT_HEIGHT_MIN_PERCENTAGE, enDsplyDimensionHeight)
    FormWidthOutsideMax = mMsg.ValueAsPt(mMsg.MSG_LIMIT_WIDTH_MAX_PERCENTAGE, enDsplyDimensionWidth)
    FormWidthOutsideMin = mMsg.ValueAsPt(mMsg.MSG_LIMIT_WIDTH_MIN_PERCENTAGE, enDsplyDimensionWidth)
    MsgButtonDefault = 1
    
xt: Exit Sub
    
eh: If ErrMsg(ErrSrc(PROC)) = vbYes Then: Stop: Resume
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
    IsFrameOrForm = TypeOf v Is MsForms.UserForm Or TypeOf v Is MsForms.Frame
End Function

Private Sub laMsgSection1Label_Click():                                                                                                 OnClickActionLabel 1:               End Sub

Private Sub laMsgSection1Label_MouseMove(ByVal Button As Integer, ByVal Shift As Integer, ByVal x As Single, ByVal Y As Single):        CursorForLabelWithOnClickAction 1:  End Sub

Private Sub laMsgSection1Text_Click():                                                                                                  OnClickActionText 1:                End Sub

Private Sub laMsgSection1Text_MouseMove(ByVal Button As Integer, ByVal Shift As Integer, ByVal x As Single, ByVal Y As Single):         CursorForTextWithOnClickAction 1:   End Sub

Private Sub laMsgSection2Label_Click():                                                                                                 OnClickActionLabel 2:               End Sub

Private Sub laMsgSection2Label_MouseMove(ByVal Button As Integer, ByVal Shift As Integer, ByVal x As Single, ByVal Y As Single):        CursorForLabelWithOnClickAction 2:  End Sub

Private Sub laMsgSection2Text_Click():                                                                                                  OnClickActionText 2:                End Sub

Private Sub laMsgSection2Text_MouseMove(ByVal Button As Integer, ByVal Shift As Integer, ByVal x As Single, ByVal Y As Single):         CursorForTextWithOnClickAction 2:   End Sub

Private Sub laMsgSection3Label_Click():                                                                                                 OnClickActionLabel 3:               End Sub

Private Sub laMsgSection3Label_MouseMove(ByVal Button As Integer, ByVal Shift As Integer, ByVal x As Single, ByVal Y As Single):        CursorForLabelWithOnClickAction 3:  End Sub

Private Sub laMsgSection3Text_Click():                                                                                                  OnClickActionText 3:                End Sub

Private Sub laMsgSection3Text_MouseMove(ByVal Button As Integer, ByVal Shift As Integer, ByVal x As Single, ByVal Y As Single):         CursorForTextWithOnClickAction 3:   End Sub

Private Sub laMsgSection4Label_Click():                                                                                                 OnClickActionLabel 4:               End Sub

Private Sub laMsgSection4Label_MouseMove(ByVal Button As Integer, ByVal Shift As Integer, ByVal x As Single, ByVal Y As Single):        CursorForLabelWithOnClickAction 4:  End Sub

Private Sub laMsgSection4Text_Click():                                                                                                  OnClickActionText 4:                End Sub

Private Sub laMsgSection4Text_MouseMove(ByVal Button As Integer, ByVal Shift As Integer, ByVal x As Single, ByVal Y As Single):         CursorForTextWithOnClickAction 4:   End Sub

Private Sub laMsgSection5Label_Click():                                                                                                 OnClickActionLabel 5:               End Sub

Private Sub laMsgSection5Label_MouseMove(ByVal Button As Integer, ByVal Shift As Integer, ByVal x As Single, ByVal Y As Single):        CursorForLabelWithOnClickAction 5:  End Sub

Private Sub laMsgSection5Text_Click():                                                                                                  OnClickActionText 5:                End Sub

Private Sub laMsgSection5Text_MouseMove(ByVal Button As Integer, ByVal Shift As Integer, ByVal x As Single, ByVal Y As Single):         CursorForTextWithOnClickAction 5:   End Sub

Private Sub laMsgSection6Label_Click():                                                                                                 OnClickActionLabel 6:               End Sub

Private Sub laMsgSection6Label_MouseMove(ByVal Button As Integer, ByVal Shift As Integer, ByVal x As Single, ByVal Y As Single):        CursorForLabelWithOnClickAction 6:  End Sub

Private Sub laMsgSection6Text_Click():                                                                                                  OnClickActionText 6:                End Sub

Private Sub laMsgSection6Text_MouseMove(ByVal Button As Integer, ByVal Shift As Integer, ByVal x As Single, ByVal Y As Single):         CursorForTextWithOnClickAction 6:   End Sub

Private Sub laMsgSection7Label_Click():                                                                                                 OnClickActionLabel 7:               End Sub

Private Sub laMsgSection7Label_MouseMove(ByVal Button As Integer, ByVal Shift As Integer, ByVal x As Single, ByVal Y As Single):        CursorForLabelWithOnClickAction 7:  End Sub

Private Sub laMsgSection7Text_Click():                                                                                                  OnClickActionText 7:                End Sub

Private Sub laMsgSection7Text_MouseMove(ByVal Button As Integer, ByVal Shift As Integer, ByVal x As Single, ByVal Y As Single):         CursorForTextWithOnClickAction 7:   End Sub

Private Sub laMsgSection8Label_Click():                                                                                                 OnClickActionLabel 8:               End Sub

Private Sub laMsgSection8Label_MouseMove(ByVal Button As Integer, ByVal Shift As Integer, ByVal x As Single, ByVal Y As Single):        CursorForLabelWithOnClickAction 8:  End Sub

Private Sub laMsgSection8Text_Click():                                                                                                  OnClickActionText 8:                End Sub

Private Sub laMsgSection8Text_MouseMove(ByVal Button As Integer, ByVal Shift As Integer, ByVal x As Single, ByVal Y As Single):         CursorForTextWithOnClickAction 8:   End Sub

Private Function MareaFrm(Optional b_properties As Boolean = True) As MsForms.Frame
' ------------------------------------------------------------------------------
' Returns the Frame of the message area section, created if yet not existing.
' ------------------------------------------------------------------------------
    Const PROC = "MareaFrm"
    
    On Error GoTo eh
    
    If MsectsActive.Count = 0 _
    Then Err.Raise AppErr(1), ErrSrc(PROC), "The Message-Area has been called/requested although there is no Message-Text or -Label specified!"
    If Not MareaFrmExists(MareaFrm) _
    Then Err.Raise AppErr(2), ErrSrc(PROC), "Userform design does not conform with expectations!"
    
    With MareaFrm
        If b_properties Then
            .Visible = True
            .Left = FormMarginLeft
            .BorderStyle = fmBorderStyleNone
            .BorderColor = Me.BackColor
            Select Case .ScrollBars
                Case fmScrollBarsNone
                    .Height = Max(50, ContentHeight(MareaFrm) + AreasFrmMarginTop + AreasFrmMarginBottom)
                    .Width = siAreasFrmWidth
                Case fmScrollBarsHorizontal
                    .Height = Max(50, ContentHeight(MareaFrm) + AreasFrmMarginTop + ScrollHorHeight(MareaFrm))
                Case fmScrollBarsVertical
                    .Width = siAreasFrmWidth + SCROLL_V_WIDTH
                    .ScrollHeight = ContentHeight(MareaFrm) + AreasFrmMarginTop + AreasFrmMarginTop
                Case fmScrollBarsBoth
                    ' width and height are to be considered final
            End Select
        End If
    End With
    VisualizationsForTestOnly MareaFrm, VISUALIZE_CLR_AREA
         
xt: Exit Function

eh: If ErrMsg(ErrSrc(PROC)) = vbYes Then: Stop: Resume
End Function

Private Function MareaFrmExists(ByRef m_frm As MsForms.Frame) As Boolean
' ------------------------------------------------------------------------------
'
' ------------------------------------------------------------------------------
    If Not frmMarea Is Nothing Then
        Set m_frm = frmMarea
        MareaFrmExists = True
    End If
End Function

Private Function MareaFrmIsActive(Optional ByRef m_frm As MsForms.Frame) As Boolean
' ------------------------------------------------------------------------------
'
' ------------------------------------------------------------------------------
    If MareaFrmExists(m_frm) Then
        If m_frm.Visible Then
            Set m_frm = frmMarea
            MareaFrmIsActive = True
        End If
    End If
End Function

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

Private Function MaxWidthSectTxtBox(ByVal frm_text As MsForms.Frame) As Single
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
    MonitorSetupTextProperties tbxStep, enMonStep
    With tbxStep
        .Top = ms_top
        .Left = 0
        .Visible = True
        .Height = 12
        .Width = Me.InsideWidth
        .BorderStyle = fmBorderStyleNone
        .BorderColor = &H80000005
        .SpecialEffect = fmSpecialEffectFlat
         ms_top = AdjustToVgrid(.Top + .Height)
    End With
    VisualizationsForTestOnly tbxStep, VISUALIZE_CLR_MON_STEPS_FRM
    Qenqueue cllSteps, tbxStep
    TimedDoEvents ErrSrc(PROC)
    
End Sub

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
                .BorderStyle = fmBorderStyleNone
                .BorderColor = &H80000005
                .SpecialEffect = fmSpecialEffectFlat
            End With
            VisualizationsForTestOnly tbxFooter, VISUALIZE_CLR_MSEC_TEXT_FRM
        End If
        MonitorSetupTextProperties tbxFooter, enMonFooter
        With tbxFooter
            .Top = AdjustToVgrid(frmSteps.Top + frmSteps.Height + 6)
            .Value = TextMonitorFooter.Text
        End With
        FormHeightInside = ContentHeight(tbxFooter.Parent) + FormMarginTop + FormMarginBottom
    End If

xt: Exit Sub

eh: If ErrMsg(ErrSrc(PROC)) = vbYes Then: Stop: Resume
End Sub

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
            .BorderStyle = fmBorderStyleNone
            .BorderColor = &H80000005
            .SpecialEffect = fmSpecialEffectFlat
        End With
        VisualizationsForTestOnly tbxHeader, VISUALIZE_CLR_MSEC_TEXT_FRM
    End If
    MonitorSetupTextProperties tbxHeader, enMonHeader
    If TextMonitorHeader.MonoSpaced Then
        AutoSizeTextBox a_tbx:=tbxHeader _
                      , a_text:=TextMonitorHeader.Text _
                      , a_width_limit:=0
    Else
        AutoSizeTextBox a_tbx:=tbxHeader _
                      , a_text:=TextMonitorHeader.Text _
                      , a_width_limit:=Me.InsideWidth - FormMarginLeft - FormMarginRight
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
    FormHeightInside = ContentHeight(frmSteps.Parent) + FormMarginTop + FormMarginBottom

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
    Dim ctl                 As MsForms.Control
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
            .Width = siFormWidthInsideMin + .Width - .InsideWidth
                        
            '~~ Establish the number of visualized monitor steps in its dedicated frame
            Set frmSteps = AddControl(ac_ctl:=Frame, ac_name:="frMonitorSteps")
            With frmSteps
                .BorderColor = Me.BackColor
                .BorderStyle = fmBorderStyleNone
                .SpecialEffect = fmSpecialEffectFlat
                .Top = siTop
                .Visible = True
                .Width = Me.InsideWidth
                siTop = 0
                For i = 1 To lMonitorStepsDisplayed
                    MonitorEstablishStep siTop
                Next i
                .Height = ContentHeight(frmSteps, False)
                '~~ The maximum height for the steps frame is the max formheight minus the height of header and footer
                siNetHeight = ContentHeight(frmSteps.Parent) - frmSteps.Height
                siStepsHeightMax = Me.FormHeightOutsideMax - siNetHeight
                NewHeight n_height:=Min(siStepsHeightMax, .Height) _
                        , n_frame:=frmSteps _
                        , n_for_visible_only:=False
            End With
            VisualizationsForTestOnly frmSteps, VISUALIZE_CLR_MON_STEPS_FRM
            NewHeight n_height:=Min(.FormHeightOutsideMax, ContentHeight(frmSteps.Parent)) _
                    , n_frame:=frmSteps.Parent
            NewWidth(frmSteps) = Min(siFormWidthInsideMax, ContentWidth(frmSteps.Parent))
        End With
        bMonitorInitialized = True
    End If

xt: Exit Sub

eh: If ErrMsg(ErrSrc(PROC)) = vbYes Then: Stop: Resume
End Sub

Private Sub MonitorSetupTextProperties(ByVal ctl As MsForms.Control, _
                                       ByVal kind_of_text As KindOfText)
' ------------------------------------------------------------------------------
' Setup the Font properties for a Label or TextBox (ctl) according to the
' corresponding udtMsgText type (kind_of_text).
' ------------------------------------------------------------------------------

    Dim Txt As udtMsgText:  Txt = MsgText(kind_of_text)
    
    With ctl.Font
        If .Bold <> Txt.FontBold Then .Bold = Txt.FontBold
        If .Italic <> Txt.FontItalic Then .Italic = Txt.FontItalic
        If .Underline <> Txt.FontUnderline Then .Underline = Txt.FontUnderline
        If Txt.MonoSpaced Then
            .Name = DFLT_TXT_MONOSPACED_FONT_NAME
            If Txt.FontSize = 0 _
            Then .Size = DFLT_TXT_MONOSPACED_FONT_SIZE _
            Else .Size = Txt.FontSize
        Else
            If Txt.FontName = vbNullString _
            Then .Name = DFLT_TXT_PROPSPACED_FONT_NAME _
            Else .Name = Txt.FontName
            If Txt.FontSize = 0 _
            Then .Size = DFLT_TXT_PROPSPACED_FONT_SIZE _
            Else .Size = Txt.FontSize
        End If
    End With
    ctl.ForeColor = Txt.FontColor
    If bVisualizeForTest Then ctl.BackColor = VISUALIZE_CLR_MSEC_TEXT_FRM
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
    Dim tbx                 As MsForms.TextBox
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
            MonitorSetupTextProperties tbx, enMonStep
            With tbx
                .Visible = True
                .Top = siTop
                .BorderStyle = fmBorderStyleNone
                .BorderColor = &H80000005
                .SpecialEffect = fmSpecialEffectFlat
            End With
            
            If TextMonitorStep.MonoSpaced Then
                AutoSizeTextBox a_tbx:=tbx _
                              , a_text:=TextMonitorStep.Text _
                              , a_width_limit:=0
            Else
                AutoSizeTextBox a_tbx:=tbx _
                              , a_width_limit:=Me.InsideWidth _
                              , a_text:=TextMonitorStep.Text
            End If
            MonitorStepsAdjustTopPosition
            NewWidth(frmSteps, False) = Min(siFormWidthInsideMax, ContentWidth(frmSteps, False)) ' applies a horizontal scroll-bar or adjust its width
            NewWidth(frmSteps.Parent) = ContentWidth(frmSteps.Parent)
            
            siNetHeight = Me.Height - (frmSteps.Height - frmSteps.Top) - HeightDiffFormOutIn
            NewHeight n_height:=Min(MonitorHeightMaxSteps, ContentHeight(frmSteps, False)) _
                    , n_frame:=frmSteps _
                    , n_for_visible_only:=False _
                    , n_y_action:=fmScrollActionBegin
            
            lStepsDisplayed = lStepsDisplayed + 1
            
            If Not tbxFooter Is Nothing Then
                tbxFooter.Top = AdjustToVgrid(frmSteps.Top + frmSteps.Height + 6)
                FormHeightInside = tbxFooter.Top + tbxFooter.Height + FormMarginTop + FormMarginBottom
            Else
                FormHeightInside = ContentHeight(frmSteps.Parent) + FormMarginTop + FormMarginBottom
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
                AutoSizeTextBox a_tbx:=tbx _
                              , a_text:=TextMonitorStep.Text _
                              , a_width_limit:=0
            Else
                AutoSizeTextBox a_tbx:=tbx _
                              , a_width_limit:=Me.InsideWidth _
                              , a_text:=TextMonitorStep.Text
            End If
            MonitorStepsAdjustTopPosition
            NewWidth(frmSteps, False) = Min(siFormWidthInsideMax, ContentWidth(frmSteps, False)) ' applies a horizontal scroll-bar or adjust its width
            NewWidth(frmSteps.Parent) = ContentWidth(frmSteps.Parent) + 20
            
            siNetHeight = Me.Height - (frmSteps.Height - frmSteps.Top)
            NewHeight n_height:=Min(MonitorHeightMaxSteps, ContentHeight(frmSteps, False)) + ScrollHorHeight(frmSteps) _
                    , n_frame:=frmSteps _
                    , n_for_visible_only:=False _
                    , n_y_action:=fmScrollActionEnd
        
            If Not tbxFooter Is Nothing Then
                tbxFooter.Top = AdjustToVgrid(frmSteps.Top + frmSteps.Height + 6)
                FormHeightInside = tbxFooter.Top + tbxFooter.Height + FormMarginTop + FormMarginBottom
            Else
                FormHeightInside = ContentHeight(frmSteps.Parent) + FormMarginTop + FormMarginBottom
            End If
        End If
    End If
        
    TimedDoEvents ErrSrc(PROC)
    NewWidth(frmSteps) = Min(siFormWidthInsideMax, ContentWidth(frmSteps.Parent) + 15)
    FormHeightInside = ContentHeight(frmSteps.Parent) + FormMarginTop + FormMarginBottom
    
xt: Exit Sub

eh: If ErrMsg(ErrSrc(PROC)) = vbYes Then: Stop: Resume
End Sub

Private Sub MonitorStepsAdjustTopPosition()
' ------------------------------------------------------------------------------
' - Adjusts each visible control's top position considering its current height.
' ------------------------------------------------------------------------------
    Const PROC = "MonitorStepsAdjustTopPosition"
    
    On Error GoTo eh
    Dim siTop   As Single
    Dim ctl     As MsForms.Control
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

Private Function MsectFrameHeight(ByVal m_sect As Long)
    Dim frm As MsForms.Frame:   Set frm = MsectFrm(m_sect)

    If Not ScrollVerApplied(frm) _
    Then MsectFrameHeight = ContentHeight(frm) _
    Else MsectFrameHeight = frm.Height
        
End Function

Private Function MsectFrm(ByVal m_sect As Long, _
                 Optional ByVal m_properties As Boolean = True) As MsForms.Frame
' ------------------------------------------------------------------------------
' Returns a message section's (m_sect) Frame with the properties .Top and .Width
' already set.
' ------------------------------------------------------------------------------
    Const PROC = "MsectFrm"
    
    On Error GoTo eh
    Dim si As Single
    
    If Not MsectFrmExists(m_sect, MsectFrm) _
    Then Err.Raise AppErr(1), ErrSrc(PROC), "Userform design does not conform with expectations!"
    
    If m_properties Then
        With MsectFrm
            .Visible = True
            .Top = MsectFrmTop(m_sect)
            .Left = AreasFrmMarginLeft
            .BorderStyle = fmBorderStyleNone
            .BorderColor = Me.BackColor
            .Font.Name = DFLT_TXT_PROPSPACED_FONT_NAME
            .Font.Size = DFLT_TXT_PROPSPACED_FONT_SIZE
            DoEvents
            Select Case .ScrollBars
                Case fmScrollBarsNone
                    .Width = siMsectFrmWidth
                    si = ContentHeight(MsectFrm) + MsectFrmMarginTop + MsectFrmMarginBottom
                    If si > MsectFrmMarginTop + MsectFrmMarginBottom _
                    Then .Height = si
                Case fmScrollBarsHorizontal
                    ' width is considered final once a horizontal scroll-bar is applied
                    si = ContentHeight(MsectFrm) + MsectFrmMarginTop + MsectFrmMarginBottom
                    If si > MsectFrmMarginTop + MsectFrmMarginBottom _
                    Then .Height = si
                Case fmScrollBarsVertical
                    .Width = siMsectFrmWidth
                    ' height is considered final once a vertical scroll-bar is applied
                Case fmScrollBarsBoth
                    ' width and height is considered final once both scroll-bars are applied
            End Select
            .Font.Size = MsectLabelFontSize(m_sect)
            DoEvents
        End With
        DoEvents
        VisualizationsForTestOnly MsectFrm, VISUALIZE_CLR_MSEC_FRM
    End If

xt: Exit Function
    
eh: If ErrMsg(ErrSrc(PROC)) = vbYes Then: Stop: Resume
End Function

Private Function MsectFrmAbove(ByVal m_sect As Long, _
                               ByRef m_sect_frm As MsForms.Frame) As Long
' ------------------------------------------------------------------------------
' Returns a message section's (m_sect) frame above section number (m_sect_above)
' when the provided section (m_sect) is the first one active the function
' returns None.
' ------------------------------------------------------------------------------
    Dim i   As Variant
    Dim cll As Collection:  Set cll = MsectsActive
    
    For i = 1 To cll.Count
        If cll(i) = m_sect And i > 1 Then
            MsectFrmAbove = cll(i - 1)
            Set m_sect_frm = MsectFrm(MsectFrmAbove)
            Exit For
        End If
    Next i
    
End Function

Private Function MsectFrmExists(ByVal m_sect As Long, _
                                ByRef m_frm As MsForms.Frame) As Boolean
    
    If dctMsectFrm.Exists(m_sect) Then
        Set m_frm = dctMsectFrm(m_sect)
        MsectFrmExists = True
    End If
    
End Function

Private Function MsectFrmHeight(ByVal m_sect As Long) As Single
    Dim frm As MsForms.Frame
    
    Set frm = MsectFrm(m_sect)
    MsectFrmHeight = frm.Height
    '~~ When no vertical scroll-bar is applied the height depends on the content's height
    If Not ScrollVerApplied(frm) _
    Then MsectFrmHeight = ContentHeight(frm) + MsectFrmMarginTop + MsectFrmMarginBottom

End Function

Private Function MsectFrmIsActive(ByVal m_sect As Long, _
                         Optional ByRef m_frm As MsForms.Frame) As Boolean
    
    If dctMsectFrm.Exists(m_sect) Then
        Set m_frm = dctMsectFrm(m_sect)
        MsectFrmIsActive = m_frm.Visible
    End If
                                     
End Function

Private Function MsectFrmTop(ByVal m_sect As Long) As Single
' ------------------------------------------------------------------------------
' Returns the top  position of a message section's frame (m_sect) which is
' FormMarginTop when it is the top most section or the top position is based on
' the vertical distance to the above message section's frame.
' ------------------------------------------------------------------------------
    Dim frm     As MsForms.Frame
    Dim iAbove  As Long
    
    MsectFrmTop = FormMarginTop
    iAbove = MsectFrmAbove(m_sect, frm)
    If iAbove <> 0 Then
        MsectFrmTop = AdjustToVgrid(frm.Top + ContentHeight(frm) + 8, , 8)
    End If
    
End Function

Private Function MsectFrmWidthFinal() As Single
    Dim v   As Variant
    Dim si  As Single
    
    For Each v In MsectsActive
        si = Max(si, MsectFrm(v).Width)
    Next v
    MsectFrmWidthFinal = si + 4
    
End Function

Private Function MsectFrmWidthLabel(ByVal m_sect As Long) As Single
    
    If MsectTxtWithLftPosLbl(m_sect) _
    Then MsectFrmWidthLabel = siMsectLabelWidthAll _
    Else MsectFrmWidthLabel = siMsectFrmWidth - MsectFrmMarginLeft - MsectFrmMarginRight
    
End Function

Private Function MsectFrmWidthLabelLimit(ByVal m_sect As Long) As Single
    
    If MsectHasOnlyLabel(m_sect) Or MsectLabelAbove(m_sect) _
    Then MsectFrmWidthLabelLimit = siMsectFrmWidth - MsectFrmMarginLeft _
    Else MsectFrmWidthLabelLimit = siMsectLabelWidthAll
    
End Function

Private Function MsectHasLabelAndText(ByVal m_sect As Long) As Boolean

    If Msg.Section(m_sect).Label.Text <> vbNullString _
    And Msg.Section(m_sect).Text.Text <> vbNullString Then
        MsectHasLabelAndText = True
    End If
    
End Function

Private Function MsectHasLabelLeft(ByVal m_sect As Single) As Boolean
    MsectHasLabelLeft = MsectHasLabelAndText(m_sect) And MsectLabelPos(m_sect) <> enLabelAboveSectionText
End Function

Private Function MsectHasOnlyLabel(ByVal m_sect As Long) As Boolean
    MsectHasOnlyLabel = Msg.Section(m_sect).Label.Text <> vbNullString _
                    And Msg.Section(m_sect).Text.Text = vbNullString
End Function

Private Function MsectHasOnlyText(ByVal m_sect As Long) As Boolean
    MsectHasOnlyText = Msg.Section(m_sect).Label.Text = vbNullString _
                   And Msg.Section(m_sect).Text.Text <> vbNullString
End Function

Private Function MsectLabel(ByVal m_sect As Long, _
                   Optional ByVal m_properties As Boolean = True) As MsForms.Label
' ------------------------------------------------------------------------------
' Returns the Label of the message section (m_sect), created in the
' corresponding MsectFrm when not yet existing.
' ------------------------------------------------------------------------------
    Const PROC      As String = "MsectLabel"
    
    On Error GoTo eh
    If Not MsectLabelExists(m_sect, MsectLabel) _
    Then Err.Raise AppErr(1), ErrSrc(PROC), "Userform design does not conform with expectations!"
    
    If m_properties Then
        With MsectLabel
            .Visible = True
            .Top = MsectFrmMarginTop
            .BorderStyle = fmBorderStyleNone
            .BorderColor = Me.BackColor
            If MsectHasLabelLeft(m_sect) _
            Then .Top = MsectLabelTop
            .Width = MsectFrmWidthLabel(m_sect)
            .Font.Size = MsectLabelFontSize(m_sect)
            DoEvents
        End With
        DoEvents
        VisualizationsForTestOnly MsectLabel, VISUALIZE_CLR_MSEC_LBL
    End If
       
xt: Exit Function
    
eh: If ErrMsg(ErrSrc(PROC)) = vbYes Then: Stop: Resume
End Function

Private Function MsectLabelAlign(ByVal m_sect As Long) As Long
    If MsectLabelAbove(m_sect) _
    Or MsectHasOnlyLabel(m_sect) Then
        MsectLabelAlign = fmTextAlignLeft
    Else
        Select Case LabelAllPos
            Case enLposLeftAlignedRight
                MsectLabelAlign = fmTextAlignRight
            Case enLposLeftAlignedCenter
                MsectLabelAlign = fmTextAlignCenter
            Case enLposLeftAlignedLeft
                MsectLabelAlign = fmTextAlignLeft
        End Select
    End If
End Function

Private Function MsectLabelExists(ByVal m_sect As Long, _
                                  ByRef m_lbl As MsForms.Label) As Boolean
                                
    If dctMsectLabel.Exists(m_sect) Then
        Set m_lbl = dctMsectLabel(m_sect)
        MsectLabelExists = True
    End If
                                
End Function

Private Function MsectLabelFontName(ByVal m_sect As Long) As String
    
    With Msg.Section(m_sect).Label
        If .MonoSpaced Then
            If .FontName <> vbNullString _
            Then MsectLabelFontName = .FontName _
            Else MsectLabelFontName = DFLT_TXT_MONOSPACED_FONT_NAME
        Else
            If .FontName <> vbNullString _
            Then MsectLabelFontName = .FontName _
            Else MsectLabelFontName = DFLT_TXT_PROPSPACED_FONT_NAME
        End If
    End With

End Function

Private Function MsectLabelFontSize(ByVal m_sect As Long) As Single
    MsgSectLbl = MsgLabel(m_sect)
    
    With Msg.Section(m_sect).Label
        If .MonoSpaced Then
            If .FontSize <> 0 _
            Then MsectLabelFontSize = .FontSize _
            Else MsectLabelFontSize = DFLT_TXT_MONOSPACED_FONT_SIZE
        Else
            If .FontSize <> 0 _
            Then MsectLabelFontSize = .FontSize _
            Else MsectLabelFontSize = DFLT_TXT_PROPSPACED_FONT_SIZE
        End If
    End With
    
End Function

Private Function MsectLabelIsActive(ByVal m_sect As Long, _
                           Optional ByRef m_sect_lbl As MsForms.Label, _
                           Optional ByRef m_sect_frm As MsForms.Frame) As Boolean
    
    Dim frm As MsForms.Frame
    Dim lbl As MsForms.Label
    
    Set frm = MsectFrm(m_sect)
    If frm.Visible Then
        Set m_sect_frm = frm
        Set lbl = MsectLabel(m_sect)
        If lbl.Visible Then
            Set m_sect_lbl = lbl
            MsectLabelIsActive = True
        End If
    End If

End Function

Private Function MsectLabelPos(ByVal m_sect As Long) As enLabelPos
    If MsectHasOnlyLabel(m_sect) _
    Then MsectLabelPos = enLabelAboveSectionText _
    Else MsectLabelPos = LabelAllPos
End Function

Private Function MsectLabelTbx(ByVal m_sect As Long) As MsForms.TextBox
    Const PROC = "MsectLabelTbx"
    
    On Error Resume Next
    Set MsectLabelTbx = MsectFrm(m_sect).Controls("tbMsgSection" & m_sect & "Label")
    If Err.Number <> 0 _
    Then Err.Raise AppErr(1), ErrSrc(PROC), "Userform design does not conform with expectations! " & vbLf & _
                                            "(section provided = " & m_sect & ")"

xt: Exit Function
    
eh: If ErrMsg(ErrSrc(PROC)) = vbYes Then: Stop: Resume
End Function

Private Function MsectLabelTbxHeight(ByVal m_sect As Long) As Single:   MsectLabelTbxHeight = MsectLabelTbx(m_sect).Height: End Function

Private Function MsectLabelText(ByVal m_sect) As String:                MsectLabelText = MsgLabel(m_sect).Text:             End Function

Private Sub MsectsActivate()
'------------------------------------------------------------------------------
' Activates all message sections (.Visible = True) for wich a section
' Label or a section Text has been provided and collects the corresponding
' section number in a collection. This collection ensures that only definitely
' provided sections are collected disregarding any gaps (e.g. 3,4,7,9). This
' approach allows "conditional sections" which may or may not be displayed
' depending on conditions.
' ------------------------------------------------------------------------------
    Dim i As Long
    
    For i = 1 To lMaxNoOfMsgSects
        If Msg.Section(i).Label.Text <> vbNullString _
        Or Msg.Section(i).Text.Text <> vbNullString _
        Then cllMsectsActive.Add i
    Next i
    
End Sub

Private Function MsectsActive() As Collection: Set MsectsActive = cllMsectsActive:  End Function

Private Function MsectTextFontName(ByVal m_sect As Long) As String
    MsgSectTxt = MsgText(enSectText, m_sect)

    With Msg.Section(m_sect).Text
        If .MonoSpaced Then
            If .FontName <> vbNullString _
            Then MsectTextFontName = .FontName _
            Else MsectTextFontName = DFLT_TXT_MONOSPACED_FONT_NAME
        Else
            If .FontName <> vbNullString _
            Then MsectTextFontName = .FontName _
            Else MsectTextFontName = DFLT_TXT_PROPSPACED_FONT_NAME
        End If
    End With

End Function

Private Function MsectTextFontSize(ByVal m_sect As Long) As Single
    MsgSectTxt = MsgText(enSectText, m_sect)

    With Msg.Section(m_sect).Text
        If .MonoSpaced Then
            If .FontSize <> 0 _
            Then MsectTextFontSize = .FontSize _
            Else MsectTextFontSize = DFLT_TXT_MONOSPACED_FONT_SIZE
        Else
            If .FontSize <> 0 _
            Then MsectTextFontSize = .FontSize _
            Else MsectTextFontSize = DFLT_TXT_PROPSPACED_FONT_SIZE
        End If
    End With

End Function

Private Function MsectTextFrm(ByVal m_sect As Long, _
                     Optional ByVal m_properties As Boolean = True) As MsForms.Frame
' ------------------------------------------------------------------------------
' Returns the frame of the TextBox of the section (m_sect), created in the
' corresponding MsectFrm when not yet existing. The Frame's top
' position is 0 or, when there is a visible above Label underneath it.
' ------------------------------------------------------------------------------
    Const PROC = "MsectTextFrm"
    
    On Error GoTo eh
    Dim si As Single
    If Not MsectTextFrmExists(m_sect, MsectTextFrm) _
    Then Err.Raise AppErr(1), ErrSrc(PROC), "Userform design does not conform with expectations! " & vbLf & _
                                            "(section provided = " & m_sect & ")"
    If MsectHasOnlyLabel(m_sect) _
    Then Err.Raise AppErr(2), ErrSrc(PROC), "Obtaining a section's text frame when it has only a Label is regarded a severe logic error!"

    
    If m_properties Then
        With MsectTextFrm
            .Visible = True
            .Top = MsectFrmMarginTop
            .Left = MsectTextFrmLeft(m_sect)
            .BorderStyle = fmBorderStyleNone
            .BorderColor = Me.BackColor
            
            Select Case .ScrollBars
                Case fmScrollBarsNone
                    .Width = MsectTextFrmWidth(m_sect)
                    si = ContentHeight(MsectTextFrm) + MsectFrmMarginTop + MsectFrmMarginBottom
                    If si > MsectFrmMarginTop + MsectFrmMarginBottom Then
                        .Height = si
                    End If
                Case fmScrollBarsHorizontal
                    si = ContentHeight(MsectTextFrm) + MsectFrmMarginTop + SCROLL_H_HEIGHT
                    If si > MsectFrmMarginTop + MsectFrmMarginBottom Then
                        .Height = si
                    End If
                    '~~ The following is just a logic assertion!
                    If .Width <> MsectTextFrmMax(m_sect) - MsectFrmMarginRight _
                    And .Width <> MsectTextFrmWidth(m_sect) _
                    Then Err.Raise (3), ErrSrc(PROC), "There's a horizontal scroll-bar applied for the message text " & _
                                                     "frame but the frame's width is not the specified maximum width!"
                Case fmScrollBarsVertical
                    .Width = MsectTextFrmMax(m_sect) + ScrollVerWidth(MsectTextFrm)
                    ' height muts no longer be altered
                Case fmScrollBarsBoth
                    ' width and height must no longer be altered
            End Select
            .Font.Name = MsectTextFontName(m_sect)
            DoEvents
            .Font.Size = MsectTextFontSize(m_sect)
            DoEvents
        End With
        DoEvents
        VisualizationsForTestOnly MsectTextFrm, VISUALIZE_CLR_MSEC_TEXT_FRM
    End If

xt: Exit Function
    
eh: If ErrMsg(ErrSrc(PROC)) = vbYes Then: Stop: Resume
End Function

Private Function MsectTextFrmExists(ByVal m_sect As Long, _
                                    ByRef m_frm As MsForms.Frame) As Boolean
                                
    If dctMsectTextFrm.Exists(m_sect) Then
        Set m_frm = dctMsectTextFrm(m_sect)
        MsectTextFrmExists = True
    End If
                                
End Function

Private Function MsectTextFrmHeight(ByVal m_sect As Long)
    Dim frm As MsForms.Frame
    
    Set frm = MsectTextFrm(m_sect)
    If Not ScrollVerApplied(frm) Then
        MsectTextFrmHeight = ContentHeight(frm) + ScrollHorHeight(frm)
    Else
        MsectTextFrmHeight = frm.Height
    End If
End Function

Private Function MsectTextFrmIsActive(ByVal m_sect As Long, _
                             Optional ByRef m_sect_frm As MsForms.Frame, _
                             Optional ByRef m_text_frm As MsForms.Frame) As Boolean
' ------------------------------------------------------------------------------
'
' ------------------------------------------------------------------------------
    Dim frmSect As MsForms.Frame
    Dim frmText As MsForms.Frame
    
    Set frmSect = MsectFrm(m_sect)
    If frmSect.Visible Then
        Set m_sect_frm = frmSect
        Set frmText = MsectTextFrm(m_sect)
        If frmText.Visible Then
            Set m_text_frm = frmText
            MsectTextFrmIsActive = True
        End If
    End If

End Function

Private Function MsectTextFrmLeft(m_sect) As Single
' ------------------------------------------------------------------------------
' When the Label is left of the text positioned and a Label had been specified
' by the caller the left pos of the Text frame is intended by the Label width,
' else the Text is positioned at left = 0.
' ------------------------------------------------------------------------------
    If MsectHasLabelLeft(m_sect) _
    Then MsectTextFrmLeft = siMsectLabelWidthAll + MsectLabelTextMargin

End Function

Private Function MsectTextFrmMax(ByVal m_sect As Long) As Single
    MsectTextFrmMax = siMsectFrmWidthMax - MsectTextFrmLeft(m_sect) - MsectFrmMarginRight
End Function

Private Function MsectTextFrmTop() As Single:                   MsectTextFrmTop = 0:                                            End Function

Private Function MsectTextLbl(ByVal m_sect As Long, _
                     Optional ByVal m_properties As Boolean = True) As MsForms.Label
' ------------------------------------------------------------------------------
' Returns the Label control representing a message section's (m_sect) Text.
' ------------------------------------------------------------------------------
    Const PROC = "MsectTextLbl"
    
    On Error GoTo eh
    If Not MsectTextLblExists(m_sect, MsectTextLbl) _
    Then Err.Raise AppErr(1), ErrSrc(PROC), "Userform design does not conform with expectations!"
    
    If m_properties Then
        With MsectTextLbl
            .Top = MsectFrmMarginTop
            .Left = MsectFrmMarginLeft
            .Font.Name = MsectTextFontName(m_sect)
            DoEvents
            .Font.Size = MsectTextFontSize(m_sect)
            DoEvents
        End With
        DoEvents
        VisualizationsForTestOnly MsectTextLbl, VISUALIZE_CLR_MSEC_LBL
    End If
    If MsectHasLabelAndText(m_sect) Then
        MsectTextFrm m_sect ' provides a displayed text frame
    End If
    MsectFrm m_sect ' provides a displayed section frame
    
xt: Exit Function
    
eh: If ErrMsg(ErrSrc(PROC)) = vbYes Then: Stop: Resume
End Function

Private Function MsectTextLblExists(ByVal m_sect As Long, _
                                    ByRef m_lbl As MsForms.Label) As Boolean
' ------------------------------------------------------------------------------
'
' ------------------------------------------------------------------------------
                                   
    If dctMsectTextLbl.Exists(m_sect) Then
        Set m_lbl = dctMsectTextLbl(m_sect)
        MsectTextLblExists = True
    End If
                                
End Function

'Private Function MsectTextMonospaced(ByVal m_sect) As String:   MsectTextMonospaced = Msg.Section(m_sect).Text.MonoSpaced:   End Function

Private Function MsectTextSetupDone(ByVal m_sect As Long) As Boolean
    MsectTextSetupDone = dctSectTextSetup.Exists(m_sect)
End Function

'Private Function MsectTextText(ByVal m_sect) As String:         MsectTextText = MsgText(enSectText, m_sect).Text:               End Function

Private Function MsectTextWidthLimit(ByVal m_sect As Long) As Single
' ------------------------------------------------------------------------------
' ! The width limit exclusively applies for the setup of prop-spaced section !
' ! text.                                                                    !
' ------------------------------------------------------------------------------

    If MsectHasLabelAndText(m_sect) Then
        MsectTextWidthLimit = siMsectTextFrmWidthWithLposLbl ' - MsectTextFrmMarginLeft - MsectTextFrmMarginRight
    Else
        MsectTextWidthLimit = siMsectTextFrmWidthOnlyText ' - MsectTextFrmMarginLeft - MsectTextFrmMarginRight
    End If
    
End Function

Private Function MsectTxtWithLftPosLbl(ByVal m_sect As Long) As Boolean
    
    MsectTxtWithLftPosLbl = Msg.Section(m_sect).Label.Text <> vbNullString _
                        And Msg.Section(m_sect).Text.Text <> vbNullString _
                        And MsectLabelPos(m_sect) <> enLabelAboveSectionText
    
End Function

Private Function NewDict(ByRef dct As Dictionary) As Dictionary
' ------------------------------------------------------------------------------
' Returns the Dictionary (dct), getting rid of an old.
' ------------------------------------------------------------------------------
    Set dct = Nothing
    Set dct = New Dictionary
    Set NewDict = dct
End Function

Private Sub NewHeight(ByVal n_height As Single, _
             Optional ByRef n_frame As Object, _
             Optional ByVal n_for_visible_only As Boolean = True, _
             Optional ByVal n_y_action As fmScrollAction = fmScrollActionBegin, _
             Optional ByVal n_threshold_height_diff As Single = 5, _
             Optional ByRef n_scroll_applied As Boolean = False)
' ------------------------------------------------------------------------------
' Mimics a height change event. Applies a vertical scroll-bar when the content
' height of the frame or form (n_frame) is greater than the height of
' the frame or form by considering a threshold (n_threshold_height_diff) in
' order to avoid a usesless scroll-bar for a redicolous height difference. In
' case the new height is less the the frame's height a vertical scrollbar is
' removed.
' ------------------------------------------------------------------------------
    Const PROC = "NewHeight"

    On Error GoTo eh
    Dim siContentHeight     As Single

    siContentHeight = ContentHeight(n_frame, n_for_visible_only)

    If n_frame Is Nothing Then Err.Raise AppErr(1), ErrSrc(PROC), "The required argument 'n_frame' is Nothing!"
    If Not IsFrameOrForm(n_frame) _
    Then Err.Raise AppErr(1), ErrSrc(PROC), "The provided argument is neither a Frame nor a Form!"

    n_frame.Height = n_height

    If siContentHeight - (n_frame.Height - ScrollHorHeight(n_frame)) > n_threshold_height_diff Then
        ScrollVerApply s_frame:=n_frame _
                     , s_content_height:=siContentHeight _
                     , s_y_action:=n_y_action _
                     , s_applied:=n_scroll_applied
'        n_frame.Width = n_frame.Width + SCROLL_V_WIDTH
    ElseIf ScrollVerApplied(n_frame) Then
        ScrollVerRemove n_frame
    End If

xt: Exit Sub

eh: If ErrMsg(ErrSrc(PROC)) = vbYes Then: Stop: Resume
End Sub

Private Sub OnClickActionLabel(ByVal o_section As Long)
    Dim sItem As String
    sItem = MsgLabel(o_section).OnClickAction
    mMsg.ShellRun sItem, WIN_NORMAL
End Sub

Private Sub OnClickActionText(ByVal o_section As Long)
    Dim sItem As String
    sItem = MsgText(enSectText, o_section).OnClickAction
    mMsg.ShellRun sItem, WIN_NORMAL
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
                .StartUpPosition = 0
                .Left = pos_left
                .Top = pos_top
            End With
        Case IsNumeric(pos_top_left)
            With Me
                .Top = Application.Top + 5
                .Left = Application.Left + 5
            End With
    End Select
    
    '~~ First make sure the bottom right fits,
    '~~ then check if the top-left is still on the screen (which gets priority).
    With Me
        If ((.Left + .Width) > (VirtualScreenLeftPts + VirtualScreenWidthPts)) Then .Left = ((VirtualScreenLeftPts + VirtualScreenWidthPts) - .Width)
        If ((.Top + .Height) > (VirtualScreenTopPts + VirtualScreenHeightPts)) Then .Top = ((VirtualScreenTopPts + VirtualScreenHeightPts) - .Height)
        If (.Left < VirtualScreenLeftPts) Then .Left = VirtualScreenLeftPts Else .Left = pos_left
        If (.Top < VirtualScreenTopPts) Then .Top = VirtualScreenTopPts Else .Top = pos_top
    End With
    
xt: Exit Sub

eh: If ErrMsg(ErrSrc(PROC)) = vbYes Then: Stop: Resume
End Sub

Private Function PrcntgHeightBareaFrm() As Single
    PrcntgHeightBareaFrm = Round(BareaFrm.Height / (MareaFrm.Height + BareaFrm.Height), 2)
End Function

Private Function PrcntgHeightMareaFrm() As Single
    PrcntgHeightMareaFrm = Round(MareaFrm.Height / (MareaFrm.Height + BareaFrm.Height), 2)
End Function

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

Private Function ScrollHorApplied(ByVal s_frame As Variant) As Boolean
' ------------------------------------------------------------------------------
' Returns True when the control (s_frame) has a horizontal scrollbar applied. When
' no control is provided it is the UserForm which is ment.
' ------------------------------------------------------------------------------
    If IsFrameOrForm(s_frame) Then
        Select Case s_frame.ScrollBars
            Case fmScrollBarsBoth, fmScrollBarsHorizontal: ScrollHorApplied = True
        End Select
    End If
End Function

Private Sub ScrollHorApply(ByRef s_frame As Variant, _
                           ByVal s_content_width, _
                  Optional ByVal s_x_action As fmScrollAction = fmScrollActionBegin, _
                  Optional ByRef s_applied_or_changed As Boolean = False)
' ------------------------------------------------------------------------------
' Applies for the frame control (s_frame) a horizontal scroll-bar when
' yet none is applied. In case one is already applied only the scroll's width
' is applied.
' ------------------------------------------------------------------------------
    Const PROC = "ScrollHorApply"
    
    On Error GoTo eh
    Dim si As Single
    
    If Not IsFrameOrForm(s_frame) _
    Then Err.Raise AppErr(1), ErrSrc(PROC), "The provided argument is neither a Frame nor a Form!"
        
    With s_frame
        If Not ScrollHorApplied(s_frame) And s_content_width > s_frame.Width Then
            Select Case .ScrollBars
                Case fmScrollBarsBoth
                    .KeepScrollBarsVisible = fmScrollBarsBoth
                Case fmScrollBarsHorizontal
                    .KeepScrollBarsVisible = fmScrollBarsHorizontal
                Case fmScrollBarsVertical
                    .ScrollBars = fmScrollBarsBoth
                    .KeepScrollBarsVisible = fmScrollBarsBoth
                    s_applied_or_changed = True
                Case fmScrollBarsNone
                    .ScrollBars = fmScrollBarsHorizontal
                    .KeepScrollBarsVisible = fmScrollBarsHorizontal
                    s_applied_or_changed = True
            End Select
        End If
        If ScrollHorApplied(s_frame) Then
            si = ContentWidth(s_frame)
            If .ScrollWidth <> si Then
                .ScrollWidth = si
                s_applied_or_changed = True
            End If
            .Scroll xAction:=s_x_action
'            .Height = ContentHeight(s_frame) + ScrollHorHeight(s_frame)
        Else
            .Height = .Height + ScrollHorHeight(s_frame)
        End If
    End With

xt: Exit Sub
    
eh: If ErrMsg(ErrSrc(PROC)) = vbYes Then: Stop: Resume
End Sub

Private Function ScrollHorHeight(ByVal s_frame As Variant) As Single
    If IsFrameOrForm(s_frame) Then
        If ScrollHorApplied(s_frame) Then ScrollHorHeight = SCROLL_H_HEIGHT
    End If
End Function

Private Sub ScrollHorRemove(ByRef s_frame As Variant)
' ------------------------------------------------------------------------------
' Removes a vertical scroll-bar.
' ------------------------------------------------------------------------------
    Const PROC = "ScrollHorRemove"
    
    On Error GoTo eh
    If Not IsFrameOrForm(s_frame) _
    Then Err.Raise AppErr(1), ErrSrc(PROC), "The provided argument is neither a Frame nor a Form!"
    
    With s_frame
        If ScrollHorApplied(s_frame) Then
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

Private Function ScrollVerApplied(Optional ByVal s_frame As Variant = Nothing) As Boolean
' ------------------------------------------------------------------------------
' Returns True when the control (s_frame) has a vertical scrollbar applied. When no
' control is provided it is the UserForm which is ment.
' ------------------------------------------------------------------------------
    If IsFrameOrForm(s_frame) Then
        Select Case s_frame.ScrollBars
            Case fmScrollBarsBoth, fmScrollBarsVertical: ScrollVerApplied = True
        End Select
    End If
End Function

Private Sub ScrollVerApply(ByRef s_frame As Variant, _
                           ByVal s_content_height As Single, _
                  Optional ByVal s_y_action As fmScrollAction = fmScrollActionBegin, _
                  Optional ByRef s_applied As Boolean = False)
' ------------------------------------------------------------------------------
' Aplies for the Frame (s_frame) a vertical scroll-bar when yet none is
' applied and adjusts/increases the width accordingly. The scroll-bar's height
' is adjusted by considering an already displayed horizontal scroll-bar.
' ------------------------------------------------------------------------------
    Const PROC = "ScrollVerApply"
    
    On Error GoTo eh
    If Not IsFrameOrForm(s_frame) _
    Then Err.Raise AppErr(1), ErrSrc(PROC), "The provided argument is neither a Frame nor a Form!"
    
    With s_frame
        If Not ScrollVerApplied(s_frame) Then
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
            s_applied = True
        End If
        .Scroll yAction:=s_y_action
        .ScrollHeight = s_content_height
        
    End With

xt: Exit Sub
    
eh: If ErrMsg(ErrSrc(PROC)) = vbYes Then: Stop: Resume
End Sub

Private Sub ScrollVerForHeightExceedingFrames(ByVal exceeding_height As Single, _
                                              ByRef s_finalized As Boolean)
' ------------------------------------------------------------------------------
' Either because the message area occupies 60% or more of the total height or
' because both, the message area and the buttons area us about the same height,
' it - or only the section text occupying 65% or more - will be reduced by the
' exceeding height amount (exceeding_height) and will get a vertical scrollbar.
' ------------------------------------------------------------------------------
    Const PROC = "ScrollVerForHeightExceedingFrames"
    
    On Error GoTo eh
    Dim bScrollApplied      As Boolean
    Dim frm                 As MsForms.Frame
    Dim i                   As Long
    Dim siExceedingHeight   As Single
    Dim siPercentageBarea   As Single
    Dim siPercentageMarea   As Single
    
    siExceedingHeight = Me.InsideHeight - FormHeightInsideMax
    
    If siExceedingHeight > 0 Then
        If MsectsActive.Count = 1 Then
            Set frm = MsectTextFrm(MsectsActive(1))
            If Not ScrollVerApplied(frm) Then
                NewHeight n_height:=frm.Height - siExceedingHeight _
                        , n_frame:=frm
                s_finalized = False
                bScrollApplied = True
            End If
            GoTo xt
            
        End If
        
        If FrameHeightExeedsThreshold(m_percentage_threshold:=65 _
                                    , m_frm:=frm) Then
            '~~ One specific frame occupies n % or more of the form's inside height.
            '~~ So this frame's height is reduced to fit the maximum height and the
            '~~ frame gets a vertical scroll-bar with the content's height
            If Not ScrollVerApplied(frm) Then
                NewHeight n_height:=frm.Height - exceeding_height _
                        , n_frame:=frm
                s_finalized = False
                bScrollApplied = True
            End If
            GoTo xt
        Else
            '~~ Neither a message section, the messsage area as a whole, nor the buttons area has
            '~~ a dominating height. Thus the message area and the buttons area's height is
            '~~ reduced proportionally and both get a vertical scroll-bar
            s_finalized = False
            siPercentageMarea = frmMarea.Height / (Me.InsideHeight / 100)
            siPercentageBarea = frmBarea.Height / (Me.InsideHeight / 100)
            If Not ScrollVerApplied(frmMarea) Then
                NewHeight n_height:=ContentHeight(frmMarea) - (exceeding_height * (siPercentageMarea / 100)) _
                        , n_frame:=frmMarea
                s_finalized = False
            End If
            If Not ScrollVerApplied(frmBarea) Then
                NewHeight n_height:=ContentHeight(frmBarea) - (exceeding_height * (siPercentageBarea / 100)) _
                        , n_frame:=frmBarea
                s_finalized = False
            End If
        End If
    End If

xt: Exit Sub

eh: If ErrMsg(ErrSrc(PROC)) = vbYes Then: Stop: Resume
End Sub

Private Sub ScrollVerRemove(ByRef sr_frame_form As Variant)
' ------------------------------------------------------------------------------
' Removes a vertical scroll-bar.
' ------------------------------------------------------------------------------
    Const PROC = "ScrollVerRemove"
    
    On Error GoTo eh
    If Not IsFrameOrForm(sr_frame_form) _
    Then Err.Raise AppErr(1), ErrSrc(PROC), "The provided argument is neither a Frame nor a Form!"
    
    With sr_frame_form
        If Not ScrollVerApplied(sr_frame_form) Then
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

Private Function ScrollVerWidth(ByVal sw_frame_form As Variant) As Single
    If IsFrameOrForm(sw_frame_form) Then
        If ScrollVerApplied(sw_frame_form) Then ScrollVerWidth = SCROLL_V_WIDTH
    End If
End Function

Public Sub Setup()
    Const PROC = "Setup"
    
    On Error GoTo eh
    '~~ Start the setup as if there wouldn't be any message - which might be the case
    Me.StartUpPosition = 2
    FormHeightInside = 400                  ' just to start with - will be expanded up to the max (default) height specified
    MsectsActivate
    IndicateCntrlCaptions False
    ' ----------------------------------------------------------------------------------------
    ' At first the Title, any monospaced message sections and the buttonsare setup as the
    ' width determining items. I.e. the initial minimum form width may be expanded up to the
    ' specified max when determined.
    ' ----------------------------------------------------------------------------------------
    Setup00WidthDeterminingItems
    IndicateCntrlCaptions False             ' may be set to True for test purpose
    
    ' -----------------------------------------------------------------------------------------------
    ' At this point the form has reached its final width (all proportionally spaced message sections
    ' are adjusted to it). However, the message height is only final in case there are just buttons
    ' but no message. The setup of proportional spaced message sections determines the final message
    ' height. When it exeeds the maximum height specified one or two vertical scrollbars are applied.
    ' -----------------------------------------------------------------------------------------------
    Setup10WidthAdaptingItems
                
    ' -----------------------------------------------------------------------------------------------
    ' When the message form height exceeds the specified or the default message height the height of
    ' the message area and or the buttons area is reduced and a vertical is applied.
    ' When both areas are about the same height (neither is taller the than 60% of the total heigth)
    ' both will get a vertical scrollbar, else only the one which uses 60% or more of the height.
    ' -----------------------------------------------------------------------------------------------
    Setup20Finalize
    
    PositionOnScreen "10;10"
    bSetUpDone = True ' To indicate for the Activate event that the setup had already be done beforehand
    BareaFrmCenter
    If bVisualizeForTest Then Stop ' for test within test environment only
    VisualizationsForTestOnly , , True ' reset visualization in case
    
xt: Exit Sub

eh: If ErrMsg(ErrSrc(PROC)) = vbYes Then: Stop: Resume
End Sub

Private Sub Setup00WidthDeterminingItems()

    If Not bSetupDoneTitle _
    Then Setup01Title
    '~~ Setup of any monospaced message sections, the second element which potentially effects the final message width.
    '~~ In case the section width exceeds the maximum width specified a horizontal scrollbar is applied.
    Setup03MonoSpacedSections
        
    '~~ Setup the reply buttons. This is the third element which may effect the final message's width.
    '~~ In case the widest buttons row exceeds the maximum width specified for the message
    '~~ a horizontal scrollbar is applied.
    If ButtonsProvided Then
        Setup05Buttons
    End If

End Sub

Private Sub Setup01Title()
' ------------------------------------------------------------------------------
' Setup the message form for the provided title (sMsgTitle) optimized with the
' provided minimum width (setup_width_min) and the provided maximum width
' (setup_width_max) by using a certain factor (setup_factor) for the calculation
' of the width required to display an untruncated title - as long as the maximum
' widht is not exeeded.
' The correction of the template length Label is a function (percentage) of the
' lenght.
' ------------------------------------------------------------------------------
    Const PROC = "Setup01Title"
    
    On Error GoTo eh
    
    '~~ The extra title Label is only used to adjust the form width and remains hidden
    With Me.laMsgTitle
        .Visible = True
        .Height = 12
        With .Font
            .Bold = False
            .Name = Me.Font.Name
        End With
        .Caption = vbNullString
        .AutoSize = True
        .Caption = " " & sMsgTitle    ' some left margin
        .Width = .Width * 0.95
        '~~ The title setup may expand the initial form's inside width. In case it does
        '~~ all subordinate frames' width is adjusted whereby the maximum widh is the limit.
        FormWidthInside = .Width ' considers a new expanded width
        .Visible = False
    End With
    Me.Caption = sMsgTitle
        
xt: Exit Sub
    
eh: If ErrMsg(ErrSrc(PROC)) = vbYes Then: Stop: Resume
End Sub

Private Sub Setup03MonoSpacedSections()
' --------------------------------------------------------------------------------------
' Setup of all sections for which a text is provided indicated mono-spaced.
' Note: The number of message sections is only determined by the number of elements in
'       MsgText.
' --------------------------------------------------------------------------------------
    Const PROC = "Setup03MonoSpacedSections"
    
    On Error GoTo eh
    Dim v   As Variant
    
    For Each v In MsectsActive
        With Msg.Section(v).Text
            If .Text <> vbNullString And .MonoSpaced = True Then
                Setup04MonoSpacedSection v
            End If
        End With
    Next v
    bSetupDoneMonoSpacedSects = True

xt: Exit Sub
    
eh: If ErrMsg(ErrSrc(PROC)) = vbYes Then: Stop: Resume
End Sub

Private Sub Setup04MonoSpacedSection(ByVal s_sect As Long)
' ------------------------------------------------------------------------------
' Setup the monospaced message section (s_sect).
' ------------------------------------------------------------------------------
Const PROC = "Setup04MonoSpacedSection"
    
    On Error GoTo eh
    Dim sText   As String
    Dim lbl     As MsForms.Label:   Set lbl = MsectTextLbl(s_sect)
    
    AutoSizeApplyTextProperties s_sect, lbl, sText, a_monospaced:=True
    
    tbxTemp.Left = 0
    AutoSizeLabelViaTextBox a_sect:=s_sect _
                          , a_lbl:=lbl _
                          , a_text:=sText _
                          , a_width_limit:=0
                          
    FormWidthInside(s_sect) = FormWidthInsideMonospaced(s_sect, lbl.Width)
    WidthsDebug s_sect, " " & PROC & " after setup monospaced section " & s_sect & " "
                  
    dctSectTextSetup.Add s_sect, vbNullString
    TempTbx lbl, t_remove:=True
    
xt: Exit Sub
    
eh: If ErrMsg(ErrSrc(PROC)) = vbYes Then: Stop: Resume
End Sub

Private Sub Setup05Buttons()
' -------------------------------------------------------------------------------
' Setup the reply buttons based on the comma delimited string of button captions
' and row breaks indicated by a vbLf, vbCr, or vbCrLf.
' ---------------------------------------------------------------------
    Const PROC = "Setup05Buttons"
    
    On Error GoTo eh
    Dim cmbBttn     As MsForms.CommandButton
    Dim frmBarea    As MsForms.Frame
    Dim frmBttnRow  As MsForms.Frame
    Dim frmBttns    As MsForms.Frame
    Dim siWidth     As Single
    Dim v           As Variant
    
    If cllMsgBttns.Count = 0 Then GoTo xt
    
    Set frmBarea = BareaFrm
    Set frmBttns = BttnsFrm
    lSetupRows = 1
    lSetupRowButtons = 0
    Set frmBttnRow = dctBttnsRowFrm(1)
    Set cmbBttn = dctBttns(1 & "-" & 1)
    
    Me.Height = 100 ' just to start with
    frmBarea.Top = AreasMarginVertical
    frmBttnRow.Top = BttnsFrm.Top
    cmbBttn.Top = frmBttnRow.Top
    cmbBttn.Width = DFLT_BTTN_MIN_WIDTH
    
    For Each v In cllMsgBttns
        If IsNumeric(v) Then v = mMsg.BttnArg(v)
        Select Case v
            Case vbOKOnly, vbOKCancel, vbYesNo, vbRetryCancel, vbYesNoCancel, vbAbortRetryIgnore, vbYesNo, vbResumeOk
                Setup06ButtonFromValue v
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
                            Setup06Button sb_row:=lSetupRows, sb_button:=lSetupRowButtons, sb_caption:=v, sb_ret_value:=v
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
    Setup05ButtonsSizeAndPosition1Buttons
    Setup05ButtonsSizeAndPosition2ButtonRows
    Setup05ButtonsSizeAndPosition3ButtonsFrame
    
    '~~ The following may expand the BareaFrm's width up to the default or specified
    '~~ maximum. In case it does all related controls' width are adjusted correspondingly
    BareaFrmWidth = ContentWidth(frmBttns) + AreasFrmMarginLeft + AreasFrmMarginRight
    With frmBarea
        .Height = ContentHeight(BareaFrm) + AreasFrmMarginTop + AreasFrmMarginBottom
    End With
    FormHeightAdjust ' height
    bButtonsSetup = True
    
xt: Exit Sub

eh: If ErrMsg(ErrSrc(PROC)) = vbYes Then: Stop: Resume
End Sub

Private Sub Setup05ButtonsSizeAndPosition1Buttons()
' ------------------------------------------------------------------------------
' Unify all applied/visible button's size by assigning the maximum width and
' height provided with their setup, and adjust their resulting left position.
' ------------------------------------------------------------------------------
    Const PROC = "Setup05ButtonsSizeAndPosition1Buttons"
    
    On Error GoTo eh
    Dim siLeft          As Single
    Dim frmRow          As MsForms.Frame    ' Frame for the buttons in a row
    Dim v               As Variant
    Dim lRow            As Long
    Dim lButton         As Long
    Dim siHeightBarea   As Single
    Dim cmb             As MsForms.CommandButton
    
    For lRow = 1 To dctBttnsRowFrm.Count
        siLeft = BttnRowMarginLeft
        Set frmRow = dctBttnsRowFrm(lRow)
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
                            .Top = BttnsRowVerticalMargin
                            siLeft = .Left + .Width + H_MARGIN_BTTNS
                            If IsNumeric(MsgButtonDefault) Then
                                If lButton = MsgButtonDefault Then .Default = True
                            Else
                                If .Caption = MsgButtonDefault Then .Default = True
                            End If
                        End With
                    End If
                End If
                siHeightBarea = siHeightBarea + siMaxButtonHeight + H_MARGIN_BTTNS
            Next v
        End If
    Next lRow
    FormHeightAdjust
        
xt: Exit Sub
    
eh: If ErrMsg(ErrSrc(PROC)) = vbYes Then: Stop: Resume
End Sub

Private Sub Setup05ButtonsSizeAndPosition2ButtonRows()
' ------------------------------------------------------------------------------
' Adjust all applied/visible button rows height to the maximum buttons height
' and the row frames width to the number of the displayed buttons considering a
' certain margin between the buttons (H_MARGIN_BTTNS) and a margin at the
' left and the right.
' ------------------------------------------------------------------------------
    Const PROC = "Setup05ButtonsSizeAndPosition2ButtonRows"
    
    On Error GoTo eh
    Dim frmRow          As MsForms.Frame
    Dim siTop           As Single
    Dim v               As Variant
    Dim lButtons        As Long
    Dim siHeight        As Single
    Dim WidthBttnsFrm   As Single
    Dim dct             As Dictionary:      Set dct = AppliedBttnRows
    
    '~~ Adjust button row's width and height
    siHeight = AppliedButtonRowHeight
    siTop = BttnsRowVerticalMargin
    For Each v In dct
        Set frmRow = v
        lButtons = dct(v)
        If frmRow.Visible Then
            With frmRow
                .Top = siTop
                .Height = siHeight
                '~~ Provide some extra space for the button's design
                .Width = ContentWidth(frmRow) + BttnRowMarginLeft + BttnRowMarginRight
                siTop = .Top + .Height + VSPACE_BTTN_ROWS
            End With
        End If
    Next v
    Set dct = Nothing

xt: Exit Sub
    
eh: If ErrMsg(ErrSrc(PROC)) = vbYes Then: Stop: Resume
End Sub

Private Sub Setup05ButtonsSizeAndPosition3ButtonsFrame()
' ------------------------------------------------------------------------------
' Adjust the frame around all button row frames to the maximum width calculated
' by the adjustment of each of the rows frame.
' ------------------------------------------------------------------------------
    Const PROC = "Setup05ButtonsSizeAndPosition3ButtonsFrame"
    
    On Error GoTo eh
    Dim v           As Variant
    Dim siWidth     As Single
    Dim siHeight    As Single
    Dim frm         As MsForms.Frame
    
    If frmBttnsFrm.Visible Then
        siWidth = ContentWidth(frmBttnsFrm)
        siHeight = ContentHeight(frmBttnsFrm)
        With frmBttnsFrm
            .Top = 0
            BttnsFrm.Height = siHeight
            BttnsFrm.Width = siWidth
            '~~ Center all button rows within the buttons frame
            For Each v In dctBttnsRowFrm
                Set frm = dctBttnsRowFrm(v)
                If frm.Visible Then
                    FrameCenterHorizontal f_center:=frm, f_within:=frmBttnsFrm
                End If
            Next v
        End With
    End If

xt: Exit Sub
    
eh: If ErrMsg(ErrSrc(PROC)) = vbYes Then: Stop: Resume
End Sub

Private Sub Setup06Button(ByVal sb_row As Long, _
                          ByVal sb_button As Long, _
                          ByVal sb_caption As String, _
                          ByVal sb_ret_value As Variant)
' -------------------------------------------------------------------------------
' Setup an applied reply button's (sb_row, sb_button) visibility and caption,
' calculate the maximum width and height, keep a record of the setup
' reply sb_index's return value.
' -------------------------------------------------------------------------------
    Const PROC = "Setup06Button"
    
    On Error GoTo eh
    Dim cmb As MsForms.CommandButton
    
    If sb_row = 0 Then sb_row = 1
    Set cmb = dctBttns(sb_row & "-" & sb_button)
    
    With cmb
        .AutoSize = True
        .WordWrap = False ' the longest line determines the sb_index's width
        .Caption = Replace(sb_caption, "\,", ",") ' an escaped , is considered
        .AutoSize = False
        .Height = .Height + 1 ' safety margin to ensure proper multilin caption display
        .BackColor = &H8000000F
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

Private Sub Setup06ButtonFromValue(ByVal lButtons As Long)
' -------------------------------------------------------------------------------
' Setup a row of standard VB MsgBox reply command buttons
' -------------------------------------------------------------------------------
    Const PROC = "Setup06ButtonFromValue"
    
    On Error GoTo eh
    Dim ResumeErrorLine As String: ResumeErrorLine = "Resume" & vbLf & "Error Line"
    Dim PassOn          As String: PassOn = "Pass on Error to" & vbLf & "Entry Procedure"
    
    Select Case lButtons
        Case vbOKOnly
            lSetupRowButtons = lSetupRowButtons + 1
            Setup06Button sb_row:=lSetupRows, sb_button:=lSetupRowButtons, sb_caption:="Ok", sb_ret_value:=vbOK
        Case vbOKCancel
            lSetupRowButtons = lSetupRowButtons + 1
            Setup06Button sb_row:=lSetupRows, sb_button:=lSetupRowButtons, sb_caption:="Ok", sb_ret_value:=vbOK
            lSetupRowButtons = lSetupRowButtons + 1
            Setup06Button sb_row:=lSetupRows, sb_button:=lSetupRowButtons, sb_caption:="Cancel", sb_ret_value:=vbCancel
        Case vbYesNo
            lSetupRowButtons = lSetupRowButtons + 1
            Setup06Button sb_row:=lSetupRows, sb_button:=lSetupRowButtons, sb_caption:="Yes", sb_ret_value:=vbYes
            lSetupRowButtons = lSetupRowButtons + 1
            Setup06Button sb_row:=lSetupRows, sb_button:=lSetupRowButtons, sb_caption:="No", sb_ret_value:=vbNo
        Case vbRetryCancel
            lSetupRowButtons = lSetupRowButtons + 1
            Setup06Button sb_row:=lSetupRows, sb_button:=lSetupRowButtons, sb_caption:="Retry", sb_ret_value:=vbRetry
            lSetupRowButtons = lSetupRowButtons + 1
            Setup06Button sb_row:=lSetupRows, sb_button:=lSetupRowButtons, sb_caption:="Cancel", sb_ret_value:=vbCancel
        Case vbResumeOk
            lSetupRowButtons = lSetupRowButtons + 1
            Setup06Button sb_row:=lSetupRows, sb_button:=lSetupRowButtons, sb_caption:=ResumeErrorLine, sb_ret_value:=vbResume
            lSetupRowButtons = lSetupRowButtons + 1
            Setup06Button sb_row:=lSetupRows, sb_button:=lSetupRowButtons, sb_caption:="Ok", sb_ret_value:=vbOK
        Case vbYesNoCancel
            lSetupRowButtons = lSetupRowButtons + 1
            Setup06Button sb_row:=lSetupRows, sb_button:=lSetupRowButtons, sb_caption:="Yes", sb_ret_value:=vbYes
            lSetupRowButtons = lSetupRowButtons + 1
            Setup06Button sb_row:=lSetupRows, sb_button:=lSetupRowButtons, sb_caption:="No", sb_ret_value:=vbNo
            lSetupRowButtons = lSetupRowButtons + 1
            Setup06Button sb_row:=lSetupRows, sb_button:=lSetupRowButtons, sb_caption:="Cancel", sb_ret_value:=vbCancel
        Case vbAbortRetryIgnore
            lSetupRowButtons = lSetupRowButtons + 1
            Setup06Button sb_row:=lSetupRows, sb_button:=lSetupRowButtons, sb_caption:="Abort", sb_ret_value:=vbAbort
            lSetupRowButtons = lSetupRowButtons + 1
            Setup06Button sb_row:=lSetupRows, sb_button:=lSetupRowButtons, sb_caption:="Retry", sb_ret_value:=vbRetry
            lSetupRowButtons = lSetupRowButtons + 1
            Setup06Button sb_row:=lSetupRows, sb_button:=lSetupRowButtons, sb_caption:="Ignore", sb_ret_value:=vbIgnore
        Case vbResumeOk
            lSetupRowButtons = lSetupRowButtons + 1
            Setup06Button sb_row:=lSetupRows, sb_button:=lSetupRowButtons, sb_caption:="Resume" & vbLf & "Error Line", sb_ret_value:=vbResume
            lSetupRowButtons = lSetupRowButtons + 1
            Setup06Button sb_row:=lSetupRows, sb_button:=lSetupRowButtons, sb_caption:="Ok", sb_ret_value:=vbOK
    
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

Private Sub Setup10WidthAdaptingItems()
    Setup13Labels
    Setup11PropSpacedSections
End Sub

Private Sub Setup11PropSpacedSections()
' -------------------------------------------------------------------------------
' Loop through all provided message sections for which a text is provided and is
' not Monospaced and setup the section.
' -------------------------------------------------------------------------------
    Const PROC = "Setup11PropSpacedSections"
    
    On Error GoTo eh
    Dim v   As Variant
    Dim frm As MsForms.Frame
    
    For Each v In MsectsActive
        With Msg.Section(v).Text
            If .Text <> vbNullString And .MonoSpaced = False Then
                Setup12PropSpacedSection v
            End If
        End With
    Next v
    bSetupDonePropSpacedSects = True
    
xt: Exit Sub
    
eh: If ErrMsg(ErrSrc(PROC)) = vbYes Then: Stop: Resume
End Sub

Private Sub Setup12PropSpacedSection(ByVal s_sect As Long)
' ------------------------------------------------------------------------------
' Sets up the proportional-spaced section (s_sect).
' ------------------------------------------------------------------------------
    Const PROC = "Setup12PropSpacedSection"
    
    On Error GoTo eh
    Dim lblText As MsForms.Label:   Set lblText = MsectTextLbl(s_sect)
    Dim sText   As String:          sText = Msg.Section(s_sect).Text.Text
    Dim frmText As MsForms.Frame:   Set frmText = MsectTextFrm(s_sect)
    
    AutoSizeApplyTextProperties s_sect, lblText, sText, a_monospaced:=False
    tbxTemp.Left = 0
    AutoSizeLabelViaTextBox a_sect:=s_sect _
                          , a_lbl:=lblText _
                          , a_text:=sText _
                          , a_width_limit:=MsectTextWidthLimit(s_sect)
                
    WidthsDebug s_sect, " setup prop-speced section done "
    dctSectTextSetup.Add s_sect, vbNullString
    TempTbx lblText, t_remove:=True
    
xt: Exit Sub
    
eh: If ErrMsg(ErrSrc(PROC)) = vbYes Then: Stop: Resume
End Sub

Private Sub Setup13Labels()
' -------------------------------------------------------------------------------
' Loops through all provided message sections for which a text is provided and is
' not Monospaced and setup the section.
' Note 1: The number of message sections is only determined by the number of
'         elements in MsgText.
' Note 2: The setup of Labels must be preceeded by all width determining setup!
' -------------------------------------------------------------------------------
    Const PROC = "Setup13Labels"
    
    On Error GoTo eh
    Dim v   As Variant
    
    With Msg
        For Each v In MsectsActive
            If .Section(v).Label.Text <> vbNullString Then
                Setup14Label s_sect:=v
            End If
        Next v
    End With
    
xt: Exit Sub
    
eh: If ErrMsg(ErrSrc(PROC)) = vbYes Then: Stop: Resume
End Sub

Private Sub Setup14Label(ByVal s_sect As Long)
' ------------------------------------------------------------------------------
'
' ------------------------------------------------------------------------------
    Const PROC = "Setup14Label"
    
    Dim lbl     As MsForms.Label
    Dim sText   As String
    
    If MsgLabel(s_sect).Text <> vbNullString Then
        Set lbl = MsectLabel(s_sect)
        AutoSizeApplyLabelProperties a_sect:=s_sect _
                                   , a_lbl:=lbl _
                                   , a_text:=sText
        tbxTemp.Left = MsectFrmMarginLeft
        AutoSizeLabelViaTextBox s_sect, lbl, sText, MsectFrmWidthLabelLimit(s_sect)
        dctSectLabelSetup.Add s_sect, vbNullString
        With lbl
            .Left = MsectFrmMarginLeft
            If MsectHasOnlyLabel(s_sect) Then
                .Width = siMsectFrmWidth + MsectFrmMarginLeft
            Else
                .Width = siMsectLabelWidthAll
            End If
            .TextAlign = MsectLabelAlign(s_sect)
            .Font.Size = MsectLabelFontSize(s_sect)
        End With
        DoEvents
    End If
    WidthsDebug s_sect, " " & PROC & " done for section " & s_sect & " "
    
End Sub

Private Sub Setup20Finalize()
' ------------------------------------------------------------------------------
' Finalizes all frame control's height and top positions and their widths.
' ------------------------------------------------------------------------------
    Const PROC = "Setup20Finalize"
    
    On Error GoTo eh
    Dim bDoneHeights    As Boolean
    Dim bDoneWidths     As Boolean
    
    FinalizeWidths bDoneWidths
    FinalizeHeights bDoneHeights
    FinalizeWidths bDoneWidths
        
xt: Exit Sub
    
eh: If ErrMsg(ErrSrc(PROC)) = vbYes Then: Stop: Resume
End Sub

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
       .StartUpPosition = 0
       .Left = PosLeft
       .Top = PosTop
    End With

End Sub

Public Function StringLengthPt(ByVal s_string As String, _
                      Optional ByVal s_Font_name As String = vbNullString, _
                      Optional ByVal s_Font_size As Single = 0) As Single
    
    Dim tbx As MsForms.Label
    
    If s_Font_name = vbNullString Then s_Font_name = Me.Font.Name
    If s_Font_size = 0 Then s_Font_size = Me.Font.Size
    
    On Error Resume Next
    Me.Controls.Remove "tbxTemp"
    Set tbx = Me.Controls.Add("Forms.TextBox.1", "tbxTemp", False)
    With tbx
        .Width = 8
        .Font.Name = s_Font_name
        .Font.Size = s_Font_size
    End With
    AutoSizeTextBox tbx, s_string, 0
    StringLengthPt = tbx.Width ' / Application.PointsToScreenPixelsX(1)
    Me.Controls.Remove "tbxTemp"
    
End Function

Private Sub TempTbx(ByVal t_lbl As MsForms.Label, _
           Optional ByVal t_remove As Boolean = False)
' ------------------------------------------------------------------------------
' Removes from - or provides in - a frame control (t_frm) a temporary TextBox
' control (global tbxTemp).
' ------------------------------------------------------------------------------
    If t_remove Then
        On Error Resume Next
        t_lbl.Parent.Controls.Remove TEMP_TBX_NAME
        Exit Sub
    Else
        Set tbxTemp = t_lbl.Parent.Controls.Add("Forms.TextBox.1", TEMP_TBX_NAME, False)
    End If
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
    TimedDoEvents = s
    
End Function

Private Sub TimerBegin()
    cyTimerTicksBegin = TimerSysCurrentTicks
End Sub

Private Function TimerEnd() As Currency
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

Private Sub VisualizationsForTestOnly(Optional ByVal v_ctl As MsForms.Control, _
                                      Optional ByVal v_backcolor As Long, _
                                      Optional ByVal v_reset As Boolean = False)
' ------------------------------------------------------------------------------
' Visualizes the Control (v_ctl) with the BackColor (v_backcolor) when
' bVisualizeForTest  is TRUE.
' ------------------------------------------------------------------------------
    Const PROC = "VisualizationsForTestOnly"
    
    On Error GoTo eh
    Dim v As Variant
    
    If v_reset Then
        For Each v In Me.Controls
            With v
                If TypeOf v Is Frame Then
                    On Error Resume Next
                    .Caption = vbNullString
                End If
                
                If Not TypeOf v Is CommandButton Then
                    If .BackColor <> Me.BackColor Then
                        .BackColor = Me.BackColor
                        On Error Resume Next
                        .BorderColor = Me.BackColor
                        On Error Resume Next
                        BorderStyle = fmBorderStyleSingle
                        DoEvents
                    End If
                End If
           End With
        Next v
        Exit Sub
    End If
    
    If bVisualizeForTest Then
        With v_ctl
            If TypeOf v_ctl Is Frame Then
                On Error Resume Next
                .Caption = vbNullString
            End If
            .BackColor = v_backcolor
            If Not dctVisualizedForTest.Exists(v_ctl) Then
                dctVisualizedForTest.Add v_ctl, v_ctl.Name ' keep a record of the visualized control
            End If
        End With
    End If
    
xt: Exit Sub
    
eh: If ErrMsg(ErrSrc(PROC)) = vbYes Then: Stop: Resume
End Sub

Function WidthExpanded(ByVal w_this As Single, _
                       ByVal w_min As Single, _
                       ByVal w_max As Single, _
                       ByVal w_now As Single, _
                       ByRef w_new As Single) As Boolean
' ------------------------------------------------------------------------------
' The determination of a message form's width (caused by the Title, the Buttons,
' and mono-spaced sections setup) is a key setup issue.
' Initial widths:
' ---------------
' With the spec of the max outside form width the following max widths are
' determined: - siFormWidthInsideMax
'             - siAreasFrmWidthMax
'             - siWidthBttnsFrmMax
'             - siMsectFrmWidthMax
'             - siMsectTextFrmWidthOnlyTextMax
'             - siMsectTextFrmWidthWithLposLblMax
'
' With the spec of the mon outside form width the following min widths are
' determined: - siFormWidthInsideMin
'             - siAreasFrmWidthMin
'             - siWidthBttnsFrmMin
'             - siMsectFrmWithMin
'             - siMsectTextFrmWidthOnlyTextMin
'             - siMsectTextWithLabelLeftFrmWithMin
' which also become the value for the current effective width
'             - siFormWidthInside
'             - siAreasFrmWidth
'             - siWidthBttnsFrm
'             - siMsectFrmWith
'             - siMsectTextFrmWidthOnlyText
'             - siMsectTextWithLabelLeftFrmWith
'
' Effective widths:
' -----------------
' The effective widths are used as the width limit for the setup of all
' proportional spaced sections and Top positioned Labels - resulting in
' autosizing the height. The following setup triggers an effective width which
' subsequently triggers the effective widths of all subordinate frames.
' - Title:              FormInsideWidth
' - Buttons:            WidthBttnsFrm
' - Monospaced section: MsectTextFrmWidthOnlyText or
'                       MsectTextWithLabelLeftFrmWith
'
' Horizontal Scroll-Bars:
' -----------------------
' - When the content width of any monospaced section text frame exceeds the
'   width of the text frame, the text frame is provided with a horizontal
'   scroll-bar.
' - When the content width of the buttons frame exceeds the width of the parent
'   frame, the buttons frame  is provided with a horizontal scroll-bar.
' ------------------------------------------------------------------------------

    w_new = Min(w_this, w_max)  ' limits to max
    w_new = Max(w_new, w_min) ' possible expands
    If w_new > w_now Then
        WidthExpanded = True
        Debug.Print "Width expanded from " & w_now & " to " & w_new
        If w_new = w_max Then
            Debug.Print " ... and has reached its max of " & w_max
        End If
    End If
End Function

Public Sub WidthsDebug(Optional ByVal m_sect As Long = 0, _
                  Optional ByVal m_phase As String = vbNullString)
' ------------------------------------------------------------------------------
' For test and debugging only!
' ------------------------------------------------------------------------------
    
    Debug.Print "====== Widths after/at: " & m_phase & "==================================================="
    Debug.Print "FormWidthOutside: Max = " & FormWidthOutsideMax & " (dpi=" & DsplyWidthDPI & ", pt=" & DsplyWidthPT & ", " & FormWidthOutsideMax / (DsplyWidthPT / 100) & "% )"
    Debug.Print "----------------- Min = " & FormWidthOutsideMin & " (dpi=" & DsplyWidthDPI & ", pt=" & DsplyWidthPT & ", " & FormWidthOutsideMin / (DsplyWidthPT / 100) & "% )"
    Debug.Print "                  now = " & Me.Width
    Debug.Print "FormWidthInside : Max = " & siFormWidthInsideMax
    Debug.Print "----------------- Min = " & siFormWidthInsideMin
    Debug.Print "                  now = " & siFormWidthInside
    Debug.Print "AreasFrmWidth   : Max = " & AreasFrmWidthMax
    Debug.Print "-- Barea/Marea -- Min = " & AreasFrmWidthMin
    Debug.Print "                  now = " & siAreasFrmWidth
    Debug.Print "BttnsFrm        : Max = " & AreasFrmWidthMax
    Debug.Print "----------------- Min = " & AreasFrmWidthMin
    Debug.Print "                  now = " & BttnsFrm.Width

End Sub

