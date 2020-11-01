VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} fMsg 
   ClientHeight    =   9255.001
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
' UserForm fMsg
' Provides all means for a message with up to 3 separated text sections, either
' proportional- or mono-spaced, with an optional label, and up to 7 reply buttons.
'
' Design:
' Since the implementation is merely design driven its setup is is essential.
' Design changes must comply with the concept. For details please refer to:
'
' Implementation:
'
' Properties:
'
'
' lScreenWidth. Rauschenberger Berlin March 2020
' --------------------------------------------------------------------------
Const MIN_BUTTON_WIDTH          As Single = 70      ' Default minimum reply button width
Const FONT_MONOSPACED_NAME      As String = "Courier New"   ' Default monospaced font
Const FONT_MONOSPACED_SIZE      As Single = 9       ' Default monospaced font size
Const FORM_MAX_HEIGHT_POW       As Long = 90        ' Max form height as a percentage of the screen size
Const FORM_MAX_WIDTH_POW        As Long = 80        ' Max form width as a percentage of the screen size
Const FORM_MIN_WIDTH            As Single = 300     ' Default minimum message form width
Const TEST_WITH_FRAME_BORDERS   As Boolean = False  ' For test purpose only! Display frames with visible border
Const TEST_WITH_FRAME_CAPTIONS  As Boolean = False  ' For test purpose only! Display frames with their test captions (erased by default)
Const HSPACE_BUTTONS            As Single = 4       ' Horizontal space between reply buttons
Const HSPACE_BUTTON_AREA        As Single = 10      ' Minimum margin between buttons area and form when centered
Const HSPACE_LEFT               As Single = 0       ' Left margin for labels and text boxes
Const HSPACE_RIGHT              As Single = 15      ' Horizontal right space for labels and text boxes
Const HSPACE_SCROLLBAR          As Single = 18      ' Additional horizontal space required for a frame with a vertical scroll bar
Const NEXT_ROW                  As String = vbLf    ' Reply button row break
Const VSPACE_AREAS              As Single = 10      ' Vertical space between message area and replies area
Const VSPACE_BOTTOM             As Single = 50      ' Bottom space after the last displayed reply row
Const VSPACE_BUTTON_ROWS        As Single = 5       ' Vertical space between button rows
Const VSPACE_LABEL              As Single = 0       ' Vertical space between label and the following text
Const VSPACE_SCROLLBAR          As Single = 12      ' Additional vertical space required for a frame with a horizontal scroll barr
Const VSPACE_SECTIONS           As Single = 5       ' Vertical space between displayed message sections
Const VSPACE_TEXTBOXES          As Single = 18      ' Vertical bottom marging for all textboxes
Const VSPACE_TOP                As Single = 2       ' Top position for the first displayed control

' Functions to get the displays DPI
' Used for getting the metrics of the system devices.
'
Const SM_XVIRTUALSCREEN As Long = &H4C&
Const SM_YVIRTUALSCREEN As Long = &H4D&
Const SM_CXVIRTUALSCREEN As Long = &H4E&
Const SM_CYVIRTUALSCREEN As Long = &H4F&
Const LOGPIXELSX = 88
Const LOGPIXELSY = 90
Const TWIPSPERINCH = 1440
Private Declare PtrSafe Function GetSystemMetrics32 Lib "user32" Alias "GetSystemMetrics" (ByVal nIndex As Long) As Long
Private Declare PtrSafe Function GetDC Lib "user32" (ByVal hWnd As Long) As Long
Private Declare PtrSafe Function GetDeviceCaps Lib "gdi32" (ByVal hDC As Long, ByVal nIndex As Long) As Long
Private Declare PtrSafe Function ReleaseDC Lib "user32" (ByVal hWnd As Long, ByVal hDC As Long) As Long

Dim siMinButtonWidth            As Single
Dim bDisplayFramesWithCaptions  As Boolean
Dim bDoneButtonsArea            As Boolean
Dim bDoneHeightDecrement        As Boolean
Dim bHscrollbarButtonsArea      As Boolean
Dim bVscrollbarButtonsArea      As Boolean
Dim bVscrollbarMsgArea          As Boolean
Dim bDoneMonoSpacedSections     As Boolean
Dim bDoneMsgArea                As Boolean
Dim bDonePropSpacedSections     As Boolean
Dim bDoneSetup                  As Boolean
Dim bDoneTitle                  As Boolean
Dim bFormEvents                 As Boolean
Dim bTestFrameWithBorders       As Boolean
Dim bWithFrames                 As Boolean          ' for test purpose only, defaults to False
Dim cllDsgnAreas                As Collection   ' Collection of the two primary/top frames
Dim cllDsgnButtonRows           As Collection   ' Collection of the designed reply button row frames
Dim cllDsgnButtons              As Collection   ' Collection of the collection of the designed reply buttons of a certain row
Dim cllDsgnButtonsFrame         As Collection
Dim cllDsgnRowButtons           As Collection       ' Collection of a designed reply button row's buttons
Dim cllDsgnSections             As Collection   '
Dim cllDsgnSectionsLabel        As Collection
Dim cllDsgnSectionsText         As Collection   ' Collection of section frames
Dim cllDsgnSectionsTextFrame    As Collection
Dim dctApplButtonRows           As Dictionary   ' Dictionary of applied/used/visible button rows (key=frame, item=row)
Dim dctApplButtons              As Dictionary   ' Dictionary of applied buttons (key=CommandButton, item=row)
Dim dctApplButtonsRetVal        As Dictionary   ' Dictionary of the applied buttons' reply value (key=CommandButton)
Dim dctApplied                  As Dictionary   ' Dictionary of all applied controls (versus just designed)
Dim dctSectionsLabel            As Dictionary   ' Section specific label either provided via properties MsgLabel or Msg
Dim dctSectionsMonoSpaced       As Dictionary   ' Section specific monospace option either provided via properties MsgMonospaced or Msg
Dim dctSectionsText             As Dictionary   ' Section specific text either provided via properties MsgText or Msg
Dim lMaxFormHeightPoW           As Long             ' % of the screen height
Dim lMaxFormWidthPoW            As Long             ' % of the screen width
Dim lMinFormWidthPoW            As Long             ' % of the screen width - calculated when min form width in pt is assigend
Dim lReplyIndex                 As Long             ' Index of the clicked reply button (a value ranging from 1 to 49)
Dim sErrSrc                     As String
Dim siHmarginButtons            As Single
Dim siHmarginFrames             As Single           ' Test property, value defaults to 0
Dim siMaxButtonHeight           As Single
Dim siMaxButtonRowWidth         As Single
Dim siMaxButtonWidth            As Single
Dim siMaxFormHeight             As Single           ' above converted to excel userform height
Dim siMaxFormWidth              As Single           ' above converted to excel userform width
Dim siMaxSectionWidth           As Single
Dim siMinFormWidth              As Single
Dim siMonoSpacedFontSize        As Single
Dim siVmarginButtons            As Single
Dim siVmarginFrames             As Single           ' Test property, value defaults to 0
Dim sMonoSpacedFontName         As String
Dim sTitle                      As String
Dim sTitleFontName              As String
Dim sTitleFontSize              As String           ' Ignored when sTitleFontName is not provided
Dim vButtons                    As Variant
Dim vReplyValue                 As Variant
Dim wVirtualScreenHeight        As Single
Dim wVirtualScreenLeft          As Single
Dim wVirtualScreenTop           As Single
Dim wVirtualScreenWidth         As Single
Dim lSetupRowButtons            As Long ' number of buttons setup in a row
Dim lSetupRows                  As Long ' number of setup button rows

Private Sub UserForm_Initialize()
        
    On Error GoTo on_error
    siMinButtonWidth = MIN_BUTTON_WIDTH
    siHmarginButtons = HSPACE_BUTTONS
    siVmarginButtons = VSPACE_BUTTON_ROWS
    bFormEvents = False
    GetScreenMetrics                                            ' This environment screen's width and height
    Me.MinFormWidth = FORM_MIN_WIDTH                            ' Default UserForm width (als calculated as % of screen size)
    Me.MaxFormWidthPrcntgOfScreenSize = FORM_MAX_WIDTH_POW
    Me.MaxFormHeightPrcntgOfScreenSize = FORM_MAX_HEIGHT_POW
    sMonoSpacedFontName = FONT_MONOSPACED_NAME                  ' Default monospaced font
    siMonoSpacedFontSize = FONT_MONOSPACED_SIZE                 ' Default monospaced font
    Me.width = siMinFormWidth
    bDisplayFramesWithCaptions = False
    bTestFrameWithBorders = False
    Me.Height = VSPACE_AREAS * 4
    bWithFrames = False
    siHmarginFrames = 0     ' Ensures proper command buttons framing, may be used for test purpose
    Me.VmarginFrames = 0    ' Ensures proper command buttons framing and vertical positioning of controls
    bDoneSetup = False
    bDoneTitle = False
    bDoneButtonsArea = False
    bDoneMonoSpacedSections = False
    bDonePropSpacedSections = False
    bDoneMsgArea = False
    bDoneHeightDecrement = False
    
exit_sub:
    Exit Sub
    
on_error:
    Debug.Print err.Description: Stop: Resume
End Sub

Private Sub CollectDesignControls()
' ----------------------------------------------------------------------
' Collects all designed controls without concidering any control's name.
' ----------------------------------------------------------------------
    On Error GoTo on_error
    
    ProvideCollection cllDsgnAreas
    Collect into:=cllDsgnAreas, ctltype:="Frame", fromparent:=Me, ctlheight:=10, ctlwidth:=Me.width - siHmarginFrames
    DsgnButtonsArea.width = 10  ' Will be adjusted to the max replies row width during setup
    
    ProvideCollection cllDsgnSections
    Collect into:=cllDsgnSections, ctltype:="Frame", fromparent:=DsgnMsgArea, ctlheight:=50, ctlwidth:=DsgnMsgArea.width - siHmarginFrames
    ProvideCollection cllDsgnSectionsLabel
    Collect into:=cllDsgnSectionsLabel, ctltype:="Label", fromparent:=cllDsgnSections, ctlheight:=15, ctlwidth:=DsgnMsgArea.width - (siHmarginFrames * 2)
    ProvideCollection cllDsgnSectionsTextFrame
    Collect into:=cllDsgnSectionsTextFrame, ctltype:="Frame", fromparent:=cllDsgnSections, ctlheight:=20, ctlwidth:=DsgnMsgArea.width - (siHmarginFrames * 2)
    ProvideCollection cllDsgnSectionsText
    Collect into:=cllDsgnSectionsText, ctltype:="TextBox", fromparent:=cllDsgnSectionsTextFrame, ctlheight:=20, ctlwidth:=DsgnMsgArea.width - (siHmarginFrames * 3)
    ProvideCollection cllDsgnButtonsFrame
    Collect into:=cllDsgnButtonsFrame, ctltype:="Frame", fromparent:=DsgnButtonsArea, ctlheight:=10, ctlwidth:=10
    ProvideCollection cllDsgnButtonRows
    Collect into:=cllDsgnButtonRows, ctltype:="Frame", fromparent:=cllDsgnButtonsFrame, ctlheight:=10, ctlwidth:=10
        
    Dim v As Variant
    ProvideCollection cllDsgnButtons
    For Each v In cllDsgnButtonRows
        ProvideCollection cllDsgnRowButtons
        Collect into:=cllDsgnRowButtons, ctltype:="CommandButton", fromparent:=v, ctlheight:=10, ctlwidth:=siMinButtonWidth
        If cllDsgnRowButtons.Count > 0 Then
            cllDsgnButtons.Add cllDsgnRowButtons
        End If
    Next v
    ProvideDictionary dctApplied ' provides a clean or new dictionary for collection applied controls
    ProvideDictionary dctApplButtons
    ProvideDictionary dctApplButtonsRetVal
    ProvideDictionary dctApplButtonRows

exit_proc:
    Exit Sub
    
on_error:
    Debug.Print err.Description: Stop: Resume
End Sub

Private Sub ProvideCollection(ByRef cll As Collection)
' ----------------------------------------------------
' Provides a clean/new Collection.
' ----------------------------------------------------
    If Not cll Is Nothing Then Set cll = Nothing
    Set cll = New Collection
End Sub

Private Sub ProvideDictionary(ByRef dct As Dictionary)
' ----------------------------------------------------
' Provides a clean or new Dictionary.
' ----------------------------------------------------
    If Not dct Is Nothing Then dct.RemoveAll Else Set dct = New Dictionary
End Sub

Private Sub UserForm_Terminate()
    Set cllDsgnAreas = Nothing
    Set cllDsgnButtonRows = Nothing
    Set cllDsgnButtons = Nothing
    Set cllDsgnButtonsFrame = Nothing
    Set cllDsgnRowButtons = Nothing
    Set cllDsgnSections = Nothing
    Set cllDsgnSectionsLabel = Nothing
    Set cllDsgnSectionsText = Nothing
    Set cllDsgnSectionsTextFrame = Nothing
    Set dctApplButtonRows = Nothing
    Set dctApplButtons = Nothing
    Set dctApplButtonsRetVal = Nothing
    Set dctApplied = Nothing
    Set dctSectionsLabel = Nothing
    Set dctSectionsMonoSpaced = Nothing
    Set dctSectionsText = Nothing
End Sub

Private Property Get ClickedButtonIndex(Optional ByVal cmb As MSForms.CommandButton) As Long
    
    Dim i   As Long
    Dim v   As Variant
    
    For Each v In dctApplButtonsRetVal
        i = i + 1
        If v Is cmb Then
            ClickedButtonIndex = i
            Exit For
        End If
    Next v

End Property

Private Property Get AppliedButtonRetVal(Optional ByVal Button As MSForms.CommandButton) As Variant
    AppliedButtonRetVal = dctApplButtonsRetVal(Button)
End Property

Private Property Let AppliedButtonRetVal(Optional ByVal Button As MSForms.CommandButton, ByVal v As Variant)
    dctApplButtonsRetVal.Add Button, v
End Property

Private Property Get AppliedButtonRowHeight() As Single:                                        AppliedButtonRowHeight = siMaxButtonHeight + (siVmarginFrames * 2) + 2:       End Property
Private Property Get AppliedButtonRowWidth(Optional ByVal buttons As Long) As Single
    '~~ Extra space required for the button's design
    AppliedButtonRowWidth = CInt((siMaxButtonWidth * buttons) + (siHmarginButtons * (buttons - 1)) + (siHmarginFrames * 2)) + 5
End Property

Private Property Let AppliedControl(ByVal v As Variant)
    If dctApplied Is Nothing Then Set dctApplied = New Dictionary
    If Not IsApplied(v) Then dctApplied.Add v, v.Name
End Property

Private Property Get ButtonsFrameHeight() As Single
    Dim l As Long:  l = dctApplButtonRows.Count
    ButtonsFrameHeight = (AppliedButtonRowHeight * l) + (siVmarginButtons * (l - 1)) + (siVmarginFrames * 2) + 2
End Property

Private Property Get ButtonsFrameWidth() As Single:                                     ButtonsFrameWidth = siMaxButtonRowWidth + (siHmarginFrames * 2):            End Property

Private Property Get DsgnButton(Optional ByVal row As Long, Optional ByVal Button As Long) As MSForms.CommandButton
    Set DsgnButton = cllDsgnButtons(row)(Button)
End Property

Private Property Get DsgnButtonRow(Optional ByVal row As Long) As MSForms.Frame:        Set DsgnButtonRow = cllDsgnButtonRows(row):                                 End Property

Private Property Get DsgnButtonRows() As Collection:                                    Set DsgnButtonRows = cllDsgnButtonRows:                                     End Property

Private Property Get DsgnButtonsArea() As MSForms.Frame:                                Set DsgnButtonsArea = cllDsgnAreas(2):                                      End Property

Private Property Get DsgnButtonsFrame() As MSForms.Frame:                               Set DsgnButtonsFrame = cllDsgnButtonsFrame(1):                              End Property

Private Property Get DsgnMsgArea() As MSForms.Frame:                                    Set DsgnMsgArea = cllDsgnAreas(1):                                          End Property

Private Property Get DsgnSection(Optional section As Long) As MSForms.Frame:            Set DsgnSection = cllDsgnSections(section):                                 End Property

Private Property Get DsgnSectionLabel(Optional section As Long) As MSForms.Label:       Set DsgnSectionLabel = cllDsgnSectionsLabel(section):                       End Property

Private Property Get DsgnSections() As Collection:                                      Set DsgnSections = cllDsgnSections:                                         End Property

Private Property Get DsgnSectionText(Optional section As Long) As MSForms.TextBox:      Set DsgnSectionText = cllDsgnSectionsText(section):                         End Property

Private Property Get DsgnSectionTextFrame(Optional ByVal section As Long):              Set DsgnSectionTextFrame = cllDsgnSectionsTextFrame(section):               End Property

Private Property Get DsgnTextFrame(Optional ByVal section As Long) As MSForms.Frame:    Set DsgnTextFrame = cllDsgnSectionsTextFrame(section):                      End Property

Private Property Get DsgnTextFrames() As Collection:                                    Set DsgnTextFrames = cllDsgnSectionsTextFrame:                              End Property

Private Property Let FormWidth(ByVal w As Single)
    Dim siInOutDiff As Single:  siInOutDiff = Me.width - Me.InsideWidth
    Me.width = Max(Me.width, siMinFormWidth, w + siInOutDiff)
End Property

Private Property Let HeightDecrementButtonsArea(ByVal b As Boolean)
    bVscrollbarButtonsArea = b
    bDoneHeightDecrement = b
End Property

Private Property Let HeightDecrementMsgArea(ByVal b As Boolean)
    bVscrollbarMsgArea = b
    bDoneHeightDecrement = b
End Property

Public Property Let HmarginButtons(ByVal si As Single):                                 siHmarginButtons = si:                                                      End Property

Public Property Let HmarginFrames(ByVal si As Single):                                  siHmarginFrames = si:                                                       End Property

Private Property Get IsApplied(Optional ByVal v As Variant) As Boolean
    If dctApplied Is Nothing _
    Then IsApplied = False _
    Else IsApplied = dctApplied.Exists(v)
End Property

Private Property Get MaxButtonsAreaWidth() As Single:                                   MaxButtonsAreaWidth = MaxFormWidthUsable - HSPACE_BUTTON_AREA:              End Property

Public Property Get MaxFormHeight() As Single:                                          MaxFormHeight = siMaxFormHeight:                                            End Property

Public Property Get MaxFormHeightPrcntgOfScreenSize() As Long:                          MaxFormHeightPrcntgOfScreenSize = lMaxFormHeightPoW:                        End Property

Public Property Let MaxFormHeightPrcntgOfScreenSize(ByVal l As Long)
    lMaxFormHeightPoW = l
    siMaxFormHeight = wVirtualScreenHeight * (Min(l, 99) / 100)   ' maximum form height based on screen size
End Property

Public Property Get MaxFormWidth() As Single:                                           MaxFormWidth = siMaxFormWidth:                                              End Property

Public Property Get MaxFormWidthPrcntgOfScreenSize() As Long:                           MaxFormWidthPrcntgOfScreenSize = lMaxFormWidthPoW:                          End Property

Public Property Let MaxFormWidthPrcntgOfScreenSize(ByVal l As Long)
    If l < lMinFormWidthPoW Then l = lMinFormWidthPoW ' The maximum form with cannot be less than the minimum form with
    lMaxFormWidthPoW = l
    siMaxFormWidth = wVirtualScreenWidth * (Min(l, 99) / 100) ' maximum form width based on screen size, limited to 99%
End Property

Private Property Get MaxFormWidthUsable() As Single
    MaxFormWidthUsable = siMaxFormWidth - (Me.width - Me.InsideWidth)
End Property

Private Property Get MaxMsgAreaWidth() As Single:                                       MaxMsgAreaWidth = MaxFormWidthUsable - siHmarginFrames:                     End Property

Private Property Get MaxRowsHeight() As Single:                                         MaxRowsHeight = siMaxButtonHeight + (siVmarginFrames * 2):                  End Property

Private Property Get MaxSectionWidth() As Single:                                       MaxSectionWidth = MaxMsgAreaWidth - siHmarginFrames - HSPACE_SCROLLBAR:     End Property

Private Property Get MaxTextBoxFrameWidth() As Single:                                  MaxTextBoxFrameWidth = MaxSectionWidth - siHmarginFrames:                   End Property

Private Property Get MaxTextBoxWidth() As Single:                                       MaxTextBoxWidth = MaxTextBoxFrameWidth - siHmarginFrames:                   End Property

Public Property Let MinButtonWidth(ByVal si As Single):                                 siMinButtonWidth = si:                                                      End Property

Public Property Get MinFormWidth() As Single:                                           MinFormWidth = siMinFormWidth:                                          End Property

Public Property Let MinFormWidth(ByVal si As Single)
    siMinFormWidth = Max(si, 200) ' cannot be specified less
    '~~ The maximum form width must never not become less than the minimum width
    If siMaxFormWidth < siMinFormWidth Then
       siMaxFormWidth = siMinFormWidth
    End If
    lMinFormWidthPoW = CInt((siMinFormWidth / wVirtualScreenWidth) * 100)
End Property

Public Property Get MinFormWidthPrcntg() As Long:                                       MinFormWidthPrcntg = lMinFormWidthPoW:                                  End Property

Friend Property Let Msg(ByRef tMsg As tMessage)
    With tMsg
        MsgLabel(1) = .section(1).sLabel: MsgText(1) = .section(1).sText:   MsgMonoSpaced(1) = .section(1).bMonspaced
        MsgLabel(2) = .section(2).sLabel: MsgText(2) = .section(2).sText:   MsgMonoSpaced(2) = .section(2).bMonspaced
        MsgLabel(3) = .section(3).sLabel: MsgText(3) = .section(3).sText:   MsgMonoSpaced(3) = .section(3).bMonspaced
    End With
End Property

Public Property Let MsgButtons(ByVal v As Variant)
    Select Case TypeName(v)
        Case "Long", "String":  vButtons = v
        Case Else:              Set vButtons = v
    End Select
End Property

Public Property Get MsgLabel(Optional ByVal section As Long) As String
    If dctSectionsLabel Is Nothing _
    Then MsgLabel = vbNullString _
    Else MsgLabel = IIf(dctSectionsLabel.Exists(section), dctSectionsLabel(section), vbNullString)
End Property

Public Property Let MsgLabel(Optional ByVal section As Long, ByVal s As String)
    If dctSectionsLabel Is Nothing Then Set dctSectionsLabel = New Dictionary
    dctSectionsLabel(section) = s
End Property

Public Property Get MsgMonoSpaced(Optional ByVal section As Long) As Boolean
    If dctSectionsMonoSpaced Is Nothing Then
        MsgMonoSpaced = False
    Else
        With dctSectionsMonoSpaced
            If .Exists(section) _
            Then MsgMonoSpaced = .Item(section) _
            Else MsgMonoSpaced = False
        End With
    End If
End Property

Public Property Let MsgMonoSpaced(Optional ByVal section As Long, ByVal b As Boolean)
    If dctSectionsMonoSpaced Is Nothing Then Set dctSectionsMonoSpaced = New Dictionary
    dctSectionsMonoSpaced(section) = b
End Property

Public Property Get MsgText(Optional ByVal section As Long) As String
    If dctSectionsText Is Nothing Then
        MsgText = vbNullString
    Else
        With dctSectionsText
            If .Exists(section) _
            Then MsgText = .Item(section) _
            Else MsgText = vbNullString
        End With
    End If
End Property

Public Property Let MsgText(Optional ByVal section As Long, ByVal s As String)
    If dctSectionsText Is Nothing Then Set dctSectionsText = New Dictionary
    dctSectionsText(section) = s
End Property

Public Property Let MsgTitle(ByVal s As String):                                        sTitle = s: SetupTitle:                                                     End Property

Private Property Get PrcntgHeightButtonsArea() As Single
    PrcntgHeightButtonsArea = Round(DsgnButtonsArea.Height / (DsgnMsgArea.Height + DsgnButtonsArea.Height), 2)
End Property

Private Property Get PrcntgHeightMsgArea() As Single
    PrcntgHeightMsgArea = Round(DsgnMsgArea.Height / (DsgnMsgArea.Height + DsgnButtonsArea.Height), 2)
End Property

Public Property Get ReplyIndex() As Long:       ReplyIndex = lReplyIndex:   Unload Me:  End Property

Public Property Get ReplyValue() As Variant:    ReplyValue = vReplyValue:   Unload Me:  End Property

Public Property Let TestFrameWithBorders(ByVal b As Boolean)
    
    Dim ctl As MSForms.Control
       
    For Each ctl In Me.Controls
        If TypeName(ctl) = "Frame" Or TypeName(ctl) = "TextBox" Then
            ctl.BorderColor = -2147483638   ' active frame, allows with style none to hide the frame
            If b = False _
            Then ctl.BorderStyle = fmBorderStyleNone _
            Else ctl.BorderStyle = fmBorderStyleSingle
        End If
    Next ctl
    
End Property

Public Property Let TestFrameWithCaptions(ByVal b As Boolean):                          bDisplayFramesWithCaptions = b:                                             End Property

Public Property Get VmarginButtons() As Single:                                         VmarginButtons = siVmarginButtons:                                          End Property

Public Property Let VmarginButtons(ByVal si As Single):                                 siVmarginButtons = si:                                                      End Property

Public Property Get VmarginFrames() As Single:                                          VmarginFrames = siVmarginFrames:                                            End Property

Public Property Let VmarginFrames(ByVal si As Single):                                  siVmarginFrames = VgridPos(si):                                             End Property

Public Sub AdjustStartupPosition(ByRef pUserForm As Object, _
                        Optional ByRef pOwner As Object)
    
    On Error Resume Next
        
    Select Case pUserForm.StartupPosition
        Case Manual, WindowsDefault ' Do nothing
        Case CenterOwner            ' Position centered on top of the 'Owner'. Usually this is Application.
            If Not pOwner Is Nothing Then Set pOwner = Application
            With pUserForm
                .StartupPosition = 0
                .left = pOwner.left + ((pOwner.width - .width) / 2)
                .Top = pOwner.Top + ((pOwner.Height - .Height) / 2)
            End With
        Case CenterScreen           ' Assign the Left and Top properties after switching to Manual positioning.
            With pUserForm
                .StartupPosition = Manual
                .left = (wVirtualScreenWidth - .width) / 2
                .Top = (wVirtualScreenHeight - .Height) / 2
            End With
    End Select
    '~~ Avoid falling off screen. Misplacement can be caused by multiple screens when the primary display
    '~~ is not the left-most screen (which causes "pOwner.Left" to be negative). First make sure the bottom
    '~~ right fits, then check if the top-left is still on the screen (which gets priority).
    With pUserForm
        If ((.left + .width) > (wVirtualScreenLeft + wVirtualScreenWidth)) Then .left = ((wVirtualScreenLeft + wVirtualScreenWidth) - .width)
        If ((.Top + .Height) > (wVirtualScreenTop + wVirtualScreenHeight)) Then .Top = ((wVirtualScreenTop + wVirtualScreenHeight) - .Height)
        If (.left < wVirtualScreenLeft) Then .left = wVirtualScreenLeft
        If (.Top < wVirtualScreenTop) Then .Top = wVirtualScreenTop
    End With
    
End Sub

Public Function AppErr(ByVal lNo As Long) As Long
' -----------------------------------------------------------------------
' Converts a positive (i.e. an "application" error number into a negative
' number by adding vbObjectError. Converts a negative number back into a
' positive i.e. the original programmed application error number.
' Usage example:
'    Err.Raise AppErr(1), .... ' when an application error is detected
'    If Err.Number < 0 Then    ' when the error is displayed
'       MsgBox "Application error " & AppErr(Err.Number)
'    Else
'       MsgBox "VB Rutime Error " & Err.Number
'    End If
' -----------------------------------------------------------------------
    AppErr = IIf(lNo < 0, AppErr = lNo - vbObjectError, AppErr = vbObjectError + lNo)
End Function

Private Sub ApplyScrollBarHorizontal(ByVal fr As MSForms.Frame, _
                                     ByVal widthnew As Single)
                                     
    Dim siScrollWidth   As Single
    
    With fr
        siScrollWidth = .width + 1
        .width = widthnew
        .Height = .Height + VSPACE_SCROLLBAR
    End With
    Select Case fr.ScrollBars
        Case fmScrollBarsBoth
        Case fmScrollBarsHorizontal
            fr.scrollwidth = siScrollWidth
            fr.Scroll xAction:=fmScrollActionNoChange, yAction:=fmScrollActionEnd
        Case fmScrollBarsNone, fmScrollBarsVertical
            fr.ScrollBars = fmScrollBarsHorizontal
            fr.scrollwidth = siScrollWidth
            fr.Scroll xAction:=fmScrollActionNoChange, yAction:=fmScrollActionEnd
    End Select
End Sub

' Apply a vertical scroll bar to the frame (scrollframe) and reduce
' the frames height by a percentage (heightreduction). The original
' frame's height becomes the height of the scroll bar.
' ----------------------------------------------------------------------
Private Sub ApplyScrollBarVertical(ByVal scrollframe As MSForms.Frame, _
                                   ByVal newheight As Single)
        
    Dim siScrollHeight As Single: siScrollHeight = scrollframe.Height + VSPACE_SCROLLBAR
        
    With scrollframe
        .Height = newheight
        Select Case .ScrollBars
            Case fmScrollBarsHorizontal
                .ScrollBars = fmScrollBarsBoth
                .ScrollHeight = siScrollHeight
                .KeepScrollBarsVisible = fmScrollBarsBoth
            Case fmScrollBarsNone
                .ScrollBars = fmScrollBarsVertical
                .ScrollHeight = siScrollHeight
                .KeepScrollBarsVisible = fmScrollBarsVertical
        End Select
    End With
    
End Sub

Private Sub ButtonClicked(ByVal cmb As MSForms.CommandButton)
' -----------------------------------------------------------
' Return the value of the clicked reply button (button).
' When there is only one applied reply button the form is
' unloaded with the click of it. Otherwise the form is just
' hidden waiting for the caller to obtain the return value or
' index which then unloads the form.
' -----------------------------------------------------------
    
    vReplyValue = AppliedButtonRetVal(cmb)
    lReplyIndex = ClickedButtonIndex(cmb)
    If dctApplButtonsRetVal.Count = 1 Then
        Unload Me
    Else
        Me.Hide ' The form will be unloaded when the ReplyValue is fetched by the caller
    End If
    
End Sub

' Center the frame (fr) horizontally within the frame (frin)
' which defaults to the UserForm when not provided.
' -------------------------------------------------------------
Private Sub CenterHorizontal(ByVal centerfr As MSForms.Frame, _
          Optional ByVal infr As MSForms.Frame = Nothing)
    
    If infr Is Nothing _
    Then centerfr.left = (Me.InsideWidth - centerfr.width) / 2 _
    Else centerfr.left = (infr.width - centerfr.width) / 2
    
End Sub

' Center the frame (fr) vertically within the frame (frin).
' -----------------------------------------------------------
Private Sub CenterVertical(ByVal centerfr As MSForms.Frame, _
                           ByVal infr As MSForms.Frame)
    centerfr.Top = (infr.Height / 2) - (centerfr.heigth / 2)
End Sub

' The reply button click event is the only code using
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

' Returns all controls of type (ctltype) which do have a parent (fromparent)
' as collection (into) by assigning the an initial height (ctlheight) and width (ctlwidth).
' -----------------------------------------------------------------------------------------
Private Sub Collect(ByRef into As Variant, _
                    ByVal fromparent As Variant, _
                    ByVal ctltype As String, _
                    ByVal ctlheight As Single, _
                    ByVal ctlwidth As Single)

    Dim ctl As MSForms.Control
    Dim v   As Variant
     
    On Error GoTo on_error
        
    Select Case TypeName(fromparent)
        Case "Collection"
            '~~ Parent is each frame in the collection
            For Each v In fromparent
                For Each ctl In Me.Controls
                    If TypeName(ctl) = ctltype And ctl.Parent Is v Then
                        With ctl
                            .Visible = False
                            .Height = ctlheight
                            .width = ctlwidth
                        End With
                        into.Add ctl
                    End If
               Next ctl
            Next v
        Case Else
            For Each ctl In Me.Controls
                If TypeName(ctl) = ctltype And ctl.Parent Is fromparent Then
                    With ctl
                        .Visible = False
                        .Height = ctlheight
                        .width = ctlwidth
                    End With
                    Select Case TypeName(into)
                        Case "Collection"
                            into.Add ctl
                        Case Else
                            Set into = ctl
                    End Select
                End If
            Next ctl
    End Select

exit_proc:
    Exit Sub
    
on_error:
    Debug.Print err.Description: Stop: Resume
End Sub

 
' Returns pixels (device dependent) to points (used by Excel).
' --------------------------------------------------------------------
Private Sub ConvertPixelsToPoints(ByRef X As Single, ByRef Y As Single)
    
    Dim hDC            As Long
    Dim RetVal         As Long
    Dim PixelsPerInchX As Long
    Dim PixelsPerInchY As Long
 
    On Error Resume Next
    hDC = GetDC(0)
    PixelsPerInchX = GetDeviceCaps(hDC, LOGPIXELSX)
    PixelsPerInchY = GetDeviceCaps(hDC, LOGPIXELSY)
    RetVal = ReleaseDC(0, hDC)
    X = X * TWIPSPERINCH / 20 / PixelsPerInchX
    Y = Y * TWIPSPERINCH / 20 / PixelsPerInchY

End Sub

Private Sub Debug_Sizes(ByVal stage As String, Optional ByVal frSectionMonoSpaced As MSForms.Frame = Nothing)
'    With Me
'        Debug.Print vbLf & stage
'
'        Debug.Print String(Len(stage), "-")
'            Debug.Print "Form (inside) width  = " & Format(.InsideWidth, "##0") & " (specified max = " & Format(.MaxFormWidth, "##0") & "pt = " & .MaxFormWidthPrcntgOfScreenSize & "%)"
'        If IsApplied(DsgnMsgArea) Then _
'            Debug.Print "Message Area  width  = " & Format(DsgnMsgArea.Width, "##0")
'        If Not frSectionMonoSpaced Is Nothing Then _
'            Debug.Print "Monosp. sect. width  = " & frSectionMonoSpaced.Width
'        If IsApplied(DsgnButtonsArea) Then
'            Debug.Print "Buttons Frame width  = " & Format(DsgnButtonsFrame.Width, "##0")
'            Debug.Print "Buttons Area  width  = " & Format(DsgnButtonsArea.Width, "##0")
'        End If
'            Debug.Print "Form (inside) height = " & Format(.InsideHeight, "##0") & " (specified max = " & Format(.MaxFormHeight, "##0") & ")"
'        If IsApplied(DsgnMsgArea) Then _
'            Debug.Print "Message Area  height = " & Format(DsgnMsgArea.Height, "##0") & " (" & PrcntgHeightMsgArea * 100 & "%)"
'        If IsApplied(DsgnButtonsArea) Then
'            Debug.Print "Max button    height = " & siMaxButtonHeight & " (" & dctApplButtonRows.Count & " setup)"
'            Debug.Print "Buttons Frame height = " & Format(DsgnButtonsFrame.Height, "##0")
'            Debug.Print "Buttons Area  height = " & Format(DsgnButtonsArea.Height, "##0") & " (" & PrcntgHeightButtonsArea * 100 & "%)"
'        End If
'    End With
End Sub

' When False (the default) captions are removed from all frames
' Else they remain visible for testing purpose
' -------------------------------------------------------------
Private Sub DisplayFramesWithCaptions(Optional ByVal b As Boolean = True)
            
    Dim ctl As MSForms.Control
       
    If Not b Then
        For Each ctl In Me.Controls
            If TypeName(ctl) = "Frame" Then
                ctl.caption = vbNullString
            End If
        Next ctl
    End If

End Sub

' Return a collection of applied/use/visible buttons in row buttonrow.
' --------------------------------------------------------------------
Private Function DsgnRowButtons(ByVal buttonrow As Long) As Collection
    Set DsgnRowButtons = cllDsgnButtons(buttonrow)
End Function

' Get coordinates of top-left corner and size of entire screen (stretched over
' all monitors) and convert to Points.
' ----------------------------------------------------------------------------
Private Sub GetScreenMetrics()
    
    wVirtualScreenLeft = GetSystemMetrics32(SM_XVIRTUALSCREEN)
    wVirtualScreenTop = GetSystemMetrics32(SM_YVIRTUALSCREEN)
    wVirtualScreenWidth = GetSystemMetrics32(SM_CXVIRTUALSCREEN)
    wVirtualScreenHeight = GetSystemMetrics32(SM_CYVIRTUALSCREEN)
    '
    ConvertPixelsToPoints wVirtualScreenLeft, wVirtualScreenTop
    ConvertPixelsToPoints wVirtualScreenWidth, wVirtualScreenHeight

End Sub

Private Function Max(ParamArray va() As Variant) As Variant
' ------------------------------------------------------
' Returns the maximum value of all values provided (va).
' ------------------------------------------------------
    Dim v   As Variant
    
    Max = va(LBound(va)): If LBound(va) = UBound(va) Then Exit Function
    For Each v In va
        If v > Max Then Max = v
    Next v
    
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

' Reduce the final form height to the maximum height specified by reducing
' one of the two areas by the total exceeding height applying a vertcal
' scroll bar or reducing the height of both areas proportionally and applying
' a vertical scroll bar for both.
' --------------------------------------------------------------------------
Private Sub ReduceAreasHeight(ByVal totalexceedingheight As Single)
    
    Dim frMsgArea               As MSForms.Frame:   Set frMsgArea = DsgnMsgArea
    Dim frButtonsArea           As MSForms.Frame:   Set frButtonsArea = DsgnButtonsArea
    Dim siAreasExceedingHeight  As Single
    
    With Me
        '~~ Reduce the height to the max height specified
        siAreasExceedingHeight = .Height - siMaxFormHeight
        .Height = siMaxFormHeight
        
        If PrcntgHeightMsgArea >= 0.6 Then
            '~~ When the message area requires 60% or more of the total height only this frame
            '~~ will be reduced and applied with a vertical scroll bar.
            ApplyScrollBarVertical scrollframe:=frMsgArea, _
                                         newheight:=frMsgArea.Height - totalexceedingheight
            HeightDecrementMsgArea = True
            
        ElseIf PrcntgHeightButtonsArea >= 0.6 Then
            '~~ When the buttons area requires 60% or more it will be reduced and applied with a vertical scroll bar.
            ApplyScrollBarVertical scrollframe:=frButtonsArea, _
                                     newheight:=frButtonsArea.Height - totalexceedingheight
            HeightDecrementButtonsArea = True

        Else
            '~~ When one area of the two requires less than 60% of the total areas heigth
            '~~ both will be reduced in the height and get a vertical scroll bar.
            ApplyScrollBarVertical scrollframe:=frMsgArea, _
                                      newheight:=frMsgArea.Height * PrcntgHeightMsgArea
            HeightDecrementMsgArea = True
            ApplyScrollBarVertical scrollframe:=frButtonsArea, _
                                   newheight:=frButtonsArea.Height * PrcntgHeightButtonsArea
            HeightDecrementButtonsArea = True
        End If
    End With
    
End Sub

' Reposition all applied/setup control. Executed optionally whenever a control's
' setup had been done (i.e. when a designed control has become an applied/used one),
' obligatory before any height decrement may be due and after any height decrement.
' -----------------------------------------------------------------------------------
Private Sub ResizeAndReposition()
    On Error GoTo on_error
    
    If IsApplied(DsgnMsgArea) Then
        ResizeAndRepositionMsgSections
        ResizeAndRepositionMsgArea
    End If
    If IsApplied(DsgnButtonsArea) Then
        ResizeAndRepositionButtons
        ResizeAndRepositionButtonRows
        ResizeAndRepositionButtonsFrame
        ResizeAndRepositionButtonsArea
    End If
    ResizeAndRepositionAreas
    
exit_proc:
    Exit Sub
    
on_error:
    Debug.Print err.Description: Stop: Resume
End Sub

Private Sub ResizeAndRepositionAreas()

    Dim v       As Variant
    Dim siTop   As Single
    
    siTop = siVmarginFrames
    For Each v In cllDsgnAreas
        With v
            If IsApplied(v) Then
                .Visible = True
                .Top = siTop
                siTop = VgridPos(.Top + .Height + VSPACE_AREAS)
            End If
        End With
    Next v
    Me.Height = VgridPos(siTop + (VSPACE_AREAS * 3))
    
End Sub

Private Sub ResizeAndRepositionButtonRows()
' ---------------------------------------------------------
' Assign all applied/visible button rows the maximum buttons
' height and the width for the number of displayed buttons
' by keeping a record of the maximum button row's width.
' ---------------------------------------------------------
    
    On Error GoTo on_error
    
    Dim cllButtonRows   As Collection:      Set cllButtonRows = DsgnButtonRows
    Dim frRow           As MSForms.Frame
    Dim siTop           As Single
    Dim vButton         As Variant
    Dim lRow            As Long
    Dim dct             As New Dictionary
    Dim v               As Variant
    Dim lButtons        As Long
    Dim siHeight        As Single
    
    '~~ Collect the applied/visible rows and their number of applied/visible buttons
    For lRow = 1 To cllButtonRows.Count
        Set frRow = cllButtonRows(lRow)
        If IsApplied(frRow) Then
            With frRow
                .Visible = True
                For Each vButton In DsgnRowButtons(lRow)
                    If IsApplied(vButton) Then lButtons = lButtons + 1
                Next vButton
                dct.Add frRow, lButtons
            End With
        End If
        lButtons = 0
    Next lRow
    
    '~~ Adjust button row's width and height
    siHeight = AppliedButtonRowHeight
    siTop = siVmarginFrames
    For Each v In dct
        lButtons = dct(v)
        Set frRow = v
        With frRow
            .Top = siTop
            .Height = siHeight
            .width = AppliedButtonRowWidth(lButtons)
            siMaxButtonRowWidth = Max(siMaxButtonRowWidth, .width)
            siTop = .Top + .Height + siVmarginButtons
        End With
    Next v
    Set dct = Nothing
    
exit_proc:
    Exit Sub
    
on_error:
    Debug.Print err.Description: Stop: Resume
End Sub

Private Sub ResizeAndRepositionButtons()
' ---------------------------------------------------------------
' Assign all applied/visible buttons the maximum width and height
' calculated during their setup and adjust their left position.
' ---------------------------------------------------------------
    
    On Error GoTo on_error
    
    Dim cllButtonRows   As Collection:      Set cllButtonRows = DsgnButtonRows
    Dim siLeft          As Single
    Dim frRow           As MSForms.Frame
    Dim vButton         As Variant
    Dim lRow            As Long
    
    For lRow = 1 To cllButtonRows.Count
        siLeft = siHmarginFrames
        Set frRow = cllButtonRows(lRow)
        If IsApplied(frRow) Then
            For Each vButton In DsgnRowButtons(lRow)
                If IsApplied(vButton) Then
                    With vButton
                        .Visible = True
                        .left = siLeft
                        .width = siMaxButtonWidth
                        .Height = siMaxButtonHeight
                        .Top = siVmarginFrames
                        siLeft = .left + .width + siHmarginButtons
                    End With
                End If
            Next vButton
        End If
    Next lRow
        
exit_proc:
    Exit Sub
    
on_error:
    Debug.Print err.Description: Stop: Resume
End Sub

Private Sub ResizeAndRepositionButtonsArea()
' ----------------------------------------------------
' Adjust buttons frame to max button row width and the
' surrounding area's width and heigth is adjusted
' ----------------------------------------------------
    On Error GoTo on_error
    
    Dim frArea      As MSForms.Frame:   Set frArea = DsgnButtonsArea
    Dim frButtons   As MSForms.Frame:   Set frButtons = DsgnButtonsFrame
        
    If IsApplied(frButtons) Then
        With frArea
            .Visible = True
            Select Case .ScrollBars
                Case fmScrollBarsBoth
                    .Height = frButtons.Height + siVmarginFrames + VSPACE_SCROLLBAR
                    .width = frButtons.width + (siHmarginFrames * 2) + HSPACE_SCROLLBAR ' space reserved or used
                    frButtons.left = 0
                Case fmScrollBarsHorizontal
                    .Height = frButtons.Height + (siVmarginFrames + 2) + VSPACE_SCROLLBAR
                    frButtons.left = 0
                Case fmScrollBarsNone
                    .Height = frButtons.Height + (siVmarginFrames * 2)
                    .width = frButtons.width + (siHmarginFrames * 2)
                Case fmScrollBarsVertical
                    .width = frButtons.width + (siHmarginFrames * 2) + HSPACE_SCROLLBAR ' space reserved or used
            End Select
            
            FormWidth = (.width + siHmarginFrames * 2)
            .left = siHmarginFrames
        End With
        If frArea.ScrollBars = fmScrollBarsNone _
        Then CenterHorizontal frButtons, frArea
    
        CenterHorizontal centerfr:=frArea
    End If
    
exit_proc:
    Exit Sub
    
on_error:
    Debug.Print err.Description: Stop: Resume
End Sub

Private Sub ResizeAndRepositionButtonsFrame()
    On Error GoTo on_error
    
    Dim fr  As MSForms.Frame: Set fr = DsgnButtonsFrame
    Dim v   As Variant
    
    If IsApplied(fr) Then
        With DsgnButtonsFrame
            .Visible = True
            .Top = siVmarginFrames
            .width = ButtonsFrameWidth
            .Height = ButtonsFrameHeight
            If bVscrollbarButtonsArea _
            Then .left = siHmarginFrames _
            Else .left = siHmarginFrames + (HSPACE_SCROLLBAR / 2)
        End With
    End If
    '~~ Center all button rows
    For Each v In DsgnButtonRows
        If IsApplied(v) Then CenterHorizontal centerfr:=v, infr:=fr
    Next v
    
exit_proc:
    Exit Sub
    
on_error:
    Debug.Print err.Description: Stop: Resume
End Sub

Private Sub ResizeAndRepositionMsgArea()
' --------------------------------------------------------
' Re-position all applied/used message sections vertically
' and adjust the Message Area height accordingly.
' --------------------------------------------------------
    
    On Error GoTo on_error
    
    Dim frArea      As MSForms.Frame: Set frArea = DsgnMsgArea
    Dim frSection   As MSForms.Frame
    Dim lSection    As Long
    Dim siTop       As Single
            
    If IsApplied(frArea) Then
        siTop = siVmarginFrames
        Me.Height = Max(Me.Height, frArea.Top + frArea.Height + (VSPACE_AREAS * 4))
    End If
    
exit_proc:
    Exit Sub
    
on_error:
    Debug.Print err.Description: Stop: Resume
End Sub

Private Sub ResizeAndRepositionMsgSections()
' -----------------------------------------------------------
' Assign all displayed message sections the required height
' and finally adjust the message area's height.
' -----------------------------------------------------------
    On Error GoTo on_error
    
    Dim frSection       As MSForms.Frame
    Dim i               As Long
    Dim la              As MSForms.Label
    Dim frText          As MSForms.Frame
    Dim tb              As MSForms.TextBox
    Dim siTop           As Single
    Dim siTopSection    As Single
    
    siTopSection = 6
    For i = 1 To cllDsgnSections.Count
        siTop = 0
        If IsApplied(DsgnSection(i)) Then
            Set frSection = DsgnSection(i)
            Set la = DsgnSectionLabel(i)
            Set frText = DsgnSectionTextFrame(i)
            Set tb = DsgnSectionText(i)
            If IsApplied(la) Then
                With la
                    .Visible = True
                    .Top = siTop
                    siTop = VgridPos(.Top + .Height)
                End With
            End If
            
            If IsApplied(tb) Then
                With tb
                    .Visible = True
                    .Top = siVmarginFrames
                End With
                With frText
                    .Visible = True
                    .Top = siTop
                    .Height = tb.Height + (siVmarginFrames * 2)
                    siTop = .Top + .Height + siVmarginFrames
                    If .ScrollBars = fmScrollBarsBoth Or frText.ScrollBars = fmScrollBarsHorizontal Then
                        .Height = tb.Top + tb.Height + VSPACE_SCROLLBAR + siVmarginFrames
                    Else
                        .Height = tb.Top + tb.Height + siVmarginFrames
                    End If
                End With
            End If
                
            If IsApplied(frSection) Then
                With frSection
                    .Top = siTopSection
                    .Visible = True
                    .Height = frText.Top + frText.Height + siVmarginFrames
                    siTopSection = VgridPos(.Top + .Height + siVmarginFrames + VSPACE_SECTIONS)
                End With
            End If
                
            Select Case DsgnMsgArea.ScrollBars
                Case fmScrollBarsBoth, fmScrollBarsVertical:    frSection.left = siHmarginFrames
                Case fmScrollBarsHorizontal, fmScrollBarsNone:  frSection.left = siHmarginFrames + (VSPACE_SCROLLBAR / 2)
            End Select
        End If
    Next i
    
    DsgnMsgArea.Height = frSection.Top + frSection.Height + siVmarginFrames
    
exit_proc:
    Exit Sub
    
on_error:
    Debug.Print err.Description: Stop: Resume
End Sub

Public Sub Setup()
    
    On Error GoTo on_error
        
    CollectDesignControls
       
    DisplayFramesWithCaptions bDisplayFramesWithCaptions ' may be True for test purpose
    
    '~~ ----------------------------------------------------------------------------------------
    '~~ The  p r i m a r y  setup of the title, the message sections and the reply buttons
    '~~ returns their individual widths which determines the minimum required message form width
    '~~ This setup ends width the final message form width and all elements adjusted to it.
    '~~ ----------------------------------------------------------------------------------------
    Me.width = siMinFormWidth ' Setup starts with the minimum message form width
    Me.StartupPosition = 2
    '~~ Setup of those elements which determine the final form width
    If Not bDoneTitle Then SetupTitle
    
    '~~ Setup monospaced message sections
    SetupMsgSectionsMonoSpaced
    ResizeAndReposition
    Debug_Sizes "Monospaced sections setup and resized:"
    
    '~~ Setup the reply buttons
    SetupButtons vButtons
    ResizeAndReposition
    Debug_Sizes "Monospaced sections and buttons setup:"
        
    '~~ At this point the form width is final - possibly with its specified minimum width.
    '~~ The message area width is adjusted to the form's width
    DsgnMsgArea.width = Me.InsideWidth - siHmarginFrames
    
    '~~ Setup proportional spaced message sections (use the given width)
    SetupMsgSectionsPropSpaced
    Debug_Sizes "Message and buttons area setup, reposition due:"
    ResizeAndReposition
    Debug_Sizes "Message and buttons area setup, repositio done:"
            
    '~~ At this point the form height is final. It may however exceed the specified maximum form height.
    '~~ In case the message and/or the buttons area (frame) may be reduced to fit and be provided with
    '~~ a vertical scroll bar. When one area of the two requires less than 60% of the total heigth of both
    '~~ areas, both get a vertical scroll bar, else only the one which uses 60% or more of the height.
    If Me.Height > siMaxFormHeight Then
        Debug_Sizes "Height exceeding max specified"
        '~~ Reduce height to maximum specified and adjust height of message section(s) accordingly
        ReduceAreasHeight totalexceedingheight:=Me.Height - siMaxFormHeight
        bDoneHeightDecrement = True
        Debug_Sizes "Areas had been reduced to fit specified maximum height:"
    End If
    
    ResizeAndReposition
    Debug_Sizes "All done! Setup and (possibly) height reduced:"

    AdjustStartupPosition Me
    bDoneSetup = True ' To indicate for the Activate event that the setup had already be done beforehand
    
exit_proc:
    Exit Sub

on_error:
    Debug.Print err.Description: Stop: Resume
End Sub

Private Sub SetupButton(ByVal buttonrow As Long, _
                        ByVal buttonindex As Long, _
                        ByVal buttoncaption As String, _
                        ByVal buttonreturnvalue As Variant)
' -----------------------------------------------------------------
' Setup an applied reply buttonindex's (buttonindex) visibility and
' caption, calculate the maximum buttonindex width and height,
' keep a record of the setup reply buttonindex's return value.
' -----------------------------------------------------------------
    On Error GoTo on_error
    
    Dim cmb As MSForms.CommandButton:   Set cmb = DsgnButton(buttonrow, buttonindex)
    
    With cmb
        .Visible = True
        .AutoSize = True
        .WordWrap = False ' the longest line determines the buttonindex's width
        .caption = buttoncaption
        .AutoSize = False
        .Height = .Height + 1 ' safety margin to ensure proper multilin caption display
        siMaxButtonHeight = Max(siMaxButtonHeight, .Height)
        siMaxButtonWidth = Max(siMaxButtonWidth, .width, siMinButtonWidth)
    End With
    dctApplButtons.Add cmb, buttonrow
    AppliedButtonRetVal(cmb) = buttonreturnvalue ' keep record of the setup buttonindex's reply value
    AppliedControl = cmb
    AppliedControl = DsgnButtonRow(buttonrow)
    
exit_proc:
    Exit Sub
    
on_error:
    Debug.Print err.Description: Stop: Resume
End Sub

Private Sub SetupButtons(ByVal vButtons As Variant)
' --------------------------------------------------------------------------------------
' Setup and position the applied reply buttons and calculate the max reply button width.
' Note: When the provided vButtons argument is a string it wil be converted into a
'       collection and the procedure is performed recursively with it.
' --------------------------------------------------------------------------------------
    On Error GoTo on_error
    
    Dim frArea  As MSForms.Frame:   Set frArea = DsgnButtonsArea
    Dim cll     As Collection
    Dim v       As Variant
    
    AppliedControl = frArea
    AppliedControl = DsgnButtonsFrame
    lSetupRows = 1
    
    '~~ Setup all reply button by calculatig their maximum width and height
    Select Case TypeName(vButtons)
        Case "Long":        SetupButtonsFromValue vButtons ' buttons are specified by an MsgBox buttons value only
        Case "String":      SetupButtonsFromString vButtons
        Case "Collection":  SetupButtonsFromCollection vButtons
        Case "Dictionary":  SetupButtonsFromCollection vButtons
        Case Else
'            MsgBox "The format of the provided ""buttons"" argument is not supported!" & vbLf & _
'                   "The message will be setup with an Ok only button", vbExclamation
            SetupButtons vbOKOnly
    End Select
            
    ResizeAndReposition
    
    If frArea.width > MaxButtonsAreaWidth Then
        Debug_Sizes "Buttons area width exceeds maximum width specified:"
        ApplyScrollBarHorizontal fr:=frArea, widthnew:=MaxButtonsAreaWidth
        bHscrollbarButtonsArea = True
        Me.width = siMaxFormWidth
        frArea.Height = frArea.Height + VSPACE_SCROLLBAR
        CenterHorizontal frArea
        Debug_Sizes "Buttons area width decremented:"
    End If

    bDoneButtonsArea = True
    
exit_proc:
    Exit Sub
    
on_error:
    Debug.Print err.Description: Stop: Resume
End Sub

' Setup the reply buttons based on the comma delimited string of button captions
' and row breaks indicated by a vbLf, vbCr, or vbCrLf.
' ------------------------------------------------------------------------------
Private Sub SetupButtonsFromCollection(ByVal cllButtons As Collection)

    On Error GoTo on_error
    
    Dim v           As Variant
    
    lSetupRows = 1
    lSetupRowButtons = 0
    
    For Each v In cllButtons
        Select Case v
            Case vbOKOnly
                SetupButtonsFromValue v
            Case vbOKCancel, vbYesNo, vbRetryCancel
                SetupButtonsFromValue v
            Case vbYesNoCancel, vbAbortRetryIgnore
                SetupButtonsFromValue v
            Case Else
                If v <> vbNullString Then
                    If v = vbLf Or v = vbCr Or v = vbCrLf Then
                        '~~ prepare for the next row
                        If lSetupRows <= 7 Then ' ignore exceeding rows
                            If Not dctApplButtonRows.Exists(DsgnButtonRow(lSetupRows)) Then dctApplButtonRows.Add DsgnButtonRow(lSetupRows), lSetupRows
                            AppliedControl = DsgnButtonRow(lSetupRows)
                            lSetupRows = lSetupRows + 1
                            lSetupRowButtons = 0
                        Else
                            MsgBox "Setup of button row " & lSetupRows & " ignored! The maximimum applicable rows is 7."
                        End If
                    Else
                        lSetupRowButtons = lSetupRowButtons + 1
                        If lSetupRowButtons <= 7 Then
                            DsgnButtonRow(lSetupRows).Visible = True
                            SetupButton buttonrow:=lSetupRows, buttonindex:=lSetupRowButtons, buttoncaption:=v, buttonreturnvalue:=v
                        Else
                            MsgBox "Setup of a button " & lSetupRowButtons & " in row " & lSetupRows & " ignored! The maximimum applicable buttons per row is 7."
                        End If
                    End If
                End If
        End Select
    Next v
    If lSetupRows <= 7 Then
        If Not dctApplButtonRows.Exists(DsgnButtonRow(lSetupRows)) Then dctApplButtonRows.Add DsgnButtonRow(lSetupRows), lSetupRows
        AppliedControl = DsgnButtonRow(lSetupRows)
    End If
    DsgnButtonsArea.Visible = True
    
exit_proc:
    Exit Sub
    
on_error:
    Debug.Print err.Description: Stop: Resume
End Sub

Private Sub SetupButtonsFromString(ByVal sButtons As String)
    
    Dim cll As New Collection
    Dim v   As Variant
    
    For Each v In Split(vButtons, ",")
        cll.Add v
    Next v
    SetupButtons cll
    
End Sub

Private Sub SetupButtonsFromValue(ByVal lButtons As Long)
' -------------------------------------------------------
' Setup a row of standard VB MsgBox reply command buttons
' -------------------------------------------------------
    On Error GoTo on_error
    
    Select Case lButtons
        Case vbOKOnly
            lSetupRowButtons = lSetupRowButtons + 1
            SetupButton buttonrow:=lSetupRows, buttonindex:=lSetupRowButtons, buttoncaption:="Ok", buttonreturnvalue:=vbOK
        Case vbOKCancel
            lSetupRowButtons = lSetupRowButtons + 1
            SetupButton buttonrow:=lSetupRows, buttonindex:=lSetupRowButtons, buttoncaption:="Ok", buttonreturnvalue:=vbOK
            lSetupRowButtons = lSetupRowButtons + 1
            SetupButton buttonrow:=lSetupRows, buttonindex:=lSetupRowButtons, buttoncaption:="Cancel", buttonreturnvalue:=vbCancel
        Case vbYesNo
            lSetupRowButtons = lSetupRowButtons + 1
            SetupButton buttonrow:=lSetupRows, buttonindex:=lSetupRowButtons, buttoncaption:="Yes", buttonreturnvalue:=vbYes
            lSetupRowButtons = lSetupRowButtons + 1
            SetupButton buttonrow:=lSetupRows, buttonindex:=lSetupRowButtons, buttoncaption:="No", buttonreturnvalue:=vbNo
        Case vbRetryCancel
            lSetupRowButtons = lSetupRowButtons + 1
            SetupButton buttonrow:=lSetupRows, buttonindex:=lSetupRowButtons, buttoncaption:="Retry", buttonreturnvalue:=vbRetry
            lSetupRowButtons = lSetupRowButtons + 1
            SetupButton buttonrow:=lSetupRows, buttonindex:=lSetupRowButtons, buttoncaption:="Cancel", buttonreturnvalue:=vbCancel
        Case vbYesNoCancel
            lSetupRowButtons = lSetupRowButtons + 1
            SetupButton buttonrow:=lSetupRows, buttonindex:=lSetupRowButtons, buttoncaption:="Yes", buttonreturnvalue:=vbYes
            lSetupRowButtons = lSetupRowButtons + 1
            SetupButton buttonrow:=lSetupRows, buttonindex:=lSetupRowButtons, buttoncaption:="No", buttonreturnvalue:=vbNo
            lSetupRowButtons = lSetupRowButtons + 1
            SetupButton buttonrow:=lSetupRows, buttonindex:=lSetupRowButtons, buttoncaption:="Cancel", buttonreturnvalue:=vbCancel
        Case vbAbortRetryIgnore
            lSetupRowButtons = lSetupRowButtons + 1
            SetupButton buttonrow:=lSetupRows, buttonindex:=lSetupRowButtons, buttoncaption:="Abort", buttonreturnvalue:=vbAbort
            lSetupRowButtons = lSetupRowButtons + 1
            SetupButton buttonrow:=lSetupRows, buttonindex:=lSetupRowButtons, buttoncaption:="Retry", buttonreturnvalue:=vbRetry
            lSetupRowButtons = lSetupRowButtons + 1
            SetupButton buttonrow:=lSetupRows, buttonindex:=lSetupRowButtons, buttoncaption:="Ignore", buttonreturnvalue:=vbIgnore
        Case Else
            MsgBox "The value provided for the ""buttons"" argument is not a known VB MsgBox value"
    End Select
    DsgnButtonsArea.Visible = True
    DsgnButtonRow(lSetupRows).Visible = True
    If Not dctApplButtonRows.Exists(DsgnButtonRow(lSetupRows)) Then dctApplButtonRows.Add DsgnButtonRow(lSetupRows), lSetupRows
    AppliedControl = DsgnButtonRow(lSetupRows)
    AppliedControl = DsgnButtonsFrame
    
exit_proc:
    Exit Sub
    
on_error:
    Debug.Print err.Description: Stop: Resume
End Sub

' Setup a message section with its label when one is specified
' and return the message's width when greater than any other.
' Note: All height adjustments except the one for the text box
'       are done by the ResizeAndReposition
' -------------------------------------------------------------
Private Sub SetupMsgSection(ByVal section As Long)
    On Error GoTo on_error
    
    Dim frArea      As MSForms.Frame
    Dim frSection   As MSForms.Frame
    Dim la          As MSForms.Label
    Dim tbText      As MSForms.TextBox
    Dim frText      As MSForms.Frame
    Dim sLabel      As String
    Dim sText       As String
    Dim bMonospaced As Boolean

    Set frArea = DsgnMsgArea
    Set frSection = DsgnSection(section)
    Set la = DsgnSectionLabel(section)
    Set tbText = DsgnSectionText(section)
    Set frText = DsgnTextFrame(section)
    
    sLabel = Me.MsgLabel(section)
    sText = Me.MsgText(section)
    bMonospaced = Me.MsgMonoSpaced(section)
    
    frSection.width = frArea.width
    la.width = frSection.width
    frText.width = frSection.width
    tbText.width = frSection.width
        
    If sText <> vbNullString Then
    
        AppliedControl = frArea
        AppliedControl = frSection
        AppliedControl = frText
        AppliedControl = tbText
                
        If sLabel <> vbNullString Then
            Set la = DsgnSectionLabel(section)
            With la
                .width = Me.InsideWidth - (siHmarginFrames * 2)
                .caption = sLabel
            End With
            frText.Top = la.Top + la.Height
            AppliedControl = la
        Else
            frText.Top = 0
        End If
        
        If bMonospaced Then
            SetupMsgSectionMonoSpaced section, sText  ' returns the maximum width required for monospaced section
        Else ' proportional spaced
            SetupMsgSectionPropSpaced section, sText
        End If
        tbText.SelStart = 0
        
    End If
    
exit_proc:
    Exit Sub
    
on_error:
    Debug.Print err.Description: Stop: Resume
End Sub

' Setup the applied monospaced message section (section) with the text (text),
' and apply width and adjust surrounding frames accordingly.
' Note: All height adjustments except the one for the text box
'       are done by the ResizeAndReposition
' --------------------------------------------------------------------
Private Sub SetupMsgSectionMonoSpaced(ByVal section As Long, _
                                       ByVal text As String)
    On Error GoTo on_error
    
    Dim frArea          As MSForms.Frame:   Set frArea = DsgnMsgArea
    Dim frText          As MSForms.Frame:   Set frText = DsgnSectionTextFrame(section)
    Dim tbText          As MSForms.TextBox: Set tbText = DsgnSectionText(section)
    Dim frSection       As MSForms.Frame:   Set frSection = DsgnSection(section)
    
    '~~ Setup the textbox
    With tbText
        .Visible = True
        .MultiLine = True
        .WordWrap = False
        .Font.Name = sMonoSpacedFontName
        .Font.Size = siMonoSpacedFontSize
        .AutoSize = True
        .value = text
        .AutoSize = False
        .SelStart = 0
        .left = siHmarginFrames
        .Height = .Height + 2 ' ensure text is not squeeced
        frText.width = .width + (siHmarginFrames * 2)
        frText.left = siHmarginFrames
                   
        frSection.width = frText.width + (siHmarginFrames * 2)
        
        '~~ The area width considers that there might be a need to apply a vertival scroll bar
        '~~ When the space finally isn't required, the sections are centered within the area
        frArea.width = Max(frArea.width, frSection.left + frSection.width + siHmarginFrames + HSPACE_SCROLLBAR)
        FormWidth = frArea.width + siHmarginFrames + 7
        
        If .width > MaxTextBoxWidth Then
            frSection.width = MaxSectionWidth
            frArea.width = MaxMsgAreaWidth
            Me.width = siMaxFormWidth
            ApplyScrollBarHorizontal fr:=frText, widthnew:=MaxTextBoxFrameWidth
        End If
        
    End With
    siMaxSectionWidth = Max(siMaxSectionWidth, frSection.width)
    
    '~~ Keep record of which controls had been applied
    AppliedControl = frArea
    AppliedControl = frSection
    AppliedControl = frText
    AppliedControl = tbText
    
    Debug_Sizes "Monospaced sections setup:", frSection
    
exit_proc:
    Exit Sub
    
on_error:
    Debug.Print err.Description: Stop: Resume
End Sub

' Setup the proportional spaced Message Section (section) with the text (text)
' Note 1: When proportional spaced Message Sections are setup the width of the
'         Message Form is already final.
' Note 2: All height adjustments except the one for the text box
'         are done by the ResizeAndReposition
' -----------------------------------------------------------------------------
Private Sub SetupMsgSectionPropSpaced(ByVal section As Long, _
                                        ByVal text As String)
    
    Dim frArea      As MSForms.Frame
    Dim frSection   As MSForms.Frame
    Dim frText      As MSForms.Frame
    Dim tbText      As MSForms.TextBox
    
    Set frArea = DsgnMsgArea
    Set frSection = DsgnSection(section)
    Set frText = DsgnSectionTextFrame(section)
    Set tbText = DsgnSectionText(section)
        
    '~~ For proportional spaced message sections the width is determined by the area width
    With frSection
        .width = frArea.width - siHmarginFrames - HSPACE_SCROLLBAR
        .left = HSPACE_LEFT
        siMaxSectionWidth = Max(siMaxSectionWidth, .width)
    End With
    With frText
        .width = frSection.width - siHmarginFrames
        .left = HSPACE_LEFT
    End With
    
    With tbText
        .Visible = True
        .MultiLine = True
        .AutoSize = True
        .WordWrap = True
        .width = frText.width - siHmarginFrames
        .value = text
        .SelStart = 0
        .left = HSPACE_LEFT
        frText.width = .left + .width + siHmarginFrames
    End With
    
    '~~ Keep record of which controls had been applied
    AppliedControl = frArea
    AppliedControl = frSection
    AppliedControl = frText
    AppliedControl = tbText

End Sub

Private Sub SetupMsgSectionsMonoSpaced()
    On Error GoTo on_error
                             
    If MsgText(1) <> vbNullString And MsgMonoSpaced(1) = True Then SetupMsgSection section:=1
    If MsgText(2) <> vbNullString And MsgMonoSpaced(2) = True Then SetupMsgSection section:=2
    If MsgText(3) <> vbNullString And MsgMonoSpaced(3) = True Then SetupMsgSection section:=3
    bDoneMonoSpacedSections = True
    
exit_proc:
    Exit Sub
    
on_error:
    Debug.Print err.Description: Stop: Resume
End Sub

Private Sub SetupMsgSectionsPropSpaced()
    On Error GoTo on_error
    
    If MsgText(1) <> vbNullString And MsgMonoSpaced(1) = False Then SetupMsgSection section:=1
    If MsgText(2) <> vbNullString And MsgMonoSpaced(2) = False Then SetupMsgSection section:=2
    If MsgText(3) <> vbNullString And MsgMonoSpaced(3) = False Then SetupMsgSection section:=3
    bDonePropSpacedSections = True
    bDoneMsgArea = True
    
exit_proc:
    Exit Sub
    
on_error:
    Debug.Print err.Description: Stop: Resume
End Sub

' When a specific font name and/or size is specified, the extra title label is actively used
' and the UserForm's title bar is not displayed - which means that there is no X to cancel.
' ------------------------------------------------------------------------------------------
Private Sub SetupTitle()
    
    Dim siTop           As Single
    Dim siTitleWidth    As Single
    
    siTop = 0
    With Me
        '~~ When a font name other then the standard UserForm font name is
        '~~ provided the extra hidden title label which mimics the title bar
        '~~ width is displayed. Otherwise it remains hidden.
        If sTitleFontName <> vbNullString And sTitleFontName <> .Font.Name Then
            With .laMsgTitle   ' Hidden by default
                .Visible = True
                .Top = siTop
                siTop = VgridPos(.Top + .Height)
                .Font.Name = sTitleFontName
                If sTitleFontSize <> 0 Then
                    .Font.Size = sTitleFontSize
                End If
                .AutoSize = True
                .caption = " " & sTitle    ' some left margin
                siTitleWidth = .width + HSPACE_RIGHT
            End With
            AppliedControl = .laMsgTitle
            .laMsgTitleSpaceBottom.Visible = True
        Else
            '~~ The extra title label is only used to adjust the form width and remains hidden
            With .laMsgTitle
                With .Font
                    .Bold = False
                    .Name = Me.Font.Name
                    .Size = 8.65   ' Value which comes to a length close to the length required
                End With
                .Visible = False
                .AutoSize = True
                .caption = " " & sTitle    ' some left margin
                siTitleWidth = .width + 30
            End With
            .caption = " " & sTitle    ' some left margin
            .laMsgTitleSpaceBottom.Visible = False
        End If
                
        .laMsgTitleSpaceBottom.width = siTitleWidth
        FormWidth = siTitleWidth
    End With
    bDoneTitle = True
    
End Sub

Private Sub UserForm_Activate()
    '~~ To avoid screen flicker the setup may has been done already.
    '~~ However for test purpose Setup may run with the Activate event i.e. the .Show
    If bDoneSetup = True Then
        bDoneSetup = False
    Else
        Setup
    End If
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
