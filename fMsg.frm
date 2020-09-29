VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} fMsg 
   ClientHeight    =   5475
   ClientLeft      =   45
   ClientTop       =   390
   ClientWidth     =   12690
   OleObjectBlob   =   "fMsg.frx":0000
   StartUpPosition =   1  'Fenstermitte
End
Attribute VB_Name = "fMsg"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
' --------------------------------------------------------------------------
' UserForm fMsg Provides all means for a message with
'               - up to 3 separated text messages, each either with a
'                 proportional or a fixed font
'               - each of the 3 messages with an optional label
'               - 4 reply buttons either specified with replies known
'                 from the VB MsgBox or any test string.
'
' W. Rauschenberger Berlin March 2020
' --------------------------------------------------------------------------
Const H_MARGIN                  As Single = 15
Const V_MARGIN                  As Single = 10
Const MIN_FORM_WIDTH            As Single = 200
Const MIN_REPLY_WIDTH           As Single = 70
Dim lFixedFontMessageLines      As Long
Dim sTitle                      As String
Dim sErrSrc                     As String
Dim vReplies                    As Variant
Dim aReplies                    As Variant
Dim lReplies                    As Long
Dim siFormWidth                 As Single
Dim sTitleFontName              As String
Dim sTitleFontSize              As String ' Ignored when sTitleFontName is not provided
Dim sProportionalFontMessage    As String
Dim siTopNext                   As Single
Dim bWithLabel                  As Boolean
Dim sMsg1Proportional           As String
Dim sMsg2Proportional           As String
Dim sMsg3Proportional           As String
Dim sMsg1Fixed                  As String
Dim sMsg2Fixed                  As String
Dim sMsg3Fixed                  As String
Dim sLabelMessage1              As String
Dim sLabelMessage2              As String
Dim sLabelMessage3              As String
Dim siTitleWidth                As Single
Dim siMaxFixedTextWidth         As Single
Dim siMaxReplyWidth             As Single
Dim siMaxReplyHeight            As Single

Private Sub UserForm_Initialize()
    siFormWidth = MIN_FORM_WIDTH ' Default
End Sub

Public Property Let ErrSrc(ByVal s As String):                  sErrSrc = s:                                    End Property
Public Property Let FormWidth(ByVal si As Single):              siFormWidth = si:                               End Property
Public Property Let LabelMessage1(ByVal s As String):           sLabelMessage1 = s:                             End Property
Public Property Let LabelMessage2(ByVal s As String):           sLabelMessage2 = s:                             End Property
Public Property Let LabelMessage3(ByVal s As String):           sLabelMessage3 = s:                             End Property
Private Property Get LabelMsg1() As MSForms.Label:              Set LabelMsg1 = Me.laMsg1:                      End Property
Private Property Get LabelMsg2() As MSForms.Label:              Set LabelMsg2 = Me.laMsg2:                      End Property
Private Property Get LabelMsg3() As MSForms.Label:              Set LabelMsg3 = Me.laMsg3:                      End Property
Public Property Let Message1Fixed(ByVal s As String):           sMsg1Fixed = s:                                 End Property
Public Property Let Message1Proportional(ByVal s As String):    sMsg1Proportional = s:                          End Property
Public Property Let Message2Fixed(ByVal s As String):           sMsg2Fixed = s:                                 End Property
Public Property Let Message2Proportional(ByVal s As String):    sMsg2Proportional = s:                          End Property
Public Property Let Message3Fixed(ByVal s As String):           sMsg3Fixed = s:                                 End Property
Public Property Let Message3Proportional(ByVal s As String):    sMsg3Proportional = s:                          End Property
Private Property Get Msg1Fixed() As MSForms.TextBox:            Set Msg1Fixed = Me.tbMsg1Fixed:                 End Property
Private Property Get Msg1Proportional() As MSForms.TextBox:     Set Msg1Proportional = Me.tbMsg1Proportional:   End Property
Private Property Get Msg2Fixed() As MSForms.TextBox:            Set Msg2Fixed = Me.tbMsg2Fixed:                 End Property
Private Property Get Msg2Proportional() As MSForms.TextBox:     Set Msg2Proportional = Me.tbMsg2Proportional:   End Property
Private Property Get Msg3Fixed() As MSForms.TextBox:            Set Msg3Fixed = Me.tbMsg3Fixed:                 End Property
Private Property Get Msg3Proportional() As MSForms.TextBox:     Set Msg3Proportional = Me.tbMsg3Proportional:   End Property
Public Property Let Title(ByVal s As String):                   sTitle = s:                                     End Property
Public Property Let TitleFontName(ByVal s As String):           sTitleFontName = s:                             End Property
Public Property Let TitleFontSize(ByVal l As Long):             sTitleFontSize = l:                             End Property

Private Property Get TopNext(Optional ByVal ctl As Variant) As Single
Dim tb  As MSForms.TextBox
Dim la  As MSForms.Label
Dim cb  As MSForms.CommandButton

    TopNext = siTopNext ' Return the current position for control

    With ctl            ' Increase the top position for the next control
        .Top = siTopNext
        Select Case TypeName(ctl)
            Case "TextBox"
                Set tb = ctl
                siTopNext = tb.Top + tb.Height + V_MARGIN
            Case "CommandButton"
                Set cb = ctl
                siTopNext = cb.Top + cb.Height + V_MARGIN
            Case "Label"
                Set la = ctl
                Select Case la.Name
                    Case "la"
                        siTopNext = .laTitleSpaceBottom.Top + .laTitleSpaceBottom.Height + V_MARGIN
                    Case Else ' Message label
                        siTopNext = la.Top + la.Height
                End Select
        End Select
    End With
End Property

Public Property Let Replies(ByVal v As Variant)
    vReplies = v
    aReplies = Split(v, ",")
    lReplies = UBound(aReplies) + 1
End Property

Private Sub SetupFixedMsgTextWith( _
            ByVal tb As MSForms.TextBox, _
            ByVal sText As String, _
            ByVal bFixed As Boolean)
' ----------------------------------------
' Setup the width of textbox (tb)
' considering the test (sText) and fixed
' or proportional font (bFixed).
' ----------------------------------------
Dim sSplit      As String
Dim v           As Variant
Dim siMaxWidth  As Single

    If bFixed Then
        '~~ A fixed font Textbox's width is determined by the maximum text line length,
        '~~ determined by means of an autosized width-template
        If InStr(sText, vbLf) <> 0 Then sSplit = vbLf
        If InStr(sText, vbCrLf) <> 0 Then sSplit = vbCrLf
        '~~ Find the width which fits the largest text line
        With Me.tbMsgFixedWidthTemplate
            .MultiLine = False
            .WordWrap = False
            For Each v In Split(sText, sSplit)
                .Value = v
                siMaxWidth = Max(siMaxWidth, .Width)
            Next v
        End With
        
        tb.Width = Max(siMaxWidth, Me.laTitle.Width) + H_MARGIN
        siMaxFixedTextWidth = mBasic.Max(siMaxFixedTextWidth, tb.Width)
    End If

End Sub

Private Sub cmbReply1_Click()
    With Me.cmbReply1
        Select Case UCase(.Caption)
            Case "OK":      ReplyWith vbOK
            Case "YES":     ReplyWith vbYes
            Case "NO":      ReplyWith vbNo
            Case "CANCEL":  ReplyWith vbCancel
            Case Else:      ReplyWith .Caption
        End Select
    End With
End Sub

Private Sub cmbReply2_Click()
    With Me.cmbReply2
        Select Case UCase(.Caption)
            Case "OK":      ReplyWith vbOK
            Case "YES":     ReplyWith vbYes
            Case "NO":      ReplyWith vbNo
            Case "CANCEL":  ReplyWith vbCancel
            Case Else:      ReplyWith .Caption
        End Select
    End With
End Sub

Private Sub cmbReply3_Click()
    With Me.cmbReply3
        Select Case UCase(.Caption)
            Case "OK":      ReplyWith vbOK
            Case "YES":     ReplyWith vbYes
            Case "NO":      ReplyWith vbNo
            Case "CANCEL":  ReplyWith vbCancel
            Case Else:      ReplyWith .Caption
        End Select
    End With
End Sub

Private Sub cmbReply4_Click()
    With Me.cmbReply4
        Select Case UCase(.Caption)
            Case "OK":      ReplyWith vbOK
            Case "YES":     ReplyWith vbYes
            Case "NO":      ReplyWith vbNo
            Case "CANCEL":  ReplyWith vbCancel
            Case Else:      ReplyWith .Caption
        End Select
    End With
End Sub

Private Sub cmbReply5_Click()
    With Me.cmbReply5
        Select Case UCase(.Caption)
            Case "OK":      ReplyWith vbOK
            Case "YES":     ReplyWith vbYes
            Case "NO":      ReplyWith vbNo
            Case "CANCEL":  ReplyWith vbCancel
            Case Else:      ReplyWith .Caption
        End Select
    End With
End Sub

Private Sub ReplyWith(ByVal v As Variant)
    mBasic.MsgReply = v
    Unload Me
End Sub

Private Sub SetupMessageFixed( _
            ByVal la As MSForms.Label, _
            ByVal sLabelText As String, _
            ByVal tb As MSForms.TextBox, _
            ByVal sTextBoxText As String)
' -----------------------------------------------------
' Any fixed font message's width is adjusted to the
' maximum line width.
' -----------------------------------------------------
Dim v           As Variant
Dim siMaxWidth  As Single

    If sTextBoxText <> vbNullString Then
        With la
            '~~ Error Path: Adjust top position and height
            If sLabelText <> vbNullString Then
                With la
                    .Caption = sLabelText
                    .Visible = True
                    .Top = TopNext(la) ' Assign and increase for next
                End With
            End If
        End With
        
        With tb
            .Visible = True
            SetupFixedMsgTextWith tb, sTextBoxText, bFixed:=True
            .MultiLine = True
            .WordWrap = True
            .AutoSize = True
            .Value = sTextBoxText
            .Top = TopNext(tb)  ' Assign and increase for next
        End With
        
        With Me
            .Width = mBasic.Max(MIN_FORM_WIDTH, _
                                 siFormWidth, _
                                 .laTitle.Width, _
                                 tb.Left + tb.Width + H_MARGIN)
            .laTitle.Width = .Width
            .laTitleSpaceBottom.Width = .Width
        End With
        
    End If
End Sub

Private Sub SetupMessageProportional( _
            ByVal la As MSForms.Label, _
            ByVal sLabelText As String, _
            ByVal tb As MSForms.TextBox, _
            ByVal sTextBoxText As String)
' ---------------------------------------
' Adjust message width to form width
' ---------------------------------------
    If sTextBoxText <> vbNullString Then
        '~~ Setup Message Label
        If sLabelText <> vbNullString Then
            With la
                .Caption = sLabelText
                .Visible = True
                .Top = TopNext(la)  ' Assign and increase for next control
            End With
        End If
        
        '~~ Setup Message Textbox
        With tb
            .Visible = True
            .MultiLine = True
            .WordWrap = True
            .Width = Me.Width - H_MARGIN
            .AutoSize = True
            .Value = sTextBoxText
            .Top = TopNext(tb)  ' Assign and increase for next control
        End With
    End If

End Sub

Private Sub SetupRepliesTopPos()
Dim siTop   As Single

    With Me
        With .cmbReply1
            .Top = TopNext(Me.cmbReply1)
            siTop = .Top
            .Height = siMaxReplyHeight
        End With
        With .cmbReply2
            .Top = siTop
            .Height = siMaxReplyHeight
        End With
        With .cmbReply3
            .Top = siTop
            .Height = siMaxReplyHeight
        End With
        With .cmbReply4
            .Top = siTop
            .Height = siMaxReplyHeight
        End With
        With .cmbReply5
            .Top = siTop
            .Height = siMaxReplyHeight
        End With
        .Height = siTop + siMaxReplyHeight + (V_MARGIN * 5)
    End With
    
End Sub

Private Sub SetupReplyButton( _
            ByVal cmb As MSForms.CommandButton, _
            ByVal s As String)
' -----------------------------------------------
' Setup Command Button's visibility and text.
' -----------------------------------------------
    With cmb
        .Visible = True
        .Caption = s
        siMaxReplyHeight = mBasic.Max(siMaxReplyHeight, .Height)
    End With
End Sub

Private Sub SetupReplyButtons(ByVal vReplies As Variant)
' ------------------------------------------------------
' Setup and position the reply buttons. Returns the max
' reply button width.
' ------------------------------------------------------

    With Me
        '~~ Setup button caption
        Select Case vReplies
            Case vbOKOnly, "OK"
                lReplies = 1
                SetupReplyButton .cmbReply1, "Ok"
                siMaxReplyWidth = Max(siMaxReplyWidth, .cmbReply1.Width, MIN_REPLY_WIDTH)
            Case vbYesNo
                lReplies = 2
                SetupReplyButton .cmbReply1, "Yes"
                SetupReplyButton .cmbReply2, "No"
                siMaxReplyWidth = Max(siMaxReplyWidth, .cmbReply2.Width, MIN_REPLY_WIDTH)
            Case vbOKCancel
                lReplies = 2
                SetupReplyButton .cmbReply1, "OK"
                SetupReplyButton .cmbReply2, "Cancel"
                siMaxReplyWidth = Max(siMaxReplyWidth, .cmbReply2.Width, MIN_REPLY_WIDTH)
            Case vbYesNoCancel
                lReplies = 3
                SetupReplyButton .cmbReply1, "Yes"
                SetupReplyButton .cmbReply2, "No"
                SetupReplyButton .cmbReply3, "Cancel"
                siMaxReplyWidth = Max(siMaxReplyWidth, .cmbReply3.Width, MIN_REPLY_WIDTH)
            Case Else
                Select Case lReplies
                    Case 1
                        SetupReplyButton .cmbReply1, aReplies(0)
                        siMaxReplyWidth = Max(siMaxReplyWidth, .cmbReply1.Width, MIN_REPLY_WIDTH)
                    Case 2
                        SetupReplyButton .cmbReply1, aReplies(0)
                        SetupReplyButton .cmbReply2, aReplies(1)
                        siMaxReplyWidth = Max(siMaxReplyWidth, .cmbReply1.Width, .cmbReply2.Width, MIN_REPLY_WIDTH)
                    Case 3
                        SetupReplyButton .cmbReply1, aReplies(0)
                        SetupReplyButton .cmbReply2, aReplies(1)
                        SetupReplyButton .cmbReply3, aReplies(2)
                        siMaxReplyWidth = Max(siMaxReplyWidth, .cmbReply1.Width, .cmbReply2.Width, .cmbReply3.Width, MIN_REPLY_WIDTH)
                    Case 4
                        SetupReplyButton .cmbReply1, aReplies(0)
                        SetupReplyButton .cmbReply2, aReplies(1)
                        SetupReplyButton .cmbReply3, aReplies(2)
                        SetupReplyButton .cmbReply4, aReplies(3)
                        siMaxReplyWidth = Max(siMaxReplyWidth, .cmbReply1.Width, .cmbReply2.Width, .cmbReply3.Width, .cmbReply4.Width, MIN_REPLY_WIDTH)
                    Case 5
                        SetupReplyButton .cmbReply1, aReplies(0)
                        SetupReplyButton .cmbReply2, aReplies(1)
                        SetupReplyButton .cmbReply3, aReplies(2)
                        SetupReplyButton .cmbReply4, aReplies(3)
                        SetupReplyButton .cmbReply5, aReplies(4)
                        siMaxReplyWidth = Max(siMaxReplyWidth, .cmbReply1.Width, .cmbReply2.Width, .cmbReply3.Width, .cmbReply4.Width, .cmbReply5.Width, MIN_REPLY_WIDTH)
                End Select
        End Select
    End With

End Sub

Private Sub SetupReplyButtonsHPos()
' --------------------------------------------------------------
' Setup for each reply button its left position.
' --------------------------------------------------------------

    With Me
        Select Case lReplies
            Case 1
                With .cmbReply1
                    .Width = siMaxReplyWidth
                    .Left = (Me.Width / 2) - (siMaxReplyWidth / 2) ' center
                End With
            Case 2
                With .cmbReply1
                    .Width = siMaxReplyWidth
                    .Left = (Me.Width / 2) - (V_MARGIN / 2) - siMaxReplyWidth ' left from center
                End With
                With .cmbReply2
                    .Width = siMaxReplyWidth
                    .Left = Me.cmbReply1.Left + siMaxReplyWidth + V_MARGIN ' right from center
                End With
            Case 3
                With .cmbReply2
                    .Width = siMaxReplyWidth
                    .Left = (Me.Width / 2) - (siMaxReplyWidth / 2) ' center
                End With
                With .cmbReply1
                    .Width = siMaxReplyWidth
                    .Left = Me.cmbReply2.Left - siMaxReplyWidth - V_MARGIN ' left from center
                End With
                With .cmbReply3
                    .Width = siMaxReplyWidth
                    .Left = Me.cmbReply2.Left + siMaxReplyWidth + V_MARGIN ' Right from center
                End With
            Case 4
                With .cmbReply1
                    .Width = siMaxReplyWidth
                    .Left = V_MARGIN
                End With
                With .cmbReply2
                    .Width = siMaxReplyWidth
                    .Left = Me.cmbReply1.Left + siMaxReplyWidth + V_MARGIN ' right from center
                End With
                With .cmbReply3
                    .Width = siMaxReplyWidth
                    .Left = Me.cmbReply2.Left + siMaxReplyWidth + V_MARGIN ' left from left
                End With
                With .cmbReply4
                    .Width = siMaxReplyWidth ' right from right
                    .Left = Me.cmbReply3.Left + siMaxReplyWidth + V_MARGIN ' Right from center
                End With
            Case 5
                With .cmbReply3                                     ' Center 3rd reply button
                    .Width = siMaxReplyWidth
                    .Left = (Me.Width / 2) - (siMaxReplyWidth / 2)
                End With
                With .cmbReply2                                     ' position 2nd to left from 3rd
                    .Width = siMaxReplyWidth
                    .Left = Me.cmbReply3.Left - siMaxReplyWidth - V_MARGIN
                End With
                With .cmbReply1                                     ' position 1st to left from 2nd
                    .Width = siMaxReplyWidth
                    .Left = Me.cmbReply2.Left - siMaxReplyWidth - V_MARGIN
                End With
                With .cmbReply4                                     ' position 4th right from 3rd
                    .Width = siMaxReplyWidth
                    .Left = Me.cmbReply3.Left + siMaxReplyWidth + V_MARGIN
                End With
                With .cmbReply5                                     ' position 5th right from 4th
                    .Width = siMaxReplyWidth
                    .Left = Me.cmbReply4.Left + siMaxReplyWidth + V_MARGIN
                End With
        End Select
    End With

End Sub

Private Sub SetupMessageTexts()
    If sMsg1Proportional <> vbNullString _
    Then SetupMessageProportional LabelMsg1, sLabelMessage1, Msg1Proportional, sMsg1Proportional
    If sMsg1Fixed <> vbNullString _
    Then SetupMessageFixed LabelMsg1, sLabelMessage1, Msg1Fixed, sMsg1Fixed
    
    If sMsg2Proportional <> vbNullString _
    Then SetupMessageProportional LabelMsg2, sLabelMessage2, Msg2Proportional, sMsg2Proportional
    If sMsg2Fixed <> vbNullString _
    Then SetupMessageFixed LabelMsg2, sLabelMessage2, Msg2Fixed, sMsg2Fixed
    
    If sMsg3Proportional <> vbNullString _
    Then SetupMessageProportional LabelMsg3, sLabelMessage3, Msg3Proportional, sMsg3Proportional
    If sMsg3Fixed <> vbNullString _
    Then SetupMessageFixed LabelMsg3, sLabelMessage3, Msg3Fixed, sMsg3Fixed
End Sub

Private Sub SetupTitle()
' ----------------------------------------------------------------
' When a font name other than the system's font name is provided
' an extra title label mimics the title bar.
' In any case the title label is used to determine the form width
' by autosize of the label.
' ----------------------------------------------------------------
    
    With Me
        If sTitleFontName <> vbNullString And sTitleFontName <> .Font.Name Then
            '~~ A title with a specific font is displayed in a dedicated title label
            With .laTitle   ' Hidden by default
                .Top = TopNext(Me.laTitle)
                .Font.Name = sTitleFontName
                If sTitleFontSize <> 0 Then
                    .Font.Size = sTitleFontSize
                End If
                .Visible = True
                siTopNext = .Top + .Height + (V_MARGIN / 2)
            End With
            
        Else
            .Caption = sTitle
            .laTitleSpaceBottom.Visible = False
            With .laTitle
                '~~ The title label is used to adjust the form width
                With .Font
                    .Bold = False
                    .Name = Me.Font.Name
                    .Size = 8.7
                End With
                .Visible = False
                siTitleWidth = .Width + H_MARGIN
            End With
            siTopNext = V_MARGIN / 2
        End If
        
        With .laTitle
            '~~ The title label is used to adjust the form width
            With .Font
                .Bold = False
                .Size = 8.7
            End With
            .AutoSize = True
            .Caption = " " & sTitle    ' some left margin
            .AutoSize = False
            siTitleWidth = .Width + H_MARGIN
        End With
        
        .Width = siTitleWidth   ' not the finalwidth though
        .laTitleSpaceBottom.Width = .laTitle.Width
    
    End With

End Sub

Private Sub AdjustWidthOfVisibleProportionalText()
Dim siFormWidth As Single

    With Me
        siFormWidth = .Width
        With .tbMsg1Proportional
            If .Visible Then
                .Width = siFormWidth - (V_MARGIN * 2)
            End If
        End With
        With .tbMsg2Proportional
            If .Visible Then
                .Width = siFormWidth - (V_MARGIN * 2)
            End If
        End With
        With .tbMsg3Proportional
            If .Visible Then
                .Width = siFormWidth - (V_MARGIN * 2)
            End If
        End With
        
    End With
End Sub

Private Sub AdjustTopPosOfVisibleElements()
    SetupRepliesTopPos
End Sub

Private Sub UserForm_Activate()
    
    With Me
        
        SetupTitle
        
        SetupReplyButtons vReplies
        
        SetupMessageTexts
        
        '~~ Final adjustment of the message window's (UserForm's) width considering:
        '~~ - the title width
        '~~ - the maximum fixed message text width
        '~~ - the width and number of the displayed reply buttons
        '~~ - the specified minimum windo width
        .Width = Max( _
                      siTitleWidth, _
                     ((siMaxReplyWidth + V_MARGIN) * lReplies) + (V_MARGIN * 2), _
                     siMaxFixedTextWidth, _
                     MIN_FORM_WIDTH)
        .Height = .cmbReply1.Top + .cmbReply1.Height + (H_MARGIN * 2) ' not final yet
        
        SetupReplyButtonsHPos
        AdjustWidthOfVisibleProportionalText
        AdjustTopPosOfVisibleElements
        
    End With

End Sub
