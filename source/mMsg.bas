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
Public Type TypeMsgLabel
        FontBold As Boolean
        FontColor As XlRgbColor
        FontItalic As Boolean
        FontName As String
        FontSize As Long
        FontUnderline As Boolean
        Monospaced As Boolean ' overwrites any FontName
        Text As String
End Type
Public Type TypeMsgText
        FontBold As Boolean
        FontColor As XlRgbColor
        FontItalic As Boolean
        FontName As String
        FontSize As Long
        FontUnderline As Boolean
        Monospaced As Boolean ' overwrites any FontName
        Text As String
End Type
Public Type TypeMsgSection
       Label As TypeMsgLabel
       Text As TypeMsgText
End Type
Public Type TypeMsg
    Section(1 To 4) As TypeMsgSection
End Type

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
'
' See: https://warbe-maker.github.io/vba/common/2020/11/17/Common-VBA-Message-Form.html
'
' W. Rauschenberger, Berlin, Nov 2020
' -------------------------------------------------------------------------------------
    Dim i As Long
    
    With fMsg
        .MaxFormHeightPrcntgOfScreenSize = box_max_height ' percentage of screen size
        .MaxFormWidthPrcntgOfScreenSize = box_max_width   ' percentage of screen size
        .MinFormWidth = box_min_width                     ' defaults to 300 pt. the absolute minimum is 200 pt
        .MinButtonWidth = box_min_button_width
        .MsgTitle = box_title
        .MsgText(1).Text = box_msg
        .MsgText(1).Monospaced = box_monospaced
        .MsgButtons = box_buttons
        .DefaultButton = box_button_default
        '+------------------------------------------------------------------------+
        '|| Setup prior showing the form improves the performance significantly  ||
        '|| and avoids any flickering message window with its setup.             ||
        '|| For testing purpose it may be appropriate to out-comment the Setup.  ||
        .Setup '                                                                 ||
        '+------------------------------------------------------------------------+
        .show
    End With
    
    ' -----------------------------------------------------------------------------
    ' Obtaining the reply value/index is only possible when more than one button is
    ' displayed! When the user had a choice the form is hidden when the button is
    ' pressed and the UserForm is unloade when the return value/index (either of
    ' the two) is obtained!
    ' -----------------------------------------------------------------------------
    If box_returnindex Then Box = fMsg.ReplyIndex Else Box = fMsg.ReplyValue

End Function

Public Function Buttons(ParamArray msg_buttons() As Variant) As Collection
' --------------------------------------------------------------------------
' Returns a collection of the items provided by msg_buttons. When more
' than 7 items are provided the function adds a button row break.
' --------------------------------------------------------------------------
    
    Dim cll As New Collection
    Dim i   As Long
    Dim j   As Long         ' buttons in a row counter
    Dim k   As Long: k = 1  ' button rows counter
    Dim l   As Long         ' total buttons count
    
    On Error Resume Next
    i = LBound(msg_buttons)
    If Err.Number <> 0 Then GoTo xt
    For i = LBound(msg_buttons) To UBound(msg_buttons)
        If (k = 7 And j = 7) Or l = 49 Then GoTo xt
        Select Case msg_buttons(i)
            Case vbLf, vbCrLf, vbCr
                cll.Add msg_buttons(i)
                j = 0
                k = k + 1
            Case vbOKOnly, vbOKCancel, vbAbortRetryIgnore, vbYesNoCancel, vbYesNo, vbRetryCancel
                If j = 7 Then
                    cll.Add vbLf
                    j = 0
                    k = k + 1
                End If
                cll.Add msg_buttons(i)
                j = j + 1
                l = l + 1
            Case Else
                If TypeName(msg_buttons(i)) = "String" Then
                    ' Any invalid buttons value will be ignored without notice
                    If j = 7 Then
                        cll.Add vbLf
                        j = 0
                        k = k + 1
                    End If
                    cll.Add msg_buttons(i)
                    j = j + 1
                    l = l + 1
                End If
        End Select
    Next i
    
xt: Set Buttons = cll

End Function
                                     
Public Function Dsply(ByVal dsply_title As String, _
                      ByRef dsply_msg As TypeMsg, _
             Optional ByVal dsply_buttons As Variant = vbOKOnly, _
             Optional ByVal dsply_button_default = 1, _
             Optional ByVal dsply_returnindex As Boolean = False, _
             Optional ByVal dsply_min_width As Long = 300, _
             Optional ByVal dsply_max_width As Long = 80, _
             Optional ByVal dsply_max_height As Long = 70, _
             Optional ByVal dsply_min_button_width = 70) As Variant
' -------------------------------------------------------------------------------------
' Common VBA Message Display: A service using the Common VBA Message Form as an
' alternative MsgBox.
' Note: In case there is only one single string to be displayed the argument
'       dsply_type will remain unused while the messag is provided via the
'       dsply_strng and dsply_strng_monospaced arguments instead.
'
' See: https://warbe-maker.github.io/vba/common/2020/11/17/Common-VBA-Message-Form.html
'
' W. Rauschenberger, Berlin, Nov 2020
' -------------------------------------------------------------------------------------
    Dim i As Long
    
    With fMsg
        .MaxFormHeightPrcntgOfScreenSize = dsply_max_height ' percentage of screen size
        .MaxFormWidthPrcntgOfScreenSize = dsply_max_width   ' percentage of screen size
        .MinFormWidth = dsply_min_width                     ' defaults to 300 pt. the absolute minimum is 200 pt
        .MinButtonWidth = dsply_min_button_width
        .MsgTitle = dsply_title
        For i = 1 To fMsg.NoOfDesignedMsgSections
            '~~ Save the label and the text udt into a Dictionary by transfering it into an array
            .MsgLabel(i) = dsply_msg.Section(i).Label
            .MsgText(i) = dsply_msg.Section(i).Text
        Next i
        
        .MsgButtons = dsply_buttons
        .DefaultButton = dsply_button_default
        '+------------------------------------------------------------------------+
        '|| Setup prior showing the form improves the performance significantly  ||
        '|| and avoids any flickering message window with its setup.             ||
        '|| For testing purpose it may be appropriate to out-comment the Setup.  ||
        .Setup '                                                                 ||
        '+------------------------------------------------------------------------+
        .show
    End With
    
    ' -----------------------------------------------------------------------------
    ' Obtaining the reply value/index is only possible when more than one button is
    ' displayed! When the user had a choice the form is hidden when the button is
    ' pressed and the UserForm is unloade when the return value/index (either of
    ' the two) is obtained!
    ' -----------------------------------------------------------------------------
    If dsply_returnindex Then Dsply = fMsg.ReplyIndex Else Dsply = fMsg.ReplyValue

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
    Optional ByVal err_dscrptn As String = vbNullString)
' ------------------------------------------------------
' This Common Component does not have its own error
' handling. Instead it passes on any error to the
' caller's error handling.
' ------------------------------------------------------
    
    If err_no = 0 Then err_no = Err.Number
    If err_dscrptn = vbNullString Then err_dscrptn = Err.Description

    Err.Raise Number:=err_no, Source:=err_source, Description:=err_dscrptn

End Sub

