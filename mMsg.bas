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
Public Type tMsgSection                 ' ---------------------
       sLabel As String                 ' Structure of the
       sText As String                  ' UserForm's message
       bMonspaced As Boolean            ' area which consists
End Type                                ' of 4 message sections
Public Type tMsg                        ' Attention: 4 is a
       section(1 To 4) As tMsgSection   ' design constant!
End Type                                ' ---------------------

Public Function Box(ByVal msg_title As String, _
           Optional ByVal msg As String = vbNullString, _
           Optional ByVal msg_monospaced As Boolean = False, _
           Optional ByVal msg_buttons As Variant = vbOKOnly, _
           Optional ByVal msg_returnindex As Boolean = False, _
           Optional ByVal msg_min_width As Long = 400, _
           Optional ByVal msg_max_width As Long = 80, _
           Optional ByVal msg_max_height As Long = 70, _
           Optional ByVal msg_min_button_width = 70) As Variant
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
        .MaxFormHeightPrcntgOfScreenSize = msg_max_height ' percentage of screen size
        .MaxFormWidthPrcntgOfScreenSize = msg_max_width   ' percentage of screen size
        .MinFormWidth = msg_min_width                     ' defaults to 300 pt. the absolute minimum is 200 pt
        .MinButtonWidth = msg_min_button_width
        .MsgTitle = msg_title
        .MsgText(1) = msg
        .MsgMonoSpaced(1) = msg_monospaced
        .MsgButtons = msg_buttons
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
    If msg_returnindex Then Box = fMsg.ReplyIndex Else Box = fMsg.ReplyValue

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
                                     
Public Function Dsply(ByVal msg_title As String, _
                      ByRef msg As tMsg, _
             Optional ByVal msg_buttons As Variant = vbOKOnly, _
             Optional ByVal msg_returnindex As Boolean = False, _
             Optional ByVal msg_min_width As Long = 300, _
             Optional ByVal msg_max_width As Long = 80, _
             Optional ByVal msg_max_height As Long = 70, _
             Optional ByVal msg_min_button_width = 70) As Variant
' -------------------------------------------------------------------------------------
' Common VBA Message Display: A service using the Common VBA Message Form as an
' alternative MsgBox.
' Note: In case there is only one single string to be displayed the argument
'       msg_type will remain unused while the messag is provided via the
'       msg_strng and msg_strng_monospaced arguments instead.
'
' See: https://warbe-maker.github.io/vba/common/2020/11/17/Common-VBA-Message-Form.html
'
' W. Rauschenberger, Berlin, Nov 2020
' -------------------------------------------------------------------------------------
    Dim i As Long
    
    With fMsg
        .MaxFormHeightPrcntgOfScreenSize = msg_max_height ' percentage of screen size
        .MaxFormWidthPrcntgOfScreenSize = msg_max_width   ' percentage of screen size
        .MinFormWidth = msg_min_width                     ' defaults to 300 pt. the absolute minimum is 200 pt
        .MinButtonWidth = msg_min_button_width
        .MsgTitle = msg_title
        For i = 1 To fMsg.NoOfDesignedMsgSections
            .MsgLabel(i) = msg.section(i).sLabel
            .MsgText(i) = msg.section(i).sText
            .MsgMonoSpaced(i) = msg.section(i).bMonspaced
        Next i
        
        .MsgButtons = msg_buttons
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
    If msg_returnindex Then Dsply = fMsg.ReplyIndex Else Dsply = fMsg.ReplyValue

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

