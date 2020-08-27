Attribute VB_Name = "mErrHndlr"
Option Explicit
#Const ErrMsg = "Custom"    ' System = Error displayed by MsgBox,
                            ' Custom = Error displayed by fMsgFrm which is
                            '          without the message box's limitations in size
                            '          and with automated adjustment in width and height
' --------------------------------------------------------------------------------------
' Standard  Module mErrHndlr
'           Global error handling for any VBA Project.
'           - When a call stack is maintained - at least by BoP/EoP statements in the
'             entry procedure (subsequent BoP/EoP are helpfull to maintain an execution
'             trace only):
'             - The full error path, i.e. from the error causing procedure up to the
'               entry procedure is displayed along with the error message
'             - Error number and description are passed on from the error causing
'               procedure up to the entry procedure - an advantage specifically for
'               an unatended regression test as follows:
'
'               BoP ErrSrc(PROC)
'               On Error Resume Next
'               <tested procedure>
'               Debug.Assert Err.Number = n or in case the error is a programmed
'               application error: Debug.Asser AppErr(Err.Number) = n
'               EoP ErrSrc(PROC)
'
'           - With the Conditional Compile Argument "ExecTrace = 1", an execution
'             trace is displayed in the imediate window - whereby the extent depends
'             on the use of BoP/EoP and BoT/EoT statements - with the highest possible
'             precision.
'           - The local Conditional Compile Constant 'ErrMsg = "Custom"' allows the use
'             of the dedicate UserForm lErrMsg which provideds a better readability.
'
' Methods:
' - ErrHndlr Either passes on the error to the caller or when the entry procedure is
'            reached, displays the error with a complete path from the entry procedure
'            to the procedure with the error.
' - BoP      Maintains the call stack at the Begin of a Procedure (optional when using
'            this common error handler)
' - EoP      Maintains the call stack at the End of a Procedure, triggers the display
'            of the Execution Trace when the entry procedure is finished and the
'            Conditional Compile Argument ExecTrace = 1
' - BoT      Begin of Trace. In contrast to BoP this is for any group of code lines
'            within a procedure
' - EoT      End of trace corresponding with the BoT.
' - ErrMsg   Displays the error message in a proper formated manner
'                           The local Conditional Compile Constant 'ErrMsg = "Custom"'
'                           allows the use of the dedicate UserForm fErrMsg which
'                           provideds a significant better readability.
'                           ErrMsg may be used with or without a call stack.
'
' Usage:                    Private/Public Sub/Function any()
'                           Const PROC = "any"  ' procedure's name as error source
'
'                              On Error GoTo on_error
'                              BoP ErrSrc(PROC)   ' puts the procedure on the call stack
'
'                              ' <any code>
'
' exit_proc:
'                               ' <any "finally" code like re-protecting an unprotected sheet for instance>
'                               EoP ErrSrc(PROC)   ' takes the procedure off from the call stack
'                               Exit Sub/Function
'
' on_error:
'                            #If Debugging = 1 Then
'                                Stop: Resume    ' allows to exactly locate the line where the error occurs
'                            #End If
'
' Note: When the call stack is not maintained the ErrHndlr will display the message
'       immediately with the procedure the error occours. When the call stack is
'       maintained, the error message will display the call path to the error beginning
'       with the first (entry) procedure in which the call stack is maintained all the
'       call sequence down to the procedure where the error occoured.
'
' Uses: mBasic - regarded available through the CompMan Addin. When the Addin is not used
'                the module must be imported into this project
'
' Requires: Reference to "Microsoft Scripting Runtime"
'
' W. Rauschenberger, Berlin January 2020
' -----------------------------------------------------------------------------
' ~~ Begin of Declarations for withdrawing the title bar ------------------------------------
'Private Declare PtrSafe Function GetForegroundWindow Lib "User32.dll" () As LongPtr
'Private Declare PtrSafe Function GetWindowLong Lib "User32.dll" _
'                          Alias "GetWindowLongPtrA" (ByVal hwnd As LongPtr, _
'                                                     ByVal nIndex As Long) As LongPtr
'Private Declare PtrSafe Function SetWindowLong Lib "User32.dll" _
'                          Alias "SetWindowLongPtrA" (ByVal hwnd As LongPtr, _
'                                                     ByVal nIndex As Long, _
'                                                     ByVal dwNewLong As LongPtr) As LongPtr
'Private Declare PtrSafe Function DrawMenuBar Lib "User32.dll" (ByVal hwnd As LongPtr) As Long
'Private Const GWL_STYLE  As Long = (-16)
'Private Const WS_CAPTION As Long = &HC00000
' ~~ End of Declarations for withdrawing the title bar --------------------------------------
Public CallStack    As clsCallStack
Public dicTrace     As Dictionary       ' Procedure execution trancing records
Private cllErrPath  As Collection

Public Sub BoP(ByVal sErrSource As String)
' ---------------------------------------------
' Begin of Procedure. Maintains the call stack.
' ---------------------------------------------
    If CallStack Is Nothing Then
        Set CallStack = New clsCallStack
    ElseIf CallStack.StackIsEmpty Then
        Set CallStack = Nothing
        Set CallStack = New clsCallStack
    End If
    CallStack.TraceBegin sErrSource   ' implicite start of the procedure's exec trace
    CallStack.StackPush sErrSource

End Sub

Public Sub BoT(ByVal s As String)
' ---------------------------------------
' Explicit execution trace start for (s).
' ---------------------------------------
#If ExecTrace Then
    CallStack.TraceBegin s
#End If
End Sub

Public Sub EoP(ByVal sErrSource As String)
' -------------------------------------------
' End of Procedure. Maintains the call stack.
' -------------------------------------------
    If Not CallStack Is Nothing Then
        CallStack.TraceEnd sErrSource
        CallStack.StackPop sErrSource

        If CallStack.StackIsEmpty Then
            If CallStack.ErrorPath = vbNullString Then
                Set CallStack = Nothing
            End If
        End If
    End If
End Sub

Public Sub EoT(ByVal s As String)
' -------------------------------------
' Explicit execution trace end for (s).
' -------------------------------------
    CallStack.TraceEnd s
End Sub

Public Sub ErrHndlr(ByVal lErrNo As Long, _
                    ByVal sErrSource As String, _
                    ByVal sErrText As String, _
                    ByVal sErrLine As String)
' -----------------------------------------------
' When the caller (sErrSource) is the entry
' procedure the error is displayed with the path
' to the error. Otherwise the error is raised
' again to pass it on to the calling procedure.
' The .ErrorPath string is maintained with all
' the way up to the calling procedure.
' -----------------------------------------------
Const PROC      As String = "ErrHndlr"
Static sLine    As String   ' provided error line (if any) for the the finally displayed message
   
   
    If lErrNo = 0 Then
        MsgBox "Apparently an ""Exit ..."" statement before the error handling is missing! The error handling has been aproached with a 0 error number!", vbExclamation, _
               "Problem deteced with " & ErrSrc(PROC)
        Exit Sub
    End If
    
    If CallStack Is Nothing Then Set CallStack = New clsCallStack
    If sErrLine <> 0 Then sLine = sErrLine
    
    With CallStack
        If .ErrorSource = vbNullString Then
            '~~ When the ErrorSource property is still empty, this indicates that the
            '~~ error handler is executed the first time This is the error raising procedure. Backtracking to the entry procedure is due
            Set cllErrPath = Nothing: Set cllErrPath = New Collection
            .ErrorSource = sErrSource
            .SourceErrorNo = lErrNo
            .ErrorNumber = lErrNo
            .ErrorDescription = sErrText
            .ErrorPath = .ErrorPath & sErrSource & " (" & ErrorDetails(lErrNo, sErrLine) & ")" & vbLf
            .TraceError sErrSource & ": " & ErrorDetails(lErrNo, sErrLine) & " """ & sErrText & """"
        ElseIf .ErrorNumber <> lErrNo Then
            '~~ The error number had changed during the process
            '~~ of passing the error on to the entry procedure
            .ErrorPath = .ErrorPath & sErrSource & " (" & ErrorDetails(lErrNo, sErrLine) & ")" & vbLf
            .TraceError sErrSource & ": " & ErrorDetails(lErrNo, sErrLine) & " """ & sErrText & """"
            .ErrorNumber = lErrNo
        Else
            '~~ This is the error handling called during the "backtracing" process,
            '~~ i.e. the process when the error is passed on up to the entry procedure
            .ErrorPath = .ErrorPath & sErrSource & vbLf
        End If
        
        If .EntryProc <> sErrSource Then ' And Not .ErrorPath <> vbNullString Then
            '~~ This is the call of the error handling for the error causing procedure or
            '~~ any of the procedures up to the entry procedure which has yet not been reached.
            '~~ The "backtrace" error path is maintained ....
            cllErrPath.Add sErrSource
            '~~ ... and the error is passed on to the calling procedure.
            Err.Raise lErrNo, sErrSource, sErrText
        
        ElseIf .EntryProc = sErrSource Then
            '~~ The entry procedure has been reached
            '~~ The "backtrace" error path is maintained ....
            cllErrPath.Add sErrSource
            '~~ .. and the error is displayed
            ErrMsg .SourceErrorNo, .ErrorSource, .ErrorDescription, sLine
            
#If ExecTrace Then
            '~~ Display of the full execution trace which had been maintained by
            '~~ the BoP and EoP and the BoT and EoT statements executed
            DsplyTrace
#End If
        End If
    End With

End Sub

Public Sub ErrMsg(ByVal lErrNo As Long, _
                  ByVal sErrSrc As String, _
                  ByVal sErrDesc As String, _
                  ByVal sErrLine As String)
' -------------------------------------------
' Displays the error message either by means
' of MsgBox or, when the Conditional Compile
' Argument ErrMsg = "Custom", by means of the
' Common Component fMsg. In any case the path
' to the error may be displayed, provided
' BoP/EoP statements had been executed with
' the entry procedure.
'
' W. Rauschenberger, Berlin, June 2020
' -------------------------------------------
Dim sErrMsg     As String
Dim sTitle      As String
Dim sErrPath    As String
Dim sIndicate   As String
Dim i           As Long
Dim sErrText    As String
Dim sErrInfo    As String
Dim iIndent     As Long

    '~~ Additional info about the error line in case one had been provided
    If sErrLine = vbNullString Or sErrLine = "0" Then
        sIndicate = vbNullString
    Else
        sIndicate = " (at line " & sErrLine & ")"
    End If
    sTitle = sTitle & sIndicate
        
    '~~ Path from the entry procedure (the first which uses BoP/EoP)
    '~~ all the way down to the procedure in which the error occoured.
    '~~ When the call stack had not been maintained the path is empty.
    If Not CallStack Is Nothing Then
        If Not CallStack.ErrorPath = vbNullString Then
            CallStack.TraceEndTime = Now()
            CallStack.StackUnwind
        End If
    End If
    
    For i = cllErrPath.Count To 1 Step -1
        If i = cllErrPath.Count Then
            sErrPath = cllErrPath(i) & vbLf
        ElseIf i = 1 Then
            sErrPath = sErrPath & mBasic.Space((iIndent) * 2) & "|_" & cllErrPath(i) & sIndicate
        Else
            sErrPath = sErrPath & mBasic.Space((iIndent) * 2) & "|_" & cllErrPath(i) & vbLf
        End If
        iIndent = iIndent + 1
    Next i
    '~~ Prepare the Title with the error number and the procedure which caused the error
    Select Case lErrNo
        Case Is > 0:    sTitle = "VBA Error " & lErrNo
        Case Is < 0:    sTitle = "Application Error " & AppErr(lErrNo)
    End Select
    sTitle = sTitle & " in:  " & sErrSrc & sIndicate
         
    '~~ Consider the error description may include an additional information about the error
    '~~ possible only when the error is raised by Err.Raise
    If InStr(sErrDesc, DCONCAT) <> 0 Then
        sErrText = Split(sErrDesc, DCONCAT)(0)
        sErrInfo = Split(sErrDesc, DCONCAT)(1)
    Else
        sErrText = sErrDesc
        sErrInfo = vbNullString
    End If
                       
#If ErrMsg = "Custom" Then
    '~~ Display the error message by means of the Common UserForm fMsg
    mBasic.ErrMsg lErrNo:=lErrNo, sTitle:=sTitle, sErrDesc:=sErrText, sErrPath:=sErrPath, sErrInfo:=sErrInfo
#Else
    '~~ Assemble error message to be displayed by MsgBox
    sErrMsg = "Source: " & vbLf & sErrSrc & sIndicate & vbLf & vbLf & _
              "Error: " & vbLf & sErrText
    If sErrPath <> vbNullString Then
        sErrMsg = sErrMsg & vbLf & vbLf & "Call Stack:" & vbLf & sErrPath
    End If
    If sErrInfo <> vbNullString Then
        sErrMsg = sErrMsg & vbLf & "About: " & vbLf & sErrInfo
    End If
    MsgBox sErrMsg, vbCritical, sTitle
#End If
End Sub

Private Sub DsplyTrace()
' ------------------------------------------------------------
' Displays the execution trace when the entry procedure has
' been reached.
' Note: The call stack is primarily used to detect whether or
'       not there was an initial entry procedure. It is not
'       used to maintain the error path which is done in any
'       case along with the process of passing on the error
'       to the calling procedure.
' ------------------------------------------------------------
    If CallStack Is Nothing Then
        Set CallStack = New clsCallStack
    End If
    CallStack.TraceDsply
    Set CallStack = Nothing
End Sub

Private Function ErrorDetails(ByVal lErrNo As Long, _
                              ByVal sErrLine As String) As String
' -----------------------------------------------------------------
' Returns kind of error, error number, and error line if available.
' -----------------------------------------------------------------
Dim s As String
    If lErrNo < 0 Then
        s = "App error " & AppErr(lErrNo)
    Else
        s = "VB error " & lErrNo
    End If
    If sErrLine <> 0 Then
        s = s & " at line " & sErrLine
    End If
    ErrorDetails = s
End Function

Private Function ErrSrc(ByVal sProc As String) As String
    ErrSrc = ThisWorkbook.Name & ">mErrHndlr" & ">" & sProc
End Function
