## Common VBA Error Services
All services are invoked through procedures copied into each module (see [Preparing the module](#preparing-the-module) and [Preparing procedures](#preparing-procedures). As a result an error message will look as follows:<br>
![](Assets/StraightToTheErrorLineOptimum.png)

Provided are:
  - **Type of the error** (distinction of Application Error, VB Runtime error, and Database error)
  - **Error Number**
  - **Error Description** of the error (_err.Description_)
  - **Error source** (procedure which raised the error)
  - **Error path** (see [The path to the error](#the-path-to-the-error) for the required conditions) 
  - **Additional information about an error** (optional when concatenated to the error description with ||)
  - **A _Resume Error Line_ button** (optional when the _Conditional Compile Argument_ `Debugging = 1`)
  - **Error line** (when available)
  
  > This error services are used with all my _[Common VBA Components][7]_ in all my VB-Projects which are all prepared to function completely autonomously (download, import, use) but at the same time to integrate with my personal 'standard' VB-Project design. See [Conflicts with personal and public _Common Components_][5] for more details.

## Installation
- Download [_mErH.bas_][1] and [mMsg.bas][4] and import them to your VB-Project  
- Download [fMsg.frm][2] and [fMsg.frx][3] and import _fMsg.frm_ to your VB-Project
- Set the _Conditional Compile Argument_ `ErHComp = 1`

## Usage
For a complete demonstration of how a Workbook/VB-Project uses these services download [StraightToTheErrorLine.xlsm][8]

### Preparing the module
The following will be copied into every module which uses the services
```vb
Private Function AppErr(ByVal app_err_no As Long) As Long
' ------------------------------------------------------------------------------
' Ensures that a programmed 'Application' error number not conflicts with the
' number of a 'VB Runtime Error' or any other system error.
' - Returns a given positive 'Application Error' number (app_err_no) into a
'   negative by adding the system constant vbObjectError
' - Returns the original 'Application Error' number when called with a negative
'   error number.
' ------------------------------------------------------------------------------
    If app_err_no >= 0 Then AppErr = app_err_no + vbObjectError Else AppErr = Abs(app_err_no - vbObjectError)
End Function

Private Sub BoP(ByVal b_proc As String, ParamArray b_arguments() As Variant)
' ------------------------------------------------------------------------------
' (B)egin-(o)f-(P)rocedure named (b_proc). Procedure to be copied as Private
' into any module potentially either using the Common VBA Error Service and/or
' the Common VBA Execution Trace Service. Has no effect when Conditional Compile
' Arguments are 0 or not set at all.
' ------------------------------------------------------------------------------
    Dim s As String: If UBound(b_arguments) >= 0 Then s = Join(b_arguments, ",")
#If ErHComp = 1 Then
    mErH.BoP b_proc, s
#ElseIf ExecTrace = 1 Then
    mTrc.BoP b_proc, s
#End If
End Sub

Private Sub EoP(ByVal e_proc As String, _
       Optional ByVal e_inf As String = vbNullString)
' ------------------------------------------------------------------------------
' (E)nd-(o)f-(P)rocedure named (e_proc). Procedure to be copied as Private Sub
' into any module potentially either using the Common VBA Error Service and/or
' the Common VBA Execution Trace Service. Has no effect when Conditional Compile
' Arguments are 0 or not set at all.
' ------------------------------------------------------------------------------
#If ErHComp = 1 Then
    mErH.EoP e_proc
#ElseIf ExecTrace = 1 Then
    mTrc.EoP e_proc, e_inf
#End If
End Sub

Private Function ErrMsg(ByVal err_source As String, _
               Optional ByVal err_no As Long = 0, _
               Optional ByVal err_dscrptn As String = vbNullString, _
               Optional ByVal err_line As Long = 0) As Variant
' ------------------------------------------------------------------------------
' Universal error message display service. See:
' https://warbe-maker.github.io/vba/common/2022/02/15/Personal-and-public-Common-Components.html
'
' - Displays a debugging option button when the Conditional Compile Argument
'   'Debugging = 1'
' - Displays an optional additional "About the error:" section when a string is
'   concatenated with the error message by two vertical bars (||)
' - Invokes mErH.ErrMsg when the Conditional Compile Argument ErHComp = !
' - Invokes mMsg.ErrMsg when the Conditional Compile Argument MsgComp = ! (and
'   the mErH module is not installed / MsgComp not set)
' - Displays the error message by means of VBA.MsgBox when neither of the two
'   components is installed
'
' Uses:
' - AppErr For programmed application errors (Err.Raise AppErr(n), ....) to
'          turn them into negative and in the error message back into a
'          positive number.
' - ErrSrc To provide an unambiguous procedure name by prefixing is with the
'          module name.
'
' See:
' https://github.com/warbe-maker/Common-VBA-Error-Services
'
' W. Rauschenberger Berlin, Feb 2022
' ------------------------------------------------------------------------------
#If ErHComp = 1 Then
    '~~ When Common VBA Error Services (mErH) is available in the VB-Project
    '~~ (which includes the mMsg component) the mErh.ErrMsg service is invoked.
    ErrMsg = mErH.ErrMsg(err_source, err_no, err_dscrptn, err_line): GoTo xt
#ElseIf MsgComp = 1 Then
    '~~ When (only) the Common Message Service (mMsg, fMsg) is available in the
    '~~ VB-Project, mMsg.ErrMsg is invoked for the display of the error message.
    ErrMsg = mMsg.ErrMsg(err_source, err_no, err_dscrptn, err_line): GoTo xt
#End If
    '~~ When neither of the Common Component is available in the VB-Project
    '~~ the error message is displayed by means of the VBA.MsgBox
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
    
    '~~ Consider extra information is provided with the error description
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
            If err_dscrptn Like "*DAO*" _
            Or err_dscrptn Like "*ODBC*" _
            Or err_dscrptn Like "*Oracle*" _
            Then ErrType = "Database Error " _
            Else ErrType = "VB Runtime Error "
    End Select
    
    If err_source <> vbNullString Then ErrSrc = " in: """ & err_source & """"   ' assemble ErrSrc from available information"
    If err_line <> 0 Then ErrAtLine = " at line " & err_line                    ' assemble ErrAtLine from available information
    ErrTitle = Replace(ErrType & ErrNo & ErrSrc & ErrAtLine, "  ", " ")         ' assemble ErrTitle from available information
       
    ErrText = "Error: " & vbLf & ErrDesc & vbLf & vbLf & "Source: " & vbLf & err_source & ErrAtLine
    If ErrAbout <> vbNullString Then ErrText = ErrText & vbLf & vbLf & "About: " & vbLf & ErrAbout
    
#If Debugging Then
    ErrBttns = vbYesNo
    ErrText = ErrText & vbLf & vbLf & "Debugging:" & vbLf & "Yes    = Resume Error Line" & vbLf & "No     = Terminate"
#Else
    ErrBttns = vbCritical
#End If
    ErrMsg = MsgBox(Title:=ErrTitle, Prompt:=ErrText, Buttons:=ErrBttns)

xt:
End Function

Private Function ErrSrc(ByVal sProc As String) As String
    ErrSrc = "mBasic." & sProc
End Function
```
### Preparing the procedure(s)
Procedures will follow the scheme outlined by the following examples (also see the [StraightToTheErrorLine.xlsm][8] Workbook for a demonstration. 
```vb
Private Sub TestProc()
    Const PROC = "TestProc"
    
    On Error GoTo eh
    BoP ErrSrc(PROC)
    '
    TestTestProc    ' this one will raise the error
    '
xt: EoP ErrSrc(PROC)
    Exit Sub

eh: Select Case ErrMsg(ErrSrc(PROC))
        Case vbResume:  Stop: Resume
        Case Else:      GoTo xt
    End Select
End Sub

Private Sub TestTestProc()
    Const PROC = "TestTestProc"
    
    On Error GoTo eh
    Dim wb As Workbook
    
    BoP ErrSrc(PROC)
    Debug.Print wb.Name ' will raise a VB Runtime error no 91

xt: EoP ErrSrc(PROC)
    Exit Sub

eh: Select Case ErrMsg(ErrSrc(PROC))
        Case vbResume:  Stop: Resume
        Case Else:      GoTo xt
    End Select
End Sub
```
When an error is displayed and the _Resume Error Line_ button is pressed the following two F8 key strokes end up at the error line. When it is again executed without any change the same error message will pop-up again of course. 

### Activating the _Resume Error Line_ button
Set the _Conditional Compile Argument_ `Debuggig = 1` (may of course be set to 0 when the VB-Project becomes productive)

### Regression test issues
I use the following two services for the Regression tests set up with all my _[Common VBA Components][7].

#### _Asserted_ error(s) service
Specifying those error numbers which are regarded asserted bypasses the display of the corresponding error message when the _Regression_ property is set to True.
Note: When the asserted error is a programmed _Application Error_ (for example `Err.Raise AppErr(1), ErrSrc(PROC), "error description"`) the Asserted error also is AppErr(1)

### _Regression_ test activation
When at the beginning of a series of test procedures `mErH.Regression` is set to `True` the display of _[Asserted](#asserted-error-service)_ errors will be suspended thereby supporting an 'unattended regression test (performing a series of procedures using `Debug.Assert` to assert test results).

## The path to the error
At this point it should be clear that a path to the error depends on the following:
1. Procedures with an error handling (the more procedures having one the better)
2. An unambiguous procedure name provided with `BoP/EoP` statements and with the call of the error message function
3. The use of this _Common VBA Error Services_ component

A path to the error can be provided in two ways, of which both are combined by the mErH module for a best possible result.

| Approach | Pro | Con |
| -------- | --- | --- |
| ***Bottom up***:<br>The path is assembled when the error is passed on up to the _Entry Procedure_. | This approach assembles a complete path which includes also procedures not having `BoP/EoP` statements provided the _Entry Procedure is known (has `BoP/EoP` statements).| The path will not be available when the error message is called immediately with the error raising procedure. This is the case either when the _Entry Procedure_ is unknown, i.e. no procedure with `BoP/EoP` statements had been invoked, or when the debugging option displays an extra button to _Resume the Error Line_ which means that even when the _Entry Procedure_ is known the message has to be displayed immediately to provide the choice.| 
| ***Top down***:<br>A stack is maintained with each invoked procedure by their `BoP/EoP` statements. | The path will already be complete when the error raising procedure is invoked and thus available even when the extra Debugging option button is displayed. | The completeness of the path depends on the completeness of `BoP/EoP` statements.|

> ***Conclusion***: Procedures with an error handling (those with On error Goto eh) should also have `BoP/EoP` statements - quasi as a default). Providing potential _Entry Procedures_ with `BoP/EoP` statements should be obligatory.



## Contribution, development, test, maintenance
Any contribution of any kind will be welcome. The dedicated _Common VBA Component Workbook_ **[ErH.xlsm][6]** is used for development, maintenance, and last but not least for the testing.


[1]:https://gitcdn.link/cdn/warbe-maker/Common-VBA-Error-Services/master/source/mErH.bas
[2]:https://gitcdn.link/cdn/warbe-maker/Common-VBA-Error-Services/master/source/fMsg.frm
[3]:https://gitcdn.link/cdn/warbe-maker/Common-VBA-Error-Services/master/source/fMsg.frx
[4]:https://gitcdn.link/cdn/warbe-maker/Common-VBA-Error-Services/master/source/mMsg.bas
[5]:https://warbe-maker.github.io/vba/common/2022/02/15/Personal-and-public-Common-Components.html
[6]:https://gitcdn.link/cdn/warbe-maker/Common-VBA-Error-Services/master/source/ErH.xlsm
[7]:https://warbe-maker.github.io/vba/common/2021/02/19/Common-VBA-Components.html
[8]:https://gitcdn.link/cdn/warbe-maker/Straight-to-the-error-line-demo/master/StraightToTheErrorLine.xlsm
