Attribute VB_Name = "mCompManClient"
Option Explicit
' ----------------------------------------------------------------------
' Standard Module mCompManClient
' Interface between a Workbook/VB-Project and the 'Component Management'
' for: - 'Export of changed components'
'      - Update of outdated used 'Common Components' by re-importing an
'        up-to-date component's Export File whereby this corresponding
'        'raw' component is hosted in another, possibly dedicated
'        Workbook.
'
' W. Rauschenberger, Berlin May 2022
'
' See also Github repo:
' https://github.com/warbe-maker/Excel-VB-Components-Management-Services
' ----------------------------------------------------------------------
Const COMPMAN_ADDIN = "CompMan.xlam"
Const COMPMAN_DEVLP = "CompMan.xlsb"

Dim Busy        As Boolean ' prevent parallel execution of a service
Dim WbServicing As String

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

Public Sub CompManService(ByVal cms_service_name As String, _
                          ByVal cms_hosted_common_components As String, _
                          Optional ByVal cms_modeless As Boolean = True)
' ----------------------------------------------------------------------------
' Execution of the CompMan service (cms_service_name) preferably via the
' "CompMan Development Instance" as the servicing Workbook. Only when not
' available/open the "CompMan AddIn Instance" (CompMan.xlam) becomes
' the servicing Workbook - which maynot be open either or open but paused.
' When the service is requested "component-by-component" (cms_modeless = True)
' - which is only relevant for the update service - the update of outdated
' components is performed item-by-item via a modeless displayed message.
' ----------------------------------------------------------------------------
    Const PROC = "CompManService"
    
    On Error GoTo eh
    Dim vDone As Variant
    
    If Busy Then
        '~~ This should avaoid any trouble caused by DoEvents used throughout the execution of the service.
        '~~ When the service is already busy and the Save icon is immedately clicked again the service
        '~~ may run twice at the same time and may frak out.
        Debug.Print "Terminated because a previous task is still busy!"
        Exit Sub
    End If
    Busy = True
    
    If CompManServiceAvailable(cms_service_name) _
    Then Application.Run WbServicing & "!mCompMan." & cms_service_name, ThisWorkbook, cms_hosted_common_components, cms_modeless

xt: Busy = False
    Exit Sub

eh: Select Case ErrMsg(ErrSrc(PROC))
        Case vbResume:  Stop: Resume
        Case Else:      GoTo xt
    End Select
End Sub

Private Function CompManServiceAvailable(ByVal csa_service As String) As Boolean
' ----------------------------------------------------------------------------
' Returns TRUE and the servicing Workbook/component (csa_servicing_wb_comp)
' when the service (csa_service) is available for "ThisWorkbook". Because the
' CompManClient does not have all the required information the check is
' forwarded to the "RunTest" service of the potentially servicing Workbook
' which is preferrably the development instance (when available/open) and
' second the Addin instance when available (open) and not paused.
' ----------------------------------------------------------------------------
    
    Dim Result              As Long
    Dim ResultByAddin       As Long
    Dim ResultByDev         As Long
    Dim AvailableByAddin    As Boolean
    Dim AvailableByDev      As Boolean
    
    '~~ 1. Check the availability of servicing Workbooks
    On Error Resume Next
    ResultByAddin = Application.Run(COMPMAN_ADDIN & "!mCompMan.RunTest", csa_service, ThisWorkbook)
    AvailableByAddin = Err.Number = 0
    
    On Error Resume Next
    ResultByDev = Application.Run(COMPMAN_DEVLP & "!mCompMan.RunTest", csa_service, ThisWorkbook)
    AvailableByDev = Err.Number = 0
    
    Select Case True
        Case AvailableByDev = True And Not csa_service Like "Update*"
            WbServicing = COMPMAN_DEVLP ' Use of available dev instance is given priority
            Result = ResultByDev
            CompManServiceAvailable = True
        Case AvailableByAddin = False And csa_service Like "Update*"
            WbServicing = vbNullString
            GoTo xt
        Case AvailableByAddin = True And AvailableByDev = False
            WbServicing = COMPMAN_ADDIN
            Result = ResultByAddin
            CompManServiceAvailable = True
        Case AvailableByAddin = True And AvailableByDev = True And csa_service Like "Update*"
            WbServicing = COMPMAN_ADDIN
            Result = ResultByAddin
            CompManServiceAvailable = True
    End Select
     
    '~~ 2. Check if the available servicing Workbook is able to provide the requested service
    Select Case Result
        Case AppErr(1)
            Application.StatusBar = "The configuration of Compman is invalid!"
        Case AppErr(2)  ' The serviced Workbook is located outside the serviced folder (silent service denial)
        Case AppErr(3)
            If WbServicing = COMPMAN_DEVLP Then
                Application.StatusBar = "The servicing 'CompMan Addin Instance' is currently paused!"
            End If
    End Select

xt: Exit Function

End Function

Private Function ErrMsg(ByVal err_source As String, _
               Optional ByVal err_no As Long = 0, _
               Optional ByVal err_dscrptn As String = vbNullString, _
               Optional ByVal err_line As Long = 0) As Variant
' ------------------------------------------------------------------------------
' Universal error message display service including a debugging option active
' when the Conditional Compile Argument 'Debugging = 1' and an optional
' additional "About the error:" section displaying text connected to an error
' message by two vertical bars (||).
'
' A copy of this function is used in each procedure with an error handling
' (On error Goto eh).
'
' The function considers the Common VBA Error Handling Component (ErH) which
' may be installed (Conditional Compile Argument 'ErHComp = 1') and/or the
' Common VBA Message Display Component (mMsg) installed (Conditional Compile
' Argument 'MsgComp = 1'). Only when none of the two is installed the error
' message is displayed by means of the VBA.MsgBox.
'
' Usage: Example with the Conditional Compile Argument 'Debugging = 1'
'
'        Private/Public <procedure-name>
'            Const PROC = "<procedure-name>"
'
'            On Error Goto eh
'            ....
'        xt: Exit Sub/Function/Property
'
'        eh: Select Case ErrMsg(ErrSrc(PROC))
'               Case vbResume:  Stop: Resume
'               Case Else:      GoTo xt
'            End Select
'        End Sub/Function/Property
'
'        The above may appear a lot of code lines but will be a godsend in case
'        of an error!
'
' Uses:  - For programmed application errors (Err.Raise AppErr(n), ....) the
'          function AppErr will be used which turns the positive number into a
'          negative one. The error message will regard a negative error number
'          as an 'Application Error' and will use AppErr to turn it back for
'          the message into its original positive number. Together with the
'          ErrSrc there will be no need to maintain numerous different error
'          numbers for a VB-Project.
'        - The caller provides the source of the error through the module
'          specific function ErrSrc(PROC) which adds the module name to the
'          procedure name.
'
' W. Rauschenberger Berlin, Nov 2021
' ------------------------------------------------------------------------------
#If ErHComp = 1 Then
    '~~ ------------------------------------------------------------------------
    '~~ When the Common VBA Error Handling Component (mErH) is installed in the
    '~~ VB-Project (which includes the mMsg component) the mErh.ErrMsg service
    '~~ is preferred since it provides some enhanced features like a path to the
    '~~ error.
    '~~ ------------------------------------------------------------------------
    ErrMsg = mErH.ErrMsg(err_source, err_no, err_dscrptn, err_line)
    GoTo xt
#ElseIf MsgComp = 1 Then
    '~~ ------------------------------------------------------------------------
    '~~ When only the Common Message Services Component (mMsg) is installed but
    '~~ not the mErH component the mMsg.ErrMsg service is preferred since it
    '~~ provides an enhanced layout and other features.
    '~~ ------------------------------------------------------------------------
    ErrMsg = mMsg.ErrMsg(err_source, err_no, err_dscrptn, err_line)
    GoTo xt
#End If
    '~~ -------------------------------------------------------------------
    '~~ When neither the mMsg nor the mErH component is installed the error
    '~~ message is displayed by means of the VBA.MsgBox
    '~~ -------------------------------------------------------------------
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
    ErrBttns = vbYesNo
    ErrText = ErrText & vbLf & vbLf & _
              "Debugging:" & vbLf & _
              "Yes    = Resume Error Line" & vbLf & _
              "No     = Terminate"
#Else
    ErrBttns = vbCritical
#End If
    
    ErrMsg = MsgBox(Title:=ErrTitle _
                  , Prompt:=ErrText _
                  , Buttons:=ErrBttns)
xt: Exit Function

End Function

Private Function ErrSrc(ByVal sProc As String) As String
    ErrSrc = "mCompManClient." & sProc
End Function

Private Function IsString(ByVal v As Variant, _
                 Optional ByVal vbnullstring_is_a_string = False) As Boolean
' ----------------------------------------------------------------------------
' Returns TRUE when v is neither an object nor numeric.
' ----------------------------------------------------------------------------
    Dim s As String
    On Error Resume Next
    s = v
    If Err.Number = 0 Then
        If Not IsNumeric(v) Then
            If (s = vbNullString And vbnullstring_is_a_string) _
            Or s <> vbNullString _
            Then IsString = True
        End If
    End If
End Function
