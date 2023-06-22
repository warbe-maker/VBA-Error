Attribute VB_Name = "mCompManClient"
Option Explicit
' ----------------------------------------------------------------------
' Standard Module mCompManClient
' ==============================
' CompMan client interface. To be imported into any Workbook for being
' serviced by CompMan's: - "Export Changed Components"
'                        - "Update Outdated Common Components"
'                        - "Synchronize VB-Projects"
'
' W. Rauschenberger, Berlin Apr 2023
'
' See https://github.com/warbe-maker/Excel-VB-Components-Management-Services
' ----------------------------------------------------------------------
' --- The below constants must not be changed to Private since they are used byCompMan
Public Const COMPMAN_DEVLP              As String = "CompMan.xlsb"
Public Const SRVC_EXPORT_ALL            As String = "ExportAll"
Public Const SRVC_EXPORT_ALL_DSPLY      As String = "Export All Components"
Public Const SRVC_EXPORT_CHANGED        As String = "ExportChangedComponents"
Public Const SRVC_EXPORT_CHANGED_DSPLY  As String = "Export Changed Components"
Public Const SRVC_SYNCHRONIZE           As String = "SynchronizeVBProjects"
Public Const SRVC_SYNCHRONIZE_DSPLY     As String = "Synchronize VB-Projects"
Public Const SRVC_UPDATE_OUTDATED       As String = "UpdateOutdatedCommonComponents"
Public Const SRVC_UPDATE_OUTDATED_DSPLY As String = "Update Outdated Common Components"
' --- The above constants must not be changed to Private since they are used byCompMan

Private Const COMPMAN_ADDIN             As String = "CompMan.xlam"
Private Const vbResume                  As Long = 6 ' return value (equates to vbYes)
Private Busy                            As Boolean ' prevent parallel execution of a service

Private Property Let DisplayedServiceStatus(ByVal s As String)
    With Application
        .StatusBar = vbNullString
        .StatusBar = s
    End With
End Property

Private Property Get IsDevInstance() As Boolean
    IsDevInstance = ThisWorkbook.Name = mCompManClient.COMPMAN_DEVLP
End Property

Private Property Get IsAddinInstance() As Boolean
    IsAddinInstance = ThisWorkbook.Name = COMPMAN_ADDIN
End Property

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

Public Sub CompManService(ByVal cms_name As String, _
                 Optional ByVal cms_hosted_common_components As String = vbNullString)
' ----------------------------------------------------------------------------
' Execution of the CompMan service (cms_name) preferably via the "CompMan
' Development Instance" as the servicing Workbook. Only when not available the
' "CompMan AddIn Instance" (COMPMAN_ADDIN) becomes the servicing
' Workbook - which maynot be open either or open but paused.
' Note: cms_unused is for backwards compatibility only
' ----------------------------------------------------------------------------
    Const PROC = "CompManService"
    
    On Error GoTo eh
    Dim sWbkServicingName   As String
    
    If IsAddinInstance Then
        Application.StatusBar = "None of CompMan's services is applicable for CompMan's Add-in instance!"
        GoTo xt
    End If

    '~~ Avoid any trouble caused by DoEvents used throughout the execution of any service
    '~~ when a service is already currently busy. This may be the case when Workbook-Save
    '~~ is clicked twice.
    If Busy Then
        Debug.Print "Terminated because a previous task is still busy!"
        Exit Sub
    End If
    Busy = True
    
    sWbkServicingName = WbkServicingName(cms_name)
    If sWbkServicingName <> vbNullString Then
        If cms_name = mCompManClient.SRVC_SYNCHRONIZE _
        Then Application.Run sWbkServicingName & "!mCompMan." & mCompManClient.SRVC_SYNCHRONIZE, ThisWorkbook _
        Else Application.Run sWbkServicingName & "!mCompMan." & cms_name, ThisWorkbook, cms_hosted_common_components
    End If
    If Not ThisWorkbook.Saved Then
        Application.DisplayAlerts = False
        Application.EnableEvents = False
        ThisWorkbook.Save
        Application.DisplayAlerts = True
    End If
    
xt: Busy = False
    Application.EnableEvents = True
    Exit Sub

eh: Select Case ErrMsg(ErrSrc(PROC))
        Case vbResume:  Stop: Resume
        Case Else:      GoTo xt
    End Select
End Sub

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
    '~~ When Common VBA Error Services (mErH) is availabel in the VB-Project
    '~~ (which includes the mMsg component) the mErh.ErrMsg service is invoked.
    ErrMsg = mErH.ErrMsg(err_source, err_no, err_dscrptn, err_line)
    GoTo xt
#ElseIf MsgComp = 1 Then
    '~~ When (only) the Common Message Service (mMsg, fMsg) is available in the
    '~~ VB-Project, mMsg.ErrMsg is invoked for the display of the error message.
    ErrMsg = mMsg.ErrMsg(err_source, err_no, err_dscrptn, err_line)
    GoTo xt
#End If
    '~~ When neither the mMsg nor the mErH component is installed the error
    '~~ message is displayed by means of the VBA.MsgBox
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

Private Function WbkServicingName(ByVal csa_service As String) As String
' ----------------------------------------------------------------------------
' Returns the name of the Workbook providing the requested service which may
' be a vbNullString when the service neither can be provided by an open
' CompMan development instance Workbook nor by an available CompMan Add-in
' instance.
' Notes: - When the requested service is not "update" an available development
'          instance is given priority over an also available Add-in instance.
'        - When the requested service is "update" and the serviced Workbook
'          is the development instance the service is only available when the
'          Add-in instance is avaialble.
'        - Even when a servicing Workbook (the Add-in and or the development
'          instance is available, CompMan may still not be configured
'          correctly!
' Uses: mCompMan.RunTest
' ----------------------------------------------------------------------------
    Const PROC = "WbkServicingName"
    
    Dim ServicedByAddinResult           As Long
    Dim ServicedByWrkbkResult           As Long
    Dim ServiceAvailableByAddin         As Boolean
    Dim ServiceAvailableByCompMan       As Boolean
    Dim ResultRequiredAddinNotAvailable As Long
    Dim ResultConfigInvalid             As Long
    Dim ResultOutsideCfgFolder          As Long
    Dim ResultRequiredDevInstncNotOpen  As Long
    
    ResultConfigInvalid = AppErr(1)              ' Configuration for the service is invalid
    ResultOutsideCfgFolder = AppErr(2)           ' Outside the for the service required folder
    ResultRequiredAddinNotAvailable = AppErr(3)  ' Required Addin for DevInstance update paused or not open
    ResultRequiredDevInstncNotOpen = AppErr(4) '
    
    '~~ Availability check CompMan Add-in
    On Error Resume Next
    ServicedByAddinResult = Application.Run(COMPMAN_ADDIN & "!mCompMan.RunTest", csa_service, ThisWorkbook)
    ServiceAvailableByAddin = Err.Number = 0
    '~~ Availability check CompMan Workbook
    On Error Resume Next
    ServicedByWrkbkResult = Application.Run(COMPMAN_DEVLP & "!mCompMan.RunTest", csa_service, ThisWorkbook)
    ServiceAvailableByCompMan = Err.Number = 0
    
    Select Case True
        '~~ Display/indicate why the service cannot be provided
        Case ServicedByWrkbkResult = ResultConfigInvalid
            Select Case csa_service
                Case SRVC_SYNCHRONIZE:      DisplayedServiceStatus = vbNullString ' "'" & SRVC_SYNCHRONIZE_DSPLY & "' service denied (no Sync-Target- and or Sync-Archive-Folder configured)!"
                Case SRVC_UPDATE_OUTDATED:  DisplayedServiceStatus = "The enabled/requested '" & SRVC_UPDATE_OUTDATED_DSPLY & "' service had been denied due to an invalid or missing configuration (see Config Worksheet)!"
                Case SRVC_EXPORT_CHANGED:   DisplayedServiceStatus = "The enabled/requested'" & SRVC_EXPORT_CHANGED_DSPLY & "' service had been denied due to an invalid or missing configuration (see Config Worksheet)!"
            End Select
        Case ServicedByWrkbkResult = ResultOutsideCfgFolder
            Select Case csa_service
                Case SRVC_SYNCHRONIZE:      Debug.Print "The enabled/requested '" & SRVC_SYNCHRONIZE_DSPLY & "' service had silently been denied! (Workbook has not been opened from within the configured 'Sync-Target-Folder')"
                Case SRVC_UPDATE_OUTDATED:  Debug.Print "The enabled/requested '" & SRVC_EXPORT_CHANGED_DSPLY & "' service had silently been denied! (Workbook has not been opened from within the configured 'Dev-and-Test-Folder')"
                Case SRVC_EXPORT_CHANGED:   Debug.Print "The enabled/requested '" & SRVC_UPDATE_OUTDATED_DSPLY & "' service had silently been denied! (Workbook has not been opened from within the configured 'Dev-and-Test-Folder')"
            End Select
        Case ServicedByWrkbkResult = ResultRequiredAddinNotAvailable
            DisplayedServiceStatus = "The required Add-in is not available for the 'Update' service for the Development-Instance!"
        Case ServicedByWrkbkResult = ResultRequiredDevInstncNotOpen
            DisplayedServiceStatus = mCompManClient.COMPMAN_DEVLP & " is the Workbook reqired for the " & SRVC_SYNCHRONIZE & " but it is not open!"
        
        '~~ When neither of the above is True the servicing Workbook instance is decided
        Case IsDevInstance And csa_service = SRVC_UPDATE_OUTDATED And ServiceAvailableByAddin:  WbkServicingName = COMPMAN_ADDIN
        Case Not IsDevInstance And ServiceAvailableByCompMan:                                   WbkServicingName = COMPMAN_DEVLP
        Case Not IsDevInstance And Not ServiceAvailableByCompMan And ServiceAvailableByAddin:   WbkServicingName = COMPMAN_ADDIN
        Case Not ServiceAvailableByCompMan And ServiceAvailableByAddin:                         WbkServicingName = COMPMAN_ADDIN
        Case ServiceAvailableByCompMan And Not ServiceAvailableByAddin:                         WbkServicingName = COMPMAN_DEVLP
        Case ServiceAvailableByCompMan And ServiceAvailableByAddin:                             WbkServicingName = COMPMAN_DEVLP
        Case Else
            '~~ Silent service denial
            Debug.Print "CompMan services are not available, neither by open Workbook nor by CompMan Add-in!"
    End Select
        
xt: Exit Function

eh: Select Case ErrMsg(ErrSrc(PROC))
        Case vbResume:  Stop: Resume
        Case Else:      GoTo xt
    End Select
End Function

