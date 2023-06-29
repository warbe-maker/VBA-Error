Attribute VB_Name = "mCompManClient"
Option Explicit
' ----------------------------------------------------------------------
' Standard Module mCompManClient
'
' Interface between a Workbook/VB-Project and the 'Component Management'
' for providing the services:
' - Export of changed components
' - Update of outdated used 'Common Components' (by re-importing an
'   up-to-date component's Export File whereby this corresponding 'raw'
'   component is hosted in another, possibly dedicated Workbook).
' - Synchronization of a Sync-Target-Workbook with its up-to-date
' Sync-Source-Workbook
'
' W. Rauschenberger, Berlin May 2022
'
' See also Github repo:
' https://github.com/warbe-maker/Excel-VB-Components-Management-Services
' ----------------------------------------------------------------------
' CompMan's global specifications essential for CompMan clients
Public Const SRVC_UPDATE_OUTDATED   As String = "UpdateOutdatedCommonComponents"
Public Const SRVC_SYNCHRONIZE       As String = "SynchronizeVBProjects"
Public Const SRVC_EXPORT_CHANGED    As String = "ExportChangedComponents"
Public Const COMPMAN_ADDIN          As String = "CompMan.xlam"
Public Const COMPMAN_DEVLP          As String = "CompMan.xlsb"

Private Const vbResume              As Long = 6 ' return value (equates to vbYes)
Private Busy                        As Boolean ' prevent parallel execution of a service

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
                 Optional ByVal cms_hosted_common_components As String = vbNullString, _
                 Optional ByVal cms_unused As Boolean)
' ----------------------------------------------------------------------------
' Execution of the CompMan service (cms_name) preferably via the "CompMan
' Development Instance" as the servicing Workbook. Only when not available the
' "CompMan AddIn Instance" (mCompManClient.COMPMAN_ADDIN) becomes the servicing
' Workbook - which maynot be open either or open but paused.
' Note: cms_unused is for backwards compatibility only
' ----------------------------------------------------------------------------
    Const PROC = "CompManService"
    
    On Error GoTo eh
    Dim vDone       As Variant
    Dim sServicing  As String
    
    '~~ Avoid any trouble caused by DoEvents used throughout the execution of any service
    '~~ when a service is already currently busy. This may be the case when Workbook-Save
    '~~ is clicked twice.
    If Busy Then
        Debug.Print "Terminated because a previous task is still busy!"
        Exit Sub
    End If
    Busy = True
    
    sServicing = WbServicing(cms_name)
    If sServicing <> vbNullString Then
        If cms_name = mCompManClient.SRVC_SYNCHRONIZE _
        Then Application.Run sServicing & "!mCompMan." & mCompManClient.SRVC_SYNCHRONIZE, ThisWorkbook _
        Else Application.Run sServicing & "!mCompMan." & cms_name, ThisWorkbook, cms_hosted_common_components
    End If
    
xt: Busy = False
    Application.EnableEvents = True
    Exit Sub

eh: Select Case ErrMsg(ErrSrc(PROC))
        Case vbResume:  Stop: Resume
        Case Else:      GoTo xt
    End Select
End Sub

Private Function WbServicing(ByVal csa_service As String) As String
' ----------------------------------------------------------------------------
' Returns the name of the Workbook providing the requested service which may
' be a vbNullString when the service cannot neither be provided by an open
' CompMan development instance Workbook nor by an available CompMan Addin
' instance.
' Notes: - When the requested service is not "update" an available development
'          instance is given priority over an also available Addin instance.
'        - When the requested service is "update" and the serviced Workbook
'          is the development instance the service is only available when the
'          Addin instance is avaialble.
'        - Even when a servicing Workbook (the Addin and or the development
'          instance is available, CompMan may still not be configured
'          correctly!
' Uses: mCompMan.RunTest
' ----------------------------------------------------------------------------
    Const PROC              As String = "WbServicing"
    
    Dim Result              As Long
    Dim ResultByAddin       As Long
    Dim ResultByDev         As Long
    Dim AvailableByAddin    As Boolean
    Dim AvailableByDev      As Boolean
    
    '~~ 1. Check the availability of servicing Workbooks
    On Error Resume Next
    ResultByAddin = Application.Run(COMPMAN_ADDIN & "!mCompMan.RunTest", csa_service, ThisWorkbook)
    AvailableByAddin = Err.Number = 0
    
    If AvailableByAddin And ResultByAddin <> AppErr(1) And ResultByAddin <> AppErr(2) Then
        '~~ Only when CompMan configured correctly and complete for the requested service (not AppErr(1))
        '~~ and the serviced Workbook has been opened from the service-obligatory folder (not AppErr(2))
        '~~ another try with the development instance makes sense
        On Error Resume Next
        ResultByDev = Application.Run(COMPMAN_DEVLP & "!mCompMan.RunTest", csa_service, ThisWorkbook)
        AvailableByDev = Err.Number = 0
        
        On Error GoTo eh
        If Not csa_service = mCompManClient.SRVC_UPDATE_OUTDATED _
           And Not csa_service = mCompManClient.SRVC_SYNCHRONIZE Then
            '~~ When the requested service is neither update nor synchronize and the CompMan development
            '~~ instance is available it is given priority over a possibly also available CompMan Addin instance.
            Select Case True
                Case AvailableByDev
                    WbServicing = mCompManClient.COMPMAN_DEVLP
                    Result = ResultByDev
                Case Not AvailableByDev And AvailableByAddin
                    WbServicing = mCompManClient.COMPMAN_ADDIN
                    Result = ResultByAddin
            End Select
        Else
            '~~ When the requested service is either update or synchronize and the serviced Workbook
            '~~ is the CompMan development instance the service is only available when the
            '~~ Addin instance is avaialble.
            Select Case True
                Case AvailableByAddin And ThisWorkbook.Name = mCompManClient.COMPMAN_DEVLP
                    WbServicing = mCompManClient.COMPMAN_ADDIN
                    Result = ResultByAddin
                Case AvailableByDev And ThisWorkbook.Name <> mCompManClient.COMPMAN_DEVLP
                    WbServicing = mCompManClient.COMPMAN_DEVLP
                    Result = ResultByDev
                Case AvailableByAddin
                    WbServicing = mCompManClient.COMPMAN_ADDIN
                    Result = ResultByAddin
                Case Else
                    DsplyStatus "Update sercvice not available by Addin! (" & mCompManClient.COMPMAN_DEVLP & " cannot update its own components)"
           End Select
        End If
    End If
    
    If WbServicing <> vbNullString Then
        '~~ When a servicing Workbook is available its result from RunTest must not be any of the following
        '~~ Application error
        Select Case Result
            Case AppErr(1), AppErr(2)
                WbServicing = vbNullString
            Case AppErr(3)
                DsplyStatus csa_service & " ( by " & WbServicing & ") for " & ThisWorkbook.Name & ": " & _
                                        "Denied! (the corresponding 'Sync-Source-Workbook' has not been found in CompMan's 'Serviced-Folder'!"
                WbServicing = vbNullString
        End Select
    End If

xt: Exit Function

eh: Select Case ErrMsg(ErrSrc(PROC))
        Case vbResume:  Stop: Resume
        Case Else:      GoTo xt
    End Select
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

Private Sub DsplyStatus(ByVal s As String)
    With Application
        .StatusBar = vbNullString
        .StatusBar = s
    End With
End Sub

