Attribute VB_Name = "mCompManClient"
Option Explicit
' ----------------------------------------------------------------------------
' Standard Module mCompManClient: CompMan client interface. To be imported
' =============================== into any Workbook for - potentially - being
' serviced by CompMan's "Export Changed Components",
'                       "Update Outdated Common Components",
'                    or "Synchronize VB-Projects" service.
'
' W. Rauschenberger, Berlin Oct 2023
'
' See https://github.com/warbe-maker/VB-Components-Management
' ----------------------------------------------------------------------------
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
Private sEventsLvl                      As String
Private bWbkExecChange                  As Boolean

' --- Begin of declarations to get all Workbooks of all running Excel instances
Private Declare PtrSafe Function FindWindowEx Lib "user32" Alias "FindWindowExA" (ByVal hWnd1 As LongPtr, ByVal hWnd2 As LongPtr, ByVal lpsz1 As String, ByVal lpsz2 As String) As LongPtr
Private Declare PtrSafe Function GetClassName Lib "user32" Alias "GetClassNameA" (ByVal hWnd As LongPtr, ByVal lpClassName As String, ByVal nMaxCount As LongPtr) As LongPtr
Private Declare PtrSafe Function IIDFromString Lib "ole32" (ByVal lpsz As LongPtr, ByRef lpiid As UUID) As LongPtr
Private Declare PtrSafe Function AccessibleObjectFromWindow Lib "oleacc" (ByVal hWnd As LongPtr, ByVal dwId As LongPtr, ByRef riid As UUID, ByRef ppvObject As Object) As LongPtr

Type UUID 'GUID
    Data1 As Long
    Data2 As Integer
    Data3 As Integer
    Data4(7) As Byte
End Type

Const IID_IDispatch As String = "{00020400-0000-0000-C000-000000000046}"
Const OBJID_NATIVEOM As LongPtr = &HFFFFFFF0
' --- End of declarations to get all Workbooks of all running Excel instances
' --- Error declarations
Const ERR_EXISTS_CMP01 = "The Component (parameter vComp) for the Component's existence check is neihter a Component object nor a string (a Component's name)!"
Const ERR_EXISTS_CVW01 = "The CustomView (parameter vCv) for the CustomView's existence check is neither a string (CustomView's name) nor a CustomView object!"
Const ERR_EXISTS_FLE01 = "The File (parameter vFile) for the File's existence check is neither a full path/file name nor a file object!"
Const ERR_EXISTS_OWB01 = "The Workbook (parameter vWb) is not open (it may have been open and already closed)!"
Const ERR_EXISTS_OWB02 = "A Workbook named '<>' is not open in any application instance!"
Const ERR_EXISTS_OWB03 = "The Workbook (parameter vWb) of which the open object is requested is ""Nothing"" (neither a Workbook object nor a Workbook's name or fullname)!"
Const ERR_EXISTS_PRC01 = "The item (parameter v) for the Procedure's existence check is neither a Component object nor a CodeModule object!"
Const ERR_EXISTS_RNG01 = "The Worksheet (parameter vWs) for the Range's existence check does not exist in Workbook (vWb)!"
Const ERR_EXISTS_RNG02 = "The Range (parameter vRange) for the Range's existence check is ""Nothing""!"
Const ERR_EXISTS_REF01 = "The Reference (parameter vRef) for the Reference's existence check is neither a valid GUID (a string enclosed in { } ) nor a Reference object!"
Const ERR_EXISTS_WBK01 = "The Workbook (parameter vWb) is neither a Workbook object nor a Workbook's name or fullname)!"
Const ERR_EXISTS_WSH01 = "The Worksheet (parameter vWs) for the Worksheet's existence check is neither a Worksheet object nor a Worksheet's name or modulename!"
Const ERR_EXISTS_GOW01 = "A Workbook (parameter vWb) named '<>' is not open!"
Const ERR_EXISTS_GOW02 = "A Workbook with the provided name (parameter vWb) is open. However it's location is '<>1' and not '<>2'!"
Const ERR_EXISTS_GOW03 = "A Workbook named '<>' (parameter vWb) is not open. A full name must be provided to get it opened!"
Const ERR_EXISTS_GOW04 = "The Workbook (parameter vWb) is a Workbook object not/no longer open!"
Const ERR_EXISTS_GOW05 = "The Workbook (parameter vWb) is neither a Workbook object nor a string (name or fullname)!"
Const ERR_EXISTS_GOW06 = "A Workbook file named '<>' (parameter vWb) does not exist!"

Public Property Get ServiceName(Optional ByVal s As String) As String
    Select Case s
        Case SRVC_EXPORT_CHANGED:   ServiceName = SRVC_EXPORT_CHANGED_DSPLY
        Case SRVC_SYNCHRONIZE:      ServiceName = SRVC_SYNCHRONIZE_DSPLY
        Case SRVC_UPDATE_OUTDATED:  ServiceName = SRVC_UPDATE_OUTDATED_DSPLY
    End Select
End Property

Private Property Let DisplayedServiceStatus(ByVal s As String)
    With Application
        .StatusBar = vbNullString
        .StatusBar = s
    End With
End Property

Private Property Get IsAddinInstance() As Boolean
    IsAddinInstance = ThisWorkbook.Name = COMPMAN_ADDIN
End Property

Private Property Get IsDevInstance() As Boolean
    IsDevInstance = ThisWorkbook.Name = mCompManClient.COMPMAN_DEVLP
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

'#If Win64 Then
    Private Function checkHwnds(ByRef xlApps() As Application, hWnd As LongPtr) As Boolean
'#Else
'    Private Function checkHwnds(ByRef xlApps() As Application, hWnd As Long) As Boolean
'#End If
' -----------------------------------------------------------------------------------------
'
' -----------------------------------------------------------------------------------------
    Const PROC = "checkHwnds"

    On Error GoTo eh
    Dim i       As Long
    
    If UBound(xlApps) = 0 Then GoTo xt

    For i = LBound(xlApps) To UBound(xlApps)
        If xlApps(i).hWnd = hWnd Then
            checkHwnds = False
            GoTo xt
        End If
    Next i

    checkHwnds = True
    
xt: Exit Function
    
eh: If ErrMsg(ErrSrc(PROC)) = vbYes Then: Stop: Resume
End Function

Public Sub CompManService(ByVal c_service_proc As String, _
                 Optional ByVal c_hosted_common_components As String = vbNullString)
' ----------------------------------------------------------------------------
' Execution of the CompMan service (c_service_proc) preferably via the "CompMan
' Development Instance" as the servicing Workbook. Only when not available the
' "CompMan AddIn Instance" (COMPMAN_ADDIN) becomes the servicing
' Workbook - which maynot be open either or open but paused.
' Note: c_unused is for backwards compatibility only
' ----------------------------------------------------------------------------
    Const PROC = "CompManService"
    
    On Error GoTo eh
    Dim sServicingWbkName   As String
        
'    If c_service_proc = mCompManClient.SRVC_EXPORT_CHANGED And ThisWorkbook.Saved Then GoTo xt
    
    Progress p_service_name:=ServiceName(c_service_proc) _
           , p_serviced_wbk_name:=ThisWorkbook.Name
    
    sEventsLvl = vbNullString
    mCompManClient.Events ErrSrc(PROC) & "." & c_service_proc, False
    If IsAddinInstance Then
        Application.StatusBar = "None of CompMan's services is applicable for CompMan's Add-in instance!"
        GoTo xt
    End If

    '~~ Avoid any trouble caused by DoEvents used throughout the execution of any service
    '~~ when a service is already currently busy. This may be the case when Workbook-Save
    '~~ is clicked twice.
    If Busy Then
        Progress p_service_name:=ServiceName(c_service_proc) _
               , p_serviced_wbk_name:=ThisWorkbook.Name _
               , p_service_info:="Terminated because a previous task is still busy!"
        Exit Sub
    End If
    Busy = True
    
    sServicingWbkName = ServicingWbkName(c_service_proc)
                                   
    If sServicingWbkName <> vbNullString Then
        Progress p_service_name:=ServiceName(c_service_proc) _
               , p_serviced_wbk_name:=ThisWorkbook.Name _
               , p_by_servicing_wbk_name:=sServicingWbkName
        If c_service_proc = mCompManClient.SRVC_SYNCHRONIZE _
        Then Application.Run sServicingWbkName & "!mCompMan." & mCompManClient.SRVC_SYNCHRONIZE, ThisWorkbook _
        Else Application.Run sServicingWbkName & "!mCompMan." & c_service_proc, ThisWorkbook, c_hosted_common_components
    Else
        Progress p_service_name:=ServiceName(c_service_proc) _
               , p_serviced_wbk_name:=ThisWorkbook.Name _
               , p_service_info:="Workbook saved (CompMan-Service not applicable)"
    End If
'    If Not ThisWorkbook.Saved Then
'        With Application
'            .DisplayAlerts = False
'            ThisWorkbook.Save
'            .DisplayAlerts = True
'        End With
'    End If
    
xt: Busy = False
    mCompManClient.Events ErrSrc(PROC) & "." & c_service_proc, True
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

Public Sub Events(ByVal e_src As String, _
                  ByVal e_b As Boolean, _
         Optional ByVal e_reset As Boolean = False)
' ------------------------------------------------------------------------------
' Follow-Up (trace) of Application.EnableEvents False/True - proves consistency.
' Recognizes the execution chang from the initiating Workbook to the service
' executing Workbook.
' ------------------------------------------------------------------------------
    Const PROC = "Events"
    
    On Error GoTo eh
    Static sLastExecWrkbk   As String
    Dim v                   As Variant
    Dim wbk                 As Workbook
    Dim dct                 As Dictionary
    
    If e_reset Then
        sEventsLvl = vbNullString
        bWbkExecChange = False
        sLastExecWrkbk = vbNullString
        Exit Sub
    End If
    
    If Not e_b Then
        EventsApp False
        Debug.Print sEventsLvl & ">> " & ThisWorkbook.Name & "." & e_src & " (Application.EnableEvents = False)"
        If sLastExecWrkbk <> vbNullString And ThisWorkbook.Name <> sLastExecWrkbk And Not bWbkExecChange Then
            bWbkExecChange = True
            sEventsLvl = sEventsLvl & "   "
        End If
        sEventsLvl = sEventsLvl & "   "
        sLastExecWrkbk = ThisWorkbook.Name
    Else
        sEventsLvl = Left(sEventsLvl, Len(sEventsLvl) - 3)
        sLastExecWrkbk = ThisWorkbook.Name
        EventsApp True
        Debug.Print sEventsLvl & "<< " & ThisWorkbook.Name & "." & e_src & " (Application.EnableEvents = True)"
    End If

    If sEventsLvl = vbNullString Then
        Events vbNullString, False, True
    End If
    
xt: Exit Sub
    
eh: If ErrMsg(ErrSrc(PROC)) = vbYes Then: Stop: Resume
End Sub

Public Sub EventsApp(a_events As Boolean)
' ------------------------------------------------------------------------------
'
' ------------------------------------------------------------------------------
    Const PROC = "EventsApp"

    On Error GoTo eh
#If Win64 Then
    Dim hWndMain As LongPtr
#Else
    Dim hWndMain As Long
#End If
    Dim appThis As Application
    Dim appNext As Application
    
    hWndMain = FindWindowEx(0&, 0&, "XLMAIN", vbNullString)
    Do While hWndMain <> 0
        Set appNext = GetExcelObjectFromHwnd(hWndMain)
        If Not appNext Is Nothing Then
            If Not appNext Is appThis Then
                Set appThis = appNext
                If appThis.EnableEvents <> a_events Then
                    appThis.EnableEvents = a_events
                End If
            End If
        End If
        hWndMain = FindWindowEx(0&, hWndMain, "XLMAIN", vbNullString)
    Loop

xt: Exit Sub
    
eh: If ErrMsg(ErrSrc(PROC)) = vbYes Then: Stop: Resume
End Sub

'#If Win64 Then
    Private Function GetExcelObjectFromHwnd(ByVal hWndMain As LongPtr) As Application
'#Else
'    Private Function GetExcelObjectFromHwnd(ByVal hWndMain As Long) As Application
'#End If
'
'#If Win64 Then
    Dim hWndDesk As LongPtr
    Dim hWnd As LongPtr
'#Else
'    Dim hWndDesk As Long
'    Dim hWnd As Long
'#End If
' -----------------------------------------------------------------------------------
'
' -----------------------------------------------------------------------------------
    Dim sText   As String
    Dim lRet    As Long
    Dim iid     As UUID
    Dim ob      As Object
    
    hWndDesk = FindWindowEx(hWndMain, 0&, "XLDESK", vbNullString)

    If hWndDesk <> 0 Then
        hWnd = FindWindowEx(hWndDesk, 0, vbNullString, vbNullString)

        Do While hWnd <> 0
            sText = String$(100, Chr$(0))
            lRet = CLng(GetClassName(hWnd, sText, 100))
            If Left$(sText, lRet) = "EXCEL7" Then
                Call IIDFromString(StrPtr(IID_IDispatch), iid)
                If AccessibleObjectFromWindow(hWnd, OBJID_NATIVEOM, iid, ob) = 0 Then 'S_OK
                    Set GetExcelObjectFromHwnd = ob.Application
                    GoTo xt
                End If
            End If
            hWnd = FindWindowEx(hWndDesk, hWnd, vbNullString, vbNullString)
        Loop
        
    End If
    
xt:
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

Public Sub Progress(ByVal p_service_name As String, _
           Optional ByVal p_serviced_wbk_name As String = vbNullString, _
           Optional ByVal p_by_servicing_wbk_name As String = vbNullString, _
           Optional ByVal p_progress_figures As Boolean = False, _
           Optional ByVal p_service_op As String = vbNullString, _
           Optional ByVal p_no_comps_serviced As Long = 0, _
           Optional ByVal p_no_comps_outdated As Long = 0, _
           Optional ByVal p_no_comps_total As Long = 0, _
           Optional ByVal p_no_comps_ignored As Long = 0, _
           Optional ByVal p_service_info As String = vbNullString)
' --------------------------------------------------------------------------
' Universal message of the export and the update service's progress in the
' form:
' <service> (by <by>) for <serviced>: <n> of <m> <op> [(component [, component] ..]
' <n> = Number of objects the service has been provided for (p_items_serviced)
' <m> = Total number of objects to be (ptentially) serviced
' <op> = The performed operation
' Whereby the progress is indicated in two ways: an increasing number of
' dots for the items collected for being serviced and a decreasing number
' of dots indication the items already serviced.
'
' Example:
' Export ... (by CompMan....) for ......: 1 of 50 exported (clsServices)
' --------------------------------------------------------------------------
    Const PROC                  As String = "Progress"
    Const SRVC_PROGRESS_SCHEME  As String = "<srvc> <by> <serviced>: <n> of <m> <dots> <op> <info>"
    
    On Error GoTo eh
    Dim sMsg    As String
    Dim lDots   As Long
    
    sMsg = Replace(SRVC_PROGRESS_SCHEME, "<srvc>", p_service_name)
    sMsg = Replace(sMsg, "<serviced>", "for " & p_serviced_wbk_name)
    If p_by_servicing_wbk_name <> vbNullString _
    Then sMsg = Replace(sMsg, "<by>", "(by " & p_by_servicing_wbk_name & ")") _
    Else sMsg = Replace(sMsg, "<by>", vbNullString)
    
    If p_progress_figures Then
        sMsg = Replace(sMsg, "<n>", p_no_comps_serviced)
        If p_no_comps_outdated <> 0 Then
            sMsg = Replace(sMsg, "<m>", p_no_comps_outdated)
        Else
            sMsg = Replace(sMsg, "<m>", p_no_comps_total)
        End If
        sMsg = Replace(sMsg, "<op>", p_service_op)
        lDots = p_no_comps_total - p_no_comps_ignored - p_no_comps_serviced
        If lDots >= 0 Then
            sMsg = Replace(sMsg, "<dots>", String(lDots, "."))
        Else
            sMsg = Replace(sMsg, "<dots>", vbNullString)
        End If
    Else
        sMsg = Replace(sMsg, "<n>", vbNullString)
        sMsg = Replace(sMsg, "of <m>", vbNullString)
        sMsg = Replace(sMsg, "<op>", vbNullString)
        sMsg = Replace(sMsg, "<dots>", vbNullString)
        sMsg = sMsg & " please wait!"
    End If
    
    sMsg = Replace(sMsg, "<info>", p_service_info)
    sMsg = Replace(sMsg, "  ", " ")
    If Len(sMsg) > 255 Then sMsg = Left(sMsg, 250) & " ..."
    With Application
        .ScreenUpdating = False
        .StatusBar = Trim(sMsg)
        .ScreenUpdating = True
    End With
    
xt: Exit Sub

eh: Select Case mBasic.ErrMsg(ErrSrc(PROC))
        Case vbResume:  Stop: Resume
        Case Else:      GoTo xt
    End Select
End Sub

Private Function ServicingWbkName(ByVal s_service_proc As String) As String
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
    Const PROC = "ServicingWbkName"
    
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
    ServicedByAddinResult = Application.Run(COMPMAN_ADDIN & "!mCompMan.RunTest", s_service_proc, ThisWorkbook)
    ServiceAvailableByAddin = Err.Number = 0
    '~~ Availability check CompMan Workbook
    On Error Resume Next
    ServicedByWrkbkResult = Application.Run(COMPMAN_DEVLP & "!mCompMan.RunTest", s_service_proc, ThisWorkbook)
    ServiceAvailableByCompMan = Err.Number = 0
    
    Select Case True
        '~~ Display/indicate why the service cannot be provided
        Case ServicedByWrkbkResult = ResultConfigInvalid
            Select Case s_service_proc
                Case SRVC_SYNCHRONIZE:      DisplayedServiceStatus = vbNullString ' "'" & SRVC_SYNCHRONIZE_DSPLY & "' service denied (no Sync-Target- and or Sync-Archive-Folder configured)!"
                Case SRVC_UPDATE_OUTDATED:  DisplayedServiceStatus = "The enabled/requested '" & SRVC_UPDATE_OUTDATED_DSPLY & "' service had been denied due to an invalid or missing configuration (see Config Worksheet)!"
                Case SRVC_EXPORT_CHANGED:   DisplayedServiceStatus = "The enabled/requested'" & SRVC_EXPORT_CHANGED_DSPLY & "' service had been denied due to an invalid or missing configuration (see Config Worksheet)!"
            End Select
        Case ServicedByWrkbkResult = ResultOutsideCfgFolder
            Progress p_service_name:=ServiceName(s_service_proc) _
                   , p_serviced_wbk_name:=ThisWorkbook.Name _
                   , p_service_info:="Service not applicable"
            Select Case s_service_proc
                Case SRVC_SYNCHRONIZE:      Debug.Print "The enabled/requested '" & SRVC_SYNCHRONIZE_DSPLY & "' service had silently been denied! (Workbook has not been opened from within the configured 'Sync-Target-Folder')"
                Case SRVC_UPDATE_OUTDATED:  Debug.Print "The enabled/requested '" & SRVC_EXPORT_CHANGED_DSPLY & "' service had silently been denied! (Workbook has not been opened from within the configured 'Dev-and-Test-Folder')"
                Case SRVC_EXPORT_CHANGED
                    Debug.Print "The enabled/requested '" & SRVC_UPDATE_OUTDATED_DSPLY & "' service had silently been denied! (Workbook has not been opened from within the configured 'Dev-and-Test-Folder')"
            End Select
        Case ServicedByWrkbkResult = ResultRequiredAddinNotAvailable
            DisplayedServiceStatus = "The required Add-in is not available for the 'Update' service for the Development-Instance!"
        Case ServicedByWrkbkResult = ResultRequiredDevInstncNotOpen
            DisplayedServiceStatus = mCompManClient.COMPMAN_DEVLP & " is the Workbook reqired for the " & SRVC_SYNCHRONIZE & " but it is not open!"
        
        '~~ When neither of the above is True the servicing Workbook instance is decided
        Case IsDevInstance And s_service_proc = SRVC_UPDATE_OUTDATED And ServiceAvailableByAddin:   ServicingWbkName = COMPMAN_ADDIN
        Case Not IsDevInstance And ServiceAvailableByCompMan:                                       ServicingWbkName = COMPMAN_DEVLP
        Case Not IsDevInstance And Not ServiceAvailableByCompMan And ServiceAvailableByAddin:       ServicingWbkName = COMPMAN_ADDIN
        Case Not ServiceAvailableByCompMan And ServiceAvailableByAddin:                             ServicingWbkName = COMPMAN_ADDIN
        Case ServiceAvailableByCompMan And Not ServiceAvailableByAddin:                             ServicingWbkName = COMPMAN_DEVLP
        Case ServiceAvailableByCompMan And ServiceAvailableByAddin:                                 ServicingWbkName = COMPMAN_DEVLP
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

