Attribute VB_Name = "mCompManClient"
Option Explicit
' ----------------------------------------------------------------------
' Standard Module mCompManClient, optionally used by any Workbook to:
' - update used 'Common-Components' (hosted, developed, tested,
'   and provided, by another Workbook) with the Workbook_open event
' - export any changed VBComponent with the Workbook_Before_Save event.
'
' W. Rauschenberger, Berlin March 2021
'
' See also Github repo:
' https://github.com/warbe-maker/Excel-VB-Components-Management-Services
' ----------------------------------------------------------------------

Public Sub CompManService(ByVal cm_service As String, _
                          ByVal hosted As String)
' ----------------------------------------------------------------------------
' Execution of the CompMan service (cm_service) preferably via the CompMan
' Development instance when available (assuming it is for testing). Only when
' not available the CompMan AddIn services (CompMan.xlam) are used.
' ----------------------------------------------------------------------------
    Const COMPMAN_BY_ADDIN = "CompMan.xlam!mCompMan."
    Const COMPMAN_BY_DEVLP = "CompMan.xlsb!mCompMan."
    
    On Error Resume Next
    Application.Run COMPMAN_BY_DEVLP & cm_service, ThisWorkbook, hosted
    If Err.Number = 1004 Then
        On Error Resume Next
        Application.Run COMPMAN_BY_ADDIN & cm_service, ThisWorkbook, hosted
        If Err.Number = 1004 Then
            Application.StatusBar = "'" & cm_service & "' neither available by '" & COMPMAN_BY_ADDIN & "' nor by '" & COMPMAN_BY_DEVLP & "'!"
        End If
    End If

xt: Exit Sub

End Sub

