VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "wbErrHndlrTest"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = True
Option Explicit
' ------------------------------------------------------------------------------
' Workbook ErH.xlsm: This is the Common Component Workbook for the (hosted)
' ================== Common Component (Standard module) mErH. The hosted
' Component is registered by means of the CompMan Addin *) when the Workbook is
' closed (procedure ExportChangedComponents). Any used Common Components
' (hosted in other Workbooks) are updated when their code had changed by means of
' the SRVC_UPDATE_OUTDATED service.
'
' Uses:
' -----
' fMsg/mMsg for the display of error messages (test only)
' mTrc      for the execution trace of the test procedures (test docu only)
'
' W. Rauschenberger, Berlin, June 2023
' https://github.com/warbe-maker/VBA-Error
' ------------------------------------------------------------------------------
Const HOSTED_RAWS = "mErH"

Private Sub Workbook_Open()
    mCompManClient.CompManService mCompManClient.SRVC_UPDATE_OUTDATED, HOSTED_RAWS
End Sub

Private Sub Workbook_BeforeSave(ByVal SaveAsUI As Boolean, Cancel As Boolean)
    mCompManClient.CompManService mCompManClient.SRVC_EXPORT_CHANGED, HOSTED_RAWS
End Sub

