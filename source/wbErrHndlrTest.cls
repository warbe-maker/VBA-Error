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
' Workbook ErH.xlsm
'
'          This is the Common Component Workbook for the (hosted) Common
'          Component (Standard module) mErH.
'          The hosted Component is registered by means of the CompMan Addin *)
'          when the Workbook is closed (procedure ExportChangedComponents).
'          Any used Common Components (hosted in other Workbooks) are updated
'          when their code had changed by means of the CompMan Addin procedure
'          UpdateUsedCommCompsTheOriginHasChanged when the Workbook is opened *)

'          *) provided the Conditional Compile Argument CompMan = 1 and the
'             CompMan Addin Workbook is available, i.e. it is referenced.
'
' Uses: fMsg Alternative VBA MsgBox
'
'          See the Github blog post
'          "A comprehensive common VBA Error Handler inspired by the best of the web"
' https://warbe-maker.github.io/vba/common/2020/10/02/Comprehensive-Common-VBA-Error-Handler.html
'
' W. Rauschenberger, Berlin, Nov 2020
' ------------------------------------------------------------------------------
Const HOSTED_RAWS = "mErH"

Private Sub Workbook_Open()
    mCompManClient.CompManService "UpdateRawClones", HOSTED_RAWS
End Sub

Private Sub Workbook_BeforeSave(ByVal SaveAsUI As Boolean, Cancel As Boolean)
    mCompManClient.CompManService "ExportChangedComponents", HOSTED_RAWS
End Sub

