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
' Workbook ErrMsg.xlsm
'
'          This is the development and test Workbook for the Common Component
'          "mErrHndlr" with its sub-components clsCallStack and clsCallStackItem.
'          Its Hosted Components are registered by means of the CompMan Addin
'          feature ExportChangedComponents when the Workbook is closed.
'          Any used Common Components (hosted in other Workbooks) are updated
'          (when their code had changed) by means of the CompMan Addin feature
'          UpdateUsedCommCompsTheOriginHasChanged when the Workbook is opened.
'
'          See the comments in module mErrHndlr for details about this common
'          component
'
' W. Rauschenberger, Berlin, June 2020
' ------------------------------------------------------------------------------
Const HOSTED_COMMON_COMPS = "mErrHndlr, clsCallStack, clsCallStackItem"

Private Sub Workbook_Open()
#If CompMan Then
    mCompMan.UpdateUsedCommCompsTheOriginHasChanged wbTarget:=ThisWorkbook, sHosted:=HOSTED_COMMON_COMPS
#End If
End Sub

Private Sub Workbook_BeforeSave(ByVal SaveAsUI As Boolean, Cancel As Boolean)
#If CompMan Then
    mCompMan.ExportChangedComponents wb:=ThisWorkbook, sHosted:=HOSTED_COMMON_COMPS
#End If
End Sub
