VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "wbErHDemo"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = True
Option Explicit

Const HOSTED_RAWS = "mMsg, fMsg"

Private Sub Workbook_Open()
   mCompManClient.CompManService "UpdateOutdatedCommonComponents", HOSTED_RAWS
End Sub

Private Sub Workbook_BeforeSave(ByVal SaveAsUI As Boolean, Cancel As Boolean)
    mCompManClient.CompManService "ExportChangedComponents", HOSTED_RAWS
End Sub

