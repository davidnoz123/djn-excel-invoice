VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "ThisWorkbook"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = True
Option Explicit

Private Sub Workbook_BeforeClose(Cancel As Boolean)
Call AutoBackup.Terminate("BeforeClose")
End Sub

Private Sub Workbook_Open()
On Error Resume Next
Call AutoBackup.xRunAutoBackup
End Sub

Private Sub Workbook_SheetChange(ByVal Sh As Object, ByVal target As Range)
Call AutoBackup.SignalChange(0)
End Sub








