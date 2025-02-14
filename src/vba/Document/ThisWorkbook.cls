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
If ThisWorkbook.name = "avenrose_florist.xlsm" Then
  Call RefreshFromOneDrive
  Call AutoBackup.xRunAutoBackup
ElseIf InStr(ThisWorkbook.name, "update.") = 1 Then
  On Error GoTo exit_with_error
  Dim Application_ScreenUpdating: Application_ScreenUpdating = Application.ScreenUpdating
  Application.ScreenUpdating = False
  Call UpdateFromThisWorkbook(0)
exit_with_error:
  Application.ScreenUpdating = Application_ScreenUpdating
  ThisWorkbook.Close
End If
End Sub

Private Sub Workbook_SheetChange(ByVal Sh As Object, ByVal target As Range)
Call AutoBackup.SignalChange(0)
End Sub



