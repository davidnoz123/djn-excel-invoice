Attribute VB_Name = "Handlers"
Option Explicit

Sub RibbonOnLoad(ribbon As IRibbonUI)

End Sub

Private Function ControlGenericHandler_CreateObject(sClass As String) As Object
Dim o As Object
Select Case LCase(Trim(sClass))
Case "ribbondocumentation": Set o = New RibbonDocumentation
Case "ribboncontrolwrapper": Set o = New RibbonControlWrapper
Case Else
  Call ErrorExit("Handlers.ControlGenericHandler1:Unexpected class name:" + sClass)
End Select
Set ControlGenericHandler_CreateObject = o
End Function

Private Sub ControlGenericHandler1(control As IRibbonControl, eventName As String, index As Integer, ByRef returnedVal)
Dim ss: ss = Split(control.id, ".")
If UBound(ss) < 1 Then
  Call ErrorExit("Handlers.ControlGenericHandler1:UBound(ss) <> 1:'" + control.id + "'")
End If
Dim sClass As String: sClass = LCase(Trim(ss(0)))
Dim o As Object: Set o = ControlGenericHandler_CreateObject(sClass)
On Error GoTo 10
Dim sMethod As String: sMethod = LCase(Trim(ss(1)))
Call CallByName(o, sMethod, VbMethod, control, eventName, index, returnedVal)
Exit Sub
10:
Call ErrorExit("Handlers.ControlGenericHandler1:" + Err.Description)
End Sub

Private Sub CustomUI_GetLabel(control As IRibbonControl, ByRef returnedVal)
Call ControlGenericHandler1(control, "getLabel", -1, returnedVal)
End Sub
Private Sub CustomUI_GetItemCount(control As IRibbonControl, ByRef returnedVal)
Call ControlGenericHandler1(control, "getItemCount", -1, returnedVal)
End Sub
Private Sub CustomUI_GetItemLabel(control As IRibbonControl, index As Integer, ByRef returnedVal)
Call ControlGenericHandler1(control, "getItemLabel", index, returnedVal)
End Sub
Private Sub CustomUI_GetSelectedItemIndex(control As IRibbonControl, ByRef returnedVal)
Call ControlGenericHandler1(control, "getSelectedItemIndex", -1, returnedVal)
End Sub
Sub CustomUI_OnAction(control As IRibbonControl, Optional id As String = "", Optional index As Integer = -1)
Dim returnedVal
Call ControlGenericHandler1(control, "onAction", index, returnedVal)
End Sub
Private Sub CustomUI_GetText(control As IRibbonControl, ByRef returnedVal)
Call ControlGenericHandler1(control, "getText", -1, returnedVal)
End Sub
Private Sub CustomUI_OnChange(control As IRibbonControl, ByRef returnedVal)
Call ControlGenericHandler1(control, "onChange", -1, returnedVal)
End Sub
Private Sub CustomUI_GetPressed(control As IRibbonControl, ByRef returnedVal)
Dim v
Call ControlGenericHandler1(control, "getPressed", -1, v)
If TypeName(v) = "Boolean" Then
  returnedVal = v
Else
  returnedVal = False
End If
End Sub
Private Sub CustomUI_GetScreentip(control As IRibbonControl, ByRef returnedVal)
Call ControlGenericHandler1(control, "getScreentip", -1, returnedVal)
End Sub
Private Sub CustomUI_GetSupertip(control As IRibbonControl, ByRef returnedVal)
Call ControlGenericHandler1(control, "getSupertip", -1, returnedVal)
End Sub




