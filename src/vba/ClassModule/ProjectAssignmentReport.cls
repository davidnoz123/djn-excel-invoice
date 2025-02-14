VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "ProjectAssignmentReport"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit


Sub ProcessCustomUIEvent(control As IRibbonControl, eventName As String, index As Integer, ByRef returnedVal)
Dim v
Dim pd As New PersistentData
Call pd.Init("ProjectAssignmentReportParams")
Dim ss: ss = Split(control.id, ".")
If UBound(ss) < 2 Then
  Call ErrorExit("ProjectAssignmentReport.ProcessCustomUIEvent:UBound(ss) < 2:" + control.id)
End If

Dim sTag As String: sTag = LCase(Trim(ss(2)))
Select Case sTag
Case "runreport":
Case Else
  Call ErrorExit("ProjectAssignmentReport.ProcessCustomUIEvent:Unexpected tag name:" + ss(2))
End Select
Dim pdoc As RibbonDocumentation
Dim sEventName As String:  sEventName = LCase(Trim(eventName))
Select Case sEventName
Case "getlabel"
  Select Case sTag
  Case "runreport": returnedVal = "Run"
  End Select
Case "gettext":
  Select Case sTag
  Case "push_table_names_list", "pull_table_names_list":
    Call pd.GetValue(sTag + "." + "gettext", v, Empty)
    If IsEmpty(v) Then
      v = ""
      Call pd.SetValue(sTag + "." + "gettext", "")
    End If
    returnedVal = v
  Case Else
    Call ErrorExit("ProjectAssignmentReport.ProcessCustomUIEvent:Unexpected tag name:" + ss(2))
  End Select
Case "onchange":
  Select Case sTag
  Case "push_table_names_list", "pull_table_names_list":
    Call pd.SetValue(sTag + "." + "gettext", returnedVal)
  Case Else

  End Select
Case "onaction"
  Select Case sTag
  Case "push_table_names_list", "pull_table_names_list":
    
  Case "runreport":
    Call OnAction_runreport(pd)
  Case "localotherwiseremote":
    Call pd.GetValue(sTag + "." + "gettext", v, Empty)
    If IsEmpty(v) Or TypeName(v) <> "Boolean" Then
      Debug.Assert False
    Else
      Call pd.SetValue(sTag + "." + "gettext", Not v)
    End If
  End Select
Case "getscreentip", "getsupertip":
  Set pdoc = New RibbonDocumentation: Call pdoc.Init("ProjectAssignmentReportParams")
  Dim sScreenTip As String, sSuperTip As String
  Call pdoc.GetTipDoc(sTag, sScreenTip, sSuperTip)
  If sEventName = "getscreentip" Then
    returnedVal = sScreenTip
  Else
    returnedVal = sSuperTip
  End If
Case "getpressed"
  Select Case sTag
  Case "localotherwiseremote":
    Call pd.GetValue(sTag + "." + "gettext", v, Empty)
    If Not IsEmpty(v) And TypeName(v) = "Boolean" Then
      returnedVal = v
    Else
      Call pd.SetValue(sTag + "." + "gettext", False)
      returnedVal = False
    End If
  End Select
Case Else
  Call ErrorExit("ProjectAssignmentReport.ProcessCustomUIEvent:Unexpected event name:" + eventName)
End Select

End Sub

Sub OnAction_runreport(pd As PersistentData)
RunReport
End Sub

