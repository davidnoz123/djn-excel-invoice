VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "RibbonDocumentation"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

Dim wsDocs As Worksheet

Sub ProcessCustomUIEvent(control As IRibbonControl, eventName As String, index As Integer, ByRef returnedVal)
Dim v
Dim pd As New PersistentData
Call pd.Init("RibbonDocumentation")
Dim ss: ss = Split(control.id, ".")
If UBound(ss) < 2 Then
  Call ErrorExit("RibbonDocumentation.ProcessCustomUIEvent:UBound(ss) < 2:" + control.id)
End If

Dim sTag As String: sTag = LCase(Trim(ss(2)))
Select Case sTag
Case "edittooltips":
Case Else
  Call ErrorExit("RibbonDocumentation.ProcessCustomUIEvent:Unexpected tag name:" + ss(2))
End Select

Dim pdoc As RibbonDocumentation
Dim sEventName As String:  sEventName = LCase(Trim(eventName))
Select Case sEventName
Case "getlabel"
  Select Case sTag
  Case "edittooltips": returnedVal = "EditTips"
  End Select
Case "onaction"
  Select Case sTag
  Case "edittooltips":
    EditAll
  End Select
Case "getscreentip", "getsupertip":
  Set pdoc = New RibbonDocumentation: Call pdoc.Init("RibbonDocumentationDocs")
  Dim sScreenTip As String, sSuperTip As String
  Call pdoc.GetTipDoc(sTag, sScreenTip, sSuperTip)
  If sEventName = "getscreentip" Then
    returnedVal = sScreenTip
  Else
    returnedVal = sSuperTip
  End If
Case Else
  Call ErrorExit("RibbonDocumentation.ProcessCustomUIEvent:Unexpected event name:" + eventName)
End Select

End Sub

Sub Init(sWorksheetName As String)
If WsExists(ThisWorkbook, sWorksheetName) Then
  Set wsDocs = ThisWorkbook.Worksheets(sWorksheetName)
Else
  Set wsDocs = SafeGetWorksheet(ThisWorkbook, sWorksheetName)
  wsDocs.Visible = xlSheetHidden
End If
End Sub

Private Sub Class_Initialize()
Set wsDocs = Nothing
End Sub

Private Function FindTipDocRow(sKey As String) As Long
Dim rr As Range
Set rr = wsDocs.Columns(1).Find(What:=sKey, After:=wsDocs.Cells(1, 1), LookIn:=xlValues, LookAt:=xlWhole, SearchOrder:=xlByRows, SearchDirection:=xlNext, MatchCase:=False, SearchFormat:=False)
Dim iRow As Long
If Not rr Is Nothing Then
  iRow = rr.row
Else
  iRow = wsDocs.UsedRange.Rows(wsDocs.UsedRange.Rows.count).row + 1
  wsDocs.Cells(iRow, 1) = sKey
  wsDocs.Cells(iRow, 2) = sKey
  wsDocs.Cells(iRow, 3) = "No help yet"
End If
FindTipDocRow = iRow
End Function

Sub GetTipDoc(sKey As String, ByRef sScreenTip As String, ByRef sSuperTip As String)
Dim iRow As Long: iRow = FindTipDocRow(sKey)
sScreenTip = CStr(wsDocs.Cells(iRow, 2))
sSuperTip = CStr(wsDocs.Cells(iRow, 3))
End Sub

Sub SetTipDoc(sKey As String, ByVal sScreenTip As String, ByVal sSuperTip As String)
Dim iRow As Long: iRow = FindTipDocRow(sKey)
wsDocs.Cells(iRow, 2) = sScreenTip
wsDocs.Cells(iRow, 3) = sSuperTip
End Sub

Sub GetKeys(ByRef col As Collection)
Set col = New Collection
Dim a: a = wsDocs.Range(wsDocs.Cells(1, 1), wsDocs.Cells(wsDocs.UsedRange.Rows(wsDocs.UsedRange.Rows.count).row, 1))
If IsArray(a) Then
  Dim i As Long
  For i = LBound(a) To UBound(a)
    If Trim(a(i, 1)) <> "" Then
      col.Add CStr(a(i, 1))
    End If
  Next i
End If
End Sub

Sub EditAll()
Dim col As New Collection
col.Add "GAMSModelDocs"
col.Add "RibbonDocumentationDocs"
Dim i As Long
For i = 1 To col.count
  Call Edit(CStr(col(i)))
Next i
End Sub

Sub Edit(sWorksheetName As String)
Call Init(sWorksheetName)
Dim fso As New FileSystemObject
Dim wsh As New WshShell
Dim col As Collection
Call GetKeys(col)
Dim i As Long, k As Long
For i = 1 To col.count
  Dim sKey As String: sKey = col(i)
  Dim sScreenTip As String, sSuperTip As String
  Call GetTipDoc(sKey, sScreenTip, sSuperTip)
  Dim sTmpTxt As String: sTmpTxt = ThisWorkbook_Path + "\" + "tmp.txt"
  Dim ts As TextStream
  Set ts = fso.CreateTextFile(sTmpTxt, True)
  Call ts.Write(sScreenTip + vbCrLf + Replace(sSuperTip, vbLf, vbCrLf))
  Call ts.Close
  Call wsh.Run("notepad """ + sTmpTxt + """", vbMaximizedFocus, True)
  Set ts = fso.OpenTextFile(sTmpTxt, ForReading)
  Dim sOut As String: sOut = ts.ReadAll
  Call ts.Close
  Dim ss: ss = Split(sOut, vbCrLf)
  sScreenTip = ""
  sSuperTip = ""
  For k = LBound(ss) To UBound(ss)
    If k = LBound(ss) Then
      sScreenTip = ss(k)
    Else
      sSuperTip = sSuperTip + IIf(sSuperTip = "", "", vbLf) + ss(k)
    End If
  Next k
  Call SetTipDoc(sKey, sScreenTip, sSuperTip)
Next i
End Sub




