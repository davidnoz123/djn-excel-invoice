VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "PersistentData"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

Dim ws As Worksheet

Sub Init(wsName As String)
If WsExists(ThisWorkbook, wsName) Then
  Set ws = ThisWorkbook.Worksheets(wsName)
Else
  Set ws = SafeGetWorksheet(ThisWorkbook, wsName, True)
  ws.Visible = xlSheetHidden
End If
End Sub

Sub SetValue(sKey As String, v)
Dim a: a = ws.UsedRange
If Not IsArray(a) Then
  ws.Cells(1, 1) = sKey
  ws.Cells(1, 2) = v
Else
  Dim iColBeg As Long: iColBeg = ws.UsedRange.Columns(1).Column
  Dim iRowBeg As Long: iRowBeg = ws.UsedRange.Rows(1).row
  Dim iRowEnd As Long: iRowEnd = ws.UsedRange.Rows(ws.UsedRange.Rows.count).row
  Dim i As Long
  For i = LBound(a) To UBound(a)
    If a(i, 1) = sKey Then
      ws.Cells(iRowBeg + i - LBound(a), iColBeg + 1) = v
      GoTo finished
    End If
  Next i
  ws.Cells(iRowEnd + 1, iColBeg + 0) = sKey
  ws.Cells(iRowEnd + 1, iColBeg + 1) = v
finished:
End If
End Sub

Function GetValue(sKey As String, ByRef v, defaultValue) As Boolean
Dim a: a = ws.UsedRange
If Not IsArray(a) Then
  v = defaultValue
Else
  Dim i As Long
  For i = LBound(a) To UBound(a)
    If a(i, 1) = sKey Then
      v = a(i, 2)
      GoTo finished
    End If
  Next i
  v = defaultValue
finished:
End If
End Function


