Attribute VB_Name = "Utils"
Option Explicit

Sub ErrorExit(msg As String)
Call Logg("Error : " + Replace(msg, vbCrLf, " "))
MsgBox msg, vbCritical, "Error"
End
End Sub

Sub Logg(msg As String)
On Error Resume Next
Debug.Print msg
Application.StatusBar = Left(Replace(msg, vbLf, " "), 256)
End Sub

Function Safe_ThisWorkbook_Worksheets(name As String) As Worksheet
On Error GoTo exit_with_failure
Set Safe_ThisWorkbook_Worksheets = ThisWorkbook.Worksheets(name)
Exit Function
exit_with_failure:
Call ErrorExit("Missing required worksheet:'" + name + "'")
End Function

Function FindCol(a, name, ByVal start_pos As Long) As Long
If start_pos > UBound(a) Then
  start_pos = LBound(a)
End If
If a(LBound(a), start_pos) = name Then
  FindCol = start_pos
Else
  Dim i As Long
  For i = LBound(a, 2) To UBound(a, 2)
    If a(LBound(a), i) = name Then
      FindCol = i
      GoTo 10
    End If
  Next i
  Call ErrorExit("MissingColumnHeader:" + name)
10:
End If
End Function

Function FindNextWiseCell(rTopLeft As Range, ByVal bGetRowOtherwiseColumn As Boolean, Optional ByVal What As String = "") As Long
Dim ws As Worksheet: Set ws = rTopLeft.Parent
Dim iRow_Beg As Long: iRow_Beg = rTopLeft.row
Dim iCol_Beg As Long: iCol_Beg = rTopLeft.Column
Dim rr As Range
If bGetRowOtherwiseColumn Then
  Set rr = ws.Columns(iCol_Beg).Find(What:=What, After:=ws.Cells(iRow_Beg, iCol_Beg), LookIn:=xlValues, LookAt _
          :=xlPart, searchorder:=xlByRows, SearchDirection:=xlNext, MatchCase:= _
          False, SearchFormat:=False)
  If rr Is Nothing Then
    FindNextWiseCell = 0
  Else
    FindNextWiseCell = rr.row - 1
  End If
Else
  Set rr = ws.Rows(iRow_Beg).Find(What:=What, After:=ws.Cells(iRow_Beg, iCol_Beg), LookIn:=xlValues, LookAt _
          :=xlPart, searchorder:=xlByRows, SearchDirection:=xlNext, MatchCase:= _
          False, SearchFormat:=False)
  If rr Is Nothing Then
    FindNextWiseCell = 0
  Else
    FindNextWiseCell = rr.Column - 1
  End If
End If
End Function

Function GetLastNonBlankCell(rTopLeft As Range, ByVal bGetRowOtherwiseColumn As Boolean) As Long
GetLastNonBlankCell = FindNextWiseCell(rTopLeft, bGetRowOtherwiseColumn, What:="")
End Function

Function GetData(rTopLeft As Range, Optional ByVal nRows As Long = 0, Optional ByVal nCols As Long = 0) As Variant()
Dim ws As Worksheet: Set ws = rTopLeft.Parent
Dim iRow_Beg As Long: iRow_Beg = rTopLeft.row
Dim iCol_Beg As Long: iCol_Beg = rTopLeft.Column

Dim iRow_End As Long
If nRows <= 0 Then
  iRow_End = GetLastNonBlankCell(rTopLeft, True)
Else
  iRow_End = iRow_Beg + nRows - 1
End If

Dim iCol_End As Long
If nCols <= 0 Then
  iCol_End = GetLastNonBlankCell(rTopLeft, False)
Else
  iCol_End = iCol_Beg + nCols - 1
End If

Dim a()
If iRow_Beg = iRow_End And iCol_Beg = iCol_End Then
  ReDim a(1 To 1, 1 To 1): a(1, 1) = ws.Cells(iRow_Beg, iCol_Beg)
Else
  If iRow_Beg <= iRow_End And iCol_Beg <= iCol_End Then
    a = ws.Range(ws.Cells(iRow_Beg, iCol_Beg), ws.Cells(iRow_End, iCol_End))
  End If
End If

GetData = a

End Function


Function NumberOfDimensions(a) As Long
If Not IsArray(a) Then
  NumberOfDimensions = 0
Else
  On Error GoTo FinalDimension
  Dim ErrorCheck As Long
  For NumberOfDimensions = 1 To 60000
     ErrorCheck = LBound(a, NumberOfDimensions)
  Next NumberOfDimensions
FinalDimension:
NumberOfDimensions = NumberOfDimensions - 1
End If
End Function

Function PutData(rTopLeft As Range, a, Optional bTranspose As Boolean = False) As Range
Dim ws As Worksheet: Set ws = rTopLeft.Parent
Dim iRow_Beg As Long: iRow_Beg = rTopLeft.row
Dim iCol_Beg As Long: iCol_Beg = rTopLeft.Column
Dim iRow_End As Long
Dim iCol_End As Long
Dim n As Long: n = NumberOfDimensions(a)
Dim r As Range
Select Case n
Case 0
  rTopLeft = Empty
Case 1
  If bTranspose Then
    iRow_End = iRow_Beg
    iCol_End = iCol_Beg + UBound(a) - LBound(a)
    Set r = ws.Range(ws.Cells(iRow_Beg, iCol_Beg), ws.Cells(iRow_End, iCol_End)): r = a
  Else
    iRow_End = iRow_Beg + UBound(a) - LBound(a)
    iCol_End = iCol_Beg
    Set r = ws.Range(ws.Cells(iRow_Beg, iCol_Beg), ws.Cells(iRow_End, iCol_End)): r = WorksheetFunction.Transpose(a)
  End If
Case 2
  If bTranspose Then
    iRow_End = iRow_Beg + UBound(a, 2) - LBound(a, 2)
    iCol_End = iCol_Beg + UBound(a) - LBound(a)
    Set r = ws.Range(ws.Cells(iRow_Beg, iCol_Beg), ws.Cells(iRow_End, iCol_End)): r = WorksheetFunction.Transpose(a)
  Else
    iRow_End = iRow_Beg + UBound(a) - LBound(a)
    iCol_End = iCol_Beg + UBound(a, 2) - LBound(a, 2)
    Set r = ws.Range(ws.Cells(iRow_Beg, iCol_Beg), ws.Cells(iRow_End, iCol_End)): r = a
  End If
Case Else
  Call ErrorExit("PutData:Unexpected number of dimensions")
End Select
Set PutData = r
End Function

Function WsExists(wb As Workbook, sName As String) As Boolean
On Error GoTo 10
Dim ws As Worksheet
Set ws = wb.Sheets(sName)
WsExists = True
Exit Function
10:
WsExists = False
End Function

Function SafeGetWorksheet(wb As Workbook, sWorksheetName As String, Optional ByVal bEnsureEmpty As Boolean = False) As Worksheet
If wb Is Nothing Then
  Call ErrorExit("Utils.SafeGetWorksheet:wb Is Nothing")
End If
On Error GoTo 10
Set SafeGetWorksheet = wb.Sheets(sWorksheetName)
If bEnsureEmpty Then
  SafeGetWorksheet.Cells.Delete
  Call ClearAllNames(SafeGetWorksheet)
  Call ClearAllShapes(SafeGetWorksheet)
End If
Exit Function
10:
Set SafeGetWorksheet = SafeAddWorksheet(wb, sWorksheetName)
End Function

Function SafeAddWorksheet(wb As Workbook, sWorksheetName As String) As Worksheet
On Error Resume Next
Dim rSelection As Range: Set rSelection = Selection
On Error GoTo 10
Dim bScreenUpdating As Boolean: bScreenUpdating = Application.ScreenUpdating
Application.ScreenUpdating = False
Set SafeAddWorksheet = wb.Sheets.Add()
Application.ScreenUpdating = bScreenUpdating
SafeAddWorksheet.name = sWorksheetName
GoTo 20
10:
Call ErrorExit("Utils.SafeAddWorksheet: Failed to add worksheet : " + sWorksheetName)
20:
On Error Resume Next
rSelection.Parent.Activate
rSelection.Select
End Function

Sub ClearAllNames(ws As Worksheet)
While ws.Names.count > 0
  Dim n As name
  Set n = ws.Names(1)
  n.Delete
Wend
End Sub

Sub ClearAllShapes(ws As Worksheet)
ws.Pictures.Delete
While ws.Shapes.count > 0
  Dim s As Shape
  Set s = ws.Shapes(1)
  s.Delete
Wend
End Sub

Function GetRangeFromString(rangeStr As String) As Range
On Error GoTo exit_with_failure
Dim ws As Worksheet
Dim arr() As String
arr = Split(rangeStr, "!")
If UBound(arr) <> 1 Then
    Set GetRangeFromString = Nothing
Else
    Set ws = ThisWorkbook.Sheets(Replace(arr(0), "'", ""))
    Set GetRangeFromString = ws.Range(arr(1))
End If
Exit Function
exit_with_failure:
End Function

Function SafeFindKeyValueInUsedRange(ws As Worksheet, key As String, Optional default_value)
Dim rr As Range: Set rr = ws.UsedRange.Find(key)
If Not rr Is Nothing Then
    SafeFindKeyValueInUsedRange = ws.Cells(rr.row, rr.Column + 1)
Else
    If IsMissing(default_value) Then
        ws.Parent.Activate
        ws.Activate
        Call ErrorExit("Missing key-value cell for key:'" + key + "'")
    End If
    SafeFindKeyValueInUsedRange = default_value
End If
End Function
