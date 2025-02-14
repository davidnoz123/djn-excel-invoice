Attribute VB_Name = "Utils"
Option Explicit

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



'For iRow_End = iRow_Beg To 999999
'  DoEvents
'  If ws.Cells(iRow_End, iCol_Beg) = "" Then
'    iRow_End = iRow_End - 1
'    Exit For
'  End If
'Next iRow_End

'For iCol_End = iCol_Beg To 999999
'  DoEvents
'  If ws.Cells(iRow_Beg, iCol_End) = "" Then
'    iCol_End = iCol_End - 1
'    Exit For
'  End If
'Next iCol_End
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
  'Sets up the error handler.
  On Error GoTo FinalDimension
  'Visual Basic for Applications arrays can have up to 60000
  'dimensions; this allows for that.
  Dim ErrorCheck As Long
  For NumberOfDimensions = 1 To 60000
     'It is necessary to do something with the LBound to force it
     'to generate an error.
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
' deletes a sheet named strSheetName in the active workbook
'Application.DisplayAlerts = False
'SafeGetWorksheet.Delete
'Application.DisplayAlerts = True
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

Public Function ShellSortCollectionOfString(c, Optional comp As VbCompareMethod = vbBinaryCompare) As Collection
Dim v
Dim i As Long
If IsArray(c) Then
  ReDim a(LBound(c) To UBound(c))
  ReDim permIndices(LBound(c) To UBound(c))
  For i = LBound(c) To UBound(c)
    a(i) = c(i)
    permIndices(i) = i
  Next i
Else
  ReDim a(1 To c.count)
  ReDim permIndices(1 To c.count)
  i = 1
  For Each v In c
    a(i) = v
    permIndices(i) = i
    i = i + 1
  Next v
End If
Call ShellSortAString(a, permIndices:=permIndices)
Dim ret As New Collection
For i = LBound(permIndices) To UBound(permIndices)
  Call ret.Add(a(permIndices(i)))
Next
Set ShellSortCollectionOfString = ret
End Function

Public Sub ShellSortAString(a, Optional comp As VbCompareMethod = vbBinaryCompare, Optional permIndices = Empty)
' Note for a,b as long: Compare(a,b) = -1   if Object(a) < Object(b)
'                 Compare(a,b) = 0      if Object(a) = Object(b)
'                 Compare(a,b) = 1      if Object(a) > Object(b)
Dim f As Long
Dim s As Long
Dim t As Long
Dim n2 As Long
Dim ls As Long
Dim i As Long
Dim is2 As Long
Dim j As Long
Dim js As Long
Dim l As Long
Dim Done As Boolean

Dim tmp
      
If IsEmpty(permIndices) Then
  f = LBound(a)
  l = UBound(a)
  n2 = (l - f + 1) \ 2
  s = 1023
  For t = 1 To 10
    If (s <= n2) Then
      ls = l - s
      For i = f To ls
        is2 = i + s
        j = i
        js = is2
        Done = (StrComp(a(js), a(j), comp) >= 0)
        While Not Done
          tmp = a(js): a(js) = a(j): a(j) = tmp
          js = j
          j = j - s
          If (j < f) Then
            Done = True
          Else
            Done = (StrComp(a(js), a(j), comp) >= 0)
          End If
        Wend
      Next i
    End If
    s = s \ 2
  Next t
Else
  f = LBound(permIndices)
  l = UBound(permIndices)
  n2 = (l - f + 1) \ 2
  s = 1023
  For t = 1 To 10
    If (s <= n2) Then
      ls = l - s
      For i = f To ls
        is2 = i + s
        j = i
        js = is2
        Done = (StrComp(a(permIndices(js)), a(permIndices(j)), comp) >= 0)
        While Not Done
          tmp = permIndices(js): permIndices(js) = permIndices(j): permIndices(j) = tmp
          js = j
          j = j - s
          If (j < f) Then
            Done = True
          Else
            Done = (StrComp(a(permIndices(js)), a(permIndices(j)), comp) >= 0)
          End If
        Wend
      Next i
    End If
    s = s \ 2
  Next t
End If

End Sub

Public Sub ShellSortPairwiseLogic(a, permIndices)
' Note for a,b as long: Compare(a,b) = -1   if Object(a) < Object(b)
'                 Compare(a,b) = 0      if Object(a) = Object(b)
'                 Compare(a,b) = 1      if Object(a) > Object(b)
Dim f As Long
Dim s As Long
Dim t As Long
Dim n2 As Long
Dim ls As Long
Dim i As Long
Dim is2 As Long
Dim j As Long
Dim js As Long
Dim l As Long
Dim Done As Boolean

Dim tmp
      
f = LBound(permIndices)
l = UBound(permIndices)
n2 = (l - f + 1) \ 2
s = 1023
For t = 1 To 10
  If (s <= n2) Then
    ls = l - s
    For i = f To ls
      is2 = i + s
      j = i
      js = is2
      Done = a(permIndices(js), permIndices(j))
      While Not Done
        tmp = permIndices(js): permIndices(js) = permIndices(j): permIndices(j) = tmp
        js = j
        j = j - s
        If (j < f) Then
          Done = True
        Else
          Done = a(permIndices(js), permIndices(j))
        End If
      Wend
    Next i
  End If
  s = s \ 2
Next t

End Sub


Function WorkbookIsOpen(sPath As String, ByRef wb As Workbook) As Boolean
Dim fso As New FileSystemObject
Dim sName As String: sName = fso.GetFileName(sPath)
On Error GoTo 10
Set wb = Application.Workbooks(sName)
WorkbookIsOpen = True
Exit Function
10:
Set wb = Nothing
WorkbookIsOpen = False
End Function
