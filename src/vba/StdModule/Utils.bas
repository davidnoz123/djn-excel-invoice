Attribute VB_Name = "Utils"
Option Explicit

Sub ErrorExit(msg As String)
Call Logg("Error : " + Replace(msg, vbCrLf, " "))
MsgBox msg, vbCritical, "Error"
End
End Sub

Sub ErrorExitX(msg As String, Optional rng As Range = Nothing)
If Not rng Is Nothing Then
    On Error Resume Next
    rng.Parent.Parent.Activate
    rng.Parent.Activate
    rng.Select
End If
Call ErrorExit(msg)
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

Function FindCol(a, name) As Long
Dim i As Long
For i = LBound(a, 2) To UBound(a, 2)
  If a(LBound(a), i) = name Then
    FindCol = i
    GoTo 10
  End If
Next i
ReDim b(LBound(a, 2) To UBound(a, 2)) As String
For i = LBound(a, 2) To UBound(a, 2)
   b(i) = CStr(a(LBound(a), i))
Next i
Call ErrorExit("MissingExpectedColumnHeader:'" + name + "'" + vbCrLf + "AvailableHeaders:['" + Join(b, "','") + "']")
10:
End Function

Function FindNextWiseCell(rTopLeft As Range, ByVal bGetRowOtherwiseColumn As Boolean, Optional ByVal What As String = "") As Long
Dim ws As Worksheet: Set ws = rTopLeft.Parent
Dim iRow_Beg As Long: iRow_Beg = rTopLeft.row
Dim iCol_Beg As Long: iCol_Beg = rTopLeft.Column
Dim rr As Range
If bGetRowOtherwiseColumn Then
  Set rr = ws.Columns(iCol_Beg).Find(What:=What, After:=ws.Cells(iRow_Beg, iCol_Beg), LookIn:=xlValues, LookAt _
          :=xlPart, SearchOrder:=xlByRows, SearchDirection:=xlNext, MatchCase:= _
          False, SearchFormat:=False)
  If rr Is Nothing Then
    FindNextWiseCell = 0
  Else
    FindNextWiseCell = rr.row - 1
  End If
Else
  Set rr = ws.Rows(iRow_Beg).Find(What:=What, After:=ws.Cells(iRow_Beg, iCol_Beg), LookIn:=xlValues, LookAt _
          :=xlPart, SearchOrder:=xlByRows, SearchDirection:=xlNext, MatchCase:= _
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

Function WbExists(sName As String) As Boolean
On Error GoTo 10
Dim wb As Workbook
Set wb = Workbooks(sName)
WbExists = True
Exit Function
10:
WbExists = False
End Function

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
Dim bScreenUpdating As Boolean: bScreenUpdating = Application.ScreenUpdating
On Error GoTo 10
Application.ScreenUpdating = False
Set SafeAddWorksheet = wb.Sheets.Add()
Application.ScreenUpdating = bScreenUpdating
SafeAddWorksheet.name = sWorksheetName
GoTo 20
10:
Application.ScreenUpdating = bScreenUpdating
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

Sub ExportWorksheetPDFSilent(ws As Worksheet, pdf_name As String, delete_worksheet As Boolean, open_after_publish As Boolean)
Dim bScreenUpdating As Boolean: bScreenUpdating = Application.ScreenUpdating
On Error GoTo 80
Application.ScreenUpdating = False
ws.ExportAsFixedFormat Type:=xlTypePDF, Filename:=pdf_name, Quality:=xlQualityStandard, IncludeDocProperties:=True, IgnorePrintAreas:=False, OpenAfterPublish:=open_after_publish
GoTo 90
80:
On Error GoTo 0
Application.ScreenUpdating = bScreenUpdating
If delete_worksheet Then Call DeleteWorksheetSilent(ws)
Call ErrorExit("Failure exporting PDF:'" + Err.Description + "'")
90:
Application.ScreenUpdating = bScreenUpdating
If delete_worksheet Then Call DeleteWorksheetSilent(ws)
End Sub

Sub DeleteWorksheetSilent(ws)
Dim bDisplayAlerts As Boolean: bDisplayAlerts = Application.DisplayAlerts
Application.DisplayAlerts = False
ws.Delete
Application.DisplayAlerts = bDisplayAlerts
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

Function SafeFindKeyValueInRange(r As Range, key As String, Optional default_value)
Dim rr As Range: Set rr = r.Find(key)
If Not rr Is Nothing Then
    SafeFindKeyValueInRange = r.Parent.Cells(rr.row, rr.Column + 1)
Else
    If IsMissing(default_value) Then
        r.Parent.Parent.Activate
        r.Parent.Activate
        Call ErrorExit("Missing key-value cell for key:'" + key + "'")
    End If
    SafeFindKeyValueInRange = default_value
End If
End Function

Sub ErrorExitWithWord(doc As Word.Document, msg As String, Optional r As Word.Range = Nothing)
On Error GoTo 10
doc.Application.Activate
doc.Activate
If Not r Is Nothing Then r.Select
10:
Call ErrorExit(msg)
End Sub

Function c2a(c)
ReDim a(c.count - 1)
Dim kk, k As Long: k = 0
For Each kk In c
    a(k) = kk
    k = k + 1
Next kk
c2a = a
End Function


Function ParseCellReference(ref As String) As Variant
    Dim matches As Object
    Dim dirPath As String, wbName As String, wsName As String, cellRef As String
    
    ref = Trim(ref)
    
    ' Create regex object
    Dim regex As New VBScript_RegExp_55.RegExp
    regex.Global = True
    regex.IgnoreCase = True
    
    ' Regular expression pattern to match directory, workbook, sheet, and cell
    Dim a2 As String: a2 = "([^!\[\]]+\\)?"      ' Capture the **directory path**
    Dim a3 As String: a3 = "(\[?[^'!\]]+\]?)?"   ' Capture the **workbook name** inside `[ ]`
    Dim a4 As String: a4 = "([^\\\]!']+)?"       ' Capture the **worksheet name**
    Dim a5 As String: a5 = "(![^\\\]]*)?"        ' Capture the **cell reference**
    regex.Pattern = "^[']?" + a2 + a3 + a4 + "[']?" + a5 + "$"
    

    ' Execute regex match
    If regex.Test(ref) Then
        Set matches = regex.Execute(ref)
        
        ' Extract components
        If matches.count > 0 Then
            Dim i As Long: i = 0
            dirPath = matches(0).SubMatches(i)
            i = i + 1
            wbName = matches(0).SubMatches(i)
            If Len(wbName) > 0 Then
                If Left(wbName, 1) = "[" Then
                    wbName = Mid(wbName, 2, Len(wbName) - 2)
                Else
                    wsName = wbName
                    wbName = ""
                End If
            End If
            i = i + 1
            If wsName = "" Then
                wsName = matches(0).SubMatches(i)
            End If
            i = i + 1
            cellRef = matches(0).SubMatches(i)  ' Worksheet name
            If Len(cellRef) > 0 Then
                cellRef = Mid(cellRef, 2)
            End If
        End If
    End If
    
    ' Return array with results
    ParseCellReference = Array(dirPath, wbName, wsName, cellRef)
End Function

Sub aaaRunParseCellReferenceTests(xx)
    Dim testCases As Variant
    Dim i As Integer
    Dim Result As Variant
    Dim inputRef As String
    Dim expectedDir As String, expectedWb As String, expectedWs As String, expectedCell As String
    Dim pass As Boolean

    ' Define test cases: {Input, Expected Directory, Expected Workbook, Expected Worksheet, Expected Cell}
    testCases = Array( _
        Array("'C:\Users\John\Docs\[Workbook.xlsx]Sheet1'!A1", "C:\Users\John\Docs\", "Workbook.xlsx", "Sheet1", "A1"), _
        Array("[Workbook.xlsx]Sheet2!B5", "", "Workbook.xlsx", "Sheet2", "B5"), _
        Array("Sheet3!C10", "", "", "Sheet3", "C10"), _
        Array("'Sheet 1'!D20", "", "", "Sheet 1", "D20"), _
        Array("'C:\Folder\[Test File.xlsx]Sheet5'!E7", "C:\Folder\", "Test File.xlsx", "Sheet5", "E7"), _
        Array("[Data.xlsx]Sheet6!F9", "", "Data.xlsx", "Sheet6", "F9"), _
        Array("'[Another Workbook.xlsx]Sheet7'!G12", "", "Another Workbook.xlsx", "Sheet7", "G12"), _
        Array("C:\Folder\[Book.xlsx]SheetX!H15", "C:\Folder\", "Book.xlsx", "SheetX", "H15"), _
        Array("SheetOnly", "", "", "SheetOnly", ""), _
        Array("'[Workbook.xlsx]Sheet 10'!I5", "", "Workbook.xlsx", "Sheet 10", "I5"), _
        Array("'C:\Users\[MyWorkbook.xlsx]DataSheet'!J1", "C:\Users\", "MyWorkbook.xlsx", "DataSheet", "J1"), _
        Array("[Report.xlsx]AnnualReport!K8", "", "Report.xlsx", "AnnualReport", "K8"), _
        Array("'D:\Projects\[Budget.xlsx]Summary'!L4", "D:\Projects\", "Budget.xlsx", "Summary", "L4"), _
        Array("'[Client Data.xlsx]Overview'!M2", "", "Client Data.xlsx", "Overview", "M2"), _
        Array("'C:\Finance\[2024Report.xlsx]Revenue'!N9", "C:\Finance\", "2024Report.xlsx", "Revenue", "N9") _
    )

    ' Loop through test cases
    For i = LBound(testCases) To UBound(testCases)
        'If i <> 8 Then GoTo skip
        
        inputRef = testCases(i)(0)
        expectedDir = testCases(i)(1)
        expectedWb = testCases(i)(2)
        expectedWs = testCases(i)(3)
        expectedCell = testCases(i)(4)

        Result = ParseCellReference(inputRef)
        
        ' Check if results match expected values
        pass = (Result(0) = expectedDir) And (Result(1) = expectedWb) And (Result(2) = expectedWs) And (Result(3) = expectedCell)

        If pass Then
            'Debug.Print "? Test " & i & " PASSED: " & inputRef
        Else
            Debug.Print "? Test " & i & " FAILED: " & inputRef
            Debug.Print "   Expected: [" & expectedDir & "], [" & expectedWb & "], [" & expectedWs & "], [" & expectedCell & "]"
            Debug.Print "   Actual:   [" & Result(0) & "], [" & Result(1) & "], [" & Result(2) & "], [" & Result(3) & "]"
        End If
skip:
    Next i
    
    Debug.Print "? All tests completed."
End Sub



Function Disjunction(Range1 As Range, Range2 As Range) As Range
    Dim cell As Range
    Dim Result As Range
    
    ' Loop through Range1 and add cells not in Range2
    For Each cell In Range1
        If Intersect(cell, Range2) Is Nothing Then
            If Result Is Nothing Then
                Set Result = cell
            Else
                Set Result = Union(Result, cell)
            End If
        End If
    Next cell
    
    ' Loop through Range2 and add cells not in Range1
    For Each cell In Range2
        If Intersect(cell, Range1) Is Nothing Then
            If Result Is Nothing Then
                Set Result = cell
            Else
                Set Result = Union(Result, cell)
            End If
        End If
    Next cell
    
    ' Return result
    Set Disjunction = Result
End Function


Function GetNonFormulaCells(inputRange As Range) As Range
    Dim cell As Range
    Dim formulaRange As Range
    
    ' Loop through each cell in the input range
    For Each cell In inputRange
        ' Check if the cell contains a formula
        If Not cell.HasFormula Then
            ' Build the formula range dynamically
            If formulaRange Is Nothing Then
                Set formulaRange = cell
            Else
                Set formulaRange = Union(formulaRange, cell)
            End If
        End If
    Next cell
    
    ' Return the resulting range (could be Nothing if no formulas found)
    Set GetNonFormulaCells = formulaRange
End Function

Private Sub SafeGetWorkbook_SafeSaveAs(wb As Workbook, sPath As String, ByVal FileFormat As XlFileFormat)
Dim bDisplayAlerts As Boolean: bDisplayAlerts = Application.DisplayAlerts
Application.DisplayAlerts = False
On Error GoTo 10
Call wb.SaveAs(sPath, FileFormat:=FileFormat)
Application.DisplayAlerts = bDisplayAlerts
Exit Sub
10:
Application.DisplayAlerts = bDisplayAlerts
Call ErrorExit("Utils.SafeGetWorkbook_SafeSaveAs:Failed to save workbook:" + Err.Description)
End Sub

Function SafeGetWorkbook(sPath As String, ByRef bFileAlreadyOpen As Boolean, Optional sTemplate As String = "", Optional ByVal bEnsureEmpty As Boolean = False, Optional ByVal bUpdateLinks As Boolean = True) As Workbook
Dim wb As Workbook: Set wb = Nothing
If Trim(sPath) = "" Then
  GoTo exit_now
End If

If sTemplate <> "" And bEnsureEmpty Then
  Call ErrorExit("Utils.SafeGetWorkbook:Inconsistent arguments:sTemplate <> """" And bEnsureEmpty")
End If
  
Dim fso As New FileSystemObject
Dim sWorkbook As String: sWorkbook = fso.GetFileName(sPath)

If LCase(sWorkbook) = LCase(ThisWorkbook.name) Then
  If bEnsureEmpty Then
    Call ErrorExit("Utils.SafeGetWorkbook:LCase(sWorkbook) = LCase(ThisWorkbook.Name) And bEnsureEmpty")
  End If
  bFileAlreadyOpen = True
  Set wb = ThisWorkbook
Else
  Dim bReopenTemplate As Boolean: bReopenTemplate = False
  Dim bDisplayAlerts As Boolean: bDisplayAlerts = Application.DisplayAlerts
  Dim bScreenUpdating As Boolean: bScreenUpdating = Application.ScreenUpdating
  
  Dim wbActive As Workbook: Set wbActive = ActiveWorkbook
    
  bFileAlreadyOpen = WorkbookIsOpen(sPath, wb)
  If bFileAlreadyOpen Then
    If bEnsureEmpty Then
      Call wb.Close(SaveChanges:=False)
    Else
      If fso.GetAbsolutePathName(wb.path + "\\" + wb.name) = fso.GetAbsolutePathName(sPath) Then
        If sTemplate = "" Then
          GoTo exit_now
        End If
      End If
      Call wb.Close(SaveChanges:=True)
    End If
  End If
  
  If fso.FileExists(sPath) And Not bEnsureEmpty And sTemplate = "" Then
    Application.ScreenUpdating = False
    On Error Resume Next
    Set wb = Application.Workbooks.Open(sPath, UpdateLinks:=bUpdateLinks)
    If Not wbActive Is Nothing Then wbActive.Activate
    Application.ScreenUpdating = bScreenUpdating
  Else
  
    Dim bTemplateFileExists As Boolean: bTemplateFileExists = fso.FileExists(sTemplate)
    If bTemplateFileExists Then
      Dim wbTemplate As Workbook
      Dim bTemplateIsOpen As Boolean: bTemplateIsOpen = WorkbookIsOpen(sTemplate, wbTemplate)
      If bTemplateIsOpen Then
        If fso.GetAbsolutePathName(sTemplate) = fso.GetAbsolutePathName(wbTemplate.FullNameURLEncoded) Then
          bReopenTemplate = True
          Call wbTemplate.Close(SaveChanges:=True)
        End If
      End If
    End If
  
    Application.ScreenUpdating = False
    On Error Resume Next
    Set wb = Application.Workbooks.Add(IIf(bTemplateFileExists, sTemplate, ""))
    If Not wbActive Is Nothing Then wbActive.Activate
    Application.ScreenUpdating = bScreenUpdating
    
    Dim FileFormat As XlFileFormat
    Dim sExt As String: sExt = LCase(Trim(fso.GetExtensionName(sPath)))
    Select Case sExt
    Case "xlsx": FileFormat = xlWorkbookDefault
    Case "xlsm": FileFormat = xlOpenXMLWorkbookMacroEnabled
    Case "xlsb": FileFormat = 50
    Case "csv": FileFormat = xlCSV
    Case Else
      Call ErrorExit("Utils.SafeGetWorkbook:Unexpected file extension:" + CStr(sExt))
    End Select
    
    Call SafeGetWorkbook_SafeSaveAs(wb, sPath, FileFormat)
    
    If bReopenTemplate Then
      Application.ScreenUpdating = False
      On Error Resume Next
      Call Application.Workbooks.Open(sTemplate)
      If Not wbActive Is Nothing Then wbActive.Activate
      Application.ScreenUpdating = bScreenUpdating
    End If
  End If
End If

exit_now:

Set SafeGetWorkbook = wb

End Function



Function EndsWith(ByVal s As String, ByVal ending As String) As Boolean
    If Len(ending) > Len(s) Then
        EndsWith = False
    Else
        EndsWith = (Right$(s, Len(ending)) = ending)
    End If
End Function

Function Mondayised(d As Date) As Date
    Select Case Weekday(d, vbMonday)
        Case 6: Mondayised = d + 2   ' Sat ? Monday
        Case 7: Mondayised = d + 1   ' Sun ? Monday
        Case Else: Mondayised = d    ' Mon–Fri ? actual date
    End Select
End Function

Function FourthMonday(year As Long, month As Long) As Date
    Dim d As Date
    d = DateSerial(year, month, 1)
    
    ' Move to first Monday
    Do While Weekday(d, vbMonday) <> 1
        d = d + 1
    Loop
    
    ' Add 3 more Mondays ? fourth Monday
    FourthMonday = d + 21
End Function

Function EasterDate(Y As Long) As Date
    Dim a As Integer, b As Integer, c As Integer
    Dim d As Integer, e As Integer, f As Integer
    Dim g As Integer, h As Integer, i As Integer
    Dim k As Integer, L As Integer, m As Integer
    
    a = Y Mod 19
    b = Y \ 100
    c = Y Mod 100
    d = b \ 4
    e = b Mod 4
    f = (b + 8) \ 25
    g = (b - f + 1) \ 3
    h = (19 * a + b - d - g + 15) Mod 30
    i = c \ 4
    k = c Mod 4
    L = (32 + 2 * e + 2 * i - h - k) Mod 7
    m = (a + 11 * h + 22 * L) \ 451
    
    EasterDate = DateSerial(Y, (h + L - 7 * m + 114) \ 31, ((h + L - 7 * m + 114) Mod 31) + 1)
End Function

Function MatarikiDate(year As Long) As Date
    Select Case year
        Case 2022: MatarikiDate = #6/24/2022#
        Case 2023: MatarikiDate = #7/14/2023#
        Case 2024: MatarikiDate = #6/28/2024#
        Case 2025: MatarikiDate = #6/20/2025#
        Case 2026: MatarikiDate = #7/10/2026#
        Case 2027: MatarikiDate = #6/25/2027#
        Case 2028: MatarikiDate = #7/14/2028#
        Case 2029: MatarikiDate = #7/6/2029#
        Case 2030: MatarikiDate = #6/21/2030#
        Case Else:
            MatarikiDate = 0   ' Unknown future year — update as needed
    End Select
End Function

Function NZPublicHolidays(year As Long) As Collection
    Dim col As New Collection
    Dim d As Date
    
    ' --- Fixed-date holidays with Mondayisation ---
    col.Add Mondayised(DateSerial(year, 1, 1))   ' New Year’s Day
    col.Add Mondayised(DateSerial(year, 1, 2))   ' Day After New Year’s Day
    col.Add Mondayised(DateSerial(year, 2, 6))   ' Waitangi Day
    col.Add Mondayised(DateSerial(year, 4, 25))  ' ANZAC Day
    col.Add Mondayised(DateSerial(year, 6, 3))   ' King’s Birthday (first Monday in June)
    ' Christmas
    col.Add Mondayised(DateSerial(year, 12, 25)) ' Christmas Day
    col.Add Mondayised(DateSerial(year, 12, 26)) ' Boxing Day
    
    ' --- Easter (Good Friday + Easter Monday) ---
    Dim easter As Date
    easter = EasterDate(year)
    
    col.Add easter - 2   ' Good Friday
    col.Add easter + 1   ' Easter Monday
    
    ' --- Labour Day (4th Monday of October) ---
    col.Add FourthMonday(year, 10)
    
    ' --- Matariki ---
    Dim mata As Date
    mata = MatarikiDate(year)
    If mata <> 0 Then col.Add mata
    
    Set NZPublicHolidays = col
End Function

Function AddBusinessDays(startDate As Date, nDays As Long, holidays As Collection) As Date
    Dim d As Date
    Dim h As Variant
    
    d = startDate

    Do While nDays > 0
        d = d + 1
        
        ' Weekend?
        If Weekday(d, vbMonday) > 5 Then
            GoTo SkipDay
        End If
        
        ' Holiday?
        For Each h In holidays
            If d = h Then GoTo SkipDay
        Next h
        
        ' Valid business day
        nDays = nDays - 1
        
SkipDay:
    Loop

    AddBusinessDays = d
End Function

Function NextBusinessDay(d As Date, holidays As Collection) As Date
    Dim h As Variant
    
    Do
        ' Weekend?
        If Weekday(d, vbMonday) > 5 Then
            d = d + 1
            GoTo ContinueLoop
        End If
        
        ' Holiday?
        For Each h In holidays
            If d = h Then
                d = d + 1
                GoTo ContinueLoop
            End If
        Next h
        
        Exit Do   ' Found business day
        
ContinueLoop:
    Loop
    
    NextBusinessDay = d
End Function

Function InvoiceDueDate(invoiceDate As Date, termsDays As Long) As Date
    Dim targetDate As Date
    Dim holidays As Collection
    
    ' Build holidays for the year of invoice or due date
    Set holidays = NZPublicHolidays(year(invoiceDate))
    
    ' Add the term days directly (calendar days)
    targetDate = DateAdd("d", termsDays, invoiceDate)
    
    ' Round up to the next business day
    InvoiceDueDate = NextBusinessDay(targetDate, holidays)
End Function


' Convert any file (e.g. PNG or JPG) to Base64 text
Function FileToBase64(path As String) As String
    Dim stm As Object
    Dim xml As Object
    Dim bytes() As Byte

    Set stm = CreateObject("ADODB.Stream")
    stm.Type = 1          ' adTypeBinary
    stm.Open
    stm.LoadFromFile path
    bytes = stm.Read
    stm.Close

    Set xml = CreateObject("MSXML2.DOMDocument")
    With xml.createElement("b64")
        .DataType = "bin.base64"
        .NodeTypedValue = bytes
        FileToBase64 = .Text
    End With
End Function
