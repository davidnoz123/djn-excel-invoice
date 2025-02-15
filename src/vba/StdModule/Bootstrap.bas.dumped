Attribute VB_Name = "Bootstrap"
Option Explicit

'Required references ...
'Microsoft Visual Basic for Applications Extensibility 5.3
'Microsoft Scripting Runtime
'Microsoft VBScript Regular Expressions 5.5

'''''''''''''''''''''''''''''''''''
' Used by FormatMessage
'''''''''''''''''''''''''''''''''''
Private Const FORMAT_MESSAGE_ALLOCATE_BUFFER As Long = &H100
Private Const FORMAT_MESSAGE_ARGUMENT_ARRAY  As Long = &H2000
Private Const FORMAT_MESSAGE_FROM_HMODULE  As Long = &H800
Private Const FORMAT_MESSAGE_FROM_STRING  As Long = &H400
Private Const FORMAT_MESSAGE_FROM_SYSTEM  As Long = &H1000
Private Const FORMAT_MESSAGE_MAX_WIDTH_MASK  As Long = &HFF
Private Const FORMAT_MESSAGE_IGNORE_INSERTS  As Long = &H200
Private Const FORMAT_MESSAGE_TEXT_LEN  As Long = &HA0 ' from VC++ ERRORS.H file

'''''''''''''''''''''''''''''''''''
' Windows API Declare
'''''''''''''''''''''''''''''''''''
#If Win64 Then
Private Declare PtrSafe Function FormatMessage Lib "kernel32" Alias "FormatMessageA" (ByVal dwFlags As Long, lpSource As Any, ByVal dwMessageId As Long, ByVal dwLanguageId As Long, ByVal lpBuffer As String, ByVal nSize As Long, Arguments As LongPtr) As Long
#Else

Private Declare Function FormatMessage Lib "kernel32" _
    Alias "FormatMessageA" ( _
    ByVal dwFlags As Long, _
    ByVal lpSource As Any, _
    ByVal dwMessageId As Long, _
    ByVal dwLanguageId As Long, _
    ByVal lpBuffer As String, _
    ByVal nSize As Long, _
    ByRef Arguments As Long) As Long
    
#End If

Private Enum eevbpatVBProjectAccessType
  evbpatRemove
  evbpatImport
  evbpatDelete
  evbpatString
End Enum


    
Sub Application_StatusBar(sMsg As String)
On Error Resume Next
Application.StatusBar = sMsg
End Sub

Private Function BrowseFolder(Title As String, _
        Optional InitialFolder As String = vbNullString, _
        Optional InitialView As Office.MsoFileDialogView = _
            msoFileDialogViewList) As String
    Dim v As Variant
    Dim InitFolder As String
    With Application.FileDialog(msoFileDialogFolderPicker)
        .Title = Title
        .InitialView = InitialView
        If Len(InitialFolder) > 0 Then
            If Dir(InitialFolder, vbDirectory) <> vbNullString Then
                InitFolder = InitialFolder
                If Right(InitFolder, 1) <> "\" Then
                    InitFolder = InitFolder & "\"
                End If
                .InitialFileName = InitFolder
            End If
        End If
        .Show
        On Error Resume Next
        Err.Clear
        v = .SelectedItems(1)
        If Err.Number <> 0 Then
            v = vbNullString
        End If
    End With
    BrowseFolder = CStr(v)
End Function

Sub CopyAllSheets(iDummy As Long) 'iDummy As Long
Dim wb As Workbook: Set wb = ThisWorkbook
Dim wbIn As Workbook: Set wbIn = Workbooks.Open("C:\analytics\projects\nbki\socnet\socnet20.xlsm", ReadOnly:=True)
Dim ws As Worksheet
For Each ws In wbIn.Sheets
  'Debug.Print ws.Name
  Call ws.Copy(After:=wb.Sheets(wb.Sheets.count))
Next ws
wbIn.Close
End Sub

Sub ErrorExit(msg As String)
'Call SafeActivateExcel
Call Logg("Error : " + Replace(msg, vbCrLf, " "))
Dim sHostName As String
sHostName = Environ$("computername")
Select Case sHostName
Case "ANZWL00001"
  'Debug.Assert False
Case Else
  
End Select
'Debug.Print iLbascas
'Debug.Print iUbascas
MsgBox msg, vbCritical, "Error"
End
End Sub


Sub Exit_(iDummy As Long)
End
End Sub

Function ExportVBComponent(sPath As String, oVBComponent As Object) As String
ExportVBComponent = ""
Dim sExt As String
sExt = VBAComponent_ExtFromType(oVBComponent.Type)
If sExt = "" Then

ElseIf oVBComponent.Type = 100 And oVBComponent.CodeModule.CountOfLines = 0 Then
  ' Don't output Worksheet code modules if they are empty
Else
  Dim sFileName As String
  sFileName = sPath + "\" + oVBComponent.name + "." + sExt
  oVBComponent.Export sFileName
  ExportVBComponent = sFileName
End If
End Function

Function GetNewSubFolder(sFolderPath As String) As String
Dim fso As New FileSystemObject
Dim dt As Date
dt = Now
Dim s As String
s = format(dt, "yyyy.mm.dd.hh.MM.ss")
GetNewSubFolder = fso.GetAbsolutePathName(sFolderPath + "\VbaBak." + s)
Call fso.CreateFolder(GetNewSubFolder)
End Function

Sub GetVbaBakFolderNames(wbTarg As Workbook, ByRef sBackupRoot As String, ByRef sBackupName As String, Optional sFolderNamePrefix As String = "VbaBak")
Dim iLoop As Long: iLoop = 0
sBackupRoot = wbTarg.path
sBackupName = sFolderNamePrefix + "." + wbTarg.name
Dim fso As New FileSystemObject
While fso.FolderExists(sBackupRoot + "\" + sBackupName + IIf(iLoop = 0, "", "." + format(iLoop, "000")))
  'DoEvents
  iLoop = iLoop + 1
Wend
sBackupName = sBackupName + IIf(iLoop = 0, "", "." + format(iLoop, "000"))
End Sub

Function GetVBAComponent(wb As Workbook, sName As String) As Object
On Error GoTo 10
Set GetVBAComponent = wb.VBProject.VBComponents(sName)
Exit Function
10:
Set GetVBAComponent = Nothing
End Function

Sub Logg(msg As String)
On Error Resume Next
Debug.Print msg
'Set Logg = Sheets("Logg")
Application.StatusBar = Left(Replace(msg, vbLf, " "), 256)
'Call Logg.Range("1:1").Insert
'Logg.Cells(1, 1) = Format(Now, "yyyy.mm.dd.hh.MM.ss")
'Logg.Cells(1, 2) = msg
End Sub

Sub RemoveVBACode(iDummy As Long) 'iDummy As Long
Dim vbComp As Object
Dim wb As Workbook
Set wb = ThisWorkbook 'Workbooks("FootballBetting02.xlsm")
For Each vbComp In wb.VBProject.VBComponents
  Select Case vbComp.Type
  Case 1 'vbext_ct_StdModule
    wb.VBProject.VBComponents.Remove vbComp
  Case 2 'vbext_ct_ClassModule
    wb.VBProject.VBComponents.Remove vbComp
  Case 3 'vbext_ct_MSForm
    wb.VBProject.VBComponents.Remove vbComp
  Case Else

  End Select
Next vbComp
End Sub

Sub VBAComponent_AddCode(wb As Workbook, eCodeModuleType As vbext_ComponentType, sCodeModuleName As String, sCode As String)
Dim vbc As VBComponent
For Each vbc In wb.VBProject.VBComponents
  If vbc.Type = eCodeModuleType And vbc.name = sCodeModuleName Then
    Call vbc.CodeModule.AddFromString(sCode)
  End If
Next vbc
End Sub

Function VBAComponent_ExtFromType(iType As Long) As String
Select Case iType
Case 1 'VBIDE.vbext_ComponentType
  VBAComponent_ExtFromType = "bas"
Case 2
  VBAComponent_ExtFromType = "cls"
Case 3
  VBAComponent_ExtFromType = "frm"
Case 100
  VBAComponent_ExtFromType = "cls" ' i.e., WorkSheet Class
Case Else
  VBAComponent_ExtFromType = ""
End Select
End Function

Function VBAComponent_TypeFromExt(sExt As String) As Long
Select Case LCase(Trim(sExt))
Case "bas"
  VBAComponent_TypeFromExt = 1
Case "cls"
  VBAComponent_TypeFromExt = 2
Case "frm"
  VBAComponent_TypeFromExt = 3
Case Else
  VBAComponent_TypeFromExt = 0
End Select
End Function

Sub Warn(msg As String)
Dim Logg As Worksheet
On Error Resume Next
Debug.Print "Warning:" + msg
'Set Logg = Sheets("Logg")
Application.StatusBar = "Warning:" + msg
'Call Logg.Range("1:1").Insert
'Logg.Cells(1, 1) = Format(Now, "yyyy.mm.dd.hh.MM.ss")
'Logg.Cells(1, 2) = msg
End Sub

Sub ExportVBACode(sFolderName As String, wb As Workbook)
Dim vbc As Object
For Each vbc In wb.VBProject.VBComponents
  Call ExportVBComponent(sFolderName, vbc)
Next vbc
End Sub

Sub zExportVBACode(x) 'iDummy As Long
Dim vbc As Object
Dim wb As Workbook
Set wb = ThisWorkbook
Dim sFolderName As String
sFolderName = BrowseFolder("ExportVBACode") ', wb.Path
If sFolderName <> "" Then
  Call ExportVBACode(sFolderName, wb)
End If
End Sub

Private Function StandardiseVBAComponentContent(ss As String) As String
Dim s As String: s = LCase(ss)
Dim iLenPrev As Long: iLenPrev = -1
If InStr(s, vbTab) > 0 Then
  Debug.Assert False
End If
While iLenPrev <> Len(s)
  iLenPrev = Len(s)
  s = Replace(s, "  ", " ")
  s = Replace(s, vbCrLf + vbCrLf, vbCrLf)
  s = Replace(s, vbCrLf + " " + vbCrLf, vbCrLf)
Wend
StandardiseVBAComponentContent = s
End Function

Private Function StandardiseVBAFileContent(sFile As String) As String
Dim fso As New FileSystemObject
If Not fso.FileExists(sFile) Then
  Debug.Assert False
End If
Dim ts As TextStream: Set ts = fso.OpenTextFile(sFile, ForReading)
Dim s As String: s = ts.ReadAll
ts.Close
StandardiseVBAFileContent = StandardiseVBAComponentContent(s)
End Function

Private Sub GetVBAComponentsDictionary(wb As Workbook, dicModulesByExt As Scripting.Dictionary, ByRef iCountComponents As Long)
Dim oVBComponent As VBComponent
Dim vba_ext As String
Set dicModulesByExt = New Scripting.Dictionary
iCountComponents = 0
For Each oVBComponent In wb.VBProject.VBComponents
  iCountComponents = iCountComponents + 1
  Dim dicModulesByName As Scripting.Dictionary
  vba_ext = VBAComponent_ExtFromType(oVBComponent.Type)
  If dicModulesByExt.Exists(vba_ext) Then
    Set dicModulesByName = dicModulesByExt(vba_ext)
  Else
    Set dicModulesByName = New Scripting.Dictionary
    Call dicModulesByExt.Add(vba_ext, dicModulesByName)
  End If
  dicModulesByName.Add oVBComponent.name, oVBComponent
Next oVBComponent
End Sub

Private Function ProgrammaticAccessAllowed() As Boolean
Dim c As VBIDE.VBComponent
On Error GoTo 10
Set c = ThisWorkbook.VBProject.VBComponents(1)
ProgrammaticAccessAllowed = True
Exit Function
10:
ProgrammaticAccessAllowed = False
End Function

Private Function SafeAddWorksheet(wb As Workbook, sWorksheetName As String) As Worksheet
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

Private Function SafeGetSheetByCodeName(wb As Workbook, sCodeName As String) As Object
Dim ws As Object
Dim i As Long
For i = 1 To wb.Sheets.count
  If StrComp(wb.Sheets(i).CodeName, sCodeName, vbTextCompare) = 0 Then
    ' We've got a Worksheet with this code name - we'll load the code into this
    Set ws = wb.Sheets(i)
    GoTo found
  End If
Next i
Dim iCount As Long: iCount = 0
Dim sName As String: sName = sCodeName
While True
  DoEvents
  On Error GoTo 10
  Set ws = wb.Worksheets(sName)
  GoTo name_exists
10:
  GoTo found_free_name
name_exists:
  iCount = iCount + 1
  sName = sCodeName + CStr(iCount)
Wend
found_free_name:
On Error GoTo 0
Set ws = Nothing
'Set ws = SafeAddWorksheet(wb, sName)
'If Not SafeSetCodeName(ws, sCodeName) Then
'  Debug.Assert False
'End If
found:
Set SafeGetSheetByCodeName = ws
End Function

Private Function SafeSetCodeName(ws As Object, sCodeName As String) As Boolean
Dim vbProj As VBIDE.VBProject
Dim vbComps As VBIDE.VBComponents
Dim vbComp As VBIDE.VBComponent
Dim vbProps As VBIDE.Properties
Dim CodeNameProp As VBIDE.Property
On Error GoTo 10
Set vbProj = ws.Parent.VBProject
Set vbComps = vbProj.VBComponents
Set vbComp = vbComps(ws.CodeName)
Set vbProps = vbComp.Properties
Set CodeNameProp = vbProps("_Codename")
CodeNameProp.Value = sCodeName
SafeSetCodeName = True
Exit Function
10:
SafeSetCodeName = False
End Function

Private Function SafeGetBool(v, vDefault)
On Error GoTo 10
SafeGetBool = CBool(v)
Exit Function
10:
Debug.Assert False
SafeGetBool = vDefault
End Function

Private Sub SafeSetBit(ByRef iValue As Long, v, ByVal iBit As Long)
If SafeGetBool(v, False) Then
  iValue = iValue Or iBit
Else
  iValue = iValue And (Not iBit)
End If
End Sub

Private Function ImportVBACode_AccessVBProject(ByVal eWhichAccess As eevbpatVBProjectAccessType, sVB_Name As String, colImportFailureMessages As Collection, wb As Workbook, oVBComponent As VBComponent, sVal As String) As Boolean
Dim sWhich As String
On Error GoTo 10
Select Case eWhichAccess
Case evbpatRemove
  sWhich = "Remove"
  Debug.Print sWhich + ":Beg:" + sVB_Name
  Call wb.VBProject.VBComponents.Remove(oVBComponent)
  Debug.Print sWhich + ":End:" + sVB_Name
Case evbpatImport
  sWhich = "Import"
  Debug.Print sWhich + ":Beg:" + sVB_Name
  Call wb.VBProject.VBComponents.Import(sVal)
  Debug.Print sWhich + ":End:" + sVB_Name
Case evbpatDelete
  sWhich = "Delete"
  Debug.Print sWhich + ":Beg:" + sVB_Name
  If oVBComponent.CodeModule.CountOfLines > 0 Then
    Call oVBComponent.CodeModule.DeleteLines(1, oVBComponent.CodeModule.CountOfLines)
  End If
  Debug.Print sWhich + ":End:" + sVB_Name
Case evbpatString
  sWhich = "String"
  Debug.Print sWhich + ":Beg:" + sVB_Name
  Call oVBComponent.CodeModule.AddFromString(sVal)
  Debug.Print sWhich + ":End:" + sVB_Name
Case Else
  Call ErrorExit("ImportVBACode_AccessVBProject:Unexpected access:" + CStr(eWhichAccess))
End Select
ImportVBACode_AccessVBProject = True
Exit Function
10:
Dim sErr As String: sErr = Err.Description
Debug.Print sWhich + ":Err:" + sVB_Name + ":" + sErr
ImportVBACode_AccessVBProject = False
colImportFailureMessages.Add sErr
End Function

Sub ImportVBACode(sFolderName As String, wb As Workbook, ByRef colImportFailureMessages As Collection)
Dim fso As New FileSystemObject

Set colImportFailureMessages = New Collection

Dim bResult As Boolean

If Not fso.FolderExists(sFolderName) Then
  Call ErrorExit("BootStrap.ImportVBACode:Missing folder:" + sFolderName)
Else
  Dim dicModulesByExt As New Scripting.Dictionary
  Dim oVBComponent As VBComponent
  Dim vba_ext As String
  
  Dim iCountOrig As Long
  
  Call GetVBAComponentsDictionary(wb, dicModulesByExt, iCountOrig)
        
  Dim cVbFiles As New Collection
  
  Dim folder_ As Folder
  Set folder_ = fso.GetFolder(sFolderName)
  Dim file_ As File
  For Each file_ In folder_.Files
    Dim sExt As String
    sExt = fso.GetExtensionName(file_.name)
    If VBAComponent_TypeFromExt(sExt) > 0 Then
      Call cVbFiles.Add(file_.path)
    End If
  Next file_
  
  Dim reMethodAttributesCommenting As New VBScript_RegExp_55.RegExp
  reMethodAttributesCommenting.Pattern = "(^|" + vbCrLf + ")" + "\s*Attribute\s+"
  reMethodAttributesCommenting.Global = True
  reMethodAttributesCommenting.IgnoreCase = True
  
  Dim re As New VBScript_RegExp_55.RegExp
  Const sAttributeLinePattern0 = "\s*Attribute\s+VB_Name*\s*=\s*[^" + vbCrLf + "]+\s*" + vbCrLf
  Const sAttributeLinePattern1 = "\s*Attribute\s+[A-Za-z][A-Z_a-z0-9]*\s*=\s*[^" + vbCrLf + "]+\s*" + vbCrLf
  re.Pattern = "(^|" + vbCrLf + ")" + sAttributeLinePattern0 + "(" + sAttributeLinePattern1 + ")*"
  're.MultiLine = True
  re.Global = True
  re.IgnoreCase = True

  Dim sBackupFolderName As String
  sBackupFolderName = ""

  Dim i As Long
  For i = 1 To cVbFiles.count
    Dim vbc As String: vbc = cVbFiles(i)
    Dim sNew As String: sNew = CStr(vbc)
    Dim sNewExt As String: sNewExt = LCase(fso.GetExtensionName(sNew))
    Dim sNewBase As String: sNewBase = fso.GetBaseName(sNew)
        
    ' Get the file contents
    Dim ts As TextStream: Set ts = fso.OpenTextFile(sNew)
    Dim codeRaw As String: codeRaw = ts.ReadAll
    ts.Close
    
    Dim mc As VBScript_RegExp_55.MatchCollection: Set mc = re.Execute(codeRaw)
    
    If mc.count = 0 Then
      ' We haven't found the expected "hidden" header in the file?
      Debug.Assert False
      GoTo skip_import_from_file
    Else
      ' Load the parameters from the "hidden" header
      Dim sVB_Name As String: sVB_Name = ""
      Const bitsVB_Creatable = 1
      Const bitsVB_GlobalNameSpace = 2
      Const bitsVB_Exposed = 4
      Const bitsVB_PredeclaredId = 8
      Dim iAssignedBooleanState As Long: iAssignedBooleanState = 0
      
      Dim mm As VBScript_RegExp_55.Match: Set mm = mc(0)
      Dim ssa: ssa = Split(mm.Value, vbCrLf)
      Dim j As Long
      For j = LBound(ssa) To UBound(ssa)
        Dim sAttributeLine As String: sAttributeLine = Replace(Replace(ssa(j), vbCrLf, ""), " ", "")
        If sAttributeLine <> "" Then
          If InStr(LCase(sAttributeLine), "attribute") <> 1 Then
            ' This should be impossible if the regexp etc. is working correctly
            Debug.Assert False
            GoTo skip_import_from_file
          Else
            sAttributeLine = Mid(sAttributeLine, Len("attribute") + 1, Len(sAttributeLine))
            Dim ss: ss = Split(sAttributeLine, "=")
            If UBound(ss) <> 1 Then
              ' We should have key value pairs?
              Debug.Assert False
              GoTo skip_import_from_file
            Else
              Dim LCase_ss_0 As String: LCase_ss_0 = LCase(ss(0))
              Select Case LCase_ss_0
              Case "vb_name":
                sVB_Name = Replace(ss(1), """", "")
              Case "vb_creatable":
                Call SafeSetBit(iAssignedBooleanState, ss(1), bitsVB_Creatable)
              Case "vb_globalnamespace":
                Call SafeSetBit(iAssignedBooleanState, ss(1), bitsVB_GlobalNameSpace)
              Case "vb_exposed":
                Call SafeSetBit(iAssignedBooleanState, ss(1), bitsVB_Exposed)
              Case "vb_predeclaredid":
                Call SafeSetBit(iAssignedBooleanState, ss(1), bitsVB_PredeclaredId)
              Case Else
                If InStr(LCase_ss_0, "vb_procdata.vb_invoke_func") > 0 Then
                
                Else
                  Debug.Assert False ' We haven't seen this attribute before - ensure we aren't missing anything by ignoring it
                End If
              End Select
            End If
          End If
        End If
      Next j
    
      If sVB_Name = "" Then
        ' We haven't found the component name in the "hidden" header??
        Debug.Assert False
        GoTo skip_import_from_file
      Else
        ' Remove the "hidden" header from the file content
        Dim codeTrm As String: codeTrm = Mid(codeRaw, mm.FirstIndex + mm.Length + 1, Len(codeRaw))
        
        codeTrm = reMethodAttributesCommenting.Replace(codeTrm, vbCrLf + "'Attribute ")
        
        Dim bIsBookObject As Boolean: bIsBookObject = False
        If sNewExt = "cls" Then
          Select Case iAssignedBooleanState
          Case 0, bitsVB_Exposed: ' The usual case of a class module
          Case bitsVB_Exposed + bitsVB_PredeclaredId:
            ' Do we have this case if and only if we have a worksheet class??
            If StrComp(sVB_Name, "ThisWorkbook", vbTextCompare) = 0 Then
              ' This is the ThisWorkbook object
              bIsBookObject = True
            Else
              Dim ws As Object: Set ws = SafeGetSheetByCodeName(wb, sVB_Name)
              If ws Is Nothing Then
                ' We haven't got a sheet module with this name - skip it to be safe
                Debug.Print "NoMatchingSheet:" + sVB_Name
                Debug.Assert False
                GoTo skip_import_from_file
              End If
              bIsBookObject = True
            End If
          Case Else
            Debug.Assert False ' Are we missing the handling of some other cases?
          End Select
        End If
        
        ' Do we already have a component with this name
        Set oVBComponent = GetVBAComponent(wb, sVB_Name)
        If oVBComponent Is Nothing Then
          If Not bIsBookObject Then
            ' We haven't got component with this name - just import from file
          Else
            ' This may or may not work ok - step through and see what can be done safely
            Debug.Assert False
            GoTo skip_import_from_file
          End If
        Else
          ' We have a name collision
          vba_ext = VBAComponent_ExtFromType(oVBComponent.Type)
          If vba_ext <> LCase(fso.GetExtensionName(sNew)) Then
            Call ErrorExit("BootStrap.zImportVBACode:Extensions clash for VBA components: " + sNew + " : " + oVBComponent.name + "." + vba_ext)
          Else
            If sBackupFolderName = "" Then
              sBackupFolderName = GetNewSubFolder(sFolderName)
            End If
            Dim sBak As String: sBak = sBackupFolderName + "\" + oVBComponent.name + "." + vba_ext
            Call oVBComponent.Export(sBak)
            Dim sBakCode As String: sBakCode = StandardiseVBAFileContent(sBak)
            Dim sNewCode As String: sNewCode = StandardiseVBAFileContent(sNew)
            If sBakCode = sNewCode Then
              ' The code is the same - don't bother loading it
              GoTo skip_import_from_file
            Else
              If oVBComponent.name = "BootStrap" Then
                ' Skip loading changes to this file
                Debug.Assert False
                GoTo skip_import_from_file
              End If
              bResult = ImportVBACode_AccessVBProject(evbpatDelete, sVB_Name, colImportFailureMessages, wb, oVBComponent, "")
              If Not bResult Then
                bResult = ImportVBACode_AccessVBProject(evbpatRemove, sVB_Name, colImportFailureMessages, wb, oVBComponent, "")
                If Not bResult Then
                  ' Can't delete the code or remove the component - ignore it
                  Debug.Assert False
                  GoTo skip_import_from_file
                End If
              Else
                bResult = ImportVBACode_AccessVBProject(evbpatString, sVB_Name, colImportFailureMessages, wb, oVBComponent, codeTrm)
                If bResult Then
                  ' We're successfully finished for this component
                  GoTo skip_import_from_file
                Else
                  ' Failed to load the code as a string - try removing the component and importing it from file
                  bResult = ImportVBACode_AccessVBProject(evbpatRemove, sVB_Name, colImportFailureMessages, wb, oVBComponent, "")
                  If Not bResult Then
                    Debug.Assert False
                    GoTo skip_import_from_file
                  End If
                End If
              End If
            End If
          End If
        End If
      End If
    End If
    ' Do an import of the file
    bResult = ImportVBACode_AccessVBProject(evbpatImport, sVB_Name, colImportFailureMessages, wb, oVBComponent, sNew)
    If Not bResult Then
      Debug.Assert False
    End If
skip_import_from_file:
  Next i
End If

Debug.Print "Complete"

End Sub

Sub zImportVBACode(x) 'iDummy As Long
Dim sFolderName As String
sFolderName = BrowseFolder("ImportVBACode") '"C:\analytics\projects\fonterra\analytics\vba\smartprice.78.b.xlsm" '"C:\analytics\projects\fonterra\analytics\vba\smartprice.78.b.xlsm" ' BrowseFolder("ImportVBACode") ' BrowseFolder("ImportVBACode") ' - bRefreshFromFileOnly = " + CStr(bRefreshFromFileOnly)) ', wb.Path
' "C:\analytics\projects\powerop\vba\XLDATA\Data421DM\03_WrapObjc\"
' "C:\analytics\projects\fonterra\analytics\vba\smartprice.66.a.xlsm\"
Dim colImportFailureMessages As Collection
Call ImportVBACode(sFolderName, ThisWorkbook, colImportFailureMessages)
If colImportFailureMessages.count > 0 Then
  Debug.Assert False
End If
End Sub

Sub zPrintVBACode(iDummy As Long) 'iDummy As Long
Dim wb As Workbook
Set wb = ThisWorkbook
'Dim cmc As New CodeModuleCollection
'Call cmc.Add(wb)
'Dim ws As Worksheet
'Set ws = Sheets("ModuleList")
'ws.UsedRange.Clear
'Call cmc.DumpToSheet(ws, 1, 1)
End Sub



Public Function GetSystemErrorMessageText(ErrorNumber As Long) As String
''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
' GetSystemErrorMessageText
'
' This function gets the system error message text that corresponds
' to the error code parameter ErrorNumber. This value is the value returned
' by Err.LastDLLError or by GetLastError, or occasionally as the returned
' result of a Windows API function.
'
' These are NOT the error numbers returned by Err.Number (for these
' errors, use Err.Description to get the description of the error).
'
' In general, you should use Err.LastDllError rather than GetLastError
' because under some circumstances the value of GetLastError will be
' reset to 0 before the value is returned to VBA. Err.LastDllError will
' always reliably return the last error number raised in an API function.
'
' The function returns vbNullString is an error occurred or if there is
' no error text for the specified error number.
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''

Dim ErrorText As String
Dim TextLen As Long
Dim FormatMessageResult As Long
Dim LangID As Long

''''''''''''''''''''''''''''''''
' Initialize the variables
''''''''''''''''''''''''''''''''
LangID = 0&   ' Default language
ErrorText = String$(FORMAT_MESSAGE_TEXT_LEN, vbNullChar)
TextLen = FORMAT_MESSAGE_TEXT_LEN

''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
' Call FormatMessage to get the text of the error message text
' associated with ErrorNumber.
''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
FormatMessageResult = FormatMessage( _
                        dwFlags:=FORMAT_MESSAGE_FROM_SYSTEM Or _
                                 FORMAT_MESSAGE_IGNORE_INSERTS, _
                        lpSource:=0&, _
                        dwMessageId:=ErrorNumber, _
                        dwLanguageId:=LangID, _
                        lpBuffer:=ErrorText, _
                        nSize:=TextLen, _
                        Arguments:=0&)

If FormatMessageResult = 0& Then
    ''''''''''''''''''''''''''''''''''''''''''''''''''
    ' An error occured. Display the error number, but
    ' don't call GetSystemErrorMessageText to get the
    ' text, which would likely cause the error again,
    ' getting us into a loop.
    ''''''''''''''''''''''''''''''''''''''''''''''''''
    MsgBox "An error occurred with the FormatMessage" & _
           " API function call." & vbCrLf & _
           "Error: " & CStr(Err.LastDllError) & _
           " Hex(" & Hex(Err.LastDllError) & ")."
    GetSystemErrorMessageText = "An internal system error occurred with the" & vbCrLf & _
        "FormatMessage API function: " & CStr(Err.LastDllError) & ". No futher information" & vbCrLf & _
        "is available."
    Exit Function
End If
''''''''''''''''''''''''''''''''''''''''''''''''''''''
' If FormatMessageResult is not zero, it is the number
' of characters placed in the ErrorText variable.
' Take the left FormatMessageResult characters and
' return that text.
''''''''''''''''''''''''''''''''''''''''''''''''''''''
ErrorText = Left$(ErrorText, FormatMessageResult)
'''''''''''''''''''''''''''''''''''''''''''''
' Get rid of the trailing vbCrLf, if present.
'''''''''''''''''''''''''''''''''''''''''''''
If Len(ErrorText) >= 2 Then
    If Right$(ErrorText, 2) = vbCrLf Then
        ErrorText = Left$(ErrorText, Len(ErrorText) - 2)
    End If
End If

''''''''''''''''''''''''''''''''
' Return the error text as the
' result.
''''''''''''''''''''''''''''''''
GetSystemErrorMessageText = ErrorText

End Function











