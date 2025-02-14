Attribute VB_Name = "Updater"
Option Explicit

#If VBA7 Then
Private Declare PtrSafe Function SetCurrentDirectoryA Lib "kernel32" (ByVal lpPathName As String) As Long
#Else
Private Declare Function SetCurrentDirectoryA Lib "kernel32" (ByVal lpPathName As String) As Long
#End If

Sub SetUNCPath(sPath As String)
Dim lReturn As Long
lReturn = SetCurrentDirectoryA(sPath)
If lReturn = 0 Then MsgBox "Error setting path."
End Sub

Function SelectUpdateWorkbook(workbook_to_update_path As String) As String
Dim fso As New FileSystemObject
Dim Title As String: Title = "Select Update For '" + fso.GetFileName(workbook_to_update_path) + "'"
Dim ws As Object: Set ws = CreateObject("WScript.Shell")
Dim odr As String
On Error GoTo 10
odr = ws.RegRead("HKEY_CURRENT_USER\Environment\OneDrive")
GoTo 20
10:
odr = ""
20:
If odr <> "" Then
  SetUNCPath odr
Else
  SetUNCPath fso.GetParentFolderName(workbook_to_update_path)
End If
Dim base_name As String: base_name = fso.GetBaseName(workbook_to_update_path)
Dim ret: ret = Application.GetOpenFilename("Excel Macro-Enabled Workbook (*.xlsm),*.xlsm", , Title, , False)
If VarType(ret) = vbBoolean Then
  SelectUpdateWorkbook = vbNullString
ElseIf VarType(ret) <> vbString Then
  Call ErrorExit("VarType(ret) <> vbString")
Else
  SelectUpdateWorkbook = fso.GetAbsolutePathName(ret)
  If LCase(fso.GetAbsolutePathName(workbook_to_update_path)) = LCase(SelectUpdateWorkbook) Then
    Call ErrorExit("Can't update workbook with itself: " + SelectUpdateWorkbook)
  End If
End If
End Function

Sub UpdateExcelContent(src_path As String, dst_path As String, internal_files_to_copy As Collection)
Dim zfSrc As New ZipFile: Call zfSrc.OpenFile(src_path).Show
Dim zfDst As New ZipFile: Call zfDst.OpenFile(dst_path).Show
Dim fn
For Each fn In internal_files_to_copy
  Dim s: s = Split(fn, "\")
  ReDim Preserve s(LBound(s) To UBound(s) - 1)
  Dim v
  If IsObject(zfSrc.ItemsByPath(fn)) Then
    Set v = zfSrc.ItemsByPath(fn)
  Else
    v = zfSrc.ItemsByPath(fn)
  End If
  If IsEmpty(v) Then
    Dim ss As String: ss = "It seems like file extensions are not visible i.e., " + fn + vbCrLf + _
    "1. Start Windows Explorer, you can do this by opening up any folder." + vbCrLf + _
    "2. Click the View menu." + vbCrLf + _
    "3. Check the box next to ""File name Extensions""" + vbCrLf
    Call ErrorExit(ss)
  End If
  Call zfDst.CopyToZip(Join(s, "\"), v) '
Next fn
Call zfDst.SaveChanges
End Sub

Sub OldVersionOfUpdate(x) ' Use UpdateFromThisWorkbook instead
Dim fso As New FileSystemObject
Dim wsMaster As Worksheet: Set wsMaster = Safe_ThisWorkbook_Worksheets("Master")

Dim TARGET_WORKBOOK_r As Range: Set TARGET_WORKBOOK_r = wsMaster.UsedRange.Find("TARGET_WORKBOOK", LookAt:=xlWhole)
If TARGET_WORKBOOK_r Is Nothing Then
  wsMaster.Parent.Activate
  wsMaster.Activate
  Call ErrorExit("Missing cell for TARGET_WORKBOOK")
End If
Dim TARGET_WORKBOOK As String: TARGET_WORKBOOK = wsMaster.Cells(TARGET_WORKBOOK_r.row, TARGET_WORKBOOK_r.Column + 1)
Dim dst_path As String: dst_path = ThisWorkbook_Path + "\" + TARGET_WORKBOOK
If Not fso.FileExists(dst_path) Then
  wsMaster.Parent.Activate
  wsMaster.Activate
  Call ErrorExit("File TARGET_WORKBOOK='" + TARGET_WORKBOOK + "' does not exist in folder: '" + ThisWorkbook_Path + "'")
End If
Dim src_path As String: src_path = SelectUpdateWorkbook(dst_path)
If src_path <> vbNullString Then

  Dim UPDATE_FILES_r As Range: Set UPDATE_FILES_r = wsMaster.UsedRange.Find("UPDATE_FILES", LookAt:=xlWhole)
  If UPDATE_FILES_r Is Nothing Then
    wsMaster.Parent.Activate
    wsMaster.Activate
    Call ErrorExit("Missing cell for UPDATE_FILES")
  End If
  Dim UPDATE_FILES: UPDATE_FILES = wsMaster.Cells(UPDATE_FILES_r.row, UPDATE_FILES_r.Column + 1)
  UPDATE_FILES = Split(UPDATE_FILES, ";")
  Dim internal_files_to_copy As New Collection
  Dim i As Long
  For i = LBound(UPDATE_FILES) To UBound(UPDATE_FILES)
    internal_files_to_copy.Add Trim(UPDATE_FILES(i))
  Next i
  Call UpdateExcelContent(src_path, dst_path, internal_files_to_copy)
  MsgBox "Source: " + fso.GetAbsolutePathName(src_path) + vbLf + "Destination: " + fso.GetAbsolutePathName(dst_path) + vbLf + "Complete", vbInformation + vbOKOnly, "Excel Update"
End If
End Sub

Sub UpdateFromThisWorkbook(x)
Dim fso As New FileSystemObject
Dim wsMaster As Worksheet: Set wsMaster = Safe_ThisWorkbook_Worksheets("Master")
wsMaster.Parent.Activate
wsMaster.Activate

Dim Title As String: Title = "Select Workbook To Have Software Updated"
Dim dst_path: dst_path = Application.GetOpenFilename("Excel Macro-Enabled Workbook (*.xlsm),*.xlsm", , Title, , False)
If VarType(dst_path) = vbBoolean Then
  dst_path = vbNullString
ElseIf VarType(dst_path) <> vbString Then
  Call ErrorExit("VarType(dst_path) <> vbString")
End If

If dst_path <> vbNullString Then
  Dim base As String, extn As String: base = fso.GetBaseName(dst_path): extn = fso.GetExtensionName(dst_path)
  base = fso.GetParentFolderName(dst_path) + "\" + base + "."
  extn = "." + extn
  Dim count As Long: count = 0
  Do
    count = count + 1
    Dim src_path As String: src_path = base + format(count, "00") + extn
  Loop Until Not fso.FileExists(src_path)
  Call ThisWorkbook.SaveCopyAs(src_path)
  
  Dim UPDATE_FILES_r As Range: Set UPDATE_FILES_r = wsMaster.UsedRange.Find("UPDATE_FILES", LookAt:=xlWhole)
  If UPDATE_FILES_r Is Nothing Then
    wsMaster.Parent.Activate
    wsMaster.Activate
    Call ErrorExit("Missing cell for UPDATE_FILES")
  End If
  Dim UPDATE_FILES: UPDATE_FILES = wsMaster.Cells(UPDATE_FILES_r.row, UPDATE_FILES_r.Column + 1)
  UPDATE_FILES = Split(UPDATE_FILES, ";")
  Dim internal_files_to_copy As New Collection
  Dim i As Long
  For i = LBound(UPDATE_FILES) To UBound(UPDATE_FILES)
    internal_files_to_copy.Add Trim(UPDATE_FILES(i))
  Next i
  Call UpdateExcelContent(src_path, CStr(dst_path), internal_files_to_copy)
  Call fso.DeleteFile(src_path)
  MsgBox "Source: " + fso.GetAbsolutePathName(ThisWorkbook_FullName) + vbLf + "Destination: " + fso.GetAbsolutePathName(dst_path) + vbLf + "Complete", vbInformation + vbOKOnly, "Excel Update"
End If
End Sub


