VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "ZipFile"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

Dim fso As New FileSystemObject
Dim objShell As Object
Dim copiesSinceLastSaveCount As Long
Dim origin_file As String
Dim source_file As String
Public ItemsByPath As Scripting.Dictionary
Public Paths As Collection

#If VBA7 Then ' Excel 2010 or later
Private Declare PtrSafe Sub Sleep Lib "kernel32" (ByVal Milliseconds As LongPtr)
#Else ' Excel 2007 or earlier
Private Declare PtrSafe Sub Sleep Lib "kernel32" (ByVal Milliseconds As Long)
#End If

Private Function ZipFileWorkspaceDir() As Variant
ZipFileWorkspaceDir = ThisWorkbook_Path + "\ZipFile"
If Not fso.FolderExists(ZipFileWorkspaceDir) Then
  Call fso.CreateFolder(ZipFileWorkspaceDir)
End If
End Function

Private Function DeletePath() As Variant
DeletePath = ZipFileWorkspaceDir + "\Delete"
If Not fso.FolderExists(DeletePath) Then
  Call fso.CreateFolder(DeletePath)
End If
End Function

Private Function TempPath() As Variant
TempPath = ZipFileWorkspaceDir + "\Temp"
If Not fso.FolderExists(TempPath) Then
  Call fso.CreateFolder(TempPath)
End If
End Function

Private Function BackPath() As Variant
BackPath = ZipFileWorkspaceDir + "\Back"
If Not fso.FolderExists(BackPath) Then
  Call fso.CreateFolder(BackPath)
End If
End Function

Private Sub SafeDeleteFileSystem(path)
If fso.FileExists(path) Then
  Call fso.DeleteFile(path, True)
ElseIf fso.FolderExists(path) Then
  Call fso.DeleteFolder(path, True)
End If
End Sub

Private Sub SafeDeleteZip(path)
If ItemsByPath.Exists(path) Then
  Dim f: Set f = ItemsByPath(path)
  Dim tmp: tmp = DeletePath
  Dim s: s = Split(path, "\")
  Call SafeDeleteFileSystem(tmp + "\" + s(UBound(s)))
  Call objShell.Namespace(tmp).MoveHere(f)
End If
End Sub

Private Function IsFileSystem(o) As Boolean
If TypeOf o Is Shell32.FolderItem Then
  IsFileSystem = o.IsFileSystem
Else
  IsFileSystem = o.Self.IsFileSystem
End If
End Function

Private Function IsFolder(o) As Boolean
If TypeOf o Is Shell32.FolderItem Then
  IsFolder = o.IsFolder
Else
  IsFolder = o.Self.IsFolder
End If
End Function

Private Function GetParentFolder(o) As Object
If TypeOf o Is Shell32.FolderItem Then
  Set GetParentFolder = o.Parent
Else
  Set GetParentFolder = o.ParentFolder
End If
End Function

Private Function GetName(o) As String
If TypeOf o Is Shell32.FolderItem Then
  GetName = o.name
Else
  GetName = o.Self.name
End If
End Function

Private Function GetPath(o) As String
If TypeOf o Is Shell32.FolderItem Then
  GetPath = o.path
Else
  GetPath = o.Self.path
End If
End Function

Private Function GetFolder(o) As Object
If TypeOf o Is Shell32.FolderItem Then
  Set GetFolder = o.GetFolder
Else
  Set GetFolder = o
End If
End Function

Private Sub RecurseNamespaceFolder(items_by_path As Scripting.Dictionary, p, Optional ByVal path As String = "")   '
' https://docs.microsoft.com/en-us/windows/win32/shell/objects
Dim f: Set f = GetFolder(p)
If path = "" Then
  Call items_by_path.Add("\", f) ' The root is called "\"
Else
  path = path + GetName(f)
  Call items_by_path.Add(path, f)
End If
path = path + "\"
Dim c
For Each c In f.Items
  If c.IsFolder Then
    Call RecurseNamespaceFolder(items_by_path, c, path:=path)
  Else
    Call items_by_path.Add(path + c.name, c)
  End If
Next c
End Sub

Function Show() As ZipFile
Dim fn
For Each fn In Paths
  Dim f: Set f = ItemsByPath(fn)
  'Call IsFolder(f)
  'Call GetPath(f)
  Debug.Print fn
Next fn
Set Show = Me
End Function

Function OpenFile(ByVal file_name As String) As ZipFile
' https://docs.microsoft.com/en-us/windows/win32/shell/objects
If Not fso.FileExists(file_name) Then
  If Not fso.FolderExists(file_name) Then
    Call ErrorExit("ZipFile.OpenFile:File does not exist:'" + file_name + "'")
  Else
    Call ErrorExit("ZipFile.OpenFile:Not expecting folders:'" + file_name + "'")
  End If
End If
origin_file = Trim(file_name)
If LCase(Right(origin_file, 4)) = ".zip" Then
  source_file = origin_file
Else
  source_file = TempPath + "\" + fso.GetBaseName(origin_file) + ".zip"
  Call fso.CopyFile(origin_file, source_file, True)
End If
Set ItemsByPath = New Scripting.Dictionary
ItemsByPath.CompareMode = TextCompare
Dim source_file_v As Variant: source_file_v = source_file
Dim root As Object: Set root = objShell.Namespace(source_file_v)
If root Is Nothing Then
  Call ErrorExit("ZipFile.OpenFile:Failed to open file:'" + source_file + "'")
End If
Call RecurseNamespaceFolder(ItemsByPath, root)
Set Paths = ShellSortCollectionOfString(ItemsByPath.Keys)
Set OpenFile = Me
End Function

Private Sub CopyHere(nsDst As Object, src)
' https://docs.microsoft.com/en-us/windows/win32/shell/objects
Dim count_init As Long: count_init = nsDst.Items.count
nsDst.CopyHere src
Dim beg As Date: beg = Now
Dim count As Long
Const sleep_milliseconds = 50
Const count_modulus = 1000 \ sleep_milliseconds
Dim src_name: src_name = Empty
count = -1
While nsDst.Items.count <= count_init
  DoEvents
  Sleep sleep_milliseconds
  count = count + 1
  If count Mod count_modulus = 0 Then
    Dim diff_seconds As Long: diff_seconds = DateDiff("s", beg, Now)
    If diff_seconds >= 15 And diff_seconds Mod 5 = 0 Then
    'If True Then
      If IsEmpty(src_name) Then
        If IsObject(src) Then
          src_name = GetName(src)
        Else
          src_name = src
        End If
      End If
      Dim ret: ret = MsgBox("Zip copy of:-" + vbLf + src_name + vbLf + "Into:-" + vbLf + GetPath(nsDst) + vbLf + "Has taken more than " + Str(diff_seconds) + " seconds." + vbLf + "Continue to wait?", vbYesNo + vbQuestion, "Zip Copy")
      If ret = vbNo Then
        End
      End If
    End If
  End If
Wend
End Sub

Sub SaveChanges()
If copiesSinceLastSaveCount > 0 And origin_file <> source_file Then
  Dim base As String, extn As String: base = fso.GetBaseName(origin_file): extn = fso.GetExtensionName(origin_file)
  base = BackPath + "\" + base + "."
  extn = "." + extn
  Dim count As Long: count = 0
  Do
    count = count + 1
    Dim back_file As String: back_file = base + format(count, "00") + extn
  Loop Until Not fso.FileExists(back_file)
  Call fso.CopyFile(origin_file, back_file, True)
  Call fso.CopyFile(source_file, origin_file, True)
End If
copiesSinceLastSaveCount = 0
End Sub

Sub CopyToZip(dst As String, src)
' https://docs.microsoft.com/en-us/windows/win32/shell/objects
If Not ItemsByPath.Exists(dst) Then
  Call ErrorExit("ZipFile.CopyToZip:Destination does not exist:'" + dst + "'")
End If

Dim fdst As Object: Set fdst = ItemsByPath(dst)
If Not IsFolder(fdst) Then
  Call ErrorExit("ZipFile.CopyToZip:Only expecting folders as destinations:'" + dst + "'")
End If

If IsObject(src) Then
  ' We're given a Namespace object ... copy it temporarily to the file system
  If IsFolder(src) Then
    Call ErrorExit("ZipFile.CopyToZip:Only expecting files as sources:'" + GetName(src) + "'")
  End If
  Dim tmp: tmp = TempPath + "\" + GetName(src)
  Call SafeDeleteFileSystem(tmp)
  Call CopyHere(objShell.Namespace(fso.GetParentFolderName(tmp)), src)
  src = tmp
End If

Dim fsrc
If fso.FileExists(src) Then
  Set fsrc = fso.GetFile(src)
ElseIf fso.FolderExists(src) Then
  'Set fsrc = fso.GetFolder(vsrc)
  Call ErrorExit("ZipFile.CopyToZip:Only expecting files as sources:'" + src + "'")
Else
  Call ErrorExit("ZipFile.CopyToZip:Source does not exist:'" + src + "'")
End If

' Make sure the destination doesn't already exist
Call SafeDeleteZip(IIf(dst = "\", "", dst) + "\" + fsrc.name)

' Run the copy to the zip file
Call CopyHere(fdst, src)

copiesSinceLastSaveCount = copiesSinceLastSaveCount + 1

End Sub

Private Sub Class_Initialize()
Set objShell = CreateObject("Shell.Application")
copiesSinceLastSaveCount = 0
End Sub

