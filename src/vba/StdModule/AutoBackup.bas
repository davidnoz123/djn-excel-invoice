Attribute VB_Name = "AutoBackup"
Option Explicit

Dim earliesttime As Double
Dim fso As New FileSystemObject
Dim last_backup_file As Scripting.File
Dim change_count As Long
Dim change_count_last_backup_file As Long

Const BACKUP_SUBDIR = "\backup"
Const DAYS_PER_SECOND = 1# / (24# * 60# * 60#)
Const DELAY_SECONDS_DEFAULT = 60# ' 2# 60#
Const MAX_MINUTES_BETWEEN_UNSAVED_WORKBOOK_AND_LAST_BACKUP = 15#   ' 0.1 15#
Const MAX_MINUTES_BETWEEN_SAVES = MAX_MINUTES_BETWEEN_UNSAVED_WORKBOOK_AND_LAST_BACKUP / 5#


Function ThisWorkbook_Path() As String
ThisWorkbook_Path = ThisWorkbook.path
End Function
Function ThisWorkbook_FullName() As String
ThisWorkbook_FullName = ThisWorkbook_Path + "\" + ThisWorkbook.name
End Function

Sub SignalChange(x)
change_count = change_count + 1
End Sub

Function LastSavedTimeStamp() As Date
LastSavedTimeStamp = ActiveWorkbook.BuiltinDocumentProperties("Last Save Time")
End Function

Sub Restart(x)
earliesttime = Now + DELAY_SECONDS_DEFAULT * DAYS_PER_SECOND
Application.OnTime earliesttime, "xRunAutoBackup"
End Sub

Sub Terminate(Optional caller As String = "None")
Err.Clear
On Error GoTo exit_with_failure
Application.OnTime earliesttime:=earliesttime, procedure:="xRunAutoBackup", schedule:=False
Exit Sub
exit_with_failure:
'Debug.Print caller, Err.Description
End Sub

Sub ThisWorkbook_SaveCopyAs(file_name)
Err.Clear
On Error GoTo exit_with_failure
Call ThisWorkbook.SaveCopyAs(file_name)
Exit Sub
exit_with_failure:
Call MsgBox("Failure backing up workbook " + ThisWorkbook.name + " to '" + file_name + "':" + Err.Description, vbOKOnly + vbExclamation, "Auto Backup")
End
End Sub

Function BackUpThisWorkbook() As String
Dim tss As String: tss = format(Now(), "yyyy.mm.dd.HH.MM.ss")
Dim backup_dir As String:  backup_dir = ThisWorkbook_Path + BACKUP_SUBDIR
If Not fso.FolderExists(backup_dir) Then
  Call fso.CreateFolder(backup_dir)
End If
Dim bak As String: bak = backup_dir + "\" + tss + "." + fso.GetFileName(ThisWorkbook_FullName)
Call ThisWorkbook_SaveCopyAs(bak)
BackUpThisWorkbook = bak
End Function

Sub xTerminate()
Terminate "xTerminate"
End Sub

Sub xRunAutoBackup()
'Exit Sub
Call Terminate("xRunAutoBackup")
Dim backup_dir As String:  backup_dir = ThisWorkbook_Path + BACKUP_SUBDIR
Err.Clear
On Error GoTo exit_with_start
If Not fso.FolderExists(backup_dir) Then
  Call fso.CreateFolder(backup_dir)
End If
Dim is_first_call As Boolean: is_first_call = False
If last_backup_file Is Nothing Then
  is_first_call = True
  Dim rgx As New VBScript_RegExp_55.RegExp
  rgx.Pattern = "([0-9]{4}\.[0-9]{2}\.[0-9]{2}\.[0-9]{2}\.[0-9]{2}\.[0-9]{2})\." + fso.GetFileName(ThisWorkbook_FullName)
  rgx.IgnoreCase = True
  Dim ff As Scripting.File
  For Each ff In fso.GetFolder(backup_dir).Files
    Dim mc As VBScript_RegExp_55.MatchCollection
    Set mc = rgx.Execute(ff.name)
    If Not mc Is Nothing Then
      If last_backup_file Is Nothing Then
        Set last_backup_file = ff
      ElseIf StrComp(last_backup_file.name, ff.name, vbTextCompare) < 0 Then
        Set last_backup_file = ff
      End If
    End If
  Next ff
End If
Dim tss As String: tss = format(Now(), "yyyy.mm.dd.HH.MM.ss")
Dim log_prefix As String: log_prefix = "xRunAutoBackup." + tss + ":"
If Not is_first_call And change_count_last_backup_file = change_count Then
  ' This is not the first call and we haven't seen any changes to the worksheets since the last time we
  Debug.Print log_prefix + "SkipS"
Else
  ' The workbook is now unsaved or this is the first call here
  Dim bRunBackup As Boolean: bRunBackup = False
  If last_backup_file Is Nothing Then
    ' There are no backup files
    bRunBackup = True
  ElseIf last_backup_file.DateLastModified < Now() - MAX_MINUTES_BETWEEN_UNSAVED_WORKBOOK_AND_LAST_BACKUP * 60# * DAYS_PER_SECOND Then
    ' We have gone long enough between backups
    bRunBackup = True
  End If
  If Not bRunBackup Then
    Debug.Print log_prefix + "SkipT"
  Else
    Debug.Print log_prefix + "Runn"
    change_count_last_backup_file = change_count
    Set last_backup_file = fso.GetFile(BackUpThisWorkbook)
  End If
End If

exit_with_start:
If Err.Description <> "" Then
  Debug.Print log_prefix + Err.Description
  If Err.Description = "File not found" Then
    Set last_backup_file = Nothing
  End If
End If
Call Restart(0) ' Schedule to run again
End Sub


