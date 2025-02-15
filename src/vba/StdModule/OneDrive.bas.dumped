Attribute VB_Name = "OneDrive"
Option Explicit

Dim Global_ThisWorkbook_Path As Variant
Dim Global_OneDrive_password As String

Function DownloadURL(ByVal url As String, ByVal user_name As String, ByVal password As String, ByVal destination_file As String, ByRef statusText As String) As Long
If InStr(LCase(url), "http") <> 1 Then
  Call ErrorExit("URL does not look correct:'" + url + "'")
End If
Dim WinHttpReq As Object: Set WinHttpReq = CreateObject("Microsoft.XMLHTTP")
WinHttpReq.Open "GET", url, False, user_name, password
On Error GoTo 10
WinHttpReq.Send
DownloadURL = WinHttpReq.status
statusText = WinHttpReq.statusText
If WinHttpReq.status = 200 Then
  Dim oStream: Set oStream = CreateObject("ADODB.Stream")
  oStream.Open
  oStream.Type = 1
  oStream.Write WinHttpReq.responseBody
  oStream.SaveToFile destination_file, 2
  oStream.Close
End If
Exit Function
10:
Call ErrorExit("Failure downloading url '" + url + "':'" + Err.Description + "':'" + destination_file + "'")
End Function

Function ThisWorkbook_Path(Optional dummy As Variant = Empty) As String
Dim tmp As String: tmp = ThisWorkbook.path
Dim fso As Object
If InStr(LCase(tmp), "http") <> 1 Then
  Set fso = New Scripting.FileSystemObject
  If Not fso.FolderExists(tmp) Then
    Call ErrorExit("FolderDoesNotExist:'" + tmp + "'")
  End If
  ThisWorkbook_Path = tmp
Else
  If IsEmpty(Global_ThisWorkbook_Path) Or VarType(Global_ThisWorkbook_Path) <> vbString Then
    Set fso = New Scripting.FileSystemObject
    Dim name As String: name = ThisWorkbook.name
    Global_ThisWorkbook_Path = Empty
    Dim ws As Object: Set ws = CreateObject("WScript.Shell")
    Dim odr As String: odr = ws.RegRead("HKEY_CURRENT_USER\Environment\OneDrive")
    If odr = "" Then
      ' ERROR
    Else
      Dim a: a = Split(tmp, "/")
      Dim i As Long, j As Long
      For i = UBound(a) To LBound(a) Step -1
        ReDim aa(i To UBound(a))
        For j = i To UBound(aa)
          aa(j) = a(j)
        Next j
        Dim path As String: path = odr + "\" + Replace(Join(aa, "\"), "%20", " ")
        Dim fn As String: fn = path + "\" + name
        If fso.FileExists(fn) Then
          Global_ThisWorkbook_Path = path
          Exit For
        End If
      Next i
      If IsEmpty(Global_ThisWorkbook_Path) Then
        ' ERROR
      End If
    End If
  End If
  ThisWorkbook_Path = Global_ThisWorkbook_Path
End If
End Function

Function ThisWorkbook_FullName(Optional dummy As Variant = Empty) As String
ThisWorkbook_FullName = ThisWorkbook_Path + "\" + ThisWorkbook.name
End Function

Sub xThisWorkbook_FullName_Test(x)
Global_ThisWorkbook_Path = Empty
Debug.Print ThisWorkbook_FullName()
End Sub

Sub CopyWorksheets(dst As Workbook, src As Workbook)
If dst.name = src.name Then
  Call ErrorExit("dst.name = src.name:'" + dst.name + "'")
End If
Dim ws_src As Worksheet, ws_dst As Worksheet
For Each ws_src In src.Worksheets
  If WsExists(dst, ws_src.name) Then
    Set ws_dst = dst.Worksheets(ws_src.name)
    ws_dst.UsedRange.ClearContents
    Dim a: a = GetData(ws_src.UsedRange.Cells(1), ws_src.UsedRange.Rows.count, ws_src.UsedRange.Columns.count)
    Call PutData(ws_dst.Cells(ws_src.UsedRange.Rows(1).row, ws_src.UsedRange.Columns(1).Column), a)
  End If
Next ws_src
End Sub


Sub RefreshFromOneDrive()
Dim url As String: url = ThisWorkbook.FullName ' Get the OneDrive URL for this workbook

If InStr(LCase(url), "http") <> 1 Then
  Exit Sub
End If

Dim wb As Workbook
Dim password As String: password = InputBox("Enter OneDrive password to refresh Workbook" + vbLf + "Cancel to skip", "", "")

If password = "" Then
  Exit Sub
End If

Logg "Getting Office Account Email..."
Dim user_name As String: user_name = OfficeAccountEmail
Logg "Getting Office Account Email:" + user_name

Dim destination_base As String: destination_base = "OneDrive_download_" + ThisWorkbook.name

If WorkbookIsOpen(destination_base, wb) Then
  Logg "Closing Workbook:" + destination_base + " ..."
  Call wb.Close(SaveChanges:=False)
  Logg "Closing Workbook:" + destination_base + " Complete"
End If

Dim destination_file As String: destination_file = Environ("TEMP") + "\" + destination_base

Dim fso As New FileSystemObject
If fso.FileExists(destination_file) Then
  Logg "Deleting old Workbook:" + destination_file + " ..."
  Call fso.DeleteFile(destination_file, True)
  Logg "Deleting old Workbook:" + destination_file + " Complete"
End If

Dim statusText As String
Logg "Downloading Workbook:" + destination_file + " ..."
Dim status As Long: status = DownloadURL(url, user_name, password, destination_file, statusText)
If status <> 200 Then
  Call ErrorExit("Failure downloading this workbook from OneDrive:" + CStr(status) + ":'" + statusText + "'")
End If
Logg "Downloading Workbook:" + destination_file + " Complete"

Logg "Backing up ThisWorkbook: ..."
BackUpThisWorkbook
Logg "Backing up ThisWorkbook: Complete"

Dim Application_EnableEvents As Boolean: Application_EnableEvents = Application.EnableEvents
Application.EnableEvents = False

Logg "Opening Workbook:" + destination_file + ":" + CStr(Application_EnableEvents) + " ..."
Set wb = Workbooks.Open(destination_file, ReadOnly:=True)
Logg "Opening Workbook:" + destination_file + ":" + CStr(Application_EnableEvents) + " Complete"

Call CopyWorksheets(ThisWorkbook, wb)

Logg "Closing Workbook:" + destination_file + " ..."
Call wb.Close(SaveChanges:=False)
Logg "Closing Workbook:" + destination_file + " Complete"

Application.EnableEvents = Application_EnableEvents

If fso.FileExists(destination_file) Then
  Logg "Deleting new Workbook:" + destination_file + " ..."
  Call fso.DeleteFile(destination_file, True)
  Logg "Deleting new Workbook:" + destination_file + " Complete"
End If

Call MsgBox("Refresh complete", vbOKOnly + vbInformation)

End Sub

