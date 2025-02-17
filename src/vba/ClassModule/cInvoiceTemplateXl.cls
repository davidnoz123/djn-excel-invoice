VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "cInvoiceTemplateXl"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

Const INVOICE_TABLE_TEMPLATES_XLS = "invoice_templates.xlsx"

' https://create.microsoft.com/en-us/templates/invoices
' https://create.microsoft.com/en-us/search?filters=invoices
Dim fso As New FileSystemObject
Dim ws_tmp As Worksheet

Function SafeGetTemplateWorkbook() As Workbook
Dim wb As Workbook
For Each wb In Workbooks
  If INVOICE_TABLE_TEMPLATES_XLS = wb.name Then
    Set SafeGetTemplateWorkbook = wb
    GoTo 10
  End If
Next wb
Dim template_file As String: template_file = fso.GetAbsolutePathName(ThisWorkbook.Path + "\" + INVOICE_TABLE_TEMPLATES_XLS)
If Not fso.FileExists(template_file) Then
    Call ErrorExit("Missing Invoice Templates Excel Workbook:'" + template_file + "'")
End If
Dim template_base As String: template_base = fso.GetFileName(template_file)
Set SafeGetTemplateWorkbook = Workbooks.Open(template_file, ReadOnly:=True)
10:
End Function

Sub Init(template_name As String)
Dim wb As Workbook: Set wb = SafeGetTemplateWorkbook
If Not WsExists(wb, template_name) Then
    Call ErrorExit("Missing worksheet in Invoice Templates Excel Workbook:'" + template_name + "'")
End If
Set ws_tmp = wb.Worksheets(template_name)
End Sub

Sub SearchAndReplace(ws As Worksheet, search_and_replace As Scripting.Dictionary)
Dim k
For Each k In search_and_replace.Keys()
  MainForm.Do_Events "SearchAndReplace " + search_and_replace(k) + " ..."
  ws.Cells.Replace What:=k, Replacement:=search_and_replace(k), LookAt:=xlPart, SearchOrder:=xlByRows, MatchCase:=False, SearchFormat:=False, ReplaceFormat:=False, FormulaVersion:=xlReplaceFormula2
Next k
MainForm.Do_Events "SearchAndReplace Complete"
End Sub

Function StringifyInvoiceContactDetail(s As String, Optional prefix As String = "")
Dim ret As String: ret = Trim(s)
If ret <> "" Then
  ret = prefix + ret + vbCrLf
End If
StringifyInvoiceContactDetail = ret
End Function

Private Sub ReplaceTokens(ws As Worksheet, master_record As cMainRecord, invoice_number As Long, date_ As Date)
' Substitute the variable strings in the document
Dim search_and_replace As New Scripting.Dictionary
search_and_replace.CompareMode = TextCompare
search_and_replace.Add "{INVOICE NUMBER}", format(invoice_number, "0000000")
search_and_replace.Add "{INVOICE DATE}", format(date_, "")
search_and_replace.Add "{Invoiced Contact Name}", StringifyInvoiceContactDetail(master_record.RecipientNameInvoice)
search_and_replace.Add "{Invoiced Company Name}", StringifyInvoiceContactDetail(master_record.CompanyName)
search_and_replace.Add "{Invoiced Street Address}", StringifyInvoiceContactDetail(master_record.StreetAddress)
search_and_replace.Add "{Invoiced Town, City Post Code}", StringifyInvoiceContactDetail(master_record.TownCityPostCode)
search_and_replace.Add "{Invoiced Country}", StringifyInvoiceContactDetail(master_record.Country)
search_and_replace.Add "{Invoiced Phone}", StringifyInvoiceContactDetail(master_record.Phone, "Phone: ")
search_and_replace.Add "{Invoiced Email}", StringifyInvoiceContactDetail(master_record.EmailAddressx, "Email: ")
Call SearchAndReplace(ws, search_and_replace)
End Sub

Function CreateInvoiceDocument(transColl As Collection, invoice_number As Long, date_ As Date) As Worksheet
Dim bScreenUpdating As Boolean: bScreenUpdating = Application.ScreenUpdating
On Error GoTo 10
Application.ScreenUpdating = False
ws_tmp.Copy After:=ws_tmp
Dim ws_out As Worksheet: Set ws_out = ws_tmp.Parent.Sheets(ws_tmp.index + 1)

Dim master_record As cMainRecord: Set master_record = transColl(1).MainRecord
Call ReplaceTokens(ws_out, master_record, invoice_number, date_)

GoTo 20
10:
Application.ScreenUpdating = bScreenUpdating
Call ErrorExit("Failure copying template worksheet")
20:
Application.ScreenUpdating = bScreenUpdating
Set CreateInvoiceDocument = ws_out
End Function


