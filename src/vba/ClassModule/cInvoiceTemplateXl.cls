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

' https://create.microsoft.com/en-us/templates/invoices
' https://create.microsoft.com/en-us/search?filters=invoices

Const INVOICE_TABLE_TEMPLATES_XLS = "invoice_templates.xlsx"

Dim fso As New FileSystemObject
Dim ws_tmp As Worksheet
Dim TransactionColumn2InvoiceHeaderColumn As Scripting.Dictionary

Function SafeGetTemplateWorkbook() As Workbook
If True Then
    Set SafeGetTemplateWorkbook = ThisWorkbook
Else
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
End If
End Function

Sub Init(main_ As cMainData, template_name As String)
Dim wb As Workbook: Set wb = SafeGetTemplateWorkbook
If Not WsExists(wb, template_name) Then
    Call ErrorExit("Missing worksheet in Invoice Templates Excel Workbook:'" + template_name + "'")
End If
Set ws_tmp = wb.Worksheets(template_name)
If ws_tmp.ListObjects.count <> 1 Then
    ws_tmp.Parent.Activate
    ws_tmp.Activate
    Call ErrorExit("Invoice Template Worksheet missing the table used for transaction records:'" + template_name + "'")
End If
Set TransactionColumn2InvoiceHeaderColumn = New Scripting.Dictionary
Dim lo As ListObject: Set lo = ws_tmp.ListObjects(1)
Dim c As Range
For Each c In lo.HeaderRowRange
    If main_.INVOICE_COLUMN_2_TRANSACTION_COLUMN_MAP.Exists(c.Text) Then
        Dim tc As String: tc = main_.INVOICE_COLUMN_2_TRANSACTION_COLUMN_MAP(c.Text)
        Call TransactionColumn2InvoiceHeaderColumn.Add(tc, c.Column)
    End If
Next c

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

Private Sub ReplaceTokens(ws As Worksheet, main_record As cMainRecord, invoice_number As Long, date_ As Date)
' Substitute the variable strings in the document
Dim search_and_replace As New Scripting.Dictionary
search_and_replace.CompareMode = TextCompare
search_and_replace.Add "{INVOICE NUMBER}", format(invoice_number, "0000000")
search_and_replace.Add "{INVOICE DATE}", format(date_, "")
search_and_replace.Add "{Invoiced Contact Name}", StringifyInvoiceContactDetail(main_record.RecipientNameInvoice)
search_and_replace.Add "{Invoiced Company Name}", StringifyInvoiceContactDetail(main_record.CompanyName)
search_and_replace.Add "{Invoiced Street Address}", StringifyInvoiceContactDetail(main_record.StreetAddress)
search_and_replace.Add "{Invoiced Town, City Post Code}", StringifyInvoiceContactDetail(main_record.TownCityPostCode)
search_and_replace.Add "{Invoiced Country}", StringifyInvoiceContactDetail(main_record.Country)
search_and_replace.Add "{Invoiced Phone}", StringifyInvoiceContactDetail(main_record.Phone, "Phone: ")
search_and_replace.Add "{Invoiced Email}", StringifyInvoiceContactDetail(main_record.EmailAddressx, "Email: ")
Call SearchAndReplace(ws, search_and_replace)
End Sub

Private Function ReshapeTableAndInvoice(lo As ListObject, rows_expect As Long) As Long

Dim ws_out As Worksheet: Set ws_out = lo.Parent

Dim row_lot As Long, col_lol As Long, row_lob As Long, col_lor As Long
row_lot = lo.DataBodyRange.row
col_lol = lo.DataBodyRange.Column
row_lob = row_lot + lo.DataBodyRange.Rows.count - 1
col_lor = col_lol + lo.DataBodyRange.Columns.count - 1

ReshapeTableAndInvoice = row_lot

Dim row_top  As Long, row_bot As Long, col_lhs As Long, col_rhs As Long
row_top = ws_out.UsedRange.Rows(1).row
col_lhs = ws_out.UsedRange.Columns(1).Column
row_bot = ws_out.UsedRange.Rows(ws_out.UsedRange.Rows.count).row
col_rhs = ws_out.UsedRange.Columns(ws_out.UsedRange.Columns.count).Column

ReDim row_heights(row_top To row_bot) As Double

Dim k As Long
For k = row_top To row_bot
    row_heights(k) = ws_out.Rows(k).Height
Next k

Dim rows_actual As Long: rows_actual = lo.DataBodyRange.Rows.count
If rows_actual <> rows_expect Then
    Dim rr As Range
    If rows_actual > rows_expect Then
        Set rr = ws_out.Range(ws_out.Cells(row_lot, col_lol), ws_out.Cells(row_lot + rows_actual - rows_expect - 1, col_lor))
        Call rr.Delete(Shift:=xlUp)
        
        Set rr = ws_out.Range(ws_out.Cells(row_bot - (rows_actual - rows_expect - 1), col_lhs), ws_out.Cells(row_bot, col_rhs))
        Call rr.Delete(Shift:=xlUp)
        
        For k = row_lob + 1 To row_bot
            ws_out.Rows(k - (rows_actual - rows_expect)).Select
            ws_out.Rows(k - (rows_actual - rows_expect)).RowHeight = row_heights(k)
        Next k
    Else
        'Set rr = ws_out.Range(ws_out.Cells(row, col), ws_out.Cells(row + rows_actual - rows_expect, col_end))
        ''Call ws_out.Range(ws_out.Cells(row, col), ws_out.Cells(row + rows_actual, col_end)).Insert(Shift:=xlDown)
    End If
End If

End Function

Function CreateInvoiceWorksheet(main_ As cMainData, transColl As Collection, invoice_number As Long, date_ As Date) As Worksheet
Dim bScreenUpdating As Boolean: bScreenUpdating = Application.ScreenUpdating
On Error GoTo 10
Application.ScreenUpdating = True

' Create a copy of the template worksheet
ws_tmp.Copy After:=ws_tmp
Dim ws_out As Worksheet: Set ws_out = ws_tmp.Parent.Sheets(ws_tmp.index + 1)

' Replace the token strings with the MainRecord values
Dim main_record As cMainRecord: Set main_record = transColl(1).MainRecord
Call ReplaceTokens(ws_out, main_record, invoice_number, date_)

' Reshape the ListObject (Excel table) and the Invoice cells
Dim rows_expect As Long: rows_expect = transColl.count
Dim lo As ListObject: Set lo = ws_out.ListObjects(1)
Dim row_lot As Long: row_lot = ReshapeTableAndInvoice(lo, rows_expect)

' Copy the values of the transaction records into the cells
Dim row As Long, col As Long: row = row_lot
Dim tr As cTransactionRecord
For Each tr In transColl
    Dim tc
    For Each tc In TransactionColumn2InvoiceHeaderColumn
        col = TransactionColumn2InvoiceHeaderColumn(tc)
        Dim val: val = CallByName(tr, tc, VbGet)
        ' Adjust the UnitPrice with the tax rate so the Transactions sum (which is what shops would charge) is the same as the invoice "TOTAL" field.
        If tc = "UnitPrice" Then
            val = val / (1 - main_.TAX_RATE)
        End If
        ws_out.Cells(row, col) = val
    Next tc
    row = row + 1
Next tr

GoTo 20
10:
Application.ScreenUpdating = bScreenUpdating
Call ErrorExit("Failure creating Invoice Worksheet:'" + Err.Description + "'")
20:
Application.ScreenUpdating = bScreenUpdating
Set CreateInvoiceWorksheet = ws_out
End Function


