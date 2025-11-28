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

Const INVOICE_TEMP_WORKBOOK = "InvoiceTemplateTemp.xlsx"

Dim fso As New FileSystemObject
Dim ws_tmp As Worksheet
Dim TransactionColumn2InvoiceHeaderColumn As Scripting.Dictionary
Dim TAX_RATE_value As Double
Dim Invoice_Subtotal_formula As String

Sub Init(main_ As cCustomerData, template_name As String)
Dim ref: ref = ParseCellReference(template_name)
Dim wb As Workbook: Set wb = Nothing
If ref(1) <> "" Then
    Dim drName As String: drName = ref(0)
    If drName = "" Then
        drName = ThisWorkbook.path
    Else
        Dim fso As New FileSystemObject
        drName = fso.GetAbsolutePathName(drName)
    End If
    Dim wbName As String: wbName = ref(1)
    If WbExists(wbName) Then
        Set wb = Workbooks(wbName)
    Else
        Dim full_path As String: full_path = drName + "\" + wbName
        If fso.FileExists(full_path) Then
            Set wb = Workbooks.Open(full_path) ' , ReadOnly:=True
        Else
            Call ErrorExit("Failed to find template_name:'" + template_name + "' FileDoesNotExist:'" + full_path + "'")
        End If
    End If
End If
If wb Is Nothing Then Set wb = ThisWorkbook
Dim worksheet_name As String: worksheet_name = ref(2)
If Not WsExists(wb, worksheet_name) Then
    Call ErrorExit("Missing worksheet in Invoice Templates Excel Workbook:'" + template_name + "'")
End If
Set ws_tmp = wb.Worksheets(worksheet_name)
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

TAX_RATE_value = -1#
Invoice_Subtotal_formula = ""

If True Then
    Const TAX_RATE_rng_name = "TAX_RATE"
    Dim n As name
    For Each n In ws_tmp.Names
        If EndsWith(n.name, TAX_RATE_rng_name) Then
            If TAX_RATE_value >= 0 Then
                Call ErrorExitX("More than one named range on the sheet has a name ending with:'" + TAX_RATE_rng_name + "'", n.RefersToRange)
            End If
            If n.name <> ws_tmp.name + "!" + TAX_RATE_rng_name Then
                Call ErrorExitX("Unexpected scope for named range:'" + TAX_RATE_rng_name + "'", n.RefersToRange)
            End If
            If n.RefersToRange.Cells.CountLarge <> 1 Then
                Call ErrorExitX("Unexpected multi-cell range for name:'" + TAX_RATE_rng_name + "'", n.RefersToRange)
            End If
            If Not IsNumeric(n.RefersToRange.Value) Then
                Call ErrorExitX("Unexpected non-numeric value in named range:'" + TAX_RATE_rng_name + "'", n.RefersToRange)
            End If
            If n.RefersToRange.Value < 0 Or n.RefersToRange.Value > 1# Then
                Call ErrorExitX("Unexpected out-of-range (i.e., [0.0, 1.0]) numeric value in named range:'" + TAX_RATE_rng_name + "'", n.RefersToRange)
            End If
            TAX_RATE_value = CDbl(n.RefersToRange.Value)
        End If
    Next n
    
    If TAX_RATE_value = -1# Then
        Call ErrorExitX("Unexpected missing named range in worksheet:'" + TAX_RATE_rng_name + "'", ws_tmp.UsedRange)
    End If
End If

If True Then
    Dim tmp As Range: Set tmp = Invoice_Subtotal_formula_cell(ws_tmp)
End If

End Sub

Private Function Invoice_Subtotal_formula_cell(ws As Worksheet) As Range
Dim ret As Range: Set ret = Nothing

Dim re As New VBScript_RegExp_55.RegExp
Const Invoice_Subtotal_formula_regex_pattern = "^=SUM\(SimpleInvoice[0-9]+\[Price\]\)$"
re.Pattern = Invoice_Subtotal_formula_regex_pattern
re.Global = False
re.IgnoreCase = False
    
Dim c As Range
For Each c In ws.UsedRange.SpecialCells(xlCellTypeFormulas)
    If re.Test(c.Formula) Then
        If Invoice_Subtotal_formula <> "" Then
            Call ErrorExitX("More than one cell formula matches regular expression:'" + Invoice_Subtotal_formula_regex_pattern + "'", c)
        End If
        Set ret = c
    End If
Next c
If ret Is Nothing Then
    Call ErrorExitX("Unexpected worksheet missing formula matching regular expression:'" + Invoice_Subtotal_formula_regex_pattern + "'", ws.UsedRange)
End If
Set Invoice_Subtotal_formula_cell = ret
End Function

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

Private Sub ReplaceTokens(ws As Worksheet, main_record As cCustomerRecord, invoice_number As Long, date_ As Date, dueDate As Date)
' Substitute the variable strings in the document
Dim search_and_replace As New Scripting.Dictionary
search_and_replace.CompareMode = TextCompare
search_and_replace.Add "{InvoiceNo}", format(invoice_number, "0000000")
search_and_replace.Add "{InvoiceDate}", format(date_, "")
search_and_replace.Add "{DaysToDueDate}", CStr(CLng(dueDate - date_))


search_and_replace.Add "{InvoicedEmail}", StringifyInvoiceContactDetail(main_record.EmailAddressx, "Email: ")

Dim field_name
For Each field_name In main_record.FieldName2Value.Keys()
    Dim field_value: field_value = main_record.FieldName2Value(field_name)
    search_and_replace.Add "{" + field_name + "}", StringifyInvoiceContactDetail(CStr(field_value))
Next field_name

Call SearchAndReplace(ws, search_and_replace)
End Sub

Private Function ReshapeTableAndInvoice(lo As ListObject, rows_expect As Long) As Long

If lo.DataBodyRange Is Nothing Then
    Call ErrorExitX("Table has an empty DataBodyRange", lo.HeaderRowRange)
End If

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

'Application.ScreenUpdating = True

' Clear out all the values of table data range cells
GetNonFormulaCells(lo.DataBodyRange).ClearContents

Dim ff As Range, dd As Range
Dim rows_actual As Long: rows_actual = lo.DataBodyRange.Rows.count
If rows_actual = rows_expect Then
    ' The table has the right number of cells
Else
    ' The table doesn't have the right number of cells
    Dim rr As Range
    If rows_actual > rows_expect Then
        ' We've got too many cells in the table
        
        ' Get the cell range of the table cells that we might consider to delete
        'ws_out.Parent.Activate
        'ws_out.Activate
        Set rr = ws_out.Range(ws_out.Cells(row_lot, col_lol), ws_out.Cells(row_lot + rows_actual - rows_expect - 1, col_lor))
        
        'rr.Select
        
        ' Get the random of the overhang of the worksheet after we delete the table rows in range 'rr'
        Set dd = ws_out.Range(ws_out.Cells(row_lot, col_lhs), ws_out.Cells(row_lot + rows_actual - rows_expect - 1, col_rhs))
        
        'dd.Select
        
        ' Get the disjunction of rr and dd ... if these cells contain values, then we don't want to delete anything.
        Set ff = Disjunction(rr, dd)
        'ff.Select
        Dim tt As Range: Set tt = ff.Find("*")
        
        If Not tt Is Nothing Then
            ' Don't delete anything
            ReshapeTableAndInvoice = rr.row
        Else
            ' We will only be deleting empty cells ... this is fine
            
            ' Delete the excess rows in the table
            Call rr.Delete(Shift:=xlUp)
            
            ' Delete the excess rows outside the table
            For Each tt In ff.Areas
                Call tt.Delete(Shift:=xlUp)
            Next tt
            
        End If

    Else
        ' We've got too few cells in the table
        
        Dim row_lob_tmp As Long
        Dim src As Range, dst As Range
        
        ' Add new rows to the table (as fast as we can)
        row_lob_tmp = row_lob
        Dim found As Boolean: found = False
        While Not found
            If 2 * (row_lob_tmp - row_lot + 1) > rows_expect Then
                found = True
                row_lob_tmp = row_lot + rows_expect - (row_lob_tmp - row_lot + 1) - 1
            End If
            Set rr = ws_out.Range(ws_out.Cells(row_lot, col_lol), ws_out.Cells(row_lob_tmp, col_lor))
            Call rr.Insert(Shift:=xlDown)
            row_lob_tmp = row_lob_tmp + (row_lob_tmp - row_lot) + 1
        Wend
        
        If True Then
            ' Paste the formatting into the new cells now to the left and right up to the end of the invoice

            If col_lol > col_lhs Then
                Set src = ws_out.Range(ws_out.Cells(row_lob + 1, col_lhs), ws_out.Cells(row_bot, col_lol - 1))
                src.Copy
                Set dst = ws_out.Cells(row_lob + 1 - (rows_actual - rows_expect), col_lhs)
                Call dst.PasteSpecial(xlPasteAll)
                
                Set src = ws_out.Range(ws_out.Cells(row_lob, col_lhs), ws_out.Cells(row_lob, col_lol - 1))
                src.Copy
                Set dst = ws_out.Range(ws_out.Cells(row_lob + 1, col_lhs), ws_out.Cells(row_lob - (rows_actual - rows_expect), col_lol - 1))
                Call dst.PasteSpecial(xlPasteFormats)
                dst.ClearContents
            End If
            
            If col_rhs > col_lor Then
                Set src = ws_out.Range(ws_out.Cells(row_lob + 1, col_lor + 1), ws_out.Cells(row_bot, col_rhs))
                src.Copy
                Set dst = ws_out.Cells(row_lob + 1 - (rows_actual - rows_expect), col_lor + 1)
                Call dst.PasteSpecial(xlPasteAll)
                
                Set src = ws_out.Range(ws_out.Cells(row_lob, col_lor + 1), ws_out.Cells(row_lob, col_rhs))
                src.Copy
                Set dst = ws_out.Range(ws_out.Cells(row_lob + 1, col_lor + 1), ws_out.Cells(row_lob - (rows_actual - rows_expect), col_rhs))
                Call dst.PasteSpecial(xlPasteFormats)
                dst.ClearContents
            End If
        End If
        
        ' Reset the row heights after the table rows are deleted
        For k = row_lob + 1 To row_lob + rows_expect - 2
            'ws_out.Rows(k).Select
            ws_out.Rows(k).RowHeight = row_heights(row_lob)
        Next k
        

        
    End If
    
    ' Reset the row heights after the table rows are deleted
    For k = row_lob + 1 To row_bot
        'ws_out.Rows(k - (rows_actual - rows_expect)).Select
        ws_out.Rows(k - (rows_actual - rows_expect)).RowHeight = row_heights(k)
    Next k
    
End If

End Function

Sub ClearInvoiceTempWorkbook()
Dim bScreenUpdating As Boolean: bScreenUpdating = Application.ScreenUpdating
Dim bDisplayAlerts As Boolean: bDisplayAlerts = Application.DisplayAlerts
On Error GoTo 10
Application.ScreenUpdating = False
Application.DisplayAlerts = False

Dim wb_out As Workbook
Dim tmp_xlsx As String: tmp_xlsx = ThisWorkbook.path + "\" + INVOICE_TEMP_WORKBOOK
If Not WorkbookIsOpen(tmp_xlsx, wb_out) Then
    Dim bFileAlreadyOpen As Boolean
    Set wb_out = SafeGetWorkbook(tmp_xlsx, bFileAlreadyOpen)
End If

Dim i As Long: i = wb_out.Sheets.count
While i > 1
    wb_out.Sheets(i).Delete
    i = i - 1
Wend

GoTo 20
10:
Application.DisplayAlerts = bDisplayAlerts
Application.ScreenUpdating = bScreenUpdating
Call ErrorExit("Failure deleting tmporary invoice worksheet:'" + Err.Description + "'")
20:
Application.DisplayAlerts = bDisplayAlerts
Application.ScreenUpdating = bScreenUpdating
End Sub

Function CreateInvoiceWorksheet(main_ As cCustomerData, transColl As Collection, invoice_number As Long, date_ As Date, ByRef inv_data As cInvoiceData) As Worksheet
Dim bScreenUpdating As Boolean: bScreenUpdating = Application.ScreenUpdating
On Error GoTo 10
Application.ScreenUpdating = False

Set inv_data = New cInvoiceData

inv_data.TAX_RATE = TAX_RATE_value
inv_data.dueDate = InvoiceDueDate(date_, 90)

Dim wb_out As Workbook
Dim tmp_xlsx As String: tmp_xlsx = ThisWorkbook.path + "\" + INVOICE_TEMP_WORKBOOK
If Not WorkbookIsOpen(tmp_xlsx, wb_out) Then
    Dim bFileAlreadyOpen As Boolean
    Set wb_out = SafeGetWorkbook(tmp_xlsx, bFileAlreadyOpen)
End If

Dim ws_cur As Worksheet: Set ws_cur = ActiveSheet ' Get the current active sheet

' Create a copy of the template worksheet
ws_tmp.Copy After:=wb_out.Sheets(wb_out.Sheets.count) ' Copy ws_tmp and put at the end of wb_out

ws_cur.Activate ' The call to ws_tmp.Copy has changed the active sheet ... restore the active sheet

Dim ws_out As Worksheet: Set ws_out = wb_out.Sheets(wb_out.Sheets.count)

' Replace the token strings with the CustomerRecord values
Dim main_record As cCustomerRecord: Set main_record = transColl(1).CustomerRecord
Call ReplaceTokens(ws_out, main_record, invoice_number, date_, inv_data.dueDate)

' Reshape the ListObject (i.e., the Excel table) and the Invoice cells
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
        Dim val: val = tr.FieldName2Value(tc)
        ws_out.Cells(row, col) = val
    Next tc
    row = row + 1
Next tr

Dim tmp As Range: Set tmp = Invoice_Subtotal_formula_cell(ws_out)

inv_data.Invoice_Subtotal = tmp.Value



GoTo 20
10:
Application.ScreenUpdating = bScreenUpdating
Call ErrorExit("Failure creating Invoice Worksheet:'" + Err.Description + "'")
20:
Application.ScreenUpdating = bScreenUpdating
Set CreateInvoiceWorksheet = ws_out
End Function


