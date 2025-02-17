VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "cInvoiceTemplate"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

Const INVOICE_TABLE_TEMPLATES_DOC = "invoice_table_templates.docx"

Dim fso As New FileSystemObject
Dim template As Word.Document
Dim template_for_tables As Word.Document
Dim table_index As Long
Dim keys_to_table_info As New Scripting.Dictionary

Function SafeGetWordApp() As Word.Application
On Error GoTo 10
Set SafeGetWordApp = GetObject(, "Word.Application")
GoTo 20
10:
Set SafeGetWordApp = New Word.Application
20:
SafeGetWordApp.Visible = True
End Function

Function SafeGetTemplateDocument(template_file As String)
template_file = fso.GetAbsolutePathName(template_file)
Dim template_base As String: template_base = fso.GetFileName(template_file)
Dim WordApp As Word.Application: Set WordApp = SafeGetWordApp
Dim d As Word.Document
For Each d In WordApp.Documents
  If template_file = d.FullName Then
    GoTo 10
  End If
  If template_base = d.name Then
    'WordApp.Activate
    'd.Activate
    'Call ErrorExit("CloseWordDocumentBeforeProceeding:" + template_base)
  End If
Next d
Set d = WordApp.Documents.Open(template_file, ReadOnly:=True)
10:
Set SafeGetTemplateDocument = d
End Function

Private Function GetTemplateForTablesDocument() As Word.Document
If template_for_tables Is Nothing Then
  Dim template_file As String: template_file = fso.GetAbsolutePathName(ThisWorkbook_Path + "\" + INVOICE_TABLE_TEMPLATES_DOC)
  Dim template_base As String: template_base = fso.GetFileName(template_file)
  Dim WordApp As Word.Application: Set WordApp = SafeGetWordApp
  Dim d As Word.Document
  For Each d In WordApp.Documents
    If template_file = d.FullName Then
      GoTo 10
    End If
    If template_base = d.name Then
      'WordApp.Activate
      'd.Activate
      'Call ErrorExit("CloseWordDocumentBeforeProceeding:" + template_base)
    End If
  Next d
  Set d = WordApp.Documents.Open(template_file, ReadOnly = True)
10:
  Set template_for_tables = d
End If
Set GetTemplateForTablesDocument = template_for_tables
End Function

Sub Init(template_file As String)
Set template = SafeGetTemplateDocument(template_file)
End Sub

Sub GetItemsTable(doc As Word.Document, ByRef tableIndex As Long, ByRef mn As Scripting.Dictionary, a, Optional ByVal row_start_indicator_string As String = vbNullString)
Dim doc_full_path As String: doc_full_path = doc.Path + "\" + doc.name
Dim mx As New Scripting.Dictionary
Dim i As Long
Dim mm As New Scripting.Dictionary
mm.CompareMode = TextCompare
Dim indicator As Long: indicator = 1
For i = LBound(a) To UBound(a)
  If a(i) = row_start_indicator_string Then indicator = -1
  Call mm.Add(a(i), indicator)
Next i
For tableIndex = 1 To doc.Tables.count
  Dim t As Word.Table: Set t = doc.Tables(tableIndex)
  Dim r As Word.row
  i = 0
  Set mn = New Scripting.Dictionary
  mn.CompareMode = TextCompare
  For Each r In t.Rows
    i = i + 1
    Dim c As Word.Cell
    Dim j As Long: j = 0
    MainForm.Do_Events "GetItemsTables " + CStr(tableIndex) + " " + CStr(i) + " ..."
    For Each c In r.Cells
      j = j + 1
      Dim cellText As String: cellText = Left(c.Range.Text, Len(c.Range.Text) - 2)
      If cellText = "" Then
        GoTo next_c
      End If
      If Not mm.Exists(cellText) Then
        GoTo next_t
      End If
      If mn.Exists(cellText) Then
        GoTo next_t
      End If
      indicator = mm(cellText)
      If indicator > 0 Then
        If i <> 1 Then
          GoTo next_t
        End If
        Call mn.Add(cellText, j)
      Else
        If j <> r.Cells.count - 1 Then
          GoTo next_t
        End If
        Call mn.Add(cellText, i)
      End If
next_c:
    Next c
  Next r
  If mn.count = mm.count Then
    GoTo found
  End If
next_t:
If mn.count > mx.count Then
  Set mx = mn
End If
Next tableIndex
Call ErrorExitWithWord(doc, "Missing table in Word Document with header:" + Join(a, ", ") + vbCrLf + "See Word Document:" + vbCrLf + "'" + doc_full_path + "'")
found:
End Sub

Private Sub AddItemsTable(document_key As String, doc As Word.Document, ByRef tableIndex As Long, ByRef mn As Scripting.Dictionary, a, Optional ByVal row_start_indicator_string = vbNullString)
Call GetTemplateForTablesDocument
Dim itd As cInvoiceTableDetails
If keys_to_table_info.Exists(document_key) Then
  Set itd = keys_to_table_info(document_key)
  tableIndex = itd.tableIndex
  Set mn = itd.mn
Else
  Call GetItemsTable(template_for_tables, tableIndex, mn, a, row_start_indicator_string:=row_start_indicator_string)
  Set itd = New cInvoiceTableDetails
  itd.tableIndex = tableIndex
  Set itd.mn = mn
  Call keys_to_table_info.Add(document_key, itd)
End If

Dim t As Word.Table: Set t = template_for_tables.Tables(tableIndex)
Dim r As Word.Range
Set r = doc.Content
r.Find.Text = "{Items Table}"
r.Find.Execute
If Not r.Find.found Then
  Call ErrorExitWithWord(doc, "Missing string '{Items Table}' in document", r)
End If
On Error GoTo 10
t.Range.Copy
Call r.PasteAndFormat(wdFormatOriginalFormatting)
GoTo 20
10:
Call ErrorExitWithWord(doc, "Error copying and pasting:'" + Err.Description + "'", r)
20:
Dim tNew As Word.Table: Set tNew = r.Tables(1)
Dim i As Long
For i = 1 To doc.Tables.count
  If tNew.Range.Start = doc.Tables(i).Range.Start Then
    tableIndex = i
    GoTo found_new_table
  End If
Next i
Call ErrorExitWithWord(doc, "Failed to find table", r)
found_new_table:
End Sub


Sub SearchAndReplace(doc As Word.Document, search_and_replace As Scripting.Dictionary)
Dim k
For Each k In search_and_replace.Keys()
  While True
    MainForm.Do_Events "SearchAndReplace " + search_and_replace(k) + " ..."
    Dim r As Word.Range
    Set r = doc.Content
    r.Find.Text = k
    r.Find.Execute
    If Not r.Find.found Then
      GoTo 99
    End If
    On Error GoTo 10
    r = search_and_replace(k)
    GoTo 20
10:
    Call ErrorExitWithWord(doc, "Failure in search and replace", r)
20:
  Wend
99:
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

Private Sub ReplaceTokens(doc As Word.Document, master_record As cMainRecord, invoice_number As Long, date_ As Date)
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

Call SearchAndReplace(doc, search_and_replace)
End Sub

Private Sub Table_Standard(doc As Word.Document, transColl As Collection)

Dim master_record As cMainRecord: Set master_record = transColl(1).MainRecord

' Get the items table in the document
Dim a:  a = Array( _
"QUANTITY", _
"DESCRIPTION", _
"DELIVERY DATE", _
"UNIT PRICE", _
"TOTAL", _
"SUBTOTAL", _
"TAX", _
"TOTAL DUE" _
)
Dim m_n As Scripting.Dictionary
Call AddItemsTable("Standard", doc, table_index, m_n, a, "SUBTOTAL")
Dim t As Word.Table: Set t = doc.Tables(table_index)
Dim col_QUANTITY As Long: col_QUANTITY = m_n("QUANTITY")
Dim col_DESCRIPTION As Long: col_DESCRIPTION = m_n("DESCRIPTION")
Dim col_DELIVERY_DATE As Long: col_DELIVERY_DATE = m_n("DELIVERY DATE")
Dim col_UNIT_PRICE As Long: col_UNIT_PRICE = m_n("UNIT PRICE")
Dim col_TOTAL As Long: col_TOTAL = m_n("TOTAL")
Dim row_SUBTOTAL As Long: row_SUBTOTAL = m_n("SUBTOTAL")
Dim row_TAX As Long: row_TAX = m_n("TAX")
Dim row_TOTAL_DUE As Long: row_TOTAL_DUE = m_n("TOTAL DUE")

Dim row_last_summary As Long: row_last_summary = WorksheetFunction.Max(row_SUBTOTAL, row_TAX, row_TOTAL_DUE)

Dim row_first_summary As Long: row_first_summary = WorksheetFunction.Min(row_SUBTOTAL, row_TAX, row_TOTAL_DUE)
Dim row_first_summary_orig As Long:: row_first_summary_orig = row_first_summary

' Populate the items table
Dim tr As cTransactionRecord
Dim i As Long: i = 1
Dim bef_tax_sum As Double: bef_tax_sum = 0#
Dim aft_tax_sum As Double: aft_tax_sum = 0#
Dim tax_sum As Double: tax_sum = 0#
For Each tr In transColl
  If Not master_record Is tr.MainRecord Then
    Call ErrorExit("Not master_record Is tr.MainRecord")
  End If
  i = i + 1
  If i >= row_first_summary - 2 Then
    row_first_summary = row_first_summary + 1
    Call t.Rows.Add(t.Rows(row_first_summary - 2))
  End If
  MainForm.Do_Events "ItemsTable " + CStr(tr.Description)
  t.Rows(i).Cells(col_QUANTITY).Range = CStr(tr.Quantity)
  t.Rows(i).Cells(col_DESCRIPTION).Range = CStr(tr.Description)
  t.Rows(i).Cells(col_DELIVERY_DATE).Range = format(tr.DeliveryDate, "yyyy-mm-dd")
  
  Dim aft_tax As Double: aft_tax = tr.UnitPrice
  Dim bef_tax As Double: bef_tax = Round(aft_tax / (1# + master_record.Parent.TAX_RATE), 2)
  
  t.Rows(i).Cells(col_UNIT_PRICE).Range = FormatCurrency(bef_tax, 2)
  
  aft_tax = tr.Quantity * aft_tax
  bef_tax = tr.Quantity * bef_tax
  
  t.Rows(i).Cells(col_TOTAL).Range = FormatCurrency(bef_tax, 2)
  
  bef_tax_sum = bef_tax_sum + bef_tax
  aft_tax_sum = aft_tax_sum + aft_tax
  tax_sum = tax_sum + (aft_tax - bef_tax)
Next tr

t.Rows(row_first_summary - 1).Delete

row_SUBTOTAL = row_SUBTOTAL + row_first_summary - row_first_summary_orig - 1
row_TAX = row_TAX + row_first_summary - row_first_summary_orig - 1
row_TOTAL_DUE = row_TOTAL_DUE + row_first_summary - row_first_summary_orig - 1

t.Rows(row_SUBTOTAL).Cells(t.Rows(row_SUBTOTAL).Cells.count).Range = FormatCurrency(bef_tax_sum, 2)
t.Rows(row_TAX).Cells(t.Rows(row_TAX).Cells.count).Range = FormatCurrency(tax_sum, 2)
t.Rows(row_TOTAL_DUE).Cells(t.Rows(row_TOTAL_DUE).Cells.count).Range = FormatCurrency(aft_tax_sum, 2)

End Sub

Function CreateInvoiceDocument(transColl As Collection, invoice_number As Long, date_ As Date) As Word.Document
Dim doc As Word.Document
Set doc = template.Application.Documents.Add(template.FullName)
Dim master_record As cMainRecord: Set master_record = transColl(1).MainRecord
Call ReplaceTokens(doc, master_record, invoice_number, date_)
Call Table_Standard(doc, transColl)
Set CreateInvoiceDocument = doc
End Function


Private Sub Class_Initialize()
Set template_for_tables = Nothing
End Sub
