Attribute VB_Name = "Main"
Option Explicit

Function SafeGetOutlookApp() As Outlook.Application
On Error GoTo 10
Set SafeGetOutlookApp = GetObject(, "Outlook.Application")
GoTo 20
10:
Set SafeGetOutlookApp = New Outlook.Application
20:
End Function

Function GetTransactionRecords(xx)
' Load the selected transaction records
Dim cws As Worksheet: Set cws = Safe_ThisWorkbook_Worksheets("Customers")
Dim tws As Worksheet: Set tws = Safe_ThisWorkbook_Worksheets("Transactions")

Dim main_ As New cCustomerData
Call main_.Init(cws.Cells(1, 1))
Dim td As New cTransactionData
Call td.Initx(tws)

tws.Parent.Activate
tws.Activate
Dim ret As New Collection
Dim r
For Each r In Selection.Rows
  If r.row > 1 Then
    If r.row > tws.UsedRange.Rows(tws.UsedRange.Rows.count).row Then
      Exit For
    End If
    Dim t As cTransactionRecord: Set t = New cTransactionRecord
    Call t.Inity(main_, td, r.row)
    If t.CustomerID <> "" Then
      ret.Add t
    End If
  End If
Next r
Set GetTransactionRecords = ret
End Function

Sub aProcessRun_Callback(transaction_records As Collection)
' Process the transaction records in transaction_records
Dim main_ As New cCustomerData

Dim tws As Worksheet: Set tws = Safe_ThisWorkbook_Worksheets("Transactions")

Dim t As cTransactionRecord
Dim v
For Each t In transaction_records
  v = t.CustomerRecord.EmailAddressx
  v = t.CustomerRecord.EmailTemplatex
  v = t.CustomerRecord.InvoiceTemplatex
Next t

Dim custId2Trans As New Scripting.Dictionary
Dim itc As New cInvoiceTemplateCollection

Dim count As Long, invoice_count As Long

invoice_count = 0
For Each t In transaction_records
  MainForm.Do_Events "Loading Invoice Template " + t.CustomerRecord.InvoiceTemplatex + " ..."
  Call itc.GetInvoiceTemplate(main_, t.CustomerRecord.InvoiceTemplatex)
  Dim date2Trans  As Scripting.Dictionary
  If custId2Trans.Exists(t.CustomerID) Then
    Set date2Trans = custId2Trans(t.CustomerID)
  Else
    Set date2Trans = New Scripting.Dictionary
    Call custId2Trans.Add(t.CustomerID, date2Trans)
  End If
  Dim transColl As Collection
  If date2Trans.Exists(t.InvoiceDate) Then
    Set transColl = date2Trans(t.InvoiceDate)
  Else
    invoice_count = invoice_count + 1
    Set transColl = New Collection
    Call date2Trans.Add(t.InvoiceDate, transColl)
  End If
  transColl.Add t
Next t

Dim fso As New FileSystemObject
Dim temp_folder As String: temp_folder = ThisWorkbook_Path
temp_folder = temp_folder + "\temp"
If Not fso.FolderExists(temp_folder) Then
  MainForm.Do_Events "Creating output folder ..."
  Call fso.CreateFolder(temp_folder)
End If

Dim it As New cInvoiceTemplateXl
Call it.ClearInvoiceTempWorkbook

Dim CustomerID
count = 0
For Each CustomerID In custId2Trans.Keys()
  Set date2Trans = custId2Trans(CustomerID)
  Dim date_
  For Each date_ In date2Trans.Keys()
    count = count + 1
    Dim prefix As String: prefix = "Processing Invoice " + Str(count) + " of " + Str(invoice_count) + ": " + CustomerID + " " + format(CDate(date_), "yyyy-mm-dd") + " "
    MainForm.log_prefix = prefix
    
    MainForm.Do_Events "Getting Invoice Template ..."
    Set transColl = date2Trans(date_)
    Dim mr As cCustomerRecord: Set mr = transColl(1).CustomerRecord
    Dim invoice_number As Long: invoice_number = main_.NextInvoiceNumber
        
    MainForm.log_prefix = prefix + "Creating Invoice: "
    MainForm.Do_Events ""
        
    MainForm.log_prefix = prefix
    MainForm.Do_Events "Saving Invoice ..."
    Dim now_str As String: now_str = format(Now(), "yyyy-mm-dd-HH-MM-ss")
    Dim invoice_base_name As String: invoice_base_name = temp_folder + "\" + now_str + "." + mr.CustomerID + "." + format(CDate(date_), "yyyy-mm-dd") + "." + format(invoice_number, "0000000")
    Dim invoice_pdf As String: invoice_pdf = invoice_base_name + ".pdf"
    
    Set it = itc.GetInvoiceTemplate(main_, mr.InvoiceTemplatex)
    Dim ws As Worksheet: Set ws = it.CreateInvoiceWorksheet(main_, transColl, invoice_number, CDate(date_))
    If MainForm.ExcelInvoiceOnly Then
        
    Else
        Call ExportWorksheetPDFSilent(ws, invoice_pdf, True, MainForm.EmailOptionNone)
        'Call doc.SaveAs2(invoice_base_name + ".docx")
        'Call doc.SaveAs2(invoice_pdf, Word.wdExportFormatPDF)
            
        If mr.EmailAddressx = "" Then
          'If False Then
          '  Call doc.Close(False)
          'End If
        ElseIf Not MainForm.EmailOptionNone Then
          'Call doc.Close(False)
          
          If MainForm.EmailOptionCreateOnly Or MainForm.EmailOptionSend Then
            MainForm.Do_Events "Creating Email ..."
            Dim ie As cInvoiceEmail: Set ie = New cInvoiceEmail
            Dim display_email As Boolean: display_email = MainForm.EmailOptionCreateOnly
            Call ie.Inite(invoice_pdf, main_, mr, invoice_number, CDate(date_), display_email, main_.SUBJECT_LINE_PREFIX)
          End If
          
          If MainForm.EmailOptionSend Then
            MainForm.Do_Events "Sending Email ..."
            On Error GoTo 10
            ie.email.Send
            On Error GoTo 0
            GoTo 20
10:
            Call ErrorExit("Failure sending email:'" + Err.Description + "'")
20:
          End If
        End If
        
        For Each t In transColl
          tws.Cells(t.rowIndex, t.Parent.col_ProcessedWhen) = now_str
          tws.Cells(t.rowIndex, t.Parent.col_InvoiceNo) = invoice_number
        Next t
    End If
    
  Next date_
Next CustomerID
End Sub

Sub RunReport()
Dim transaction_records As Collection: Set transaction_records = GetTransactionRecords(Empty)
MainForm.ResetState transaction_records
MainForm.Show
End Sub

Sub aaMain()
'Dim r As Range: Set r = Range("[Journal_fl.xlsm]Sheet1!A1")
'End
Application.DisplayAlerts = True
Application.ScreenUpdating = True
If True Then
    Dim transaction_records As Collection: Set transaction_records = GetTransactionRecords(Empty)
    MainForm.ResetState transaction_records
    MainForm.ExcelInvoiceOnly = True
    MainForm.EmailOptionNone = False
    MainForm.EmailOptionCreateOnly = False
    MainForm.EmailOptionSend = False
    Call aProcessRun_Callback(transaction_records)
End If
End Sub

