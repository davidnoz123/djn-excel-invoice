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

Dim master As New cMasterData
Call master.Init(cws.Cells(1, 1))
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
    Call t.Inity(master, td, r.row)
    If t.CustomerID <> "" Then
      ret.Add t
    End If
  End If
Next r
Set GetTransactionRecords = ret
End Function

Sub aProcessRun_Callback(transaction_records As Collection)
' Process the transaction records in transaction_records
Dim master As New cMasterData

Dim tws As Worksheet: Set tws = Safe_ThisWorkbook_Worksheets("Transactions")

Dim t As cTransactionRecord
Dim v
For Each t In transaction_records
  v = t.MasterRecord.EmailAddressx
  v = t.MasterRecord.EmailTemplatex
  v = t.MasterRecord.InvoiceTemplatex
Next t

Dim custId2Trans As New Scripting.Dictionary
Dim itc As New cInvoiceTemplateCollection

Dim count As Long, invoice_count As Long

invoice_count = 0
For Each t In transaction_records
  MainForm.Do_Events "Loading Invoice Template " + t.MasterRecord.InvoiceTemplatex + " ..."
  Call itc.GetInvoiceTemplate(t.MasterRecord.InvoiceTemplatex)
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
    Dim mr As cMasterRecord: Set mr = transColl(1).MasterRecord
    Dim invoice_number As Long: invoice_number = master.NextInvoiceNumber
    Dim it As cInvoiceTemplate: Set it = itc.GetInvoiceTemplate(mr.InvoiceTemplatex)
        
    MainForm.log_prefix = prefix + "Creating Invoice: "
    MainForm.Do_Events ""
    Dim doc As Word.Document: Set doc = it.CreateInvoiceDocument(transColl, invoice_number, CDate(date_))
    MainForm.log_prefix = prefix
    
    Dim now_str As String: now_str = format(Now(), "yyyy-mm-dd-HH-MM-ss")
    
    MainForm.Do_Events "Saving Invoice ..."
    Dim invoice_base_name As String: invoice_base_name = temp_folder + "\" + now_str + "." + mr.CustomerID + "." + format(CDate(date_), "yyyy-mm-dd") + "." + format(invoice_number, "0000000")
    Dim invoice_pdf As String: invoice_pdf = invoice_base_name + ".pdf"
    Call doc.SaveAs2(invoice_base_name + ".docx")
    Call doc.SaveAs2(invoice_pdf, Word.wdExportFormatPDF)
        
    If mr.EmailAddressx = "" Then
      'If False Then
      '  Call doc.Close(False)
      'End If
    ElseIf Not MainForm.EmailOptionNone Then
      Call doc.Close(False)
      
      If MainForm.EmailOptionCreateOnly Or MainForm.EmailOptionSend Then
        MainForm.Do_Events "Creating Email ..."
        Dim ie As cInvoiceEmail: Set ie = New cInvoiceEmail
        Call ie.Init(invoice_pdf, mr, invoice_number, CDate(date_))
      End If
      
      If MainForm.EmailOptionSend Then
        MainForm.Do_Events "Sending Email ..."
        ie.email.Send
      End If
    End If
    
    For Each t In transColl
      tws.Cells(t.rowIndex, t.Parent.col_Status) = now_str
      tws.Cells(t.rowIndex, t.Parent.col_InvoiceNo) = invoice_number
    Next t
    
  Next date_
Next CustomerID

End Sub

Sub RunReport()
Dim transaction_records As Collection: Set transaction_records = GetTransactionRecords(Empty)
MainForm.ResetState transaction_records
MainForm.Show
End Sub
