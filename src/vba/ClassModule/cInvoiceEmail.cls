VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "cInvoiceEmail"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

Dim fso As New FileSystemObject
Public email As Outlook.MailItem

Sub Inite(invoice_pdf As String, main_ As cCustomerData, main_record As cCustomerRecord, invoice_number As Long, date_ As Date, display_email As Boolean, subject_line_prefix_s As String, ByVal inv_data As cInvoiceData)
Dim t As TextStream: Set t = fso.OpenTextFile(main_record.EmailTemplatex)
Dim email_template_s As String: email_template_s = t.ReadAll

email_template_s = Replace(email_template_s, "{YourBusinessName}", main_.YOUR_BUSINESS_NAME)

' #caedfb invoice blue
' #d0d0d0 invoice grey
' #8c2532 ass_logo_maroon
' #f0aa3e ass_logo_gold

    ' email_version_01.html
    'HTML = Replace(HTML, "{LOGO_SRC}", "data:image/jpeg;base64," & base64Logo)
    'HTML = Replace(HTML, "{CUSTOMER_NAME}", "D Nielsen")
    'HTML = Replace(HTML, "{INVOICE_NUMBER}", "372")
    'HTML = Replace(HTML, "{INVOICE_DATE}", "30/11/2025")
    'HTML = Replace(HTML, "{DUE_DATE}", "28/02/2026")
    'HTML = Replace(HTML, "{SUBTOTAL}", "$495.65")
    'HTML = Replace(HTML, "{TAX_RATE}", "15.00")
    'HTML = Replace(HTML, "{TAX_AMOUNT}", "$74.35")
    'HTML = Replace(HTML, "{TOTAL}", "$570.00")
    'CustomerID    Barracuda Engineering Ltd
    'EmailAddress  info@barracudaengineering.co.nz
    'EmailAddressOld
    'EmailTemplate email_version_01.HTML
    'InvoiceTemplate template2
    'RecipientNameEmail Barry & Pavle
    'RecipientNameInvoice
    'CompanyName   Barracuda Engineering Ltd
    'StreetAddress 19B Rylock Place, Pakuranga
    'TownCityPostCode
    'Country
    'Phone
    
Dim adhoc_replacements As New Scripting.Dictionary
adhoc_replacements("INVOICE_NUMBER") = CStr(invoice_number)
adhoc_replacements("INVOICE_DATE") = CStr(date_)
adhoc_replacements("DUE_IN_DAYS") = CStr(CLng(inv_data.dueDate - date_))
adhoc_replacements("DUE_DATE") = CStr(inv_data.dueDate)
adhoc_replacements("SUBTOTAL") = FormatCurrency(inv_data.Invoice_Subtotal)
adhoc_replacements("TAX_RATE") = format(Round(100# * inv_data.TAX_RATE, 2), "0.00")
adhoc_replacements("TAX_AMOUNT") = FormatCurrency(inv_data.Invoice_Subtotal * inv_data.TAX_RATE)
adhoc_replacements("TOTAL") = FormatCurrency(inv_data.Invoice_Subtotal * (1# + inv_data.TAX_RATE))

Dim field_name, field_value
For Each field_name In main_record.FieldName2Value.Keys()
    email_template_s = Replace(email_template_s, "{" + field_name + "}", CStr(main_record.FieldName2Value(field_name)))
Next field_name
For Each field_name In adhoc_replacements.Keys()
    email_template_s = Replace(email_template_s, "{" + field_name + "}", CStr(adhoc_replacements(field_name)))
Next field_name

Const use_logo = True
Const use_outlook_2_send = False

If use_logo Then
    Dim logoPath As String: logoPath = ThisWorkbook.path + "\" + "logo.jpeg"
    Dim base64Logo: base64Logo = FileToBase64(logoPath)
    email_template_s = Replace(email_template_s, "{LOGO_SRC}", "data:image/jpeg;base64," & base64Logo)
End If

Debug.Print email_template_s

Dim OutlookApp As Outlook.Application: Set OutlookApp = SafeGetOutlookApp()
Set email = OutlookApp.CreateItem(olMailItem)
' https://docs.microsoft.com/en-us/office/vba/api/outlook.attachments.add
' https://www.browserling.com/tools/image-to-base64
On Error GoTo exit_with_failure

If use_outlook_2_send Then
    ' Find the account you want to send from
    Dim acct As Outlook.Account      ' make sure this is declared As Outlook.Account
    For Each acct In OutlookApp.Session.Accounts
        If LCase(acct.SmtpAddress) = LCase(main_.YOUR_BUSINESS_EMAIL) Then
            GoTo found_email_account
        End If
    Next acct
    Call ErrorExit("Failed to find Outlook Account with SmtpAddress equal to:'" + main_.YOUR_BUSINESS_EMAIL + "'")
found_email_account:
End If

With email
  If display_email Then
    .Display
    Dim insp
    Set insp = .GetInspector
    insp.WindowState = olMinimized
  End If
  .BodyFormat = olFormatHTML
  .HTMLBody = email_template_s + .HTMLBody
  .To = main_record.EmailAddressx
  '.CC = "sdf@gamil.com"
  '.BCC = "hello@gamil.com;hi@gmail.com"
  .Subject = subject_line_prefix_s + ":" + format(invoice_number, "0000000")
  
  Call .Attachments.Add(invoice_pdf)
End With

If use_outlook_2_send Then
    email.SendUsingAccount = acct
End If

Exit Sub
exit_with_failure:
Call ErrorExit("Unable to continue after failure creating email:'" + Err.Description + "'")
End Sub

Private Sub Class_Initialize()
Set email = Nothing
End Sub
