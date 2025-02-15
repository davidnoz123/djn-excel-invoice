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

Sub Init(invoice_pdf As String, master_record As cMainRecord, invoice_number As Long, date_ As Date)
Dim t As TextStream: Set t = fso.OpenTextFile(master_record.EmailTemplatex)
Dim email_template_s As String: email_template_s = t.ReadAll

email_template_s = Replace(email_template_s, "{RecipientName}", master_record.RecipientNameEmail)

Dim OutlookApp As Outlook.Application: Set OutlookApp = SafeGetOutlookApp()
Set email = OutlookApp.CreateItem(olMailItem)
' https://docs.microsoft.com/en-us/office/vba/api/outlook.attachments.add
' https://www.browserling.com/tools/image-to-base64
On Error GoTo exit_with_failure
With email
  .BodyFormat = olFormatHTML
  .Display
  .HTMLBody = email_template_s + .HTMLBody
  .To = master_record.EmailAddressx
  '.CC = "sdf@gamil.com"
  '.BCC = "hello@gamil.com;hi@gmail.com"
  .Subject = "Rosehaven Florist Invoice:" + format(invoice_number, "0000000")
  Call .Attachments.Add(invoice_pdf)
End With
Exit Sub
exit_with_failure:
Call ErrorExit("Unable to continue after failure creating email:'" + Err.Description + "'")
End Sub

Private Sub Class_Initialize()
Set email = Nothing
End Sub
