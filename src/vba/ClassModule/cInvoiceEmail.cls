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

Sub Inite(invoice_pdf As String, main_ As cMainData, main_record As cMainRecord, invoice_number As Long, date_ As Date, display_email As Boolean, subject_line_prefix_s As String)
Dim t As TextStream: Set t = fso.OpenTextFile(main_record.EmailTemplatex)
Dim email_template_s As String: email_template_s = t.ReadAll

'email_template_s = Replace(email_template_s, "{RecipientName}", main_record.RecipientNameEmail)
email_template_s = Replace(email_template_s, "{YourBusinessName}", main_.YOUR_BUSINESS_NAME)

Dim field_name
For Each field_name In main_record.FieldName2Value.Keys()
    Dim field_value: field_value = main_record.FieldName2Value(field_name)
    email_template_s = Replace(email_template_s, "{" + field_name + "}", CStr(field_value))
Next field_name

Dim OutlookApp As Outlook.Application: Set OutlookApp = SafeGetOutlookApp()
Set email = OutlookApp.CreateItem(olMailItem)
' https://docs.microsoft.com/en-us/office/vba/api/outlook.attachments.add
' https://www.browserling.com/tools/image-to-base64
On Error GoTo exit_with_failure
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
Exit Sub
exit_with_failure:
Call ErrorExit("Unable to continue after failure creating email:'" + Err.Description + "'")
End Sub

Private Sub Class_Initialize()
Set email = Nothing
End Sub
