VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "cMasterData"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

Public INVOICE_TEMPLATE_DEFAULT As String
Public INVOICE_EMAIL_TEMPLATE_DEFAULT As String

Dim wsMaster As Worksheet
Public fso As New FileSystemObject
Public TopLeft As Range

Public col_CustomerID As Long
Public col_EmailAddress As Long
Public col_EmailTemplate As Long
Public col_InvoiceTemplate As Long
Public col_RecipientNameEmail As Long
Public col_RecipientNameInvoice As Long
Public col_CompanyName As Long
Public col_StreetAddress As Long
Public col_TownCityPostCode As Long
Public col_Country As Long
Public col_Phone As Long
Public col_SubCustomerCommission As Long
Public GST_RATE As Double

'Dim NEXT_INVOICE_NUMBER_r As Range

Dim NEXT_INVOICE_NUMBER_txt As String

Dim d
Dim row_by_customer_id As Scripting.Dictionary
Dim rec_by_customer_id As Scripting.Dictionary

Public EmailAddressRegExp As New VBScript_RegExp_55.RegExp

Function NextInvoiceNumber() As Long
Dim ts As TextStream
Dim ret As Long
If Not fso.FileExists(NEXT_INVOICE_NUMBER_txt) Then
  Set ts = fso.CreateTextFile(NEXT_INVOICE_NUMBER_txt)
  Call ts.WriteLine("1")
  ts.Close
  ret = MsgBox("A new file '" + NEXT_INVOICE_NUMBER_txt + "' has just been created to track invoice numbers starting from '1'." + vbCrLf + "Feel free to edit this file to change where the invioice numbering starts." + vbCrLf + "Do you want to continue?", vbYesNo + vbQuestion, "Next Invoice Number")
  If ret = vbNo Then
    End
  End If
End If
Set ts = fso.OpenTextFile(NEXT_INVOICE_NUMBER_txt, ForReading)
Dim s As String: s = ts.ReadLine()
ts.Close
On Error GoTo exit_with_failure
ret = CLng(s)
GoTo success
exit_with_failure:
ret = MsgBox("The first line of file '" + NEXT_INVOICE_NUMBER_txt + "' does not contain a number to use as the next invoice number." + vbCrLf + "Edit this file to have a number on the first line before continuing.", vbOKOnly + vbCritical, "Next Invoice Number")
End
success:
Set ts = fso.OpenTextFile(NEXT_INVOICE_NUMBER_txt, ForWriting)
Call ts.WriteLine(CStr(ret + 1))
ts.Close
NextInvoiceNumber = ret
End Function

Property Get RecordsByCustomerID(CustomerID) As cMasterRecord
Dim r As cMasterRecord
If rec_by_customer_id.Exists(CustomerID) Then
  Set r = rec_by_customer_id(CustomerID)
Else
  If Not row_by_customer_id.Exists(CustomerID) Then
    Set r = Nothing
  Else
    Set r = New cMasterRecord: Call r.Init(Me, d, row_by_customer_id(CustomerID))
    Call rec_by_customer_id.Add(CustomerID, r)
  End If
End If
Set RecordsByCustomerID = r
End Property

Sub Init(rTopLeft As Range)

Set TopLeft = rTopLeft
Set row_by_customer_id = New Scripting.Dictionary
Set rec_by_customer_id = New Scripting.Dictionary

d = GetData(rTopLeft)
col_CustomerID = FindCol(d, "CustomerID", 1)
col_EmailAddress = FindCol(d, "EmailAddress", 2)
col_EmailTemplate = FindCol(d, "EmailTemplate", 3)
col_InvoiceTemplate = FindCol(d, "InvoiceTemplate", 4)
col_RecipientNameEmail = FindCol(d, "RecipientNameEmail", 5)
col_RecipientNameInvoice = FindCol(d, "RecipientNameInvoice", 5)
col_CompanyName = FindCol(d, "CompanyName", 6)
col_StreetAddress = FindCol(d, "StreetAddress", 7)
col_TownCityPostCode = FindCol(d, "TownCityPostCode", 8)
col_Country = FindCol(d, "Country", 9)
col_Phone = FindCol(d, "Phone", 10)
col_SubCustomerCommission = FindCol(d, "SubCustomerCommission", 11)

Dim i As Long
For i = LBound(d) + 1 To UBound(d)
  Dim cid As String: cid = Trim(d(i, col_CustomerID))
  If cid <> "" Then
    If row_by_customer_id.Exists(cid) Then
      rTopLeft.Parent.Parent.Activate
      rTopLeft.Parent.Activate
      rTopLeft.Rows(i).EntireRow.Select
      Call ErrorExit("DuplicateCustomerID:" + cid)
    End If
    Call row_by_customer_id.Add(cid, i)
  End If
Next i
End Sub

Private Sub Class_Initialize()
' https://stackoverflow.com/questions/201323/how-can-i-validate-an-email-address-using-a-regular-expression
EmailAddressRegExp.Pattern = "(?:[a-z0-9!#$%&'*+/=?^_`{|}~-]+(?:\.[a-z0-9!#$%&'*+/=?^_`{|}~-]+)*|""(?:[\x01-\x08\x0b\x0c\x0e-\x1f\x21\x23-\x5b\x5d-\x7f]|\\[\x01-\x09\x0b\x0c\x0e-\x7f])*"")@(?:(?:[a-z0-9](?:[a-z0-9-]*[a-z0-9])?\.)+[a-z0-9](?:[a-z0-9-]*[a-z0-9])?|\[(?:(?:(2(5[0-5]|[0-4][0-9])|1[0-9][0-9]|[1-9]?[0-9]))\.){3}(?:(2(5[0-5]|[0-4][0-9])|1[0-9][0-9]|[1-9]?[0-9])|[a-z0-9-]*[a-z0-9]:(?:[\x01-\x08\x0b\x0c\x0e-\x1f\x21-\x5a\x53-\x7f]|\\[\x01-\x09\x0b\x0c\x0e-\x7f])+)\])"
INVOICE_TEMPLATE_DEFAULT = "invoice_default.docx"
INVOICE_EMAIL_TEMPLATE_DEFAULT = "email_default.html"

Set wsMaster = SafeGetWorksheet(ThisWorkbook, "Master")

NEXT_INVOICE_NUMBER_txt = ThisWorkbook_Path + "\" + "NEXT_INVOICE_NUMBER.txt"
'Set NEXT_INVOICE_NUMBER_r = wsMaster.UsedRange.Find("NEXT_INVOICE_NUMBER")
'If NEXT_INVOICE_NUMBER_r Is Nothing Then
'  wsMaster.Parent.Activate
'  wsMaster.Activate
'  Call ErrorExit("Missing cell for NEXT_INVOICE_NUMBER")
'End If
'Set NEXT_INVOICE_NUMBER_r = wsMaster.Cells(NEXT_INVOICE_NUMBER_r.row, NEXT_INVOICE_NUMBER_r.Column + 1)
Dim GST_RATE_r As Range: Set GST_RATE_r = wsMaster.UsedRange.Find("GST_RATE")
If GST_RATE_r Is Nothing Then
  wsMaster.Parent.Activate
  wsMaster.Activate
  Call ErrorExit("Missing cell for GST_RATE")
End If
GST_RATE = wsMaster.Cells(GST_RATE_r.row, GST_RATE_r.Column + 1)

End Sub
