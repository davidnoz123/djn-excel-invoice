VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "cCustomerData"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

Public INVOICE_TEMPLATE_DEFAULT As String
Public INVOICE_EMAIL_TEMPLATE_DEFAULT As String

Const INVOICE_EMAIL_TEMPLATE_DEFAULT_CONTENT = "<div>" + vbCrLf + _
  "<p>Dear {RecipientName}</p>" + vbCrLf + _
  "<p>Please find the attached invoice.</p>" + vbCrLf + _
  "<p>Best Wishes</p>" + vbCrLf + _
  "<p>{YourBusinessName}</p>" + vbCrLf + _
"</div>"


Dim wsMain As Worksheet
Public fso As New FileSystemObject
Public TopLeft As Range

Public col_CustomerID As Long
Public col_EmailAddress As Long
Public col_EmailTemplate As Long
Public col_InvoiceTemplate As Long

Public FieldName2ColIndex As New Scripting.Dictionary

Public SUBJECT_LINE_PREFIX As String
Public YOUR_BUSINESS_NAME As String

Dim NEXT_INVOICE_NUMBER_txt As String
Public INVOICE_COLUMN_2_TRANSACTION_COLUMN_MAP As New Scripting.Dictionary

Dim d
Dim row_by_customer_id As Scripting.Dictionary
Dim rec_by_customer_id As Scripting.Dictionary

Public EmailAddressRegExp As New VBScript_RegExp_55.RegExp

Function NextInvoiceNumber() As Long
' NOTE: We could include this state in the Excel spreadsheet but it would mean the workbook needs to be saved after each run (which may not always happen)
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

Property Get RecordsByCustomerID(CustomerID) As cCustomerRecord
Dim r As cCustomerRecord
If rec_by_customer_id.Exists(CustomerID) Then
  Set r = rec_by_customer_id(CustomerID)
Else
  If Not row_by_customer_id.Exists(CustomerID) Then
    Set r = Nothing
  Else
    Set r = New cCustomerRecord: Call r.Init(Me, d, row_by_customer_id(CustomerID))
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
col_CustomerID = FindCol(d, "CustomerID")
col_EmailAddress = FindCol(d, "EmailAddress")
col_EmailTemplate = FindCol(d, "EmailTemplate")
col_InvoiceTemplate = FindCol(d, "InvoiceTemplate")

Set FieldName2ColIndex = New Scripting.Dictionary

Dim i As Long
For i = LBound(d, 2) To UBound(d, 2)
  Call FieldName2ColIndex.Add(d(LBound(d), i), i)
Next i

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
INVOICE_EMAIL_TEMPLATE_DEFAULT = ThisWorkbook.Path + "\" + "email_default.html"
Dim fso As New FileSystemObject
If Not fso.FileExists(INVOICE_EMAIL_TEMPLATE_DEFAULT) Then
    Dim ts As TextStream
    Set ts = fso.CreateTextFile(INVOICE_EMAIL_TEMPLATE_DEFAULT)
    Call ts.WriteLine(INVOICE_EMAIL_TEMPLATE_DEFAULT_CONTENT)
    ts.Close
End If

Set wsMain = SafeGetWorksheet(ThisWorkbook, "Main")

YOUR_BUSINESS_NAME = SafeFindKeyValueInRange(wsMain.UsedRange, "YOUR_BUSINESS_NAME", "Your Business")
SUBJECT_LINE_PREFIX = SafeFindKeyValueInRange(wsMain.UsedRange, "SUBJECT_LINE_PREFIX", YOUR_BUSINESS_NAME + " Invoice")
NEXT_INVOICE_NUMBER_txt = ThisWorkbook_Path + "\" + SafeFindKeyValueInRange(wsMain.UsedRange, "NEXT_INVOICE_NUMBER_FILENAME", "NEXT_INVOICE_NUMBER.txt")

Dim td As New cTransactionData

Dim tmp As New Scripting.Dictionary
Dim r As Range: Set r = Nothing
While True
    If r Is Nothing Then
        Set r = wsMain.UsedRange.Find("INVOICE_COLUMN_2_TRANSACTION_COLUMN_MAP")
    Else
        Set r = wsMain.UsedRange.Find("INVOICE_COLUMN_2_TRANSACTION_COLUMN_MAP", After:=r)
    End If
    If r Is Nothing Then GoTo exit_while
    If tmp.Exists(r.AddressLocal) Then
        GoTo exit_while
    End If
    Call tmp.Add(r.AddressLocal, r)
    Dim rv As Range: Set rv = wsMain.Cells(r.row, r.Column + 1)
    Dim vv As Variant: vv = rv
    If VarType(vv) = vbString Then
        If Trim(vv) <> "" Then
            Dim a: a = Split(vv, ";")
            Dim key As String, val As String
            If UBound(a) - LBound(a) <> 1 Then
                rv.Parent.Parent.Activate
                rv.Parent.Activate
                rv.Select
                Call ErrorExit("INVOICE_COLUMN_2_TRANSACTION_COLUMN_MAP value is not a semicolon delimited pair of strings:'" + vv + "'")
            End If
            key = a(0)
            val = a(1)
            'If Not td.MappableTransactionColumnSet.Exists(val) Then
            '    rv.Parent.Parent.Activate
            '    rv.Parent.Activate
            '    rv.Select
            '    Dim aa: aa = c2a(td.MappableTransactionColumnSet)
            '    Call ErrorExit("Unexpected TRANSACTION_COLUMN:'" + val + "' Expecting one of:'" + Join(aa, "', '") + "'")
            'End If
            If INVOICE_COLUMN_2_TRANSACTION_COLUMN_MAP.Exists(key) Then
                If INVOICE_COLUMN_2_TRANSACTION_COLUMN_MAP(key) <> val Then
                    rv.Parent.Parent.Activate
                    rv.Parent.Activate
                    rv.Select
                    Call ErrorExit("INVOICE_COLUMN string already mapped to TRANSACTION_COLUMN:'" + key + "' '" + INVOICE_COLUMN_2_TRANSACTION_COLUMN_MAP(key) + "'")
                End If
            End If
            Call INVOICE_COLUMN_2_TRANSACTION_COLUMN_MAP.Add(key, val)
        End If
    End If
    '
Wend
exit_while:

End Sub
