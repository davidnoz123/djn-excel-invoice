VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "cCustomerRecord"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

Public CustomerID As String

Public FieldName2Value As New Scripting.Dictionary

Dim EmailAddress_v As Variant
Dim EmailAddress_s As String

Dim EmailTemplate_v As Variant
Dim EmailTemplate_s As String
Dim EmailTemplate_is_default As Boolean

Dim InvoiceTemplate_v As Variant
Dim InvoiceTemplate_s As String
Dim InvoiceTemplate_is_default As Boolean

Dim rowIndex As Long

Public Parent As cCustomerData

Property Get EmailAddressx() As String
If IsEmpty(EmailAddress_v) Then
  EmailAddress_v = EmailAddress_s
  If EmailAddress_s <> "" Then
    If False Then
      Dim mc As VBScript_RegExp_55.MatchCollection
      Set mc = Parent.EmailAddressRegExp.Execute(EmailAddress_s)
      If mc Is Nothing Then
        Call x_error_exit(rowIndex, Parent.col_EmailAddress, "Invalid Email Address '" + EmailAddress_s + "'")
      End If
    End If
  End If
End If
EmailAddressx = EmailAddress_v
End Property

Property Get EmailTemplatex() As String
If IsEmpty(EmailTemplate_v) Then
  If Parent.fso.FileExists(EmailTemplate_s) Then
    EmailTemplate_v = Parent.fso.GetAbsolutePathName(EmailTemplate_s)
  Else
    EmailTemplate_v = Parent.fso.GetAbsolutePathName(ThisWorkbook_Path + "\" + EmailTemplate_s)
  End If
  If Not Parent.fso.FileExists(EmailTemplate_v) Then
    Call x_error_exit(rowIndex, Parent.col_EmailTemplate, "Missing " + IIf(EmailTemplate_is_default, "default ", "") + "EmailTemplate file '" + EmailTemplate_v + "'")
  End If
End If
EmailTemplatex = EmailTemplate_v
End Property

Property Get InvoiceTemplatex() As String
If True Then
    InvoiceTemplatex = InvoiceTemplate_s
Else
    If IsEmpty(InvoiceTemplate_v) Then
      If Parent.fso.FileExists(InvoiceTemplate_s) Then
        InvoiceTemplate_v = Parent.fso.GetAbsolutePathName(InvoiceTemplate_s)
      Else
        InvoiceTemplate_v = Parent.fso.GetAbsolutePathName(ThisWorkbook_Path + "\" + InvoiceTemplate_s)
      End If
      If Not Parent.fso.FileExists(InvoiceTemplate_v) Then
        Call x_error_exit(rowIndex, Parent.col_InvoiceTemplate, "Missing " + IIf(InvoiceTemplate_is_default, "default ", "") + "InvoiceTemplate file '" + InvoiceTemplate_v + "'")
      End If
    End If
    InvoiceTemplatex = InvoiceTemplate_v
End If
End Property

Private Sub x_error_exit(row As Long, col As Long, msg As String)
Parent.TopLeft.Parent.Parent.Activate
Parent.TopLeft.Parent.Activate
Parent.TopLeft.Cells(row, col).Select
Call ErrorExit(msg)
End Sub

Sub Init(Parent_ As cCustomerData, d, row As Long)
Set Parent = Parent_
rowIndex = row

CustomerID = d(row, Parent.col_CustomerID)
EmailAddress_s = Trim(d(row, Parent.col_EmailAddress))
EmailTemplate_s = Trim(d(row, Parent.col_EmailTemplate))
InvoiceTemplate_s = Trim(d(row, Parent.col_InvoiceTemplate))

Set FieldName2Value = New Scripting.Dictionary
Dim field_name
For Each field_name In Parent_.FieldName2ColIndex.Keys()
    Dim col As Long: col = Parent_.FieldName2ColIndex(field_name)
    Call FieldName2Value.Add(field_name, d(row, col))
Next field_name

EmailAddress_v = Empty

InvoiceTemplate_v = Empty
InvoiceTemplate_is_default = (InvoiceTemplate_s = "")
If InvoiceTemplate_is_default Then
  InvoiceTemplate_s = Parent_.INVOICE_TEMPLATE_DEFAULT
End If

EmailTemplate_v = Empty
EmailTemplate_is_default = (EmailTemplate_s = "")
If EmailTemplate_is_default Then
  EmailTemplate_s = Parent_.INVOICE_EMAIL_TEMPLATE_DEFAULT
End If

End Sub
