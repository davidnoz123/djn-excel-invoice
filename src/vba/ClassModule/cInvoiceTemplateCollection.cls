VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "cInvoiceTemplateCollection"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit


Dim fso As New FileSystemObject
Dim fb2it As New Scripting.Dictionary

Function GetInvoiceTemplate(template_name As String) As cInvoiceTemplateXl
If fb2it.Exists(template_name) Then
  Set GetInvoiceTemplate = fb2it(template_name)
Else
  Dim it As New cInvoiceTemplateXl
  Call it.Init(template_name)
  Call fb2it.Add(template_name, it)
  Set GetInvoiceTemplate = it
End If
End Function


