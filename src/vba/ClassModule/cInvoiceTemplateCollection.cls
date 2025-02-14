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

Function GetInvoiceTemplate(file_base) As cInvoiceTemplate
file_base = LCase(fso.GetFileName(file_base))
If fb2it.Exists(file_base) Then
  Set GetInvoiceTemplate = fb2it(file_base)
Else
  Dim it As New cInvoiceTemplate
  Call it.Init(ThisWorkbook_Path + "\" + file_base)
  Call fb2it.Add(file_base, it)
  Set GetInvoiceTemplate = it
End If
End Function


