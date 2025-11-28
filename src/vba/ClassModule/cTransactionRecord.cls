VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "cTransactionRecord"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

Public rowIndex As Long
Public CustomerRecord As cCustomerRecord

Public CustomerID As String
Public invoiceDate As Date
Public ProcessedWhen As String
Public InvoiceNo As Long

Public FieldName2Value As New Scripting.Dictionary

Public Parent As cTransactionData

Private Sub x_error_exit(row As Long, col As Long, msg As String, Optional SelectRow = False)
Parent.tws.Parent.Activate
Parent.tws.Activate
If SelectRow Then
  Parent.tws.Range(Parent.tws.Cells(row, Parent.colBeg), Parent.tws.Cells(row, Parent.colEnd)).Select
Else
  Parent.tws.Cells(row, col).Select
End If
Call ErrorExit(msg)
End Sub

Sub Inity(main_ As cCustomerData, td As cTransactionData, rowIndex_)
Set Parent = td
rowIndex = rowIndex_
Dim rec: rec = td.GetRecord(rowIndex)
Me.CustomerID = Trim(rec(LBound(rec), td.col_CustomerID))
If Me.CustomerID <> "" Then
  Dim mr As cCustomerRecord: Set mr = main_.RecordsByCustomerID(Me.CustomerID)
  If mr Is Nothing Then
    Call x_error_exit(rowIndex, td.col_CustomerID, "Unknown CustomerID '" + CStr(Me.CustomerID) + "'", True)
  End If
  Set Me.CustomerRecord = mr
  Dim vInvoiceDate As Variant: vInvoiceDate = rec(LBound(rec), td.col_InvoiceDate)
  If IsEmpty(vInvoiceDate) Or Not IsDate(vInvoiceDate) Then
    Call x_error_exit(rowIndex, td.col_InvoiceDate, "Missing or invalid InvoiceDate in transaction record. Provide dates in the form similar to '2021-07-21'")
  End If
  
  Me.invoiceDate = CDate(vInvoiceDate)
  Me.ProcessedWhen = rec(LBound(rec), td.col_ProcessedWhen)
  Me.InvoiceNo = rec(LBound(rec), td.col_InvoiceNo)
  
  Set FieldName2Value = New Scripting.Dictionary
  Dim field_name
  For Each field_name In Parent.FieldName2ColIndex.Keys()
      Dim col As Long: col = Parent.FieldName2ColIndex(field_name)
      Call FieldName2Value.Add(field_name, rec(LBound(rec), col))
  Next field_name
  
End If
End Sub




