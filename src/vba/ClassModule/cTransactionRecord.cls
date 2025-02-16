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
Public MainRecord As cMainRecord

Public CustomerID As String
Public InvoiceDate As Date
Public DeliveryDate As Date
Public Description As String
Public Quantity As Double
Public UnitPrice As Double
Public ProcessedWhen As String
Public InvoiceNo As Long

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

Sub Inity(main_ As cMainData, td As cTransactionData, rowIndex_)
Set Parent = td
rowIndex = rowIndex_
Dim rec: rec = td.GetRecord(rowIndex)
Me.CustomerID = Trim(rec(LBound(rec), td.col_CustomerID))
If Me.CustomerID <> "" Then
  Dim mr As cMainRecord: Set mr = main_.RecordsByCustomerID(Me.CustomerID)
  If mr Is Nothing Then
    Call x_error_exit(rowIndex, td.col_CustomerID, "Unknown CustomerID '" + CStr(Me.CustomerID) + "'", True)
  End If
  Set Me.MainRecord = mr
  Dim vInvoiceDate As Variant: vInvoiceDate = rec(LBound(rec), td.col_InvoiceDate)
  If IsEmpty(vInvoiceDate) Or Not IsDate(vInvoiceDate) Then
    Call x_error_exit(rowIndex, td.col_InvoiceDate, "Missing or invalid InvoiceDate in transaction record. Provide dates in the form similar to '2021-07-21'")
  End If
  Dim vDeliveryDate As Variant: vDeliveryDate = rec(LBound(rec), td.col_DeliveryDate)
  If IsEmpty(vDeliveryDate) Or Not IsDate(vDeliveryDate) Then
    Call x_error_exit(rowIndex, td.col_DeliveryDate, "Missing or invalid DeliveryDate in transaction record. Provide dates in the form similar to '2021-07-21'")
  End If
  
  Me.InvoiceDate = CDate(vInvoiceDate)
  Me.DeliveryDate = CDate(vDeliveryDate)
  Me.Description = rec(LBound(rec), td.col_DESCRIPTION)
  Me.Quantity = rec(LBound(rec), td.col_QUANTITY)
  Me.UnitPrice = rec(LBound(rec), td.col_UnitPrice)
  Me.ProcessedWhen = rec(LBound(rec), td.col_ProcessedWhen)
  Me.InvoiceNo = rec(LBound(rec), td.col_InvoiceNo)
  
End If
End Sub




