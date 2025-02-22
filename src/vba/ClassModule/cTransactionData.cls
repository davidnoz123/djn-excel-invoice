VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "cTransactionData"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

Public col_CustomerID As Long
'Public col_DeliveryDate As Long
Public col_InvoiceDate As Long
'Public col_DESCRIPTION As Long
'Public col_QUANTITY As Long
'Public col_UnitPrice As Long
Public col_ProcessedWhen As Long
Public col_InvoiceNo As Long
Public header

Public FieldName2ColIndex As New Scripting.Dictionary

Public tws As Worksheet
Public colBeg As Long
Public colEnd As Long
'Public MappableTransactionColumnSet As New Scripting.Dictionary

Function GetRecord(rowIndex)
GetRecord = tws.Range(tws.Cells(rowIndex, colBeg), tws.Cells(rowIndex, colEnd))
End Function

Sub Initx(ws As Worksheet)
Set tws = ws
colBeg = tws.UsedRange.Columns(1).Column
colEnd = tws.UsedRange.Columns(tws.UsedRange.Columns.count).Column
header = tws.Range(tws.Cells(1, colBeg), tws.Cells(1, colEnd))

col_CustomerID = FindCol(header, "CustomerID", 1)
'col_DeliveryDate = FindCol(header, "DeliveryDate", 2)
col_InvoiceDate = FindCol(header, "InvoiceDate", 3)
'col_DESCRIPTION = FindCol(header, "Description", 4)
'col_QUANTITY = FindCol(header, "Quantity", 5)
'col_UnitPrice = FindCol(header, "UnitPrice", 6)
col_ProcessedWhen = FindCol(header, "ProcessedWhen", 8)
col_InvoiceNo = FindCol(header, "InvoiceNo", 9)

Set FieldName2ColIndex = New Scripting.Dictionary

Dim i As Long
For i = LBound(header, 2) To UBound(header, 2)
  Call FieldName2ColIndex.Add(header(LBound(header), i), i)
Next i

End Sub

Private Sub Class_Initialize()
'Call MappableTransactionColumnSet.Add("DeliveryDate", Empty)
'Call MappableTransactionColumnSet.Add("Description", Empty)
'Call MappableTransactionColumnSet.Add("Quantity", Empty)
'Call MappableTransactionColumnSet.Add("UnitPrice", Empty)
End Sub
