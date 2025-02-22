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
Public col_InvoiceDate As Long
Public col_ProcessedWhen As Long
Public col_InvoiceNo As Long
Public header

Public FieldName2ColIndex As New Scripting.Dictionary

Public tws As Worksheet
Public colBeg As Long
Public colEnd As Long

Function GetRecord(rowIndex)
GetRecord = tws.Range(tws.Cells(rowIndex, colBeg), tws.Cells(rowIndex, colEnd))
End Function

Sub Initx(ws As Worksheet)
Set tws = ws
colBeg = tws.UsedRange.Columns(1).Column
colEnd = tws.UsedRange.Columns(tws.UsedRange.Columns.count).Column
header = tws.Range(tws.Cells(1, colBeg), tws.Cells(1, colEnd))

col_CustomerID = FindCol(header, "CustomerID")
col_InvoiceDate = FindCol(header, "InvoiceDate")
col_ProcessedWhen = FindCol(header, "ProcessedWhen")
col_InvoiceNo = FindCol(header, "InvoiceNo")

Set FieldName2ColIndex = New Scripting.Dictionary

Dim i As Long
For i = LBound(header, 2) To UBound(header, 2)
  Call FieldName2ColIndex.Add(header(LBound(header), i), i)
Next i

End Sub

