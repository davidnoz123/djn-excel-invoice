VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "Sheet1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = True
Option Explicit

Dim TempCombo_InitialValue As Variant

Private Sub Worksheet_BeforeDoubleClick(ByVal target As Range, Cancel As Boolean)
If target.row > 1 Then
  Dim ws As Worksheet: Set ws = target.Parent
  Dim is_DeliveryDate As Boolean: is_DeliveryDate = ws.Cells(1, target.Column) = "DeliveryDate"
  If is_DeliveryDate Or ws.Cells(1, target.Column) = "InvoiceDate" Then
    Set DatePickerForm.target = target.Cells(1, 1)
    DatePickerForm.Show vbModal
    Cancel = True
    If is_DeliveryDate And IsDate(target) Then
      Dim r As Range: Set r = ws.Rows(1).EntireRow.Find("InvoiceDate", LookAt:=xlWhole)
      If r Is Nothing Then
        ws.Parent.Activate
        ws.Activate
        ws.Rows(1).EntireRow.Select
        Call ErrorExit("Missing column with header 'InvoiceDate'")
      End If
      Set r = ws.Cells(target.row, r.Column)
      If IsEmpty(r) Then
        Dim tmp: tmp = DateAdd("m", 1, target)
        r = DateAdd("d", -1, DateSerial(Year(tmp), Month(tmp), 1))
      End If
    End If
  End If
End If
End Sub

Private Sub Worksheet_SelectionChange(ByVal target As Range)
Dim xWs As Worksheet
Set xWs = Application.ActiveSheet

On Error Resume Next

Dim xCombox As OLEObject
Set xCombox = xWs.OLEObjects("TempCombo")

' Added this to auto select all text when activating the combox box.
xCombox.SetFocus

With xCombox
    .ListFillRange = vbNullString
    .LinkedCell = vbNullString
    .Visible = False
End With
TempCombo_InitialValue = Empty


Dim xStr As String
Dim xArr


If target.Rows.count = 1 And target.Columns.count = 1 And target.Validation.Type = xlValidateList Then
    ' The target cell contains Data Validation.

    target.Validation.InCellDropdown = False


    ' Cancel the "SelectionChange" event.
    Dim Cancel As Boolean
    Cancel = True


    xStr = target.Validation.Formula1
    xStr = Right(xStr, Len(xStr) - 1)

    If xStr = vbNullString Then Exit Sub

    With xCombox
        .Visible = True
        .Left = target.Left
        .Top = target.Top
        .Width = target.Width + 5
        .Height = target.Height + 5
        .ListFillRange = xStr

        If .ListFillRange = vbNullString Then
            xArr = Split(xStr, ",")
            Me.TempCombo.List = xArr
        End If

        .LinkedCell = target.Address
        
        TempCombo_InitialValue = target.Value

    End With

    xCombox.Activate
    Me.TempCombo.DropDown

End If
End Sub













