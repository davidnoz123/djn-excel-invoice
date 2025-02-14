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
'Update by Extendoffice: 2018/9/21
' Update by Chris Brackett 2018-11-30

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

Private Sub TempComboHandleChange(ActiveCell_ As Range)
On Error GoTo 99
Dim r As Range: Set r = Application.Evaluate(ActiveCell_.Validation.Formula1)
Dim f As Range: Set f = r.Find(ActiveCell_.Value, LookIn:=xlValues, LookAt:=xlWhole)
If f Is Nothing Then
  Call MsgBox("Unrecognised Value:" + CStr(ActiveCell_.Value), vbCritical + vbOKOnly, "Error")
  ActiveCell_.Value = TempCombo_InitialValue
Else
  If Not IsEmpty(TempCombo_InitialValue) Then
    If TempCombo_InitialValue <> ActiveCell_.Value Then
      Dim ret As Long: ret = MsgBox("Change from:" + CStr(TempCombo_InitialValue), vbQuestion + vbOKCancel + vbDefaultButton2, "Confirm Change")
      If ret = vbCancel Then
        Debug.Print ActiveCell_.Value
        ActiveCell_.Value = TempCombo_InitialValue
      End If
    End If
  End If
End If
99:
End Sub

Private Sub TempCombo_KeyDown( _
                ByVal KeyCode As MSForms.ReturnInteger, _
                ByVal Shift As Integer)
    Select Case KeyCode
        Case 9  ' Tab key
            Call TempComboHandleChange(Application.ActiveCell)
            Application.ActiveCell.Offset(0, 1).Activate
        Case 13 ' Pause key
            Call TempComboHandleChange(Application.ActiveCell)
            Application.ActiveCell.Offset(1, 0).Activate
    End Select
End Sub







