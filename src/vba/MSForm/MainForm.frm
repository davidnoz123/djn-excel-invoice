VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} MainForm 
   Caption         =   "Generate Invoices"
   ClientHeight    =   1980
   ClientLeft      =   105
   ClientTop       =   450
   ClientWidth     =   11565
   OleObjectBlob   =   "MainForm.frx":0000
   StartUpPosition =   1  'CenterOwner
End
Attribute VB_Name = "MainForm"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False





Option Explicit

Dim transaction_records_ As Collection
Public log_prefix As String

Sub Do_Events(msg)
Me.label.Caption = log_prefix + msg
DoEvents
If ToggleState.Caption = "Stopping" Then
  Me.Hide
  End
End If
End Sub

Private Sub ToggleState_Click()
If ToggleState.Caption = "Start" Then
  ToggleState.Caption = "Stop"
  Call aProcessRun_Callback(transaction_records_)
  ToggleState.Caption = "Exit"
  Me.label.Caption = "Complete"
ElseIf ToggleState.Caption = "Stop" Then
  ToggleState.Caption = "Stopping"
  ToggleState.Enabled = False
Else
  Me.Hide
End If
End Sub

Sub ResetState(transaction_records As Collection)
Set transaction_records_ = transaction_records
log_prefix = ""
ToggleState.Enabled = True
If transaction_records.count > 0 Then
  ToggleState.Caption = "Start"
  Me.label.Caption = "Ready to process " + Str(transaction_records.count) + " transaction records"
Else
  ToggleState.Caption = "Exit"
  Me.label.Caption = "No Records Selected"
End If
End Sub

Private Sub UserForm_Initialize()
EmailOptionNone.Value = True
End Sub
