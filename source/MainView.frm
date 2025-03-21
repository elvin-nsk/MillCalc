VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} MainView 
   Caption         =   "MillCalc"
   ClientHeight    =   8040
   ClientLeft      =   45
   ClientTop       =   375
   ClientWidth     =   5625
   OleObjectBlob   =   "MainView.frx":0000
   StartUpPosition =   1  'CenterOwner
End
Attribute VB_Name = "MainView"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Option Explicit

'===============================================================================

Private Sub UserForm_Initialize()
  Caption = APP_DISPLAYNAME & " (v" & APP_VERSION & ")"
  Me.lbOutsideElements.Visible = False
End Sub

Private Sub btnOK_Click()
  FormCancel
End Sub

'===============================================================================

Private Sub FormОК()
  Me.Hide
End Sub

Private Sub FormCancel()
  Me.Hide
End Sub

'===============================================================================

Private Sub OnlyInt(ByVal KeyAscii As MSForms.ReturnInteger)
  Select Case KeyAscii
    Case Asc("0") To Asc("9")
    Case Else
      KeyAscii = 0
  End Select
End Sub

Private Sub OnlyNum(ByVal KeyAscii As MSForms.ReturnInteger)
  Select Case KeyAscii
    Case Asc("0") To Asc("9")
    Case Asc(",")
    Case Else
      KeyAscii = 0
  End Select
End Sub

Private Sub CheckRangeDbl(TextBox As MSForms.TextBox, ByVal Min As Double, Optional ByVal Max As Double = 2147483647)
  With TextBox
    If .Value = "" Then .Value = CStr(Min)
    If CDbl(.Value) > Max Then .Value = CStr(Max)
    If CDbl(.Value) < Min Then .Value = CStr(Min)
  End With
End Sub

Private Sub CheckRangeLng(TextBox As MSForms.TextBox, ByVal Min As Long, Optional ByVal Max As Long = 2147483647)
  With TextBox
    If .Value = "" Then .Value = CStr(Min)
    If CLng(.Value) > Max Then .Value = CStr(Max)
    If CLng(.Value) < Min Then .Value = CStr(Min)
  End With
End Sub

Private Sub UserForm_QueryClose(Сancel As Integer, CloseMode As Integer)
  If CloseMode = VbQueryClose.vbFormControlMenu Then
    Сancel = True
    FormCancel
  End If
End Sub
