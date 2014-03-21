VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} ufXAx 
   Caption         =   "X軸タイトルの選択"
   ClientHeight    =   2625
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   6420
   OleObjectBlob   =   "ufXAx.frx":0000
   StartUpPosition =   1  'オーナー フォームの中央
End
Attribute VB_Name = "ufXAx"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Const app As String = "Excel"
Private Const sec As String = "Labels"

Public Function Label() As String
 Label = GetSetting(app, sec, "x", "")
End Function


Private Sub setX(Label As String)
  SaveSetting app, sec, "oldx", GetSetting(app, sec, "x", "")
  SaveSetting app, sec, "x", Label
  Unload Me
End Sub

Private Sub cp(ByVal Button As Integer)
 If Button = xlSecondaryButton Then
   Me.TextBox1.Text = Me.ActiveControl.Caption
   Me.TextBox1.SetFocus
End If
End Sub

Private Sub setc()
  setX Me.ActiveControl.Caption

End Sub

Private Sub CommandButton1_Click()
  setc
End Sub

Private Sub CommandButton1_MouseDown(ByVal Button As Integer, ByVal Shift As Integer, ByVal X As Single, ByVal Y As Single)
cp Button
End Sub

Private Sub CommandButton11_Click()
setc
End Sub

Private Sub CommandButton11_MouseDown(ByVal Button As Integer, ByVal Shift As Integer, ByVal X As Single, ByVal Y As Single)
cp Button
End Sub

Private Sub CommandButton12_Click()
setc
End Sub

Private Sub CommandButton12_MouseDown(ByVal Button As Integer, ByVal Shift As Integer, ByVal X As Single, ByVal Y As Single)
cp Button
End Sub

Private Sub CommandButton13_Click()
setc
End Sub

Private Sub CommandButton13_MouseDown(ByVal Button As Integer, ByVal Shift As Integer, ByVal X As Single, ByVal Y As Single)
cp Button
End Sub

Private Sub CommandButton14_Click()
setc
End Sub

Private Sub CommandButton14_MouseDown(ByVal Button As Integer, ByVal Shift As Integer, ByVal X As Single, ByVal Y As Single)
cp Button
End Sub

Private Sub CommandButton15_Click()
 setX ""
End Sub

Private Sub CommandButton16_Click()
setc
End Sub

Private Sub CommandButton16_MouseDown(ByVal Button As Integer, ByVal Shift As Integer, ByVal X As Single, ByVal Y As Single)
cp Button
End Sub

Private Sub CommandButton3_Click()
 setX Me.TextBox1.Text
End Sub



Private Sub CommandButton5_Click()
setc
End Sub

Private Sub CommandButton5_KeyDown(ByVal KeyCode As MSForms.ReturnInteger, ByVal Shift As Integer)
cp Button
End Sub

Private Sub CommandButton6_Click()
setc
End Sub

Private Sub CommandButton6_MouseDown(ByVal Button As Integer, ByVal Shift As Integer, ByVal X As Single, ByVal Y As Single)
cp Button
End Sub

Private Sub CommandButton8_Click()
setc
End Sub

Private Sub CommandButton8_MouseDown(ByVal Button As Integer, ByVal Shift As Integer, ByVal X As Single, ByVal Y As Single)
cp Button
End Sub

Private Sub UserForm_Initialize()
 Me.TextBox1.Text = GetSetting(app, sec, "x", "")
 If Me.TextBox1.Text = "" Then Me.TextBox1.Text = GetSetting(app, sec, "oldx", "")
 Me.TextBox1.SetFocus
End Sub
