VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} ufYAx 
   Caption         =   "Y軸タイトルの選択"
   ClientHeight    =   5730
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   6420
   OleObjectBlob   =   "ufYAx.frx":0000
   StartUpPosition =   1  'オーナー フォームの中央
End
Attribute VB_Name = "ufYAx"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Const app As String = "Excel"
Private Const sec As String = "Labels"

Public Function Label() As String
 Label = GetSetting(app, sec, "y", "")
End Function



Private Sub setX(Label As String)
  SaveSetting app, sec, "oldy", GetSetting(app, sec, "y", "")
  SaveSetting app, sec, "y", Label
  SaveSetting app, sec, "rotate", CStr(Me.CheckBox1.Value)
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

Private Sub cbCancel_Click()
 setX ""
End Sub

Private Sub cbOK_Click()
 setX Me.TextBox1.Text
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

Private Sub CommandButton16_Click()
setc
End Sub

Private Sub CommandButton16_MouseDown(ByVal Button As Integer, ByVal Shift As Integer, ByVal X As Single, ByVal Y As Single)
cp Button
End Sub

Private Sub CommandButton17_Click()
setc
End Sub

Private Sub CommandButton17_MouseDown(ByVal Button As Integer, ByVal Shift As Integer, ByVal X As Single, ByVal Y As Single)
cp Button
End Sub

Private Sub CommandButton18_Click()
 setc
End Sub

Private Sub CommandButton18_MouseDown(ByVal Button As Integer, ByVal Shift As Integer, ByVal X As Single, ByVal Y As Single)
cp Button
End Sub

Private Sub CommandButton19_Click()
setc
End Sub

Private Sub CommandButton19_MouseDown(ByVal Button As Integer, ByVal Shift As Integer, ByVal X As Single, ByVal Y As Single)
cp Button
End Sub

Private Sub CommandButton20_Click()
setc
End Sub

Private Sub CommandButton20_MouseDown(ByVal Button As Integer, ByVal Shift As Integer, ByVal X As Single, ByVal Y As Single)
cp Button
End Sub

Private Sub CommandButton21_Click()
setc
End Sub

Private Sub CommandButton21_MouseDown(ByVal Button As Integer, ByVal Shift As Integer, ByVal X As Single, ByVal Y As Single)
cp Button
End Sub

Private Sub CommandButton22_Click()
setc
End Sub

Private Sub CommandButton22_MouseDown(ByVal Button As Integer, ByVal Shift As Integer, ByVal X As Single, ByVal Y As Single)
cp Button
End Sub

Private Sub CommandButton23_Click()
setc
End Sub

Private Sub CommandButton23_MouseDown(ByVal Button As Integer, ByVal Shift As Integer, ByVal X As Single, ByVal Y As Single)
cp Button
End Sub

Private Sub CommandButton24_Click()
setc
End Sub

Private Sub CommandButton24_MouseDown(ByVal Button As Integer, ByVal Shift As Integer, ByVal X As Single, ByVal Y As Single)
cp Button
End Sub

Private Sub CommandButton25_Click()
setc
End Sub

Private Sub CommandButton25_MouseDown(ByVal Button As Integer, ByVal Shift As Integer, ByVal X As Single, ByVal Y As Single)
cp Button
End Sub

Private Sub CommandButton26_Click()
setc
End Sub

Private Sub CommandButton26_MouseDown(ByVal Button As Integer, ByVal Shift As Integer, ByVal X As Single, ByVal Y As Single)
cp Button

End Sub

Private Sub CommandButton27_Click()
setc
End Sub

Private Sub CommandButton27_MouseDown(ByVal Button As Integer, ByVal Shift As Integer, ByVal X As Single, ByVal Y As Single)
cp Button
End Sub

Private Sub CommandButton28_Click()
setc
End Sub

Private Sub CommandButton28_MouseDown(ByVal Button As Integer, ByVal Shift As Integer, ByVal X As Single, ByVal Y As Single)
cp Button
End Sub

Private Sub CommandButton29_Click()
cp xlSecondaryButton
End Sub

Private Sub CommandButton29_MouseDown(ByVal Button As Integer, ByVal Shift As Integer, ByVal X As Single, ByVal Y As Single)
cp Button
End Sub

Private Sub CommandButton30_Click()
setc
End Sub

Private Sub CommandButton30_MouseDown(ByVal Button As Integer, ByVal Shift As Integer, ByVal X As Single, ByVal Y As Single)
cp Button
End Sub

Private Sub CommandButton31_Click()
setc
End Sub

Private Sub CommandButton31_MouseDown(ByVal Button As Integer, ByVal Shift As Integer, ByVal X As Single, ByVal Y As Single)
cp Button
End Sub

Private Sub CommandButton32_Click()
setc
End Sub

Private Sub CommandButton32_MouseDown(ByVal Button As Integer, ByVal Shift As Integer, ByVal X As Single, ByVal Y As Single)
cp Button
End Sub

Private Sub CommandButton33_Click()
setc
End Sub

Private Sub CommandButton33_MouseDown(ByVal Button As Integer, ByVal Shift As Integer, ByVal X As Single, ByVal Y As Single)
cp Button
End Sub

Private Sub CommandButton34_Click()
setc
End Sub

Private Sub CommandButton34_MouseDown(ByVal Button As Integer, ByVal Shift As Integer, ByVal X As Single, ByVal Y As Single)
cp Button
End Sub

Private Sub CommandButton35_Click()
setc
End Sub

Private Sub CommandButton35_MouseDown(ByVal Button As Integer, ByVal Shift As Integer, ByVal X As Single, ByVal Y As Single)
cp Button
End Sub

Private Sub CommandButton36_Click()
setc
End Sub

Private Sub CommandButton36_MouseDown(ByVal Button As Integer, ByVal Shift As Integer, ByVal X As Single, ByVal Y As Single)
cp Button
End Sub

Private Sub CommandButton37_Click()
setc
End Sub

Private Sub CommandButton37_MouseDown(ByVal Button As Integer, ByVal Shift As Integer, ByVal X As Single, ByVal Y As Single)
cp Button
End Sub

Private Sub CommandButton38_Click()
setc
End Sub

Private Sub CommandButton38_MouseDown(ByVal Button As Integer, ByVal Shift As Integer, ByVal X As Single, ByVal Y As Single)
cp Button
End Sub

Private Sub CommandButton39_Click()
setc
End Sub

Private Sub CommandButton39_MouseDown(ByVal Button As Integer, ByVal Shift As Integer, ByVal X As Single, ByVal Y As Single)
cp Button
End Sub

Private Sub CommandButton5_Click()
setc
End Sub

Private Sub CommandButton5_MouseDown(ByVal Button As Integer, ByVal Shift As Integer, ByVal X As Single, ByVal Y As Single)
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
 Me.TextBox1.Text = GetSetting(app, sec, "y", "")
 If Me.TextBox1.Text = "" Then Me.TextBox1.Text = GetSetting(app, sec, "oldy", "")
 If GetSetting(app, sec, "rotate", "False") = "True" Then Me.CheckBox1.Value = True
 Me.TextBox1.SetFocus
End Sub

