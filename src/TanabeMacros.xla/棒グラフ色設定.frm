VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} �_�O���t�F�ݒ� 
   Caption         =   "�_�O���t���̓h�ׂ��F�ݒ�"
   ClientHeight    =   4845
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   7695
   OleObjectBlob   =   "�_�O���t�F�ݒ�.frx":0000
   StartUpPosition =   1  '�I�[�i�[ �t�H�[���̒���
End
Attribute VB_Name = "�_�O���t�F�ݒ�"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
''' �e�N�X�`���`��������̃t�H�[�}�b�g
'
'  ;�ōs�ɋ�؂���D
' �͂��߂̍s�͐F���D
' �c��̍s�͐F�ݒ�D
' �s�̓J���}��5��ɋ�؂���D
' ��1��F �p�^�[��  msoFillType
' ��2��F �O�ʂ̐F �iSchemeColor)
' ��3��F �w�ʂ̐F  (SchemeColor)
' ��4��F �p�^�[���̈�������1
' ��5��F �p�^�[���̈�������2
'
' ��4�C5��̒l�͑�1��̃p�^�[���̎�ނɈˑ�����D

' �p�^�[���̔ԍ��Ǝ�ނ̊֌W�͈ȉ��̂Ƃ���
' 1: �\���b�h�h��Ԃ�
' 2: �p�^�[��
' 3: �O���f�[�V����
' 4: �e�N�X�`��
' 5: �G�N�Z���ł͎g�p�ł��Ȃ� �i�o�b�N�O���E���h�j
' 6: �s�N�`�� �t�@�C�������K�v�Ȃ̂Ŏg�p�s�\

' �O���f�[�V�����ł͏ꏊ������Ȃ��̂ŁC
' �O���f�[�V�����̐ݒ�(���̍s)��:�ŋ�؂��đ�4��ɐݒ肷��D
'    �^�C�v�F�p�x�F�X�^�C���F�o���G�[�V�����F�v���Z�b�g

'    g_type = CLng(a(0))
'    g_degree = CLng(a(1))
'    g_style = CLng(a(2))
'    g_variant = CLng(a(3))
'    g_preset = CLng(a(4))


'' �Q�l�I�u�W�F�N�g�u���E�U�̒l
'Const msoFillSolid = 1
'Const msoFillGradient = 3
'Const msoFillMixed = -2 (&HFFFFFFFE)
'Const msoFillPatterned = 2
'Const msoFillPicture = 6
'Const msoFillTextured = 4
'Const msoFillBackground = 5


' �����̃e�N�X�`���`��
 Const pattern_bw As String = "14;2,1,2,10,0;2,1,2,26,0;2,1,2,4,0;2,1,2,14,0;2,1,2,38,0;2,1,2,23,0;2,1,2,2,0;2,1,2,13,0;2,1,2,6,0;2,1,2,31,0;2,1,2,33,0;2,1,2,42,0;2,1,2,17,0;2,1,2,39,0"
 Const pattern_bw20 As String = "20;2,1,2,10,0;2,1,2,26,0;2,1,2,4,0;2,1,2,14,0;2,1,2,38,0;2,1,2,23,0;2,1,2,2,0;2,1,2,13,0;2,1,2,6,0;2,1,2,31,0;2,1,2,33,0;2,1,2,42,0;2,1,2,17,0;2,1,2,39,0;1,15,2,0,0;1,48,2,0,0;1,16,2,0,0;1,56,2,0,0;1,1,2,0,0;2,1,2,24,0"
 Const pattern_gray As String = "14;1,1,1,0,0;2,1,2,12,0;1,56,1,0,0;2,1,2,10,0;2,1,2,9,0;2,1,2,8,0;1,16,1,0,0;1,48,1,0,0;2,1,2,5,0;1,15,1,0,0;2,1,2,3,0;2,1,2,2,0;2,1,2,1,0;1,2,1,0,0"
'Private Const pattern_enquete As String = "5;2,1,2,6,0;2,1,2,3,0;2,1,2,19,0;2,1,2,16,0;1,1,1,0,0"
 Const pattern_enquete As String = "5;2,1,2,6,0;2,1,2,3,0;2,1,2,1,0;2,1,2,8,0;1,1,1,0,0"

' �J���[�̃e�N�X�`���`��
 Const pattern_enquete_color As String = "5;1,37,2,0,0;1,34,2,0,0;1,35,2,0,0;1,38,2,0,0;1,7,2,0,0"

' save setting
Private Const app = "Excel", sec = "SeriesFillColor", keyClose = "AutoClose", keyUserText = "UserText"

Function �����p�^�[��20() As String
    �����p�^�[��20 = pattern_bw20
End Function

Function �����p�^�[��5() As String
    �����p�^�[��5 = pattern_enquete
End Function


Private Sub applyBarFill4SingleSeries(info_string As String)
On Error GoTo eee

Dim info
Dim s As Series
Dim ff As ChartFillFormat
Dim n As Integer
Dim i As Integer
Dim p As Point

Set s = ActiveChart.SeriesCollection(1)

n = s.Points().Count

info = Split(info_string, ";")
If val(info(0)) < n Then n = CInt(info(0))


 
For i = 1 To n
  Set p = s.Points(i)
  Set ff = p.Fill
  Dim t, fc, bc, arg1, arg2, arr
  arr = Split(info(i), ",")
  t = CLng(arr(0)) 'ff.Type
  fc = CLng(arr(1)) 'ff.ForeColor.RGB
  bc = CLng(arr(2)) 'ff.BackColor.RGB
  
  ' set color
  ff.ForeColor.SchemeColor = fc
  ff.BackColor.SchemeColor = bc
  
  Select Case t
  Case msoFillGradient ' 3
    ' �O���f�[�V����
    Dim a, g_type, g_style, g_variant, g_degree, g_preset
    a = Split(arr(4), ":")
    'a(0) = ff.GradientColorType
    'a(1) = ff.GradientDegree
    'a(2) = ff.GradientStyle
    'a(3) = ff.GradientVariant
    g_type = CLng(a(0))
    g_degree = CLng(a(1))
    g_style = CLng(a(2))
    g_variant = CLng(a(3))
    g_preset = CLng(a(4))
    'arg1 = 4
    'arg2 = Join(a, ":")
    Select Case CLng(a(0))
    Case msoGradientColorMixed
      '???
    Case msoGradientOneColor
      ff.OneColorGradient g_style, g_variant, g_degree
    Case msoGradientPresetColors
      ff.PresetGradient g_style, g_variant, g_preset
    Case msoGradientTwoColors
      ff.TwoColorGradient g_stype, g_variant
    End Select
  Case msoFillMixed
     ' ???
    'arg1 = 0
    'arg2 = 0
  Case msoFillPatterned
    '�p�^�[��
    'arg1 = ff.Pattern
    'arg2 = 0
    ff.Patterned CLng(arr(3))
  Case msoFillPicture
'     arg1 = 0
'     arg2 = 0
    If arr(3) <> "0" Then
      If arr(4) = "0" Then
        ff.UserPicture arr(3)
      Else
        Dim cfg
        cfg = Split(arr(4), ":")
        ff.UserPicture arr(3), CLng(cfg(0)), CLng(cfg(1)), CLng(cfg(2))
      End If
    End If
  Case msoFillSolid
    ff.Solid
    ff.ForeColor.SchemeColor = fc
  Case msoFillTextured
     If CLng(arr(3)) = msoTexturePreset Then
       ff.PresetTextured CLng(arr(4))
     Else
       ff.UserTextured arr(4)
     End If
  End Select
Next

Exit Sub

eee:

   MsgBox "�G���[���������Ă��܂��D"

End Sub

Sub applyBarFill(info_string As String)
Dim s As Series

If ActiveChart.ChartGroups(1).VaryByCategories Then
    applyBarFill4SingleSeries info_string
Else
    applyBarFill4MultiSeries info_string
End If

End Sub


Private Sub applyBarFill4MultiSeries(info_string As String)
On Error GoTo eee

Dim info
Dim s As Series
Dim ff As ChartFillFormat
Dim n As Integer
Dim i As Integer

n = ActiveChart.SeriesCollection.Count
info = Split(info_string, ";")

If val(info(0)) < n Then n = CInt(info(0))

 
For i = 1 To n
  Set s = ActiveChart.SeriesCollection(i)
  Set ff = s.Fill
  Dim t, fc, bc, arg1, arg2, arr
  arr = Split(info(i), ",")
  t = CLng(arr(0)) 'ff.Type
  fc = CLng(arr(1)) 'ff.ForeColor.RGB
  bc = CLng(arr(2)) 'ff.BackColor.RGB
  
  
  ' set color
  ff.ForeColor.SchemeColor = fc
  ff.BackColor.SchemeColor = bc
  
  Select Case t  ' msofilltype
  Case msoFillGradient
    ' �O���f�[�V����
    Dim a, g_type, g_style, g_variant, g_degree, g_preset
    a = Split(arr(4), ":")
    'a(0) = ff.GradientColorType
    'a(1) = ff.GradientDegree
    'a(2) = ff.GradientStyle
    'a(3) = ff.GradientVariant
    g_type = CLng(a(0))
    g_degree = CLng(a(1))
    g_style = CLng(a(2))
    g_variant = CLng(a(3))
    g_preset = CLng(a(4))
    'arg1 = 4
    'arg2 = Join(a, ":")
    Select Case CLng(a(0))
    Case msoGradientColorMixed
      '???
    Case msoGradientOneColor
      ff.OneColorGradient g_style, g_variant, g_degree
    Case msoGradientPresetColors
      ff.PresetGradient g_style, g_variant, g_preset
    Case msoGradientTwoColors
      ff.TwoColorGradient g_stype, g_variant
    End Select
  Case msoFillMixed
     ' ???
    'arg1 = 0
    'arg2 = 0
  Case msoFillPatterned
    '�p�^�[��
    'arg1 = ff.Pattern
    'arg2 = 0
    ff.Patterned CLng(arr(3))
  Case msoFillPicture
'     arg1 = 0
'     arg2 = 0
    If arr(3) <> "0" Then
      If arr(4) = "0" Then
        ff.UserPicture arr(3)
      Else
        Dim cfg
        cfg = Split(arr(4), ":")
        ff.UserPicture arr(3), CLng(cfg(0)), CLng(cfg(1)), CLng(cfg(2))
      End If
    End If
  Case msoFillSolid
    ff.Solid
  Case msoFillTextured
     If CLng(arr(3)) = msoTexturePreset Then
       ff.PresetTextured CLng(arr(4))
     Else
       ff.UserTextured arr(4)
     End If
  End Select
Next

Exit Sub

eee:

   MsgBox "�G���[���������Ă��܂��D"

End Sub

Private Sub check_and_applyBarFill(ptn As String)
  Me.Label1.Caption = ""
  On Error GoTo eee
  Select Case ActiveChart.ChartType
  Case xlBar, xlBarClustered, xlBarStacked, xlBarStacked100, _
       xlColumn, xlColumnClustered, xlColumnStacked, xlColumnStacked100
  Case Else
    Me.Label1.Caption = "�Ή����Ă��Ȃ��O���t�^�C�v�ł�"
  End Select
  applyBarFill ptn

  If Me.cbClose.Value Then Unload Me
Exit Sub
eee:
   Me.Label1.Caption = "�G���[�ł��D�O���t���I������Ă��Ȃ��\��������܂�"

End Sub


Private Sub Image1_Click()
  check_and_applyBarFill pattern_bw
End Sub

Private Sub Image2_Click()
  check_and_applyBarFill pattern_gray
End Sub

Private Sub Image3_Click()
  check_and_applyBarFill pattern_enquete
End Sub

Private Sub Image4_Click()
  check_and_applyBarFill pattern_enquete_color
End Sub

Private Sub Image5_Click()
  check_and_applyBarFill pattern_bw20
End Sub

Private Sub lblUseText_Click()
  check_and_applyBarFill Me.tbColorText.Text
End Sub

Private Sub UserForm_DblClick(ByVal Cancel As MSForms.ReturnBoolean)
  Unload Me

End Sub


Private Sub UserForm_Initialize()
  Me.tbColorText.Text = GetSetting(app, sec, keyUserText, "")
  Me.cbClose.Value = CBool(GetSetting(app, sec, keyClose, "False"))
End Sub

Private Sub UserForm_Terminate()
  SaveSetting app, sec, keyClose, CStr(Me.cbClose.Value)
  SaveSetting app, sec, keyUserText, Me.tbColorText.Text
End Sub
