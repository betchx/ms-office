VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} �u�b�N�}�[�N���ҏW 
   Caption         =   "�u�b�N�}�[�N���̕ҏW"
   ClientHeight    =   1230
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   6210
   OleObjectBlob   =   "�u�b�N�}�[�N���ҏW.frx":0000
   StartUpPosition =   1  '�I�[�i�[ �t�H�[���̒���
End
Attribute VB_Name = "�u�b�N�}�[�N���ҏW"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Public ���� As String

Private Sub cbCancel_Click()
   ���� = ""
   Me.Hide
End Sub

Private Sub cbOK_Click()
    ���� = TextBox1.text
    Me.Hide
End Sub

Private Sub cbReset_Click()
  TextBox1.text = �u�b�N�}�[�N�\�����ւ̕ϊ�(���)
End Sub

Private Sub TextBox1_Change()
  Dim bk As String
  bk = �u�b�N�}�[�N�\�����ւ̕ϊ�(TextBox1.text)
  If TextBox1.text <> bk Then
    ' �g�p�ł��Ȃ��������^����ꂽ�ꍇ�͎g������̂ɕϊ����čĐݒ�
    TextBox1.text = bk
  End If
End Sub

Private Sub UserForm_Initialize()

' ngs = Array(" ", "�@", "(", ")", "-", "?", ".", ",", "/", "!", "*", "%", "#", "'", "=", "^", "~", "\", "|", Chr(10), Chr(13))
 
  ���� = ""
 
  Label1.Caption = _
"�u�b�N�}�[�N��ǉ��E�ҏW���Ă��������D" & vbCrLf & _
"�ő�20�����ł��D�X�y�[�X���͎g���܂���D" & vbCrLf & _
"���̑��̎g�p�ł��Ȃ������͑S�p�ɕϊ�����܂��D"

End Sub

Sub ���ݒ�(Optional ��� As String = "")
'"�u�b�N�}�[�N�͍ő�20�����܂łŁC��������n�܂邱�Ƃ͂ł��܂���D" & vbCrLf & _
'"�܂��C���s��X�y�[�X�i���p�S�p�Ƃ��j�Ǝ��̔��p�����͎g�p�ł��܂���F ()-?.,/!*%#'=^~\|"
  If ��� = "" Then
    If Selection.Start = Selection.End Then
      ��� = �u�b�N�}�[�N�\�����ւ̕ϊ�(Selection.Paragraphs(1).Range.text)
    Else
      ��� = �u�b�N�}�[�N�\�����ւ̕ϊ�(Selection.text)
    End If
  End If
  TextBox1.text = �u�b�N�}�[�N�\�����ւ̕ϊ�(���)
  TextBox1.SetFocus
End Sub
