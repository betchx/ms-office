VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} uf���͈͂̕ۑ��I�� 
   Caption         =   "���͈͂̕ۑ��I��"
   ClientHeight    =   825
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   4350
   OleObjectBlob   =   "uf���͈͂̕ۑ��I��.frx":0000
   StartUpPosition =   1  '�I�[�i�[ �t�H�[���̒���
End
Attribute VB_Name = "uf���͈͂̕ۑ��I��"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Public �I�� As Integer

Private Sub cbBoth_Click()
  �I�� = 3
  Me.Hide
End Sub

Private Sub cbX_Click()
  �I�� = 1
  Me.Hide
End Sub

Private Sub cbY_Click()
  �I�� = 2
  Me.Hide
End Sub

Private Sub UserForm_Initialize()
  �I�� = 0
End Sub
