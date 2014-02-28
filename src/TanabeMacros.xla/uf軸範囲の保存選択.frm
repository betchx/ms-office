VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} uf軸範囲の保存選択 
   Caption         =   "軸範囲の保存選択"
   ClientHeight    =   825
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   4350
   OleObjectBlob   =   "uf軸範囲の保存選択.frx":0000
   StartUpPosition =   1  'オーナー フォームの中央
End
Attribute VB_Name = "uf軸範囲の保存選択"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Public 選択 As Integer

Private Sub cbBoth_Click()
  選択 = 3
  Me.Hide
End Sub

Private Sub cbX_Click()
  選択 = 1
  Me.Hide
End Sub

Private Sub cbY_Click()
  選択 = 2
  Me.Hide
End Sub

Private Sub UserForm_Initialize()
  選択 = 0
End Sub
