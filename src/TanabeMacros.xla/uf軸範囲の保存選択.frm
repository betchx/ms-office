VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} uf²ÍÍÌÛ¶Ið 
   Caption         =   "²ÍÍÌÛ¶Ið"
   ClientHeight    =   825
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   4350
   OleObjectBlob   =   "uf²ÍÍÌÛ¶Ið.frx":0000
   StartUpPosition =   1  'I[i[ tH[Ì
End
Attribute VB_Name = "uf²ÍÍÌÛ¶Ið"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Public Ið As Integer

Private Sub cbBoth_Click()
  Ið = 3
  Me.Hide
End Sub

Private Sub cbX_Click()
  Ið = 1
  Me.Hide
End Sub

Private Sub cbY_Click()
  Ið = 2
  Me.Hide
End Sub

Private Sub UserForm_Initialize()
  Ið = 0
End Sub
