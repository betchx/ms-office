VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} ���ޑI�� 
   Caption         =   "���ޑI��"
   ClientHeight    =   3660
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   5625
   OleObjectBlob   =   "���ޑI��.frx":0000
   StartUpPosition =   1  '�I�[�i�[ �t�H�[���̒���
End
Attribute VB_Name = "���ޑI��"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private detail(7, 15)

Private Sub UserForm_Initialize()

 Me.ListBoxType.AddItem "�|��"
 detail(1, 1) = "t6"
 Me.ListBoxType.AddItem "H�|��"
 Me.ListBoxType.AddItem "�R�`�|"
 Me.ListBoxType.AddItem "�p�^�|��"
 Me.ListBoxType.AddItem "�Z�p�{���g(HTB)"
 Me.ListBoxType.AddItem "�g���V�A�{���g(TCB)"
 Me.ListBoxType.AddItem "�A���J�[�{���g"
End Sub


