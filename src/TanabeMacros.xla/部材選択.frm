VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} 部材選択 
   Caption         =   "部材選択"
   ClientHeight    =   3660
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   5625
   OleObjectBlob   =   "部材選択.frx":0000
   StartUpPosition =   1  'オーナー フォームの中央
End
Attribute VB_Name = "部材選択"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private detail(7, 15)

Private Sub UserForm_Initialize()

 Me.ListBoxType.AddItem "鋼鈑"
 detail(1, 1) = "t6"
 Me.ListBoxType.AddItem "H鋼材"
 Me.ListBoxType.AddItem "山形鋼"
 Me.ListBoxType.AddItem "角型鋼鈑"
 Me.ListBoxType.AddItem "六角ボルト(HTB)"
 Me.ListBoxType.AddItem "トルシアボルト(TCB)"
 Me.ListBoxType.AddItem "アンカーボルト"
End Sub


