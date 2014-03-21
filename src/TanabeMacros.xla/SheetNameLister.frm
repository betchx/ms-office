VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} SheetNameLister 
   Caption         =   "シート名選択"
   ClientHeight    =   3225
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   4710
   OleObjectBlob   =   "SheetNameLister.frx":0000
   StartUpPosition =   1  'オーナー フォームの中央
End
Attribute VB_Name = "SheetNameLister"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Public selected_name As String


Private Sub ListBox1_選択完了()
  Dim idx As Integer
  idx = Me.ListBox1.ListIndex
  If idx = -1 Then Exit Sub
  selected_name = Me.ListBox1.List(idx)
  Me.Hide
End Sub

Private Sub ListBox1_DblClick(ByVal Cancel As MSForms.ReturnBoolean)
  Call ListBox1_選択完了
End Sub


Private Sub ListBox1_KeyUp(ByVal KeyCode As MSForms.ReturnInteger, ByVal Shift As Integer)
  Select Case KeyCode.Value
  Case KeyCodeConstants.vbKeyReturn
    Call ListBox1_選択完了
  Case KeyCodeConstants.vbKeyEscape
    Me.Hide
  End Select
End Sub

Private Sub ListBox1_MouseUp(ByVal Button As Integer, ByVal Shift As Integer, ByVal X As Single, ByVal Y As Single)
  If Button = 1 Then
    Call ListBox1_選択完了
  End If
End Sub

Private Sub UserForm_Initialize()
  selected_name = ""
  Call UserForm_Resize
  Dim sheet_names() As String
  Dim i As Integer
  With ActiveWorkbook.Worksheets
    ReDim sheet_names(.Count)
    For i = 1 To .Count
      sheet_names(i - 1) = .item(i).name
    Next i
  End With
  Me.ListBox1.List = sheet_names
  Me.ListBox1.SetFocus
End Sub

Private Sub UserForm_Resize()
  With ListBox1
    .Top = 0
    .Left = 0
    .Width = Me.InsideWidth
    .Height = Me.InsideHeight
  End With
End Sub

