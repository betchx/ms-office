VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} �u�b�N�}�[�N�͈͂̕ҏW 
   Caption         =   "�u�b�N�}�[�N�̃����W���C��"
   ClientHeight    =   2400
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   7845
   OleObjectBlob   =   "�u�b�N�}�[�N�͈͂̕ҏW.frx":0000
   StartUpPosition =   1  '�I�[�i�[ �t�H�[���̒���
End
Attribute VB_Name = "�u�b�N�}�[�N�͈͂̕ҏW"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Dim arr()



Private Function t(s As Long, e As Long) As String
  Dim w As String
  w = ActiveDocument.Range(s, e).text
  w = Replace(w, " ", "\s")
  w = Replace(w, "�@", "��")
  w = Replace(w, vbCr, "\r")
  t = Replace(w, vbLf, "\n")
End Function

Private Function �L�����b�g��������쐬(pos As Long, Optional �L�����b�g As String = "|") As String
 Dim s As Long, e As Long
 s = pos - 10
 If s < 1 Then s = 1
 e = pos + 9
 If e > ActiveDocument.Content.End Then e = ActiveDocument.Content.End
 
   �L�����b�g��������쐬 = _
   t(s, pos) & _
   �L�����b�g & _
   t(pos, e)
  
End Function

Private Sub �X�^�[�g�X�V()
 Me.tbStart.text = �L�����b�g��������쐬(cbm().Start, " [ ")
End Sub

Private Sub �G���h�X�V()
 Me.tbEnd.text = �L�����b�g��������쐬(cbm().End, " ] ")
End Sub

Private Sub �͈͍X�V()
 Dim i As Long
 Dim b As Bookmark
 If Len(Me.listBookmark.text) = 0 Then
   Me.tbStart.text = ""
   Me.tbEnd.text = ""
   sbStart.Enabled = False
   sbEnd.Enabled = False
 Else
   i = Me.listBookmark.ListIndex
   Set b = ActiveDocument.Bookmarks(i + 1)
   �X�^�[�g�X�V
    With Me.sbStart
      .Value = 1
      .Max = b.End
      .Value = b.Start
    End With
   sbStart.Enabled = True
   �G���h�X�V
    With Me.sbEnd
      .Value = .Max
      .Min = b.Start
      .Value = b.End
    End With
   sbEnd.Enabled = True
 End If
End Sub


Private Sub listBookmark_Change()
  �͈͍X�V
End Sub


Private Sub sbEnd_SpinDown()
    If cbm().End > cbm().Start Then _
        cbm().End = cbm().End - 1
    �G���h�X�V
End Sub

Private Sub sbEnd_SpinUp()
    If cbm().End < ActiveDocument.Content.End Then _
        cbm().End = cbm().End + 1
    �G���h�X�V
End Sub

Private Sub sbStart_SpinDown()
  If cbm().Start > 1 Then _
      cbm().Start = cbm.Start - 1
  �X�^�[�g�X�V
End Sub

Private Sub sbStart_SpinUp()
  If cbm().Start < cbm().End Then _
      cbm().Start = cbm.Start + 1
  �X�^�[�g�X�V
End Sub

' Current Book Mark
Private Function cbm() As Bookmark
  Set cbm = ActiveDocument.Bookmarks(listBookmark.ListIndex + 1)
End Function



Private Sub tbEnd_Enter()
'  ActiveDocument.Content.Characters(cbm().End).Select
  cbm.Select
  Selection.Collapse wdCollapseEnd
End Sub

Private Sub tbStart_Enter()
  cbm().Select
  Selection.Collapse
End Sub

Private Sub UserForm_Activate()
 Dim b As Bookmark
 Dim i As Long, n As Long
 
 n = ActiveDocument.Bookmarks.Count
    
 ReDim arr(0 To n - 1)

 For i = 1 To n
   Set b = ActiveDocument.Bookmarks(i)
   arr(i - 1) = b.name
 Next

 Me.listBookmark.List = arr

 Me.sbStart.Min = 1
 Me.sbEnd.Max = ActiveDocument.Content.End

 If Selection.Bookmarks.Count > 0 Then
   Set b = Selection.Bookmarks(1)
   For i = 0 To n - 1
       If arr(i) = b.name Then
         Me.listBookmark.ListIndex = i
         Exit For
      End If
   Next
 End If

End Sub


Private Sub UserForm_Click()
   UserForm_Activate
End Sub
