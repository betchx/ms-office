VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} ブックマーク範囲の編集 
   Caption         =   "ブックマークのレンジを修正"
   ClientHeight    =   2400
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   7845
   OleObjectBlob   =   "ブックマーク範囲の編集.frx":0000
   StartUpPosition =   1  'オーナー フォームの中央
End
Attribute VB_Name = "ブックマーク範囲の編集"
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
  w = Replace(w, "　", "□")
  w = Replace(w, vbCr, "\r")
  t = Replace(w, vbLf, "\n")
End Function

Private Function キャレット部文字列作成(pos As Long, Optional キャレット As String = "|") As String
 Dim s As Long, e As Long
 s = pos - 10
 If s < 1 Then s = 1
 e = pos + 9
 If e > ActiveDocument.Content.End Then e = ActiveDocument.Content.End
 
   キャレット部文字列作成 = _
   t(s, pos) & _
   キャレット & _
   t(pos, e)
  
End Function

Private Sub スタート更新()
 Me.tbStart.text = キャレット部文字列作成(cbm().Start, " [ ")
End Sub

Private Sub エンド更新()
 Me.tbEnd.text = キャレット部文字列作成(cbm().End, " ] ")
End Sub

Private Sub 範囲更新()
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
   スタート更新
    With Me.sbStart
      .Value = 1
      .Max = b.End
      .Value = b.Start
    End With
   sbStart.Enabled = True
   エンド更新
    With Me.sbEnd
      .Value = .Max
      .Min = b.Start
      .Value = b.End
    End With
   sbEnd.Enabled = True
 End If
End Sub


Private Sub listBookmark_Change()
  範囲更新
End Sub


Private Sub sbEnd_SpinDown()
    If cbm().End > cbm().Start Then _
        cbm().End = cbm().End - 1
    エンド更新
End Sub

Private Sub sbEnd_SpinUp()
    If cbm().End < ActiveDocument.Content.End Then _
        cbm().End = cbm().End + 1
    エンド更新
End Sub

Private Sub sbStart_SpinDown()
  If cbm().Start > 1 Then _
      cbm().Start = cbm.Start - 1
  スタート更新
End Sub

Private Sub sbStart_SpinUp()
  If cbm().Start < cbm().End Then _
      cbm().Start = cbm.Start + 1
  スタート更新
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
