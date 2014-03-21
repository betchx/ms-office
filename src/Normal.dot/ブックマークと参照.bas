Attribute VB_Name = "ブックマークと参照"
Option Explicit


Sub 半角スペース記入(Optional times As Integer = 1)
   Dim i As Integer
   For i = 1 To times
     Selection.TypeText text:=" "
   Next
End Sub


' 改段落記入 Macro
Sub 改段落記入(Optional times As Integer = 1)
   Dim i As Integer
   For i = 1 To times
    Selection.TypeParagraph
   Next
End Sub


' 参照先選択 ユーザーフォームを表示
Sub 参照()
  参照先選択.Show vbModeless
End Sub


' 選択文字列からブックマークを作成
Sub 選択文字列からブックマーク()
  Dim tag As String
  tag = ブックマーク可能文字への変換(Trim(Selection.text))
  ActiveDocument.Bookmarks.Add tag
End Sub

Sub ブックマーク()
    Dim e As New ブックマーク名編集
    Dim r As Range
    Set r = Selection.Range
    If Selection.Start = Selection.End Then
      Select Case MsgBox("範囲が選択されていません．選択範囲を段落へ拡張しますか？", vbYesNoCancel, "確認")
      Case vbCancel
        Exit Sub
      Case vbYes
        Set r = Selection.Paragraphs(1).Range
        r.MoveEnd wdCharacter, -1
      End Select
    End If
    ' 選択範囲の前後にある空白等を削除
    r.MoveStartWhile CSet:=" 　" & vbTab, Count:=r.End - r.Start
    r.MoveEndWhile CSet:=" 　" & vbTab & wdCRLF, Count:=r.Start - r.End
    e.Show vbModal
    If Len(e.結果) > 0 Then
      ActiveDocument.Bookmarks.Add e.結果, r
    End If
    Unload e
    Set e = Nothing
  
End Sub


'  TrixExシリーズ 末尾に改行記号がある場合にそれらを削除してからトリムする

Function TrimEx(ByVal target As String) As String
  Dim n As Integer
  Dim i As Integer
  i = 1
  n = Len(target)
  If n < 2 Then
    TrimEx = target
  Else
    If Mid(target, n - 1) = vbCrLf Then n = n - 2
    If Right(target, 1) = vbCr Then n = n - 1
    If Right(target, 1) = vbLf Then n = n - 1
    TrimEx = Trim(Left(target, n))
  End If

End Function

Function RTrimEx(ByVal target As String) As String
  Dim n As Integer
  
  n = Len(target)
  If n < 2 Then
    RTrimEx = target
  Else
    If Mid(target, n - 1) = vbCrLf Then n = n - 2
    If Right(target, 1) = vbCr Then n = n - 1
    If Right(target, 1) = vbLf Then n = n - 1
    RTrimEx = RTrim(Left(target, n))
  End If

End Function

Function LTrimEx(ByVal target As String) As String
  Dim n As Integer
  
  n = Len(target)
  If n < 2 Then
    LTrimEx = target
  Else
    If Mid(target, n - 1) = vbCrLf Then n = n - 2
    If Right(target, 1) = vbCr Then n = n - 1
    If Right(target, 1) = vbLf Then n = n - 1
    LTrimEx = LTrim(Left(target, n))
  End If

End Function



Function ブックマーク可能文字への変換(ByVal target As String) As String
  Dim tag As String, ngs, rep, i As Integer
 
  tag = TrimEx(target)
  
  Select Case Left(tag, 1)
  Case "1", "2", "3", "4", "5", "6", "7", "8", "9", "0", "１", "２", "３", "４", "５", "６", "７", "８", "９", "０"
    tag = "＿" + tag
  End Select
  ngs = Array(" ", "　", vbTab, "(", ")", "-", "?", ".", ",", "/", "!", "*", "%", "#", "'", "=", "^", "~", "\", "|", Chr(10), Chr(13))
  rep = Array("", "", "", "（", "）", "－", "", "．", "，", "／", "", "", "％", "", "’", "＝", "", "", "￥", "｜", "", "")
  
  For i = 0 To UBound(ngs)
    tag = Replace(tag, ngs(i), rep(i))
  Next
  
  If LenB(tag) > 40 Then
    tag = Replace(tag, "　", "")
    If LenB(tag) > 400 Then tag = Left(tag, 40)
    If LenB(tag) > 40 Then
      tag = Replace(tag, "（", "_")
      tag = Replace(tag, "）", "_")
    End If
    Dim tr
    tr = "￥｜－＝，．％"
    For i = 1 To Len(tr)
      If LenB(tag) <= 40 Then Exit For
      tag = Replace(tag, Mid(tr, i, 1), "")
    Next
    
    If LenB(tag) > 80 Then tag = Left(tag, 40)
    Do While LenB(tag) > 40
      tag = Left(tag, Len(tag) - 1)
    Loop
  End If
  
  ブックマーク可能文字への変換 = tag

End Function


' ブックマークの開始と終了を修正
Sub ブックマークの範囲を編集()
   ブックマーク範囲の編集.Show vbModeless
End Sub

