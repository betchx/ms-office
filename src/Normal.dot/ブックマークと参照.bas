Attribute VB_Name = "�u�b�N�}�[�N�ƎQ��"
Option Explicit


Sub ���p�X�y�[�X�L��(Optional times As Integer = 1)
   Dim i As Integer
   For i = 1 To times
     Selection.TypeText text:=" "
   Next
End Sub


' ���i���L�� Macro
Sub ���i���L��(Optional times As Integer = 1)
   Dim i As Integer
   For i = 1 To times
    Selection.TypeParagraph
   Next
End Sub


' �Q�Ɛ�I�� ���[�U�[�t�H�[����\��
Sub �Q��()
  �Q�Ɛ�I��.Show vbModeless
End Sub


' �I�𕶎��񂩂�u�b�N�}�[�N���쐬
Sub �I�𕶎��񂩂�u�b�N�}�[�N()
  Dim tag As String
  tag = �u�b�N�}�[�N�\�����ւ̕ϊ�(Trim(Selection.text))
  ActiveDocument.Bookmarks.Add tag
End Sub

Sub �u�b�N�}�[�N()
    Dim e As New �u�b�N�}�[�N���ҏW
    Dim r As Range
    Set r = Selection.Range
    If Selection.Start = Selection.End Then
      Select Case MsgBox("�͈͂��I������Ă��܂���D�I��͈͂�i���֊g�����܂����H", vbYesNoCancel, "�m�F")
      Case vbCancel
        Exit Sub
      Case vbYes
        Set r = Selection.Paragraphs(1).Range
        r.MoveEnd wdCharacter, -1
      End Select
    End If
    ' �I��͈͂̑O��ɂ���󔒓����폜
    r.MoveStartWhile CSet:=" �@" & vbTab, Count:=r.End - r.Start
    r.MoveEndWhile CSet:=" �@" & vbTab & wdCRLF, Count:=r.Start - r.End
    e.Show vbModal
    If Len(e.����) > 0 Then
      ActiveDocument.Bookmarks.Add e.����, r
    End If
    Unload e
    Set e = Nothing
  
End Sub


'  TrixEx�V���[�Y �����ɉ��s�L��������ꍇ�ɂ������폜���Ă���g��������

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



Function �u�b�N�}�[�N�\�����ւ̕ϊ�(ByVal target As String) As String
  Dim tag As String, ngs, rep, i As Integer
 
  tag = TrimEx(target)
  
  Select Case Left(tag, 1)
  Case "1", "2", "3", "4", "5", "6", "7", "8", "9", "0", "�P", "�Q", "�R", "�S", "�T", "�U", "�V", "�W", "�X", "�O"
    tag = "�Q" + tag
  End Select
  ngs = Array(" ", "�@", vbTab, "(", ")", "-", "?", ".", ",", "/", "!", "*", "%", "#", "'", "=", "^", "~", "\", "|", Chr(10), Chr(13))
  rep = Array("", "", "", "�i", "�j", "�|", "", "�D", "�C", "�^", "", "", "��", "", "�f", "��", "", "", "��", "�b", "", "")
  
  For i = 0 To UBound(ngs)
    tag = Replace(tag, ngs(i), rep(i))
  Next
  
  If LenB(tag) > 40 Then
    tag = Replace(tag, "�@", "")
    If LenB(tag) > 400 Then tag = Left(tag, 40)
    If LenB(tag) > 40 Then
      tag = Replace(tag, "�i", "_")
      tag = Replace(tag, "�j", "_")
    End If
    Dim tr
    tr = "���b�|���C�D��"
    For i = 1 To Len(tr)
      If LenB(tag) <= 40 Then Exit For
      tag = Replace(tag, Mid(tr, i, 1), "")
    Next
    
    If LenB(tag) > 80 Then tag = Left(tag, 40)
    Do While LenB(tag) > 40
      tag = Left(tag, Len(tag) - 1)
    Loop
  End If
  
  �u�b�N�}�[�N�\�����ւ̕ϊ� = tag

End Function


' �u�b�N�}�[�N�̊J�n�ƏI�����C��
Sub �u�b�N�}�[�N�͈̔͂�ҏW()
   �u�b�N�}�[�N�͈͂̕ҏW.Show vbModeless
End Sub

