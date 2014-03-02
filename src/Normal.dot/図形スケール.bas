Attribute VB_Name = "�}�`�X�P�[��"
Option Explicit
Const app As String = "WordMacro"
Const sec As String = "Triming"

Sub �ꊇ�X�P�[��()
  Dim i
  Dim s As Single, rate As Single, h As Single, w As Single
  Dim res As String
  Dim pos As Long
  Const key As String = "scale"
'  Const message As String = "Size in % or mm." & vbCrLf _
'    & " Add '%' for scale: " & vbCrLf _
'    & "      possitive value for scaling from current size," & vbCrLf _
'    & "      negative value for scaling from original size if possible." & vbCrLf _
'    & " otherwise scaling for constant width or height:" & vbCrLf _
'    & "      possitive value for width," & vbCrLf _
'    & "      negative value for new height "
  Const message As String = "�T�C�Y�� % �� ���l�i�P�ʂ�mm) �Ŏw�肵�Ă��������D" & vbCrLf _
    & " �Ō�� '%' ������ꍇ " & vbCrLf _
    & "      ���̒l�͌��݂̃T�C�Y�ɑ΂��銄���ŃX�P�[�����܂�" & vbCrLf _
    & "      ���̒l�̏ꍇ�͉\�Ȃ�΃I���W�i���T�C�Y�ɑ΂��銄���ŃX�P�[�����܂��D" & vbCrLf _
    & " ���l�Ŏw�肵���ꍇ�͕��������͍��������̒l�ƂȂ�l�ɃX�P�[�����܂�:" & vbCrLf _
    & "      ���̒l�̏ꍇ�͕�," & vbCrLf _
    & "      ���̒l�̏ꍇ�͍��� "
  
  res = StrConv(InputBox(message, "�T�C�Y�ݒ�", GetSetting(app, "size", key, "100%")), vbNarrow)
  
  If res = "" Then Exit Sub
  
  SaveSetting app, "size", key, res
  
  pos = InStr(res, "%")
  If pos > 0 Then
    ' %����Ȃ̂Ŋ�����Scale
    s = CSng(Left(res, pos - 1))
    If s < 0 Then
       ' scale by original
        s = s * -1
        �I��͈͓��̐}�������ŃX�P�[�� s, ���̃T�C�Y����:=True
    Else
        �I��͈͓��̐}�������ŃX�P�[�� s
    End If
  Else
    s = CSng(res)
    If s > 0# Then
        �I��͈͓��̐}�𕝂ŃX�P�[�� s
    Else
        �I��͈͓��̐}�������ŃX�P�[�� -s
    End If
  End If
End Sub

Private Sub �I��͈͓��̐}�������ŃX�P�[��(���� As Single, Optional ���̃T�C�Y���� As Boolean = False)
    Dim i
    Dim �\�ȑΏ� As Boolean
    Dim p As Paragraph
    Dim r As Range
    
    Select Case Selection.Type
    Case wdSelectionInlineShape
        For Each i In Selection.InlineShapes
            i.ScaleHeight = ����
            i.ScaleWidth = ����
        Next
    Case wdSelectionShape
        For Each i In Selection.ShapeRange
            �\�ȑΏ� = i.Type = msoPicture Or i.Type = msoOLEControlObject
            i.ScaleHeight ���� * 0.01, ���̃T�C�Y���� And �\�ȑΏ�
            i.ScaleWidth ���� * 0.01, ���̃T�C�Y���� And �\�ȑΏ�
        Next
    Case wdSelectionNormal
        For Each p In Selection.Paragraphs
            Set r = p.Range
            If r.InlineShapes.Count > 0 Then
                For Each i In p.Range.InlineShapes
                    i.ScaleHeight = ����
                    i.ScaleWidth = ����
                Next i
            End If
        Next p
        If Selection.ShapeRange.Count > 0 Then
            For Each i In Selection.ShapeRange
                �\�ȑΏ� = i.Type = msoPicture Or i.Type = msoOLEControlObject
                i.ScaleHeight ���� * 0.01, ���̃T�C�Y���� And �\�ȑΏ�
                i.ScaleWidth ���� * 0.01, ���̃T�C�Y���� And �\�ȑΏ�
            Next i
        End If
    End Select

End Sub

Private Sub �I��͈͓��̐}�𕝂ŃX�P�[��(mm As Single)
    Dim s
    Dim p As Paragraph
    Dim r As Range
    Dim i
    Dim Top, Left
    Dim rate As Single
    s = mm2pnt(mm)
    Select Case Selection.Type
    Case wdSelectionShape
        For Each i In Selection.ShapeRange
            Top = i.Top
            Left = i.Left
            i.ScaleHeight s / i.Width, False
            i.Width = s
            i.Top = Top
            i.Left = Left
        Next
    Case wdSelectionInlineShape
        For Each i In Selection.InlineShapes
            i.ScaleWidth = 1
            rate = s / i.Width * 100#
            i.ScaleHeight = rate
            i.Width = s
        Next
    Case wdSelectionNormal
        For Each p In Selection.Paragraphs
            Set r = p.Range
            If �V�F�[�v����(r) Then
                For Each i In r.ShapeRange
                    Top = i.Top
                    Left = i.Left
                    i.Height = i.Height * s / i.Width
                    i.Width = s
                    i.Top = Top
                    i.Left = Left
                Next i
            End If
            On Error GoTo 0
            If r.InlineShapes.Count > 0 Then
                Dim ils As InlineShape
                For Each i In p.Range.InlineShapes
                    Set ils = i
                    ils.Reset
                    i.ScaleHeight = s / i.Width * 100
                    i.Width = s
                Next i
            End If
        Next p
    End Select
End Sub

Private Function �V�F�[�v����(r As Range) As Boolean
  On Error GoTo eee:
  If r.ShapeRange.Count > 0 Then
    �V�F�[�v���� = True
  End If
  Exit Function
eee:
  �V�F�[�v���� = False
End Function


Private Sub �I��͈͓��̐}�������ŃX�P�[��(mm As Single)
    Dim s
    Dim p As Paragraph
    Dim r As Range
    Dim i
    
    s = mm2pnt(mm)
    Select Case Selection.Type
    Case wdSelectionShape
        For Each i In Selection.ShapeRange
            i.ScaleWidth s / i.Height, False
            i.Height = s
        Next
    Case wdSelectionInlineShape
        For Each i In Selection.InlineShapes
            i.ScaleWidth = s / i.Height * 100
            i.Height = s
        Next
    Case wdSelectionNormal
        For Each p In Selection.Paragraphs
            Set r = p.Range
            If r.ShapeRange.Count > 0 Then
                For Each i In r.ShapeRange
                    i.ScaleWidth s / i.Height, False
                    i.Height = s
                Next i
            End If
            If r.InlineShapes.Count > 0 Then
                For Each i In p.Range.InlineShapes
                    i.ScaleWidth = s / i.Height * 100
                    i.Height = s
                Next i
            End If
        Next p
    End Select
End Sub



Function mm2pnt(mm)
'  mm2pnt = mm * 72 / 25.4    '(mm * 160#) / 64#
  mm2pnt = Application.MillimetersToPoints(mm)
End Function

Function pt2mm(pt)
'  pt2mm = pt * 25.4 / 72# '(pt * 64#) / 160#
  pt2mm = Application.PointsToMillimeters(pt)
End Function




