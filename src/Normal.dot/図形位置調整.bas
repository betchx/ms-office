Attribute VB_Name = "�}�`�ʒu����"
Option Explicit


Sub �}�`�E�[()
    Selection.ShapeRange.RelativeHorizontalPosition = _
        wdRelativeHorizontalPositionMargin
    Selection.ShapeRange.Left = wdShapeRight
End Sub
Sub �}�`���[()
    Selection.ShapeRange.RelativeHorizontalPosition = _
        wdRelativeHorizontalPositionMargin
    Selection.ShapeRange.Left = wdShapeLeft
End Sub
Sub �}�`���E����()
    Selection.ShapeRange.RelativeHorizontalPosition = _
        wdRelativeHorizontalPositionMargin
    Selection.ShapeRange.Left = wdShapeCenter
End Sub



Sub �}�`��[()
    Selection.ShapeRange.RelativeVerticalPosition = _
        wdRelativeVerticalPositionMargin
    Selection.ShapeRange.Top = wdShapeTop
End Sub
Sub �}�`���[()
    Selection.ShapeRange.RelativeVerticalPosition = _
        wdRelativeVerticalPositionMargin
    Selection.ShapeRange.Top = wdShapeBottom
End Sub
Sub �}�`�㉺����()
    Selection.ShapeRange.RelativeVerticalPosition = _
        wdRelativeVerticalPositionMargin
    Selection.ShapeRange.Top = wdShapeCenter
End Sub



Sub �A���J�[�Œ�()
'
' �A���J�[�Œ� Macro
' �L�^�� 2011/09/16 �L�^�� �c��
'
'  Dim s
 ' For Each s In Selection.ShapeRang
  Selection.ShapeRange.LockAnchor = True
End Sub


Sub �A���J�[�t���[()
'
' �A���J�[�Œ� Macro
' �L�^�� 2011/09/16 �L�^�� �c��
'
'  Dim s
 ' For Each s In Selection.ShapeRang
  Selection.ShapeRange.LockAnchor = False
End Sub

