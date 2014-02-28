Attribute VB_Name = "�g���~���O"
Option Explicit

Const app As String = "WordMacro"
Const sec As String = "Triming"
Const keyBottom As String = "bottom"
Const keyTop As String = "top"
Const keyLeft As String = "left"
Const keyRight As String = "right"


Sub �g���~���O���ۑ�()
    Dim pf As PictureFormat
    If Selection.Type = wdSelectionInlineShape Or Selection.InlineShapes.Count > 0 Then
        Set pf = Selection.InlineShapes(1).PictureFormat
    ElseIf Selection.Type = wdSelectionShape Or Selection.ShapeRange.Count > 0 Then
        Set pf = Selection.ShapeRange.PictureFormat
    Else
      Exit Sub
    End If
    SaveSetting app, sec, keyBottom, CStr(pf.CropBottom)
    SaveSetting app, sec, keyTop, CStr(pf.CropTop)
    SaveSetting app, sec, keyLeft, CStr(pf.CropLeft)
    SaveSetting app, sec, keyRight, CStr(pf.CropRight)
End Sub

Private Sub �g���~���O���̔��f(ByRef pf As PictureFormat)
    pf.CropBottom = CSng(GetSetting(app, sec, keyBottom, "0.0"))
    pf.CropTop = CSng(GetSetting(app, sec, keyTop, "0.0"))
    pf.CropLeft = CSng(GetSetting(app, sec, keyLeft, "0.0"))
    pf.CropRight = CSng(GetSetting(app, sec, keyRight, "0.0"))
End Sub


Sub �g���~���O���ݒ�()
    Dim s As InlineShape
    Dim p As Shape
    If Selection.Type = wdSelectionInlineShape Or Selection.InlineShapes.Count > 0 Then
        For Each s In Selection.InlineShapes
            'Set pf = Selection.InlineShapes(1).PictureFormat
            �g���~���O���̔��f s.PictureFormat
        Next
    End If
    If Selection.Type = wdSelectionShape Then
        For Each p In Selection.ShapeRange
          �g���~���O���̔��f p.PictureFormat
        Next
    End If
End Sub


