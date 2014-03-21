Attribute VB_Name = "トリミング"
Option Explicit

Const app As String = "WordMacro"
Const sec As String = "Triming"
Const keyBottom As String = "bottom"
Const keyTop As String = "top"
Const keyLeft As String = "left"
Const keyRight As String = "right"


Sub トリミング情報保存()
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

Private Sub トリミング情報の反映(ByRef pf As PictureFormat)
    pf.CropBottom = CSng(GetSetting(app, sec, keyBottom, "0.0"))
    pf.CropTop = CSng(GetSetting(app, sec, keyTop, "0.0"))
    pf.CropLeft = CSng(GetSetting(app, sec, keyLeft, "0.0"))
    pf.CropRight = CSng(GetSetting(app, sec, keyRight, "0.0"))
End Sub


Sub トリミング情報設定()
    Dim s As InlineShape
    Dim p As Shape
    If Selection.Type = wdSelectionInlineShape Or Selection.InlineShapes.Count > 0 Then
        For Each s In Selection.InlineShapes
            'Set pf = Selection.InlineShapes(1).PictureFormat
            トリミング情報の反映 s.PictureFormat
        Next
    End If
    If Selection.Type = wdSelectionShape Then
        For Each p In Selection.ShapeRange
          トリミング情報の反映 p.PictureFormat
        Next
    End If
End Sub


