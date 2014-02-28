Attribute VB_Name = "図形位置調整"
Option Explicit


Sub 図形右端()
    Selection.ShapeRange.RelativeHorizontalPosition = _
        wdRelativeHorizontalPositionMargin
    Selection.ShapeRange.Left = wdShapeRight
End Sub
Sub 図形左端()
    Selection.ShapeRange.RelativeHorizontalPosition = _
        wdRelativeHorizontalPositionMargin
    Selection.ShapeRange.Left = wdShapeLeft
End Sub
Sub 図形左右中央()
    Selection.ShapeRange.RelativeHorizontalPosition = _
        wdRelativeHorizontalPositionMargin
    Selection.ShapeRange.Left = wdShapeCenter
End Sub



Sub 図形上端()
    Selection.ShapeRange.RelativeVerticalPosition = _
        wdRelativeVerticalPositionMargin
    Selection.ShapeRange.Top = wdShapeTop
End Sub
Sub 図形下端()
    Selection.ShapeRange.RelativeVerticalPosition = _
        wdRelativeVerticalPositionMargin
    Selection.ShapeRange.Top = wdShapeBottom
End Sub
Sub 図形上下中央()
    Selection.ShapeRange.RelativeVerticalPosition = _
        wdRelativeVerticalPositionMargin
    Selection.ShapeRange.Top = wdShapeCenter
End Sub



Sub アンカー固定()
'
' アンカー固定 Macro
' 記録日 2011/09/16 記録者 田辺
'
'  Dim s
 ' For Each s In Selection.ShapeRang
  Selection.ShapeRange.LockAnchor = True
End Sub


Sub アンカーフリー()
'
' アンカー固定 Macro
' 記録日 2011/09/16 記録者 田辺
'
'  Dim s
 ' For Each s In Selection.ShapeRang
  Selection.ShapeRange.LockAnchor = False
End Sub

