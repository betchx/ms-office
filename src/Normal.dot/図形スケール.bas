Attribute VB_Name = "図形スケール"
Option Explicit
Const app As String = "WordMacro"
Const sec As String = "Triming"

Sub 一括スケール()
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
  Const message As String = "サイズを % か 数値（単位はmm) で指定してください．" & vbCrLf _
    & " 最後に '%' がある場合 " & vbCrLf _
    & "      正の値は現在のサイズに対する割合でスケールします" & vbCrLf _
    & "      負の値の場合は可能ならばオリジナルサイズに対する割合でスケールします．" & vbCrLf _
    & " 数値で指定した場合は幅もしくは高さがその値となる様にスケールします:" & vbCrLf _
    & "      正の値の場合は幅," & vbCrLf _
    & "      負の値の場合は高さ "
  
  res = StrConv(InputBox(message, "サイズ設定", GetSetting(app, "size", key, "100%")), vbNarrow)
  
  If res = "" Then Exit Sub
  
  SaveSetting app, "size", key, res
  
  pos = InStr(res, "%")
  If pos > 0 Then
    ' %ありなので割合でScale
    s = CSng(Left(res, pos - 1))
    If s < 0 Then
       ' scale by original
        s = s * -1
        選択範囲内の図を割合でスケール s, 元のサイズから:=True
    Else
        選択範囲内の図を割合でスケール s
    End If
  Else
    s = CSng(res)
    If s > 0# Then
        選択範囲内の図を幅でスケール s
    Else
        選択範囲内の図を高さでスケール -s
    End If
  End If
End Sub

Private Sub 選択範囲内の図を割合でスケール(割合 As Single, Optional 元のサイズから As Boolean = False)
    Dim i
    Dim 可能な対象 As Boolean
    Dim p As Paragraph
    Dim r As Range
    
    Select Case Selection.Type
    Case wdSelectionInlineShape
        For Each i In Selection.InlineShapes
            i.ScaleHeight = 割合
            i.ScaleWidth = 割合
        Next
    Case wdSelectionShape
        For Each i In Selection.ShapeRange
            可能な対象 = i.Type = msoPicture Or i.Type = msoOLEControlObject
            i.ScaleHeight 割合 * 0.01, 元のサイズから And 可能な対象
            i.ScaleWidth 割合 * 0.01, 元のサイズから And 可能な対象
        Next
    Case wdSelectionNormal
        For Each p In Selection.Paragraphs
            Set r = p.Range
            If r.InlineShapes.Count > 0 Then
                For Each i In p.Range.InlineShapes
                    i.ScaleHeight = 割合
                    i.ScaleWidth = 割合
                Next i
            End If
        Next p
        If Selection.ShapeRange.Count > 0 Then
            For Each i In Selection.ShapeRange
                可能な対象 = i.Type = msoPicture Or i.Type = msoOLEControlObject
                i.ScaleHeight 割合 * 0.01, 元のサイズから And 可能な対象
                i.ScaleWidth 割合 * 0.01, 元のサイズから And 可能な対象
            Next i
        End If
    End Select

End Sub

Private Sub 選択範囲内の図を幅でスケール(mm As Single)
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
            If シェープあり(r) Then
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

Private Function シェープあり(r As Range) As Boolean
  On Error GoTo eee:
  If r.ShapeRange.Count > 0 Then
    シェープあり = True
  End If
  Exit Function
eee:
  シェープあり = False
End Function


Private Sub 選択範囲内の図を高さでスケール(mm As Single)
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




