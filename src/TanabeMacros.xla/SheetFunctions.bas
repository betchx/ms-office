Attribute VB_Name = "SheetFunctions"
Option Explicit


Public Function 鉄筋断面積(ByVal 径 As Variant, _
                            Optional ByVal 異形鉄筋 As Boolean = False) As Double
    If IsObject(径) Then
        Select Case TypeName(径)
        Case "Range"
            Dim r As Range
            Set r = 径
            If r.Columns.Count > 1 Then
                If r.Rows.Count > 1 Then
                    径 = 0 ' 範囲が広すぎて判定できないため
                Else
                    径 = r.Columns(1, Application.ThisCell.Column).Value
                End If
            Else
                If r.Rows.Count > 1 Then
                    径 = r.Rows(Application.ThisCell.row).Value
                Else
                    If r.Value = "" Then
                      径 = 0#
                      Exit Function
                    End If
                    径 = r.Value
                End If
            End If
        Case Else
            MsgBox "タイプ(" & TypeName(径) & ")は現在サポートしていません"
        End Select
    End If

    If Not IsNumeric(径) Then
        If Left(径, 1) = "D" Then
            径 = CInt(val(Mid(径, 2)))
            異形鉄筋 = True
        ElseIf Left(径, 1) = "φ" Then
            径 = val(Mid(径, 2))
            異形鉄筋 = False
        Else
            径 = 0
        End If
    End If
    If 異形鉄筋 Then
      If 径 > 51 Or 径 < 4 Then
        鉄筋断面積 = 0
      Else
        鉄筋断面積 = _
          Array(0, 0, 0, 0, 14.05, 21.98, 31.67, 0, 49.51, 0, 71.33, _
                    0, 0, 126.7, 0, 0, 198.6, 0, 0, 286.5, 0, _
                    0, 387.1, 0, 0, 506.7, 0, 0, 0, 642.4, 0, _
                    0, 794.2, 0, 0, 956.6, 0, 0, 1140#, 0, 0, _
                 1340#, 0, 0, 0, 0, 0, 0, 0, 0, 0, _
                 2027#)(径)
      End If
    Else
      鉄筋断面積 = 径 ^ 2 * 0.25 * WorksheetFunction.Pi()
    End If
    If 鉄筋断面積 = 0 Then 鉄筋断面積 = CVErr(xlErrNA)

End Function

