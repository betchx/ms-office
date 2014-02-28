Attribute VB_Name = "SheetFunctions"
Option Explicit


Public Function �S�ؒf�ʐ�(ByVal �a As Variant, _
                            Optional ByVal �ٌ`�S�� As Boolean = False) As Double
    If IsObject(�a) Then
        Select Case TypeName(�a)
        Case "Range"
            Dim r As Range
            Set r = �a
            If r.Columns.Count > 1 Then
                If r.Rows.Count > 1 Then
                    �a = 0 ' �͈͂��L�����Ĕ���ł��Ȃ�����
                Else
                    �a = r.Columns(1, Application.ThisCell.Column).Value
                End If
            Else
                If r.Rows.Count > 1 Then
                    �a = r.Rows(Application.ThisCell.row).Value
                Else
                    If r.Value = "" Then
                      �a = 0#
                      Exit Function
                    End If
                    �a = r.Value
                End If
            End If
        Case Else
            MsgBox "�^�C�v(" & TypeName(�a) & ")�͌��݃T�|�[�g���Ă��܂���"
        End Select
    End If

    If Not IsNumeric(�a) Then
        If Left(�a, 1) = "D" Then
            �a = CInt(val(Mid(�a, 2)))
            �ٌ`�S�� = True
        ElseIf Left(�a, 1) = "��" Then
            �a = val(Mid(�a, 2))
            �ٌ`�S�� = False
        Else
            �a = 0
        End If
    End If
    If �ٌ`�S�� Then
      If �a > 51 Or �a < 4 Then
        �S�ؒf�ʐ� = 0
      Else
        �S�ؒf�ʐ� = _
          Array(0, 0, 0, 0, 14.05, 21.98, 31.67, 0, 49.51, 0, 71.33, _
                    0, 0, 126.7, 0, 0, 198.6, 0, 0, 286.5, 0, _
                    0, 387.1, 0, 0, 506.7, 0, 0, 0, 642.4, 0, _
                    0, 794.2, 0, 0, 956.6, 0, 0, 1140#, 0, 0, _
                 1340#, 0, 0, 0, 0, 0, 0, 0, 0, 0, _
                 2027#)(�a)
      End If
    Else
      �S�ؒf�ʐ� = �a ^ 2 * 0.25 * WorksheetFunction.Pi()
    End If
    If �S�ؒf�ʐ� = 0 Then �S�ؒf�ʐ� = CVErr(xlErrNA)

End Function

