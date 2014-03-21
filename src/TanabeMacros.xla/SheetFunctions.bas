Attribute VB_Name = "SheetFunctions"
Option Explicit


Public Function “S‹Ø’f–ÊÏ(ByVal Œa As Variant, _
                            Optional ByVal ˆÙŒ`“S‹Ø As Boolean = False) As Double
    If IsObject(Œa) Then
        Select Case TypeName(Œa)
        Case "Range"
            Dim r As Range
            Set r = Œa
            If r.Columns.Count > 1 Then
                If r.Rows.Count > 1 Then
                    Œa = 0 ' ”ÍˆÍ‚ªL‚·‚¬‚Ä”»’è‚Å‚«‚È‚¢‚½‚ß
                Else
                    Œa = r.Columns(1, Application.ThisCell.Column).Value
                End If
            Else
                If r.Rows.Count > 1 Then
                    Œa = r.Rows(Application.ThisCell.row).Value
                Else
                    If r.Value = "" Then
                      Œa = 0#
                      Exit Function
                    End If
                    Œa = r.Value
                End If
            End If
        Case Else
            MsgBox "ƒ^ƒCƒv(" & TypeName(Œa) & ")‚ÍŒ»ÝƒTƒ|[ƒg‚µ‚Ä‚¢‚Ü‚¹‚ñ"
        End Select
    End If

    If Not IsNumeric(Œa) Then
        If Left(Œa, 1) = "D" Then
            Œa = CInt(val(Mid(Œa, 2)))
            ˆÙŒ`“S‹Ø = True
        ElseIf Left(Œa, 1) = "ƒÓ" Then
            Œa = val(Mid(Œa, 2))
            ˆÙŒ`“S‹Ø = False
        Else
            Œa = 0
        End If
    End If
    If ˆÙŒ`“S‹Ø Then
      If Œa > 51 Or Œa < 4 Then
        “S‹Ø’f–ÊÏ = 0
      Else
        “S‹Ø’f–ÊÏ = _
          Array(0, 0, 0, 0, 14.05, 21.98, 31.67, 0, 49.51, 0, 71.33, _
                    0, 0, 126.7, 0, 0, 198.6, 0, 0, 286.5, 0, _
                    0, 387.1, 0, 0, 506.7, 0, 0, 0, 642.4, 0, _
                    0, 794.2, 0, 0, 956.6, 0, 0, 1140#, 0, 0, _
                 1340#, 0, 0, 0, 0, 0, 0, 0, 0, 0, _
                 2027#)(Œa)
      End If
    Else
      “S‹Ø’f–ÊÏ = Œa ^ 2 * 0.25 * WorksheetFunction.Pi()
    End If
    If “S‹Ø’f–ÊÏ = 0 Then “S‹Ø’f–ÊÏ = CVErr(xlErrNA)

End Function

