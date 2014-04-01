Attribute VB_Name = "スタイル"
Option Explicit

Sub 段落前で改ページのトグル()
'
'
    With Selection.ParagraphFormat
        .PageBreakBefore = Not .PageBreakBefore
    End With
End Sub

Public Function MainList() As ListTemplate
    
    ' 見出し 1 でつかっているテンプレートを返す様に変更してみた． 2013/3/5
    Set MainList = ActiveDocument.Styles("見出し 1").ListTemplate
    If MainList Is Nothing Then
    
      ' テンプレートの番号が3なのは意味がある模様．
      ' どうやら0,1,2は章立てではなく，レベル有の箇条書き用に予約されている模様
      Set MainList = Application.ListGalleries(wdOutlineNumberGallery).ListTemplates(3)
    End If

End Function


Sub メインリストのリセット()
    メインリストリセットレベル1 単独更新:=False
    メインリストリセットレベル2 単独更新:=False
    メインリストリセットレベル3 単独更新:=False
    メインリストリセットレベル4 単独更新:=False
    メインリストリセットレベル5 単独更新:=False
    メインリストリセットレベル6 単独更新:=False
    メインリストリセットレベル7 単独更新:=False
    メインリストリセットレベル8 単独更新:=False
    メインリストリセットレベル9 単独更新:=False
    
    ' リストのリセット
    Dim s As String
    Selection.Collapse
    s = Selection.Style
    Selection.Range.Style = "見出し 1"
    Selection.Range.ListFormat.ApplyListTemplate MainList(), , wdListApplyToWholeList, wdWord9ListBehavior
    Selection.Range.Style = s
    'ActiveDocument.Content.ListFormat.ApplyListTemplate ListTemplate:=MainList()
End Sub

Sub メインリストリセットレベル1(Optional 単独更新 As Boolean = True)
    With MainList().ListLevels(1)
        .NumberFormat = "%1."
        .TrailingCharacter = wdTrailingSpace
        .NumberStyle = wdListNumberStyleArabic
        .NumberPosition = MillimetersToPoints(0)
        .Alignment = wdListLevelAlignLeft
        .TextPosition = MillimetersToPoints(0)
        .TabPosition = wdUndefined
        .ResetOnHigher = False
        .StartAt = 1
        With .Font
            .Bold = wdUndefined
            .Italic = wdUndefined
            .StrikeThrough = wdUndefined
            .Subscript = wdUndefined
            .Superscript = wdUndefined
            .Shadow = wdUndefined
            .Outline = wdUndefined
            .Emboss = wdUndefined
            .Engrave = wdUndefined
            .AllCaps = wdUndefined
            .Hidden = wdUndefined
            .Underline = wdUndefined
            .Color = wdUndefined
            .Size = wdUndefined
            .Animation = wdUndefined
            .DoubleStrikeThrough = wdUndefined
            .name = "Alial"
        End With
        .LinkedStyle = "見出し 1"
        If 単独更新 Then
          ActiveDocument.Styles(.LinkedStyle).LinkToListTemplate ListTemplate:=MainList(), ListLevelNumber:=1
            ' リストのリセット
            Dim s As String
            Selection.Collapse
            s = Selection.Style
            Selection.Range.Style = .LinkedStyle
            Selection.Range.ListFormat.ApplyListTemplate MainList(), , wdListApplyToWholeList, wdWord9ListBehavior
            Selection.Range.Style = s
            End If
    End With
End Sub

 Sub メインリストリセットレベル2(Optional 単独更新 As Boolean = True)
    With MainList().ListLevels(2)
        .NumberFormat = "%1.%2"
        .TrailingCharacter = wdTrailingSpace
        .NumberStyle = wdListNumberStyleArabic
        .NumberPosition = MillimetersToPoints(0)
        .Alignment = wdListLevelAlignLeft
        .TextPosition = MillimetersToPoints(0)
        .TabPosition = wdUndefined
        .ResetOnHigher = 1
        .StartAt = 1
        With .Font
            .Bold = wdUndefined
            .Italic = wdUndefined
            .StrikeThrough = wdUndefined
            .Subscript = wdUndefined
            .Superscript = wdUndefined
            .Shadow = wdUndefined
            .Outline = wdUndefined
            .Emboss = wdUndefined
            .Engrave = wdUndefined
            .AllCaps = wdUndefined
            .Hidden = wdUndefined
            .Underline = wdUndefined
            .Color = wdUndefined
            .Size = wdUndefined
            .Animation = wdUndefined
            .DoubleStrikeThrough = wdUndefined
            .name = ""
        End With
        .LinkedStyle = "見出し 2"
        If 単独更新 Then
        ActiveDocument.Styles(.LinkedStyle).LinkToListTemplate ListTemplate:=MainList(), ListLevelNumber:=2
            ' リストのリセット
            Dim s As String
            Selection.Collapse
            s = Selection.Style
            Selection.Range.Style = .LinkedStyle
            Selection.Range.ListFormat.ApplyListTemplate MainList(), , wdListApplyToWholeList, wdWord9ListBehavior
            Selection.Range.Style = s
            End If
    End With
End Sub

 Sub メインリストリセットレベル3(Optional 単独更新 As Boolean = True)
    With MainList().ListLevels(3)
        .NumberFormat = "%1.%2.%3"
        .TrailingCharacter = wdTrailingSpace
        .NumberStyle = wdListNumberStyleArabic
        .NumberPosition = MillimetersToPoints(0)
        .Alignment = wdListLevelAlignLeft
        .TextPosition = MillimetersToPoints(0)
        .TabPosition = wdUndefined
        .ResetOnHigher = 2
        .StartAt = 1
        With .Font
            .Bold = wdUndefined
            .Italic = wdUndefined
            .StrikeThrough = wdUndefined
            .Subscript = wdUndefined
            .Superscript = wdUndefined
            .Shadow = wdUndefined
            .Outline = wdUndefined
            .Emboss = wdUndefined
            .Engrave = wdUndefined
            .AllCaps = wdUndefined
            .Hidden = wdUndefined
            .Underline = wdUndefined
            .Color = wdUndefined
            .Size = wdUndefined
            .Animation = wdUndefined
            .DoubleStrikeThrough = wdUndefined
            .name = ""
        End With
        .LinkedStyle = "見出し 3"
        If 単独更新 Then
        ActiveDocument.Styles(.LinkedStyle).LinkToListTemplate ListTemplate:=MainList(), ListLevelNumber:=3
            ' リストのリセット
            Dim s As String
            Selection.Collapse
            s = Selection.Style
            Selection.Range.Style = .LinkedStyle
            Selection.Range.ListFormat.ApplyListTemplate MainList(), , wdListApplyToWholeList, wdWord9ListBehavior
            Selection.Range.Style = s
            End If
    End With
End Sub

 Sub メインリストリセットレベル4(Optional 単独更新 As Boolean = True)
    With MainList().ListLevels(4)
        .NumberFormat = "(%4)"
        .TrailingCharacter = wdTrailingSpace
        .NumberStyle = wdListNumberStyleArabic
        .NumberPosition = MillimetersToPoints(0)
        .Alignment = wdListLevelAlignLeft
        .TextPosition = MillimetersToPoints(0)
        .TabPosition = wdUndefined
        .ResetOnHigher = 3
        .StartAt = 1
        With .Font
            .Bold = wdUndefined
            .Italic = wdUndefined
            .StrikeThrough = wdUndefined
            .Subscript = wdUndefined
            .Superscript = wdUndefined
            .Shadow = wdUndefined
            .Outline = wdUndefined
            .Emboss = wdUndefined
            .Engrave = wdUndefined
            .AllCaps = wdUndefined
            .Hidden = wdUndefined
            .Underline = wdUndefined
            .Color = wdUndefined
            .Size = wdUndefined
            .Animation = wdUndefined
            .DoubleStrikeThrough = wdUndefined
            .name = ""
        End With
        .LinkedStyle = "見出し 4"
         If 単独更新 Then
           ActiveDocument.Styles(.LinkedStyle).LinkToListTemplate ListTemplate:=MainList(), ListLevelNumber:=4
            ' リストのリセット
            Dim s As String
            Selection.Collapse
            s = Selection.Style
            Selection.Range.Style = .LinkedStyle
            Selection.Range.ListFormat.ApplyListTemplate MainList(), , wdListApplyToWholeList, wdWord9ListBehavior
            Selection.Range.Style = s
            End If
    End With
    
End Sub

 Sub メインリストリセットレベル5(Optional 単独更新 As Boolean = True)
    
    With MainList().ListLevels(5)
        .NumberFormat = "%5)"
        .TrailingCharacter = wdTrailingSpace
        .NumberStyle = wdListNumberStyleLowercaseLetter
        .NumberPosition = MillimetersToPoints(0)
        .Alignment = wdListLevelAlignLeft
        .TextPosition = MillimetersToPoints(0)
        .TabPosition = wdUndefined
        .ResetOnHigher = 4
        .StartAt = 1
        With .Font
            .Bold = wdUndefined
            .Italic = wdUndefined
            .StrikeThrough = wdUndefined
            .Subscript = wdUndefined
            .Superscript = wdUndefined
            .Shadow = wdUndefined
            .Outline = wdUndefined
            .Emboss = wdUndefined
            .Engrave = wdUndefined
            .AllCaps = wdUndefined
            .Hidden = wdUndefined
            .Underline = wdUndefined
            .Color = wdUndefined
            .Size = wdUndefined
            .Animation = wdUndefined
            .DoubleStrikeThrough = wdUndefined
            .name = ""
        End With
        .LinkedStyle = "見出し 5"
        If 単独更新 Then
        ActiveDocument.Styles(.LinkedStyle).LinkToListTemplate ListTemplate:=MainList(), ListLevelNumber:=5
            ' リストのリセット
            Dim s As String
            Selection.Collapse
            s = Selection.Style
            Selection.Range.Style = .LinkedStyle
            Selection.Range.ListFormat.ApplyListTemplate MainList(), , wdListApplyToWholeList, wdWord9ListBehavior
            Selection.Range.Style = s
            End If
    End With
End Sub

'図
 Sub メインリストリセットレベル6(Optional 単独更新 As Boolean = True)
    With MainList().ListLevels(6)
        .NumberFormat = "図%6"
        .TrailingCharacter = wdTrailingNone
        .NumberStyle = wdListNumberStyleArabic
        .NumberPosition = MillimetersToPoints(0)
        .Alignment = wdListLevelAlignLeft
        .TextPosition = MillimetersToPoints(0)
        .TabPosition = wdUndefined
        .ResetOnHigher = False
        .StartAt = 1
        With .Font
            .Bold = wdUndefined
            .Italic = wdUndefined
            .StrikeThrough = wdUndefined
            .Subscript = wdUndefined
            .Superscript = wdUndefined
            .Shadow = wdUndefined
            .Outline = wdUndefined
            .Emboss = wdUndefined
            .Engrave = wdUndefined
            .AllCaps = wdUndefined
            .Hidden = wdUndefined
            .Underline = wdUndefined
            .Color = wdUndefined
            .Size = wdUndefined
            .Animation = wdUndefined
            .DoubleStrikeThrough = wdUndefined
            .name = ""
        End With
        .LinkedStyle = "図"
        If 単独更新 Then
        ActiveDocument.Styles(.LinkedStyle).LinkToListTemplate ListTemplate:=MainList(), ListLevelNumber:=6
            ' リストのリセット
            Dim s As String
            Selection.Collapse
            s = Selection.Style
            Selection.Range.Style = .LinkedStyle
            Selection.Range.ListFormat.ApplyListTemplate MainList(), , wdListApplyToWholeList, wdWord9ListBehavior
            Selection.Range.Style = s
            End If
    End With
End Sub
 Sub メインリストリセットレベル7(Optional 単独更新 As Boolean = True)
    With MainList().ListLevels(7)
        .NumberFormat = "表%7"
        .TrailingCharacter = wdTrailingNone
        .NumberStyle = wdListNumberStyleArabic
        .NumberPosition = MillimetersToPoints(0)
        .Alignment = wdListLevelAlignLeft
        .TextPosition = MillimetersToPoints(0)
        .TabPosition = wdUndefined
        .ResetOnHigher = False
        .StartAt = 1
        With .Font
            .Bold = wdUndefined
            .Italic = wdUndefined
            .StrikeThrough = wdUndefined
            .Subscript = wdUndefined
            .Superscript = wdUndefined
            .Shadow = wdUndefined
            .Outline = wdUndefined
            .Emboss = wdUndefined
            .Engrave = wdUndefined
            .AllCaps = wdUndefined
            .Hidden = wdUndefined
            .Underline = wdUndefined
            .Color = wdUndefined
            .Size = wdUndefined
            .Animation = wdUndefined
            .DoubleStrikeThrough = wdUndefined
            .name = ""
        End With
        .LinkedStyle = "表"
        If 単独更新 Then
        ActiveDocument.Styles(.LinkedStyle).LinkToListTemplate ListTemplate:=MainList(), ListLevelNumber:=7
            ' リストのリセット
            Dim s As String
            Selection.Collapse
            s = Selection.Style
            Selection.Range.Style = .LinkedStyle
            Selection.Range.ListFormat.ApplyListTemplate MainList(), , wdListApplyToWholeList, wdWord9ListBehavior
            Selection.Range.Style = s
            End If
    End With
End Sub

 Sub メインリストリセットレベル8(Optional 単独更新 As Boolean = True)
With MainList().ListLevels(8)
        .NumberFormat = "(%8)"
        .TrailingCharacter = wdTrailingNone
        .NumberStyle = wdListNumberStyleLowercaseLetter
        .NumberPosition = MillimetersToPoints(0)
        .Alignment = wdListLevelAlignLeft
        .TextPosition = MillimetersToPoints(0)
        .TabPosition = wdUndefined
        .ResetOnHigher = 6
        .StartAt = 1
        With .Font
            .Bold = wdUndefined
            .Italic = wdUndefined
            .StrikeThrough = wdUndefined
            .Subscript = wdUndefined
            .Superscript = wdUndefined
            .Shadow = wdUndefined
            .Outline = wdUndefined
            .Emboss = wdUndefined
            .Engrave = wdUndefined
            .AllCaps = wdUndefined
            .Hidden = wdUndefined
            .Underline = wdUndefined
            .Color = wdUndefined
            .Size = wdUndefined
            .Animation = wdUndefined
            .DoubleStrikeThrough = wdUndefined
            .name = ""
        End With
        .LinkedStyle = "図副題"
        If 単独更新 Then
        ActiveDocument.Styles(.LinkedStyle).LinkToListTemplate ListTemplate:=MainList(), ListLevelNumber:=8
            ' リストのリセット
            Dim s As String
            Selection.Collapse
            s = Selection.Style
            Selection.Range.Style = .LinkedStyle
            Selection.Range.ListFormat.ApplyListTemplate MainList(), , wdListApplyToWholeList, wdWord9ListBehavior
            Selection.Range.Style = s
            End If
    End With
End Sub
 Sub メインリストリセットレベル9(Optional 単独更新 As Boolean = True)
    With MainList().ListLevels(9)
        .NumberFormat = "%9"
        .TrailingCharacter = wdTrailingNone
        .NumberStyle = wdListNumberStyleNumberInCircle
        .NumberPosition = MillimetersToPoints(7)
        .Alignment = wdListLevelAlignLeft
        .TextPosition = MillimetersToPoints(10)
        .TabPosition = MillimetersToPoints(10)
        .ResetOnHigher = 1
        .StartAt = 1
        With .Font
            .Bold = wdUndefined
            .Italic = wdUndefined
            .StrikeThrough = wdUndefined
            .Subscript = wdUndefined
            .Superscript = wdUndefined
            .Shadow = wdUndefined
            .Outline = wdUndefined
            .Emboss = wdUndefined
            .Engrave = wdUndefined
            .AllCaps = wdUndefined
            .Hidden = wdUndefined
            .Underline = wdUndefined
            .Color = wdUndefined
            .Size = wdUndefined
            .Animation = wdUndefined
            .DoubleStrikeThrough = wdUndefined
            .name = ""
        End With
        .LinkedStyle = "列挙"
        If 単独更新 Then
        ActiveDocument.Styles(.LinkedStyle).LinkToListTemplate ListTemplate:=MainList(), ListLevelNumber:=9
            ' リストのリセット
            Dim s As String
            Selection.Collapse
            s = Selection.Style
            Selection.Range.Style = .LinkedStyle
            Selection.Range.ListFormat.ApplyListTemplate MainList(), , wdListApplyToWholeList, wdWord9ListBehavior
            Selection.Range.Style = s
            End If
    End With
End Sub


Private Sub 反転数式行追加()
  On Error GoTo not_exist
  If Application.IsObjectValid(ActiveDocument.Styles("反転数式行")) Then Exit Sub
not_exist:
  Dim textwidth As Single
  With Selection.PageSetup
    textwidth = .PageWidth - .LeftMargin - .RightMargin
  End With
  With ActiveDocument.Styles.Add("反転数式行", WdStyleType.wdStyleTypeParagraph)
    .BaseStyle = "標準"
    .ParagraphFormat.TabStops.Add textwidth * 0.5, WdTabAlignment.wdAlignTabCenter
'    .ParagraphFormat.TabStops.Add textwidth, WdTabAlignment.wdAlignTabRight
    .ParagraphFormat.SpaceAfter = .ParagraphFormat.LineSpacing
    .ParagraphFormat.SpaceBefore = .ParagraphFormat.LineSpacing
    .ParagraphFormat.ReadingOrder = wdReadingOrderRtl
    .ParagraphFormat.OutlineLevel = wdOutlineLevelBodyText
    .NoSpaceBetweenParagraphsOfSameStyle = True
    .NextParagraphStyle = "本文"
  End With
End Sub

 Sub レベル9を章番号付き数式行に()
    Call 反転数式行追加
    With MainList().ListLevels(9)
        .NumberFormat = "(%1.%9)"
        .TrailingCharacter = wdTrailingNone
        .NumberStyle = wdListNumberStyleArabic
        .NumberPosition = MillimetersToPoints(0)
        .Alignment = wdListLevelAlignLeft
        .TextPosition = MillimetersToPoints(0)
        .TabPosition = MillimetersToPoints(0)
        .ResetOnHigher = 1
        .StartAt = 1
        With .Font
            .Bold = wdUndefined
            .Italic = wdUndefined
            .StrikeThrough = wdUndefined
            .Subscript = wdUndefined
            .Superscript = wdUndefined
            .Shadow = wdUndefined
            .Outline = wdUndefined
            .Emboss = wdUndefined
            .Engrave = wdUndefined
            .AllCaps = wdUndefined
            .Hidden = wdUndefined
            .Underline = wdUndefined
            .Color = wdUndefined
            .Size = wdUndefined
            .Animation = wdUndefined
            .DoubleStrikeThrough = wdUndefined
            .name = ""
        End With
        .LinkedStyle = "反転数式行"
    End With
End Sub

 Sub レベル9を節番号付き数式行に()
    Call 反転数式行追加
    With MainList().ListLevels(9)
        .NumberFormat = "(%1.%2.%9)"
        .TrailingCharacter = wdTrailingNone
        .NumberStyle = wdListNumberStyleArabic
        .NumberPosition = MillimetersToPoints(0)
        .Alignment = wdListLevelAlignLeft
        .TextPosition = MillimetersToPoints(0)
        .TabPosition = MillimetersToPoints(0)
        .ResetOnHigher = 2
        .StartAt = 1
        With .Font
            .Bold = wdUndefined
            .Italic = wdUndefined
            .StrikeThrough = wdUndefined
            .Subscript = wdUndefined
            .Superscript = wdUndefined
            .Shadow = wdUndefined
            .Outline = wdUndefined
            .Emboss = wdUndefined
            .Engrave = wdUndefined
            .AllCaps = wdUndefined
            .Hidden = wdUndefined
            .Underline = wdUndefined
            .Color = wdUndefined
            .Size = wdUndefined
            .Animation = wdUndefined
            .DoubleStrikeThrough = wdUndefined
            .name = ""
        End With
        .LinkedStyle = "反転数式行"
    End With
End Sub




Sub スタイルの内容をコピー(ByVal 元 As String, ByVal 先 As String)
  Dim s As Style
  
  Set s = ActiveDocument.Styles(元)

  With ActiveDocument.Styles(先)
    .AutomaticallyUpdate = False
    With .Font
        .NameFarEast = s.Font.NameFarEast
        .NameAscii = s.Font.NameAscii
        .NameOther = s.Font.NameOther
        .name = s.Font.name
        .Size = s.Font.Size
        .Bold = s.Font.Bold
        .Italic = s.Font.Italic
        .Underline = s.Font.Underline
        .UnderlineColor = s.Font.UnderlineColor
        .StrikeThrough = s.Font.StrikeThrough
        .DoubleStrikeThrough = s.Font.DoubleStrikeThrough
        .Outline = s.Font.Outline
        .Emboss = s.Font.Emboss
        .Shadow = s.Font.Shadow
        .Hidden = s.Font.Hidden
        .SmallCaps = s.Font.SmallCaps
        .AllCaps = s.Font.AllCaps
        .Color = s.Font.Color
        .Engrave = s.Font.Engrave
        .Superscript = s.Font.Superscript
        .Subscript = s.Font.Subscript
        .Scaling = s.Font.Scaling
        .Kerning = s.Font.Kerning
        .Animation = s.Font.Animation
        .DisableCharacterSpaceGrid = s.Font.DisableCharacterSpaceGrid
        .EmphasisMark = s.Font.EmphasisMark
    End With
    Dim p As ParagraphFormat
    Set p = s.ParagraphFormat
    With .ParagraphFormat
        .LeftIndent = p.LeftIndent
        .RightIndent = p.RightIndent
        .SpaceBefore = p.SpaceBefore
        .SpaceBeforeAuto = p.SpaceBeforeAuto
        .SpaceAfter = p.SpaceAfter
        .SpaceAfterAuto = p.SpaceAfterAuto
        .LineSpacingRule = p.LineSpacingRule
        .Alignment = p.Alignment
        .WidowControl = p.WidowControl
        .KeepWithNext = p.KeepWithNext
        .KeepTogether = p.KeepTogether
        .PageBreakBefore = p.PageBreakBefore
        .NoLineNumber = p.NoLineNumber
        .Hyphenation = p.Hyphenation
        .FirstLineIndent = p.FirstLineIndent
        .OutlineLevel = p.OutlineLevel
        .CharacterUnitLeftIndent = p.CharacterUnitLeftIndent
        .CharacterUnitRightIndent = p.CharacterUnitRightIndent
        .CharacterUnitFirstLineIndent = p.CharacterUnitFirstLineIndent
        .LineUnitBefore = p.LineUnitBefore
        .LineUnitAfter = p.LineUnitAfter
        .AutoAdjustRightIndent = p.AutoAdjustRightIndent
        .DisableLineHeightGrid = p.DisableLineHeightGrid
        .FarEastLineBreakControl = p.FarEastLineBreakControl
        .WordWrap = p.WordWrap
        .HangingPunctuation = p.HangingPunctuation
        .HalfWidthPunctuationOnTopOfLine = p.HalfWidthPunctuationOnTopOfLine
        .AddSpaceBetweenFarEastAndAlpha = p.AddSpaceBetweenFarEastAndAlpha
        .AddSpaceBetweenFarEastAndDigit = p.AddSpaceBetweenFarEastAndDigit
        .BaseLineAlignment = p.BaseLineAlignment
        .Borders(wdBorderLeft).LineStyle = p.Borders(wdBorderLeft).LineStyle
        .Borders(wdBorderRight).LineStyle = p.Borders(wdBorderRight).LineStyle
        .Borders(wdBorderTop).LineStyle = p.Borders(wdBorderTop).LineStyle
        .Borders(wdBorderBottom).LineStyle = p.Borders(wdBorderBottom).LineStyle
        .Borders.DistanceFromTop = p.Borders.DistanceFromTop
        .Borders.DistanceFromLeft = p.Borders.DistanceFromLeft
        .Borders.DistanceFromBottom = p.Borders.DistanceFromBottom
        .Borders.DistanceFromRight = p.Borders.DistanceFromRight
        .Borders.Shadow = p.Borders.Shadow
    End With
    .NoSpaceBetweenParagraphsOfSameStyle = s.NoSpaceBetweenParagraphsOfSameStyle
    With .ParagraphFormat.TabStops
        .ClearAll
        Dim t As TabStop
        For Each t In s.ParagraphFormat.TabStops
            .Add t.Position, t.Alignment, t.Leader
        Next
    End With
    With .ParagraphFormat.Shading
        .Texture = p.Shading.Texture
        .ForegroundPatternColor = p.Shading.ForegroundPatternColor
        .BackgroundPatternColor = p.Shading.BackgroundPatternColor
    End With
    .LanguageID = s.LanguageID
    .NoProofing = s.NoProofing
    .LanguageID = s.LanguageID
    .NoProofing = s.NoProofing
    .Frame.Delete
  End With

End Sub


Sub 数式行スタイル確認()
  On Error GoTo add_style:
  Dim n
  n = ActiveDocument.Styles("数式行").BuiltIn
  Exit Sub

add_style:
  Dim textwidth As Single
  With Selection.PageSetup
    textwidth = .PageWidth - .LeftMargin - .RightMargin
  End With
  With ActiveDocument.Styles.Add("数式行", wdStyleTypeParagraph)
    .ParagraphFormat.TabStops.Add textwidth * 0.5, WdTabAlignment.wdAlignTabCenter
    .ParagraphFormat.TabStops.Add textwidth, WdTabAlignment.wdAlignTabRight
    .ParagraphFormat.SpaceAfter = .ParagraphFormat.LineSpacing
    .ParagraphFormat.SpaceBefore = .ParagraphFormat.LineSpacing
'    .ParagraphFormat.ReadingOrder = wdReadingOrderRtl
    .NextParagraphStyle = "本文"
  End With
End Sub

Sub 数式行の挿入()

  数式行スタイル確認

  Selection.Paragraphs(1).Style = "数式行"

  Selection.InsertBefore vbTab
  Selection.Collapse wdCollapseEnd
  Selection.InsertBefore vbTab & "( )"
  ActiveDocument.Fields.Add ActiveDocument.Range(Selection.Range.End - 2, Selection.Range.End - 1), wdFieldListNum, "Main \l 9"
  Selection.Collapse wdCollapseStart
End Sub


