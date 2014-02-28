Attribute VB_Name = "図表設定"
Option Explicit

Function 図リストレベル() As ListLevel
  Set 図リストレベル = MainList().ListLevels(6)
End Function

Function 表リストレベル() As ListLevel
  Set 表リストレベル = MainList().ListLevels(7)
End Function

Sub 図表設定番号のみ()
  図リストレベル().NumberFormat = "図%6"
  表リストレベル().NumberFormat = "表%7"
  図リストレベル().ResetOnHigher = False
  表リストレベル().ResetOnHigher = False
  設定反映
End Sub

Sub 図表設定ハイフン()
  図リストレベル().NumberFormat = "図-%6"
  表リストレベル().NumberFormat = "表-%7"
  図リストレベル().ResetOnHigher = False
  表リストレベル().ResetOnHigher = False
  設定反映
End Sub

Sub 図表設定章番号()
  図リストレベル().NumberFormat = "図%1.%6"
  表リストレベル().NumberFormat = "表%1.%7"
  図リストレベル().ResetOnHigher = 1
  表リストレベル().ResetOnHigher = 1
  設定反映
End Sub


Sub 図表設定節番号()
  図リストレベル().NumberFormat = "図%1.%2.%6"
  表リストレベル().NumberFormat = "表%1.%2.%7"
  図リストレベル().ResetOnHigher = 2
  表リストレベル().ResetOnHigher = 2
  設定反映
End Sub

Sub 図表設定小節番号()
  図リストレベル().NumberFormat = "図%1.%2.%3.%6"
  表リストレベル().NumberFormat = "表%1.%2.%3.%7"
  図リストレベル().ResetOnHigher = 3
  表リストレベル().ResetOnHigher = 3
  設定反映
End Sub

Sub 図表表題部ゴシックのONOFF()
    If 図リストレベル().Font.name = "ＭＳ ゴシック" Then
       Dim f As Font
        Set f = ActiveDocument.Styles("本文").Font
        図リストレベル().Font.name = f.name '"ＭＳ 明朝"
        表リストレベル().Font.name = f.name '"ＭＳ 明朝"
    Else
        図リストレベル().Font.name = "ＭＳ ゴシック"
        表リストレベル().Font.name = "ＭＳ ゴシック"
    End If
    設定反映
End Sub



Private Sub 設定反映()
'  ActiveDocument.Content.ListFormat.ApplyListTemplate ListTemplate:=MainList()
  Dim p As Paragraph
  For Each p In ActiveDocument.Paragraphs
      If p.Style = "表" Or p.Style = "図" _
        Then p.Range.ListFormat.ApplyListTemplate ListTemplate:=MainList()
  Next
End Sub

Private Sub 表のスタイル設定(ByVal 元 As String)
  スタイルコピー 元, "表"
  ActiveDocument.Styles("表").BaseStyle = 元
  ActiveDocument.Styles("表").NextParagraphStyle = "表本体"
End Sub

Private Sub 図のスタイル設定(ByVal 元 As String)
  スタイルコピー 元, "図"
  ActiveDocument.Styles("図").BaseStyle = 元
  ActiveDocument.Styles("図").NextParagraphStyle = "本文"
End Sub


