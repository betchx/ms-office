Attribute VB_Name = "フィールド"
Option Explicit

Function ExcelEval(s As String)
  'Excel VBAのEvaluate関数を呼び出す
  ExcelEval = ExcelApplication.Evaluate(s)
End Function

Private Function ExcelApplication() As Excel.Application
' エクセルのApplicationオブジェクトへの参照を返す．
' エクセルがまだ起動されていない場合は起動する．

  On Error GoTo new_instance
  Set ExcelApplication = Excel.Application
  Exit Function
new_instance:
  If Err.Number = 429 Then
    Set ExcelApplication = New Word.Application
    Err.Clear
  ElseIf Err.Number <> 0 Then
    Set ExcelApplication = Nothing
  End If
End Function

Sub UpdateFields()
' フィールド更新のフック
'
' UpdateFields Macro
' 選択したフィールドの実行結果を更新して表示します。
'
'  Expression (=) フィールドを拡張し，単純評価でエラーがでた場合はExcel.Evaluateによる評価を試みる
'  また，Expression フィールドの2文字目が=の場合 （{ == xxxx }の形式の場合）は，最初からExcelで評価する．

  Dim f As Field
  Dim i As Long
  Dim s As String
  Dim res
  '
  If Selection.Start = Selection.End Then
    For Each f In ActiveDocument.Fields
      If f.Code.Start <= Selection.Start And f.Code.End >= Selection.End Then
        UpdateFieldWithExpressionCheck f
        Exit For ' ネストした内側のフィールドは別途更新されるので，ここでは更新不要
      End If
    Next
  Else
    For Each f In Selection.Fields
      UpdateFieldWithExpressionCheck f
    Next f
  End If
End Sub



Private Sub UpdateFieldWithExpressionCheck(ByRef f As Field)
' 数式かどうかを確認しつつ，フィールドを更新する．
' 数式で，==が2つ連続する場合は，エクセルで評価する．

  Dim s As String
  Dim i As Integer
  If f.Type = wdFieldExpression Then
    s = f.Code.text
    i = InStr(1, s, "=")
    If Mid(s, i + 1, 1) = "=" Then
     UpdateFieldWithExcelEvalEx f ', Mid(s, i + 2)
    Else
      If Not f.Update() Then
'          UpdateFieldWithExcelEval f, Mid(s, i + 1)
      End If
    End If
  Else
    f.Update
  End If
End Sub


Private Sub UpdateFieldWithExcelEval(ByRef f As Field, expr As String)
' 数式フィールドをエクセルのEvaluateで更新する．
' ブックマーク（変数）やスイッチ等に対応していないプロトタイプ．
' (すでに未使用だが，わかりやいのでメモとして残しておく）
  Dim res
  res = ExcelEval(expr)
  f.Update ' need
  If TypeName(res) <> "Error" Then
    ' Overwrite error
    With f.Result
      .text = res
      .Bold = False
      .Italic = False
    End With
  Else
    f.Result.text = "Excel対応数式にエラーがあります「" & expr & "」"
  End If
End Sub

Private Function isNumber(expr As String) As Boolean
' 引数が数字で始まるかどうかの判定．
' 今のところ利用していない．
  
  Dim Code As Integer
  Code = Asc(Left(expr, 1))
  isNumber = Code >= Asc("0") And Code <= Asc("9")
End Function

Private Sub UpdateFieldWithExcelEvalEx(ByRef f As Field)
' 数式フィールドをエクセルの数式として評価する
' 変数（ブックマーク）やスイッチに対応したバージョン
' ただし，ワードがもつ表のセルに対する参照機能は無くなる．

  Dim res
  Dim tokens As Words
  Dim i
  Dim expr As String
  Dim bm_name As String
  Dim c As String, a As Integer
  Dim field_Switch As String
  field_Switch = ""

  Set tokens = f.Code.Words
  Dim n As Integer
  n = tokens.Count
  For i = 1 To n
    c = Left(tokens(i).text, 1)
    a = Asc(c)
    Select Case c
    Case "=", " ", Chr(19), "　", Chr(13)
      ' これらは最初のみ
      expr = ""
    Case "\"
      ' スイッチ
      Dim k As Integer
      For k = i To tokens.Count
        field_Switch = field_Switch & tokens(k)
      Next
      Exit For
    Case "0", "1", "2", "3", "4", "5", "6", "7", "8", "9", ".", "+", "-", "*", "/", "(", ")", ",", """", "'", "<", ">", ""
      ' 数値や演算子とおもわれるものはそのまま追加
      expr = expr & tokens(i).text
    Case Else
      ' 文字列等なので，関数でなければ変数とみなしてブックマークを検索する．
            Dim isFunc As Boolean
      If i = tokens.Count Then
        isFunc = False
      ElseIf Left(tokens(i + 1).text, 1) = "(" Then
        isFunc = True
      Else
        isFunc = False
      End If

      If isFunc Then
        '関数はそのまま渡す．
        expr = expr & tokens(i)
      Else
        ' 変数と思われるので，ブックマークの解決を試みる
        On Error GoTo no_bookmark
        bm_name = Trim(tokens(i))
        Dim bm As Bookmark
        Set bm = ActiveDocument.Bookmarks(bm_name)
        Dim ff As Field
        For Each ff In bm.Range.Fields
          UpdateFieldWithExpressionCheck ff
        Next
        expr = expr & bm.Range.text
        On Error GoTo 0
      End If
    End Select
  Next
  
  ' 作成した数式をエクセルで評価する
  res = ExcelEval(expr)
  
  If TypeName(res) <> "Error" Then
    ' 手動でスイッチの処理を行うのは困難でバグの原因となるので，
    ' 評価した値でコードを差し替えた上で更新し，そのあとオリジナルのコードに戻す．
    Dim original_code_text As String
    original_code_text = f.Code.text
  '  f.code.text = Left(original_code_text, 2) & res & " " & field_Switch
    f.Code.text = "=" & res & " " & field_Switch
    f.Update
    f.Code.text = original_code_text
  Else
    f.Update 'エラーになるのがわかっていても更新作業は必要
    ' 結果をエラー文字列で上書きする．
    f.Result.text = "Excel対応数式にエラーがあります「" & expr & "」"
  End If
  
  ' 更新したら，コード表示を終了させる．（通常の更新での動作にあわせる）
  f.ShowCodes = False
  Exit Sub
  
no_bookmark:
  f.Update  ' エラーになるが更新は必要
  f.Result.text = "ブックマークがありません「" & bm_name & "」"
End Sub

