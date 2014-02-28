Attribute VB_Name = "Main"
Option Explicit


Sub リンク貼り付け()
'
' リンク貼り付け Macro
' マクロ記録日 : 2011/2/2  ユーザー名 : 田辺
'
On Error Resume Next

    ActiveSheet.Paste Link:=True
End Sub

Sub オブジェクトの高さ変更()

  Dim str As String
  Dim h As Single, h0 As Single
  Dim sh As shape
  Dim ca As ChartArea
  Dim name As String
  Dim ash As Worksheet
  
  On Error GoTo eee:
  
  name = ActiveWindow.Selection.name
  
  Set ca = ActiveWindow.Selection
  
  
  h0 = ca.Height
  
    
  str = InputBox("高さを設定してください. (正数: pt, 負数: mm)" & vbCrLf & _
  "現在値：" & CStr(h0) & "pt(" & pnt2mm(h0) & "mm)" & vbCrLf & _
  "半分：" & CStr(h0 / 2) & "pt  ２倍:" & CStr(h0 * 2) & "pt", _
  "高さの変更", h0)
    
  If Len(str) > 0 Then
    h = CSng(val(str))
    If h < 0# Then h = mm2pnt(-h)
    If h > 0# Then ca.Parent.Parent.Height = h
    
  End If
  
eee:

End Sub

Sub オブジェクトの幅変更()

  Dim str As String
  Dim w As Single, w0 As Single
  Dim sh As shape
  Dim ca As ChartArea
  Dim name As String
  Dim ash As Worksheet
  Dim w_half As Single, w2 As Single
  
  
  On Error GoTo eee:
  
  name = ActiveWindow.Selection.name
  
  Set ca = ActiveWindow.Selection
  
  
  w0 = ca.Width
    
  str = InputBox("幅を設定してください. (正数: pt, 負数: mm)" & vbCrLf & _
  "現在値：" & CStr(w0) & "pt(" & pnt2mm(w0) & "mm)" & vbCrLf & _
  "半分：" & CStr(w0 / 2) & "pt  ２倍:" & CStr(w0 * 2) & "pt", _
  "幅の変更", w0)
    
  If Len(str) > 0 Then
    w = CSng(val(str))
    If w < 0# Then w = mm2pnt(-w)
    If w > 0# Then ca.Parent.Parent.Width = w
    
  End If
  
eee:

End Sub


Sub カレントセルをシートタイトルに()
Attribute カレントセルをシートタイトルに.VB_ProcData.VB_Invoke_Func = "t\n14"
' Keyboard Shortcut: Alt+1
  
 ' If ActiveCell.Count = 1 And ActiveCell.Formula <> "" Then ActiveSheet.name = ActiveCell.Formula
 If ActiveCell.Count = 1 And ActiveCell.Formula <> "" Then
   Dim reps, a, b, tgt As String, ttl As String
   ttl = ActiveCell.Value
   reps = Array(Array("", Array("*", "/", "\", "|", "&", "?", "？")))
   For Each a In reps
     tgt = a(0)
     For Each b In a(1)
       ttl = Replace(ttl, b, tgt)
     Next
   Next
       ActiveSheet.name = Left(trim(ttl), 31)
 End If
End Sub


Sub 列を順番にゴールシーク()
  Dim col As String
  Dim r As Range
  Dim cng As Range
  Dim tgt As Double
  Dim original As String
  Dim tmp As String
  
  Set r = ActiveCell
  col = InputBox("変化させるセルの入っている列を指定", Default:=Chr(Asc("A") - 1 + r.Offset(0, -1).Column))
  If Len(col) = 0 Then Exit Sub
  
  tmp = InputBox("目標値", Default:=0#)
  If Len(tmp) = 0 Then Exit Sub
  tgt = val(tmp)
  Do While Len(r.FormulaLocal) > 0
    Set cng = Range(col & Format(r.row, "0"))
    original = cng.FormulaLocal
    If Not r.GoalSeek(tgt, cng) Then
      cng.Formula = original
      cng.Font.color = vbRed
    End If
    Set r = r.Offset(1)
  Loop

End Sub


Sub 複数のCSVを現在のブックの末尾に読み込み()
    Dim csv As Workbook
    Dim book As Workbook
    Set book = ActiveWorkbook
    
    Dim targets
    targets = Application.GetOpenFilename( _
      FileFilter:="CSV files,*.csv,AllFiles(*.*),*.*", _
      Title:="対象となるcsvを選択", MultiSelect:=True)
    
    If IsArray(targets) Then
    
      Dim csvfile
      Dim sh As Worksheet
      For Each csvfile In targets
        Workbooks.OpenText csvfile, DataType:=xlDelimited, Comma:=True
        Set csv = ActiveWorkbook
        Set sh = csv.Worksheets(1)
        sh.Move After:=book.Worksheets(book.Worksheets.Count)
      Next
    End If
End Sub


Sub シート名を抽出()
  Dim b As Workbook
  Set b = ActiveWorkbook
  Dim n
  Dim r As Range
  Dim s  As Worksheet
  Set r = ActiveCell
  
  On Error GoTo kkk:
  Application.ScreenUpdating = False
  
  For Each n In b.Worksheets
    Set s = n
    r.Formula = s.name
    Set r = r.Offset(1)
  Next n
kkk:
  Application.ScreenUpdating = True
  
End Sub



