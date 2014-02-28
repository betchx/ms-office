Attribute VB_Name = "Sheet"
Option Explicit

Sub Auto_Open()

    'Ctrl+Shift+Q
    Application.OnKey "+^q", "上付き下付き"
    
    ' Ctrl+Shift+L
    Application.OnKey "+^l", "自動下付き化"
    
    ' Ctrl+Shift+U
    Application.OnKey "+^u", "自動上付き化"
    
    ' Ctrl+Shift+N
    Application.OnKey "+^n", "上付き下付き解除"
    
    ' Ctrl+Shift+H
    Application.OnKey "+^h", "自動上付き下付き順次除去"

        
    ' Ctrl+T
    Application.OnKey "^t", "カレントセルをシートタイトルに"
    
    
    ' Ctrl+Shift+E
    'Application.OnKey "+^E", "目盛を外に"
    Application.OnKey "+^E", "有効数字設定"

    ' Ctrl+Shift+_  (下線)
    Application.OnKey "+^\", "罫線クリア"
    
    ' Ctrl+Shift+R
    Application.OnKey "+^r", "Round追加"
        
    ' Ctrl+Shift+F
    Application.OnKey "+^f", "下罫線設定"
    
    ' Ctrl+Shift+B
    Application.OnKey "+^b", "選択範囲を太枠罫線で囲う"
    
    ' Ctrl+Shift+G
    Application.OnKey "+^g"
    Application.OnKey "+^g", "選択範囲に格子罫線を設定" ' <== "グリッド化"
    
    Dim i As Integer
    For i = 0 To 6
    ' Ctrl+Alt+0
      Application.OnKey "^%" & i, "NumFormatPoint" & i
    Next i
    
    ' Ctrl+Alt+7
    ' 先頭に ( を追加する．
    Application.OnKey "%7", "InsertOpenBracket"
    
    ' Ctrl+Alt+8
    ' 末尾に ( を追加する
    Application.OnKey "%8", "AppendOpenBracket"
    
    ' Ctrl+Alt+9
    ' 末尾に ) を追加する．
    Application.OnKey "%9", "AppendCloseBracket"
    
    
    ' Ctrl+Alt+U
    Application.OnKey "^%u", "単位追加"
    
    ' Ctrl+Alt+s
    Application.OnKey "%s", "シート名の選択記入"
    
    ' ×記入： Alt+:
    Application.OnKey "%:", "SetTimes"
    
    '縮小して表示： Alt+;
    Application.OnKey "%;", "縮小して表示"
       
    ' シートに名前を追加： Ctrl+Alt+N
    Application.OnKey "^%N", "シートに名前を追加"
    
    
    ' 不要行の隠蔽 :  Ctrl+Alt+H
    Application.OnKey "%^h", "列を指定して不要行を隠蔽"
    
    
    ' セル入力後の移動の有無を切り替え :  Ctrl+M
    Application.OnKey "^m", "セル入力後の移動の有無を切り替え"
    
    ' セル入力後の移動の方向を切り替え :  Ctrl+J
    Application.OnKey "^j", "セル入力後の移動の方向を切り替え"
    

End Sub


Sub formatAsTex()
    
    Dim rng As Range
    
    For Each rng In Selection
      If rng.Characters.Count > 0 Then
        Dim pos
        pos = 1
        Dim ch As Characters
        Do While pos <= rng.Characters.Count
           Select Case rng.Characters(pos, 1)
           Case "_" ' 下付き
             rng.Characters(pos, 1).Delete
             
             rng.Characters(pos, 1).Font.Subscript = True
           
           Case "^" '上付き
             rng.Characters(pos, 1).Delete
             
             rng.Characters(pos, 1).Font.Superscript = True
            
           End Select
          
        Loop
      
      End If
    Next

End Sub


Sub Round追加()
    Dim rng As Range
    Dim n As Integer
    Dim eqn As String
    Dim fmt As String
    Dim add_round As Boolean
    For Each rng In Selection
      eqn = rng.Formula
      
      ' 行うべき操作の種類を判定
      If Len(eqn) > 0 And Left(eqn, 1) = "=" Then
        'すでにRoundが追加されていないかチェックする．
        If Left(eqn, 6) = "=ROUND" Then
          add_round = False
        Else
          add_round = True
        End If
      Else
        '数式以外
        add_round = True
      End If
      
      ' Roundの操作を実施
      If add_round Then
        ' add Round
        fmt = get_format(rng)
        If Left(fmt, 2) = "0." Then
           n = Len(fmt) - 2
        Else
           Dim s As String
           s = InputBox("Roundの桁数を指定.", "桁数設定")
           If s = "" Then
             n = 0
           Else
             n = CInt(val(s))
           End If
        End If
        
        ' extract body
        eqn = Mid(eqn, 2)
        
        rng.Formula = "=Round(" & eqn & Format(n, ", 0)")
      Else
        ' Roundあり
        Select Case Mid(eqn, 7, 1)
        Case "("
            '通常のRound   ==> RounUpに変更
            rng.Formula = "=ROUNDUP" & Mid(eqn, 7)
        Case "U"
            ' RoundUp  ==> RoundDownに変更
            rng.Formula = "=RoundDown" & Mid(eqn, 9)
        Case "D"
            ' RoundDown ==> Ronudなしに
            n = Len(eqn) - 1
            Do While n > 0 And Mid(eqn, n, 1) <> ","
              n = n - 1
            Loop
            If n = 0 Then Exit Sub
            rng.Formula = "=" & Mid(eqn, 12, n - 12)
        End Select
      End If
    Next
End Sub

Private Sub Round追加_Original()
    Dim rng As Range
    Dim n As Integer
    Dim eqn As String
    Dim fmt As String
    For Each rng In Selection
      eqn = rng.Formula
      If Len(eqn) > 0 And Left(eqn, 1) = "=" Then
         fmt = get_format(rng)
         If Left(fmt, 2) = "0." Then
            n = Len(fmt) - 2
         Else
            Dim s As String
            s = InputBox("Roundの桁数を指定.", "桁数設定")
            If s = "" Then
              n = 0
            Else
              n = CInt(val(s))
            End If
         End If
         
         ' extract body
         eqn = Mid(eqn, 2)
         
         rng.Formula = "=Round(" & eqn & Format(n, ", 0)")
         
      End If
    Next
End Sub

Sub 有効数字設定()
    Dim rng As Range
    Dim n As Integer
    Dim eqn As String
    
    Dim ans
    ans = InputBox("有効桁数", Default:="3")
    n = CInt(val(ans))
    If n < 1 Then
       MsgBox "不正な有効桁数です．"
       Exit Sub
    End If
    
    For Each rng In Selection
      eqn = rng.Formula
      If Len(eqn) > 0 Then
        ' extract body
        If Left(eqn, 1) = "=" Then eqn = Mid(eqn, 2)
        Dim new_eqn
        
        new_eqn = "=if((" & eqn & ")=0,0,Round(" & eqn & ", " & Format(n, "0") & " - roundup(log10(abs(" & eqn & ")),0)))"
        rng.Formula = new_eqn
      End If
    Next
End Sub

Sub 下罫線設定()

' 罫線のクリアは行わない様にした．

   Call 選択範囲に罫線を設定(Array(xlEdgeBottom))
   
   Exit Sub

'  以下はオリジナルのマクロ
    Selection.Borders(xlDiagonalDown).LineStyle = xlNone
    Selection.Borders(xlDiagonalUp).LineStyle = xlNone
    Selection.Borders(xlEdgeLeft).LineStyle = xlNone
    Selection.Borders(xlEdgeTop).LineStyle = xlNone
    With Selection.Borders(xlEdgeBottom)
        .LineStyle = xlContinuous
        .weight = xlThin
        .ColorIndex = xlAutomatic
    End With
    Selection.Borders(xlEdgeRight).LineStyle = xlNone
End Sub



Sub 罫線クリア()
    Dim r As Range
    Set r = Selection
    r.Borders.LineStyle = xlLineStyleNone
End Sub

Private Sub 選択範囲に罫線を設定(targets, _
   Optional style As XlLineStyle = xlContinuous, _
   Optional weight As XlBorderWeight = xlThin, _
   Optional color = xlAutomatic)
    Dim tgt
    For Each tgt In targets
      With Selection.Borders(tgt)
        ' 罫線が存在しない場合はweightが225になる模様
        If .weight <> 225 Then
          .LineStyle = style
          .weight = weight
          .ColorIndex = color
        End If
      End With
    Next
   End Sub

Sub 選択範囲を太枠罫線で囲う()
  Call 選択範囲に罫線を設定( _
    Array(xlEdgeLeft, xlEdgeTop, xlEdgeBottom, xlEdgeRight) _
    , weight:=xlMedium)
End Sub


Sub 選択範囲に格子罫線を設定()
  Call 選択範囲に罫線を設定( _
    Array(xlEdgeLeft, xlEdgeTop, xlEdgeBottom, xlEdgeRight, _
          xlInsideVertical, xlInsideHorizontal) _
     )
End Sub



Sub 自動下付き化()

   Dim b As Integer
   
   b = Len(ActiveCell.Text)
   Do While ActiveCell.Characters(b, 1).Font.Subscript
        b = b - 1
        If b <= 0 Then Exit Sub
   Loop
   ActiveCell.Characters(b, 1).Font.Subscript = True

End Sub

Sub 自動上付き化()

   Dim b As Integer
   
   b = Len(ActiveCell.Text)
   Do While ActiveCell.Characters(b, 1).Font.Superscript
        b = b - 1
   Loop
   ActiveCell.Characters(b, 1).Font.Superscript = True

End Sub

Sub 自動上付き下付き順次除去()

   Dim b As Integer
   
   b = Len(ActiveCell.Text)
   Do Until ActiveCell.Characters(b, 1).Font.Superscript _
       Or ActiveCell.Characters(b, 1).Font.Subscript
        b = b - 1
        If b < 0 Then Exit Sub
   Loop
   ActiveCell.Characters(b, 1).Font.Superscript = False
   ActiveCell.Characters(b, 1).Font.Subscript = False

End Sub


Sub 上付き下付き解除()
    ActiveCell.Characters.Font.Subscript = False
    ActiveCell.Characters.Font.Superscript = False
End Sub



Sub set_format(fmt As String, Optional r As Range = Nothing)
  If r Is Nothing Then Set r = Selection
  r.NumberFormatLocal = fmt
End Sub

Function get_format(Optional ByRef r As Range = Nothing) As String
'   Dim r As Range
   If r Is Nothing Then Set r = Selection
   If IsNull(r.NumberFormat) Then
     Dim s As String
     If r.Count > 1 Then
       s = CStr(r(1).Value)
     Else
       s = r.Value
     End If
     
     Dim fmt As String
     Dim t As String
     
     Dim u As Variant
     For Each u In Array("1", "2", "3", "4", "5", "6", "7", "8", "9")
        s = Replace(s, u, "0")
     Next
     get_format = Replace(s, "-", "")
   Else
     get_format = Selection.NumberFormatLocal
   End If
End Function

Sub NumFormatPoint0()
  set_format "0"
End Sub

Sub NumFormatPoint1()
  set_format "0.0"
End Sub

Sub NumFormatPoint2()
  set_format "0.00"
End Sub
Sub NumFormatPoint3()
  set_format "0.000"
End Sub
Sub NumFormatPoint4()
  set_format "0.0000"
End Sub
Sub NumFormatPoint5()
  set_format "0.00000"
End Sub
Sub NumFormatPoint6()
  set_format "0.000000"
End Sub


Sub AppendCloseBracket()
   set_format get_format() & ")"
End Sub

Sub AppendOpenBracket()
   set_format get_format() & "("
End Sub

Sub AddBracket()
   set_format "(" & get_format() & ")"
End Sub

Sub InsertOpenBracket()
   set_format "(" & get_format()
End Sub

Sub 単位追加()
  Dim unit As String
  unit = InputBox("単位を記入してください．数値との間には半角スペースが挿入されます")
  
  Dim r As Range
  For Each r In Selection
       set_format get_format(r) & " """ & unit & """", r
  Next

End Sub


Sub 上付き下付き()
  Dim f As Font
  Set f = Selection.Font
  With f
    If .Subscript Then
        .Subscript = False
        Exit Sub
    ElseIf .Superscript Then
        .Superscript = False
        .Subscript = True
    Else
        .Superscript = True
    End If
    
  End With
  If TypeName(f.Parent) = "Range" Then
    Dim r As Range
    Set r = f.Parent
    r.HorizontalAlignment = xlLeft
    r.Offset(0, -1).HorizontalAlignment = xlRight
  End If
End Sub

Sub SetTimes()
    ActiveCell.FormulaR1C1 = "×"
End Sub


Sub ☆ジャンプ()
  JumpList.Show vbModeless
End Sub


Sub シートに名前を追加()
    On Error GoTo eee:
    Dim r As Range
    Set r = Selection
    Dim s As String
    s = InputBox("名前を記入", "シートに名前を設定", "")
    If s = "" Then Exit Sub
    Dim sh As Worksheet
    Set sh = ActiveSheet
    
    If r.MergeCells Then
      ' セルがマージされている
      If r.Cells(1, 1).MergeArea.Count = r.Count Then
        ' 単一の結合セルが選択されている
        sh.Names.Add s, r.Cells(1, 1), True
        Exit Sub
      End If
    End If
    sh.Names.Add s, r, True
    
eee:

End Sub


Sub 指定のシートに名前を追加()
  Dim tgt As Range
  Set tgt = ActiveCell.CurrentRegion
  
  Dim row As Integer
  For row = 1 To tgt.Count / 3
    Dim sheet_name As String
    sheet_name = tgt(row, 1).Value
    If sheet_name = "" Then Exit Sub
    Dim addr As String
    addr = tgt(row, 2).Value
    Dim 名前 As String
    名前 = tgt(row, 3)
    Dim sh As Worksheet
    Set sh = ActiveWorkbook.Sheets(sheet_name)
    sh.Names.Add 名前, sh.Range(addr)
  Next


End Sub


Sub 縮小して表示()
'
' 縮小して表示 Macro
' マクロ記録日 : 2012/7/26  ユーザー名 : -
'
    Selection.ShrinkToFit = True
End Sub


Sub シート名の選択記入()
  Dim lister As New SheetNameLister
  Dim a As Worksheet
  Set a = ActiveSheet
  lister.Show vbModal
  If lister.selected_name <> "" Then
    ActiveCell.Formula = lister.selected_name
    SheetNameLister.selected_name = ""
  End If
  Unload lister
  
  a.Activate

End Sub


Private Sub 参照先が無ければ名前を削除(ByRef n As name)
      If Right(n.RefersTo, 5) = "#REF!" Or Mid(n.RefersTo, 2, 5) = "#REF!" Then
        n.Delete
      End If
End Sub


Sub 現在のブックから参照先の無い名前を全て削除()
   Dim シート As Worksheet
   Dim n As name
   Dim tb As shape
   Dim s As Worksheet
   Set s = ActiveSheet
   Set tb = s.Shapes.AddTextbox(msoTextOrientationHorizontal, ActiveCell.Left, ActiveCell.Top, 100, 20)
   tb.TextFrame.Characters().Text = "Workbook"
   tb.Visible = msoTrue
   DoEvents
   For Each n In ActiveWorkbook.Names
      参照先が無ければ名前を削除 n
   Next
   For Each シート In ActiveWorkbook.Sheets
      tb.TextFrame.Characters().Text = シート.name
      DoEvents
      For Each n In シート.Names
          参照先が無ければ名前を削除 n
      Next
   Next
   
   tb.Delete

End Sub


Sub 現在のシートから参照先のない名前を削除()
   Dim n As name
   For Each n In ActiveSheet.Names
      参照先が無ければ名前を削除 n
   Next
End Sub

Sub SetLink()

Dim i As Integer
Dim s As Worksheet, s1 As Worksheet
Dim r As Range
Set s1 = ActiveWorkbook.Sheets("一覧")
Set s = ActiveWorkbook.Sheets("一覧")


For i = 1 To 20
    Set r = s.Cells(i + 3, 2)
    r.Hyperlinks.Add r, "Soft" & Format(i, "00") & "!B3"
Next
End Sub


Sub 列を指定して不要行を隠蔽()
  Dim s As String
  s = InputBox("隠蔽指定列を指定してください", "隠蔽チェック列指定", "A")
  If s <> "" Then
     不要行の隠蔽 s
  End If
End Sub



Sub 不要行の隠蔽(Optional col As String = "A")
  Dim s As Worksheet
  Dim r As Range
  Dim origin As Range
  Dim sel As Range
  Dim target As Range
  Dim start_row As Integer
  Dim last_row As Integer
  Dim check_range As String
  
  Set s = ActiveSheet
  Set origin = ActiveCell
  Set sel = Selection
  
  Select Case s.PageSetup.PrintArea
  Case "", False
    check_range = col & ":" & col
  Case Else
    Dim print_range As Range
    Set print_range = s.Range(s.PageSetup.PrintArea)
    start_row = print_range.row
    last_row = start_row + print_range.Rows.Count - 1
    check_range = col & start_row & ":" & col & last_row
  End Select
  
  On Error GoTo xxx:
  Application.ScreenUpdating = False
  
  For Each r In s.Range(check_range)
    If r.EntireRow.Hidden Then
      If r.Value = "" Then r.EntireRow.Hidden = False
    Else
      If r.Value <> "" Then r.EntireRow.Hidden = True
    End If
  Next r
  
xxx:
  
  origin.Activate
  sel.Select
  Application.ScreenUpdating = True
  
  
End Sub


Sub 格子()
'
' 格子 Macro
' マクロ記録日 : 2013/10/9  ユーザー名 : NSC999
'

'
    Selection.Borders(xlDiagonalDown).LineStyle = xlNone
    Selection.Borders(xlDiagonalUp).LineStyle = xlNone
    With Selection.Borders(xlEdgeLeft)
        .LineStyle = xlContinuous
        .weight = xlThin
        .ColorIndex = xlAutomatic
    End With
    With Selection.Borders(xlEdgeTop)
        .LineStyle = xlContinuous
        .weight = xlThin
        .ColorIndex = xlAutomatic
    End With
    With Selection.Borders(xlEdgeBottom)
        .LineStyle = xlContinuous
        .weight = xlThin
        .ColorIndex = xlAutomatic
    End With
    With Selection.Borders(xlEdgeRight)
        .LineStyle = xlContinuous
        .weight = xlThin
        .ColorIndex = xlAutomatic
    End With
    With Selection.Borders(xlInsideVertical)
        .LineStyle = xlContinuous
        .weight = xlThin
        .ColorIndex = xlAutomatic
    End With
    With Selection.Borders(xlInsideHorizontal)
        .LineStyle = xlContinuous
        .weight = xlThin
        .ColorIndex = xlAutomatic
    End With
End Sub


Sub セル入力後の移動の有無を切り替え()
'
   
    Application.MoveAfterReturn = Not Application.MoveAfterReturn
End Sub

Sub セル入力後の移動の方向を切り替え()
    Select Case Application.MoveAfterReturnDirection
    Case xlToRight
      Application.MoveAfterReturnDirection = xlDown
    Case xlDown
      Application.MoveAfterReturnDirection = xlToRight
    Case xlUp
      Application.MoveAfterReturnDirection = xlToLeft
      Application.MoveAfterReturnDirection = xlDown
    Case xlToLeft
      Application.MoveAfterReturnDirection = xlUp
      Application.MoveAfterReturnDirection = xlDown
    End Select
End Sub
'    Application.MoveAfterReturnDirection = xlDown
'    Application.MoveAfterReturnDirection = xlUp
'    Application.MoveAfterReturnDirection = xlToLeft
'    Application.MoveAfterReturnDirection = xlDown
