VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} 参照先選択 
   Caption         =   "参照先選択"
   ClientHeight    =   10140
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   15540
   OleObjectBlob   =   "参照先選択.frx":0000
   StartUpPosition =   1  'オーナー フォームの中央
End
Attribute VB_Name = "参照先選択"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Const リストの番号部の幅 As Single = 60  '32
Private Const リストのブックマーク名の幅 As Single = 150
Private Const 内容表示文字数 As Integer = 100

Private Const posTag As Integer = 0
Private Const posName As Integer = 1
Private Const posStart As Integer = 3
Private Const posValue As Integer = 2  ' 追加

Private continue_flag As Boolean
Private names()

Private Sub CommandButton1_Click()
 Apply
 Me.ListBox1.SetFocus
End Sub

Private Sub CommandButton2_Click()
 Unload Me
End Sub

' 相互参照を入力する
Private Sub Apply(Optional ByVal 図表参照スタイル設定 As Boolean = True)
    Dim i As Integer
    Dim a As Boolean, b As Boolean
    Dim tag As String, head As String
    i = ListBox1.ListIndex
    If i >= 0 Then
        If names(posTag, i) = "" Then
            'Selection.InsertCrossReference ReferenceType:="ブックマーク", ReferenceKind:= _
                wdPageNumber, ReferenceItem:=names(1, i), InsertAsHyperlink:=True, _
                IncludePosition:=False, SeparateNumbers:=False, SeparatorString:=" "
            Selection.InsertCrossReference ReferenceType:="ブックマーク", _
                ReferenceKind:=wdContentText, _
                ReferenceItem:=names(posName, i), _
                InsertAsHyperlink:=False, _
                IncludePosition:=False, _
                SeparateNumbers:=False, _
                SeparatorString:=" "
        Else
            Dim fld As Field
            
            Selection.InsertCrossReference ReferenceType:="ブックマーク", ReferenceKind:= _
                wdNumberNoContext, ReferenceItem:=names(posName, i), _
                InsertAsHyperlink:=True, _
                IncludePosition:=False, SeparateNumbers:=False, SeparatorString:=" "
            
            tag = names(posTag, i)
            head = Mid(tag, 1, 1)
            ' 文献引用の確認
            a = Right(tag, 1) = ")"
            b = InStr("123456789", head) > 0
            If a And b Then
                ' 右肩カッコ形式の参考文献引用 (土論など)
                Selection.MoveLeft Unit:=wdCharacter, Count:=1, Extend:=wdExtend
                Selection.Font.Superscript = True
            ElseIf head = "図" Or head = "表" Then
                If 図表参照スタイル設定 Then
                  Call 図表参照確認
                   Selection.MoveLeft Unit:=wdCharacter, Count:=1, Extend:=wdExtend
                   Selection.Style = "図表参照"
                   Selection.Characters.Last.Font.Reset
                   Selection.InsertAfter " "
                   Selection.InsertBefore "QUOTE "
                   Selection.Characters.First.Font.Reset
                   Selection.Fields.Add(Selection.Range, wdFieldEmpty, , False).Select
                   Selection.Fields.Update
                End If
            End If
            Selection.Collapse wdCollapseEnd
        End If
    End If
    


End Sub

Private Sub ApplyName()
    Dim i As Integer
    Dim a As Boolean, b As Boolean
    Dim tag As String, head As String
    i = ListBox1.ListIndex
    If i >= 0 Then
        If names(posTag, i) = "" Then
            Selection.InsertCrossReference ReferenceType:="ブックマーク", ReferenceKind:= _
                wdPageNumber, ReferenceItem:=names(posName, i), InsertAsHyperlink:=True, _
                IncludePosition:=False, SeparateNumbers:=False, SeparatorString:=" "
        Else
            Selection.InsertCrossReference ReferenceType:="ブックマーク", _
            ReferenceKind:=wdContentText, ReferenceItem:=names(posName, i), _
                InsertAsHyperlink:=True, _
                IncludePosition:=False, SeparateNumbers:=False, SeparatorString:=" "
        End If
    End If

End Sub


Private Sub ListBox1_DblClick(ByVal Cancel As MSForms.ReturnBoolean)
  Apply
End Sub

Private Sub 図表参照確認()
    Dim X As Style
    For Each X In ActiveDocument.Styles
      If X.NameLocal = "図表参照" Then Exit Sub
    Next X


    'Normal.dotの図表参照スタイルがデフォルトのスタイルになります．
    'ここでエラーがでて止まった場合は，Normal.dotに"図表参照"という文字スタイルを追加してください．
    Application.OrganizerCopy _
        Source:=ThisDocument.Path & "\" & ThisDocument.name, _
        Destination:=ActiveDocument.Path & "\" & ActiveDocument.name, _
        name:="図表参照", Object:=wdOrganizerObjectStyles

'        "D:\Documents\NSC999\Application Data\Microsoft\Templates\報告書用re.dot", _


End Sub

Private Sub ブックマーク名を入力()

    Selection.text = names(posName, ListBox1.ListIndex)
    Selection.Collapse wdCollapseEnd

End Sub


Private Sub ListBox1_KeyDown(ByVal KeyCode As MSForms.ReturnInteger, ByVal Shift As Integer)
    
    Select Case KeyCode
    Case vbKeyEscape
       Unload Me
    Case vbKeyReturn
        ' Ctrl+Shiftによる連続記入の場合は改段落しておくほうが使いやすい
        If (Shift And 2) <> 0 Then
          If continue_flag Then 改段落記入
        End If
        If (Shift And 4) = 0 Then 'Altの場合は入力しない
            Call Apply((Shift And 2) = 0)
        End If
        If (Shift And 2) + (Shift And 4) <> 0 Then 'ctrl or Alt
          If (Shift And 2) <> 0 Then 半角スペース記入
          ApplyName
        End If
        If (Shift And 1) = 0 Then Unload Me 'シフトが押されていない場合
    Case vbKeyDelete, vbKeyBack
        ' Remove bookmark
        'ブックマークを削除
        ActiveDocument.Bookmarks(names(posName, ListBox1.ListIndex)).Delete
        'リストから削除
        ListBox1.RemoveItem ListBox1.ListIndex
    Case vbKeyF2
        ' rename
        Dim bk As Bookmark
        Dim n
        n = ListBox1.ListIndex
        Set bk = ActiveDocument.Bookmarks(names(posName, n))
        Set bk = ブックマークの置換(bk)
        names(posName, n) = bk.name
        ListBox1.List(n, posName) = bk.name
       
    Case vbKeySpace, 229
        If (Shift And 2) <> 0 Then
          ブックマーク名を入力
        Else
            With ListBox1
                If Shift = 1 And .ListIndex > 0 Then
                  .ListIndex = .ListIndex - 1
                ElseIf .ListIndex <> .ListCount - 1 Then
                  .ListIndex = .ListIndex + 1
                End If
            End With
        End If
    End Select
    continue_flag = True
End Sub

Private Function 新ブックマーク名の取得(old_name As String) As String
    
    Dim e As New ブックマーク名編集
    e.候補 = old_name
    e.Show vbModal
    新ブックマーク名の取得 = e.結果
    Unload e
    Set e = Nothing
    
End Function

Private Function 新ブックマーク名の取得_OLD(old_name As String) As String
    Dim typed_name As String, new_name As String
    Dim res As VbMsgBoxResult
    
    typed_name = old_name
    新ブックマーク名の取得_OLD = ""  ' キャンセルした場合
    
    Do
    
      typed_name = InputBox("新しいブックマーク名を入力してください．" & vbCrLf & _
                          "旧：" & old_name, "ブックマークの変更", typed_name)
      
      new_name = ブックマーク可能文字への変換(typed_name)
      If new_name = typed_name Then
        res = vbYes
      Else
        res = MsgBox("ブックマークに使えない文字列があったので変更しました．" & vbCrLf & _
                    "変更前：""" & typed_name & """" & vbCrLf & _
                    "変更後：""" & new_name & """" & vbCrLf & _
                    "よろしいですか？ " & vbCrLf & _
                    "   はい: 変更後のもので置き換え" & vbCrLf & _
                    "   いいえ: 文字列を再修正する" & vbCrLf & _
                    "   キャンセル： ブックマーク名修正を取り消し", vbYesNoCancel, _
                    "ブックマーク自動修正の確認")
      End If
      If res = vbCancel Then Exit Function
        
    Loop Until res = vbYes
    
    新ブックマーク名の取得_OLD = new_name

End Function


Private Function ブックマークの置換(ByRef bk As Bookmark) As Bookmark
    Dim new_name As String, old_name As String
    old_name = bk.name
    new_name = 新ブックマーク名の取得(old_name)
    
    ' キャンセルのチェック
    If Len(new_name) = 0 Then Exit Function
    
    ' 新しい名前で同じ位置にブックマークを追加
    Set ブックマークの置換 = ActiveDocument.Bookmarks.Add(new_name, bk.Range)
    
    'ブックマーク参照の置き換え
    Dim f As Field
    For Each f In ActiveDocument.Fields
      If f.Type = wdFieldRef Then
        f.Code.text = Replace(f.Code.text, old_name, new_name)
      End If
    Next
    
    ' 不要になったブックマークを削除する．
    bk.Delete
    

End Function


Private Sub 説明文記入()

  Me.DescMain.Caption = _
  "Enter:ブックマーク先の番号を引用(してダイアログを閉じる)" & vbCrLf & _
  "Ctrl+Enter：番号とブックマーク文字列を追加" & vbCrLf & _
  "Alt+Enter: ブックマーク文字列のみ" & vbCrLf & _
  "上に＋Shift：ダイアログは開いたまま" & vbCrLf & _
  "ダブルクリック： 番号を記入してダイアログは開いたまま" & vbCrLf & _
  "Tab: 右上のリストに移動" & vbCrLf & _
  "Esc: キャンセル(なにもせず閉じる)" & vbCrLf & _
  "Del: 選択しているブックマークを削除" & vbCrLf & _
  "F2: 選択しているブックマーク名を修正(参照先も変更される)"
  
  Me.DescSub.Caption = _
  "ブックマーク追加：" & vbCrLf & _
  "①上で種類を選択すると下が更新" & vbCrLf & _
  "②下で対象を選択してEnter" & vbCrLf & _
  "③ブックマーク名を入力" & vbCrLf & _
  "④参照先の番号が入力される" & vbCrLf & _
  ""
  

End Sub




Private Sub AddBookmark()
  Dim Data(0 To 4)
  Dim i As Integer, n As Integer
  
  With Me.ListBoxCaptions
    If IsNull(.Value) Then Exit Sub
    For i = 0 To 4
      .BoundColumn = i
      Data(i) = .Value
    Next
  End With
  Dim r As Range
  Set r = ActiveDocument.Range(Data(3), Data(4))
  If r.text = "" Then Exit Sub
  
  Dim tag As String
  tag = ブックマーク可能文字への変換(r.text)
  tag = InputBox("ブックマークを確認して修正してください" & vbCrLf & _
                 "元:" & r.text, _
                 "ブックマーク名の確認・修正", tag)
  If tag = "" Then Exit Sub
    
  ActiveDocument.Bookmarks.Add tag, r
  
  i = ListBox1.ListCount
  ListBox1.AddItem
  ListBox1.ListIndex = i
  ListBox1.List(i, 0) = Data(1)
  ListBox1.List(i, 1) = tag
  
   n = ActiveDocument.Bookmarks.Count
   ReDim Preserve names(posStart + 1, n - 1)
   names(posTag, n - 1) = Data(1)
   names(posName, n - 1) = tag
   names(posStart, n - 1) = Data(3)
   names(posValue, n - 1) = Left(r.text, 内容表示文字数)
  
   
End Sub

Private Sub AddBookmarkAndApply()
  
  AddBookmark
  Apply
  
End Sub

Private Sub ListBoxCaptions_DblClick(ByVal Cancel As MSForms.ReturnBoolean)
  AddBookmark ' AndApply
   Call ブックマークリスト更新
   ListBoxStyle_Change
End Sub



Private Sub ListBoxCaptions_KeyDown(ByVal KeyCode As MSForms.ReturnInteger, ByVal Shift As Integer)
    Select Case KeyCode
    Case vbKeyReturn
        AddBookmarkAndApply
        If Shift = 0 Then Unload Me
        Call ブックマークリスト更新
        ListBoxStyle_Change
    Case vbKeySpace, 229
        With ListBoxCaptions
            If Shift = 1 And .ListIndex > 0 Then
              .ListIndex = .ListIndex - 1
            ElseIf .ListIndex <> .ListCount - 1 Then
              .ListIndex = .ListIndex + 1
            End If
        End With
    End Select

End Sub

Private Sub ListBoxStyle_Change()
    'Me.ListBoxCaptions.Clear
    Dim key, i
    
    key = Me.ListBoxStyle.List(Me.ListBoxStyle.ListIndex)
    Dim p As Paragraph
    Dim c As Collection
    Set c = New Collection
    
    With Me.ListBoxCaptions
      .ColumnCount = 3
      .ColumnWidths = CStr(リストの番号部の幅) & ";" & CStr(.Width - リストの番号部の幅 - 5) & ";0"
    End With
    
    Dim item(0 To 3) As String
'    For i = 1 To ActiveDocument.Paragraphs.Count
'        Set p = ActiveDocument.Paragraphs(i)
'        If p.Style.NameLocal = key Then
'            item(0) = p.Range.ListFormat.ListString
'            item(1) = p.Range.Text
'            item(2) = i
'            Me.ListBoxCaptions.AddItem item
'        End If
'    Next
    
    On Error GoTo eee:
    Dim r As Range
    Dim next_pos
    Set r = ActiveDocument.Content
    r.Find.ClearFormatting
    r.Find.Style = ActiveDocument.Styles(key)
    Do While r.Find.Execute("", Forward:=True, format:=True, Wrap:=wdFindStop)
      item(0) = r.ListFormat.ListString
      item(1) = TrimEx(r.text)
      item(2) = r.Start + Len(r.text) - Len(LTrim(r.text))  ' ここはEx不要
      item(3) = r.End - Len(r.text) + Len(RTrimEx(r.text))  ' こっちはExが必要
      next_pos = r.End + 1
      If r.Bookmarks.Count = 0 Then
          c.Add item
      End If
      ' go to rest if r doesnot reach the end of the active document
      If r.End = ActiveDocument.Content.End Then Exit Do
      r.End = ActiveDocument.Content.End
      '  r.Start = item(3) + 1   <== これだと永久ループが発生する．
      r.Start = next_pos
    Loop
    
    
    If c.Count = 0 Then
        Me.ListBoxCaptions.Clear
        Exit Sub
    End If
    
    Dim Data()
    ReDim Data(c.Count() - 1, 3)
    Dim k As Integer
    For i = 0 To c.Count - 1
      For k = 0 To 3
        Data(i, k) = c(i + 1)(k)
      Next
    Next
    
    Me.ListBoxCaptions.List() = Data

eee:

End Sub

Private Sub ブックマークリスト更新()
   Dim n
   Dim pos()
   Dim arr As New ArrayList
   Dim i, k
   Dim tag As String
   Dim b As Bookmark
   Dim bs As Bookmarks
   
   Set bs = ActiveDocument.Bookmarks
   n = ActiveDocument.Bookmarks.Count
   If n > 0 Then
       ReDim names(posStart + 1, n - 1)  ' redim のために行と列を逆にする
       For i = 1 To n
         k = i - 1
         Set b = bs(i)
    '     pos(i - 1) = b.End
         arr.Add format(b.End, "0000000000") & "," & format(k, "0")
         DoEvents
       Next
       
       arr.Sort
       
       For k = 0 To n - 1
         i = CInt(Split(arr(k), ",")(1)) + 1
         Set b = ActiveDocument.Bookmarks(i)
         tag = b.Range.ListFormat.ListString
         names(posTag, k) = tag 'b.Range.ListFormat.ListString
         names(posName, k) = b.name
         names(posStart, k) = b.Start
         names(posValue, k) = Left(b.Range.text, 内容表示文字数)
         'names(k, 2) = (b.End - b.Start) < 10 And Right(tag, 1) = ")" And InStr("123456789", left(tag,1)) > 0
         DoEvents
       Next
    
       Set arr = Nothing
       Me.LabelNum.Left = Me.ListBox1.Left + 5
       Me.LabelBookMark.Left = Me.ListBox1.Left + 5 + リストの番号部の幅
       Me.LabelTarget.Left = Me.LabelBookMark.Left + リストのブックマーク名の幅
       With Me.ListBox1
            .ColumnCount = 3
            .ColumnWidths = CStr(リストの番号部の幅) & ";" & _
                            CStr(リストのブックマーク名の幅) & ";" & _
                            CStr(.Width - リストのブックマーク名の幅 - リストの番号部の幅 - 5)
            
            ' ここで値を設定
            .Column() = names
            .ColumnHeads = False
            .SetFocus
            
            For i = 0 To n - 1
              k = names(posStart, i)
              If k > Selection.Range.Start Then Exit For
             DoEvents
            Next
            
            If i >= .ListCount Then i = .ListCount - 1
            .ListIndex = i
       End With
   Else
     Me.ListBoxStyle.SetFocus
   End If

End Sub



Private Sub UserForm_Initialize()
   
   continue_flag = False
     
   Call 説明文記入
   
   Call ブックマークリスト更新
   
   
   '' 参照用スタイルを設定
   Dim can
'   can = Array("図", "図-", "図(章)", "図(節)", _
'                "表", "表-", "表(章)", "表(節)", _
'                "見出し 1", "見出し 2", "見出し 3", _
'                "Appendix 1", "Appendix 2")

   ' 図の派生や表の派生については 図や表の設定のためのコピー元として扱っているので，
   ' 表示させると余計ややこしいということに気がついたので，削除した．
   can = Array("図", "表", "付属資料", "巻末資料", "列挙", _
                "見出し 1", "見出し 2", "見出し 3", _
                "Appendix 1", "Appendix 2")
                ' "図副題" は数が多いと問題が発生した（※）ので削除
                ' ※：ビジーループになり，応答しなくなったため，本体ごと強制終了する羽目になった．
   Me.ListBoxStyle.List() = can
   
   ' 以下は使っているものだけを乗せることを考えたもの．
   ' よくよく考えてみると，結構使い勝手が悪いと思われたので，却下．
'   Dim used_style()
'  Dim n_para As Integer
'   n_para = ActiveDocument.Paragraphs.Count
'   ReDim used_style(1 To n_para)
'   For i = 1 To n_para
'     used_style(i) = ActiveDocument.Paragraphs(i).Style
'   Next
   
   
   
   ' 右下のリストを更新するために，ListBoxStyleの最初を選択しておく．
   Me.ListBoxStyle.ListIndex = 0
   
   Exit Sub
   
   '''' 以下はメモ
   Dim used
   Set used = New Scripting.Dictionary
   Dim s_name, s As Style, para As Paragraph
   Dim c As Collection
   For Each s In ActiveDocument.Styles
      Set c = New Collection
      used.Add s.NameLocal, c
   Next
   Dim i
   For i = 1 To ActiveDocument.Paragraphs.Count
     Set para = ActiveDocument.Paragraphs(i)
     s_name = para.Style.NameLocal
     used(para.Style.NameLocal).Add i
   Next
   
   For Each s_name In can
     If used(s_name).Count > 0 Then
       Me.ListBoxStyle.AddItem s_name
     End If
   Next
   
   
   
End Sub
