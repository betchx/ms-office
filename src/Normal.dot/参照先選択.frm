VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} �Q�Ɛ�I�� 
   Caption         =   "�Q�Ɛ�I��"
   ClientHeight    =   10140
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   15540
   OleObjectBlob   =   "�Q�Ɛ�I��.frx":0000
   StartUpPosition =   1  '�I�[�i�[ �t�H�[���̒���
End
Attribute VB_Name = "�Q�Ɛ�I��"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Const ���X�g�̔ԍ����̕� As Single = 60  '32
Private Const ���X�g�̃u�b�N�}�[�N���̕� As Single = 150
Private Const ���e�\�������� As Integer = 100

Private Const posTag As Integer = 0
Private Const posName As Integer = 1
Private Const posStart As Integer = 3
Private Const posValue As Integer = 2  ' �ǉ�

Private continue_flag As Boolean
Private names()

Private Sub CommandButton1_Click()
 Apply
 Me.ListBox1.SetFocus
End Sub

Private Sub CommandButton2_Click()
 Unload Me
End Sub

' ���ݎQ�Ƃ���͂���
Private Sub Apply(Optional ByVal �}�\�Q�ƃX�^�C���ݒ� As Boolean = True)
    Dim i As Integer
    Dim a As Boolean, b As Boolean
    Dim tag As String, head As String
    i = ListBox1.ListIndex
    If i >= 0 Then
        If names(posTag, i) = "" Then
            'Selection.InsertCrossReference ReferenceType:="�u�b�N�}�[�N", ReferenceKind:= _
                wdPageNumber, ReferenceItem:=names(1, i), InsertAsHyperlink:=True, _
                IncludePosition:=False, SeparateNumbers:=False, SeparatorString:=" "
            Selection.InsertCrossReference ReferenceType:="�u�b�N�}�[�N", _
                ReferenceKind:=wdContentText, _
                ReferenceItem:=names(posName, i), _
                InsertAsHyperlink:=False, _
                IncludePosition:=False, _
                SeparateNumbers:=False, _
                SeparatorString:=" "
        Else
            Dim fld As Field
            
            Selection.InsertCrossReference ReferenceType:="�u�b�N�}�[�N", ReferenceKind:= _
                wdNumberNoContext, ReferenceItem:=names(posName, i), _
                InsertAsHyperlink:=True, _
                IncludePosition:=False, SeparateNumbers:=False, SeparatorString:=" "
            
            tag = names(posTag, i)
            head = Mid(tag, 1, 1)
            ' �������p�̊m�F
            a = Right(tag, 1) = ")"
            b = InStr("123456789", head) > 0
            If a And b Then
                ' �E���J�b�R�`���̎Q�l�������p (�y�_�Ȃ�)
                Selection.MoveLeft Unit:=wdCharacter, Count:=1, Extend:=wdExtend
                Selection.Font.Superscript = True
            ElseIf head = "�}" Or head = "�\" Then
                If �}�\�Q�ƃX�^�C���ݒ� Then
                  Call �}�\�Q�Ɗm�F
                   Selection.MoveLeft Unit:=wdCharacter, Count:=1, Extend:=wdExtend
                   Selection.Style = "�}�\�Q��"
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
            Selection.InsertCrossReference ReferenceType:="�u�b�N�}�[�N", ReferenceKind:= _
                wdPageNumber, ReferenceItem:=names(posName, i), InsertAsHyperlink:=True, _
                IncludePosition:=False, SeparateNumbers:=False, SeparatorString:=" "
        Else
            Selection.InsertCrossReference ReferenceType:="�u�b�N�}�[�N", _
            ReferenceKind:=wdContentText, ReferenceItem:=names(posName, i), _
                InsertAsHyperlink:=True, _
                IncludePosition:=False, SeparateNumbers:=False, SeparatorString:=" "
        End If
    End If

End Sub


Private Sub ListBox1_DblClick(ByVal Cancel As MSForms.ReturnBoolean)
  Apply
End Sub

Private Sub �}�\�Q�Ɗm�F()
    Dim X As Style
    For Each X In ActiveDocument.Styles
      If X.NameLocal = "�}�\�Q��" Then Exit Sub
    Next X


    'Normal.dot�̐}�\�Q�ƃX�^�C�����f�t�H���g�̃X�^�C���ɂȂ�܂��D
    '�����ŃG���[���łĎ~�܂����ꍇ�́CNormal.dot��"�}�\�Q��"�Ƃ��������X�^�C����ǉ����Ă��������D
    Application.OrganizerCopy _
        Source:=ThisDocument.Path & "\" & ThisDocument.name, _
        Destination:=ActiveDocument.Path & "\" & ActiveDocument.name, _
        name:="�}�\�Q��", Object:=wdOrganizerObjectStyles

'        "D:\Documents\NSC999\Application Data\Microsoft\Templates\�񍐏��pre.dot", _


End Sub

Private Sub �u�b�N�}�[�N�������()

    Selection.text = names(posName, ListBox1.ListIndex)
    Selection.Collapse wdCollapseEnd

End Sub


Private Sub ListBox1_KeyDown(ByVal KeyCode As MSForms.ReturnInteger, ByVal Shift As Integer)
    
    Select Case KeyCode
    Case vbKeyEscape
       Unload Me
    Case vbKeyReturn
        ' Ctrl+Shift�ɂ��A���L���̏ꍇ�͉��i�����Ă����ق����g���₷��
        If (Shift And 2) <> 0 Then
          If continue_flag Then ���i���L��
        End If
        If (Shift And 4) = 0 Then 'Alt�̏ꍇ�͓��͂��Ȃ�
            Call Apply((Shift And 2) = 0)
        End If
        If (Shift And 2) + (Shift And 4) <> 0 Then 'ctrl or Alt
          If (Shift And 2) <> 0 Then ���p�X�y�[�X�L��
          ApplyName
        End If
        If (Shift And 1) = 0 Then Unload Me '�V�t�g��������Ă��Ȃ��ꍇ
    Case vbKeyDelete, vbKeyBack
        ' Remove bookmark
        '�u�b�N�}�[�N���폜
        ActiveDocument.Bookmarks(names(posName, ListBox1.ListIndex)).Delete
        '���X�g����폜
        ListBox1.RemoveItem ListBox1.ListIndex
    Case vbKeyF2
        ' rename
        Dim bk As Bookmark
        Dim n
        n = ListBox1.ListIndex
        Set bk = ActiveDocument.Bookmarks(names(posName, n))
        Set bk = �u�b�N�}�[�N�̒u��(bk)
        names(posName, n) = bk.name
        ListBox1.List(n, posName) = bk.name
       
    Case vbKeySpace, 229
        If (Shift And 2) <> 0 Then
          �u�b�N�}�[�N�������
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

Private Function �V�u�b�N�}�[�N���̎擾(old_name As String) As String
    
    Dim e As New �u�b�N�}�[�N���ҏW
    e.��� = old_name
    e.Show vbModal
    �V�u�b�N�}�[�N���̎擾 = e.����
    Unload e
    Set e = Nothing
    
End Function

Private Function �V�u�b�N�}�[�N���̎擾_OLD(old_name As String) As String
    Dim typed_name As String, new_name As String
    Dim res As VbMsgBoxResult
    
    typed_name = old_name
    �V�u�b�N�}�[�N���̎擾_OLD = ""  ' �L�����Z�������ꍇ
    
    Do
    
      typed_name = InputBox("�V�����u�b�N�}�[�N������͂��Ă��������D" & vbCrLf & _
                          "���F" & old_name, "�u�b�N�}�[�N�̕ύX", typed_name)
      
      new_name = �u�b�N�}�[�N�\�����ւ̕ϊ�(typed_name)
      If new_name = typed_name Then
        res = vbYes
      Else
        res = MsgBox("�u�b�N�}�[�N�Ɏg���Ȃ������񂪂������̂ŕύX���܂����D" & vbCrLf & _
                    "�ύX�O�F""" & typed_name & """" & vbCrLf & _
                    "�ύX��F""" & new_name & """" & vbCrLf & _
                    "��낵���ł����H " & vbCrLf & _
                    "   �͂�: �ύX��̂��̂Œu������" & vbCrLf & _
                    "   ������: ��������ďC������" & vbCrLf & _
                    "   �L�����Z���F �u�b�N�}�[�N���C����������", vbYesNoCancel, _
                    "�u�b�N�}�[�N�����C���̊m�F")
      End If
      If res = vbCancel Then Exit Function
        
    Loop Until res = vbYes
    
    �V�u�b�N�}�[�N���̎擾_OLD = new_name

End Function


Private Function �u�b�N�}�[�N�̒u��(ByRef bk As Bookmark) As Bookmark
    Dim new_name As String, old_name As String
    old_name = bk.name
    new_name = �V�u�b�N�}�[�N���̎擾(old_name)
    
    ' �L�����Z���̃`�F�b�N
    If Len(new_name) = 0 Then Exit Function
    
    ' �V�������O�œ����ʒu�Ƀu�b�N�}�[�N��ǉ�
    Set �u�b�N�}�[�N�̒u�� = ActiveDocument.Bookmarks.Add(new_name, bk.Range)
    
    '�u�b�N�}�[�N�Q�Ƃ̒u������
    Dim f As Field
    For Each f In ActiveDocument.Fields
      If f.Type = wdFieldRef Then
        f.Code.text = Replace(f.Code.text, old_name, new_name)
      End If
    Next
    
    ' �s�v�ɂȂ����u�b�N�}�[�N���폜����D
    bk.Delete
    

End Function


Private Sub �������L��()

  Me.DescMain.Caption = _
  "Enter:�u�b�N�}�[�N��̔ԍ������p(���ă_�C�A���O�����)" & vbCrLf & _
  "Ctrl+Enter�F�ԍ��ƃu�b�N�}�[�N�������ǉ�" & vbCrLf & _
  "Alt+Enter: �u�b�N�}�[�N������̂�" & vbCrLf & _
  "��Ɂ{Shift�F�_�C�A���O�͊J�����܂�" & vbCrLf & _
  "�_�u���N���b�N�F �ԍ����L�����ă_�C�A���O�͊J�����܂�" & vbCrLf & _
  "Tab: �E��̃��X�g�Ɉړ�" & vbCrLf & _
  "Esc: �L�����Z��(�Ȃɂ���������)" & vbCrLf & _
  "Del: �I�����Ă���u�b�N�}�[�N���폜" & vbCrLf & _
  "F2: �I�����Ă���u�b�N�}�[�N�����C��(�Q�Ɛ���ύX�����)"
  
  Me.DescSub.Caption = _
  "�u�b�N�}�[�N�ǉ��F" & vbCrLf & _
  "�@��Ŏ�ނ�I������Ɖ����X�V" & vbCrLf & _
  "�A���őΏۂ�I������Enter" & vbCrLf & _
  "�B�u�b�N�}�[�N�������" & vbCrLf & _
  "�C�Q�Ɛ�̔ԍ������͂����" & vbCrLf & _
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
  tag = �u�b�N�}�[�N�\�����ւ̕ϊ�(r.text)
  tag = InputBox("�u�b�N�}�[�N���m�F���ďC�����Ă�������" & vbCrLf & _
                 "��:" & r.text, _
                 "�u�b�N�}�[�N���̊m�F�E�C��", tag)
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
   names(posValue, n - 1) = Left(r.text, ���e�\��������)
  
   
End Sub

Private Sub AddBookmarkAndApply()
  
  AddBookmark
  Apply
  
End Sub

Private Sub ListBoxCaptions_DblClick(ByVal Cancel As MSForms.ReturnBoolean)
  AddBookmark ' AndApply
   Call �u�b�N�}�[�N���X�g�X�V
   ListBoxStyle_Change
End Sub



Private Sub ListBoxCaptions_KeyDown(ByVal KeyCode As MSForms.ReturnInteger, ByVal Shift As Integer)
    Select Case KeyCode
    Case vbKeyReturn
        AddBookmarkAndApply
        If Shift = 0 Then Unload Me
        Call �u�b�N�}�[�N���X�g�X�V
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
      .ColumnWidths = CStr(���X�g�̔ԍ����̕�) & ";" & CStr(.Width - ���X�g�̔ԍ����̕� - 5) & ";0"
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
      item(2) = r.Start + Len(r.text) - Len(LTrim(r.text))  ' ������Ex�s�v
      item(3) = r.End - Len(r.text) + Len(RTrimEx(r.text))  ' ��������Ex���K�v
      next_pos = r.End + 1
      If r.Bookmarks.Count = 0 Then
          c.Add item
      End If
      ' go to rest if r doesnot reach the end of the active document
      If r.End = ActiveDocument.Content.End Then Exit Do
      r.End = ActiveDocument.Content.End
      '  r.Start = item(3) + 1   <== ���ꂾ�Ɖi�v���[�v����������D
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

Private Sub �u�b�N�}�[�N���X�g�X�V()
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
       ReDim names(posStart + 1, n - 1)  ' redim �̂��߂ɍs�Ɨ���t�ɂ���
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
         names(posValue, k) = Left(b.Range.text, ���e�\��������)
         'names(k, 2) = (b.End - b.Start) < 10 And Right(tag, 1) = ")" And InStr("123456789", left(tag,1)) > 0
         DoEvents
       Next
    
       Set arr = Nothing
       Me.LabelNum.Left = Me.ListBox1.Left + 5
       Me.LabelBookMark.Left = Me.ListBox1.Left + 5 + ���X�g�̔ԍ����̕�
       Me.LabelTarget.Left = Me.LabelBookMark.Left + ���X�g�̃u�b�N�}�[�N���̕�
       With Me.ListBox1
            .ColumnCount = 3
            .ColumnWidths = CStr(���X�g�̔ԍ����̕�) & ";" & _
                            CStr(���X�g�̃u�b�N�}�[�N���̕�) & ";" & _
                            CStr(.Width - ���X�g�̃u�b�N�}�[�N���̕� - ���X�g�̔ԍ����̕� - 5)
            
            ' �����Œl��ݒ�
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
     
   Call �������L��
   
   Call �u�b�N�}�[�N���X�g�X�V
   
   
   '' �Q�Ɨp�X�^�C����ݒ�
   Dim can
'   can = Array("�}", "�}-", "�}(��)", "�}(��)", _
'                "�\", "�\-", "�\(��)", "�\(��)", _
'                "���o�� 1", "���o�� 2", "���o�� 3", _
'                "Appendix 1", "Appendix 2")

   ' �}�̔h����\�̔h���ɂ��Ă� �}��\�̐ݒ�̂��߂̃R�s�[���Ƃ��Ĉ����Ă���̂ŁC
   ' �\��������Ɨ]�v��₱�����Ƃ������ƂɋC�������̂ŁC�폜�����D
   can = Array("�}", "�\", "�t������", "��������", "��", _
                "���o�� 1", "���o�� 2", "���o�� 3", _
                "Appendix 1", "Appendix 2")
                ' "�}����" �͐��������Ɩ�肪���������i���j�̂ō폜
                ' ���F�r�W�[���[�v�ɂȂ�C�������Ȃ��Ȃ������߁C�{�̂��Ƌ����I������H�ڂɂȂ����D
   Me.ListBoxStyle.List() = can
   
   ' �ȉ��͎g���Ă�����̂������悹�邱�Ƃ��l�������́D
   ' �悭�悭�l���Ă݂�ƁC���\�g�����肪�����Ǝv��ꂽ�̂ŁC�p���D
'   Dim used_style()
'  Dim n_para As Integer
'   n_para = ActiveDocument.Paragraphs.Count
'   ReDim used_style(1 To n_para)
'   For i = 1 To n_para
'     used_style(i) = ActiveDocument.Paragraphs(i).Style
'   Next
   
   
   
   ' �E���̃��X�g���X�V���邽�߂ɁCListBoxStyle�̍ŏ���I�����Ă����D
   Me.ListBoxStyle.ListIndex = 0
   
   Exit Sub
   
   '''' �ȉ��̓���
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
