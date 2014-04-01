Attribute VB_Name = "�t�B�[���h"
Option Explicit

Function ExcelEval(s As String)
  'Excel VBA��Evaluate�֐����Ăяo��
  ExcelEval = ExcelApplication.Evaluate(s)
End Function

Private Function ExcelApplication() As Excel.Application
' �G�N�Z����Application�I�u�W�F�N�g�ւ̎Q�Ƃ�Ԃ��D
' �G�N�Z�����܂��N������Ă��Ȃ��ꍇ�͋N������D

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
' �t�B�[���h�X�V�̃t�b�N
'
' UpdateFields Macro
' �I�������t�B�[���h�̎��s���ʂ��X�V���ĕ\�����܂��B
'
'  Expression (=) �t�B�[���h���g�����C�P���]���ŃG���[���ł��ꍇ��Excel.Evaluate�ɂ��]�������݂�
'  �܂��CExpression �t�B�[���h��2�����ڂ�=�̏ꍇ �i{ == xxxx }�̌`���̏ꍇ�j�́C�ŏ�����Excel�ŕ]������D

  Dim f As Field
  Dim i As Long
  Dim s As String
  Dim res
  '
  If Selection.Start = Selection.End Then
    For Each f In ActiveDocument.Fields
      If f.Code.Start <= Selection.Start And f.Code.End >= Selection.End Then
        UpdateFieldWithExpressionCheck f
        Exit For ' �l�X�g���������̃t�B�[���h�͕ʓr�X�V�����̂ŁC�����ł͍X�V�s�v
      End If
    Next
  Else
    For Each f In Selection.Fields
      UpdateFieldWithExpressionCheck f
    Next f
  End If
End Sub



Private Sub UpdateFieldWithExpressionCheck(ByRef f As Field)
' �������ǂ������m�F���C�t�B�[���h���X�V����D
' �����ŁC==��2�A������ꍇ�́C�G�N�Z���ŕ]������D

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
' �����t�B�[���h���G�N�Z����Evaluate�ōX�V����D
' �u�b�N�}�[�N�i�ϐ��j��X�C�b�`���ɑΉ����Ă��Ȃ��v���g�^�C�v�D
' (���łɖ��g�p�����C�킩��₢�̂Ń����Ƃ��Ďc���Ă����j
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
    f.Result.text = "Excel�Ή������ɃG���[������܂��u" & expr & "�v"
  End If
End Sub

Private Function isNumber(expr As String) As Boolean
' �����������Ŏn�܂邩�ǂ����̔���D
' ���̂Ƃ��뗘�p���Ă��Ȃ��D
  
  Dim Code As Integer
  Code = Asc(Left(expr, 1))
  isNumber = Code >= Asc("0") And Code <= Asc("9")
End Function

Private Sub UpdateFieldWithExcelEvalEx(ByRef f As Field)
' �����t�B�[���h���G�N�Z���̐����Ƃ��ĕ]������
' �ϐ��i�u�b�N�}�[�N�j��X�C�b�`�ɑΉ������o�[�W����
' �������C���[�h�����\�̃Z���ɑ΂���Q�Ƌ@�\�͖����Ȃ�D

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
    Case "=", " ", Chr(19), "�@", Chr(13)
      ' �����͍ŏ��̂�
      expr = ""
    Case "\"
      ' �X�C�b�`
      Dim k As Integer
      For k = i To tokens.Count
        field_Switch = field_Switch & tokens(k)
      Next
      Exit For
    Case "0", "1", "2", "3", "4", "5", "6", "7", "8", "9", ".", "+", "-", "*", "/", "(", ")", ",", """", "'", "<", ">", ""
      ' ���l�≉�Z�q�Ƃ���������̂͂��̂܂ܒǉ�
      expr = expr & tokens(i).text
    Case Else
      ' �����񓙂Ȃ̂ŁC�֐��łȂ���Εϐ��Ƃ݂Ȃ��ău�b�N�}�[�N����������D
            Dim isFunc As Boolean
      If i = tokens.Count Then
        isFunc = False
      ElseIf Left(tokens(i + 1).text, 1) = "(" Then
        isFunc = True
      Else
        isFunc = False
      End If

      If isFunc Then
        '�֐��͂��̂܂ܓn���D
        expr = expr & tokens(i)
      Else
        ' �ϐ��Ǝv����̂ŁC�u�b�N�}�[�N�̉��������݂�
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
  
  ' �쐬�����������G�N�Z���ŕ]������
  res = ExcelEval(expr)
  
  If TypeName(res) <> "Error" Then
    ' �蓮�ŃX�C�b�`�̏������s���͍̂���Ńo�O�̌����ƂȂ�̂ŁC
    ' �]�������l�ŃR�[�h�������ւ�����ōX�V���C���̂��ƃI���W�i���̃R�[�h�ɖ߂��D
    Dim original_code_text As String
    original_code_text = f.Code.text
  '  f.code.text = Left(original_code_text, 2) & res & " " & field_Switch
    f.Code.text = "=" & res & " " & field_Switch
    f.Update
    f.Code.text = original_code_text
  Else
    f.Update '�G���[�ɂȂ�̂��킩���Ă��Ă��X�V��Ƃ͕K�v
    ' ���ʂ��G���[������ŏ㏑������D
    f.Result.text = "Excel�Ή������ɃG���[������܂��u" & expr & "�v"
  End If
  
  ' �X�V������C�R�[�h�\�����I��������D�i�ʏ�̍X�V�ł̓���ɂ��킹��j
  f.ShowCodes = False
  Exit Sub
  
no_bookmark:
  f.Update  ' �G���[�ɂȂ邪�X�V�͕K�v
  f.Result.text = "�u�b�N�}�[�N������܂���u" & bm_name & "�v"
End Sub

