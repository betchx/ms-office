Attribute VB_Name = "Sheet"
Option Explicit

Sub Auto_Open()

    'Ctrl+Shift+Q
    Application.OnKey "+^q", "��t�����t��"
    
    ' Ctrl+Shift+L
    Application.OnKey "+^l", "�������t����"
    
    ' Ctrl+Shift+U
    Application.OnKey "+^u", "������t����"
    
    ' Ctrl+Shift+N
    Application.OnKey "+^n", "��t�����t������"
    
    ' Ctrl+Shift+H
    Application.OnKey "+^h", "������t�����t����������"

        
    ' Ctrl+T
    Application.OnKey "^t", "�J�����g�Z�����V�[�g�^�C�g����"
    
    
    ' Ctrl+Shift+E
    'Application.OnKey "+^E", "�ڐ����O��"
    Application.OnKey "+^E", "�L�������ݒ�"

    ' Ctrl+Shift+_  (����)
    Application.OnKey "+^\", "�r���N���A"
    
    ' Ctrl+Shift+R
    Application.OnKey "+^r", "Round�ǉ�"
        
    ' Ctrl+Shift+F
    Application.OnKey "+^f", "���r���ݒ�"
    
    ' Ctrl+Shift+B
    Application.OnKey "+^b", "�I��͈͂𑾘g�r���ň͂�"
    
    ' Ctrl+Shift+G
    Application.OnKey "+^g"
    Application.OnKey "+^g", "�I��͈͂Ɋi�q�r����ݒ�" ' <== "�O���b�h��"
    
    Dim i As Integer
    For i = 0 To 6
    ' Ctrl+Alt+0
      Application.OnKey "^%" & i, "NumFormatPoint" & i
    Next i
    
    ' Ctrl+Alt+7
    ' �擪�� ( ��ǉ�����D
    Application.OnKey "%7", "InsertOpenBracket"
    
    ' Ctrl+Alt+8
    ' ������ ( ��ǉ�����
    Application.OnKey "%8", "AppendOpenBracket"
    
    ' Ctrl+Alt+9
    ' ������ ) ��ǉ�����D
    Application.OnKey "%9", "AppendCloseBracket"
    
    
    ' Ctrl+Alt+U
    Application.OnKey "^%u", "�P�ʒǉ�"
    
    ' Ctrl+Alt+s
    Application.OnKey "%s", "�V�[�g���̑I���L��"
    
    ' �~�L���F Alt+:
    Application.OnKey "%:", "SetTimes"
    
    '�k�����ĕ\���F Alt+;
    Application.OnKey "%;", "�k�����ĕ\��"
       
    ' �V�[�g�ɖ��O��ǉ��F Ctrl+Alt+N
    Application.OnKey "^%N", "�V�[�g�ɖ��O��ǉ�"
    
    
    ' �s�v�s�̉B�� :  Ctrl+Alt+H
    Application.OnKey "%^h", "����w�肵�ĕs�v�s���B��"
    
    
    ' �Z�����͌�̈ړ��̗L����؂�ւ� :  Ctrl+M
    Application.OnKey "^m", "�Z�����͌�̈ړ��̗L����؂�ւ�"
    
    ' �Z�����͌�̈ړ��̕�����؂�ւ� :  Ctrl+J
    Application.OnKey "^j", "�Z�����͌�̈ړ��̕�����؂�ւ�"
    

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
           Case "_" ' ���t��
             rng.Characters(pos, 1).Delete
             
             rng.Characters(pos, 1).Font.Subscript = True
           
           Case "^" '��t��
             rng.Characters(pos, 1).Delete
             
             rng.Characters(pos, 1).Font.Superscript = True
            
           End Select
          
        Loop
      
      End If
    Next

End Sub


Sub Round�ǉ�()
    Dim rng As Range
    Dim n As Integer
    Dim eqn As String
    Dim fmt As String
    Dim add_round As Boolean
    For Each rng In Selection
      eqn = rng.Formula
      
      ' �s���ׂ�����̎�ނ𔻒�
      If Len(eqn) > 0 And Left(eqn, 1) = "=" Then
        '���ł�Round���ǉ�����Ă��Ȃ����`�F�b�N����D
        If Left(eqn, 6) = "=ROUND" Then
          add_round = False
        Else
          add_round = True
        End If
      Else
        '�����ȊO
        add_round = True
      End If
      
      ' Round�̑�������{
      If add_round Then
        ' add Round
        fmt = get_format(rng)
        If Left(fmt, 2) = "0." Then
           n = Len(fmt) - 2
        Else
           Dim s As String
           s = InputBox("Round�̌������w��.", "�����ݒ�")
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
        ' Round����
        Select Case Mid(eqn, 7, 1)
        Case "("
            '�ʏ��Round   ==> RounUp�ɕύX
            rng.Formula = "=ROUNDUP" & Mid(eqn, 7)
        Case "U"
            ' RoundUp  ==> RoundDown�ɕύX
            rng.Formula = "=RoundDown" & Mid(eqn, 9)
        Case "D"
            ' RoundDown ==> Ronud�Ȃ���
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

Private Sub Round�ǉ�_Original()
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
            s = InputBox("Round�̌������w��.", "�����ݒ�")
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

Sub �L�������ݒ�()
    Dim rng As Range
    Dim n As Integer
    Dim eqn As String
    
    Dim ans
    ans = InputBox("�L������", Default:="3")
    n = CInt(val(ans))
    If n < 1 Then
       MsgBox "�s���ȗL�������ł��D"
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

Sub ���r���ݒ�()

' �r���̃N���A�͍s��Ȃ��l�ɂ����D

   Call �I��͈͂Ɍr����ݒ�(Array(xlEdgeBottom))
   
   Exit Sub

'  �ȉ��̓I���W�i���̃}�N��
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



Sub �r���N���A()
    Dim r As Range
    Set r = Selection
    r.Borders.LineStyle = xlLineStyleNone
End Sub

Private Sub �I��͈͂Ɍr����ݒ�(targets, _
   Optional style As XlLineStyle = xlContinuous, _
   Optional weight As XlBorderWeight = xlThin, _
   Optional color = xlAutomatic)
    Dim tgt
    For Each tgt In targets
      With Selection.Borders(tgt)
        ' �r�������݂��Ȃ��ꍇ��weight��225�ɂȂ�͗l
        If .weight <> 225 Then
          .LineStyle = style
          .weight = weight
          .ColorIndex = color
        End If
      End With
    Next
   End Sub

Sub �I��͈͂𑾘g�r���ň͂�()
  Call �I��͈͂Ɍr����ݒ�( _
    Array(xlEdgeLeft, xlEdgeTop, xlEdgeBottom, xlEdgeRight) _
    , weight:=xlMedium)
End Sub


Sub �I��͈͂Ɋi�q�r����ݒ�()
  Call �I��͈͂Ɍr����ݒ�( _
    Array(xlEdgeLeft, xlEdgeTop, xlEdgeBottom, xlEdgeRight, _
          xlInsideVertical, xlInsideHorizontal) _
     )
End Sub



Sub �������t����()

   Dim b As Integer
   
   b = Len(ActiveCell.Text)
   Do While ActiveCell.Characters(b, 1).Font.Subscript
        b = b - 1
        If b <= 0 Then Exit Sub
   Loop
   ActiveCell.Characters(b, 1).Font.Subscript = True

End Sub

Sub ������t����()

   Dim b As Integer
   
   b = Len(ActiveCell.Text)
   Do While ActiveCell.Characters(b, 1).Font.Superscript
        b = b - 1
   Loop
   ActiveCell.Characters(b, 1).Font.Superscript = True

End Sub

Sub ������t�����t����������()

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


Sub ��t�����t������()
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

Sub �P�ʒǉ�()
  Dim unit As String
  unit = InputBox("�P�ʂ��L�����Ă��������D���l�Ƃ̊Ԃɂ͔��p�X�y�[�X���}������܂�")
  
  Dim r As Range
  For Each r In Selection
       set_format get_format(r) & " """ & unit & """", r
  Next

End Sub


Sub ��t�����t��()
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
    ActiveCell.FormulaR1C1 = "�~"
End Sub


Sub ���W�����v()
  JumpList.Show vbModeless
End Sub


Sub �V�[�g�ɖ��O��ǉ�()
    On Error GoTo eee:
    Dim r As Range
    Set r = Selection
    Dim s As String
    s = InputBox("���O���L��", "�V�[�g�ɖ��O��ݒ�", "")
    If s = "" Then Exit Sub
    Dim sh As Worksheet
    Set sh = ActiveSheet
    
    If r.MergeCells Then
      ' �Z�����}�[�W����Ă���
      If r.Cells(1, 1).MergeArea.Count = r.Count Then
        ' �P��̌����Z�����I������Ă���
        sh.Names.Add s, r.Cells(1, 1), True
        Exit Sub
      End If
    End If
    sh.Names.Add s, r, True
    
eee:

End Sub


Sub �w��̃V�[�g�ɖ��O��ǉ�()
  Dim tgt As Range
  Set tgt = ActiveCell.CurrentRegion
  
  Dim row As Integer
  For row = 1 To tgt.Count / 3
    Dim sheet_name As String
    sheet_name = tgt(row, 1).Value
    If sheet_name = "" Then Exit Sub
    Dim addr As String
    addr = tgt(row, 2).Value
    Dim ���O As String
    ���O = tgt(row, 3)
    Dim sh As Worksheet
    Set sh = ActiveWorkbook.Sheets(sheet_name)
    sh.Names.Add ���O, sh.Range(addr)
  Next


End Sub


Sub �k�����ĕ\��()
'
' �k�����ĕ\�� Macro
' �}�N���L�^�� : 2012/7/26  ���[�U�[�� : -
'
    Selection.ShrinkToFit = True
End Sub


Sub �V�[�g���̑I���L��()
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


Private Sub �Q�Ɛ悪������Ζ��O���폜(ByRef n As name)
      If Right(n.RefersTo, 5) = "#REF!" Or Mid(n.RefersTo, 2, 5) = "#REF!" Then
        n.Delete
      End If
End Sub


Sub ���݂̃u�b�N����Q�Ɛ�̖������O��S�č폜()
   Dim �V�[�g As Worksheet
   Dim n As name
   Dim tb As shape
   Dim s As Worksheet
   Set s = ActiveSheet
   Set tb = s.Shapes.AddTextbox(msoTextOrientationHorizontal, ActiveCell.Left, ActiveCell.Top, 100, 20)
   tb.TextFrame.Characters().Text = "Workbook"
   tb.Visible = msoTrue
   DoEvents
   For Each n In ActiveWorkbook.Names
      �Q�Ɛ悪������Ζ��O���폜 n
   Next
   For Each �V�[�g In ActiveWorkbook.Sheets
      tb.TextFrame.Characters().Text = �V�[�g.name
      DoEvents
      For Each n In �V�[�g.Names
          �Q�Ɛ悪������Ζ��O���폜 n
      Next
   Next
   
   tb.Delete

End Sub


Sub ���݂̃V�[�g����Q�Ɛ�̂Ȃ����O���폜()
   Dim n As name
   For Each n In ActiveSheet.Names
      �Q�Ɛ悪������Ζ��O���폜 n
   Next
End Sub

Sub SetLink()

Dim i As Integer
Dim s As Worksheet, s1 As Worksheet
Dim r As Range
Set s1 = ActiveWorkbook.Sheets("�ꗗ")
Set s = ActiveWorkbook.Sheets("�ꗗ")


For i = 1 To 20
    Set r = s.Cells(i + 3, 2)
    r.Hyperlinks.Add r, "Soft" & Format(i, "00") & "!B3"
Next
End Sub


Sub ����w�肵�ĕs�v�s���B��()
  Dim s As String
  s = InputBox("�B���w�����w�肵�Ă�������", "�B���`�F�b�N��w��", "A")
  If s <> "" Then
     �s�v�s�̉B�� s
  End If
End Sub



Sub �s�v�s�̉B��(Optional col As String = "A")
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


Sub �i�q()
'
' �i�q Macro
' �}�N���L�^�� : 2013/10/9  ���[�U�[�� : NSC999
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


Sub �Z�����͌�̈ړ��̗L����؂�ւ�()
'
   
    Application.MoveAfterReturn = Not Application.MoveAfterReturn
End Sub

Sub �Z�����͌�̈ړ��̕�����؂�ւ�()
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
