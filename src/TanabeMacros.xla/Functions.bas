Attribute VB_Name = "Functions"
Option Explicit
Function pnt2mm(p As Single) As Single

  pnt2mm = p / 72# * 25.4
  
End Function


Function mm2pnt(m As Single) As Single
  mm2pnt = m / 25.4 * 72#
End Function

Function r2i(r As Range) As Integer
  r2i = CInt(val(r.Value))
End Function

Function r2s(r As Range) As Single
  r2s = CSng(val(r.Value))
End Function

Function n2s(n As String, Optional i As Integer = 0) As Single
  n2s = r2s(n2r(n, i))
End Function

' ���O���烌���W���擾����
' �܂��̓A�N�e�B�u�V�[�g�ŒT���C������Ȃ��ꍇ�̓��[�N�u�b�N����T��
' �ǂ��ɂ��邩���킩���Ă���� Range("���O")�ł��T����
Function n2r(n As String, Optional i As Integer = 0) As Range
  Dim m As name
'  For Each m In ActiveSheet.Names
'    If m.Name = n Then
'      If i = 0 Then
'        n2r = m.RefersToRange
'      Else
'        n2r = m.RefersToRange.Cells(i, 1)
'     Return
'    End If
'  Next
If i = 0 Then
On Error GoTo kkk:
  Set n2r = ActiveSheet.Names(n).RefersToRange
  Exit Function
kkk:
  ' �ύX�F ThisWorkbook => ActiveWorkbook  (Personal.xls�Ɉړ������̂Łj
  For Each m In ActiveWorkbook.Names
    If m.name = n Then
      Set n2r = m.RefersToRange
      Exit Function
    End If
  Next
  Set n2r = Nothing
Else
On Error GoTo ttt:
  Set n2r = ActiveSheet.Names(n).RefersToRange.Cells(i, 1)
  Exit Function
ttt:
  For Each m In ActiveWorkbook.Names
     If m.name = n Then
       Set n2r = m.RefersToRange.Cells(i, 1)
       Exit Function
     End If
  Next
  Set n2r = Nothing
End If

Exit Function

nnn:
  Set n2r = Nothing

End Function


