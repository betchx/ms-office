Attribute VB_Name = "�}�\�ݒ�"
Option Explicit

Function �}���X�g���x��() As ListLevel
  Set �}���X�g���x�� = MainList().ListLevels(6)
End Function

Function �\���X�g���x��() As ListLevel
  Set �\���X�g���x�� = MainList().ListLevels(7)
End Function

Sub �}�\�ݒ�ԍ��̂�()
  �}���X�g���x��().NumberFormat = "�}%6"
  �\���X�g���x��().NumberFormat = "�\%7"
  �}���X�g���x��().ResetOnHigher = False
  �\���X�g���x��().ResetOnHigher = False
  �ݒ蔽�f
End Sub

Sub �}�\�ݒ�n�C�t��()
  �}���X�g���x��().NumberFormat = "�}-%6"
  �\���X�g���x��().NumberFormat = "�\-%7"
  �}���X�g���x��().ResetOnHigher = False
  �\���X�g���x��().ResetOnHigher = False
  �ݒ蔽�f
End Sub

Sub �}�\�ݒ�͔ԍ�()
  �}���X�g���x��().NumberFormat = "�}%1.%6"
  �\���X�g���x��().NumberFormat = "�\%1.%7"
  �}���X�g���x��().ResetOnHigher = 1
  �\���X�g���x��().ResetOnHigher = 1
  �ݒ蔽�f
End Sub


Sub �}�\�ݒ�ߔԍ�()
  �}���X�g���x��().NumberFormat = "�}%1.%2.%6"
  �\���X�g���x��().NumberFormat = "�\%1.%2.%7"
  �}���X�g���x��().ResetOnHigher = 2
  �\���X�g���x��().ResetOnHigher = 2
  �ݒ蔽�f
End Sub

Sub �}�\�ݒ菬�ߔԍ�()
  �}���X�g���x��().NumberFormat = "�}%1.%2.%3.%6"
  �\���X�g���x��().NumberFormat = "�\%1.%2.%3.%7"
  �}���X�g���x��().ResetOnHigher = 3
  �\���X�g���x��().ResetOnHigher = 3
  �ݒ蔽�f
End Sub

Sub �}�\�\�蕔�S�V�b�N��ONOFF()
    If �}���X�g���x��().Font.name = "�l�r �S�V�b�N" Then
       Dim f As Font
        Set f = ActiveDocument.Styles("�{��").Font
        �}���X�g���x��().Font.name = f.name '"�l�r ����"
        �\���X�g���x��().Font.name = f.name '"�l�r ����"
    Else
        �}���X�g���x��().Font.name = "�l�r �S�V�b�N"
        �\���X�g���x��().Font.name = "�l�r �S�V�b�N"
    End If
    �ݒ蔽�f
End Sub



Private Sub �ݒ蔽�f()
'  ActiveDocument.Content.ListFormat.ApplyListTemplate ListTemplate:=MainList()
  Dim p As Paragraph
  For Each p In ActiveDocument.Paragraphs
      If p.Style = "�\" Or p.Style = "�}" _
        Then p.Range.ListFormat.ApplyListTemplate ListTemplate:=MainList()
  Next
End Sub

Private Sub �\�̃X�^�C���ݒ�(ByVal �� As String)
  �X�^�C���R�s�[ ��, "�\"
  ActiveDocument.Styles("�\").BaseStyle = ��
  ActiveDocument.Styles("�\").NextParagraphStyle = "�\�{��"
End Sub

Private Sub �}�̃X�^�C���ݒ�(ByVal �� As String)
  �X�^�C���R�s�[ ��, "�}"
  ActiveDocument.Styles("�}").BaseStyle = ��
  ActiveDocument.Styles("�}").NextParagraphStyle = "�{��"
End Sub


