Attribute VB_Name = "�N���b�v�{�[�h"
Option Explicit

Sub PasteAsNormalText()
'
' PasteAsNormalText Macro
' �L�^�� 2012/02/05 �L�^�� -
'
    Selection.PasteAndFormat wdFormatPlainText '(wdPasteDefault)
End Sub

Sub ToGif()
'
' ToGif Macro
' �L�^�� 2013/03/25 �L�^�� NSC999
'
    Selection.Cut
    Selection.PasteSpecial link:=False, Datatype:=13, Placement:= _
        wdFloatOverText, DisplayAsIcon:=False
End Sub


' �C�����C���̊g�����^�t�@�C���Ƃ��ē\��t����
Sub PasteAsInlineEmf()
   
'   Selection.Collapse wdCollapseEnd
   Selection.PasteSpecial link:=False, Datatype:=wdPasteEnhancedMetafile, Placement:=wdInLine

End Sub
