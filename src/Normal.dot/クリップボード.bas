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
    Selection.PasteSpecial Link:=False, DataType:=13, Placement:= _
        wdFloatOverText, DisplayAsIcon:=False
End Sub


