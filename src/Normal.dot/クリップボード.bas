Attribute VB_Name = "クリップボード"
Option Explicit

Sub PasteAsNormalText()
'
' PasteAsNormalText Macro
' 記録日 2012/02/05 記録者 -
'
    Selection.PasteAndFormat wdFormatPlainText '(wdPasteDefault)
End Sub

Sub ToGif()
'
' ToGif Macro
' 記録日 2013/03/25 記録者 NSC999
'
    Selection.Cut
    Selection.PasteSpecial link:=False, Datatype:=13, Placement:= _
        wdFloatOverText, DisplayAsIcon:=False
End Sub


' インラインの拡張メタファイルとして貼り付ける
Sub PasteAsInlineEmf()
   
'   Selection.Collapse wdCollapseEnd
   Selection.PasteSpecial link:=False, Datatype:=wdPasteEnhancedMetafile, Placement:=wdInLine

End Sub
