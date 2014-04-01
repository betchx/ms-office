VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} ブックマーク名編集 
   Caption         =   "ブックマーク名の編集"
   ClientHeight    =   1230
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   6210
   OleObjectBlob   =   "ブックマーク名編集.frx":0000
   StartUpPosition =   1  'オーナー フォームの中央
End
Attribute VB_Name = "ブックマーク名編集"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Public 結果 As String

Private Sub cbCancel_Click()
   結果 = ""
   Me.Hide
End Sub

Private Sub cbOK_Click()
    結果 = TextBox1.text
    Me.Hide
End Sub

Private Sub cbReset_Click()
  TextBox1.text = ブックマーク可能文字への変換(候補)
End Sub

Private Sub TextBox1_Change()
  Dim bk As String
  bk = ブックマーク可能文字への変換(TextBox1.text)
  If TextBox1.text <> bk Then
    ' 使用できない文字が与えられた場合は使えるものに変換して再設定
    TextBox1.text = bk
  End If
End Sub

Private Sub UserForm_Initialize()

' ngs = Array(" ", "　", "(", ")", "-", "?", ".", ",", "/", "!", "*", "%", "#", "'", "=", "^", "~", "\", "|", Chr(10), Chr(13))
 
  結果 = ""
 
  Label1.Caption = _
"ブックマークを追加・編集してください．" & vbCrLf & _
"最大20文字です．スペース等は使えません．" & vbCrLf & _
"その他の使用できない文字は全角に変換されます．"

End Sub

Sub 候補設定(Optional 候補 As String = "")
'"ブックマークは最大20文字までで，数字から始まることはできません．" & vbCrLf & _
'"また，改行やスペース（半角全角とも）と次の半角文字は使用できません： ()-?.,/!*%#'=^~\|"
  If 候補 = "" Then
    If Selection.Start = Selection.End Then
      候補 = ブックマーク可能文字への変換(Selection.Paragraphs(1).Range.text)
    Else
      候補 = ブックマーク可能文字への変換(Selection.text)
    End If
  End If
  TextBox1.text = ブックマーク可能文字への変換(候補)
  TextBox1.SetFocus
End Sub
