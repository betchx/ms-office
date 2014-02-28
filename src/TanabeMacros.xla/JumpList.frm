VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} JumpList 
   Caption         =   "UserForm1"
   ClientHeight    =   11235
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   5415
   OleObjectBlob   =   "JumpList.frx":0000
   StartUpPosition =   1  'オーナー フォームの中央
End
Attribute VB_Name = "JumpList"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim arr()
Dim col As Integer

Private Sub ListBox1_Click()
If UBound(arr) > 0 Then ActiveSheet.Cells(arr(Me.ListBox1.ListIndex, 1), col).Activate
Unload Me
End Sub

Private Sub UserForm_Initialize()
Dim sht As Worksheet

Set sht = ActiveSheet

Dim c As Range
Dim n As Integer
n = WorksheetFunction.CountIf(sht.Range("A:A"), "☆*")
If n = 0 Then Exit Sub
ReDim arr(n - 1, 1)

col = 2
'If Range("A1").HasFormula Then
' Dim wk As Integer
' wk = CInt(Range("A1").Value)
' If wk > 0 Then col = wk
'End If
 
Do While WorksheetFunction.CountA(sht.Columns(col)) = 0
  col = col + 1
Loop

Dim i As Integer, k As Integer
Dim r As Range
k = 1

For i = 0 To n - 1
 Do
   k = k + 1
   Set r = sht.Cells(k, 1)
 Loop Until Left(r.Value, 1) = "☆"
 arr(i, 1) = k
 arr(i, 0) = r.Offset(0, col - 1).Value
Next

Me.ListBox1.ColumnCount = 1
Me.ListBox1.List = arr

End Sub
