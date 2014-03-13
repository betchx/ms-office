Attribute VB_Name = "Plots"
Option Explicit

Private Const app As String = "axes"

 ' �������F�I���_�C�A���O�ׂ̈̐ݒ�

Private Type ChooseColor
  lStructSize As Long
  hWndOwner As Long
  hInstance As Long
  rgbResult As Long
  lpCustColors As String
  flags As Long
  lCustData As Long
  lpfnHook As Long
  lpTemplateName As String
End Type

Private Declare Function ChooseColor Lib "comdlg32.dll" Alias "ChooseColorA" _
                                      (pChoosecolor As ChooseColor) As Long

Private Const CC_RGBINIT = &H1                '�F�̃f�t�H���g�l��ݒ�
Private Const CC_LFULLOPEN = &H2              '�F�̍쐬���s��������\��
Private Const CC_PREVENTFULLOPEN = &H4        '�F�̍쐬�{�^���𖳌��ɂ���
Private Const CC_SHOWHELP = &H8               '�w���v�{�^����\��
 ' �����܂ŐF�I���_�C�A���O�ׂ̈̐ݒ�
'
Sub Test_GetColorDlg()
   Dim col As Long
   col = GetColorDlg()
   MsgBox "Selected color rgb is " & Format(col / 256 / 256, "0") & "," & _
   Format((col / 256) Mod 256, "0") & ", " & Format(col Mod 256, "0")
   
End Sub


' �������� �F�I���_�C�A���O���Ăяo����RGB���擾����֐�
' Access�p��Web����R�s�[���Ĕ�����
' �C�����e�F �E�C���h�E�n���h���̃v���p�e�B����Access��Excel�ňႤ�̂ŏC��
' �R�s�[��URL: http://www.tsware.jp/tips/tips_343.htm
' �R�s�[���F 2011-11-21

Public Function GetColorDlg(Optional ByVal lngDefColor As Long = 0) As Long
'�@�\ �F �F�̐ݒ�_�C�A���O��\�����A�����őI�����ꂽ�F��RGB�l��Ԃ�
'���� �F lngDefColor �f�t�H���g�\������F
'�Ԓl �F ������ RGB�l   �L�����Z����-1  �G���[�� -2  �i�[���͍��Ȃ̂Œ��Ӂj

  Dim udtChooseColor As ChooseColor
  Dim lngRet As Long

  With udtChooseColor
    '�_�C�A���O�̐ݒ�
    .lStructSize = Len(udtChooseColor)
    .hWndOwner = Application.Hwnd  ' <= ���������ύX���K�v������
    .lpCustColors = String$(64, Chr$(0))
    .flags = CC_RGBINIT + CC_LFULLOPEN
    .rgbResult = lngDefColor
    '�_�C�A���O��\��
    lngRet = ChooseColor(udtChooseColor)
    '�_�C�A���O����̕Ԃ�l���`�F�b�N
    If lngRet <> 0 Then
      If .rgbResult > RGB(255, 255, 255) Then
        '�G���[
        GetColorDlg = -2
      Else
        '����I���ARGB�l��Ԃ�l�ɃZ�b�g
        GetColorDlg = .rgbResult
      End If
    Else
      '�L�����Z���������ꂽ�Ƃ�
      GetColorDlg = -1
    End If
  
  End With
  
End Function
' �����܂� �F�I���_�C�A���O���Ăяo����RGB���擾����֐�





Sub SaveAxesRange()

Dim a As Axis
On Error GoTo eee:

uf���͈͂̕ۑ��I��.Show vbModal

Set a = ActiveChart.Axes(xlValue)
If uf���͈͂̕ۑ��I��.�I�� > 1 Then
    ' Y on
    SaveSetting app, "y", "enable", True
    SaveSetting app, "y", "max", a.MaximumScale
    SaveSetting app, "y", "min", a.MinimumScale
    SaveSetting app, "y", "tick", a.MajorUnit
    SaveSetting app, "y", "mtick", a.MinorUnit
    SaveSetting app, "y", "crossesat", a.CrossesAt
Else
    SaveSetting app, "y", "enable", False
End If

Set a = ActiveChart.Axes(xlCategory)
If uf���͈͂̕ۑ��I��.�I�� Mod 2 = 1 Then
    SaveSetting app, "x", "enable", True
    SaveSetting app, "x", "max", a.MaximumScale
    SaveSetting app, "x", "min", a.MinimumScale
    SaveSetting app, "x", "tick", a.MajorUnit
    SaveSetting app, "x", "mtick", a.MinorUnit
    SaveSetting app, "x", "crossesat", a.CrossesAt
Else
    SaveSetting app, "x", "enable", False
End If

eee:
End Sub

Sub ApplyAxesRange()
Attribute ApplyAxesRange.VB_ProcData.VB_Invoke_Func = " \n14"
On Error GoTo eee:
Dim s As Boolean
s = Application.ScreenUpdating
Application.ScreenUpdating = False
Dim a As Axis


Set a = ActiveChart.Axes(xlValue)
If CBool(GetSetting(app, "y", "enable", "False")) Then
    a.MaximumScale = val(GetSetting(app, "y", "max", "1"))
    a.MinimumScale = val(GetSetting(app, "y", "min", "0"))
    a.MajorUnit = val(GetSetting(app, "y", "tick", "1"))
    a.MinorUnit = val(GetSetting(app, "y", "mtick", "0.1"))
    a.CrossesAt = val(GetSetting(app, "y", "crossesat", "0"))
End If

Set a = ActiveChart.Axes(xlCategory)
If CBool(GetSetting(app, "x", "enable", "False")) Then
    a.MaximumScale = val(GetSetting(app, "x", "max", "1"))
    a.MinimumScale = val(GetSetting(app, "x", "min", "0"))
    a.MajorUnit = val(GetSetting(app, "x", "tick", "1"))
    a.MinorUnit = val(GetSetting(app, "x", "mtick", "0.1"))
    a.CrossesAt = val(GetSetting(app, "x", "crossesat", "0"))
End If

eee:
Application.ScreenUpdating = s

End Sub

Sub CopyAxesRangeX2Y()
Attribute CopyAxesRangeX2Y.VB_ProcData.VB_Invoke_Func = " \n14"
   CopyAxesRange ActiveChart.Axes(xlCategory), ActiveChart.Axes(xlValue)
End Sub

Sub CopyAxesRangeY2X()
Attribute CopyAxesRangeY2X.VB_ProcData.VB_Invoke_Func = " \n14"
   CopyAxesRange ActiveChart.Axes(xlValue), ActiveChart.Axes(xlCategory)
End Sub

Private Sub CopyAxesRange(f As Axis, t As Axis)
On Error GoTo eee:
Dim s As Boolean
s = Application.ScreenUpdating
Application.ScreenUpdating = False

t.MaximumScale = f.MaximumScale
t.MinimumScale = f.MinimumScale
t.MajorUnit = f.MajorUnit
t.MinorUnit = f.MinorUnit
t.CrossesAt = f.CrossesAt

eee:
Application.ScreenUpdating = s

End Sub



Sub ApplyAxesRangeForAll()
Attribute ApplyAxesRangeForAll.VB_ProcData.VB_Invoke_Func = " \n14"
On Error GoTo eee:
Application.ScreenUpdating = False
Dim c
Dim s As Worksheet
Set s = ActiveSheet

For Each c In s.ChartObjects
   c.Activate
   Call ApplyAxesRange
Next

s.Range("G7").CurrentRegion.Select

eee:
Application.ScreenUpdating = True
End Sub



Sub XYXY�O���t()

On Error GoTo eee

Application.ScreenUpdating = False


Dim s As Worksheet
Set s = ActiveSheet

Dim r As Range
Dim ori As Range


Set r = ActiveCell.CurrentRegion
Set ori = r.Range("A1")

Dim nCol As Integer
nCol = r.Columns.Count

Dim rX As Range, rY As Range

Dim iCol As Integer

Dim co As ChartObject
Dim g As Chart
Dim i As Integer

Set g = Charts.Add()
With g
  .ChartType = xlXYScatterLinesNoMarkers
  .SetSourceData source:=Range(ori, ori.End(xlDown).Offset(0, 1)), _
      PlotBy:=xlColumns
  
  For iCol = 3 To nCol Step 2
    i = (iCol + 1) / 2
    .SeriesCollection.NewSeries
    .SeriesCollection(i).XValues = Range(ori.Offset(1, iCol - 1), ori.Offset(1, iCol - 1).End(xlDown))
    .SeriesCollection(i).values = Range(ori.Offset(1, iCol), ori.Offset(1, iCol).End(xlDown))
    .SeriesCollection(i).name = ori.Offset(0, iCol)
  Next iCol
  
  .Location Where:=xlLocationAsObject, name:=s.name
  
  ' �}�N���̐����ɂ��C�^�C�g�����͕ʓr�ݒ肷��ق����ǂ��̂ŃR�����g�A�E�g �� 2012/04/04
'  .HasTitle = True
'  .ChartTitle.Characters.Text = "GraphTitle"
'  .Axes(xlCategory, xlPrimary).HasTitle = True
'  .Axes(xlCategory, xlPrimary).AxisTitle.Characters.Text = "Xtitle"
'  .Axes(xlValue, xlPrimary).HasTitle = True
'  .Axes(xlValue, xlPrimary).AxisTitle.Characters.Text = "Ytitle"
End With

eee:
  Application.ScreenUpdating = True

End Sub

Private Function ThickerLine(X As XlBorderWeight) As XlBorderWeight
Select Case X
  Case xlHairline
    ThickerLine = xlThin
  Case xlThin
    ThickerLine = xlMedium
  Case Else
    ThickerLine = xlThick
End Select
End Function

Private Function ThinerLine(X As XlBorderWeight) As XlBorderWeight
Select Case X
  Case xlMedium
    ThinerLine = xlThin
  Case xlThick
    ThinerLine = xlMedium
  Case Else
    ThinerLine = xlHairline
End Select
End Function


Sub �O���t�̐��𑾂�()

Dim s As Series

On Error GoTo eee
Application.ScreenUpdating = False

Dim n As String

n = TypeName(Selection)

If n = "Series" Then
  If Selection.Border.ColorIndex <> xlColorIndexNone Then
    Selection.Border.weight = ThickerLine(Selection.Border.weight)
  End If
Else
  For Each s In ActiveChart.SeriesCollection
    If s.Border.ColorIndex <> xlColorIndexNone Then
      s.Border.weight = ThickerLine(s.Border.weight)
    End If
  Next s
End If

eee:

Application.ScreenUpdating = True

End Sub

Sub �O���t�̐����ׂ�()

On Error GoTo eee
Application.ScreenUpdating = False

Dim s As Series
Dim n As String

n = TypeName(Selection)

If n = "Series" Then
  If Selection.Border.ColorIndex <> xlColorIndexNone Then
    Selection.Border.weight = ThinerLine(Selection.Border.weight)
  End If
Else
  For Each s In ActiveChart.SeriesCollection
    If s.Border.ColorIndex <> xlColorIndexNone Then
      s.Border.weight = ThinerLine(s.Border.weight)
    End If
  Next s
End If

eee:

Application.ScreenUpdating = True

End Sub


Sub �O���t�̃}�[�J�[��傫��()

Dim s As Series
On Error GoTo eee
Application.ScreenUpdating = False
Dim n As String
n = TypeName(Selection)

If n = "Series" Then
  Selection.MarkerSize = Selection.MarkerSize + 1
Else
  For Each s In ActiveChart.SeriesCollection
    s.MarkerSize = s.MarkerSize + 1
  Next s
End If
eee:

Application.ScreenUpdating = True

End Sub

Sub �O���t�̃}�[�J�[��������()

Dim s As Series
On Error GoTo eee
Application.ScreenUpdating = False
Dim n As String
n = TypeName(Selection)

Dim sz As Integer
If n = "Series" Then
  sz = Selection.MarkerSize - 1
  If sz < 2 Then sz = 2
  Selection.MarkerSize = sz
Else
  For Each s In ActiveChart.SeriesCollection
    sz = s.MarkerSize - 1
    If sz < 2 Then sz = 2
    s.MarkerSize = sz
  Next s
End If
eee:
Application.ScreenUpdating = True
End Sub


Sub �O���t�̃}�[�J�[���ۂ�()

Dim s As Series
On Error GoTo eee
Application.ScreenUpdating = False
Dim n As String
n = TypeName(Selection)

Dim sz As Integer
If n = "Series" Then
  Selection.MarkerStyle = xlMarkerStyleCircle
Else
  For Each s In ActiveChart.SeriesCollection
    s.MarkerStyle = xlMarkerStyleCircle
  Next s
End If
eee:
Application.ScreenUpdating = True
End Sub

Sub �O���t�̃}�[�J�[�𖳂���()

Dim s As Series
On Error GoTo eee
Application.ScreenUpdating = False
Dim n As String
n = TypeName(Selection)

Dim sz As Integer
If n = "Series" Then
  Selection.MarkerStyle = xlMarkerStyleNone
Else
  For Each s In ActiveChart.SeriesCollection
    s.MarkerStyle = xlMarkerStyleNone
  Next s
End If
eee:
Application.ScreenUpdating = True
End Sub

Sub �O���t�̐��𖳂���()

Dim s As Series
On Error GoTo eee
Application.ScreenUpdating = False
Dim n As String
n = TypeName(Selection)

Dim sz As Integer
If n = "Series" Then
  �V���[�Y�̐������� Selection
Else
  For Each s In ActiveChart.SeriesCollection
    �V���[�Y�̐������� s
  Next s
End If
eee:
Application.ScreenUpdating = True
End Sub

Private Sub �V���[�Y�̐�������(s As Series)
 If s.MarkerStyle = xlMarkerStyleNone Then s.MarkerStyle = xlMarkerStyleAutomatic
 s.Border.ColorIndex = xlColorIndexNone
End Sub



Sub set_coror_list()
  Dim i As Integer
  Dim sc As SeriesCollection
  Set sc = ActiveChart.SeriesCollection
  For i = 1 To sc.Count
    sc.item(i).Border.ColorIndex = i - 1
    sc.item(i).MarkerBackgroundColorIndex = i - 1
    sc.item(i).MarkerForegroundColorIndex = i - 1
  Next i

End Sub


Sub �����F���g���Ȃ�����̐F��ݒ�()

Dim s As String
Dim cmax As Integer
Dim nsame As Integer

s = InputBox("�O���t�̐��Ɏg�p����F�̐����w�肵�Ă�������", "���F���ݒ�", "8")
If s <> "" Then
  cmax = CInt(val(s))
  s = InputBox("�O���t�œ����F���g�p���鐔���w�肵�Ă�������", "���F�����w��", "2")
  If s <> "" Then
    nsame = CInt(val(s))
  
    SeriesColorSetByCmax cmax, nsame
  End If
End If

End Sub

Sub ������()
  SeriesColorSetByCmax 1, ActiveChart.SeriesCollection.Count
End Sub


Sub ���ϕ\���O���t�̐ݒ�()
  On Error GoTo xxx
  �O���t�̃}�[�J�[�𖳂���
  Dim res As String
  Dim val
  Dim �Z�x
  Dim ��{���b�Z�[�W
  Dim ���b�Z�[�W
  Dim ��ʐF
   
  ��{���b�Z�[�W = "���ϒl�ȊO�̃f�[�^�̔Z���𐔒l�œ��͂��Ă��������D" & vbCrLf _
    & "��:0    50%�D�F�F127  ���F255  (�����̏ꍇ)" & vbCrLf _
    & "��:0.0  50%�D�F�F0.5  ���F1.0  (�����̏ꍇ)" & vbCrLf _
    & " �������� ? ��1�����Ƃ���ΐF�I���_�C�A���O��(�J���[��)�I���ł��܂�"
  
  ���b�Z�[�W = ��{���b�Z�[�W
  
  Do
    res = InputBox(���b�Z�[�W, "�Z�x�̐ݒ�", "192")
    If res = "" Then
      �Z�x = 192
    ElseIf res = "?" Then
      ��ʐF = GetColorDlg(RGB(192, 192, 192))
      �Z�x = 0
    ElseIf InStr(res, ".") = 0 Then
      val = CInt(res)
      If val >= 0 And val <= 255 Then
        �Z�x = val
      Else
        ���b�Z�[�W = "�Z�x��0�ȏ�255�ȉ��ł���K�v������܂��D" & vbCrLf _
           & ��{���b�Z�[�W
      End If
    Else
      val = CSng(res)
      If val >= 0# And val <= 1# Then
        �Z�x = Round(255 * val)
      Else
        ���b�Z�[�W = "�Z�x��0.0�ȏ�1.0�ȉ��ł���K�v������܂��D" & vbCrLf _
            & ��{���b�Z�[�W
      End If
    End If
  Loop While IsEmpty(�Z�x)
 

  
  If IsEmpty(��ʐF) Then ��ʐF = RGB(�Z�x, �Z�x, �Z�x)
  
  Application.ScreenUpdating = False
     ActiveChart.HasLegend = False
     Dim n As Integer, i As Integer
     n = ActiveChart.SeriesCollection.Count
    
     For i = 1 To n - 1
       �F�ݒ� ActiveChart.SeriesCollection(i), xlContinuous, ��ʐF
     Next
     Dim ave As Series
     Set ave = ActiveChart.SeriesCollection(n)
     �F�ݒ� ave, xlContinuous, RGB(0, 0, 0)
     ave.Border.weight = ThickerLine(ave.Border.weight)
  Application.ScreenUpdating = True
  
  �O���t�̊�{�ݒ� �F�ݒ���s:=False

xxx:
  
End Sub

Sub Y�ΐ�()
  ���̑ΐ��ݒ� xlValue
End Sub

Sub X�ΐ�()
  ���̑ΐ��ݒ� xlCategory
End Sub

Private Sub ���̑ΐ��ݒ�(��)
   On Error GoTo eee
   Dim ax As Axis
   Set ax = ActiveChart.Axes(��)
   Application.ScreenUpdating = False
   If ax.ScaleType = xlScaleLinear Then
    If Not ax.MinimumScaleIsAuto And ax.MinimumScale = 0 Then ax.MinimumScaleIsAuto = True
    ax.ScaleType = xlScaleLogarithmic
    If ax.MinimumScaleIsAuto Then ax.MinimumScale = 1
    ax.HasMajorGridlines = True
    ax.HasMinorGridlines = True
    ax.MajorGridlines.Border.LineStyle = xlContinuous
    ax.MinorGridlines.Border.LineStyle = xlDot
    ax.MinorGridlines.Border.color = RGB(64, 64, 64)
   Else
    ax.ScaleType = xlScaleLinear
    ax.MinimumScaleIsAuto = True
    ax.MaximumScaleIsAuto = True
    ax.HasMinorGridlines = False
    If �� = xlCategory Then ax.HasMajorGridlines = False
   End If
eee:
   Application.ScreenUpdating = True
End Sub



Sub SeriesColorSetWithLimetedColor()

Dim s As String
Dim cmax As Integer

s = InputBox("�O���t�̐��Ɏg�p����F�̐����w�肵�Ă�������", "���F���ݒ�", "8")
If s <> "" Then
  cmax = CInt(val(s))
  SeriesColorSetByCmax cmax, 1
End If

End Sub


Sub SeriesColorSet()

SeriesColorSetByCmax 8

End Sub


' �n��̐F�ݒ���s���֐�
Function SeriesColorSetByCmax(cmax As Integer, Optional nsame As Integer = 1)

Dim �\���ݒ� As Boolean
On Error GoTo eee
�\���ݒ� = Application.ScreenUpdating
Application.ScreenUpdating = False


Dim c(0 To 8)
c(1) = RGB(0, 0, 0)
c(2) = RGB(255, 0, 0)
c(3) = RGB(0, 0, 255)
c(4) = RGB(0, 255, 0)
c(5) = RGB(127, 0, 127)
c(6) = RGB(0, 127, 127)
c(7) = RGB(127, 127, 0)
c(8) = RGB(127, 127, 127)


Dim sc As SeriesCollection
Set sc = ActiveChart.SeriesCollection

Dim i As Integer
If cmax > 8 Then cmax = 8

c(0) = c(cmax)

Dim hozon
Dim ci As Integer

Dim lt As Integer
Dim ltmax As Integer
Dim lts(0 To 6)
ltmax = 7
lts(0) = xlContinuous
lts(1) = xlDash
lts(2) = xlDot
lts(3) = xlDashDot
lts(4) = xlDashDotDot
lts(5) = xlSlantDashDot
lts(6) = xlDouble


Dim idx As Integer

For i = 1 To sc.Count
  idx = (i - 1) \ nsame
  

  lt = ((idx) \ cmax) Mod ltmax
  If cmax > 1 Then
    ci = (idx + 1) Mod cmax
  
  Else
    ci = 1
  End If

'  sc.Item(i).Border.Color = c(ci)
'  hozon = sc.Item(i).ChartType
'  sc.Item(i).Border.LineStyle = lts(lt)
'  sc.Item(i).MarkerBackgroundColor = c(ci)
'  sc.Item(i).MarkerForegroundColor = c(ci)
'  sc.Item(i).ChartType = hozon
   �F�ݒ� sc.item(i), lts(lt), c(ci)
 
Next i


eee:
Application.ScreenUpdating = �\���ݒ�

End Function


Private Sub �F�ݒ�(�V���[�Y As Series, �X�^�C��, �F)
Dim �ۑ�
With �V���[�Y
'   �ۑ� = .ChartType
   If .Border.ColorIndex <> xlColorIndexNone Then
     .Border.LineStyle = �X�^�C��
     .Border.color = �F
   End If
   '�ۑ� = .MarkerStyle
   
   Select Case .MarkerStyle
   Case xlMarkerStyleCircle, xlMarkerStyleDiamond, xlMarkerStyleTriangle, xlMarkerStyleSquare
     .MarkerForegroundColor = �F
     .MarkerBackgroundColor = �F
   Case xlMarkerStyleNone
     '' Do nothing
   Case Else
'     .MarkerBackgroundColor = �F
     .MarkerForegroundColor = �F
     .MarkerBackgroundColorIndex = xlColorIndexNone
   End Select
   '.MarkerStyle = �ۑ�
'   .ChartType = �ۑ�
End With
End Sub


Sub �O���t�ɒ����}�[�J�[��ݒ�()

Dim s As Series
On Error GoTo eee
Application.ScreenUpdating = False
Dim n As String
n = TypeName(Selection)

If n = "Series" Then
  If Selection.MarkerForegroundColorIndex = xlColorIndexNone Then
    Selection.MarkerForegroundColor = Selection.Border.color
  End If
  Selection.MarkerBackgroundColorIndex = Selection.MarkerForegroundColorIndex 'Border.Color
Else
  For Each s In ActiveChart.SeriesCollection
  If s.MarkerForegroundColorIndex = xlColorIndexNone Then
    s.MarkerForegroundColor = s.Border.color
  End If
  s.MarkerBackgroundColorIndex = s.MarkerForegroundColorIndex
  Next s
End If
eee:
Application.ScreenUpdating = True

End Sub

Sub �O���t�ɒ���}�[�J�[��ݒ�()

Dim s As Series
On Error GoTo eee
Application.ScreenUpdating = False
Dim n As String
n = TypeName(Selection)

If n = "Series" Then
  If Selection.MarkerForegroundColorIndex = xlColorIndexNone Then
    Selection.MarkerForegroundColor = Selection.Border.color
  End If
  Selection.MarkerBackgroundColorIndex = xlColorIndexNone ' RGB(255, 255, 255) 'Selection.Border.Color
Else
  For Each s In ActiveChart.SeriesCollection
  If s.MarkerForegroundColorIndex = xlColorIndexNone Then
    s.MarkerForegroundColor = s.Border.color
  End If
  s.MarkerBackgroundColorIndex = xlColorIndexNone 'RGB(255, 255, 255)
  's.MarkerForegroundColor = s.Border.Color
  Next s
End If
eee:
Application.ScreenUpdating = True

End Sub


Sub �O���t�ɔ������}�[�J�[��ݒ�()

Dim s As Series
On Error GoTo eee
Application.ScreenUpdating = False
Dim n As String
n = TypeName(Selection)

If n = "Series" Then
  If Selection.MarkerForegroundColorIndex = xlColorIndexNone Then
    Selection.MarkerForegroundColor = Selection.Border.color
  End If
  Selection.MarkerBackgroundColor = RGB(255, 255, 255) 'Selection.Border.Color
Else
  For Each s In ActiveChart.SeriesCollection
  If s.MarkerForegroundColorIndex = xlColorIndexNone Then
    s.MarkerForegroundColor = s.Border.color
  End If
  s.MarkerBackgroundColor = RGB(255, 255, 255)
  's.MarkerForegroundColor = s.Border.Color
  Next s
End If
eee:
Application.ScreenUpdating = True

End Sub



Sub �O���t�̔w�i�Ȃ���()
  On Error GoTo eee:
    ActiveChart.PlotArea.Interior.ColorIndex = xlNone
'    ActiveChart.ChartArea.Interior.ColorIndex = xlNone
'    ActiveChart.ChartArea.Border.ColorIndex = xlNone
eee:
End Sub

Sub �O���t�̔w�i�𓧉߂�()
  On Error GoTo eee:
    ActiveChart.PlotArea.Interior.ColorIndex = xlNone
    ActiveChart.ChartArea.Interior.ColorIndex = xlNone
    ActiveChart.ChartArea.Border.ColorIndex = xlNone
eee:
End Sub

Sub X���L���v�V�����ݒ�()
    Dim str As String
    Dim a As Axis
    Dim obj
    ufXAx.Show
    str = ufXAx.Label()
    
    If Len(str) > 0 Then
        If TypeName(Selection) = "ChartArea" Then
          Set a = Selection.Parent.Axes(xlCategory)
          If Not a Is Nothing Then
            a.HasTitle = True
            a.AxisTitle.Characters.Text = str
          End If
        Else
            For Each obj In Selection.ShapeRange
              Set a = Nothing
              Select Case TypeName(obj)
              Case "ChartObject"
                Set a = obj.Chart.Axes(xlCategory)
              Case "Chart"
                Set a = obj.Axes(xlCategory)
              End Select
              If Not a Is Nothing Then
                a.HasTitle = True
                a.AxisTitle.Characters.Text = str
              End If
            Next
        End If
    End If

End Sub

Sub Y���L���v�V�����ݒ�()
    Dim str As String
    Dim ax As Axis
    Dim obj
    Dim isHorizontal As Boolean
    
    isHorizontal = False
    
    If Not ActiveChart Is Nothing Then
      Select Case ActiveChart.SeriesCollection(1).ChartType
      Case xlBarClustered, xlBarStacked, xlBarStacked100, _
            xl3DBarClustered, xl3DBarStacked, xl3DBarStacked100
          isHorizontal = True
      End Select
    Else
      For Each obj In Selection
        Select Case TypeName(obj)
        Case "ChartObject"
          Select Case obj.Chart.SeriesCollection(1).ChartType
          Case xlBarClustered, xlBarStacked, xlBarStacked100, _
                xl3DBarClustered, xl3DBarStacked, xl3DBarStacked100
              isHorizontal = True
              Exit For
          End Select
        Case "Chart"
          Select Case obj.SeriesCollection(1).ChartType
          Case xlBarClustered, xlBarStacked, xlBarStacked100, _
                xl3DBarClustered, xl3DBarStacked, xl3DBarStacked100
              isHorizontal = True
              Exit For
          End Select
        End Select
      Next
    End If
    
    ufYAx.Show
    str = ufYAx.Label()
    
    
    If Len(str) > 0 Then
      Dim arr
      If ActiveChart Is Nothing Then
        Set arr = Selection
      Else
        arr = Array(ActiveChart)
      End If
      For Each obj In arr
        Set ax = Nothing
        Select Case TypeName(obj)
        Case "ChartObject"
           Set ax = obj.Chart.Axes(xlValue)
        Case "Chart"
           Set ax = obj.Axes(xlValue)
        End Select
        If Not ax Is Nothing Then
          With ax
            .HasTitle = True
            If ufYAx.CheckBox1.Value And Not isHorizontal Then
              .AxisTitle.Orientation = xlHorizontal
              str = Replace(str, "|", vbLf)
              str = Replace(str, " ", vbLf)
            Else
              str = Replace(str, "|", "")
            End If
            .AxisTitle.Characters.Text = str
          End With
        End If
      Next
    End If
End Sub


Sub �O���t�̊�{�ݒ�(Optional �F�ݒ���s As Boolean = True)
'
' �O���t�̊�{�ݒ� Macro
' �}�N���L�^�� : 2011/2/21  ���[�U�[�� : �c��
'

Dim str As String
Dim a As Axis
Dim General As String

' �o�[�W�����ɂ��C�w�肷�ׂ����̂��Ⴄ
If Application.Version <= "11.0" Then
  General = "General"
Else
  General = "G/�W��"
End If

On Error GoTo eee
Application.ScreenUpdating = False
    
    Set a = ActiveChart.Axes(xlCategory)
    With a
      If Not .HasTitle Then X���L���v�V�����ݒ�
      .TickLabelPosition = xlTickLabelPositionLow
      .TickLabels.NumberFormatLocal = General
    End With
    Set a = ActiveChart.Axes(xlValue)
    With a
      If Not .HasTitle Then Y���L���v�V�����ݒ�
      .TickLabelPosition = xlTickLabelPositionLow
      .TickLabels.NumberFormatLocal = General
    End With
    
    Call �O���t�̔w�i�Ȃ���
    If �F�ݒ���s Then Call SeriesColorSet
    Call �ڐ����O��
    
'    ActiveChart.PlotArea.Select
    ActiveChart.ChartArea.Select

eee:
    Application.ScreenUpdating = True
End Sub


Sub �O���t�^�C�g���ݒ�()
  On Error GoTo eee
  Application.ScreenUpdating = False
  
  Dim str As String
  
  str = InputBox("�O���t�̃^�C�g������͂��Ă�������", "Graph Title", ActiveSheet.name) 'ActiveWorkbook.name)
  
  If Len(str) > 0 Then
    ActiveChart.HasTitle = True
    If Left(str, 1) = "=" And InStr(str, "!") = 0 Then
      Dim sht As Worksheet
      Set sht = ActiveSheet
      Dim new_str As String
      new_str = "='" & sht.name & "'!" & sht.Range(Mid(str, 1)).Address(True, True, xlR1C1)
      ActiveChart.ChartTitle.Text = new_str
    Else
      ActiveChart.ChartTitle.Text = str
    End If
  End If
    
eee:
  Application.ScreenUpdating = True



End Sub


Private Sub AddGraph()
'
' AddGraph Macro
' �}�N���L�^�� : 2011/1/27  ���[�U�[�� : �c��
'

'
    Charts.Add
    ActiveChart.ChartType = xlXYScatterLinesNoMarkers
    ActiveChart.SetSourceData source:=Sheets("ShellPushover").Range("A1:B12"), _
        PlotBy:=xlColumns
    ActiveChart.Location Where:=xlLocationAsObject, name:="ShellPushover"
    With ActiveChart
        .HasTitle = True
        .ChartTitle.Characters.Text = "Graph Title"
        .Axes(xlCategory, xlPrimary).HasTitle = True
        .Axes(xlCategory, xlPrimary).AxisTitle.Characters.Text = "X Title"
        .Axes(xlValue, xlPrimary).HasTitle = True
        .Axes(xlValue, xlPrimary).AxisTitle.Characters.Text = "Y Title"
    End With
End Sub

Private Function �t���[�g�z�u�ݒ�() As XlPlacement
  �t���[�g�z�u�ݒ� = xlMove
End Function

Sub A4�p�ɃT�C�Y�ύX()
  On Error GoTo eee
  Application.ScreenUpdating = False
  With GetChart(Selection)
    .Font.Size = 9
    .AutoScaleFont = False
    .Parent.Parent.Width = mm2pnt(170)
    .Parent.Parent.Height = mm2pnt(105)
    .Parent.Parent.Placement = �t���[�g�z�u�ݒ�()
    .Border.LineStyle = 0
  End With
eee:
  Application.ScreenUpdating = True
End Sub

Sub �O���t�V�[�g4�����p�ɃT�C�Y�ύX()
  On Error GoTo eee
  Application.ScreenUpdating = False
  With GetChart(Selection)
    .Font.Size = 8
    .AutoScaleFont = False
    .Parent.Parent.Width = mm2pnt(125)
    .Parent.Parent.Height = mm2pnt(78.5)
    .Parent.Parent.Placement = �t���[�g�z�u�ݒ�()
    .Border.LineStyle = 0
  End With
eee:
  Application.ScreenUpdating = True
End Sub

Sub A4�ŏc3�i�p�ɃT�C�Y�ύX()
  On Error GoTo eee
  Application.ScreenUpdating = False
  With GetChart(Selection)
    .Font.Size = 9
    .AutoScaleFont = False
    .Parent.Parent.Width = mm2pnt(170)
    .Parent.Parent.Height = 210 'pt
    .Parent.Parent.Placement = �t���[�g�z�u�ݒ�()
    .Border.LineStyle = 0
  End With
eee:
  Application.ScreenUpdating = True
End Sub

' ���50%�ɏk�����邱�Ƃ�O��
Sub A4�p��2�i�g�p�ɃT�C�Y�ύX()
  On Error GoTo eee
  Application.ScreenUpdating = False
  With GetChart(Selection)
    .Font.Size = 18
    .AutoScaleFont = False
    .Parent.Parent.Width = mm2pnt(160)
    .Parent.Parent.Height = mm2pnt(99)
    .Parent.Parent.Placement = �t���[�g�z�u�ݒ�()
    .Border.LineStyle = 0
  End With
eee:
  Application.ScreenUpdating = True
End Sub


Sub A3�p�ɃT�C�Y�ύX()
  On Error GoTo eee
  Application.ScreenUpdating = False
  With GetChart(Selection)
    .Font.Size = 9
    .AutoScaleFont = False
    .Parent.Parent.Width = mm2pnt(210)
    .Parent.Parent.Height = mm2pnt(170)
    .Parent.Parent.Placement = �t���[�g�z�u�ݒ�()
    .Border.LineStyle = 0
  End With
eee:
  Application.ScreenUpdating = True
End Sub



Sub �O���b�h��()
Attribute �O���b�h��.VB_ProcData.VB_Invoke_Func = "G\n14"
'
' �O���b�h�� Macro
' �}�N���L�^�� : 2011/3/26  ���[�U�[�� : �c��
'
'
  On Error GoTo eee
  Application.ScreenUpdating = False
    
    With ActiveChart.Axes(xlCategory)
        .HasMajorGridlines = True
        .HasMinorGridlines = True
    End With
    With ActiveChart.Axes(xlValue)
        .HasMajorGridlines = True
        .HasMinorGridlines = True
    End With
    With ActiveChart.Axes(xlValue).MinorGridlines.Border
        .ColorIndex = 15
        .weight = xlHairline
        .LineStyle = xlDot
    End With
    With ActiveChart.Axes(xlValue).MajorGridlines.Border
        .ColorIndex = 48
        .weight = xlHairline
        .LineStyle = xlContinuous
    End With
    With ActiveChart.Axes(xlCategory).MinorGridlines.Border
        .ColorIndex = 15
        .weight = xlHairline
        .LineStyle = xlDot
    End With
    With ActiveChart.Axes(xlCategory).MajorGridlines.Border
        .ColorIndex = 48
        .weight = xlHairline
        .LineStyle = xlContinuous
    End With
eee:
  Application.ScreenUpdating = True


End Sub


Sub �ڐ����O��()
Attribute �ڐ����O��.VB_ProcData.VB_Invoke_Func = "E\n14"

  With ActiveChart
    .Axes(xlCategory, xlPrimary).TickLabelPosition = xlTickLabelPositionLow
    .Axes(xlValue, xlPrimary).TickLabelPosition = xlTickLabelPositionLow
  End With

End Sub

Sub �����O��()
On Error Resume Next
  With ActiveChart
    .Axes(xlValue, xlPrimary).CrossesAt = .Axes(xlValue, xlPrimary).MinimumScale
    .Axes(xlCategory, xlPrimary).CrossesAt = .Axes(xlCategory, xlPrimary).MinimumScale
  End With
End Sub



Sub X���͈̔͂��k��()
   ���͈̔͂��k�� ActiveChart.Axes(xlCategory, xlPrimary), False
End Sub

Sub Y���͈̔͂��k��()
   ���͈̔͂��k�� ActiveChart.Axes(xlValue, xlPrimary), True
End Sub


Sub Y���̍ő�͈͂��k��()
   ���͈̔͂��k�� ActiveChart.Axes(xlValue, xlPrimary), False
End Sub

Sub X���̃��Z�b�g()
  ���̃��Z�b�g ActiveChart.Axes(xlCategory)
End Sub

Sub Y���̃��Z�b�g()
  ���̃��Z�b�g ActiveChart.Axes(xlValue)
End Sub

Sub �����̃��Z�b�g()
  X���̃��Z�b�g
  Y���̃��Z�b�g
End Sub


Private Sub ���͈̔͂��k��(ax As Axis, ByVal both As Boolean)
  ax.MajorUnitIsAuto = False

  ax.MaximumScale = ax.MaximumScale - ax.MajorUnit
  If both Then
    ax.MinimumScale = ax.MinimumScale + ax.MajorUnit
  End If
End Sub

Private Sub ���̃��Z�b�g(ax As Axis)
  ax.MaximumScaleIsAuto = True
  ax.MinimumScaleIsAuto = True
  ax.MajorUnitIsAuto = True
  ax.MinorUnitIsAuto = True
End Sub

Sub �O���tY�����x����]()
'
' �O���tY�����x����] Macro
' �}�N���L�^�� : 2012/2/5  ���[�U�[�� : -

    On Error Resume Next
    ActiveChart.Axes(xlValue).AxisTitle.Orientation = xlHorizontal
End Sub

Sub X���j���ڐ��ǉ�()
  �j���ڐ��ǉ� ActiveChart.Axes(xlCategory)
End Sub


Private Sub �j���ڐ��ǉ�(ax As Axis)
   ax.HasMajorGridlines = True
   ax.MajorGridlines.Border.weight = xlThin
   ax.MajorGridlines.Border.LineStyle = xlDash
End Sub

Sub �c�_�O���t�̐F�ݒ�()
'
' �c�_�O���t�̐F�ݒ� Macro
' �}�N���L�^�� : 2012/3/15  ���[�U�[�� : NSC999
'

On Error GoTo eee
    Application.ScreenUpdating = False

'
    ActiveChart.SeriesCollection(1).Select
    With Selection.Interior
        .ColorIndex = 56 ' �D�F
        .Pattern = xlSolid
    End With
    ActiveChart.SeriesCollection(2).Select
    With Selection.Interior
        .ColorIndex = 10 ' ��
        .Pattern = xlSolid
    End With
    ActiveChart.SeriesCollection(3).Select
    With Selection.Interior
        .ColorIndex = 43 ' ���C��
        .Pattern = xlSolid
    End With
    ActiveChart.SeriesCollection(4).Select
    With Selection.Interior
        .ColorIndex = 5  ' ��
        .Pattern = xlSolid
    End With
    ActiveChart.SeriesCollection(5).Select
    With Selection.Interior
        .ColorIndex = 37 ' �x�[���u���[
        .Pattern = xlSolid
    End With

eee:
    Application.ScreenUpdating = True

End Sub


Sub kiloRangeY()
Attribute kiloRangeY.VB_ProcData.VB_Invoke_Func = " \n14"
'
' kiloRangeY Macro
' �}�N���L�^�� : 2011/4/7  ���[�U�[�� : �c��

    With ActiveChart.Axes(xlValue)
        .DisplayUnit = xlThousands
        .HasDisplayUnitLabel = False
    End With
End Sub



Public Sub �n��̓h��Ԃ�����ݒ�()
  �_�O���t�F�ݒ�.Show vbModeless
End Sub


Public Sub �n��̓h��Ԃ����擾()
On Error GoTo eee

Dim info()
Dim s As Series
Dim ff As ChartFillFormat
Dim n As Integer
Dim i As Integer

n = ActiveChart.SeriesCollection.Count
ReDim info(0 To n)
info(0) = n
  
For i = 1 To n
  Set s = ActiveChart.SeriesCollection(i)
  Set ff = s.Fill
  Dim t, fc, bc, arg1, arg2
  t = ff.Type
  fc = ff.ForeColor.SchemeColor
  bc = ff.BackColor.SchemeColor
  Select Case t
  Case msoFillGradient
    ' �O���f�[�V����
    Dim a(0 To 4)
    a(0) = ff.GradientColorType
    a(1) = ff.GradientDegree
    a(2) = ff.GradientStyle
    a(3) = ff.GradientVariant
    a(4) = ff.PresetGradientType
    arg1 = 5
    arg2 = Join(a, ":")
  Case msoFillMixed
     ' ???
    arg1 = 0
    arg2 = 0
  Case msoFillPatterned
    '�p�^�[��
    arg1 = ff.Pattern
    arg2 = 0
  Case msoFillPicture
     arg1 = 0
     arg2 = 0
  Case msoFillSolid
     arg1 = 0
     arg2 = 0
  Case msoFillTextured
     arg1 = ff.TextureType
     If arg1 = msoTexturePreset Then
       arg2 = ff.PresetTexture
     Else
       arg2 = ff.TextureName
     End If
  End Select
  info(i) = Join(Array(t, fc, bc, arg1, arg2), ",")
Next


InputBox "�R�s�[���Ă�������", "�h��Ԃ���񕶎���", Join(info, ";")


eee:

End Sub

Sub �f�[�^���x����l��()
    ActiveChart.ApplyDataLabels AutoText:=True, LegendKey:=False, _
        HasLeaderLines:=False, ShowSeriesName:=False, ShowCategoryName:=False, _
        ShowValue:=True, ShowPercentage:=False, ShowBubbleSize:=False
End Sub


Sub �㔼�̌n����f�[�^�\���p�ɕύX()
'
' Macro6 Macro
' �}�N���L�^�� : 2012/4/13  ���[�U�[�� : NSC999
'
    Dim n As Integer
    n = ActiveChart.SeriesCollection.Count
'    Dim start As Integer
'    start = InputBox("�S����" & n & "�̌n�񂪂���܂��D" & vbCrLf _
'    & "���Ԉȍ~��l�\���p�ɂ��܂����H", "�ϊ��J�n�ʒu�̑I��")
'
'    If start = "" Then Exit Sub
    'On Error Resume Next
    Application.ScreenUpdating = False
    
    If ActiveChart.HasLegend Then
        ' reset legend
        ActiveChart.HasLegend = False
        ActiveChart.HasLegend = True
    End If
    
    Dim s As Series
    Dim i As Integer
    For i = n To n / 2 + 1 Step -1
      Set s = ActiveChart.SeriesCollection(i)
        s.Border.LineStyle = xlNone
        s.Interior.ColorIndex = xlNone
        s.AxisGroup = 2
        s.ApplyDataLabels AutoText:=True, LegendKey:= _
            False, ShowSeriesName:=False, ShowCategoryName:=False, ShowValue:=True, _
            ShowPercentage:=False, ShowBubbleSize:=False
        With s.DataLabels
            .HorizontalAlignment = xlCenter
            .VerticalAlignment = xlCenter
            .ReadingOrder = xlContext
            .Position = xlLabelPositionInsideBase
            .Orientation = xlHorizontal
            .Font.Background = xlOpaque
        End With
'        If ActiveChart.HasLegend Then
'          Dim e As LegendEntry
'          Set e = ActiveChart.Legend.LegendEntries(i)
'          If Not e Is Nothing Then e.Delete
'        End If
    Next
    If ActiveChart.HasLegend Then
    On Error GoTo kkk
        Do Until ActiveChart.Legend.LegendEntries().Count <= n / 2
            ActiveChart.Legend.LegendEntries(ActiveChart.Legend.LegendEntries().Count).Delete
        Loop
kkk:
    On Error GoTo eee
    End If
    With ActiveChart.Axes(xlValue, xlSecondary)
        .MajorTickMark = xlNone
        .MinorTickMark = xlNone
        .TickLabelPosition = xlNone
    End With
    ActiveChart.ChartGroups(2).GapWidth = ActiveChart.ChartGroups(1).GapWidth
    ActiveChart.ChartGroups(2).HasSeriesLines = False
eee:
    Application.ScreenUpdating = True
End Sub

Sub �㔼�̌n��͖}��Ȃ���()
    On Error GoTo kkk
    If ActiveChart.HasLegend Then
        Dim n As Integer
        n = ActiveChart.SeriesCollection.Count
        Do Until ActiveChart.Legend.LegendEntries().Count <= n / 2
            ActiveChart.Legend.LegendEntries(ActiveChart.Legend.LegendEntries().Count).Delete
        Loop
    End If
kkk:
End Sub


Sub �㔼�̌n����2���Q�Ƃ�()
'
' Macro6 Macro
' �}�N���L�^�� : 2012/4/13  ���[�U�[�� : NSC999
'
    Dim n As Integer
    n = ActiveChart.SeriesCollection.Count
'    Dim start As Integer
'    start = InputBox("�S����" & n & "�̌n�񂪂���܂��D" & vbCrLf _
'    & "���Ԉȍ~��l�\���p�ɂ��܂����H", "�ϊ��J�n�ʒu�̑I��")
'
'    If start = "" Then Exit Sub
    'On Error Resume Next
    Application.ScreenUpdating = False
    
    Dim s As Series
    Dim i As Integer
    For i = n To n / 2 + 1 Step -1
      Set s = ActiveChart.SeriesCollection(i)
        s.AxisGroup = 2
    Next
eee:
    Application.ScreenUpdating = True
End Sub

Sub ��2�����1���Ɠ�����()
  
  Dim f As Axis, s As Axis
  Set f = ActiveChart.Axes(xlValue, xlPrimary)
  Set s = ActiveChart.Axes(xlValue, xlSecondary)

  
  If f.MaximumScale > s.MaximumScale Then
    s.MaximumScale = f.MaximumScale
  Else
    f.MaximumScale = s.MaximumScale
  End If
  If f.MinimumScale < s.MinimumScale Then
    s.MinimumScale = f.MinimumScale
  Else
    f.MinimumScale = s.MinimumScale
  End If
  s.HasTitle = False
'  s.HasDisplayUnitLabel = False
  s.Delete
  
End Sub

Sub �㔼�̖_�O���t�\����O���Ɠ�����()
  Dim i As Integer
  Dim p As Series, s As Series
  Dim n As Integer
  
  n = ActiveChart.SeriesCollection.Count / 2
  For i = 1 To n
    Set p = ActiveChart.SeriesCollection(i)
    Set s = ActiveChart.SeriesCollection(i + n)
    If p.Fill.Visible Then
      s.Fill.BackColor.SchemeColor = p.Fill.BackColor.SchemeColor
      s.Fill.ForeColor.SchemeColor = p.Fill.ForeColor.SchemeColor
      Select Case p.Fill.Type
        Case msoFillGradient
            MsgBox "�O���f�[�V�����͖��Ή��ł�"
        Case msoFillMixed
            MsgBox "Mixed�͖��Ή��ł�"
        Case msoFillPatterned
            s.Fill.Patterned p.Fill.Pattern
        Case msoFillPicture
            MsgBox "�s�N�`���͖��Ή��ł�"
        Case msoFillSolid
            s.Fill.Solid
        Case msoFillTextured
            If p.Fill.TextureType = msoTexturePreset Then
                s.Fill.PresetTextured p.Fill.PresetTexture
            Else
                s.Fill.UserTextured p.Fill.TextureName
            End If
      End Select
    End If
    s.Fill.Visible = p.Fill.Visible
    
    s.Border.LineStyle = p.Border.LineStyle
    If p.Border.LineStyle <> xlLineStyleNone Then
        s.Border.ColorIndex = p.Border.ColorIndex
        s.Border.weight = p.Border.weight
    End If
  Next

End Sub


Sub �f�[�^���x�������T�C�Y�ύX(Optional sz As Double = 0#)
'Dim sz As Double
Dim s As Series

If sz = 0# Then
    sz = val(InputBox("�|�C���g�T�C�Y�H", "�f�[�^���x���̕����T�C�Y�ύX", "8"))
End If

On Error GoTo eee
Application.ScreenUpdating = False

For Each s In ActiveChart.SeriesCollection
  If s.HasDataLabels Then
      s.DataLabels.Font.Size = sz
  End If
Next

eee:
Application.ScreenUpdating = True

End Sub

Sub �f�[�^���x�������w�i�ύX()
Dim bg
Dim s As Series

For Each s In ActiveChart.SeriesCollection
  If s.HasDataLabels Then
    If s.DataLabels.Font.Background = xlOpaque Then
      bg = xlTransparent
    Else
      bg = xlOpaque
    End If
    Exit For
  End If
Next

' �ЂƂ��f�[�^���x�����Ȃ�������Ȃɂ����Ȃ�
If IsEmpty(bg) Then Exit Sub

On Error GoTo eee
Application.ScreenUpdating = False

For Each s In ActiveChart.SeriesCollection
  If s.HasDataLabels Then
      s.DataLabels.Font.Background = bg
  End If
Next

eee:
Application.ScreenUpdating = True

End Sub


Sub �f�[�^���x�����������ύX()
Dim ori
Dim s As Series

For Each s In ActiveChart.SeriesCollection
  If s.HasDataLabels Then
    If s.DataLabels.Orientation = xlHorizontal Then
      ori = xlUpward
    Else
      ori = xlHorizontal
    End If
    Exit For
  End If
Next

' �ЂƂ��f�[�^���x�����Ȃ�������Ȃɂ����Ȃ�
If IsEmpty(ori) Then Exit Sub

On Error GoTo eee
Application.ScreenUpdating = False

For Each s In ActiveChart.SeriesCollection
  If s.HasDataLabels Then
      s.DataLabels.Orientation = ori
  End If
Next

eee:
Application.ScreenUpdating = True

End Sub

Sub �}������()
'
' �}������ Macro
' �}�N���L�^�� : 2012/4/17  ���[�U�[�� : NSC999
'
    ActiveChart.Legend.Position = xlTop
End Sub

Sub �~�O���t�ݒ�()
'
' Macro3 Macro
' �}�N���L�^�� : 2012/4/17  ���[�U�[�� : NSC999
'

'
    If ActiveChart Is Nothing Then Exit Sub
    If ActiveChart.HasLegend Then ActiveChart.Legend.Delete
    ActiveChart.SeriesCollection(1).ApplyDataLabels AutoText:=True, LegendKey:= _
        False, HasLeaderLines:=True, ShowSeriesName:=False, ShowCategoryName:= _
        True, ShowValue:=True, ShowPercentage:=True, ShowBubbleSize:=False, _
        Separator:="" & Chr(10) & ""
    With ActiveChart.SeriesCollection(1).DataLabels
        .HorizontalAlignment = xlCenter
        .VerticalAlignment = xlCenter
        .ReadingOrder = xlContext
        .Position = xlLabelPositionCenter
        .Orientation = xlHorizontal
        .AutoScaleFont = False
        .Font.Background = xlOpaque
    End With
    If ActiveChart.HasTitle Then ActiveChart.ChartTitle.Delete
    ActiveChart.Parent.Select
End Sub


Sub �f�[�^���x���n�񖼒ǉ�()
'
' �f�[�^���x�����ږ��ǉ� Macro
' �}�N���L�^�� : 2012/4/17  ���[�U�[�� : NSC999
'

'
    'ActiveChart.SeriesCollection(2).Points (2)
    Selection.ApplyDataLabels AutoText:=True, _
        LegendKey:=False, ShowSeriesName:=True, ShowCategoryName:=False, _
        ShowValue:=False, ShowPercentage:=False, ShowBubbleSize:=False
End Sub

Sub �f�[�^���x���n�񖼂ƒl�̒ǉ�()
'
' �f�[�^���x���n�񖼂ƒl�̒ǉ� Macro
' �}�N���L�^�� : 2012/4/17  ���[�U�[�� : NSC999
'

'
'    ActiveSheet.ChartObjects("�O���t 5").Activate
'    ActiveChart.SeriesCollection(2).Select
'    ActiveChart.SeriesCollection(2).Points(1).Select
  On Error Resume Next
    Selection.ApplyDataLabels AutoText:=True, _
        LegendKey:=False, ShowSeriesName:=True, ShowCategoryName:=False, _
        ShowValue:=True, ShowPercentage:=False, ShowBubbleSize:=False, Separator _
        :="" & Chr(10) & ""
End Sub

Sub �ڐ����x���Ԋu��1��()
'
' �ڐ��Ԋu��1�� Macro
' �}�N���L�^�� : 2012/4/17  ���[�U�[�� : NSC999
'

'
    With ActiveChart.Axes(xlCategory)
'        .Crosses = xlMaximum
        .TickLabelSpacing = 1
        .TickMarkSpacing = 1
'        .AxisBetweenCategories = True
'        .ReversePlotOrder = True
    End With
End Sub

Sub ����2�{()

    If ActiveChart Is Nothing Then Exit Sub
    If TypeName(ActiveChart.Parent) = "Workbook" Then Exit Sub
    ActiveChart.Parent.Height = ActiveChart.Parent.Height * 2
End Sub

Sub copyChartToClipboadWithOriginalScale()
Attribute copyChartToClipboadWithOriginalScale.VB_ProcData.VB_Invoke_Func = " \n14"
  If ActiveChart Is Nothing Then Exit Sub
  Dim co As ChartObject
 
  Set co = ActiveChart.Parent
 
  co.TopLeftCell.Select
'  ActiveWindow.Selection = xlNone
  ActiveWindow.Zoom = 100
  co.Activate
  co.Chart.ChartArea.Select
  co.Chart.ChartArea.Copy
End Sub


Private Function SetMarkerTypeWithCount(Optional same As Integer = 1)
    Dim s As Series
    Dim i As Integer
    Dim shape
    Dim k As Integer
    shape = Array(xlPlus, xlCircle, xlSquare, xlTriangle, xlDiamond, xlCross, xlStar)
    Const n As Integer = 7
    
    For i = 1 To ActiveChart.SeriesCollection.Count
        Set s = ActiveChart.SeriesCollection(i)
        k = Int((i - 1) / same) + 1
        s.MarkerStyle = shape(k Mod n)
    Next
    SetMarkerTypeWithCount = True
End Function

Sub �}�[�J�[�ݒ�()
    Dim res As String
    res = InputBox("�K�p����(���������~���{)�ł��D" & vbCrLf & "�����}�[�J�[������g���܂����H", Default:="1")
    If res = "" Then Exit Sub
    SetMarkerTypeWithCount CInt(res)
End Sub


Sub �}�[�J�[���ݓh��()
    Dim i As Integer
    Dim s As Series
    For i = 1 To ActiveChart.SeriesCollection.Count
        Set s = ActiveChart.SeriesCollection(i)
        If i Mod 2 = 0 Then
            s.MarkerBackgroundColorIndex = 2
        Else
            Select Case s.MarkerStyle
            Case xlCircle, xlSquare, xlTriangle, xlDiamond
                s.MarkerBackgroundColorIndex = s.MarkerForegroundColorIndex
            Case Else
                s.MarkerBackgroundColorIndex = xlColorIndexNone
            End Select
        End If
    Next
End Sub
Sub �}�[�J�[���ݓh��_����()
    Dim i As Integer
    Dim s As Series
    For i = 1 To ActiveChart.SeriesCollection.Count
        Set s = ActiveChart.SeriesCollection(i)
        If i Mod 2 = 1 Then
            s.MarkerBackgroundColorIndex = 2
        Else
            Select Case s.MarkerStyle
            Case xlCircle, xlSquare, xlTriangle, xlDiamond
                s.MarkerBackgroundColorIndex = s.MarkerForegroundColorIndex
            Case Else
                s.MarkerBackgroundColorIndex = xlColorIndexNone
            End Select
        End If
    Next
End Sub



Private Sub remove_last_line(ByRef s As Series)
  Dim i  As Integer
  Dim n As Integer
  i = s.Points().Count
  Do While i > 0
    Dim p As Point
    Set p = s.Points(i)
    If p.Border.ColorIndex <> xlColorIndexNone Then
      p.Border.ColorIndex = xlColorIndexNone
      Exit Sub
    End If
    i = i - 1
  Loop
  
End Sub

Sub �n��̐����Ōォ�珇�Ɏ�菜��()
 
  If TypeName(Selection) = "Series" Then
    remove_last_line Selection
  Else
    Dim s As Series
    Application.ScreenUpdating = False
    For Each s In ActiveChart.SeriesCollection
      remove_last_line s
    Next
    Application.ScreenUpdating = True
  End If
End Sub

Sub RotateMajorCategory()
    Dim ax As Axis
    Set ax = ActiveChart.Axes(xlCategory)
    
    ax.TickLabels.Orientation = xlTickLabelOrientationHorizontal
    
    

End Sub


Sub �܂���ݒ�()
'
' �܂���ݒ� Macro
' �}�N���L�^�� : 2012/4/26  ���[�U�[�� : NSC999
'
    Dim setting
    
    setting = Array( _
        Array(xlCircle, 5, False), _
        Array(xlCircle, 5, True), _
        Array(xlTriangle, 5, False), _
        Array(xlTriangle, 5, True), _
        Array(xlSquare, 5, False), _
        Array(xlSquare, 5, True), _
        Null)

    

'
    ActiveSheet.ChartObjects("�O���t 46").Activate
    ActiveChart.Legend.Select
    ActiveChart.Legend.LegendEntries(1).LegendKey.Select
    With Selection.Border
        .ColorIndex = 1
        .weight = xlThin
        .LineStyle = xlContinuous
    End With
    With Selection
        .MarkerBackgroundColorIndex = 2
        .MarkerForegroundColorIndex = 1
        .MarkerStyle = xlCircle
        .Smooth = False
        .MarkerSize = 5
        .Shadow = False
    End With
    ActiveChart.Legend.LegendEntries(2).LegendKey.Select
    With Selection.Border
        .ColorIndex = 1
        .weight = xlThin
        .LineStyle = xlContinuous
    End With
    With Selection
        .MarkerBackgroundColorIndex = 1
        .MarkerForegroundColorIndex = 1
        .MarkerStyle = xlCircle
        .Smooth = False
        .MarkerSize = 5
        .Shadow = False
    End With
    ActiveChart.Legend.LegendEntries(3).LegendKey.Select
    With Selection.Border
        .ColorIndex = 1
        .weight = xlThin
        .LineStyle = xlContinuous
    End With
    With Selection
        .MarkerBackgroundColorIndex = 2
        .MarkerForegroundColorIndex = 1
        .MarkerStyle = xlTriangle
        .Smooth = False
        .MarkerSize = 5
        .Shadow = False
    End With
    ActiveChart.Legend.LegendEntries(4).LegendKey.Select
    With Selection.Border
        .ColorIndex = 1
        .weight = xlThin
        .LineStyle = xlContinuous
    End With
    With Selection
        .MarkerBackgroundColorIndex = 1
        .MarkerForegroundColorIndex = 1
        .MarkerStyle = xlTriangle
        .Smooth = False
        .MarkerSize = 5
        .Shadow = False
    End With
    ActiveChart.Legend.LegendEntries(5).LegendKey.Select
    With Selection.Border
        .ColorIndex = 1
        .weight = xlThin
        .LineStyle = xlContinuous
    End With
    With Selection
        .MarkerBackgroundColorIndex = 2
        .MarkerForegroundColorIndex = 1
        .MarkerStyle = xlSquare
        .Smooth = False
        .MarkerSize = 5
        .Shadow = False
    End With
    ActiveChart.Legend.LegendEntries(6).LegendKey.Select
    With Selection.Border
        .ColorIndex = 1
        .weight = xlThin
        .LineStyle = xlContinuous
    End With
    With Selection
        .MarkerBackgroundColorIndex = 1
        .MarkerForegroundColorIndex = 1
        .MarkerStyle = xlSquare
        .Smooth = False
        .MarkerSize = 5
        .Shadow = False
    End With
    ActiveChart.Legend.LegendEntries(7).Select
    ActiveChart.Legend.LegendEntries(7).LegendKey.Select
    With Selection.Border
        .ColorIndex = 1
        .weight = xlThin
        .LineStyle = xlContinuous
    End With
    With Selection
        .MarkerBackgroundColorIndex = 2
        .MarkerForegroundColorIndex = 1
        .MarkerStyle = xlDiamond
        .Smooth = False
        .MarkerSize = 5
        .Shadow = False
    End With
    ActiveChart.Legend.LegendEntries(8).LegendKey.Select
    With Selection.Border
        .ColorIndex = 1
        .weight = xlThin
        .LineStyle = xlContinuous
    End With
    With Selection
        .MarkerBackgroundColorIndex = 1
        .MarkerForegroundColorIndex = 1
        .MarkerStyle = xlDiamond
        .Smooth = False
        .MarkerSize = 5
        .Shadow = False
    End With
    ActiveChart.Legend.LegendEntries(9).LegendKey.Select
    With Selection.Border
        .ColorIndex = 1
        .weight = xlThin
        .LineStyle = xlContinuous
    End With
    With Selection
        .MarkerBackgroundColorIndex = 2
        .MarkerForegroundColorIndex = 1
        .MarkerStyle = xlX
        .Smooth = False
        .MarkerSize = 5
        .Shadow = False
    End With
End Sub

Sub �ő�ŏ��_�O���t�ݒ�()

  '�O���t�̊�{�ݒ� ����g�������Ȃ��̂��R�s�[
  Dim str As String
  Dim a As Axis
  On Error GoTo eee
  Application.ScreenUpdating = False

  Set a = ActiveChart.Axes(xlValue)
  With a
    If Not .HasTitle Then Y���L���v�V�����ݒ�
    .TickLabelPosition = xlTickLabelPositionLow
    .TickLabels.NumberFormatLocal = "General"
  End With
  Call �O���t�̔w�i�Ȃ���
  Call �ڐ����O��
  ActiveChart.ChartArea.Select
  
  ' �ȍ~�͊�{�ݒ�ɖ�������
  Call A4�p�ɃT�C�Y�ύX
  �_�O���t�F�ݒ�.applyBarFill �_�O���t�F�ݒ�.�����p�^�[��5()
  Call �㔼�̌n����2���Q�Ƃ�
  Call ��2�����1���Ɠ�����
  Call �㔼�̖_�O���t�\����O���Ɠ�����
  Call �㔼�̌n��͖}��Ȃ���
  Call �}������
    
eee:
    Application.ScreenUpdating = True

End Sub




Private Function GetChart(ByRef arg As Object) As Chart
  Dim o As Object
  Set o = arg
  Set GetChart = Nothing
  Do Until (TypeName(o) = "Application")
    If TypeName(o) = "Chart" Then
      Set GetChart = o
      Exit Function
    End If
    Set o = o.Parent
  Loop

End Function

Private Function GetChartObject(ByRef arg As Object) As ChartObject
  Dim o As Object
  Set o = arg
  Set GetChartObject = Nothing
  Do Until (TypeName(o) = "Application")
    If TypeName(o) = "ChartObject" Then
      Set GetChartObject = o
      Exit Function
    End If
    Set o = o.Parent
  Loop

End Function



Sub copy_and_change_data_column()
  Dim co As ChartObject, ori As ChartObject
  Dim sht As Worksheet
  Dim r As Range, TL As Range, BR As Range
  
  ' Copy Selected Chart
  Set sht = ActiveSheet
  Set ori = GetChartObject(Selection)
  
  ' Validation
  If ori Is Nothing Then Exit Sub
  If ori.Chart.SeriesCollection.count <> 1 Then Exit Sub
  
  ori.Copy
  Set TL = ori.TopLeftCell
  Set BR = ori.BottomRightCell
  sht.Cells(BR.row, TL.Column).Select
  sht.Paste
  Set co = GetChartObject(Selection)
  co.Top = ori.Top + ori.Height
  co.Left = ori.Left
  
  ' move to right column
  Dim s As Series
  Set s = co.Chart.SeriesCollection(1)
  Dim fm As String
  fm = s.FormulaR1C1
  Dim a, b, c, d, e, f, g
  a = Split(fm, ",")
  
  b = Split(a(2), "!")
  c = Split(b(1), ":")
  d = Split(c(0), "C")
  e = Split(c(1), "C")
  
  d(1) = Format(val(d(1)) + 1, "0")
  e(1) = Format(val(e(1)) + 1, "0")
  
  c(1) = Join(e, "C")
  c(0) = Join(d, "C")
  b(1) = Join(c, ":")
  a(2) = Join(b, "!")
  
  ' change caption
  f = Split(a(0), "!")
  g = Split(f(1), "C")
  g(1) = Format(val(g(1)) + 1, "0")
  f(1) = Join(g, "C")
  a(0) = Join(f, "!")
  
  s.FormulaR1C1 = Join(a, ",")

End Sub


Sub use_current_sheet_data()
  Dim co As ChartObject, ori As ChartObject
  Dim sht As Worksheet
  Dim r As Range, TL As Range, BR As Range
  
  ' Copy Selected Chart
  Set sht = ActiveSheet
  Set co = GetChartObject(Selection)
  
  ' Validation
  If co Is Nothing Then Exit Sub
  If co.Chart.SeriesCollection.count <> 1 Then Exit Sub
  
  ' Change sheet name
  Dim s As Series
  Set s = co.Chart.SeriesCollection(1)
  Dim fm As String
  fm = s.FormulaR1C1
  Dim a, b, c, d, e, f, g
  a = Split(fm, ",")
  b = Split(a(0), "(")
  c = Split(b(1), "!")
  c(0) = "'" & sht.name & "'"
  b(1) = Join(c, "!")
  a(0) = Join(b, "(")
  
  d = Split(a(1), "!")
  d(0) = c(0)
  a(1) = Join(d, "!")
  
  e = Split(a(2), "!")
  e(0) = c(0)
  a(2) = Join(e, "!")
  
  s.FormulaR1C1 = Join(a, ",")
  
  co.Chart.ChartTitle.Text = sht.name
  
  'exit sub
  Dim num As Integer
  num = sht.Range("A1").End(xlToRight).Column - 2
  Dim i
  For i = 1 To num
    Call copy_and_change_data_column
  Next
  co.Activate
  co.Select True
  co.Copy
End Sub


