Attribute VB_Name = "Temporary"
Option Explicit


Sub ���W�����v()
  JumpList.Show vbModeless
End Sub

Sub �A���P�[�g���_100�O���t�쐬()
'
' �A���P�[�g�񓚗����_�O���t�쐬 Macro
' �}�N���L�^�� : 2012/4/20  ���[�U�[�� : NSC999
'
    Dim rng As Range, sht As Worksheet
    Set sht = ActiveSheet
    Set rng = ActiveWindow.Selection


    Charts.Add
    ActiveChart.ChartType = xlBarStacked100
    ActiveChart.SetSourceData source:=rng, PlotBy:=xlRows
    ActiveChart.Location Where:=xlLocationAsObject, name:=sht.name
    
    �_�O���t�F�ݒ�.applyBarFill �_�O���t�F�ݒ�.�����p�^�[��5()
    �}������
    Call A4�ŏc3�i�p�ɃT�C�Y�ύX
    Call AAA_enquete_graph(True)
    With ActiveChart.Axes(xlValue)
        .TickLabels.NumberFormatLocal = "0%"
        .HasTitle = True
        .AxisTitle.Characters.Text = "�񓚂̑I��"
    End With
    �ڐ����x���Ԋu��1��
    �f�[�^���x����l��
    
    If rng.Columns.Count = 2 Then
      Dim ax As Axis
      Set ax = ActiveChart.Axes(xlCategory)
      ax.TickLabels.Delete
    Else
      ActiveChart.HasTitle = True
      ActiveChart.ChartTitle.Text = rng.Range("A1").Value
    End If
    
'    Call �f�[�^���x�������w�i�ύX
    ActiveChart.ChartArea.Select


End Sub


Sub �A���P�[�g�I�𗦃O���t�쐬()
'
' Macro3 Macro
' �}�N���L�^�� : 2012/4/17  ���[�U�[�� : NSC999
'
'
    Dim rng As Range, sht As Worksheet
    Set sht = ActiveSheet
    Set rng = ActiveWindow.Selection
    Dim n As Integer, i  As Integer
    n = rng.Columns.Count
    If n < 3 Then
        MsgBox "�I��͈̗͂񐔂��s�����Ă��܂��D"
        Exit Sub
    End If
    
    Charts.Add
    ActiveChart.ChartType = xlBarClustered
'    ActiveChart.SetSourceData source:=Sheets("wakate-res_-1").Range("C1310:E1314" _
        ), PlotBy:=xlColumns
    ActiveChart.SetSourceData source:=rng, PlotBy:=xlColumns
    ActiveChart.Location Where:=xlLocationAsObject, name:=sht.name
    
    n = ActiveChart.ChartGroups(1).SeriesCollection().Count
    If n Mod 2 = 1 Then
        If MsgBox("�n�񐔂������ɂȂ��Ă��܂���D�폜���܂��D", vbOKCancel) = vbOK Then _
            ActiveChart.Parent.Delete
        Exit Sub
    End If
    
    
    ' ���ɑ���
    For i = 1 To n / 2
        ActiveChart.ChartGroups(1).SeriesCollection(1).PlotOrder = n
    Next
    ' �񐔂��Q�Ȃ�}��͕s�v
    If n = 2 Then
        ActiveChart.Legend.Delete
        With ActiveChart.ChartGroups(1).SeriesCollection(1).Interior
            .ColorIndex = 48
            .Pattern = xlSolid
        End With
    Else
       �_�O���t�F�ݒ�.applyBarFill �_�O���t�F�ݒ�.�����p�^�[��20()
    End If
    Call A4�ŏc3�i�p�ɃT�C�Y�ύX
    Call �㔼�̌n����f�[�^�\���p�ɕύX
    Call AAA_enquete_graph(False)
    With ActiveChart.Axes(xlValue)
        .TickLabels.NumberFormatLocal = "0%"
        .HasTitle = True
        .AxisTitle.Characters.Text = "�񓚂̑I��"
    End With
    Call �f�[�^���x�������w�i�ύX
    ActiveChart.ChartArea.Select
End Sub

Sub �A���P�[�g�s�I�𗦃O���t�쐬()
'
' Macro3 Macro
' �}�N���L�^�� : 2012/4/17  ���[�U�[�� : NSC999
'
'
    Dim rng As Range, sht As Worksheet
    Set sht = ActiveSheet
    Set rng = ActiveWindow.Selection
    Dim n As Integer, i  As Integer
    n = rng.Rows.Count
    If n < 3 Then
        MsgBox "�I��͈͂̍s�����s�����Ă��܂��D"
        Exit Sub
    End If
    
    Charts.Add
    ActiveChart.ChartType = xlBarClustered
    ActiveChart.SetSourceData source:=rng, PlotBy:=xlRows
    ActiveChart.Location Where:=xlLocationAsObject, name:=sht.name
    
    n = ActiveChart.ChartGroups(1).SeriesCollection().Count
    If n Mod 2 = 1 Then
        If MsgBox("�n�񐔂������ɂȂ��Ă��܂���D�폜���܂��D", vbOKCancel) = vbOK Then _
            ActiveChart.Parent.Delete
        Exit Sub
    End If
    
'    ' ���ɑ���
'    For i = 1 To n / 2
'        ActiveChart.ChartGroups(1).SeriesCollection(1).PlotOrder = n
'    Next

    ' ���Ԃ𔽓]������
    For i = 1 To (n - 1)
       ActiveChart.ChartGroups(1).SeriesCollection(n).PlotOrder = i
    Next

    If n = 2 Then
        ' �񐔂��Q�Ȃ�}��͕s�v
        ActiveChart.Legend.Delete
        With ActiveChart.ChartGroups(1).SeriesCollection(1).Interior
            .ColorIndex = 48
            .Pattern = xlSolid
        End With
    Else
       �_�O���t�F�ݒ�.applyBarFill �_�O���t�F�ݒ�.�����p�^�[��20()
    End If
    Dim items As Integer
    items = n / 2 * rng.Columns.Count
    If items > 20 Then
        Call A4�p�ɃT�C�Y�ύX
    Else
        Call A4�ŏc3�i�p�ɃT�C�Y�ύX
    End If
    Call �㔼�̌n����f�[�^�\���p�ɕύX
    Call AAA_enquete_graph(False, False)
    With ActiveChart.Axes(xlValue)
        .TickLabels.NumberFormatLocal = "0%"
        .HasTitle = True
        .AxisTitle.Characters.Text = "�񓚂̑I��"
    End With
    If n = 2 Then
        Call �f�[�^���x�������w�i�ύX
    End If
    Call �f�[�^���x�������T�C�Y�ύX(8#)
    ActiveChart.ChartArea.Select
End Sub

Sub �A���P�[�g�I�𗦏c�_�O���t�쐬()
'
' Macro3 Macro
' �}�N���L�^�� : 2012/4/17  ���[�U�[�� : NSC999
'
'
    Dim rng As Range, sht As Worksheet
    Set sht = ActiveSheet
    Set rng = ActiveWindow.Selection
    Dim n As Integer, i  As Integer
    n = rng.Columns.Count
    If n < 3 Then
        MsgBox "�I��͈̗͂񐔂��s�����Ă��܂��D"
        Exit Sub
    End If
    
    Charts.Add
    ActiveChart.ChartType = xlColumnClustered
'    ActiveChart.SetSourceData source:=Sheets("wakate-res_-1").Range("C1310:E1314" _
        ), PlotBy:=xlColumns
    ActiveChart.SetSourceData source:=rng, PlotBy:=xlColumns
    ActiveChart.Location Where:=xlLocationAsObject, name:=sht.name
    
    n = ActiveChart.ChartGroups(1).SeriesCollection().Count
    
    If n Mod 2 = 1 Then
        If MsgBox("�n�񐔂������ɂȂ��Ă��܂���D�폜���܂��D", vbOKCancel) = vbOK Then _
            ActiveChart.Parent.Delete
        Exit Sub
    End If
    
    
    ' ���ɑ���
    For i = 1 To n / 2
        ActiveChart.ChartGroups(1).SeriesCollection(1).PlotOrder = n
    Next
    ' �񐔂��Q�Ȃ�}��͕s�v
    If n = 2 Then
        ActiveChart.Legend.Delete
        With ActiveChart.ChartGroups(1).SeriesCollection(1).Interior
            .ColorIndex = 48
            .Pattern = xlSolid
        End With
    Else
       �_�O���t�F�ݒ�.applyBarFill �_�O���t�F�ݒ�.�����p�^�[��20()
    End If
    Call A4�ŏc3�i�p�ɃT�C�Y�ύX
    Call �㔼�̌n����f�[�^�\���p�ɕύX
    Call AAA_enquete_graph(False)
    With ActiveChart.Axes(xlValue)
        .TickLabels.NumberFormatLocal = "0%"
        .HasTitle = True
        .AxisTitle.Characters.Text = "�񓚂̑I��"
        .AxisTitle.Orientation = xlVertical
    End With
    If n = 2 Then Call �f�[�^���x�������w�i�ύX
    ActiveChart.ChartArea.Select
End Sub

Sub �A���P�[�g�s�I�𗦏c�_�O���t�쐬()
'
' Macro3 Macro
' �}�N���L�^�� : 2012/4/17  ���[�U�[�� : NSC999
'
'
    Dim rng As Range, sht As Worksheet
    Set sht = ActiveSheet
    Set rng = ActiveWindow.Selection
    Dim n As Integer, i  As Integer
    n = rng.Rows.Count
    If n < 3 Then
        MsgBox "�I��͈͂̍s�����s�����Ă��܂��D"
        Exit Sub
    End If

    
    Charts.Add
    ActiveChart.ChartType = xlColumnClustered
    ActiveChart.SetSourceData source:=rng, PlotBy:=xlRows
    ActiveChart.Location Where:=xlLocationAsObject, name:=sht.name
    
    n = ActiveChart.ChartGroups(1).SeriesCollection().Count
    
    If n Mod 2 = 1 Then
        MsgBox "�f�[�^���������ɂȂ��Ă��܂���D"
        ActiveChart.Parent.Delete
        Exit Sub
    End If
    
    ' ���ɑ���
    For i = 1 To n / 2
        ActiveChart.ChartGroups(1).SeriesCollection(1).PlotOrder = n
    Next

    ' ���Ԃ𔽓]������
'    For i = 1 To n - 1
'       ActiveChart.ChartGroups(1).SeriesCollection(n ).PlotOrder = i
'    Next

    If n = 2 Then
        ' �񐔂��Q�Ȃ�}��͕s�v
        ActiveChart.Legend.Delete
        With ActiveChart.ChartGroups(1).SeriesCollection(1).Interior
            .ColorIndex = 48
            .Pattern = xlSolid
        End With
    Else
       �_�O���t�F�ݒ�.applyBarFill �_�O���t�F�ݒ�.�����p�^�[��20()
    End If
    Dim items As Integer
    items = n / 2 * rng.Columns.Count
    Call A4�ŏc3�i�p�ɃT�C�Y�ύX
    Call �㔼�̌n����f�[�^�\���p�ɕύX
    Call AAA_enquete_graph(False, False)
    With ActiveChart.Axes(xlValue)
        .TickLabels.NumberFormatLocal = "0%"
        .HasTitle = True
        .AxisTitle.Characters.Text = "�񓚂̑I��"
        .AxisTitle.Orientation = xlVertical
    End With
    If n = 2 Then
        Call �f�[�^���x�������w�i�ύX
    End If
    Call �f�[�^���x�������T�C�Y�ύX(8#)
    ActiveChart.ChartArea.Select
End Sub



'�A���P�[�g�p�ɃO���t��ݒ肷��D

Sub AAA_enquete_graph(Optional drawSreiesLines As Boolean = True, Optional swapOrder As Boolean = True)
'
' Macro2 Macro
' �}�N���L�^�� : 2012/4/10  ���[�U�[�� : NSC999
'

'
    �O���t�̔w�i�Ȃ���
    Dim a As Axis
    If ActiveChart Is Nothing Then Exit Sub
    Set a = ActiveChart.Axes(xlCategory)
    With a
        .TickLabels.Orientation = xlHorizontal
        .TickLabelSpacing = 1
        .TickMarkSpacing = 1
        Select Case ActiveChart.SeriesCollection(1).ChartType
        Case xlBarClustered, xlBarStacked, xlBarStacked100, _
                xl3DBarClustered, xl3DBarStacked, xl3DBarStacked100
            If swapOrder Then
                .Crosses = xlMaximum
                .ReversePlotOrder = True
            End If
        End Select
    End With
    Dim cg As ChartGroup
    
    For Each cg In ActiveChart.ChartGroups
    With cg
        .GapWidth = 50
        .HasSeriesLines = drawSreiesLines
        If drawSreiesLines Then
        With .SeriesLines.Border
            .ColorIndex = 57
            .weight = xlThin
            .LineStyle = xlContinuous
        End With
        End If
    End With
    Next
End Sub




Sub �e�X�g�_�C�A���O�\��()
  Dim d As New �_�C�A���O�e�X�g
  d.Show vbModal
   
  Set d = Nothing

End Sub



Private Sub FFT_OCT_FORMAT_IT_THEN_PRINT_IT()
Attribute FFT_OCT_FORMAT_IT_THEN_PRINT_IT.VB_Description = "�}�N���L�^�� : 2011/2/4  ���[�U�[�� : �c��"
Attribute FFT_OCT_FORMAT_IT_THEN_PRINT_IT.VB_ProcData.VB_Invoke_Func = " \n14"
'
' FFT_OCT_FORMAT_IT_THEN_PRINT_IT Macro
' �}�N���L�^�� : 2011/2/4  ���[�U�[�� : �c��
'

'
    Sheets("�}").Select
    Range("A1").Select
    ActiveWindow.SelectedSheets.PrintOut Copies:=1
    Sheets("DATA").Select
    
    Range("H7:M23").Select
    Selection.Cut Destination:=Range("B25:G41")
    
    Range("N7:S23").Select
    Selection.Cut Destination:=Range("B43:G59")
    
    Range("T7:Y22").Select
    Selection.Cut Destination:=Range("B61:G76")
    
    Range("A8:A23").Select
    Selection.Copy Destination:=Range("A26:A41")
    
    Range("A26:A41").Select
    Selection.Copy Destination:=Range("A44:A59")
    
    Range("A44:A59").Select
    Selection.Copy Destination:=Range("A61:A76")
    
    Columns("B:G").Select
    Selection.ColumnWidth = 17
    
    Range("A1:H77").Select
    ActiveSheet.PageSetup.PrintArea = "$A$1:$G$76"
    With ActiveSheet.PageSetup
        .PrintTitleRows = ""
        .PrintTitleColumns = ""
    End With
    With ActiveSheet.PageSetup
        .LeftHeader = ""
        .CenterHeader = ""
        .RightHeader = ""
        .LeftFooter = ""
        .CenterFooter = ""
        .RightFooter = ""
        .LeftMargin = Application.InchesToPoints(0.787)
        .RightMargin = Application.InchesToPoints(0.787)
        .TopMargin = Application.InchesToPoints(0.984)
        .BottomMargin = Application.InchesToPoints(0.984)
        .HeaderMargin = Application.InchesToPoints(0.512)
        .FooterMargin = Application.InchesToPoints(0.512)
        .PrintHeadings = False
        .PrintGridlines = False
        .PrintComments = xlPrintNoComments
        .CenterHorizontally = False
        .CenterVertically = False
        .Orientation = xlPortrait
        .Draft = False
        .PaperSize = xlPaperA4
        .FirstPageNumber = xlAutomatic
        .Order = xlDownThenOver
        .BlackAndWhite = False
        .Zoom = False
        .FitToPagesWide = 1
        .FitToPagesTall = 1
        .PrintErrors = xlPrintErrorsDisplayed
    End With
    ActiveWindow.SelectedSheets.PrintOut Copies:=1
    ActiveWorkbook.Save
    ActiveWindow.Close
End Sub



Private Sub PrintThenCloseIt()
'
' PrintThenCloseIt Macro
' �}�N���L�^�� : 2011/2/4  ���[�U�[�� : �c��
'
'
' Keyboard Shortcut: Ctrl+j
'
    ActiveWindow.SelectedSheets.PrintOut Copies:=1
    ActiveWindow.Close SaveChanges:=False
End Sub


Private Sub Temp_ACCimport()
  
  Dim dir As String
  Dim ch As String
  Dim wb As Workbook
  Dim accbook As Workbook
  Dim setting As Worksheet
  Dim data As Worksheet
  Dim source As Worksheet
  Dim accfile As String
  Dim source_range As Range
  Dim data_range As Range
  Dim ���� As String
  Dim �ꏊ As String
  Dim fname As String
  Dim cwd As String
  
  
  If ActiveSheet.name <> "�ݒ�" Then
    MsgBox "[�ݒ�]�V�[�g���A�N�e�B�u�ɂ��Ă�����s���Ă��������D"
    Exit Sub
  End If
  
  Set wb = ActiveWorkbook
  Set setting = wb.Sheets("�ݒ�")
  Set data = wb.Sheets("data")
  
  ch = setting.Range("B2").Formula
  ���� = CStr(setting.Range("B3").Formula)
  �ꏊ = setting.Range("B1").Formula
  
  'cwd = "D:\Documents\03948\My Documents\Analysis\morido\dai2kurami\"
  If setting.Range("B4").Formula = "" Then
    cwd = Application.GetOpenFilename( _
      FileFilter:="CSV files,*.csv,AllFiles(*.*),*.*", _
      Title:="kudari-ch1.csv��I��")
    cwd = Left$(cwd, InStrRev(cwd, "\"))
    setting.Range("B4").Formula = cwd
  Else
    cwd = setting.Range("B4").Formula
  End If
    
  
  fname = cwd & �ꏊ & ���� & ch & "ch.xls"
  
  If ���� = "���" Then
    dir = "nobori"
  Else
    dir = "kudari"
  End If
  
  accfile = cwd & dir & "-ch" & ch & ".csv"
  
  Set accbook = Application.Workbooks.Open(accfile)
  Set source = accbook.Sheets(1)
  
  Set source_range = source.Range("A1")
  Set data_range = data.Range("A1")
  source_range.CurrentRegion.Copy
  data_range.PasteSpecial xlPasteValues
    
  Set source_range = source_range.End(xlToRight).End(xlToRight)
  Set data_range = data_range.End(xlToRight).End(xlToRight)
  source_range.CurrentRegion.Copy
  data_range.PasteSpecial xlPasteValues
  
  Set source_range = source_range.End(xlToRight).End(xlToRight)
  Set data_range = data_range.End(xlToRight).End(xlToRight)
  source_range.CurrentRegion.Copy
  data_range.PasteSpecial xlPasteValues
  
  setting.Range("B1").Copy
  
  accbook.Close msoFalse
  
  
  wb.SaveAs fname
  wb.Sheets(Array("�X�y�N�g���}", "�g�`�}")).PrintOut
  
  setting.Activate
  setting.Range("B2").Select

  
End Sub


Private Sub Temp_�O���t�c�����C�����Ĉ��������()

  Dim cwd As String
  cwd = Application.GetOpenFilename( _
    FileFilter:="CSV files,*.csv,AllFiles(*.*),*.*", _
    Title:="kudari-dis_0_5.csv��I��")
  cwd = Left$(cwd, InStrRev(cwd, "\"))
  Dim i  As Integer
  Dim freqs(4) As String
  
  For i = 1 To 4
    freqs(i) = "-dis-" & CStr(i * 3 - 1) & "_5-" & CStr(i * 3) & "_5.xls"
  Next i
  freqs(0) = "-dis-0_5.xls"

  Dim dirname
  Dim fname As String
  Dim dirs(1)
  
  dirs(1) = "nobori"
  dirs(0) = "kudari"
  
  Dim wb As Workbook
  Dim ch As Chart
  Dim ax As Axis
  
  For Each dirname In dirs
    For i = 0 To 4
      fname = cwd & dirname & freqs(i)
      Set wb = Application.Workbooks.Open(fname)
      wb.Sheets("Graph1").Activate
      Set ch = ActiveChart
      Set ax = ch.Axes(xlValue)
      ax.AxisTitle.Characters.Text = "�ψ�(m)"
      ch.PrintOut
      wb.Close msoTrue
    Next i
  Next dirname

End Sub


Private Sub Temp_�O���t������4����6�b���g�債�������쐬���Ĉ��������()

  Dim cwd As String
  cwd = Application.GetOpenFilename( _
    FileFilter:="CSV files,*.csv,AllFiles(*.*),*.*", _
    Title:="kudari-dis_0_5.csv��I��")
  cwd = Left$(cwd, InStrRev(cwd, "\"))
  Dim i  As Integer
  Dim freqs(4) As String
  
  For i = 1 To 4
    freqs(i) = "-dis-" & CStr(i * 3 - 1) & "_5-" & CStr(i * 3) & "_5.xls"
  Next i
  freqs(0) = "-dis-0_5.xls"

  Dim dirname
  Dim fname As String
  Dim dirs(1)
  
  dirs(1) = "nobori"
  dirs(0) = "kudari"
  
  Dim wb As Workbook
  Dim ch As Chart
  Dim ax As Axis
  
  Dim ymax(4)
  ymax(0) = 0.0001
  ymax(1) = 0.00005
  ymax(2) = 0.00002
  ymax(3) = 0.00006
  ymax(4) = 0.00001
  
  
  For Each dirname In dirs
    For i = 0 To 4
      fname = cwd & dirname & freqs(i)
      Set wb = Application.Workbooks.Open(fname)
      wb.Sheets("Graph1").Activate
      ActiveChart.Copy After:=ActiveChart
      Set ch = ActiveChart
      Set ax = ch.Axes(xlValue)
      ax.MaximumScale = ymax(i)
      ax.MinimumScale = -ymax(i)
      ch.PrintOut
      Set ax = ch.Axes(xlCategory)
      ax.MaximumScale = 6#
      ax.MinimumScale = 4
      ch.PrintOut
      wb.Close msoTrue
    Next i
  Next dirname

End Sub




Private Sub Macro1()
'
' Macro1 Macro
' �}�N���L�^�� : 2011/3/11  ���[�U�[�� : �c��
'

'
    With Selection
        .HorizontalAlignment = xlCenter
        .VerticalAlignment = xlCenter
        .WrapText = False
        .Orientation = 0
        .AddIndent = False
        .IndentLevel = 0
        .ShrinkToFit = False
        .ReadingOrder = xlContext
        .MergeCells = False
    End With
    Selection.Merge
    With Selection
        .HorizontalAlignment = xlLeft
        .VerticalAlignment = xlCenter
        .WrapText = False
        .Orientation = 0
        .AddIndent = False
        .IndentLevel = 0
        .ShrinkToFit = False
        .ReadingOrder = xlContext
        .MergeCells = True
    End With
    Selection.NumberFormatLocal = "0.000E+00"
End Sub


Private Sub ���l�f�[�^���C���|�[�g()
    Dim b As Workbook
    Dim ab As Workbook
    Dim ts As Worksheet
    
    Set ab = ActiveWorkbook
  
    Dim dlgOpen As FileDialog
    Set dlgOpen = Application.FileDialog(msoFileDialogOpen)
    
    dlgOpen.AllowMultiSelect = True
    dlgOpen.Filters.Add "Image File", "*.xlsx", 1
    If dlgOpen.Show = -1 Then
      Dim fname As Variant
      Dim n As Integer
      Dim e As Integer
      For Each fname In dlgOpen.SelectedItems
        Set b = Application.Workbooks.Open(fname)
        Set ts = b.Worksheets.item(1)
        n = InStr(fname, "�i") + 1
        e = InStr(fname, ")")
        ts.name = Mid$(fname, n, e - n)
        ts.Copy After:=ab.Worksheets(ab.Worksheets.Count)
        b.Close False
      Next fname
    End If
    
    Set dlgOpen = Nothing

End Sub



Private Sub ����͈͐ݒ�()

' ����͈͐ݒ� Macro
' �}�N���L�^�� : 2011/3/17  ���[�U�[�� : �c��
'
' Keyboard Shortcut: Ctrl+Shift+P
'
  On Error GoTo eee
  Application.ScreenUpdating = False
    
    ActiveSheet.PageSetup.PrintArea = "$B$81:$AK$151"
    With ActiveSheet.PageSetup
        .LeftHeader = ""
        .CenterHeader = ""
        .RightHeader = ""
        .LeftFooter = ""
        .CenterFooter = ""
        .RightFooter = ""
        .LeftMargin = Application.InchesToPoints(0.787)
        .RightMargin = Application.InchesToPoints(0.787)
        .TopMargin = Application.InchesToPoints(0.984)
        .BottomMargin = Application.InchesToPoints(0.984)
        .HeaderMargin = Application.InchesToPoints(0.512)
        .FooterMargin = Application.InchesToPoints(0.512)
        .PrintHeadings = False
        .PrintGridlines = False
        .PrintComments = xlPrintNoComments
        .PrintQuality = 600
        .CenterHorizontally = False
        .CenterVertically = False
        .Orientation = xlPortrait
        .Draft = False
        .PaperSize = xlPaperA4
        .FirstPageNumber = xlAutomatic
        .Order = xlDownThenOver
        .BlackAndWhite = False
        .Zoom = False
        .FitToPagesWide = 2
        .FitToPagesTall = 1
        .PrintErrors = xlPrintErrorsDisplayed
    End With

eee:
  Application.ScreenUpdating = True

End Sub

Sub �}�̃X�P�[��20�p�[�Z���g()
'
' Macro7 Macro
' �}�N���L�^�� : 2012/9/24  ���[�U�[�� : NSC999
'

'
Dim s As ShapeRange
    Set s = Selection.ShapeRange
    With s
        .LockAspectRatio = msoTrue
        .ScaleHeight 0.2, msoTrue
        .ScaleWidth 0.2, msoTrue
    End With
End Sub

