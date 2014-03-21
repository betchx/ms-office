Attribute VB_Name = "Temporary"
Option Explicit


Sub ☆ジャンプ()
  JumpList.Show vbModeless
End Sub

Sub アンケート横棒100グラフ作成()
'
' アンケート回答率横棒グラフ作成 Macro
' マクロ記録日 : 2012/4/20  ユーザー名 : NSC999
'
    Dim rng As Range, sht As Worksheet
    Set sht = ActiveSheet
    Set rng = ActiveWindow.Selection


    Charts.Add
    ActiveChart.ChartType = xlBarStacked100
    ActiveChart.SetSourceData source:=rng, PlotBy:=xlRows
    ActiveChart.Location Where:=xlLocationAsObject, name:=sht.name
    
    棒グラフ色設定.applyBarFill 棒グラフ色設定.白黒パターン5()
    凡例を上に
    Call A4で縦3段用にサイズ変更
    Call AAA_enquete_graph(True)
    With ActiveChart.Axes(xlValue)
        .TickLabels.NumberFormatLocal = "0%"
        .HasTitle = True
        .AxisTitle.Characters.Text = "回答の選択率"
    End With
    目盛ラベル間隔を1に
    データラベルを値に
    
    If rng.Columns.Count = 2 Then
      Dim ax As Axis
      Set ax = ActiveChart.Axes(xlCategory)
      ax.TickLabels.Delete
    Else
      ActiveChart.HasTitle = True
      ActiveChart.ChartTitle.Text = rng.Range("A1").Value
    End If
    
'    Call データラベル文字背景変更
    ActiveChart.ChartArea.Select


End Sub


Sub アンケート選択率グラフ作成()
'
' Macro3 Macro
' マクロ記録日 : 2012/4/17  ユーザー名 : NSC999
'
'
    Dim rng As Range, sht As Worksheet
    Set sht = ActiveSheet
    Set rng = ActiveWindow.Selection
    Dim n As Integer, i  As Integer
    n = rng.Columns.Count
    If n < 3 Then
        MsgBox "選択範囲の列数が不足しています．"
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
        If MsgBox("系列数が偶数になっていません．削除します．", vbOKCancel) = vbOK Then _
            ActiveChart.Parent.Delete
        Exit Sub
    End If
    
    
    ' 後ろに送る
    For i = 1 To n / 2
        ActiveChart.ChartGroups(1).SeriesCollection(1).PlotOrder = n
    Next
    ' 列数が２なら凡例は不要
    If n = 2 Then
        ActiveChart.Legend.Delete
        With ActiveChart.ChartGroups(1).SeriesCollection(1).Interior
            .ColorIndex = 48
            .Pattern = xlSolid
        End With
    Else
       棒グラフ色設定.applyBarFill 棒グラフ色設定.白黒パターン20()
    End If
    Call A4で縦3段用にサイズ変更
    Call 後半の系列をデータ表示用に変更
    Call AAA_enquete_graph(False)
    With ActiveChart.Axes(xlValue)
        .TickLabels.NumberFormatLocal = "0%"
        .HasTitle = True
        .AxisTitle.Characters.Text = "回答の選択率"
    End With
    Call データラベル文字背景変更
    ActiveChart.ChartArea.Select
End Sub

Sub アンケート行選択率グラフ作成()
'
' Macro3 Macro
' マクロ記録日 : 2012/4/17  ユーザー名 : NSC999
'
'
    Dim rng As Range, sht As Worksheet
    Set sht = ActiveSheet
    Set rng = ActiveWindow.Selection
    Dim n As Integer, i  As Integer
    n = rng.Rows.Count
    If n < 3 Then
        MsgBox "選択範囲の行数が不足しています．"
        Exit Sub
    End If
    
    Charts.Add
    ActiveChart.ChartType = xlBarClustered
    ActiveChart.SetSourceData source:=rng, PlotBy:=xlRows
    ActiveChart.Location Where:=xlLocationAsObject, name:=sht.name
    
    n = ActiveChart.ChartGroups(1).SeriesCollection().Count
    If n Mod 2 = 1 Then
        If MsgBox("系列数が偶数になっていません．削除します．", vbOKCancel) = vbOK Then _
            ActiveChart.Parent.Delete
        Exit Sub
    End If
    
'    ' 後ろに送る
'    For i = 1 To n / 2
'        ActiveChart.ChartGroups(1).SeriesCollection(1).PlotOrder = n
'    Next

    ' 順番を反転させる
    For i = 1 To (n - 1)
       ActiveChart.ChartGroups(1).SeriesCollection(n).PlotOrder = i
    Next

    If n = 2 Then
        ' 列数が２なら凡例は不要
        ActiveChart.Legend.Delete
        With ActiveChart.ChartGroups(1).SeriesCollection(1).Interior
            .ColorIndex = 48
            .Pattern = xlSolid
        End With
    Else
       棒グラフ色設定.applyBarFill 棒グラフ色設定.白黒パターン20()
    End If
    Dim items As Integer
    items = n / 2 * rng.Columns.Count
    If items > 20 Then
        Call A4用にサイズ変更
    Else
        Call A4で縦3段用にサイズ変更
    End If
    Call 後半の系列をデータ表示用に変更
    Call AAA_enquete_graph(False, False)
    With ActiveChart.Axes(xlValue)
        .TickLabels.NumberFormatLocal = "0%"
        .HasTitle = True
        .AxisTitle.Characters.Text = "回答の選択率"
    End With
    If n = 2 Then
        Call データラベル文字背景変更
    End If
    Call データラベル文字サイズ変更(8#)
    ActiveChart.ChartArea.Select
End Sub

Sub アンケート選択率縦棒グラフ作成()
'
' Macro3 Macro
' マクロ記録日 : 2012/4/17  ユーザー名 : NSC999
'
'
    Dim rng As Range, sht As Worksheet
    Set sht = ActiveSheet
    Set rng = ActiveWindow.Selection
    Dim n As Integer, i  As Integer
    n = rng.Columns.Count
    If n < 3 Then
        MsgBox "選択範囲の列数が不足しています．"
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
        If MsgBox("系列数が偶数になっていません．削除します．", vbOKCancel) = vbOK Then _
            ActiveChart.Parent.Delete
        Exit Sub
    End If
    
    
    ' 後ろに送る
    For i = 1 To n / 2
        ActiveChart.ChartGroups(1).SeriesCollection(1).PlotOrder = n
    Next
    ' 列数が２なら凡例は不要
    If n = 2 Then
        ActiveChart.Legend.Delete
        With ActiveChart.ChartGroups(1).SeriesCollection(1).Interior
            .ColorIndex = 48
            .Pattern = xlSolid
        End With
    Else
       棒グラフ色設定.applyBarFill 棒グラフ色設定.白黒パターン20()
    End If
    Call A4で縦3段用にサイズ変更
    Call 後半の系列をデータ表示用に変更
    Call AAA_enquete_graph(False)
    With ActiveChart.Axes(xlValue)
        .TickLabels.NumberFormatLocal = "0%"
        .HasTitle = True
        .AxisTitle.Characters.Text = "回答の選択率"
        .AxisTitle.Orientation = xlVertical
    End With
    If n = 2 Then Call データラベル文字背景変更
    ActiveChart.ChartArea.Select
End Sub

Sub アンケート行選択率縦棒グラフ作成()
'
' Macro3 Macro
' マクロ記録日 : 2012/4/17  ユーザー名 : NSC999
'
'
    Dim rng As Range, sht As Worksheet
    Set sht = ActiveSheet
    Set rng = ActiveWindow.Selection
    Dim n As Integer, i  As Integer
    n = rng.Rows.Count
    If n < 3 Then
        MsgBox "選択範囲の行数が不足しています．"
        Exit Sub
    End If

    
    Charts.Add
    ActiveChart.ChartType = xlColumnClustered
    ActiveChart.SetSourceData source:=rng, PlotBy:=xlRows
    ActiveChart.Location Where:=xlLocationAsObject, name:=sht.name
    
    n = ActiveChart.ChartGroups(1).SeriesCollection().Count
    
    If n Mod 2 = 1 Then
        MsgBox "データ数が偶数になっていません．"
        ActiveChart.Parent.Delete
        Exit Sub
    End If
    
    ' 後ろに送る
    For i = 1 To n / 2
        ActiveChart.ChartGroups(1).SeriesCollection(1).PlotOrder = n
    Next

    ' 順番を反転させる
'    For i = 1 To n - 1
'       ActiveChart.ChartGroups(1).SeriesCollection(n ).PlotOrder = i
'    Next

    If n = 2 Then
        ' 列数が２なら凡例は不要
        ActiveChart.Legend.Delete
        With ActiveChart.ChartGroups(1).SeriesCollection(1).Interior
            .ColorIndex = 48
            .Pattern = xlSolid
        End With
    Else
       棒グラフ色設定.applyBarFill 棒グラフ色設定.白黒パターン20()
    End If
    Dim items As Integer
    items = n / 2 * rng.Columns.Count
    Call A4で縦3段用にサイズ変更
    Call 後半の系列をデータ表示用に変更
    Call AAA_enquete_graph(False, False)
    With ActiveChart.Axes(xlValue)
        .TickLabels.NumberFormatLocal = "0%"
        .HasTitle = True
        .AxisTitle.Characters.Text = "回答の選択率"
        .AxisTitle.Orientation = xlVertical
    End With
    If n = 2 Then
        Call データラベル文字背景変更
    End If
    Call データラベル文字サイズ変更(8#)
    ActiveChart.ChartArea.Select
End Sub



'アンケート用にグラフを設定する．

Sub AAA_enquete_graph(Optional drawSreiesLines As Boolean = True, Optional swapOrder As Boolean = True)
'
' Macro2 Macro
' マクロ記録日 : 2012/4/10  ユーザー名 : NSC999
'

'
    グラフの背景なしに
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




Sub テストダイアログ表示()
  Dim d As New ダイアログテスト
  d.Show vbModal
   
  Set d = Nothing

End Sub



Private Sub FFT_OCT_FORMAT_IT_THEN_PRINT_IT()
Attribute FFT_OCT_FORMAT_IT_THEN_PRINT_IT.VB_Description = "マクロ記録日 : 2011/2/4  ユーザー名 : 田辺"
Attribute FFT_OCT_FORMAT_IT_THEN_PRINT_IT.VB_ProcData.VB_Invoke_Func = " \n14"
'
' FFT_OCT_FORMAT_IT_THEN_PRINT_IT Macro
' マクロ記録日 : 2011/2/4  ユーザー名 : 田辺
'

'
    Sheets("図").Select
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
' マクロ記録日 : 2011/2/4  ユーザー名 : 田辺
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
  Dim 方向 As String
  Dim 場所 As String
  Dim fname As String
  Dim cwd As String
  
  
  If ActiveSheet.name <> "設定" Then
    MsgBox "[設定]シートをアクティブにしてから実行してください．"
    Exit Sub
  End If
  
  Set wb = ActiveWorkbook
  Set setting = wb.Sheets("設定")
  Set data = wb.Sheets("data")
  
  ch = setting.Range("B2").Formula
  方向 = CStr(setting.Range("B3").Formula)
  場所 = setting.Range("B1").Formula
  
  'cwd = "D:\Documents\03948\My Documents\Analysis\morido\dai2kurami\"
  If setting.Range("B4").Formula = "" Then
    cwd = Application.GetOpenFilename( _
      FileFilter:="CSV files,*.csv,AllFiles(*.*),*.*", _
      Title:="kudari-ch1.csvを選択")
    cwd = Left$(cwd, InStrRev(cwd, "\"))
    setting.Range("B4").Formula = cwd
  Else
    cwd = setting.Range("B4").Formula
  End If
    
  
  fname = cwd & 場所 & 方向 & ch & "ch.xls"
  
  If 方向 = "上り" Then
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
  wb.Sheets(Array("スペクトル図", "波形図")).PrintOut
  
  setting.Activate
  setting.Range("B2").Select

  
End Sub


Private Sub Temp_グラフ縦軸を修正して印刷し閉じる()

  Dim cwd As String
  cwd = Application.GetOpenFilename( _
    FileFilter:="CSV files,*.csv,AllFiles(*.*),*.*", _
    Title:="kudari-dis_0_5.csvを選択")
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
      ax.AxisTitle.Characters.Text = "変位(m)"
      ch.PrintOut
      wb.Close msoTrue
    Next i
  Next dirname

End Sub


Private Sub Temp_グラフ横軸が4から6秒を拡大した物を作成して印刷し閉じる()

  Dim cwd As String
  cwd = Application.GetOpenFilename( _
    FileFilter:="CSV files,*.csv,AllFiles(*.*),*.*", _
    Title:="kudari-dis_0_5.csvを選択")
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
' マクロ記録日 : 2011/3/11  ユーザー名 : 田辺
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


Private Sub 数値データをインポート()
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
        n = InStr(fname, "（") + 1
        e = InStr(fname, ")")
        ts.name = Mid$(fname, n, e - n)
        ts.Copy After:=ab.Worksheets(ab.Worksheets.Count)
        b.Close False
      Next fname
    End If
    
    Set dlgOpen = Nothing

End Sub



Private Sub 印刷範囲設定()

' 印刷範囲設定 Macro
' マクロ記録日 : 2011/3/17  ユーザー名 : 田辺
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

Sub 図のスケール20パーセント()
'
' Macro7 Macro
' マクロ記録日 : 2012/9/24  ユーザー名 : NSC999
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

