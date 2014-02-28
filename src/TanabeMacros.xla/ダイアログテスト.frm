VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} ダイアログテスト 
   Caption         =   "ダイアログ表示テスト"
   ClientHeight    =   5835
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   10185
   OleObjectBlob   =   "ダイアログテスト.frx":0000
   StartUpPosition =   1  'オーナー フォームの中央
End
Attribute VB_Name = "ダイアログテスト"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private values(243, 2)



Private Sub CommandButton1_Click()
Dim code
Dim rtn As Boolean
Dim d As Dialog

On Error GoTo eee

code = values(Me.ListBox1.ListIndex, 2)

Set d = Application.Dialogs(code)
If IsNull(d) Or IsEmpty(d) Then
    MsgBox "ダイアログを取得できませんでした"
Else
    rtn = d.Show()
End If
Exit Sub

eee:
 MsgBox "エラーが発生し表示できませんでした．"
 
End Sub

Private Sub UserForm_Initialize()

values(0, 0) = "   xlDialogActivate    ":   values(0, 1) = "   [ウィンドウの選択] ダイアログ ボックス  ":   values(0, 2) = xlDialogActivate
values(1, 0) = "   xlDialogActiveCellFont  ":   values(1, 1) = "   [セルの書式設定 (フォント)] ダイアログ ボックス ":   values(1, 2) = xlDialogActiveCellFont
values(2, 0) = "   xlDialogAddChartAutoformat  ":   values(2, 1) = "   [ユーザー設定のグラフ種類の追加] ダイアログ ボックス    ":   values(2, 2) = xlDialogAddChartAutoformat
values(3, 0) = "   xlDialogAddinManager    ":   values(3, 1) = "   [アドイン] ダイアログ ボックス  ":   values(3, 2) = xlDialogAddinManager
values(4, 0) = "   xlDialogAlignment   ":   values(4, 1) = "   [セルの書式設定 (配置)] ダイアログ ボックス ":   values(4, 2) = xlDialogAlignment
values(5, 0) = "   xlDialogApplyNames  ":   values(5, 1) = "   [名前の引用] ダイアログ ボックス    ":   values(5, 2) = xlDialogApplyNames
values(6, 0) = "   xlDialogApplyStyle  ":   values(6, 1) = "   [スタイル] ダイアログ ボックス  ":   values(6, 2) = xlDialogApplyStyle
values(7, 0) = "   xlDialogAppMove ":   values(7, 1) = "   [移動(アプリケーション)] ダイアログ ボックス    ":   values(7, 2) = xlDialogAppMove
values(8, 0) = "   xlDialogAppSize ":   values(8, 1) = "   [送信] ダイアログ ボックス  ":   values(8, 2) = xlDialogAppSize
values(9, 0) = "   xlDialogArrangeAll  ":   values(9, 1) = "   [ウィンドウの整列] ダイアログ ボックス  ":   values(9, 2) = xlDialogArrangeAll
values(10, 0) = "   xlDialogAssignToObject  ":  values(10, 1) = "   [オブジェクトへの登録] ダイアログ ボックス  ":  values(10, 2) = xlDialogAssignToObject
values(11, 0) = "   xlDialogAssignToTool    ":  values(11, 1) = "   [ツールに割り当て] ダイアログ ボックス  ":  values(11, 2) = xlDialogAssignToTool
values(12, 0) = "   xlDialogAttachText  ":  values(12, 1) = "   [文字の追加] ダイアログ ボックス    ":  values(12, 2) = xlDialogAttachText
values(13, 0) = "   xlDialogAttachToolbars  ":  values(13, 1) = "   [ブックへのツールバーの登録] ダイアログ ボックス    ":  values(13, 2) = xlDialogAttachToolbars
values(14, 0) = "   xlDialogAutoCorrect ":  values(14, 1) = "   [オートコレクト (オートコレクト)] ダイアログ ボックス   ":  values(14, 2) = xlDialogAutoCorrect
values(15, 0) = "   xlDialogAxes    ":  values(15, 1) = "   [軸] ダイアログ ボックス    ":  values(15, 2) = xlDialogAxes
values(16, 0) = "   xlDialogBorder  ":  values(16, 1) = "   [セルの書式設定 (罫線)] ダイアログ ボックス ":  values(16, 2) = xlDialogBorder
values(17, 0) = "   xlDialogCalculation ":  values(17, 1) = "   [計算方法の設定] ダイアログ ボックス    ":  values(17, 2) = xlDialogCalculation
values(18, 0) = "   xlDialogCellProtection  ":  values(18, 1) = "   [セルの書式設定 (保護)] ダイアログ ボックス ":  values(18, 2) = xlDialogCellProtection
values(19, 0) = "   xlDialogChangeLink  ":  values(19, 1) = "   [リンクの変更] ダイアログ ボックス  ":  values(19, 2) = xlDialogChangeLink
values(20, 0) = "   xlDialogChartAddData    ":  values(20, 1) = "   [グラフ追加データ] ダイアログ ボックス  ":  values(20, 2) = xlDialogChartAddData
values(21, 0) = "   xlDialogChartLocation   ":  values(21, 1) = "   [グラフの場所] ダイアログ ボックス  ":  values(21, 2) = xlDialogChartLocation
values(22, 0) = "   xlDialogChartOptionsDataLabelMultiple   ":  values(22, 1) = "   [グラフ オプション データ ラベル複数] ダイアログ ボックス   ":  values(22, 2) = xlDialogChartOptionsDataLabelMultiple
values(23, 0) = "   xlDialogChartOptionsDataLabels  ":  values(23, 1) = "   [グラフ オプション データ ラベル] ダイアログ ボックス   ":  values(23, 2) = xlDialogChartOptionsDataLabels
values(24, 0) = "   xlDialogChartOptionsDataTable   ":  values(24, 1) = "   [グラフ オプション データ テーブル] ダイアログ ボックス ":  values(24, 2) = xlDialogChartOptionsDataTable
values(25, 0) = "   xlDialogChartSourceData ":  values(25, 1) = "   [グラフの元データ] ダイアログ ボックス  ":  values(25, 2) = xlDialogChartSourceData
values(26, 0) = "   xlDialogChartTrend  ":  values(26, 1) = "   [グラフ トレンド] ダイアログ ボックス   ":  values(26, 2) = xlDialogChartTrend
values(27, 0) = "   xlDialogChartType   ":  values(27, 1) = "   [グラフの種類] ダイアログ ボックス  ":  values(27, 2) = xlDialogChartType
values(28, 0) = "   xlDialogChartWizard ":  values(28, 1) = "   [グラフ ウィザード] ダイアログ ボックス ":  values(28, 2) = xlDialogChartWizard
values(29, 0) = "   xlDialogCheckboxProperties  ":  values(29, 1) = "   [チェック ボックスのプロパティ] ダイアログ ボックス ":  values(29, 2) = xlDialogCheckboxProperties
values(30, 0) = "   xlDialogClear   ":  values(30, 1) = "   [消去] ダイアログ ボックス  ":  values(30, 2) = xlDialogClear
values(31, 0) = "   xlDialogColorPalette    ":  values(31, 1) = "   [オプション (色)] ダイアログ ボックス   ":  values(31, 2) = xlDialogColorPalette
values(32, 0) = "   xlDialogColumnWidth ":  values(32, 1) = "   [列幅] ダイアログ ボックス  ":  values(32, 2) = xlDialogColumnWidth
values(33, 0) = "   xlDialogCombination ":  values(33, 1) = "   [複合] ダイアログ ボックス  ":  values(33, 2) = xlDialogCombination
values(34, 0) = "   xlDialogConditionalFormatting   ":  values(34, 1) = "   [条件付き書式の設定] ダイアログ ボックス    ":  values(34, 2) = xlDialogConditionalFormatting
values(35, 0) = "   xlDialogConsolidate ":  values(35, 1) = "   [統合の設定] ダイアログ ボックス    ":  values(35, 2) = xlDialogConsolidate
values(36, 0) = "   xlDialogCopyChart   ":  values(36, 1) = "   [グラフのコピー] ダイアログ ボックス    ":  values(36, 2) = xlDialogCopyChart
values(37, 0) = "   xlDialogCopyPicture ":  values(37, 1) = "   [図のコピー] ダイアログ ボックス    ":  values(37, 2) = xlDialogCopyPicture
values(38, 0) = "   xlDialogCreateList  ":  values(38, 1) = "   [リストの作成 ] ダイアログ ボックス ":  values(38, 2) = xlDialogCreateList
values(39, 0) = "   xlDialogCreateNames ":  values(39, 1) = "   [名前の作成] ダイアログ ボックス    ":  values(39, 2) = xlDialogCreateNames
values(40, 0) = "   xlDialogCreatePublisher ":  values(40, 1) = "   [発行側の作成] ダイアログ ボックス  ":  values(40, 2) = xlDialogCreatePublisher
values(41, 0) = "   xlDialogCustomizeToolbar    ":  values(41, 1) = "   [ユーザー設定 (オプション)] ダイアログ ボックス ":  values(41, 2) = xlDialogCustomizeToolbar
values(42, 0) = "   xlDialogCustomViews ":  values(42, 1) = "   [ユーザー設定のビュー] ダイアログ ボックス  ":  values(42, 2) = xlDialogCustomViews
values(43, 0) = "   xlDialogDataDelete  ":  values(43, 1) = "   [データの削除] ダイアログ ボックス  ":  values(43, 2) = xlDialogDataDelete
values(44, 0) = "   xlDialogDataLabel   ":  values(44, 1) = "   [データ ラベル] ダイアログ ボックス ":  values(44, 2) = xlDialogDataLabel
values(45, 0) = "   xlDialogDataLabelMultiple   ":  values(45, 1) = "   [データ ラベル複数] ダイアログ ボックス ":  values(45, 2) = xlDialogDataLabelMultiple
values(46, 0) = "   xlDialogDataSeries  ":  values(46, 1) = "   [連続データ] ダイアログ ボックス    ":  values(46, 2) = xlDialogDataSeries
values(47, 0) = "   xlDialogDataValidation  ":  values(47, 1) = "   [データの入力規則 (設定)] ダイアログ ボックス   ":  values(47, 2) = xlDialogDataValidation
values(48, 0) = "   xlDialogDefineName  ":  values(48, 1) = "   [名前の定義] ダイアログ ボックス    ":  values(48, 2) = xlDialogDefineName
values(49, 0) = "   xlDialogDefineStyle ":  values(49, 1) = "   [スタイル] ダイアログ ボックス  ":  values(49, 2) = xlDialogDefineStyle
values(50, 0) = "   xlDialogDeleteFormat    ":  values(50, 1) = "   [セルの書式設定 (表示形式)] ダイアログ ボックス ":  values(50, 2) = xlDialogDeleteFormat
values(51, 0) = "   xlDialogDeleteName  ":  values(51, 1) = "   [名前の定義] ダイアログ ボックス    ":  values(51, 2) = xlDialogDeleteName
values(52, 0) = "   xlDialogDemote  ":  values(52, 1) = "   [グループ化] ダイアログ ボックス    ":  values(52, 2) = xlDialogDemote
values(53, 0) = "   xlDialogDisplay ":  values(53, 1) = "   [画面設定] ダイアログ ボックス  ":  values(53, 2) = xlDialogDisplay
values(54, 0) = "   xlDialogDocumentInspector   ":  values(54, 1) = "(x)[ドキュメント検査] ダイアログ ボックス  ":  values(54, 2) = 0 'xlDialogDocumentInspector
values(55, 0) = "   xlDialogEditboxProperties   ":  values(55, 1) = "   [編集ボックスのプロパティ] ダイアログ ボックス  ":  values(55, 2) = xlDialogEditboxProperties
values(56, 0) = "   xlDialogEditColor   ":  values(56, 1) = "   [色の編集] ダイアログ ボックス  ":  values(56, 2) = xlDialogEditColor
values(57, 0) = "   xlDialogEditDelete  ":  values(57, 1) = "   [削除] ダイアログ ボックス  ":  values(57, 2) = xlDialogEditDelete
values(58, 0) = "   xlDialogEditionOptions  ":  values(58, 1) = "   [エディション オプション] ダイアログ ボックス   ":  values(58, 2) = xlDialogEditionOptions
values(59, 0) = "   xlDialogEditSeries  ":  values(59, 1) = "   [系列の編集] ダイアログ ボックス    ":  values(59, 2) = xlDialogEditSeries
values(60, 0) = "   xlDialogErrorbarX   ":  values(60, 1) = "   [Errorbar X] ダイアログ ボックス    ":  values(60, 2) = xlDialogErrorbarX
values(61, 0) = "   xlDialogErrorbarY   ":  values(61, 1) = "   [Errorbar Y] ダイアログ ボックス    ":  values(61, 2) = xlDialogErrorbarY
values(62, 0) = "   xlDialogErrorChecking   ":  values(62, 1) = "   [エラー チェック] ダイアログ ボックス   ":  values(62, 2) = xlDialogErrorChecking
values(63, 0) = "   xlDialogEvaluateFormula ":  values(63, 1) = "   [数式の検証] ダイアログ ボックス    ":  values(63, 2) = xlDialogEvaluateFormula
values(64, 0) = "   xlDialogExternalDataProperties  ":  values(64, 1) = "   [外部データのプロパティ] ダイアログ ボックス    ":  values(64, 2) = xlDialogExternalDataProperties
values(65, 0) = "   xlDialogExtract ":  values(65, 1) = "   [抽出] ダイアログ ボックス  ":  values(65, 2) = xlDialogExtract
values(66, 0) = "   xlDialogFileDelete  ":  values(66, 1) = "   [ファイルの削除] ダイアログ ボックス    ":  values(66, 2) = xlDialogFileDelete
values(67, 0) = "   xlDialogFileSharing ":  values(67, 1) = "   [ブックの共有 (編集)] ダイアログ ボックス   ":  values(67, 2) = xlDialogFileSharing
values(68, 0) = "   xlDialogFillGroup   ":  values(68, 1) = "   [グループの入力] ダイアログ ボックス    ":  values(68, 2) = xlDialogFillGroup
values(69, 0) = "   xlDialogFillWorkgroup   ":  values(69, 1) = "   [ワークグループの入力] ダイアログ ボックス  ":  values(69, 2) = xlDialogFillWorkgroup
values(70, 0) = "   xlDialogFilter  ":  values(70, 1) = "   [オートフィルター] ダイアログ ボックス  ":  values(70, 2) = xlDialogFilter
values(71, 0) = "   xlDialogFilterAdvanced  ":  values(71, 1) = "   [フィルター オプションの設定] ダイアログ ボックス   ":  values(71, 2) = xlDialogFilterAdvanced
values(72, 0) = "   xlDialogFindFile    ":  values(72, 1) = "   [ファイルを開く] ダイアログ ボックス    ":  values(72, 2) = xlDialogFindFile
values(73, 0) = "   xlDialogFont    ":  values(73, 1) = "   [フォントの設定] ダイアログ ボックス    ":  values(73, 2) = xlDialogFont
values(74, 0) = "   xlDialogFontProperties  ":  values(74, 1) = "   [セルの書式設定 (フォント)] ダイアログ ボックス ":  values(74, 2) = xlDialogFontProperties
values(75, 0) = "   xlDialogFormatAuto  ":  values(75, 1) = "   [オートフォーマット] ダイアログ ボックス    ":  values(75, 2) = xlDialogFormatAuto
values(76, 0) = "   xlDialogFormatChart ":  values(76, 1) = "   [グラフの書式設定] ダイアログ ボックス  ":  values(76, 2) = xlDialogFormatChart
values(77, 0) = "   xlDialogFormatCharttype ":  values(77, 1) = "   [グラフの種類] ダイアログ ボックス  ":  values(77, 2) = xlDialogFormatCharttype
values(78, 0) = "   xlDialogFormatFont  ":  values(78, 1) = "   [フォントの設定] ダイアログ ボックス    ":  values(78, 2) = xlDialogFormatFont
values(79, 0) = "   xlDialogFormatLegend    ":  values(79, 1) = "   [凡例の書式設定] ダイアログ ボックス    ":  values(79, 2) = xlDialogFormatLegend
values(80, 0) = "   xlDialogFormatMain  ":  values(80, 1) = "   [メイングラフ/重ね合わせグラフ] ダイアログ ボックス ":  values(80, 2) = xlDialogFormatMain
values(81, 0) = "   xlDialogFormatMove  ":  values(81, 1) = "   [移動の書式設定] ダイアログ ボックス    ":  values(81, 2) = xlDialogFormatMove
values(82, 0) = "   xlDialogFormatNumber    ":  values(82, 1) = "   [セルの書式設定 (表示形式)] ダイアログ ボックス ":  values(82, 2) = xlDialogFormatNumber
values(83, 0) = "   xlDialogFormatOverlay   ":  values(83, 1) = "   [重ね合わせグラフの設定] ダイアログ ボックス    ":  values(83, 2) = xlDialogFormatOverlay
values(84, 0) = "   xlDialogFormatSize  ":  values(84, 1) = "   [サイズの書式設定] ダイアログ ボックス  ":  values(84, 2) = xlDialogFormatSize
values(85, 0) = "   xlDialogFormatText  ":  values(85, 1) = "   [文字書式] ダイアログ ボックス  ":  values(85, 2) = xlDialogFormatText
values(86, 0) = "   xlDialogFormulaFind ":  values(86, 1) = "   [検索] ダイアログ ボックス  ":  values(86, 2) = xlDialogFormulaFind
values(87, 0) = "   xlDialogFormulaGoto ":  values(87, 1) = "   [ジャンプ] ダイアログ ボックス  ":  values(87, 2) = xlDialogFormulaGoto
values(88, 0) = "   xlDialogFormulaReplace  ":  values(88, 1) = "   [置換] ダイアログ ボックス  ":  values(88, 2) = xlDialogFormulaReplace
values(89, 0) = "   xlDialogFunctionWizard  ":  values(89, 1) = "   [関数の挿入] ダイアログ ボックス    ":  values(89, 2) = xlDialogFunctionWizard
values(90, 0) = "   xlDialogGallery3dArea   ":  values(90, 1) = "   [オートフォーマット (3-D 面)] ダイアログ ボックス   ":  values(90, 2) = xlDialogGallery3dArea
values(91, 0) = "   xlDialogGallery3dBar    ":  values(91, 1) = "   [オートフォーマット (横棒)] ダイアログ ボックス ":  values(91, 2) = xlDialogGallery3dBar
values(92, 0) = "   xlDialogGallery3dColumn ":  values(92, 1) = "   [オートフォーマット (3-D 縦棒)] ダイアログ ボックス ":  values(92, 2) = xlDialogGallery3dColumn
values(93, 0) = "   xlDialogGallery3dLine   ":  values(93, 1) = "   [オートフォーマット (3-D 折れ線)] ダイアログ ボックス   ":  values(93, 2) = xlDialogGallery3dLine
values(94, 0) = "   xlDialogGallery3dPie    ":  values(94, 1) = "   [オートフォーマット (3-D 円)] ダイアログ ボックス   ":  values(94, 2) = xlDialogGallery3dPie
values(95, 0) = "   xlDialogGallery3dSurface    ":  values(95, 1) = "   [オートフォーマット (等高線)] ダイアログ ボックス   ":  values(95, 2) = xlDialogGallery3dSurface
values(96, 0) = "   xlDialogGalleryArea ":  values(96, 1) = "   [オートフォーマット (面)] ダイアログ ボックス   ":  values(96, 2) = xlDialogGalleryArea
values(97, 0) = "   xlDialogGalleryBar  ":  values(97, 1) = "   [オートフォーマット (横棒)] ダイアログ ボックス ":  values(97, 2) = xlDialogGalleryBar
values(98, 0) = "   xlDialogGalleryColumn   ":  values(98, 1) = "   [オートフォーマット (縦棒)] ダイアログ ボックス ":  values(98, 2) = xlDialogGalleryColumn
values(99, 0) = "   xlDialogGalleryCustom   ":  values(99, 1) = "   [オートフォーマット (フォーマットの種類)] ダイアログ ボックス   ":  values(99, 2) = xlDialogGalleryCustom
values(100, 0) = "   xlDialogGalleryDoughnut ": values(100, 1) = "   [オートフォーマット (ドーナッツ)] ダイアログ ボックス   ": values(100, 2) = xlDialogGalleryDoughnut
values(101, 0) = "   xlDialogGalleryLine ": values(101, 1) = "   [オートフォーマット (折れ線)] ダイアログ ボックス   ": values(101, 2) = xlDialogGalleryLine
values(102, 0) = "   xlDialogGalleryPie  ": values(102, 1) = "   [オートフォーマット (円)] ダイアログ ボックス   ": values(102, 2) = xlDialogGalleryPie
values(103, 0) = "   xlDialogGalleryRadar    ": values(103, 1) = "   [オートフォーマット (レーダー)] ダイアログ ボックス ": values(103, 2) = xlDialogGalleryRadar
values(104, 0) = "   xlDialogGalleryScatter  ": values(104, 1) = "   [オートフォーマット (散布図)] ダイアログ ボックス   ": values(104, 2) = xlDialogGalleryScatter
values(105, 0) = "   xlDialogGoalSeek    ": values(105, 1) = "   [ゴール シーク] ダイアログ ボックス ": values(105, 2) = xlDialogGoalSeek
values(106, 0) = "   xlDialogGridlines   ": values(106, 1) = "   [グラフ オプション (目盛線)] ダイアログ ボックス    ": values(106, 2) = xlDialogGridlines
values(107, 0) = "   xlDialogImportTextFile  ": values(107, 1) = "   [テキスト ファイルのインポート] ダイアログ ボックス ": values(107, 2) = xlDialogImportTextFile
values(108, 0) = "   xlDialogInsert  ": values(108, 1) = "   [セルの挿入] ダイアログ ボックス    ": values(108, 2) = xlDialogInsert
values(109, 0) = "   xlDialogInsertHyperlink ": values(109, 1) = "   [ハイパーリンクの挿入] ダイアログ ボックス  ": values(109, 2) = xlDialogInsertHyperlink
values(110, 0) = "   xlDialogInsertObject    ": values(110, 1) = "   [オブジェクトの挿入 (新規作成)] ダイアログ ボックス ": values(110, 2) = xlDialogInsertObject
values(111, 0) = "   xlDialogInsertPicture   ": values(111, 1) = "   [図の挿入] ダイアログ ボックス  ": values(111, 2) = xlDialogInsertPicture
values(112, 0) = "   xlDialogInsertTitle ": values(112, 1) = "   [タイトル/軸ラベルの挿入] ダイアログ ボックス   ": values(112, 2) = xlDialogInsertTitle
values(113, 0) = "   xlDialogLabelProperties ": values(113, 1) = "   [ラベルのプロパティ] ダイアログ ボックス    ": values(113, 2) = xlDialogLabelProperties
values(114, 0) = "   xlDialogListboxProperties   ": values(114, 1) = "   [リスト ボックスのプロパティ] ダイアログ ボックス   ": values(114, 2) = xlDialogListboxProperties
values(115, 0) = "   xlDialogMacroOptions    ": values(115, 1) = "   [マクロ オプション] ダイアログ ボックス ": values(115, 2) = xlDialogMacroOptions
values(116, 0) = "   xlDialogMailEditMailer  ": values(116, 1) = "   [メール編集メーラー] ダイアログ ボックス    ": values(116, 2) = xlDialogMailEditMailer
values(117, 0) = "   xlDialogMailLogon   ": values(117, 1) = "   [回覧先] ダイアログ ボックス    ": values(117, 2) = xlDialogMailLogon
values(118, 0) = "   xlDialogMailNextLetter  ": values(118, 1) = "   [次の手紙の送信] ダイアログ ボックス    ": values(118, 2) = xlDialogMailNextLetter
values(119, 0) = "   xlDialogMainChart   ": values(119, 1) = "   [メイン グラフ] ダイアログ ボックス ": values(119, 2) = xlDialogMainChart
values(120, 0) = "   xlDialogMainChartType   ": values(120, 1) = "   [メイン グラフの種類] ダイアログ ボックス   ": values(120, 2) = xlDialogMainChartType
values(121, 0) = "   xlDialogMenuEditor  ": values(121, 1) = "   [メニュー エディター] ダイアログ ボックス   ": values(121, 2) = xlDialogMenuEditor
values(122, 0) = "   xlDialogMove    ": values(122, 1) = "   [移動] ダイアログ ボックス  ": values(122, 2) = xlDialogMove
values(123, 0) = "   xlDialogMyPermission    ": values(123, 1) = "   [アクセス許可] ダイアログ ボックス  ": values(123, 2) = xlDialogMyPermission
values(124, 0) = "   xlDialogNameManager ": values(124, 1) = "(x)[名前の管理] ダイアログ ボックス    ": values(124, 2) = 0 'xlDialogNameManager
values(125, 0) = "   xlDialogNew ": values(125, 1) = "   [新規作成 (標準)] ダイアログ ボックス   ": values(125, 2) = xlDialogNew
values(126, 0) = "   xlDialogNewName ": values(126, 1) = "(x)[新しい名前] ダイアログ ボックス    ": values(126, 2) = 0 'xlDialogNewName
values(127, 0) = "   xlDialogNewWebQuery ": values(127, 1) = "   [新しい Web クエリ] ダイアログ ボックス ": values(127, 2) = xlDialogNewWebQuery
values(128, 0) = "   xlDialogNote    ": values(128, 1) = "   [コメントの挿入] ダイアログ ボックス    ": values(128, 2) = xlDialogNote
values(129, 0) = "   xlDialogObjectProperties    ": values(129, 1) = "   [オブジェクトのプロパティ] ダイアログ ボックス  ": values(129, 2) = xlDialogObjectProperties
values(130, 0) = "   xlDialogObjectProtection    ": values(130, 1) = "   [オブジェクトの保護] ダイアログ ボックス    ": values(130, 2) = xlDialogObjectProtection
values(131, 0) = "   xlDialogOpen    ": values(131, 1) = "   [ファイルを開く] ダイアログ ボックス    ": values(131, 2) = xlDialogOpen
values(132, 0) = "   xlDialogOpenLinks   ": values(132, 1) = "   [リンク元を開く] ダイアログ ボックス    ": values(132, 2) = xlDialogOpenLinks
values(133, 0) = "   xlDialogOpenMail    ": values(133, 1) = "   [メールを開く] ダイアログ ボックス  ": values(133, 2) = xlDialogOpenMail
values(134, 0) = "   xlDialogOpenText    ": values(134, 1) = "   [テキストを開く] ダイアログ ボックス    ": values(134, 2) = xlDialogOpenText
values(135, 0) = "   xlDialogOptionsCalculation  ": values(135, 1) = "   [オプション (計算方法)] ダイアログ ボックス ": values(135, 2) = xlDialogOptionsCalculation
values(136, 0) = "   xlDialogOptionsChart    ": values(136, 1) = "   [オプション (グラフ)] ダイアログ ボックス   ": values(136, 2) = xlDialogOptionsChart
values(137, 0) = "   xlDialogOptionsEdit ": values(137, 1) = "   [オプション (編集)] ダイアログ ボックス ": values(137, 2) = xlDialogOptionsEdit
values(138, 0) = "   xlDialogOptionsGeneral  ": values(138, 1) = "   [オプション (全般)] ダイアログ ボックス ": values(138, 2) = xlDialogOptionsGeneral
values(139, 0) = "   xlDialogOptionsListsAdd ": values(139, 1) = "   [オプション (ユーザー設定リスト)] ダイアログ ボックス   ": values(139, 2) = xlDialogOptionsListsAdd
values(140, 0) = "   xlDialogOptionsME   ": values(140, 1) = "   [オプション (インターナショナル)] ダイアログ ボックス   ": values(140, 2) = xlDialogOptionsME
values(141, 0) = "   xlDialogOptionsTransition   ": values(141, 1) = "   [オプション (移行)] ダイアログ ボックス ": values(141, 2) = xlDialogOptionsTransition
values(142, 0) = "   xlDialogOptionsView ": values(142, 1) = "   [オプション (表示)] ダイアログ ボックス ": values(142, 2) = xlDialogOptionsView
values(143, 0) = "   xlDialogOutline ": values(143, 1) = "   [設定] ダイアログ ボックス  ": values(143, 2) = xlDialogOutline
values(144, 0) = "   xlDialogOverlay ": values(144, 1) = "   [重ね合わせグラフ] ダイアログ ボックス  ": values(144, 2) = xlDialogOverlay
values(145, 0) = "   xlDialogOverlayChartType    ": values(145, 1) = "   [グラフの種類の重ね合わせ] ダイアログ ボックス  ": values(145, 2) = xlDialogOverlayChartType
values(146, 0) = "   xlDialogPageSetup   ": values(146, 1) = "   [ページ設定 (ページ)] ダイアログ ボックス   ": values(146, 2) = xlDialogPageSetup
values(147, 0) = "   xlDialogParse   ": values(147, 1) = "   [区切り位置] ダイアログ ボックス    ": values(147, 2) = xlDialogParse
values(148, 0) = "   xlDialogPasteNames  ": values(148, 1) = "   [名前の貼り付け] ダイアログ ボックス    ": values(148, 2) = xlDialogPasteNames
values(149, 0) = "   xlDialogPasteSpecial    ": values(149, 1) = "   [形式を選択して貼り付け] ダイアログ ボックス    ": values(149, 2) = xlDialogPasteSpecial
values(150, 0) = "   xlDialogPatterns    ": values(150, 1) = "   [セルの書式設定 (パターン)] ダイアログ ボックス ": values(150, 2) = xlDialogPatterns
values(151, 0) = "   xlDialogPermission  ": values(151, 1) = "   [アクセス許可] ダイアログ ボックス  ": values(151, 2) = xlDialogPermission
values(152, 0) = "   xlDialogPhonetic    ": values(152, 1) = "   [ふりがなの設定 (ふりがな)] ダイアログ ボックス ": values(152, 2) = xlDialogPhonetic
values(153, 0) = "   xlDialogPivotCalculatedField    ": values(153, 1) = "   [ピボット集計フィールド] ダイアログ ボックス    ": values(153, 2) = xlDialogPivotCalculatedField
values(154, 0) = "   xlDialogPivotCalculatedItem ": values(154, 1) = "   [ピボット集計アイテム] ダイアログ ボックス  ": values(154, 2) = xlDialogPivotCalculatedItem
values(155, 0) = "   xlDialogPivotClientServerSet    ": values(155, 1) = "   [ピボット クライアント サーバー セット] ダイアログ ボックス ": values(155, 2) = xlDialogPivotClientServerSet
values(156, 0) = "   xlDialogPivotFieldGroup ": values(156, 1) = "   [ピボット フィールド グループ] ダイアログ ボックス  ": values(156, 2) = xlDialogPivotFieldGroup
values(157, 0) = "   xlDialogPivotFieldProperties    ": values(157, 1) = "   [ピボット フィールド プロパティ] ダイアログ ボックス    ": values(157, 2) = xlDialogPivotFieldProperties
values(158, 0) = "   xlDialogPivotFieldUngroup   ": values(158, 1) = "   [ピボット フィールド グループ解除] ダイアログ ボックス  ": values(158, 2) = xlDialogPivotFieldUngroup
values(159, 0) = "   xlDialogPivotShowPages  ": values(159, 1) = "   [ピボット表示ページ] ダイアログ ボックス    ": values(159, 2) = xlDialogPivotShowPages
values(160, 0) = "   xlDialogPivotSolveOrder ": values(160, 1) = "   [ピボット解決順序] ダイアログ ボックス  ": values(160, 2) = xlDialogPivotSolveOrder
values(161, 0) = "   xlDialogPivotTableOptions   ": values(161, 1) = "   [ピボットテーブル オプション] ダイアログ ボックス   ": values(161, 2) = xlDialogPivotTableOptions
values(162, 0) = "   xlDialogPivotTableWizard    ": values(162, 1) = "   [ピボットテーブル/ピボットグラフ ウィザード] ダイアログ ボックス    ": values(162, 2) = xlDialogPivotTableWizard
values(163, 0) = "   xlDialogPlacement   ": values(163, 1) = "   [表示位置] ダイアログ ボックス  ": values(163, 2) = xlDialogPlacement
values(164, 0) = "   xlDialogPrint   ": values(164, 1) = "   [印刷] ダイアログ ボックス  ": values(164, 2) = xlDialogPrint
values(165, 0) = "   xlDialogPrinterSetup    ": values(165, 1) = "   [プリンターの設定] ダイアログ ボックス  ": values(165, 2) = xlDialogPrinterSetup
values(166, 0) = "   xlDialogPrintPreview    ": values(166, 1) = "   [印刷プレビュー] ダイアログ ボックス    ": values(166, 2) = xlDialogPrintPreview
values(167, 0) = "   xlDialogPromote ": values(167, 1) = "   [グループの解除] ダイアログ ボックス    ": values(167, 2) = xlDialogPromote
values(168, 0) = "   xlDialogProperties  ": values(168, 1) = "   [プロパティ (ファイルの概要)] ダイアログ ボックス   ": values(168, 2) = xlDialogProperties
values(169, 0) = "   xlDialogPropertyFields  ": values(169, 1) = "   [プロパティ フィールド] ダイアログ ボックス ": values(169, 2) = xlDialogPropertyFields
values(170, 0) = "   xlDialogProtectDocument ": values(170, 1) = "   [シートの保護] ダイアログ ボックス  ": values(170, 2) = xlDialogProtectDocument
values(171, 0) = "   xlDialogProtectSharing  ": values(171, 1) = "   [共有ブックの保護] ダイアログ ボックス  ": values(171, 2) = xlDialogProtectSharing
values(172, 0) = "   xlDialogPublishAsWebPage    ": values(172, 1) = "   [Web ページとして発行] ダイアログ ボックス  ": values(172, 2) = xlDialogPublishAsWebPage
values(173, 0) = "   xlDialogPushbuttonProperties    ": values(173, 1) = "   [プッシュ ボタンのプロパティ] ダイアログ ボックス   ": values(173, 2) = xlDialogPushbuttonProperties
values(174, 0) = "   xlDialogReplaceFont ": values(174, 1) = "   [フォントの設定] ダイアログ ボックス    ": values(174, 2) = xlDialogReplaceFont
values(175, 0) = "   xlDialogRoutingSlip ": values(175, 1) = "   [回覧先] ダイアログ ボックス    ": values(175, 2) = xlDialogRoutingSlip
values(176, 0) = "   xlDialogRowHeight   ": values(176, 1) = "   [行の高さ] ダイアログ ボックス  ": values(176, 2) = xlDialogRowHeight
values(177, 0) = "   xlDialogRun ": values(177, 1) = "   [マクロ] ダイアログ ボックス    ": values(177, 2) = xlDialogRun
values(178, 0) = "   xlDialogSaveAs  ": values(178, 1) = "   [名前を付けて保存] ダイアログ ボックス  ": values(178, 2) = xlDialogSaveAs
values(179, 0) = "   xlDialogSaveCopyAs  ": values(179, 1) = "   [コピーを名前を付けて保存] ダイアログ ボックス  ": values(179, 2) = xlDialogSaveCopyAs
values(180, 0) = "   xlDialogSaveNewObject   ": values(180, 1) = "   「新しいオブジェクトの保存] ダイアログ ボックス ": values(180, 2) = xlDialogSaveNewObject
values(181, 0) = "   xlDialogSaveWorkbook    ": values(181, 1) = "   [名前を付けて保存] ダイアログ ボックス  ": values(181, 2) = xlDialogSaveWorkbook
values(182, 0) = "   xlDialogSaveWorkspace   ": values(182, 1) = "   [作業状態の保存] ダイアログ ボックス    ": values(182, 2) = xlDialogSaveWorkspace
values(183, 0) = "   xlDialogScale   ": values(183, 1) = "   [倍率] ダイアログ ボックス  ": values(183, 2) = xlDialogScale
values(184, 0) = "   xlDialogScenarioAdd ": values(184, 1) = "   [シナリオの追加] ダイアログ ボックス    ": values(184, 2) = xlDialogScenarioAdd
values(185, 0) = "   xlDialogScenarioCells   ": values(185, 1) = "   [シナリオの登録と管理] ダイアログ ボックス  ": values(185, 2) = xlDialogScenarioCells
values(186, 0) = "   xlDialogScenarioEdit    ": values(186, 1) = "   [シナリオの追加] ダイアログ ボックス    ": values(186, 2) = xlDialogScenarioEdit
values(187, 0) = "   xlDialogScenarioMerge   ": values(187, 1) = "   [シナリオのコピー] ダイアログ ボックス  ": values(187, 2) = xlDialogScenarioMerge
values(188, 0) = "   xlDialogScenarioSummary ": values(188, 1) = "   [シナリオの情報] ダイアログ ボックス    ": values(188, 2) = xlDialogScenarioSummary
values(189, 0) = "   xlDialogScrollbarProperties ": values(189, 1) = "   [スクロール バーのプロパティ] ダイアログ ボックス   ": values(189, 2) = xlDialogScrollbarProperties
values(190, 0) = "   xlDialogSearch  ": values(190, 1) = "   [通常のファイル検索] ダイアログ ボックス    ": values(190, 2) = xlDialogSearch
values(191, 0) = "   xlDialogSelectSpecial   ": values(191, 1) = "   選択オプション] ダイアログ ボックス ": values(191, 2) = xlDialogSelectSpecial
values(192, 0) = "   xlDialogSendMail    ": values(192, 1) = "   [メッセージ (HTML 形式)] ダイアログ ボックス    ": values(192, 2) = xlDialogSendMail
values(193, 0) = "   xlDialogSeriesAxes  ": values(193, 1) = "   [系列軸] ダイアログ ボックス    ": values(193, 2) = xlDialogSeriesAxes
values(194, 0) = "   xlDialogSeriesOptions   ": values(194, 1) = "   [系列オプション] ダイアログ ボックス    ": values(194, 2) = xlDialogSeriesOptions
values(195, 0) = "   xlDialogSeriesOrder ": values(195, 1) = "   [系列の順序] ダイアログ ボックス    ": values(195, 2) = xlDialogSeriesOrder
values(196, 0) = "   xlDialogSeriesShape ": values(196, 1) = "   [系列の形状] ダイアログ ボックス    ": values(196, 2) = xlDialogSeriesShape
values(197, 0) = "   xlDialogSeriesX ": values(197, 1) = "   [系列 X] ダイアログ ボックス    ": values(197, 2) = xlDialogSeriesX
values(198, 0) = "   xlDialogSeriesY ": values(198, 1) = "   [データ系列の書式設定 (名前/値)] ダイアログ ボックス    ": values(198, 2) = xlDialogSeriesY
values(199, 0) = "   xlDialogSetBackgroundPicture    ": values(199, 1) = "   [シートの背景] ダイアログ ボックス  ": values(199, 2) = xlDialogSetBackgroundPicture
values(200, 0) = "   xlDialogSetPrintTitles  ": values(200, 1) = "   [印刷タイトルの設定] ダイアログ ボックス    ": values(200, 2) = xlDialogSetPrintTitles
values(201, 0) = "   xlDialogSetUpdateStatus ": values(201, 1) = "   [更新状態の設定] ダイアログ ボックス    ": values(201, 2) = xlDialogSetUpdateStatus
values(202, 0) = "   xlDialogShowDetail  ": values(202, 1) = "   [詳細データの表示] ダイアログ ボックス  ": values(202, 2) = xlDialogShowDetail
values(203, 0) = "   xlDialogShowToolbar ": values(203, 1) = "   [ユーザー設定 (オプション)] ダイアログ ボックス ": values(203, 2) = xlDialogShowToolbar
values(204, 0) = "   xlDialogSize    ": values(204, 1) = "   [サイズ] ダイアログ ボックス    ": values(204, 2) = xlDialogSize
values(205, 0) = "   xlDialogSort    ": values(205, 1) = "   [並べ替え] ダイアログ ボックス  ": values(205, 2) = xlDialogSort
values(206, 0) = "   xlDialogSortSpecial ": values(206, 1) = "   [並べ替え] ダイアログ ボックス  ": values(206, 2) = xlDialogSortSpecial
values(207, 0) = "   xlDialogSplit   ": values(207, 1) = "   [列の分割、 行の分割] ダイアログ ボックス   ": values(207, 2) = xlDialogSplit
values(208, 0) = "   xlDialogStandardFont    ": values(208, 1) = "   [フォントの設定] ダイアログ ボックス    ": values(208, 2) = xlDialogStandardFont
values(209, 0) = "   xlDialogStandardWidth   ": values(209, 1) = "   [標準の幅] ダイアログ ボックス  ": values(209, 2) = xlDialogStandardWidth
values(210, 0) = "   xlDialogStyle   ": values(210, 1) = "   [フォントの設定] ダイアログ ボックス    ": values(210, 2) = xlDialogStyle
values(211, 0) = "   xlDialogSubscribeTo ": values(211, 1) = "   [引用] ダイアログ ボックス  ": values(211, 2) = xlDialogSubscribeTo
values(212, 0) = "   xlDialogSubtotalCreate  ": values(212, 1) = "   [集計の設定] ダイアログ ボックス    ": values(212, 2) = xlDialogSubtotalCreate
values(213, 0) = "   xlDialogSummaryInfo ": values(213, 1) = "   [プロパティ (ファイルの概要)] ダイアログ ボックス   ": values(213, 2) = xlDialogSummaryInfo
values(214, 0) = "   xlDialogTable   ": values(214, 1) = "   [テーブル] ダイアログ ボックス  ": values(214, 2) = xlDialogTable
values(215, 0) = "   xlDialogTabOrder    ": values(215, 1) = "   [タブ オーダーの設定] ダイアログ ボックス   ": values(215, 2) = xlDialogTabOrder
values(216, 0) = "   xlDialogTextToColumns   ": values(216, 1) = "   [区切り位置] ダイアログ ボックス    ": values(216, 2) = xlDialogTextToColumns
values(217, 0) = "   xlDialogUnhide  ": values(217, 1) = "   [ウィンドウの再表示] ダイアログ ボックス    ": values(217, 2) = xlDialogUnhide
values(218, 0) = "   xlDialogUpdateLink  ": values(218, 1) = "   [リンクの更新] ダイアログ ボックス  ": values(218, 2) = xlDialogUpdateLink
values(219, 0) = "   xlDialogVbaInsertFile   ": values(219, 1) = "   [VBA 挿入ファイル] ダイアログ ボックス  ": values(219, 2) = xlDialogVbaInsertFile
values(220, 0) = "   xlDialogVbaMakeAddin    ": values(220, 1) = "   [VBA 作成アドイン] ダイアログ ボックス  ": values(220, 2) = xlDialogVbaMakeAddin
values(221, 0) = "   xlDialogVbaProcedureDefinition  ": values(221, 1) = "   [VBA 手順定義] ダイアログ ボックス  ": values(221, 2) = xlDialogVbaProcedureDefinition
values(222, 0) = "   xlDialogView3d  ": values(222, 1) = "   [3D 表示] ダイアログ ボックス   ": values(222, 2) = xlDialogView3d
values(223, 0) = "   xlDialogWebOptionsBrowsers  ": values(223, 1) = "   [Web オプション (ブラウザー)] ダイアログ ボックス   ": values(223, 2) = xlDialogWebOptionsBrowsers
values(224, 0) = "   xlDialogWebOptionsEncoding  ": values(224, 1) = "   [Web オプション (エンコード)] ダイアログ ボックス   ": values(224, 2) = xlDialogWebOptionsEncoding
values(225, 0) = "   xlDialogWebOptionsFiles ": values(225, 1) = "   [Web オプション (ファイル)] ダイアログ ボックス ": values(225, 2) = xlDialogWebOptionsFiles
values(226, 0) = "   xlDialogWebOptionsFonts ": values(226, 1) = "   [Web オプション (フォント)] ダイアログ ボックス ": values(226, 2) = xlDialogWebOptionsFonts
values(227, 0) = "   xlDialogWebOptionsGeneral   ": values(227, 1) = "   [Web オプション (全般)] ダイアログ ボックス ": values(227, 2) = xlDialogWebOptionsGeneral
values(228, 0) = "   xlDialogWebOptionsPictures  ": values(228, 1) = "   [Web オプション (図)] ダイアログ ボックス   ": values(228, 2) = xlDialogWebOptionsPictures
values(229, 0) = "   xlDialogWindowMove  ": values(229, 1) = "   [ウィンドウの移動] ダイアログ ボックス  ": values(229, 2) = xlDialogWindowMove
values(230, 0) = "   xlDialogWindowSize  ": values(230, 1) = "   [ウィンドウ サイズ] ダイアログ ボックス ": values(230, 2) = xlDialogWindowSize
values(231, 0) = "   xlDialogWorkbookAdd ": values(231, 1) = "   [シートの移動またはコピー] ダイアログ ボックス  ": values(231, 2) = xlDialogWorkbookAdd
values(232, 0) = "   xlDialogWorkbookCopy    ": values(232, 1) = "   [シートの移動またはコピー] ダイアログ ボックス  ": values(232, 2) = xlDialogWorkbookCopy
values(233, 0) = "   xlDialogWorkbookInsert  ": values(233, 1) = "   [挿入 (標準)] ダイアログ ボックス   ": values(233, 2) = xlDialogWorkbookInsert
values(234, 0) = "   xlDialogWorkbookMove    ": values(234, 1) = "   [シートの移動またはコピー] ダイアログ ボックス  ": values(234, 2) = xlDialogWorkbookMove
values(235, 0) = "   xlDialogWorkbookName    ": values(235, 1) = "   [シート名の変更] ダイアログ ボックス    ": values(235, 2) = xlDialogWorkbookName
values(236, 0) = "   xlDialogWorkbookNew ": values(236, 1) = "   [挿入 (標準)] ダイアログ ボックス   ": values(236, 2) = xlDialogWorkbookNew
values(237, 0) = "   xlDialogWorkbookOptions ": values(237, 1) = "   [シート名の変更] ダイアログ ボックス    ": values(237, 2) = xlDialogWorkbookOptions
values(238, 0) = "   xlDialogWorkbookProtect ": values(238, 1) = "   [ブックの保護] ダイアログ ボックス  ": values(238, 2) = xlDialogWorkbookProtect
values(239, 0) = "   xlDialogWorkbookTabSplit    ": values(239, 1) = "   [ブックのタブ分割] ダイアログ ボックス  ": values(239, 2) = xlDialogWorkbookTabSplit
values(240, 0) = "   xlDialogWorkbookUnhide  ": values(240, 1) = "   [再表示] ダイアログ ボックス    ": values(240, 2) = xlDialogWorkbookUnhide
values(241, 0) = "   xlDialogWorkgroup   ": values(241, 1) = "   [グループ編集] ダイアログ ボックス  ": values(241, 2) = xlDialogWorkgroup
values(242, 0) = "   xlDialogWorkspace   ": values(242, 1) = "   [作業状態設定] ダイアログ ボックス  ": values(242, 2) = xlDialogWorkspace
values(243, 0) = "   xlDialogZoom    ": values(243, 1) = "   [ズーム] ダイアログ ボックス    ": values(243, 2) = xlDialogZoom
     'リスト ボックスには 3 つのデータ列が含まれます。
    ListBox1.ColumnCount = 2

    'ListBox1 および ListBox2 にデータを読み込みます。
    ListBox1.List() = values

End Sub
