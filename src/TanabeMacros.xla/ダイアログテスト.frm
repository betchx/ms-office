VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} �_�C�A���O�e�X�g 
   Caption         =   "�_�C�A���O�\���e�X�g"
   ClientHeight    =   5835
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   10185
   OleObjectBlob   =   "�_�C�A���O�e�X�g.frx":0000
   StartUpPosition =   1  '�I�[�i�[ �t�H�[���̒���
End
Attribute VB_Name = "�_�C�A���O�e�X�g"
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
    MsgBox "�_�C�A���O���擾�ł��܂���ł���"
Else
    rtn = d.Show()
End If
Exit Sub

eee:
 MsgBox "�G���[���������\���ł��܂���ł����D"
 
End Sub

Private Sub UserForm_Initialize()

values(0, 0) = "   xlDialogActivate    ":   values(0, 1) = "   [�E�B���h�E�̑I��] �_�C�A���O �{�b�N�X  ":   values(0, 2) = xlDialogActivate
values(1, 0) = "   xlDialogActiveCellFont  ":   values(1, 1) = "   [�Z���̏����ݒ� (�t�H���g)] �_�C�A���O �{�b�N�X ":   values(1, 2) = xlDialogActiveCellFont
values(2, 0) = "   xlDialogAddChartAutoformat  ":   values(2, 1) = "   [���[�U�[�ݒ�̃O���t��ނ̒ǉ�] �_�C�A���O �{�b�N�X    ":   values(2, 2) = xlDialogAddChartAutoformat
values(3, 0) = "   xlDialogAddinManager    ":   values(3, 1) = "   [�A�h�C��] �_�C�A���O �{�b�N�X  ":   values(3, 2) = xlDialogAddinManager
values(4, 0) = "   xlDialogAlignment   ":   values(4, 1) = "   [�Z���̏����ݒ� (�z�u)] �_�C�A���O �{�b�N�X ":   values(4, 2) = xlDialogAlignment
values(5, 0) = "   xlDialogApplyNames  ":   values(5, 1) = "   [���O�̈��p] �_�C�A���O �{�b�N�X    ":   values(5, 2) = xlDialogApplyNames
values(6, 0) = "   xlDialogApplyStyle  ":   values(6, 1) = "   [�X�^�C��] �_�C�A���O �{�b�N�X  ":   values(6, 2) = xlDialogApplyStyle
values(7, 0) = "   xlDialogAppMove ":   values(7, 1) = "   [�ړ�(�A�v���P�[�V����)] �_�C�A���O �{�b�N�X    ":   values(7, 2) = xlDialogAppMove
values(8, 0) = "   xlDialogAppSize ":   values(8, 1) = "   [���M] �_�C�A���O �{�b�N�X  ":   values(8, 2) = xlDialogAppSize
values(9, 0) = "   xlDialogArrangeAll  ":   values(9, 1) = "   [�E�B���h�E�̐���] �_�C�A���O �{�b�N�X  ":   values(9, 2) = xlDialogArrangeAll
values(10, 0) = "   xlDialogAssignToObject  ":  values(10, 1) = "   [�I�u�W�F�N�g�ւ̓o�^] �_�C�A���O �{�b�N�X  ":  values(10, 2) = xlDialogAssignToObject
values(11, 0) = "   xlDialogAssignToTool    ":  values(11, 1) = "   [�c�[���Ɋ��蓖��] �_�C�A���O �{�b�N�X  ":  values(11, 2) = xlDialogAssignToTool
values(12, 0) = "   xlDialogAttachText  ":  values(12, 1) = "   [�����̒ǉ�] �_�C�A���O �{�b�N�X    ":  values(12, 2) = xlDialogAttachText
values(13, 0) = "   xlDialogAttachToolbars  ":  values(13, 1) = "   [�u�b�N�ւ̃c�[���o�[�̓o�^] �_�C�A���O �{�b�N�X    ":  values(13, 2) = xlDialogAttachToolbars
values(14, 0) = "   xlDialogAutoCorrect ":  values(14, 1) = "   [�I�[�g�R���N�g (�I�[�g�R���N�g)] �_�C�A���O �{�b�N�X   ":  values(14, 2) = xlDialogAutoCorrect
values(15, 0) = "   xlDialogAxes    ":  values(15, 1) = "   [��] �_�C�A���O �{�b�N�X    ":  values(15, 2) = xlDialogAxes
values(16, 0) = "   xlDialogBorder  ":  values(16, 1) = "   [�Z���̏����ݒ� (�r��)] �_�C�A���O �{�b�N�X ":  values(16, 2) = xlDialogBorder
values(17, 0) = "   xlDialogCalculation ":  values(17, 1) = "   [�v�Z���@�̐ݒ�] �_�C�A���O �{�b�N�X    ":  values(17, 2) = xlDialogCalculation
values(18, 0) = "   xlDialogCellProtection  ":  values(18, 1) = "   [�Z���̏����ݒ� (�ی�)] �_�C�A���O �{�b�N�X ":  values(18, 2) = xlDialogCellProtection
values(19, 0) = "   xlDialogChangeLink  ":  values(19, 1) = "   [�����N�̕ύX] �_�C�A���O �{�b�N�X  ":  values(19, 2) = xlDialogChangeLink
values(20, 0) = "   xlDialogChartAddData    ":  values(20, 1) = "   [�O���t�ǉ��f�[�^] �_�C�A���O �{�b�N�X  ":  values(20, 2) = xlDialogChartAddData
values(21, 0) = "   xlDialogChartLocation   ":  values(21, 1) = "   [�O���t�̏ꏊ] �_�C�A���O �{�b�N�X  ":  values(21, 2) = xlDialogChartLocation
values(22, 0) = "   xlDialogChartOptionsDataLabelMultiple   ":  values(22, 1) = "   [�O���t �I�v�V���� �f�[�^ ���x������] �_�C�A���O �{�b�N�X   ":  values(22, 2) = xlDialogChartOptionsDataLabelMultiple
values(23, 0) = "   xlDialogChartOptionsDataLabels  ":  values(23, 1) = "   [�O���t �I�v�V���� �f�[�^ ���x��] �_�C�A���O �{�b�N�X   ":  values(23, 2) = xlDialogChartOptionsDataLabels
values(24, 0) = "   xlDialogChartOptionsDataTable   ":  values(24, 1) = "   [�O���t �I�v�V���� �f�[�^ �e�[�u��] �_�C�A���O �{�b�N�X ":  values(24, 2) = xlDialogChartOptionsDataTable
values(25, 0) = "   xlDialogChartSourceData ":  values(25, 1) = "   [�O���t�̌��f�[�^] �_�C�A���O �{�b�N�X  ":  values(25, 2) = xlDialogChartSourceData
values(26, 0) = "   xlDialogChartTrend  ":  values(26, 1) = "   [�O���t �g�����h] �_�C�A���O �{�b�N�X   ":  values(26, 2) = xlDialogChartTrend
values(27, 0) = "   xlDialogChartType   ":  values(27, 1) = "   [�O���t�̎��] �_�C�A���O �{�b�N�X  ":  values(27, 2) = xlDialogChartType
values(28, 0) = "   xlDialogChartWizard ":  values(28, 1) = "   [�O���t �E�B�U�[�h] �_�C�A���O �{�b�N�X ":  values(28, 2) = xlDialogChartWizard
values(29, 0) = "   xlDialogCheckboxProperties  ":  values(29, 1) = "   [�`�F�b�N �{�b�N�X�̃v���p�e�B] �_�C�A���O �{�b�N�X ":  values(29, 2) = xlDialogCheckboxProperties
values(30, 0) = "   xlDialogClear   ":  values(30, 1) = "   [����] �_�C�A���O �{�b�N�X  ":  values(30, 2) = xlDialogClear
values(31, 0) = "   xlDialogColorPalette    ":  values(31, 1) = "   [�I�v�V���� (�F)] �_�C�A���O �{�b�N�X   ":  values(31, 2) = xlDialogColorPalette
values(32, 0) = "   xlDialogColumnWidth ":  values(32, 1) = "   [��] �_�C�A���O �{�b�N�X  ":  values(32, 2) = xlDialogColumnWidth
values(33, 0) = "   xlDialogCombination ":  values(33, 1) = "   [����] �_�C�A���O �{�b�N�X  ":  values(33, 2) = xlDialogCombination
values(34, 0) = "   xlDialogConditionalFormatting   ":  values(34, 1) = "   [�����t�������̐ݒ�] �_�C�A���O �{�b�N�X    ":  values(34, 2) = xlDialogConditionalFormatting
values(35, 0) = "   xlDialogConsolidate ":  values(35, 1) = "   [�����̐ݒ�] �_�C�A���O �{�b�N�X    ":  values(35, 2) = xlDialogConsolidate
values(36, 0) = "   xlDialogCopyChart   ":  values(36, 1) = "   [�O���t�̃R�s�[] �_�C�A���O �{�b�N�X    ":  values(36, 2) = xlDialogCopyChart
values(37, 0) = "   xlDialogCopyPicture ":  values(37, 1) = "   [�}�̃R�s�[] �_�C�A���O �{�b�N�X    ":  values(37, 2) = xlDialogCopyPicture
values(38, 0) = "   xlDialogCreateList  ":  values(38, 1) = "   [���X�g�̍쐬 ] �_�C�A���O �{�b�N�X ":  values(38, 2) = xlDialogCreateList
values(39, 0) = "   xlDialogCreateNames ":  values(39, 1) = "   [���O�̍쐬] �_�C�A���O �{�b�N�X    ":  values(39, 2) = xlDialogCreateNames
values(40, 0) = "   xlDialogCreatePublisher ":  values(40, 1) = "   [���s���̍쐬] �_�C�A���O �{�b�N�X  ":  values(40, 2) = xlDialogCreatePublisher
values(41, 0) = "   xlDialogCustomizeToolbar    ":  values(41, 1) = "   [���[�U�[�ݒ� (�I�v�V����)] �_�C�A���O �{�b�N�X ":  values(41, 2) = xlDialogCustomizeToolbar
values(42, 0) = "   xlDialogCustomViews ":  values(42, 1) = "   [���[�U�[�ݒ�̃r���[] �_�C�A���O �{�b�N�X  ":  values(42, 2) = xlDialogCustomViews
values(43, 0) = "   xlDialogDataDelete  ":  values(43, 1) = "   [�f�[�^�̍폜] �_�C�A���O �{�b�N�X  ":  values(43, 2) = xlDialogDataDelete
values(44, 0) = "   xlDialogDataLabel   ":  values(44, 1) = "   [�f�[�^ ���x��] �_�C�A���O �{�b�N�X ":  values(44, 2) = xlDialogDataLabel
values(45, 0) = "   xlDialogDataLabelMultiple   ":  values(45, 1) = "   [�f�[�^ ���x������] �_�C�A���O �{�b�N�X ":  values(45, 2) = xlDialogDataLabelMultiple
values(46, 0) = "   xlDialogDataSeries  ":  values(46, 1) = "   [�A���f�[�^] �_�C�A���O �{�b�N�X    ":  values(46, 2) = xlDialogDataSeries
values(47, 0) = "   xlDialogDataValidation  ":  values(47, 1) = "   [�f�[�^�̓��͋K�� (�ݒ�)] �_�C�A���O �{�b�N�X   ":  values(47, 2) = xlDialogDataValidation
values(48, 0) = "   xlDialogDefineName  ":  values(48, 1) = "   [���O�̒�`] �_�C�A���O �{�b�N�X    ":  values(48, 2) = xlDialogDefineName
values(49, 0) = "   xlDialogDefineStyle ":  values(49, 1) = "   [�X�^�C��] �_�C�A���O �{�b�N�X  ":  values(49, 2) = xlDialogDefineStyle
values(50, 0) = "   xlDialogDeleteFormat    ":  values(50, 1) = "   [�Z���̏����ݒ� (�\���`��)] �_�C�A���O �{�b�N�X ":  values(50, 2) = xlDialogDeleteFormat
values(51, 0) = "   xlDialogDeleteName  ":  values(51, 1) = "   [���O�̒�`] �_�C�A���O �{�b�N�X    ":  values(51, 2) = xlDialogDeleteName
values(52, 0) = "   xlDialogDemote  ":  values(52, 1) = "   [�O���[�v��] �_�C�A���O �{�b�N�X    ":  values(52, 2) = xlDialogDemote
values(53, 0) = "   xlDialogDisplay ":  values(53, 1) = "   [��ʐݒ�] �_�C�A���O �{�b�N�X  ":  values(53, 2) = xlDialogDisplay
values(54, 0) = "   xlDialogDocumentInspector   ":  values(54, 1) = "(x)[�h�L�������g����] �_�C�A���O �{�b�N�X  ":  values(54, 2) = 0 'xlDialogDocumentInspector
values(55, 0) = "   xlDialogEditboxProperties   ":  values(55, 1) = "   [�ҏW�{�b�N�X�̃v���p�e�B] �_�C�A���O �{�b�N�X  ":  values(55, 2) = xlDialogEditboxProperties
values(56, 0) = "   xlDialogEditColor   ":  values(56, 1) = "   [�F�̕ҏW] �_�C�A���O �{�b�N�X  ":  values(56, 2) = xlDialogEditColor
values(57, 0) = "   xlDialogEditDelete  ":  values(57, 1) = "   [�폜] �_�C�A���O �{�b�N�X  ":  values(57, 2) = xlDialogEditDelete
values(58, 0) = "   xlDialogEditionOptions  ":  values(58, 1) = "   [�G�f�B�V���� �I�v�V����] �_�C�A���O �{�b�N�X   ":  values(58, 2) = xlDialogEditionOptions
values(59, 0) = "   xlDialogEditSeries  ":  values(59, 1) = "   [�n��̕ҏW] �_�C�A���O �{�b�N�X    ":  values(59, 2) = xlDialogEditSeries
values(60, 0) = "   xlDialogErrorbarX   ":  values(60, 1) = "   [Errorbar X] �_�C�A���O �{�b�N�X    ":  values(60, 2) = xlDialogErrorbarX
values(61, 0) = "   xlDialogErrorbarY   ":  values(61, 1) = "   [Errorbar Y] �_�C�A���O �{�b�N�X    ":  values(61, 2) = xlDialogErrorbarY
values(62, 0) = "   xlDialogErrorChecking   ":  values(62, 1) = "   [�G���[ �`�F�b�N] �_�C�A���O �{�b�N�X   ":  values(62, 2) = xlDialogErrorChecking
values(63, 0) = "   xlDialogEvaluateFormula ":  values(63, 1) = "   [�����̌���] �_�C�A���O �{�b�N�X    ":  values(63, 2) = xlDialogEvaluateFormula
values(64, 0) = "   xlDialogExternalDataProperties  ":  values(64, 1) = "   [�O���f�[�^�̃v���p�e�B] �_�C�A���O �{�b�N�X    ":  values(64, 2) = xlDialogExternalDataProperties
values(65, 0) = "   xlDialogExtract ":  values(65, 1) = "   [���o] �_�C�A���O �{�b�N�X  ":  values(65, 2) = xlDialogExtract
values(66, 0) = "   xlDialogFileDelete  ":  values(66, 1) = "   [�t�@�C���̍폜] �_�C�A���O �{�b�N�X    ":  values(66, 2) = xlDialogFileDelete
values(67, 0) = "   xlDialogFileSharing ":  values(67, 1) = "   [�u�b�N�̋��L (�ҏW)] �_�C�A���O �{�b�N�X   ":  values(67, 2) = xlDialogFileSharing
values(68, 0) = "   xlDialogFillGroup   ":  values(68, 1) = "   [�O���[�v�̓���] �_�C�A���O �{�b�N�X    ":  values(68, 2) = xlDialogFillGroup
values(69, 0) = "   xlDialogFillWorkgroup   ":  values(69, 1) = "   [���[�N�O���[�v�̓���] �_�C�A���O �{�b�N�X  ":  values(69, 2) = xlDialogFillWorkgroup
values(70, 0) = "   xlDialogFilter  ":  values(70, 1) = "   [�I�[�g�t�B���^�[] �_�C�A���O �{�b�N�X  ":  values(70, 2) = xlDialogFilter
values(71, 0) = "   xlDialogFilterAdvanced  ":  values(71, 1) = "   [�t�B���^�[ �I�v�V�����̐ݒ�] �_�C�A���O �{�b�N�X   ":  values(71, 2) = xlDialogFilterAdvanced
values(72, 0) = "   xlDialogFindFile    ":  values(72, 1) = "   [�t�@�C�����J��] �_�C�A���O �{�b�N�X    ":  values(72, 2) = xlDialogFindFile
values(73, 0) = "   xlDialogFont    ":  values(73, 1) = "   [�t�H���g�̐ݒ�] �_�C�A���O �{�b�N�X    ":  values(73, 2) = xlDialogFont
values(74, 0) = "   xlDialogFontProperties  ":  values(74, 1) = "   [�Z���̏����ݒ� (�t�H���g)] �_�C�A���O �{�b�N�X ":  values(74, 2) = xlDialogFontProperties
values(75, 0) = "   xlDialogFormatAuto  ":  values(75, 1) = "   [�I�[�g�t�H�[�}�b�g] �_�C�A���O �{�b�N�X    ":  values(75, 2) = xlDialogFormatAuto
values(76, 0) = "   xlDialogFormatChart ":  values(76, 1) = "   [�O���t�̏����ݒ�] �_�C�A���O �{�b�N�X  ":  values(76, 2) = xlDialogFormatChart
values(77, 0) = "   xlDialogFormatCharttype ":  values(77, 1) = "   [�O���t�̎��] �_�C�A���O �{�b�N�X  ":  values(77, 2) = xlDialogFormatCharttype
values(78, 0) = "   xlDialogFormatFont  ":  values(78, 1) = "   [�t�H���g�̐ݒ�] �_�C�A���O �{�b�N�X    ":  values(78, 2) = xlDialogFormatFont
values(79, 0) = "   xlDialogFormatLegend    ":  values(79, 1) = "   [�}��̏����ݒ�] �_�C�A���O �{�b�N�X    ":  values(79, 2) = xlDialogFormatLegend
values(80, 0) = "   xlDialogFormatMain  ":  values(80, 1) = "   [���C���O���t/�d�ˍ��킹�O���t] �_�C�A���O �{�b�N�X ":  values(80, 2) = xlDialogFormatMain
values(81, 0) = "   xlDialogFormatMove  ":  values(81, 1) = "   [�ړ��̏����ݒ�] �_�C�A���O �{�b�N�X    ":  values(81, 2) = xlDialogFormatMove
values(82, 0) = "   xlDialogFormatNumber    ":  values(82, 1) = "   [�Z���̏����ݒ� (�\���`��)] �_�C�A���O �{�b�N�X ":  values(82, 2) = xlDialogFormatNumber
values(83, 0) = "   xlDialogFormatOverlay   ":  values(83, 1) = "   [�d�ˍ��킹�O���t�̐ݒ�] �_�C�A���O �{�b�N�X    ":  values(83, 2) = xlDialogFormatOverlay
values(84, 0) = "   xlDialogFormatSize  ":  values(84, 1) = "   [�T�C�Y�̏����ݒ�] �_�C�A���O �{�b�N�X  ":  values(84, 2) = xlDialogFormatSize
values(85, 0) = "   xlDialogFormatText  ":  values(85, 1) = "   [��������] �_�C�A���O �{�b�N�X  ":  values(85, 2) = xlDialogFormatText
values(86, 0) = "   xlDialogFormulaFind ":  values(86, 1) = "   [����] �_�C�A���O �{�b�N�X  ":  values(86, 2) = xlDialogFormulaFind
values(87, 0) = "   xlDialogFormulaGoto ":  values(87, 1) = "   [�W�����v] �_�C�A���O �{�b�N�X  ":  values(87, 2) = xlDialogFormulaGoto
values(88, 0) = "   xlDialogFormulaReplace  ":  values(88, 1) = "   [�u��] �_�C�A���O �{�b�N�X  ":  values(88, 2) = xlDialogFormulaReplace
values(89, 0) = "   xlDialogFunctionWizard  ":  values(89, 1) = "   [�֐��̑}��] �_�C�A���O �{�b�N�X    ":  values(89, 2) = xlDialogFunctionWizard
values(90, 0) = "   xlDialogGallery3dArea   ":  values(90, 1) = "   [�I�[�g�t�H�[�}�b�g (3-D ��)] �_�C�A���O �{�b�N�X   ":  values(90, 2) = xlDialogGallery3dArea
values(91, 0) = "   xlDialogGallery3dBar    ":  values(91, 1) = "   [�I�[�g�t�H�[�}�b�g (���_)] �_�C�A���O �{�b�N�X ":  values(91, 2) = xlDialogGallery3dBar
values(92, 0) = "   xlDialogGallery3dColumn ":  values(92, 1) = "   [�I�[�g�t�H�[�}�b�g (3-D �c�_)] �_�C�A���O �{�b�N�X ":  values(92, 2) = xlDialogGallery3dColumn
values(93, 0) = "   xlDialogGallery3dLine   ":  values(93, 1) = "   [�I�[�g�t�H�[�}�b�g (3-D �܂��)] �_�C�A���O �{�b�N�X   ":  values(93, 2) = xlDialogGallery3dLine
values(94, 0) = "   xlDialogGallery3dPie    ":  values(94, 1) = "   [�I�[�g�t�H�[�}�b�g (3-D �~)] �_�C�A���O �{�b�N�X   ":  values(94, 2) = xlDialogGallery3dPie
values(95, 0) = "   xlDialogGallery3dSurface    ":  values(95, 1) = "   [�I�[�g�t�H�[�}�b�g (������)] �_�C�A���O �{�b�N�X   ":  values(95, 2) = xlDialogGallery3dSurface
values(96, 0) = "   xlDialogGalleryArea ":  values(96, 1) = "   [�I�[�g�t�H�[�}�b�g (��)] �_�C�A���O �{�b�N�X   ":  values(96, 2) = xlDialogGalleryArea
values(97, 0) = "   xlDialogGalleryBar  ":  values(97, 1) = "   [�I�[�g�t�H�[�}�b�g (���_)] �_�C�A���O �{�b�N�X ":  values(97, 2) = xlDialogGalleryBar
values(98, 0) = "   xlDialogGalleryColumn   ":  values(98, 1) = "   [�I�[�g�t�H�[�}�b�g (�c�_)] �_�C�A���O �{�b�N�X ":  values(98, 2) = xlDialogGalleryColumn
values(99, 0) = "   xlDialogGalleryCustom   ":  values(99, 1) = "   [�I�[�g�t�H�[�}�b�g (�t�H�[�}�b�g�̎��)] �_�C�A���O �{�b�N�X   ":  values(99, 2) = xlDialogGalleryCustom
values(100, 0) = "   xlDialogGalleryDoughnut ": values(100, 1) = "   [�I�[�g�t�H�[�}�b�g (�h�[�i�b�c)] �_�C�A���O �{�b�N�X   ": values(100, 2) = xlDialogGalleryDoughnut
values(101, 0) = "   xlDialogGalleryLine ": values(101, 1) = "   [�I�[�g�t�H�[�}�b�g (�܂��)] �_�C�A���O �{�b�N�X   ": values(101, 2) = xlDialogGalleryLine
values(102, 0) = "   xlDialogGalleryPie  ": values(102, 1) = "   [�I�[�g�t�H�[�}�b�g (�~)] �_�C�A���O �{�b�N�X   ": values(102, 2) = xlDialogGalleryPie
values(103, 0) = "   xlDialogGalleryRadar    ": values(103, 1) = "   [�I�[�g�t�H�[�}�b�g (���[�_�[)] �_�C�A���O �{�b�N�X ": values(103, 2) = xlDialogGalleryRadar
values(104, 0) = "   xlDialogGalleryScatter  ": values(104, 1) = "   [�I�[�g�t�H�[�}�b�g (�U�z�})] �_�C�A���O �{�b�N�X   ": values(104, 2) = xlDialogGalleryScatter
values(105, 0) = "   xlDialogGoalSeek    ": values(105, 1) = "   [�S�[�� �V�[�N] �_�C�A���O �{�b�N�X ": values(105, 2) = xlDialogGoalSeek
values(106, 0) = "   xlDialogGridlines   ": values(106, 1) = "   [�O���t �I�v�V���� (�ڐ���)] �_�C�A���O �{�b�N�X    ": values(106, 2) = xlDialogGridlines
values(107, 0) = "   xlDialogImportTextFile  ": values(107, 1) = "   [�e�L�X�g �t�@�C���̃C���|�[�g] �_�C�A���O �{�b�N�X ": values(107, 2) = xlDialogImportTextFile
values(108, 0) = "   xlDialogInsert  ": values(108, 1) = "   [�Z���̑}��] �_�C�A���O �{�b�N�X    ": values(108, 2) = xlDialogInsert
values(109, 0) = "   xlDialogInsertHyperlink ": values(109, 1) = "   [�n�C�p�[�����N�̑}��] �_�C�A���O �{�b�N�X  ": values(109, 2) = xlDialogInsertHyperlink
values(110, 0) = "   xlDialogInsertObject    ": values(110, 1) = "   [�I�u�W�F�N�g�̑}�� (�V�K�쐬)] �_�C�A���O �{�b�N�X ": values(110, 2) = xlDialogInsertObject
values(111, 0) = "   xlDialogInsertPicture   ": values(111, 1) = "   [�}�̑}��] �_�C�A���O �{�b�N�X  ": values(111, 2) = xlDialogInsertPicture
values(112, 0) = "   xlDialogInsertTitle ": values(112, 1) = "   [�^�C�g��/�����x���̑}��] �_�C�A���O �{�b�N�X   ": values(112, 2) = xlDialogInsertTitle
values(113, 0) = "   xlDialogLabelProperties ": values(113, 1) = "   [���x���̃v���p�e�B] �_�C�A���O �{�b�N�X    ": values(113, 2) = xlDialogLabelProperties
values(114, 0) = "   xlDialogListboxProperties   ": values(114, 1) = "   [���X�g �{�b�N�X�̃v���p�e�B] �_�C�A���O �{�b�N�X   ": values(114, 2) = xlDialogListboxProperties
values(115, 0) = "   xlDialogMacroOptions    ": values(115, 1) = "   [�}�N�� �I�v�V����] �_�C�A���O �{�b�N�X ": values(115, 2) = xlDialogMacroOptions
values(116, 0) = "   xlDialogMailEditMailer  ": values(116, 1) = "   [���[���ҏW���[���[] �_�C�A���O �{�b�N�X    ": values(116, 2) = xlDialogMailEditMailer
values(117, 0) = "   xlDialogMailLogon   ": values(117, 1) = "   [�񗗐�] �_�C�A���O �{�b�N�X    ": values(117, 2) = xlDialogMailLogon
values(118, 0) = "   xlDialogMailNextLetter  ": values(118, 1) = "   [���̎莆�̑��M] �_�C�A���O �{�b�N�X    ": values(118, 2) = xlDialogMailNextLetter
values(119, 0) = "   xlDialogMainChart   ": values(119, 1) = "   [���C�� �O���t] �_�C�A���O �{�b�N�X ": values(119, 2) = xlDialogMainChart
values(120, 0) = "   xlDialogMainChartType   ": values(120, 1) = "   [���C�� �O���t�̎��] �_�C�A���O �{�b�N�X   ": values(120, 2) = xlDialogMainChartType
values(121, 0) = "   xlDialogMenuEditor  ": values(121, 1) = "   [���j���[ �G�f�B�^�[] �_�C�A���O �{�b�N�X   ": values(121, 2) = xlDialogMenuEditor
values(122, 0) = "   xlDialogMove    ": values(122, 1) = "   [�ړ�] �_�C�A���O �{�b�N�X  ": values(122, 2) = xlDialogMove
values(123, 0) = "   xlDialogMyPermission    ": values(123, 1) = "   [�A�N�Z�X����] �_�C�A���O �{�b�N�X  ": values(123, 2) = xlDialogMyPermission
values(124, 0) = "   xlDialogNameManager ": values(124, 1) = "(x)[���O�̊Ǘ�] �_�C�A���O �{�b�N�X    ": values(124, 2) = 0 'xlDialogNameManager
values(125, 0) = "   xlDialogNew ": values(125, 1) = "   [�V�K�쐬 (�W��)] �_�C�A���O �{�b�N�X   ": values(125, 2) = xlDialogNew
values(126, 0) = "   xlDialogNewName ": values(126, 1) = "(x)[�V�������O] �_�C�A���O �{�b�N�X    ": values(126, 2) = 0 'xlDialogNewName
values(127, 0) = "   xlDialogNewWebQuery ": values(127, 1) = "   [�V���� Web �N�G��] �_�C�A���O �{�b�N�X ": values(127, 2) = xlDialogNewWebQuery
values(128, 0) = "   xlDialogNote    ": values(128, 1) = "   [�R�����g�̑}��] �_�C�A���O �{�b�N�X    ": values(128, 2) = xlDialogNote
values(129, 0) = "   xlDialogObjectProperties    ": values(129, 1) = "   [�I�u�W�F�N�g�̃v���p�e�B] �_�C�A���O �{�b�N�X  ": values(129, 2) = xlDialogObjectProperties
values(130, 0) = "   xlDialogObjectProtection    ": values(130, 1) = "   [�I�u�W�F�N�g�̕ی�] �_�C�A���O �{�b�N�X    ": values(130, 2) = xlDialogObjectProtection
values(131, 0) = "   xlDialogOpen    ": values(131, 1) = "   [�t�@�C�����J��] �_�C�A���O �{�b�N�X    ": values(131, 2) = xlDialogOpen
values(132, 0) = "   xlDialogOpenLinks   ": values(132, 1) = "   [�����N�����J��] �_�C�A���O �{�b�N�X    ": values(132, 2) = xlDialogOpenLinks
values(133, 0) = "   xlDialogOpenMail    ": values(133, 1) = "   [���[�����J��] �_�C�A���O �{�b�N�X  ": values(133, 2) = xlDialogOpenMail
values(134, 0) = "   xlDialogOpenText    ": values(134, 1) = "   [�e�L�X�g���J��] �_�C�A���O �{�b�N�X    ": values(134, 2) = xlDialogOpenText
values(135, 0) = "   xlDialogOptionsCalculation  ": values(135, 1) = "   [�I�v�V���� (�v�Z���@)] �_�C�A���O �{�b�N�X ": values(135, 2) = xlDialogOptionsCalculation
values(136, 0) = "   xlDialogOptionsChart    ": values(136, 1) = "   [�I�v�V���� (�O���t)] �_�C�A���O �{�b�N�X   ": values(136, 2) = xlDialogOptionsChart
values(137, 0) = "   xlDialogOptionsEdit ": values(137, 1) = "   [�I�v�V���� (�ҏW)] �_�C�A���O �{�b�N�X ": values(137, 2) = xlDialogOptionsEdit
values(138, 0) = "   xlDialogOptionsGeneral  ": values(138, 1) = "   [�I�v�V���� (�S��)] �_�C�A���O �{�b�N�X ": values(138, 2) = xlDialogOptionsGeneral
values(139, 0) = "   xlDialogOptionsListsAdd ": values(139, 1) = "   [�I�v�V���� (���[�U�[�ݒ胊�X�g)] �_�C�A���O �{�b�N�X   ": values(139, 2) = xlDialogOptionsListsAdd
values(140, 0) = "   xlDialogOptionsME   ": values(140, 1) = "   [�I�v�V���� (�C���^�[�i�V���i��)] �_�C�A���O �{�b�N�X   ": values(140, 2) = xlDialogOptionsME
values(141, 0) = "   xlDialogOptionsTransition   ": values(141, 1) = "   [�I�v�V���� (�ڍs)] �_�C�A���O �{�b�N�X ": values(141, 2) = xlDialogOptionsTransition
values(142, 0) = "   xlDialogOptionsView ": values(142, 1) = "   [�I�v�V���� (�\��)] �_�C�A���O �{�b�N�X ": values(142, 2) = xlDialogOptionsView
values(143, 0) = "   xlDialogOutline ": values(143, 1) = "   [�ݒ�] �_�C�A���O �{�b�N�X  ": values(143, 2) = xlDialogOutline
values(144, 0) = "   xlDialogOverlay ": values(144, 1) = "   [�d�ˍ��킹�O���t] �_�C�A���O �{�b�N�X  ": values(144, 2) = xlDialogOverlay
values(145, 0) = "   xlDialogOverlayChartType    ": values(145, 1) = "   [�O���t�̎�ނ̏d�ˍ��킹] �_�C�A���O �{�b�N�X  ": values(145, 2) = xlDialogOverlayChartType
values(146, 0) = "   xlDialogPageSetup   ": values(146, 1) = "   [�y�[�W�ݒ� (�y�[�W)] �_�C�A���O �{�b�N�X   ": values(146, 2) = xlDialogPageSetup
values(147, 0) = "   xlDialogParse   ": values(147, 1) = "   [��؂�ʒu] �_�C�A���O �{�b�N�X    ": values(147, 2) = xlDialogParse
values(148, 0) = "   xlDialogPasteNames  ": values(148, 1) = "   [���O�̓\��t��] �_�C�A���O �{�b�N�X    ": values(148, 2) = xlDialogPasteNames
values(149, 0) = "   xlDialogPasteSpecial    ": values(149, 1) = "   [�`����I�����ē\��t��] �_�C�A���O �{�b�N�X    ": values(149, 2) = xlDialogPasteSpecial
values(150, 0) = "   xlDialogPatterns    ": values(150, 1) = "   [�Z���̏����ݒ� (�p�^�[��)] �_�C�A���O �{�b�N�X ": values(150, 2) = xlDialogPatterns
values(151, 0) = "   xlDialogPermission  ": values(151, 1) = "   [�A�N�Z�X����] �_�C�A���O �{�b�N�X  ": values(151, 2) = xlDialogPermission
values(152, 0) = "   xlDialogPhonetic    ": values(152, 1) = "   [�ӂ肪�Ȃ̐ݒ� (�ӂ肪��)] �_�C�A���O �{�b�N�X ": values(152, 2) = xlDialogPhonetic
values(153, 0) = "   xlDialogPivotCalculatedField    ": values(153, 1) = "   [�s�{�b�g�W�v�t�B�[���h] �_�C�A���O �{�b�N�X    ": values(153, 2) = xlDialogPivotCalculatedField
values(154, 0) = "   xlDialogPivotCalculatedItem ": values(154, 1) = "   [�s�{�b�g�W�v�A�C�e��] �_�C�A���O �{�b�N�X  ": values(154, 2) = xlDialogPivotCalculatedItem
values(155, 0) = "   xlDialogPivotClientServerSet    ": values(155, 1) = "   [�s�{�b�g �N���C�A���g �T�[�o�[ �Z�b�g] �_�C�A���O �{�b�N�X ": values(155, 2) = xlDialogPivotClientServerSet
values(156, 0) = "   xlDialogPivotFieldGroup ": values(156, 1) = "   [�s�{�b�g �t�B�[���h �O���[�v] �_�C�A���O �{�b�N�X  ": values(156, 2) = xlDialogPivotFieldGroup
values(157, 0) = "   xlDialogPivotFieldProperties    ": values(157, 1) = "   [�s�{�b�g �t�B�[���h �v���p�e�B] �_�C�A���O �{�b�N�X    ": values(157, 2) = xlDialogPivotFieldProperties
values(158, 0) = "   xlDialogPivotFieldUngroup   ": values(158, 1) = "   [�s�{�b�g �t�B�[���h �O���[�v����] �_�C�A���O �{�b�N�X  ": values(158, 2) = xlDialogPivotFieldUngroup
values(159, 0) = "   xlDialogPivotShowPages  ": values(159, 1) = "   [�s�{�b�g�\���y�[�W] �_�C�A���O �{�b�N�X    ": values(159, 2) = xlDialogPivotShowPages
values(160, 0) = "   xlDialogPivotSolveOrder ": values(160, 1) = "   [�s�{�b�g��������] �_�C�A���O �{�b�N�X  ": values(160, 2) = xlDialogPivotSolveOrder
values(161, 0) = "   xlDialogPivotTableOptions   ": values(161, 1) = "   [�s�{�b�g�e�[�u�� �I�v�V����] �_�C�A���O �{�b�N�X   ": values(161, 2) = xlDialogPivotTableOptions
values(162, 0) = "   xlDialogPivotTableWizard    ": values(162, 1) = "   [�s�{�b�g�e�[�u��/�s�{�b�g�O���t �E�B�U�[�h] �_�C�A���O �{�b�N�X    ": values(162, 2) = xlDialogPivotTableWizard
values(163, 0) = "   xlDialogPlacement   ": values(163, 1) = "   [�\���ʒu] �_�C�A���O �{�b�N�X  ": values(163, 2) = xlDialogPlacement
values(164, 0) = "   xlDialogPrint   ": values(164, 1) = "   [���] �_�C�A���O �{�b�N�X  ": values(164, 2) = xlDialogPrint
values(165, 0) = "   xlDialogPrinterSetup    ": values(165, 1) = "   [�v�����^�[�̐ݒ�] �_�C�A���O �{�b�N�X  ": values(165, 2) = xlDialogPrinterSetup
values(166, 0) = "   xlDialogPrintPreview    ": values(166, 1) = "   [����v���r���[] �_�C�A���O �{�b�N�X    ": values(166, 2) = xlDialogPrintPreview
values(167, 0) = "   xlDialogPromote ": values(167, 1) = "   [�O���[�v�̉���] �_�C�A���O �{�b�N�X    ": values(167, 2) = xlDialogPromote
values(168, 0) = "   xlDialogProperties  ": values(168, 1) = "   [�v���p�e�B (�t�@�C���̊T�v)] �_�C�A���O �{�b�N�X   ": values(168, 2) = xlDialogProperties
values(169, 0) = "   xlDialogPropertyFields  ": values(169, 1) = "   [�v���p�e�B �t�B�[���h] �_�C�A���O �{�b�N�X ": values(169, 2) = xlDialogPropertyFields
values(170, 0) = "   xlDialogProtectDocument ": values(170, 1) = "   [�V�[�g�̕ی�] �_�C�A���O �{�b�N�X  ": values(170, 2) = xlDialogProtectDocument
values(171, 0) = "   xlDialogProtectSharing  ": values(171, 1) = "   [���L�u�b�N�̕ی�] �_�C�A���O �{�b�N�X  ": values(171, 2) = xlDialogProtectSharing
values(172, 0) = "   xlDialogPublishAsWebPage    ": values(172, 1) = "   [Web �y�[�W�Ƃ��Ĕ��s] �_�C�A���O �{�b�N�X  ": values(172, 2) = xlDialogPublishAsWebPage
values(173, 0) = "   xlDialogPushbuttonProperties    ": values(173, 1) = "   [�v�b�V�� �{�^���̃v���p�e�B] �_�C�A���O �{�b�N�X   ": values(173, 2) = xlDialogPushbuttonProperties
values(174, 0) = "   xlDialogReplaceFont ": values(174, 1) = "   [�t�H���g�̐ݒ�] �_�C�A���O �{�b�N�X    ": values(174, 2) = xlDialogReplaceFont
values(175, 0) = "   xlDialogRoutingSlip ": values(175, 1) = "   [�񗗐�] �_�C�A���O �{�b�N�X    ": values(175, 2) = xlDialogRoutingSlip
values(176, 0) = "   xlDialogRowHeight   ": values(176, 1) = "   [�s�̍���] �_�C�A���O �{�b�N�X  ": values(176, 2) = xlDialogRowHeight
values(177, 0) = "   xlDialogRun ": values(177, 1) = "   [�}�N��] �_�C�A���O �{�b�N�X    ": values(177, 2) = xlDialogRun
values(178, 0) = "   xlDialogSaveAs  ": values(178, 1) = "   [���O��t���ĕۑ�] �_�C�A���O �{�b�N�X  ": values(178, 2) = xlDialogSaveAs
values(179, 0) = "   xlDialogSaveCopyAs  ": values(179, 1) = "   [�R�s�[�𖼑O��t���ĕۑ�] �_�C�A���O �{�b�N�X  ": values(179, 2) = xlDialogSaveCopyAs
values(180, 0) = "   xlDialogSaveNewObject   ": values(180, 1) = "   �u�V�����I�u�W�F�N�g�̕ۑ�] �_�C�A���O �{�b�N�X ": values(180, 2) = xlDialogSaveNewObject
values(181, 0) = "   xlDialogSaveWorkbook    ": values(181, 1) = "   [���O��t���ĕۑ�] �_�C�A���O �{�b�N�X  ": values(181, 2) = xlDialogSaveWorkbook
values(182, 0) = "   xlDialogSaveWorkspace   ": values(182, 1) = "   [��Ə�Ԃ̕ۑ�] �_�C�A���O �{�b�N�X    ": values(182, 2) = xlDialogSaveWorkspace
values(183, 0) = "   xlDialogScale   ": values(183, 1) = "   [�{��] �_�C�A���O �{�b�N�X  ": values(183, 2) = xlDialogScale
values(184, 0) = "   xlDialogScenarioAdd ": values(184, 1) = "   [�V�i���I�̒ǉ�] �_�C�A���O �{�b�N�X    ": values(184, 2) = xlDialogScenarioAdd
values(185, 0) = "   xlDialogScenarioCells   ": values(185, 1) = "   [�V�i���I�̓o�^�ƊǗ�] �_�C�A���O �{�b�N�X  ": values(185, 2) = xlDialogScenarioCells
values(186, 0) = "   xlDialogScenarioEdit    ": values(186, 1) = "   [�V�i���I�̒ǉ�] �_�C�A���O �{�b�N�X    ": values(186, 2) = xlDialogScenarioEdit
values(187, 0) = "   xlDialogScenarioMerge   ": values(187, 1) = "   [�V�i���I�̃R�s�[] �_�C�A���O �{�b�N�X  ": values(187, 2) = xlDialogScenarioMerge
values(188, 0) = "   xlDialogScenarioSummary ": values(188, 1) = "   [�V�i���I�̏��] �_�C�A���O �{�b�N�X    ": values(188, 2) = xlDialogScenarioSummary
values(189, 0) = "   xlDialogScrollbarProperties ": values(189, 1) = "   [�X�N���[�� �o�[�̃v���p�e�B] �_�C�A���O �{�b�N�X   ": values(189, 2) = xlDialogScrollbarProperties
values(190, 0) = "   xlDialogSearch  ": values(190, 1) = "   [�ʏ�̃t�@�C������] �_�C�A���O �{�b�N�X    ": values(190, 2) = xlDialogSearch
values(191, 0) = "   xlDialogSelectSpecial   ": values(191, 1) = "   �I���I�v�V����] �_�C�A���O �{�b�N�X ": values(191, 2) = xlDialogSelectSpecial
values(192, 0) = "   xlDialogSendMail    ": values(192, 1) = "   [���b�Z�[�W (HTML �`��)] �_�C�A���O �{�b�N�X    ": values(192, 2) = xlDialogSendMail
values(193, 0) = "   xlDialogSeriesAxes  ": values(193, 1) = "   [�n��] �_�C�A���O �{�b�N�X    ": values(193, 2) = xlDialogSeriesAxes
values(194, 0) = "   xlDialogSeriesOptions   ": values(194, 1) = "   [�n��I�v�V����] �_�C�A���O �{�b�N�X    ": values(194, 2) = xlDialogSeriesOptions
values(195, 0) = "   xlDialogSeriesOrder ": values(195, 1) = "   [�n��̏���] �_�C�A���O �{�b�N�X    ": values(195, 2) = xlDialogSeriesOrder
values(196, 0) = "   xlDialogSeriesShape ": values(196, 1) = "   [�n��̌`��] �_�C�A���O �{�b�N�X    ": values(196, 2) = xlDialogSeriesShape
values(197, 0) = "   xlDialogSeriesX ": values(197, 1) = "   [�n�� X] �_�C�A���O �{�b�N�X    ": values(197, 2) = xlDialogSeriesX
values(198, 0) = "   xlDialogSeriesY ": values(198, 1) = "   [�f�[�^�n��̏����ݒ� (���O/�l)] �_�C�A���O �{�b�N�X    ": values(198, 2) = xlDialogSeriesY
values(199, 0) = "   xlDialogSetBackgroundPicture    ": values(199, 1) = "   [�V�[�g�̔w�i] �_�C�A���O �{�b�N�X  ": values(199, 2) = xlDialogSetBackgroundPicture
values(200, 0) = "   xlDialogSetPrintTitles  ": values(200, 1) = "   [����^�C�g���̐ݒ�] �_�C�A���O �{�b�N�X    ": values(200, 2) = xlDialogSetPrintTitles
values(201, 0) = "   xlDialogSetUpdateStatus ": values(201, 1) = "   [�X�V��Ԃ̐ݒ�] �_�C�A���O �{�b�N�X    ": values(201, 2) = xlDialogSetUpdateStatus
values(202, 0) = "   xlDialogShowDetail  ": values(202, 1) = "   [�ڍ׃f�[�^�̕\��] �_�C�A���O �{�b�N�X  ": values(202, 2) = xlDialogShowDetail
values(203, 0) = "   xlDialogShowToolbar ": values(203, 1) = "   [���[�U�[�ݒ� (�I�v�V����)] �_�C�A���O �{�b�N�X ": values(203, 2) = xlDialogShowToolbar
values(204, 0) = "   xlDialogSize    ": values(204, 1) = "   [�T�C�Y] �_�C�A���O �{�b�N�X    ": values(204, 2) = xlDialogSize
values(205, 0) = "   xlDialogSort    ": values(205, 1) = "   [���בւ�] �_�C�A���O �{�b�N�X  ": values(205, 2) = xlDialogSort
values(206, 0) = "   xlDialogSortSpecial ": values(206, 1) = "   [���בւ�] �_�C�A���O �{�b�N�X  ": values(206, 2) = xlDialogSortSpecial
values(207, 0) = "   xlDialogSplit   ": values(207, 1) = "   [��̕����A �s�̕���] �_�C�A���O �{�b�N�X   ": values(207, 2) = xlDialogSplit
values(208, 0) = "   xlDialogStandardFont    ": values(208, 1) = "   [�t�H���g�̐ݒ�] �_�C�A���O �{�b�N�X    ": values(208, 2) = xlDialogStandardFont
values(209, 0) = "   xlDialogStandardWidth   ": values(209, 1) = "   [�W���̕�] �_�C�A���O �{�b�N�X  ": values(209, 2) = xlDialogStandardWidth
values(210, 0) = "   xlDialogStyle   ": values(210, 1) = "   [�t�H���g�̐ݒ�] �_�C�A���O �{�b�N�X    ": values(210, 2) = xlDialogStyle
values(211, 0) = "   xlDialogSubscribeTo ": values(211, 1) = "   [���p] �_�C�A���O �{�b�N�X  ": values(211, 2) = xlDialogSubscribeTo
values(212, 0) = "   xlDialogSubtotalCreate  ": values(212, 1) = "   [�W�v�̐ݒ�] �_�C�A���O �{�b�N�X    ": values(212, 2) = xlDialogSubtotalCreate
values(213, 0) = "   xlDialogSummaryInfo ": values(213, 1) = "   [�v���p�e�B (�t�@�C���̊T�v)] �_�C�A���O �{�b�N�X   ": values(213, 2) = xlDialogSummaryInfo
values(214, 0) = "   xlDialogTable   ": values(214, 1) = "   [�e�[�u��] �_�C�A���O �{�b�N�X  ": values(214, 2) = xlDialogTable
values(215, 0) = "   xlDialogTabOrder    ": values(215, 1) = "   [�^�u �I�[�_�[�̐ݒ�] �_�C�A���O �{�b�N�X   ": values(215, 2) = xlDialogTabOrder
values(216, 0) = "   xlDialogTextToColumns   ": values(216, 1) = "   [��؂�ʒu] �_�C�A���O �{�b�N�X    ": values(216, 2) = xlDialogTextToColumns
values(217, 0) = "   xlDialogUnhide  ": values(217, 1) = "   [�E�B���h�E�̍ĕ\��] �_�C�A���O �{�b�N�X    ": values(217, 2) = xlDialogUnhide
values(218, 0) = "   xlDialogUpdateLink  ": values(218, 1) = "   [�����N�̍X�V] �_�C�A���O �{�b�N�X  ": values(218, 2) = xlDialogUpdateLink
values(219, 0) = "   xlDialogVbaInsertFile   ": values(219, 1) = "   [VBA �}���t�@�C��] �_�C�A���O �{�b�N�X  ": values(219, 2) = xlDialogVbaInsertFile
values(220, 0) = "   xlDialogVbaMakeAddin    ": values(220, 1) = "   [VBA �쐬�A�h�C��] �_�C�A���O �{�b�N�X  ": values(220, 2) = xlDialogVbaMakeAddin
values(221, 0) = "   xlDialogVbaProcedureDefinition  ": values(221, 1) = "   [VBA �菇��`] �_�C�A���O �{�b�N�X  ": values(221, 2) = xlDialogVbaProcedureDefinition
values(222, 0) = "   xlDialogView3d  ": values(222, 1) = "   [3D �\��] �_�C�A���O �{�b�N�X   ": values(222, 2) = xlDialogView3d
values(223, 0) = "   xlDialogWebOptionsBrowsers  ": values(223, 1) = "   [Web �I�v�V���� (�u���E�U�[)] �_�C�A���O �{�b�N�X   ": values(223, 2) = xlDialogWebOptionsBrowsers
values(224, 0) = "   xlDialogWebOptionsEncoding  ": values(224, 1) = "   [Web �I�v�V���� (�G���R�[�h)] �_�C�A���O �{�b�N�X   ": values(224, 2) = xlDialogWebOptionsEncoding
values(225, 0) = "   xlDialogWebOptionsFiles ": values(225, 1) = "   [Web �I�v�V���� (�t�@�C��)] �_�C�A���O �{�b�N�X ": values(225, 2) = xlDialogWebOptionsFiles
values(226, 0) = "   xlDialogWebOptionsFonts ": values(226, 1) = "   [Web �I�v�V���� (�t�H���g)] �_�C�A���O �{�b�N�X ": values(226, 2) = xlDialogWebOptionsFonts
values(227, 0) = "   xlDialogWebOptionsGeneral   ": values(227, 1) = "   [Web �I�v�V���� (�S��)] �_�C�A���O �{�b�N�X ": values(227, 2) = xlDialogWebOptionsGeneral
values(228, 0) = "   xlDialogWebOptionsPictures  ": values(228, 1) = "   [Web �I�v�V���� (�})] �_�C�A���O �{�b�N�X   ": values(228, 2) = xlDialogWebOptionsPictures
values(229, 0) = "   xlDialogWindowMove  ": values(229, 1) = "   [�E�B���h�E�̈ړ�] �_�C�A���O �{�b�N�X  ": values(229, 2) = xlDialogWindowMove
values(230, 0) = "   xlDialogWindowSize  ": values(230, 1) = "   [�E�B���h�E �T�C�Y] �_�C�A���O �{�b�N�X ": values(230, 2) = xlDialogWindowSize
values(231, 0) = "   xlDialogWorkbookAdd ": values(231, 1) = "   [�V�[�g�̈ړ��܂��̓R�s�[] �_�C�A���O �{�b�N�X  ": values(231, 2) = xlDialogWorkbookAdd
values(232, 0) = "   xlDialogWorkbookCopy    ": values(232, 1) = "   [�V�[�g�̈ړ��܂��̓R�s�[] �_�C�A���O �{�b�N�X  ": values(232, 2) = xlDialogWorkbookCopy
values(233, 0) = "   xlDialogWorkbookInsert  ": values(233, 1) = "   [�}�� (�W��)] �_�C�A���O �{�b�N�X   ": values(233, 2) = xlDialogWorkbookInsert
values(234, 0) = "   xlDialogWorkbookMove    ": values(234, 1) = "   [�V�[�g�̈ړ��܂��̓R�s�[] �_�C�A���O �{�b�N�X  ": values(234, 2) = xlDialogWorkbookMove
values(235, 0) = "   xlDialogWorkbookName    ": values(235, 1) = "   [�V�[�g���̕ύX] �_�C�A���O �{�b�N�X    ": values(235, 2) = xlDialogWorkbookName
values(236, 0) = "   xlDialogWorkbookNew ": values(236, 1) = "   [�}�� (�W��)] �_�C�A���O �{�b�N�X   ": values(236, 2) = xlDialogWorkbookNew
values(237, 0) = "   xlDialogWorkbookOptions ": values(237, 1) = "   [�V�[�g���̕ύX] �_�C�A���O �{�b�N�X    ": values(237, 2) = xlDialogWorkbookOptions
values(238, 0) = "   xlDialogWorkbookProtect ": values(238, 1) = "   [�u�b�N�̕ی�] �_�C�A���O �{�b�N�X  ": values(238, 2) = xlDialogWorkbookProtect
values(239, 0) = "   xlDialogWorkbookTabSplit    ": values(239, 1) = "   [�u�b�N�̃^�u����] �_�C�A���O �{�b�N�X  ": values(239, 2) = xlDialogWorkbookTabSplit
values(240, 0) = "   xlDialogWorkbookUnhide  ": values(240, 1) = "   [�ĕ\��] �_�C�A���O �{�b�N�X    ": values(240, 2) = xlDialogWorkbookUnhide
values(241, 0) = "   xlDialogWorkgroup   ": values(241, 1) = "   [�O���[�v�ҏW] �_�C�A���O �{�b�N�X  ": values(241, 2) = xlDialogWorkgroup
values(242, 0) = "   xlDialogWorkspace   ": values(242, 1) = "   [��Ə�Ԑݒ�] �_�C�A���O �{�b�N�X  ": values(242, 2) = xlDialogWorkspace
values(243, 0) = "   xlDialogZoom    ": values(243, 1) = "   [�Y�[��] �_�C�A���O �{�b�N�X    ": values(243, 2) = xlDialogZoom
     '���X�g �{�b�N�X�ɂ� 3 �̃f�[�^�񂪊܂܂�܂��B
    ListBox1.ColumnCount = 2

    'ListBox1 ����� ListBox2 �Ƀf�[�^��ǂݍ��݂܂��B
    ListBox1.List() = values

End Sub
