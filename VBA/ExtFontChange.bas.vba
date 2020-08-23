Attribute VB_Name = "ExtFontChange"
Rem --------------------------------------------------------------------------------
Rem
Rem  @module        ExtFontChange
Rem
Rem  @description   Excel�u�b�N�̃t�H���g�𓝈ꂷ��}�N��
Rem
Rem  @update        2020/08/07
Rem
Rem  @author        @KotorinChunChun (GitHub / Twitter)
Rem
Rem  @license       MIT (http://www.opensource.org/licenses/mit-license.php)
Rem
Rem --------------------------------------------------------------------------------
Rem  @references
Rem    Microsoft Scripting Runtime
Rem
Rem --------------------------------------------------------------------------------
Rem  @history
Rem     2019/07/21 : �u���O�f��
Rem     2020/08/23 : GitHub�f��
Rem
Rem --------------------------------------------------------------------------------
Rem  @note
Rem   [���J��]
Rem     �������邿��񂿂�� - Excel������S�V�b�N�̂�O��I�ɋ쒀���� Part3
Rem      https://www.excel-chunchun.com/entry/FontChange3
Rem
Rem   [�Ή����Ă���t�H���g]
Rem     �E���S�V�b�N
Rem     �E������
Rem     �E�l�r �o�S�V�b�N
Rem     �E�l�r �o����
Rem     �E���C���I�@�@�@�����C���I�͗񕝂��ϓ����邽�ߐ}�`�ό`�΍���Ή�
Rem
Rem --------------------------------------------------------------------------------

Option Explicit

Rem ��������{���ɓo�^������F8�Ŏ��s
Public Sub �A�N�e�B�u�u�b�N�̃t�H���g��MSP�S�V�b�N�ɓ���()
    Call �u�b�N�̃t�H���g���w��t�H���g�ɕύX(ActiveWorkbook, "�l�r �o�S�V�b�N", "�l�r �o����")
End Sub

Public Sub �A�N�e�B�u�u�b�N�̃t�H���g����S�V�b�N�ɓ���()
    Call �u�b�N�̃t�H���g���w��t�H���g�ɕύX(ActiveWorkbook, "���S�V�b�N", "������")
End Sub

Public Sub �A�N�e�B�u�u�b�N�̃t�H���g�����C���I�ɓ���()
'    Call �u�b�N�̃t�H���g���w��t�H���g�ɕύX(ActiveWorkbook, "���C���I", "���C���I")
    Call �u�b�N�̃t�H���g���w��t�H���g�ɕύX(ActiveWorkbook, "Meiryo UI", "Meiryo UI")
End Sub

Rem �ϊ��}�N���{��
Rem ��{�͟��S�V�b�N��AfterGothic�A��������AfterMincho�����A
Rem AfterGothic������̃t�H���g�ƂȂ�
Public Sub �u�b�N�̃t�H���g���w��t�H���g�ɕύX(wb As Workbook, AfterGothic As String, AfterMincho As String)
    
    '�t�H���g�ϊ��e�[�u�����쐬
    Dim fonts As Dictionary: Set fonts = New Dictionary
    Dim item As Variant
    
    '����̃t�H���g�����������݂��邱�Ƃ�����ׁA�S�ʂ�o�^
    For Each item In GetFonts("���S�V�b�N*", "YuGothic*", "Yu Gothic*", "�l�r �o�S�V�b�N*", "Meiryo*", "���C���I*")
        If Not fonts.Exists(item) Then
            fonts.Add item, AfterGothic
        End If
    Next
    For Each item In GetFonts("������*", "YuMincho*", "Yu Mincho*", "�l�r �o����*")
        If Not fonts.Exists(item) Then
            fonts.Add item, AfterMincho
        End If
    Next
    
    '�t�H���g�̕ϊ����������s
    Dim ws As Worksheet
    
    For Each ws In wb.Worksheets
        Call �s�̍������Œ�ɕύX(ws)
    Next
    
    Call �Z���̃X�^�C���̃t�H���g��ύX(wb, fonts)
    Call Office�e�[�}�t�H���g��ύX(wb, "" & fonts.Items(1))
    
    For Each ws In wb.Worksheets
        Call �Z���̃t�H���g��ύX(ws, fonts)
        Call �}�̃t�H���g��ύX(ws, fonts)
        Call �w�b�_�t�b�^�̃t�H���g��ύX(ws, fonts)
    Next
    
End Sub

'�w�肵���P�ꂩ��n�܂�t�H���g�����X�g��
Public Function GetFonts(ParamArray FontNameFiltes() As Variant) As Collection
    Set GetFonts = New Collection
    
    If TypeName(Excel.Selection) <> "Range" Then ActiveWindow.RangeSelection.Select
    
    With Application.CommandBars("Formatting").Controls(1)
        Dim i As Long
        For i = 1 To .ListCount
            Dim fnf As Variant
            For Each fnf In FontNameFiltes
                If .List(i) Like fnf Then
'                    Debug.Print .List(i)
                    GetFonts.Add .List(i)
                End If
            Next
        Next
    End With
End Function

Sub Test_GetFonts()
    Dim item
    For Each item In GetFonts("���S�V�b�N*", "YuGothic*", "Yu Gothic*")
        Debug.Print item
    Next
End Sub

Public Sub �s�̍������Œ�ɕύX(ws As Worksheet)
    With ws
        Dim rng As Range
        For Each rng In .Range(.Cells(1, 1), .UsedRange).EntireRow
            rng.RowHeight = rng.RowHeight
        Next
    End With
End Sub

Public Sub �Z���̃X�^�C���̃t�H���g��ύX(wb As Workbook, fonts As Dictionary)
    Dim st As Style
    For Each st In wb.Styles
        With st.Font
            Dim fts
            For Each fts In fonts.Keys
                If .Name = fts Then .Name = fonts(fts)
            Next
        End With
    Next
End Sub

Sub Test_Office�e�[�}�t�H���g��ύX()
    Call Office�e�[�}�t�H���g��ύX(ActiveWorkbook, "MSP�S�V�b�N")
    Call Office�e�[�}�t�H���g��ύX(ActiveWorkbook, "MSP����")
    Call Office�e�[�}�t�H���g��ύX(ActiveWorkbook, "MS P�S�V�b�N")
    Call Office�e�[�}�t�H���g��ύX(ActiveWorkbook, "MS P �S�V�b�N")
    Call Office�e�[�}�t�H���g��ύX(ActiveWorkbook, "���S�V�b�N")
End Sub

Public Sub Office�e�[�}�t�H���g��ύX(wb As Workbook, select_font As String)
    '���p�X��Excel 2016/Office365/32bit/64bit�ł������؂��Ă��Ȃ��B
    Const FP1 = "C:\Program Files\Microsoft Office\Root\Document Themes 16\"
    Const FP2 = "C:\Program Files (x86)\Microsoft Office\Root\Document Themes 16\"
    
    Dim fnt As String
    fnt = select_font
    fnt = Replace(Replace(Replace(fnt, "��", "yu"), "����", "mincho"), "�S�V�b�N", "gothic")
    fnt = Replace(fnt, "���C���I", "meiryo")
    fnt = Replace(StrConv(fnt, vbLowerCase + vbNarrow), " ", "")
    
    Dim fp As String
    If fnt Like "yugothic" Then
        fp = "Office Theme.thmx"
        fp = ChooseExistsFile(FP1 & fp, FP2 & fp)
        Debug.Print fp
        Call wb.ApplyTheme(fp)
    Else
        Select Case True
            Case fnt Like "mspgothic*": fp = "Office 2007 - 2010.xml"
            Case fnt Like "mspmincho*": fp = "Century Schoolbook.xml"
            Case fnt Like "meiryo*": fp = "Calibri.xml"
        End Select
        fp = "Theme Fonts\" & fp
        fp = ChooseExistsFile(FP1 & fp, FP2 & fp)
        Debug.Print fp
        Call wb.Theme.ThemeFontScheme.Load(fp)
    End If
End Sub

Rem �w�肵���t�@�C���p�X�ꗗ����ŏ��Ƀt�@�C�������݂����p�X��Ԃ�
Rem
Rem @param arr_filepath �t�@�C���̃t���p�X�̉ϒ�����
Rem
Rem @return As String   �t�@�C���p�X
Rem                     ���݂��Ȃ��ꍇ�͋�""��Ԃ�
Public Function ChooseExistsFile(ParamArray arr_filepath()) As String
    Dim fso As FileSystemObject: Set fso = New FileSystemObject
    Dim fp
    For Each fp In arr_filepath
        If fso.FileExists(fp) Then
            ChooseExistsFile = fp
            Exit Function
        End If
    Next
    Stop
End Function

Public Sub �Z���̃t�H���g��ύX(ws As Worksheet, fonts As Dictionary)
    Dim rng As Range
    For Each rng In ws.UsedRange
        With rng.Font
            Dim fts
            For Each fts In fonts.Keys
                If .Name = fts Then .Name = fonts(fts)
            Next
        End With
    Next
End Sub

Sub Test_�Z���̃t�H���g��ύX()
    Dim fonts As Dictionary: Set fonts = New Dictionary
    Dim item
    For Each item In GetFonts("���S�V�b�N*", "YuGothic*", "Yu Gothic*")
        If Not fonts.Exists(item) Then
            fonts.Add item, "�l�r �o�S�V�b�N"
        End If
    Next
    Call �Z���̃t�H���g��ύX(ActiveSheet, fonts)
End Sub

Public Sub �}�̃t�H���g��ύX(ws As Worksheet, fonts As Dictionary)
    Dim shp As Shape
    For Each shp In ws.Shapes
        '��TextFrame2��Excel 2007�ȍ~
        If shp.TextFrame2.HasText Then
            With shp.TextFrame2.TextRange.Font
                Dim fts
                For Each fts In fonts.Keys
                    If .Name = fts Then .Name = fonts(fts)
                    If .NameFarEast = fts Then .NameFarEast = fonts(fts)
                Next
            End With
        End If
    Next
End Sub

Public Sub �w�b�_�t�b�^�̃t�H���g��ύX(ws As Worksheet, fonts As Dictionary)
    Dim ps As PageSetup: Set ps = ws.PageSetup
    With ps
        Dim fts
        For Each fts In fonts.Keys
            'Replace���Ɵ��S�V�b�N�Ƃ����w�b�_�t�b�^�������u�����Ă��܂��o�O����
            If HeaderFooterFont(.LeftHeader) = fts Then
                .LeftFooter = Replace(.LeftFooter, fts, fonts(fts))
            End If
            If HeaderFooterFont(.CenterHeader) = fts Then
                .CenterHeader = Replace(.CenterHeader, fts, fonts(fts))
            End If
            If HeaderFooterFont(.RightHeader) = fts Then
                .RightHeader = Replace(.RightHeader, fts, fonts(fts))
            End If
            
            If HeaderFooterFont(.LeftFooter) = fts Then
                .LeftFooter = Replace(.LeftFooter, fts, fonts(fts))
            End If
            If HeaderFooterFont(.CenterFooter) = fts Then
                .CenterFooter = Replace(.CenterFooter, fts, fonts(fts))
            End If
            If HeaderFooterFont(.RightFooter) = fts Then
                .RightFooter = Replace(.RightFooter, fts, fonts(fts))
            End If
        Next
    End With
End Sub

Rem �w�b�_�E�t�b�^�̃t�H���g�������Ԃ�
Rem �����f�[�^�݂̂̎��@�@�F�@bbbb
Rem �t�H���g�ݒ莞�@�@�@�@�F�@&"HGP�n�p�p�߯�ߑ�,�W��"dddd
Rem �t�H���g�A�T�C�Y�ݒ莞�F�@&"���S�V�b�N,�W��"&16��������
Rem �t�H���g�A�����ݒ莞�@�F�@&"Yu Gothic UI Light,����"��������
Rem �t�H���g�A�Α̐ݒ莞�@�F�@&"-,�Α�"��������
Rem �t�H���g�A�����ݒ莦�@�F�@&"-,�W��"&U��������
Rem �t�H���g�A������ݒ莞�F�@&"-,�W��"&S��������
Rem �����E�ΆE�����ݒ莞�F�@&"-,���� �Α�"&U��������
Rem �o�^���ď������ʏ펞�@�F�@&"+,�W��"��������
Public Function HeaderFooterFont(hfValue As String) As String
    Dim S() As String
    S = Split(hfValue, """")
    If UBound(S, 1) > 0 Then
        HeaderFooterFont = Split(S(1), ",")(0)
    End If
End Function

Public Sub Test_HeaderFooterFont()
    Debug.Print HeaderFooterFont("")
    Debug.Print HeaderFooterFont("&""-,�Α�""&U��������")
    Debug.Print HeaderFooterFont("&""���S�V�b�N,�W��""&16��������")
    Debug.Print HeaderFooterFont("&""-,���� �Α�""&U��������")
    Debug.Print HeaderFooterFont("&""Yu Gothic UI Light,����""��������")
End Sub

'����̓u�b�N�ł͂Ȃ��A�v���̐ݒ�����������邽�ߑΏۊO
'Public Sub Excel�A�v���P�[�V�����̕W���t�H���g��ύX()
'    Application.StandardFont = "Meiryo UI"
'    Application.StandardFontSize = 11
'End Sub
